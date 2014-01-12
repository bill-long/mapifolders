#include "stdafx.h"
#include "ValidateFolderACL.h"
#include "MAPIFolders.h"

enum {
    ePR_MEMBER_ENTRYID, 
    ePR_MEMBER_RIGHTS,  
    ePR_MEMBER_ID, 
    ePR_MEMBER_NAME, 
    NUM_COLS
};

SizedSPropTagArray(NUM_COLS, rgPropTag) =
{
    NUM_COLS,
    {
        PR_MEMBER_ENTRYID,  // Unique across directory.
        PR_MEMBER_RIGHTS,  
        PR_MEMBER_ID,       // Unique within ACL table. 
        PR_MEMBER_NAME,     // Display name.
    }
};

ValidateFolderACL::ValidateFolderACL(tstring *pstrBasePath, UserArgs::ActionScope nScope, bool fixBadACLs)
	:OperationBase(pstrBasePath, nScope)
{
	this->FixBadACLs = fixBadACLs;
	this->IsInitialized = false;
}


ValidateFolderACL::~ValidateFolderACL(void)
{
}

HRESULT ValidateFolderACL::Initialize(void)
{
	HRESULT hr = S_OK;
	this->psidAnonymous = (PSID)malloc(SECURITY_MAX_SID_SIZE);
	DWORD dwLength = SECURITY_MAX_SID_SIZE;
	std::wcout << L"Buffer size:" << dwLength << std::endl;
	if (!this->psidAnonymous)
	{
		std::wcout << L"Failed to allocate memory" << std::endl;
		hr = HRESULT_FROM_WIN32(GetLastError());
		goto Error;
	}

	if (!CreateWellKnownSid(WinAnonymousSid, NULL, this->psidAnonymous, &dwLength))
	{
		std::wcout << L"Failed to create Anonymous SID" << std::endl;
		hr = HRESULT_FROM_WIN32(GetLastError());
		goto Error;
	}

	this->IsInitialized = true;

Cleanup:
	return hr;
Error:
	goto Cleanup;
}

void ValidateFolderACL::ProcessFolder(LPMAPIFOLDER folder, std::wstring folderPath)
{
	BOOL bValidDACL = false;
	PACL pACL = NULL;
	BOOL bDACLDefaulted = false;
	HRESULT hr = S_OK;
	SPropTagArray rgPropTag = { 1, PR_NT_SECURITY_DESCRIPTOR };
	LPSPropValue lpPropValue = NULL;
	ULONG cValues = 0;

	if (!this->IsInitialized)
	{
		std::wcout << L"Operation was not initialized." << std::endl;
		goto Error;
	}

	std::wcout << L"Checking ACL on folder: " << folderPath.c_str() << std::endl;

RetryGetProps:
	hr = folder->GetProps(&rgPropTag, NULL, &cValues, &lpPropValue);
	if (FAILED(hr))
	{
		if (hr == MAPI_E_TIMEOUT)
		{
			std::wcout << L"     Encountered a timeout trying to read the security descriptor. Retrying in 5 seconds..." << std::endl;
			Sleep(5000);
			goto RetryGetProps;
		}
		else
		{
			std::wcout << L"     Failed to read security descriptor on this folder. hr = " << std::hex << hr << std::endl;
			goto Error;
		}
	}

	if (lpPropValue && lpPropValue->ulPropTag == PR_NT_SECURITY_DESCRIPTOR)
	{
		PSECURITY_DESCRIPTOR pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(lpPropValue->Value.bin.lpb);
		if (!IsValidSecurityDescriptor(pSecurityDescriptor))
		{
			std::wcout << L"     Invalid security descriptor." << std::endl;
			goto Error;
		}

		BOOL succeeded = GetSecurityDescriptorDacl(
		pSecurityDescriptor,
		&bValidDACL,
		&pACL,
		&bDACLDefaulted);

		if (!(succeeded && bValidDACL && pACL))
		{
			std::wcout << L"     Invalid DACL." << std::endl;
			goto Error;
		}

		ACL_SIZE_INFORMATION ACLSizeInfo = {0};
		succeeded = GetAclInformation(
			pACL,
			&ACLSizeInfo,
			sizeof(ACLSizeInfo),
			AclSizeInformation);

		if (!succeeded)
		{
			std::wcout << L"     Could not get ACL information." << std::endl;
			goto Error;
		}

		bool groupDenyEncountered = false;
		bool aclIsNonCanonical = false;
		BOOL foundAnonymous = false;

		for (DWORD x = 0; x < ACLSizeInfo.AceCount; x++)
		{
			if (aclIsNonCanonical)
				break;

			void* pACE = NULL;
			succeeded = GetAce(
				pACL,
				x,
				&pACE);

			if (pACE)
			{
				//
				// Most of the following code from MySecInfo.cpp in MfcMapi
				//

				BYTE	AceType = 0;
				BYTE	AceFlags = 0;
				ACCESS_MASK	Mask = 0;
				DWORD	Flags = 0;
				GUID	ObjectType = {0};
				GUID	InheritedObjectType = {0};
				SID*	SidStart = NULL;
				bool	bObjectFound = false;

				/* Check type of ACE */
				switch (((PACE_HEADER)pACE)->AceType)
				{
				case ACCESS_ALLOWED_ACE_TYPE:
					AceType = ((ACCESS_ALLOWED_ACE *) pACE)->Header.AceType;
					AceFlags = ((ACCESS_ALLOWED_ACE *) pACE)->Header.AceFlags;
					Mask = ((ACCESS_ALLOWED_ACE *) pACE)->Mask;
					SidStart = (SID *) &((ACCESS_ALLOWED_ACE *) pACE)->SidStart;
					break;
				case ACCESS_DENIED_ACE_TYPE:
					AceType = ((ACCESS_DENIED_ACE *) pACE)->Header.AceType;
					AceFlags = ((ACCESS_DENIED_ACE *) pACE)->Header.AceFlags;
					Mask = ((ACCESS_DENIED_ACE *) pACE)->Mask;
					SidStart = (SID *) &((ACCESS_DENIED_ACE *) pACE)->SidStart;
					break;
				case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
					AceType = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->Header.AceType;
					AceFlags = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->Header.AceFlags;
					Mask = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->Mask;
					Flags = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->Flags;
					ObjectType = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->ObjectType;
					InheritedObjectType = ((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->InheritedObjectType;
					SidStart = (SID *) &((ACCESS_ALLOWED_OBJECT_ACE *) pACE)->SidStart;
					bObjectFound = true;
					break;
				case ACCESS_DENIED_OBJECT_ACE_TYPE:
					AceType = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->Header.AceType;
					AceFlags = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->Header.AceFlags;
					Mask = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->Mask;
					Flags = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->Flags;
					ObjectType = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->ObjectType;
					InheritedObjectType = ((ACCESS_DENIED_OBJECT_ACE *) pACE)->InheritedObjectType;
					SidStart = (SID *) &((ACCESS_DENIED_OBJECT_ACE *) pACE)->SidStart;
					bObjectFound = true;
					break;
				}

				if (!foundAnonymous)
				{
					foundAnonymous = EqualSid(SidStart, this->psidAnonymous);
				}

				DWORD dwSidName = 0;
				DWORD dwSidDomain = 0;
				SID_NAME_USE SidNameUse;

				CORg(LookupAccountSid(
					NULL,
					SidStart,
					NULL,
					&dwSidName,
					NULL,
					&dwSidDomain,
					&SidNameUse));
				hr = S_OK;

				LPTSTR lpSidName = NULL;
				LPTSTR lpSidDomain = NULL;

			#pragma warning(push)
			#pragma warning(disable:6211)
				if (dwSidName) lpSidName = new TCHAR[dwSidName];
				if (dwSidDomain) lpSidDomain = new TCHAR[dwSidDomain];
			#pragma warning(pop)

				// Only make the call if we got something to get
				if (lpSidName || lpSidDomain)
				{
					CORg(LookupAccountSid(
						NULL,
						SidStart,
						lpSidName,
						&dwSidName,
						lpSidDomain,
						&dwSidDomain,
						&SidNameUse));
					hr = S_OK;
				}

				if (AceType == ACCESS_DENIED_ACE_TYPE && SidNameUse == SidTypeGroup)
				{
					groupDenyEncountered = true;
				}
				else if (AceType == ACCESS_ALLOWED_ACE_TYPE && SidNameUse == SidTypeGroup && groupDenyEncountered == true)
				{
					std::wcout << L"     ACL on this folder is non-canonical" << std::endl;
					aclIsNonCanonical = true;
				}

				delete[] lpSidDomain;
				delete[] lpSidName;

			}
		}

		if (!foundAnonymous)
		{
			std::wcout << L"     ACL is missing Anonymous" << std::endl;
			aclIsNonCanonical = true;
		}

		bool aclTableIsGood;
		CORg(this->CheckACLTable(folder, aclTableIsGood));

		if ((aclIsNonCanonical || !aclTableIsGood) && this->FixBadACLs)
		{
			this->FixACL(folder);
		}
	}

Cleanup:
	return;
Error:
	goto Cleanup;
}

// Returns true if there is a problem
HRESULT ValidateFolderACL::CheckACLTable(LPMAPIFOLDER folder, bool &aclTableIsGood)
{
	aclTableIsGood = true;
	HRESULT hr = S_OK;
	LPEXCHANGEMODIFYTABLE lpExchModTbl = NULL;
	LPMAPITABLE lpMapiTable = NULL;
	ULONG lpulRowCount = NULL;
	LPSRowSet pRows = NULL;
	UINT x = 0;
	UINT y = 0;

	CORg(folder->OpenProperty(PR_ACL_TABLE, &IID_IExchangeModifyTable, 0, MAPI_DEFERRED_ERRORS, (LPUNKNOWN*)&lpExchModTbl));
	CORg(lpExchModTbl->GetTable(0, &lpMapiTable));
	CORg(lpMapiTable->SetColumns((LPSPropTagArray)&rgPropTag, 0));
	CORg(HrQueryAllRows(lpMapiTable, NULL, NULL, NULL, NULL, &pRows));

	// Use a nested loop to check for duplicate entries
	for (x = 0; x < pRows->cRows; x++)
	{
		for (y = 0; y < pRows->cRows; y++)
		{
			if (x == y)
			{
				continue;
			}

			if (pRows->aRow[x].lpProps[ePR_MEMBER_ID].Value.l == pRows->aRow[y].lpProps[ePR_MEMBER_ID].Value.l)
			{
				std::wcout << "     ACL table has duplicate security principals." << std::endl;
				aclTableIsGood = false;
				break;
			}
		}

		if (!aclTableIsGood)
			break;
	}

Cleanup:
	if (pRows)
		FreeProws(pRows);
	if (lpMapiTable)
		lpMapiTable->Release();
	if (lpExchModTbl)
		lpExchModTbl->Release();

	return hr;
Error:
	goto Cleanup;
}

void ValidateFolderACL::FixACL(LPMAPIFOLDER folder)
{

}
