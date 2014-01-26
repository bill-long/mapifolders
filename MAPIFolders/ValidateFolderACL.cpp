#include "stdafx.h"
#include "ValidateFolderACL.h"
#include "MAPIFolders.h"

struct SavedACEInfo
{
	SBinary entryID;
	SBinary memberID;
	long accessRights;
};

ValidateFolderACL::ValidateFolderACL(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, bool fixBadACLs)
	:OperationBase(pstrBasePath, pstrMailbox, nScope)
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

	CORg(OperationBase::Initialize());

	if (!this->psidAnonymous)
	{
		tcout << "Failed to allocate memory" << std::endl;
		hr = HRESULT_FROM_WIN32(GetLastError());
		goto Error;
	}

	if (!CreateWellKnownSid(WinAnonymousSid, NULL, this->psidAnonymous, &dwLength))
	{
		tcout << "Failed to create Anonymous SID" << std::endl;
		hr = HRESULT_FROM_WIN32(GetLastError());
		goto Error;
	}

	this->IsInitialized = true;

Cleanup:
	return hr;
Error:
	goto Cleanup;
}

void ValidateFolderACL::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
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
		tcout << "Operation was not initialized." << std::endl;
		goto Error;
	}

	tcout << "Checking ACL on folder: " << folderPath.c_str() << std::endl;

RetryGetProps:
	hr = folder->GetProps(&rgPropTag, NULL, &cValues, &lpPropValue);
	if (FAILED(hr))
	{
		if (hr == MAPI_E_TIMEOUT)
		{
			tcout << "     Encountered a timeout trying to read the security descriptor. Retrying in 5 seconds..." << std::endl;
			Sleep(5000);
			goto RetryGetProps;
		}
		else
		{
			tcout << "     Failed to read security descriptor on this folder. hr = " << std::hex << hr << std::endl;
			goto Error;
		}
	}

	if (lpPropValue && lpPropValue->ulPropTag == PR_NT_SECURITY_DESCRIPTOR)
	{
		PSECURITY_DESCRIPTOR pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(lpPropValue->Value.bin.lpb);
		if (!IsValidSecurityDescriptor(pSecurityDescriptor))
		{
			tcout << "     Invalid security descriptor." << std::endl;
			goto Error;
		}

		BOOL succeeded = GetSecurityDescriptorDacl(
		pSecurityDescriptor,
		&bValidDACL,
		&pACL,
		&bDACLDefaulted);

		if (!(succeeded && bValidDACL && pACL))
		{
			tcout << "     Invalid DACL." << std::endl;
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
			tcout << "     Could not get ACL information." << std::endl;
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
					tcout << "     ACL on this folder is non-canonical" << std::endl;
					aclIsNonCanonical = true;
				}

				delete[] lpSidDomain;
				delete[] lpSidName;

			}
		}

		if (!foundAnonymous)
		{
			tcout << "     ACL is missing Anonymous" << std::endl;
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
	if (lpPropValue)
		MAPIFreeBuffer(lpPropValue);
	return;
Error:
	goto Cleanup;
}

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
	CORg(lpMapiTable->SetColumns((LPSPropTagArray)&rgAclTablePropTags, 0));
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
				tcout << "     ACL table has duplicate security principals." << std::endl;
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
	HRESULT hr = S_OK;
	LPEXCHANGEMODIFYTABLE lpExchModTbl = NULL;
	LPMAPITABLE lpMapiTable = NULL;
	ULONG lpulRowCount = NULL;
	LPSRowSet pRowsBefore = NULL; // The pRows that backs the stuff saved in SavedACEInfo; do not free until end of method
	LPSRowSet pRowsTemp = NULL;   // The pRows we use for the other queries; this one can be freed
	UINT x = 0;
	UINT y = 0;
	SavedACEInfo savedACEs[500];
	UINT savedACECount = 0;
	bool foundAnonymous = false;

	tcout << "     Attempting to fix ACL..." << std::endl;

	CORg(folder->OpenProperty(PR_ACL_TABLE, &IID_IExchangeModifyTable, 0, MAPI_DEFERRED_ERRORS, (LPUNKNOWN*)&lpExchModTbl));
	CORg(lpExchModTbl->GetTable(0, &lpMapiTable));
	CORg(lpMapiTable->SetColumns((LPSPropTagArray)&rgAclTablePropTags, 0));
	CORg(HrQueryAllRows(lpMapiTable, NULL, NULL, NULL, NULL, &pRowsBefore));

	tcout << std::endl << "     ACL table before changes:" << std::endl;

	for (x = 0; x < pRowsBefore->cRows; x++)
	{
		tcout << "     " << pRowsBefore->aRow[x].lpProps[ePR_MEMBER_NAME].Value.lpszW << "," << 
			pRowsBefore->aRow[x].lpProps[ePR_MEMBER_RIGHTS].Value.l << std::endl;

		if (strAnonymous == tstring(pRowsBefore->aRow[x].lpProps[ePR_MEMBER_NAME].Value.lpszW))
		{
			foundAnonymous = true;
		}

		// Default and Anonymous don't have entry IDs. We won't be removing and re-adding them,
		// so we just ignore them here.
		if (pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb > 0)
		{
			bool found = false;
			for (y = 0; y < savedACECount; y++)
			{
				if (IsEntryIdEqual(savedACEs[y].entryID, pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin))
				{
					found = true;
					break;
				}
			}

			// Only save each entry to the list once, so we drop any duplicates
			if (!found)
			{
				savedACEs[savedACECount].entryID.cb = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb;
				savedACEs[savedACECount].entryID.lpb = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.lpb;
				savedACEs[savedACECount].memberID.cb = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.cb;
				savedACEs[savedACECount].memberID.lpb = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.lpb;
				savedACEs[savedACECount].accessRights = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_RIGHTS].Value.l;
				savedACECount++;

				if (savedACECount > 500)
				{
					tcout << "     Error: ACL table has too many entries." << std::endl;
					goto Error;
				}
			}
		}
	}

	// Now remove and re-add each entry ID we found
	for (x = 0; x < savedACECount; x++)
	{
		// First, determine how many times we need to remove it
		UINT acesToRemove = 0;
		for (y = 0; y < pRowsBefore->cRows; y++)
		{
			if (IsEntryIdEqual(savedACEs[x].entryID, pRowsBefore->aRow[y].lpProps[ePR_MEMBER_ENTRYID].Value.bin))
			{
				acesToRemove++;
			}
		}

		for (y = 0; y < acesToRemove; y++)
		{
			SPropValue propRemove[1] = {0};
			ROWLIST rowListRemove = {0};

			propRemove[0].ulPropTag = PR_MEMBER_ID;
			propRemove[0].Value.bin.cb = savedACEs[x].memberID.cb;
			propRemove[0].Value.bin.lpb = savedACEs[x].memberID.lpb;

			rowListRemove.cEntries = 1;
			rowListRemove.aEntries->ulRowFlags = ROW_REMOVE;
			rowListRemove.aEntries->cValues = 1;
			rowListRemove.aEntries->rgPropVals = &propRemove[0];

			CORg(lpExchModTbl->ModifyTable(0, &rowListRemove));
		}

		// We might be reusing this, so check if we need to free
		if (pRowsTemp)
		{
			FreeProws(pRowsTemp);
		}
		pRowsTemp = NULL;

		// It's gone. Now add it back.
		SPropValue propAdd[2] = {0};
		ROWLIST rowListAdd = {0};
		propAdd[0].ulPropTag = PR_MEMBER_ENTRYID;
		propAdd[0].Value.bin.cb = savedACEs[x].entryID.cb;
		propAdd[0].Value.bin.lpb = savedACEs[x].entryID.lpb;
		propAdd[1].ulPropTag = PR_MEMBER_RIGHTS;
		propAdd[1].Value.l = savedACEs[x].accessRights;

		rowListAdd.cEntries = 1;
		rowListAdd.aEntries->ulRowFlags = ROW_ADD;
		rowListAdd.aEntries->cValues = 2;
		rowListAdd.aEntries->rgPropVals = &propAdd[0];

		CORg(lpExchModTbl->ModifyTable(0, &rowListAdd));
	}

	// Finally, if we didn't find Anonymous, add it now
	if (!foundAnonymous)
	{
		SPropValue prop[1] = {0};
		prop[0].ulPropTag = PR_MEMBER_ID;
		prop[0].Value.dbl = -1;

		ROWLIST rowList = {0};
		rowList.cEntries = 1;
		rowList.aEntries->ulRowFlags = ROW_ADD;
		rowList.aEntries->cValues = 1;
		rowList.aEntries->rgPropVals = &prop[0];

		CORg(lpExchModTbl->ModifyTable(0, &rowList));
	}

	// Done making changes. Free the table and then request it again.
	// Reusing pRowsTemp again.
	if (pRowsTemp)
		FreeProws(pRowsTemp);
	if (lpMapiTable)
		lpMapiTable->Release();
	if (lpExchModTbl)
		lpExchModTbl->Release();
	
	pRowsTemp = NULL;
	lpExchModTbl = NULL;
	lpMapiTable = NULL;

	CORg(folder->OpenProperty(PR_ACL_TABLE, &IID_IExchangeModifyTable, 0, MAPI_DEFERRED_ERRORS, (LPUNKNOWN*)&lpExchModTbl));
	CORg(lpExchModTbl->GetTable(0, &lpMapiTable));
	CORg(lpMapiTable->SetColumns((LPSPropTagArray)&rgAclTablePropTags, 0));
	CORg(HrQueryAllRows(lpMapiTable, NULL, NULL, NULL, NULL, &pRowsTemp));
	tcout << std::endl << "     ACL table after changes:" << std::endl;

	for (x = 0; x < pRowsTemp->cRows; x++)
	{
		tcout << "     " << pRowsTemp->aRow[x].lpProps[ePR_MEMBER_NAME].Value.lpszW << "," << 
			pRowsTemp->aRow[x].lpProps[ePR_MEMBER_RIGHTS].Value.l << std::endl;
	}

	tcout << std::endl;

Cleanup:
	if (pRowsTemp)
		FreeProws(pRowsTemp);
	if (pRowsBefore)
		FreeProws(pRowsBefore);
	if (lpMapiTable)
		lpMapiTable->Release();
	if (lpExchModTbl)
		lpExchModTbl->Release();

	return;
Error:
	goto Cleanup;
}
