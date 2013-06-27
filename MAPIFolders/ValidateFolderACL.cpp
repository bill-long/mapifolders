#include "stdafx.h"
#include "ValidateFolderACL.h"
#include "MAPIFolders.h"


ValidateFolderACL::ValidateFolderACL(CComPtr<IMAPISession> session)
	:OperationBase(session)
{
}


ValidateFolderACL::~ValidateFolderACL(void)
{
}


void ValidateFolderACL::ProcessFolder(LPMAPIFOLDER folder)
{
	BOOL bValidDACL = false;
	PACL pACL = NULL;
	BOOL bDACLDefaulted = false;
	HRESULT hr = S_OK;
	SPropTagArray rgPropTag = { 1, PR_NT_SECURITY_DESCRIPTOR };
	LPSPropValue lpPropValue = NULL;
	ULONG cValues = 0;

	std::string folderPathString = this->GetStringFromFolderPath(folder);
	std::cout << "Checking ACL on folder: " << folderPathString.c_str() << std::endl;

	CORg(folder->GetProps(&rgPropTag, NULL, &cValues, &lpPropValue));
	if (lpPropValue && lpPropValue->ulPropTag == PR_NT_SECURITY_DESCRIPTOR)
	{
		PSECURITY_DESCRIPTOR pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(lpPropValue->Value.bin.lpb);
		if (!IsValidSecurityDescriptor(pSecurityDescriptor))
		{
			std::wcout << "Invalid security descriptor.";
			goto Error;
		}

		BOOL succeeded = GetSecurityDescriptorDacl(
		pSecurityDescriptor,
		&bValidDACL,
		&pACL,
		&bDACLDefaulted);

		if (!(succeeded && bValidDACL && pACL))
		{
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
			goto Error;
		}

		bool groupDenyEncountered = false;
		bool aclIsNonCanonical = false;

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
				// From MySecInfo.cpp in MfcMapi
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
					std::wcout << L"ACL on this folder is non-canonical!";
					aclIsNonCanonical = true;
				}

				delete[] lpSidDomain;
				delete[] lpSidName;

			}
		}
	}

Cleanup:
	return;
Error:
	goto Cleanup;
}
