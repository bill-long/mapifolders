#include "stdafx.h"
#include "ValidateFolderACL.h"
#include "MAPIFolders.h"

struct SavedACEInfo
{
	SBinary entryID;
	CURRENCY memberID;
	long accessRights;
};

ValidateFolderACL::ValidateFolderACL(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, bool fixBadACLs, Log *log, bool useAdmin)
	:OperationBase(pstrBasePath, pstrMailbox, nScope, log, useAdmin)
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
	this->psidEveryone = (PSID)malloc(SECURITY_MAX_SID_SIZE);
	DWORD dwLength = SECURITY_MAX_SID_SIZE;

	CORg(OperationBase::Initialize());

	if (!this->psidAnonymous || !this->psidEveryone)
	{
		*pLog << "Failed to allocate memory" << "\n";
		hr = HRESULT_FROM_WIN32(GetLastError());
		goto Error;
	}

	if (!CreateWellKnownSid(WinAnonymousSid, NULL, this->psidAnonymous, &dwLength))
	{
		*pLog << "Failed to create Anonymous SID" << "\n";
		hr = HRESULT_FROM_WIN32(GetLastError());
		goto Error;
	}

	dwLength = SECURITY_MAX_SID_SIZE;
	if (!CreateWellKnownSid(WinWorldSid, NULL, this->psidEveryone, &dwLength))
	{
		*pLog << "Failed to create Everyone SID" << "\n";
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
		*pLog << "Operation was not initialized." << "\n";
		goto Error;
	}

	*pLog << "Checking ACL on folder: " << folderPath.c_str() << "\n";

RetryGetProps:
	hr = folder->GetProps(&rgPropTag, NULL, &cValues, &lpPropValue);
	if (FAILED(hr))
	{
		if (hr == MAPI_E_TIMEOUT)
		{
			*pLog << "     Encountered a timeout trying to read the security descriptor. Retrying in 5 seconds..." << "\n";
			Sleep(5000);
			goto RetryGetProps;
		}
		else
		{
			*pLog << "     Failed to read security descriptor on this folder. hr = " << std::hex << hr << "\n";
			goto Error;
		}
	}

	if (lpPropValue && lpPropValue->ulPropTag == PR_NT_SECURITY_DESCRIPTOR)
	{
		PSECURITY_DESCRIPTOR pSecurityDescriptor = SECURITY_DESCRIPTOR_OF(lpPropValue->Value.bin.lpb);
		if (!IsValidSecurityDescriptor(pSecurityDescriptor))
		{
			*pLog << "     Invalid security descriptor." << "\n";
			goto Error;
		}

		BOOL succeeded = GetSecurityDescriptorDacl(
		pSecurityDescriptor,
		&bValidDACL,
		&pACL,
		&bDACLDefaulted);

		if (!(succeeded && bValidDACL && pACL))
		{
			*pLog << "     Invalid DACL." << "\n";
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
			*pLog << "     Could not get ACL information." << "\n";
			goto Error;
		}

		bool groupDenyEncountered = false;
		bool aclIsNonCanonical = false;
		BOOL foundAnonymous = false;
		BOOL foundEveryone = false;

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

				if (foundEveryone && !EqualSid(SidStart, this->psidEveryone))
				{
					// If we found any ACE that is not Everyone, but we already have the Everyone ACE,
					// this ACL is non-canonical.
					*pLog << "     ACL on this folder is non-canonical" << "\n";
					*pLog << "     Everyone group is not last" << "\n";
					aclIsNonCanonical = true;
					continue;
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
					*pLog << "     ACL on this folder is non-canonical" << "\n";
					aclIsNonCanonical = true;
				}

				if (!foundEveryone)
				{
					foundEveryone = EqualSid(SidStart, this->psidEveryone);
				}

				delete[] lpSidDomain;
				delete[] lpSidName;

			}
		}

		if (!foundAnonymous)
		{
			*pLog << "     ACL is missing Anonymous" << "\n";
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
		// Don't bother comparing null member IDs
		if (pRows->aRow[x].lpProps[ePR_MEMBER_ID].Value.l == 0x0)
		{
			continue;
		}

		for (y = 0; y < pRows->cRows; y++)
		{
			if (x == y)
			{
				continue;
			}

			if (pRows->aRow[x].lpProps[ePR_MEMBER_ID].Value.l == pRows->aRow[y].lpProps[ePR_MEMBER_ID].Value.l)
			{
				*pLog << "     ACL table has duplicate security principals." << "\n";
				aclTableIsGood = false;
				break;
			}

			if (pRows->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb > 0 && pRows->aRow[y].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb > 0)
			{
				ULONG result = 0;
				lpSession->CompareEntryIDs(
					pRows->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb,
					(LPENTRYID)pRows->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.lpb,
					pRows->aRow[y].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb,
					(LPENTRYID)pRows->aRow[y].lpProps[ePR_MEMBER_ENTRYID].Value.bin.lpb, NULL, &result);

				if (result > 0)
				{
					*pLog << "     ACL table has duplicate security principals." << "\n";
					aclTableIsGood = false;
					break;
				}
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

	*pLog << "     Attempting to fix ACL..." << "\n";

	CORg(folder->OpenProperty(PR_ACL_TABLE, &IID_IExchangeModifyTable, 0, MAPI_DEFERRED_ERRORS, (LPUNKNOWN*)&lpExchModTbl));
	CORg(lpExchModTbl->GetTable(0, &lpMapiTable));
	CORg(lpMapiTable->SetColumns((LPSPropTagArray)&rgAclTablePropTags, 0));
	CORg(HrQueryAllRows(lpMapiTable, NULL, NULL, NULL, NULL, &pRowsBefore));

	*pLog << "\n" << "     ACL table before changes:" << "\n";

	for (x = 0; x < pRowsBefore->cRows; x++)
	{
		*pLog << "     " << pRowsBefore->aRow[x].lpProps[ePR_MEMBER_NAME].Value.lpszW << "," << 
			pRowsBefore->aRow[x].lpProps[ePR_MEMBER_RIGHTS].Value.l;
#ifdef DEBUG
		*pLog << "," <<
			pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.cb << ","; // member ID value is stored in cb... weird, but whatever
		OutputSBinary(pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin);
#endif
		*pLog << "\n";

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
				ULONG result = 0;
				CORg(this->lpSession->CompareEntryIDs(
					savedACEs[y].entryID.cb,
					(LPENTRYID) savedACEs[y].entryID.lpb,
					pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb,
					(LPENTRYID) pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.lpb,
					NULL, &result));

				if (result)
				{
					found = true;
				}
			}

			// Only save each entry to the list once, so we drop any duplicates
			if (!found)
			{
				savedACEs[savedACECount].entryID.cb = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb;
				savedACEs[savedACECount].entryID.lpb = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.lpb;
				savedACEs[savedACECount].memberID.Lo = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ID].Value.cur.Lo;
				savedACEs[savedACECount].memberID.Hi = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_ID].Value.cur.Hi;
				savedACEs[savedACECount].accessRights = pRowsBefore->aRow[x].lpProps[ePR_MEMBER_RIGHTS].Value.l;
				savedACECount++;

				if (savedACECount > 500)
				{
					*pLog << "     Error: ACL table has too many entries." << "\n";
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
			if (pRowsBefore->aRow[y].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb > 0)
			{
				ULONG result = 0;
				CORg(this->lpSession->CompareEntryIDs(
					savedACEs[x].entryID.cb,
					(LPENTRYID)savedACEs[x].entryID.lpb,
					pRowsBefore->aRow[y].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb,
					(LPENTRYID)pRowsBefore->aRow[y].lpProps[ePR_MEMBER_ENTRYID].Value.bin.lpb,
					NULL, &result));

				if (result)
				{
					acesToRemove++;
				}
			}
		}

		for (y = 0; y < acesToRemove; y++)
		{
			SPropValue propRemove[1] = {0};
			ROWLIST rowListRemove = {0};

			propRemove[0].ulPropTag = PR_MEMBER_ID;
			propRemove[0].Value.cur.Lo = savedACEs[x].memberID.Lo;
			propRemove[0].Value.cur.Hi = savedACEs[x].memberID.Hi;

			rowListRemove.cEntries = 1;
			rowListRemove.aEntries->ulRowFlags = ROW_REMOVE;
			rowListRemove.aEntries->cValues = 1;
			rowListRemove.aEntries->rgPropVals = &propRemove[0];
			
#ifdef DEBUG
			*pLog << "    Removing member ID: " << savedACEs[x].memberID.Lo << ":" << savedACEs[x].memberID.Hi << "\n";
#endif
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

#ifdef DEBUG
		*pLog << "    Adding entry ID: ";
		OutputSBinary(savedACEs[x].entryID);
		*pLog << "\n";
#endif
		CORg(lpExchModTbl->ModifyTable(0, &rowListAdd));
	}

	// Finally, if we didn't find Anonymous, add it now
	// This approach works on E15, but might not work
	// on older versions of Exchange. Need to test.
	if (!foundAnonymous)
	{
		SPropValue prop[2] = {0};
		prop[0].ulPropTag = PR_MEMBER_ID;
		// For anonymous, the first 8 bytes of the value all need to be 0xFF. The
		// easiest way to do this on both x86 and x64 is to just set cur.
		prop[0].Value.cur.Lo = 0xFFFFFFFF;
		prop[0].Value.cur.Hi = 0xFFFFFFFF;
		prop[1].ulPropTag = PR_MEMBER_RIGHTS;
		prop[1].Value.l = RIGHTS_NONE;

		ROWLIST rowList = {0};
		rowList.cEntries = 1;
		rowList.aEntries->ulRowFlags = ROW_MODIFY;
		rowList.aEntries->cValues = 2;
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
	*pLog << "\n" << "     ACL table after changes:" << "\n";

	for (x = 0; x < pRowsTemp->cRows; x++)
	{
		*pLog << "     " << pRowsTemp->aRow[x].lpProps[ePR_MEMBER_NAME].Value.lpszW << "," << 
			pRowsTemp->aRow[x].lpProps[ePR_MEMBER_RIGHTS].Value.l;
#ifdef DEBUG
		*pLog << "," <<
			pRowsTemp->aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.cb << ","; // member ID value is stored in cb... weird, but whatever
		OutputSBinary(pRowsTemp->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin);
#endif
		*pLog << "\n";
	}

	*pLog << "\n";

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