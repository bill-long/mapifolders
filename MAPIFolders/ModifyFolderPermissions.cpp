#include "stdafx.h"
#include "MAPIFolders.h"
#include "ModifyFolderPermissions.h"
#include "OperationBase.h"

ModifyFolderPermissions::ModifyFolderPermissions(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, bool remove, tstring *pstrUserString, tstring *pstrRightsString)
	:OperationBase(pstrBasePath, pstrMailbox, nScope)
{
	this->remove = remove;
	this->pstrUserString = pstrUserString;
	this->pstrRightsString = pstrRightsString;
}

HRESULT ModifyFolderPermissions::Initialize(void)
{
	HRESULT hr = S_OK;
	LPADRBOOK lpAdrBook = NULL;
	LPADRLIST lpAdrList = NULL;

	// Resolve the user string
	this->resolvedUserEID = NULL;
	if (_tstricmp(strAnonymous.c_str(), this->pstrUserString->c_str()) == 0)
	{
		if (this->remove)
		{
			tcout << _T("Error: Removing Anonymous is not permitted.") << std::endl;
			hr = MAPI_E_INVALID_PARAMETER;
			goto Error;
		}

		this->pstrUserString = new tstring(_T("Anonymous"));
	}
	else if (_tstricmp(strDefault.c_str(), this->pstrUserString->c_str()))
	{
		if (this->remove)
		{
			tcout << _T("Error: Removing Default is not permitted.") << std::endl;
			hr = MAPI_E_INVALID_PARAMETER;
			goto Error;
		}

		this->pstrUserString = new tstring(_T("Default"));
	}
	else
	{
		enum
		{
			NAME,
			EID,
			NUM_RECIP_PROPS
		};

		CORg(lpSession->OpenAddressBook(NULL, NULL, MAPI_ACCESS_READ, &lpAdrBook));
		HrAllocAdrList(NUM_RECIP_PROPS,&lpAdrList);
		if (lpAdrList)
		{
			// Setup the One Time recipient by indicating how many recipients
			// and how many properties will be set on each recipient.
			lpAdrList->cEntries = 1;	// How many recipients.
			lpAdrList->aEntries[0].cValues = NUM_RECIP_PROPS; // How many properties per recipient

			// Set the SPropValue members == the desired values.
			lpAdrList->aEntries[0].rgPropVals[NAME].ulPropTag = PR_DISPLAY_NAME;
			lpAdrList->aEntries[0].rgPropVals[NAME].Value.LPSZ = (LPTSTR) this->pstrUserString;

			lpAdrList->aEntries[0].rgPropVals[EID].ulPropTag = PR_ENTRYID;

			CORg(lpAdrBook->ResolveName(
				0L,
				fMapiUnicode,
				NULL,
				lpAdrList));

			CORg(MAPIAllocateBuffer(sizeof(SBinary), (LPVOID*) &this->resolvedUserEID));
			ZeroMemory(this->resolvedUserEID, sizeof(SBinary));
			CORg(this->CopySBinary(this->resolvedUserEID, &lpAdrList->aEntries[0].rgPropVals[EID].Value.bin, NULL));
		}
	}

	// Resolve the rights string
	if (_tcsicmp(this->pstrRightsString->c_str(), _T("Owner")))
		this->rights = ROLE_OWNER;
	else if (_tcsicmp(this->pstrRightsString->c_str(), _T("PublishingEditor")))
		this->rights = ROLE_PUBLISH_EDITOR;
	else if (_tcsicmp(this->pstrRightsString->c_str(), _T("Editor")))
		this->rights = ROLE_EDITOR;
	else if (_tcsicmp(this->pstrRightsString->c_str(), _T("PublishingAuthor")))
		this->rights = ROLE_PUBLISH_AUTHOR;
	else if (_tcsicmp(this->pstrRightsString->c_str(), _T("Author")))
		this->rights = ROLE_AUTHOR;
	else if (_tcsicmp(this->pstrRightsString->c_str(), _T("NonEditingAuthor")))
		this->rights = ROLE_NONEDITING_AUTHOR;
	else if (_tcsicmp(this->pstrRightsString->c_str(), _T("Reviewer")))
		this->rights = ROLE_REVIEWER;
	else if (_tcsicmp(this->pstrRightsString->c_str(), _T("Contributor")))
		this->rights = ROLE_CONTRIBUTOR;
	else if (_tcsicmp(this->pstrRightsString->c_str(), _T("None")))
		this->rights = ROLE_NONE;
	else if (_tcsicmp(this->pstrRightsString->c_str(), _T("All")))
		this->rights = ROLE_OWNER | RIGHTS_FOLDER_CONTACT;

Cleanup:
	if (lpAdrList)
		FreePadrlist(lpAdrList);
	if (lpAdrBook)
		lpAdrBook->Release();
	if (lpAdrBook)
		lpAdrBook->Release();
	return hr;
Error:
	goto Cleanup;
}

void ModifyFolderPermissions::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
{
	HRESULT hr = S_OK;
	LPEXCHANGEMODIFYTABLE lpExchModTbl = NULL;
	LPMAPITABLE lpMapiTable = NULL;
	ULONG lpulRowCount = NULL;
	LPSRowSet pRows = NULL;
	SPropValue   prop[1] = {0};
	ROWLIST      rowList  = {0};
	UINT x = 0;

	tcout << _T("Modifying permissions on folder: ") << folderPath.c_str() << std::endl;

	CORg(folder->OpenProperty(PR_ACL_TABLE, &IID_IExchangeModifyTable, 0, MAPI_DEFERRED_ERRORS, (LPUNKNOWN*)&lpExchModTbl));
	CORg(lpExchModTbl->GetTable(0, &lpMapiTable));
	CORg(lpMapiTable->SetColumns((LPSPropTagArray)&rgAclTablePropTags, 0));
	CORg(HrQueryAllRows(lpMapiTable, NULL, NULL, NULL, NULL, &pRows));

	bool found = false;
	for (x = 0; x < pRows->cRows; x++)
	{
		if (this->resolvedUserEID == NULL)
		{
			// Must be Default or Anonymous, so match on the name
			if (pstrUserString->compare(pRows->aRow[x].lpProps[ePR_MEMBER_NAME].Value.lpszW) == 0)
			{
				found = true;
			}
		}
		else if (pRows->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb > 0)
		{
			if (IsEntryIdEqual(*resolvedUserEID, pRows->aRow[x].lpProps[ePR_MEMBER_ENTRYID].Value.bin))
			{
				found = true;
			}
		}

		if (found)
		{
			break;
		}
	}

	if (found)
	{
		if (this->remove)
		{
			prop[0].ulPropTag  = PR_MEMBER_ID;
			prop[0].Value.bin.cb = pRows -> aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.cb;
			prop[0].Value.bin.lpb = (BYTE*)pRows -> aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.lpb;
			rowList.cEntries    = 1;
			rowList.aEntries->ulRowFlags = ROW_REMOVE;
			rowList.aEntries->cValues  = 1;
			rowList.aEntries->rgPropVals = &prop[0];
    
			CORg(lpExchModTbl->ModifyTable(0, &rowList));
		}
		else
		{
			prop[0].ulPropTag  = PR_MEMBER_ID;
			prop[0].Value.bin.cb = pRows -> aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.cb;
			prop[0].Value.bin.lpb = (BYTE*)pRows -> aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.lpb;
			prop[1].ulPropTag  = PR_MEMBER_RIGHTS;
			prop[1].Value.l   = rights;

			rowList.cEntries    = 1;
			rowList.aEntries->ulRowFlags = ROW_MODIFY;
			rowList.aEntries->cValues  = 2;
			rowList.aEntries->rgPropVals = &prop[0];
     
			hr = lpExchModTbl->ModifyTable(0, &rowList);
		}
	}
	else
	{
		if (this->remove)
		{
			tcout << _T("     Specified user not present. No changes needed.") << std::endl;
		}
		else
		{
			prop[0].ulPropTag  = PR_MEMBER_ENTRYID;
			prop[0].Value.bin.cb = resolvedUserEID->cb;
			prop[0].Value.bin.lpb = resolvedUserEID->lpb;
			prop[1].ulPropTag  = PR_MEMBER_RIGHTS;
			prop[1].Value.l   = rights;

			rowList.cEntries = 1;
			rowList.aEntries->ulRowFlags = ROW_ADD;
			rowList.aEntries->cValues  = 2;
			rowList.aEntries->rgPropVals = &prop[0]; 

			hr = lpExchModTbl->ModifyTable(0, &rowList);
		}
	}

Cleanup:
	return;
Error:
	goto Cleanup;
}

// From MAPIABFunctions.cpp in MFCMapi
HRESULT ModifyFolderPermissions::HrAllocAdrList(ULONG ulNumProps, _Deref_out_opt_ LPADRLIST* lpAdrList)
{
	if (!lpAdrList || ulNumProps > ULONG_MAX/sizeof(SPropValue)) return MAPI_E_INVALID_PARAMETER;
	HRESULT hr = S_OK;
	LPADRLIST lpLocalAdrList = NULL;

	*lpAdrList = NULL;

	// Allocate memory for new SRowSet structure.
	CORg(MAPIAllocateBuffer(CbNewSRowSet(1),(LPVOID*) &lpLocalAdrList));

	if (lpLocalAdrList)
	{
		// Zero out allocated memory.
		ZeroMemory(lpLocalAdrList, CbNewSRowSet(1));

		// Allocate memory for SPropValue structure that indicates what
		// recipient properties will be set.
		CORg(MAPIAllocateBuffer(
			ulNumProps * sizeof(SPropValue),
			(LPVOID*) &lpLocalAdrList->aEntries[0].rgPropVals));

		// Zero out allocated memory.
		if (lpLocalAdrList->aEntries[0].rgPropVals)
			ZeroMemory(lpLocalAdrList->aEntries[0].rgPropVals,ulNumProps * sizeof(SPropValue));
		if (SUCCEEDED(hr))
		{
			*lpAdrList = lpLocalAdrList;
		}
		else
		{
			FreePadrlist(lpLocalAdrList);
		}
	}

Error:
	return hr;
} // HrAllocAdrList

ModifyFolderPermissions::~ModifyFolderPermissions(void)
{
}
