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

	CORg(OperationBase::Initialize());
	
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
	else if (_tstricmp(strDefault.c_str(), this->pstrUserString->c_str()) == 0)
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
		this->resolvedUserEID = this->ResolveNameToEID(this->pstrUserString);
		if (this->resolvedUserEID == NULL)
		{
			hr = MAPI_E_NOT_FOUND;
			tcout << _T("Could not resolve the specified user name.") << std::endl;
			goto Error;
		}
	}

	// Resolve the rights string
	if (this->pstrRightsString)
	{
		if (_tcsicmp(this->pstrRightsString->c_str(), _T("Owner")) == 0)
			this->rights = ROLE_OWNER;
		else if (_tcsicmp(this->pstrRightsString->c_str(), _T("PublishingEditor")) == 0)
			this->rights = ROLE_PUBLISH_EDITOR;
		else if (_tcsicmp(this->pstrRightsString->c_str(), _T("Editor")) == 0)
			this->rights = ROLE_EDITOR;
		else if (_tcsicmp(this->pstrRightsString->c_str(), _T("PublishingAuthor")) == 0)
			this->rights = ROLE_PUBLISH_AUTHOR;
		else if (_tcsicmp(this->pstrRightsString->c_str(), _T("Author")) == 0)
			this->rights = ROLE_AUTHOR;
		else if (_tcsicmp(this->pstrRightsString->c_str(), _T("NonEditingAuthor")) == 0)
			this->rights = ROLE_NONEDITING_AUTHOR;
		else if (_tcsicmp(this->pstrRightsString->c_str(), _T("Reviewer")) == 0)
			this->rights = ROLE_REVIEWER;
		else if (_tcsicmp(this->pstrRightsString->c_str(), _T("Contributor")) == 0)
			this->rights = ROLE_CONTRIBUTOR;
		else if (_tcsicmp(this->pstrRightsString->c_str(), _T("None")) == 0)
			this->rights = ROLE_NONE;
		else if (_tcsicmp(this->pstrRightsString->c_str(), _T("All")) == 0)
			this->rights = ROLE_OWNER | RIGHTS_FOLDER_CONTACT;
	}

Cleanup:
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
			if (_tcsicmp(pstrUserString->c_str(), pRows->aRow[x].lpProps[ePR_MEMBER_NAME].Value.lpszW) == 0)
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
			SPropValue   prop[1] = {0};
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
			SPropValue   prop[2] = {0};
			prop[0].ulPropTag  = PR_MEMBER_ID;
			prop[0].Value.bin.cb = pRows -> aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.cb;
			prop[0].Value.bin.lpb = (BYTE*)pRows -> aRow[x].lpProps[ePR_MEMBER_ID].Value.bin.lpb;
			prop[1].ulPropTag  = PR_MEMBER_RIGHTS;
			prop[1].Value.l   = rights;

			rowList.cEntries    = 1;
			rowList.aEntries->ulRowFlags = ROW_MODIFY;
			rowList.aEntries->cValues  = 2;
			rowList.aEntries->rgPropVals = &prop[0];
     
			CORg(lpExchModTbl->ModifyTable(0, &rowList));
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
			SPropValue   prop[2] = {0};
			prop[0].ulPropTag  = PR_MEMBER_ENTRYID;
			prop[0].Value.bin.cb = resolvedUserEID->cb;
			prop[0].Value.bin.lpb = resolvedUserEID->lpb;
			prop[1].ulPropTag  = PR_MEMBER_RIGHTS;
			prop[1].Value.l   = rights;

			rowList.cEntries = 1;
			rowList.aEntries->ulRowFlags = ROW_ADD;
			rowList.aEntries->cValues  = 2;
			rowList.aEntries->rgPropVals = &prop[0]; 

			CORg(lpExchModTbl->ModifyTable(0, &rowList));
		}
	}

Cleanup:
	if (pRows)
		FreeProws(pRows);
	if (lpMapiTable)
		lpMapiTable->Release();
	if (lpExchModTbl)
		lpExchModTbl->Release();

	return;
Error:
	goto Cleanup;
}



ModifyFolderPermissions::~ModifyFolderPermissions(void)
{
}
