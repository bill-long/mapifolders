#include "ExportFolderPermissionsOperation.h"
#include "MAPIFolders.h"


ExportFolderPermissionsOperation::ExportFolderPermissionsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, Log *exportFile, bool useAdmin)
	:OperationBase(pstrBasePath, pstrMailbox, nScope, log, useAdmin)
{
	this->exportFile = exportFile;
}


ExportFolderPermissionsOperation::~ExportFolderPermissionsOperation()
{
}


HRESULT ExportFolderPermissionsOperation::Initialize(void)
{
	HRESULT hr = S_OK;
	CORg(OperationBase::Initialize());
Error:
	return hr;
}


void ExportFolderPermissionsOperation::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
{
	HRESULT hr = S_OK;
	LPEXCHANGEMODIFYTABLE lpExchModTbl = NULL;
	LPMAPITABLE lpMapiTable = NULL;
	ULONG lpulRowCount = NULL;
	LPSRowSet pRows = NULL;
	UINT x = 0;
	UINT y = 0;

	*pLog << "Exporting permissions for folder: " << folderPath.c_str() << "\n";

	CORg(folder->OpenProperty(PR_ACL_TABLE, &IID_IExchangeModifyTable, 0, MAPI_DEFERRED_ERRORS, (LPUNKNOWN*)&lpExchModTbl));
	CORg(lpExchModTbl->GetTable(0, &lpMapiTable));
	CORg(lpMapiTable->SetColumns((LPSPropTagArray)&rgAclTablePropTags, 0));
	CORg(HrQueryAllRows(lpMapiTable, NULL, NULL, NULL, NULL, &pRows));

	*exportFile << "SETACL\t" << folderPath.c_str();

	for (x = 0; x < pRows->cRows; x++)
	{
		SRow thisRow = pRows->aRow[x];
		LONG rights = thisRow.lpProps[ePR_MEMBER_RIGHTS].Value.l;
		tstring rightsString = _T("");

		// Check for a role
		if ((rights & ROLE_OWNER) == ROLE_OWNER)
		{
			if (rights & RIGHTS_FOLDER_CONTACT)
			{
				rightsString.append(_T("All "));
				rights &= ~RIGHTS_FOLDER_CONTACT;
			}
			else
			{
				rightsString.append(_T("Owner "));
			}

			rights &= ~ROLE_OWNER;
			rights &= ~RIGHTS_EDIT_OWN;
			rights &= ~RIGHTS_DELETE_OWN;
		}
		else if ((rights & ROLE_PUBLISH_EDITOR) == ROLE_PUBLISH_EDITOR)
		{
			rightsString.append(_T("PublishingEditor "));
			rights &= ~ROLE_PUBLISH_EDITOR;
			rights &= ~RIGHTS_EDIT_OWN;
			rights &= ~RIGHTS_DELETE_OWN;
		}
		else if ((rights & ROLE_EDITOR) == ROLE_EDITOR)
		{
			rightsString.append(_T("Editor "));
			rights &= ~ROLE_EDITOR;
			rights &= ~RIGHTS_EDIT_OWN;
			rights &= ~RIGHTS_DELETE_OWN;
		}
		else if ((rights & ROLE_PUBLISH_AUTHOR) == ROLE_PUBLISH_AUTHOR)
		{
			rightsString.append(_T("PublishingAuthor "));
			rights &= ~ROLE_PUBLISH_AUTHOR;
		}
		else if ((rights & ROLE_AUTHOR) == ROLE_AUTHOR)
		{
			rightsString.append(_T("Author "));
			rights &= ~ROLE_AUTHOR;
		}
		else if ((rights & ROLE_NONEDITING_AUTHOR) == ROLE_NONEDITING_AUTHOR)
		{
			rightsString.append(_T("NonEditingAuthor "));
			rights &= ~ROLE_NONEDITING_AUTHOR;
		}
		else if ((rights & ROLE_REVIEWER) == ROLE_REVIEWER)
		{
			rightsString.append(_T("Reviewer "));
			rights &= ~ROLE_REVIEWER;
		}
		else if ((rights & ROLE_CONTRIBUTOR) == ROLE_CONTRIBUTOR)
		{
			rightsString.append(_T("Contributor "));
			rights &= ~ROLE_CONTRIBUTOR;
		}
		else if ((rights & ROLE_NONE) == ROLE_NONE)
		{
			rightsString.append(_T("Visible "));
			rights &= ~ROLE_NONE;
		}

		// If we didn't find a role, look at the individual rights
		if (rights & RIGHTS_EDIT_OWN)
		{
			rightsString.append(_T("WriteOwn "));
			rights &= ~RIGHTS_EDIT_OWN;
		}

		if (rights & RIGHTS_EDIT_ALL)
		{
			rightsString.append(_T("Write "));
			rights &= ~RIGHTS_EDIT_ALL;
		}

		if (rights & RIGHTS_DELETE_OWN)
		{
			rightsString.append(_T("DeleteOwn "));
			rights &= ~RIGHTS_DELETE_OWN;
		}

		if (rights & RIGHTS_DELETE_ALL)
		{
			rightsString.append(_T("Delete "));
			rights &= ~RIGHTS_DELETE_ALL;
		}

		if (rights & RIGHTS_READ_ITEMS)
		{
			rightsString.append(_T("Read "));
			rights &= ~RIGHTS_READ_ITEMS;
		}

		if (rights & RIGHTS_CREATE_ITEMS)
		{
			rightsString.append(_T("Create "));
			rights &= ~RIGHTS_CREATE_ITEMS;
		}

		if (rights & RIGHTS_CREATE_SUBFOLDERS)
		{
			rightsString.append(_T("CreateSubfolder "));
			rights &= ~RIGHTS_CREATE_SUBFOLDERS;
		}

		if (rights & RIGHTS_FOLDER_OWNER)
		{
			rightsString.append(_T("o "));
			rights &= ~RIGHTS_FOLDER_OWNER;
		}

		if (rights & RIGHTS_FOLDER_CONTACT)
		{
			rightsString.append(_T("Contact "));
			rights &= ~RIGHTS_FOLDER_CONTACT;
		}

		// We should never actually hit this since ROLE_NONE == RIGHTS_FOLDER_VISIBLE
		if (rights & RIGHTS_FOLDER_VISIBLE)
		{
			rightsString.append(_T("Visible "));
			rights &= ~RIGHTS_FOLDER_VISIBLE;
		}

		// At this point there should be nothing left in rights
		if (rights != 0)
		{
			*pLog << _T("    Error: Unexpected rights encountered.\n");
			hr = E_FAIL;
			goto Error;
		}

		if (rightsString.size() == 0)
		{
			// We found no rights at all
			rightsString.append(_T("None "));
		}

		rightsString.resize(rightsString.size() - 1);

		tstring nameString;
		if (thisRow.lpProps[ePR_MEMBER_ID].Value.cur.Lo == 0xFFFFFFFF && thisRow.lpProps[ePR_MEMBER_ID].Value.cur.Hi == 0xFFFFFFFF)
		{
			// This is Anonymous
			nameString = _T("Anonymous");
		}
		else if (thisRow.lpProps[ePR_MEMBER_ID].Value.cur.Lo == 0x0 && thisRow.lpProps[ePR_MEMBER_ID].Value.cur.Hi == 0x0)
		{
			// This is Everyone
			nameString = _T("Default");
		}
		else
		{
			// We need to resolve the Entry ID to a leg DN.
			ULONG objType = 0;
			LPMAPIPROP resolvedEIDProps = NULL;
			
			if (lpAdrBook == NULL)
			{
				CORg(this->lpSession->OpenAddressBook(NULL, NULL, NULL, &lpAdrBook));
			}
			
			hr = this->lpAdrBook->OpenEntry(
				thisRow.lpProps[ePR_MEMBER_ENTRYID].Value.bin.cb,
				(LPENTRYID)thisRow.lpProps[ePR_MEMBER_ENTRYID].Value.bin.lpb,
				NULL, MAPI_BEST_ACCESS, &objType, (LPUNKNOWN *)&resolvedEIDProps);

			if (hr != S_OK)
			{
				// If we couldn't resolve it just output whatever name we have
				nameString = thisRow.lpProps[ePR_MEMBER_NAME].Value.LPSZ;
				*pLog << _T("    Couldn't resolve EID for: ") << nameString.c_str() << "\n";
			}
			else
			{
				LPSPropValue legDN = NULL;
				hr = HrGetOneProp(resolvedEIDProps, PR_EMAIL_ADDRESS, &legDN);
				if (hr != S_OK)
				{
					// If we couldn't resolve it just output whatever name we have
					nameString = thisRow.lpProps[ePR_MEMBER_NAME].Value.LPSZ;
					*pLog << _T("    Couldn't read PR_EMAIL_ADDRESS for: ") << nameString.c_str() << "\n";
				}
				else
				{
					nameString = legDN->Value.LPSZ;
				}
			}
		}

		// At this point we have both a name and a rights string, so write them out
		*exportFile << "\t" << nameString.c_str() << "\t" << rightsString.c_str();
	}

	*exportFile << "\n";

Cleanup:
	return;
Error:
	goto Cleanup;
}