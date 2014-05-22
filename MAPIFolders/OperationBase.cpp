#include "stdafx.h"
#include "OperationBase.h"
#include "MAPIFolders.h"
#include <algorithm>

OperationBase::OperationBase(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log)
{
	this->lpAdminMDB = NULL;
	this->strBasePath = pstrBasePath;
	this->strMailbox = pstrMailbox;
	this->nScope = nScope;
	this->pLog = log;
	this->lpAdrBook = NULL;
}

OperationBase::~OperationBase(void)
{
}


void OperationBase::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
{
	return;
}

std::string OperationBase::GetStringFromFolderPath(LPMAPIFOLDER folder)
{
	HRESULT hr = NULL;
	SPropTagArray rgPropTag = { 1, PR_FOLDER_PATHNAME };
	LPSPropValue lpPropValue = NULL;
	ULONG cValues = 0;
	LPCSTR lpszFolderPath = NULL;
	LPCSTR lpszReturnFolderPath = NULL;

	CORg(folder->GetProps(&rgPropTag, NULL, &cValues, &lpPropValue));
	if (lpPropValue && lpPropValue->ulPropTag == PR_FOLDER_PATHNAME)
	{
		lpszFolderPath = lpPropValue->Value.lpszA;
		std::string folderPath(lpszFolderPath);
		std::replace(folderPath.begin(), folderPath.end(), '?', '\\');
		return folderPath;
	}

Error:
	return NULL;
}

HRESULT OperationBase::Initialize(void)
{
	HRESULT hr = S_OK;
	lpAdminMDB = NULL;
	lpRootFolder = NULL;
	lpStartingFolder = NULL;
	lpSession = NULL;
	startingPath = tstring(_T(""));

	MAPIINIT_0  MAPIINIT= { 0, MAPI_MULTITHREAD_NOTIFICATIONS};

	CORg(MAPIInitialize (&MAPIINIT));

	CORg(MAPILogonEx(0, NULL, NULL, MAPI_LOGON_UI | MAPI_EXTENDED | MAPI_ALLOW_OTHERS | MAPI_LOGON_UI | MAPI_EXPLICIT_PROFILE, &lpSession));
		
	if (strMailbox != NULL)
	{
		lpRootFolder = GetMailboxRoot(lpSession);
	}
	else
	{
		lpRootFolder = GetPFRoot(lpSession);
	}
	if (lpRootFolder == NULL)
	{
		*pLog << "Failed to get root folder" << "\n";
		hr = MAPI_E_NOT_FOUND;
		goto Error;
	}

	lpStartingFolder = GetStartingFolder(lpSession, &startingPath);
	if (lpStartingFolder == NULL)
	{
		*pLog << "Failed to find the specified folder" << "\n";
		hr = MAPI_E_NOT_FOUND;
		goto Error;
	}

Cleanup:
	// Don't release anything here, we need this stuff to run the operation
	return hr;
Error:
	goto Cleanup;
}

void OperationBase::DoOperation()
{
	TraverseFolders(lpSession, lpStartingFolder, startingPath);

	// Now we can release stuff
	if (lpRootFolder)
		lpRootFolder->Release();
	if (lpAdminMDB)
		lpAdminMDB->Release();
	if (lpSession)
		lpSession->Release();
}

LPMAPIFOLDER OperationBase::GetMailboxRoot(IMAPISession *pSession)
{
	HRESULT hr = S_OK;
	SBinary mailboxEID = {0};
	LPMAPIFOLDER lpRoot = NULL;
	LPADRLIST lpAdrList = NULL;
	LPMDB lpDefaultMDB = NULL;
	LPEXCHANGEMANAGESTORE lpXManageStore = NULL;
	LPTSTR szServerName = NULL;
	LPCTSTR	szServerNamePTR = NULL;
	LPTSTR lpszMsgStoreDN = NULL;

	*pLog << "Resolving mailbox: " << this->strMailbox->c_str() << "\n";

	enum
	{
		NAME,
		NUM_RECIP_PROPS
	};

	CORg(lpSession->OpenAddressBook(NULL, NULL, NULL, &lpAdrBook));
	HrAllocAdrList(NUM_RECIP_PROPS,&lpAdrList);
	if (lpAdrList)
	{
		lpAdrList->cEntries = 1;	// How many recipients.
		lpAdrList->aEntries[0].cValues = NUM_RECIP_PROPS; // How many properties per recipient

		// Set the SPropValue members == the desired values.
		lpAdrList->aEntries[0].rgPropVals[NAME].ulPropTag = PR_DISPLAY_NAME;
		lpAdrList->aEntries[0].rgPropVals[NAME].Value.LPSZ = (LPTSTR) this->strMailbox->c_str();

		CORg(lpAdrBook->ResolveName(
			0L,
			fMapiUnicode,
			NULL,
			lpAdrList));

		bool foundEXAddressType = false;
		LPTSTR emailAddress = NULL;
		for (UINT x = 0; x < lpAdrList->aEntries[0].cValues; x++)
		{
			SPropValue thisProp = lpAdrList->aEntries[0].rgPropVals[x];
			if (thisProp.ulPropTag == PR_ADDRTYPE)
			{
				*pLog << "    PR_ADDRTYPE: " << thisProp.Value.LPSZ << "\n";
				if (tstring(_T("EX")).compare(thisProp.Value.LPSZ) == 0)
				{
					foundEXAddressType = true;
				}
			}
			else if (thisProp.ulPropTag == PR_DISPLAY_NAME)
			{
				*pLog << "    PR_DISPLAY_NAME: " << thisProp.Value.LPSZ << "\n";
			}
			else if (thisProp.ulPropTag == PR_EMAIL_ADDRESS)
			{
				*pLog << "    PR_EMAIL_ADDRESS: " << thisProp.Value.LPSZ << "\n";
				emailAddress = thisProp.Value.LPSZ;
			}
		}

		if (!foundEXAddressType || emailAddress == NULL)
		{
			hr = MAPI_E_NOT_FOUND;
			goto Error;
		}

		CORg(OpenDefaultMessageStore(lpSession, &lpDefaultMDB));

		CORg(lpDefaultMDB->QueryInterface(
		IID_IExchangeManageStore,
		(LPVOID*) &lpXManageStore));

		CORg(GetServerName(lpSession, &szServerName));
		szServerNamePTR = szServerName;

		CORg(BuildServerDN(
			szServerNamePTR,
			_T("/cn=Microsoft Private MDB"), // STRING_OK
			&lpszMsgStoreDN));

#ifdef UNICODE
		{
			char *szAnsiMsgStoreDN = NULL;
			char *szAnsiMailboxDN = NULL;
			CORg(UnicodeToAnsi(lpszMsgStoreDN,&szAnsiMsgStoreDN,-1));
			CORg(UnicodeToAnsi(emailAddress,&szAnsiMailboxDN,-1));

			CORg(lpXManageStore->CreateStoreEntryID(
				szAnsiMsgStoreDN,
				szAnsiMailboxDN,
				OPENSTORE_TAKE_OWNERSHIP | OPENSTORE_USE_ADMIN_PRIVILEGE,
				&mailboxEID.cb,
				(LPENTRYID*) &mailboxEID.lpb));
			delete[] szAnsiMsgStoreDN;
			delete[] szAnsiMailboxDN;
		}
#else
		EC_MAPI(lpXManageStore->CreateStoreEntryID(
			(LPSTR) lpszMsgStoreDN,
			(LPSTR) lpszMailboxDN,
			ulFlags,
			&mailboxEID->cb,
			(LPENTRYID*) &mailboxEID->lpb));
#endif
	}

	if (mailboxEID.cb == 0)
	{
		*pLog << _T("Could not resolve mailbox.") << "\n";
		hr = MAPI_E_NOT_FOUND;
		goto Error;
	}

	CORg(CallOpenMsgStore(lpSession, NULL, &mailboxEID,
		MDB_NO_DIALOG |
		MDB_NO_MAIL |     // spooler not notified of our presence
		MDB_TEMPORARY |   // message store not added to MAPI profile
		MAPI_BEST_ACCESS, // normally WRITE, but allow access to RO store
		&lpAdminMDB));

	ULONG ulObjType = NULL;
	CORg(lpAdminMDB->OpenEntry(NULL, NULL, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *) &lpRoot));

Cleanup:
	if (lpAdrList)
		FreePadrlist(lpAdrList);
	return lpRoot;
Error:
	goto Cleanup;
}

LPMAPIFOLDER OperationBase::GetPFRoot(IMAPISession *pSession)
{
	bool foundPFStore = FALSE;
	HRESULT hr = S_OK;
	CComPtr<IMAPITable> spTable;
	SRowSet *pmrows = NULL;
	SBinary publicEntryID;
	SBinary adminEntryID = {0};
	LPMAPIFOLDER lpRoot = NULL;
	LPMDB lpMDB = NULL;
	LPEXCHANGEMANAGESTORE lpXManageStore = NULL;
	LPSPropValue lpServerName = NULL;
	LPTSTR	szServerDN = NULL;

	enum MAPIColumns
	{
		COL_ENTRYID = 0,
		COL_DISPLAYNAME,
		COL_MDB_PROVIDER,
		cCols				// End marker
	};

	static const SizedSPropTagArray(cCols, mcols) = 
	{	
		cCols,
		{
			PR_ENTRYID,
			PR_DISPLAY_NAME,
			PR_MDB_PROVIDER
		}
	};

	CORg(pSession->GetMsgStoresTable(0, &spTable));

	CORg(spTable->SetColumns((LPSPropTagArray)&mcols, TBL_BATCH));
	CORg(spTable->SeekRow(BOOKMARK_BEGINNING, 0, 0));

	CORg(spTable->QueryRows(50, 0, &pmrows));

	*pLog << "Found " << pmrows->cRows << " stores in MAPI profile:" << "\n";
	for (UINT i=0; i != pmrows->cRows; i++)
	{
		SRow *prow = pmrows->aRow + i;
		LPCTSTR pwz = NULL;
		if (PR_DISPLAY_NAME == prow->lpProps[COL_DISPLAYNAME].ulPropTag)
			pwz = prow->lpProps[COL_DISPLAYNAME].Value.LPSZ;
		if (pwz)
			*pLog << " " << std::setw(4) << i << ": " << (pwz ? pwz : _T("<NULL>")) << "\n";

		if (IsEqualMAPIUID(prow->lpProps[COL_MDB_PROVIDER].Value.bin.lpb,pbExchangeProviderPublicGuid))
		{
			publicEntryID = prow->lpProps[COL_ENTRYID].Value.bin;
			foundPFStore = TRUE;
		}
	}

	if (!foundPFStore)
	{
		*pLog << "No public folder database in MAPI profile." << "\n";
		goto Error;
	}

	if (publicEntryID.cb < 1)
	{
		*pLog << "Could not get public folder store entry ID." << "\n";
		goto Error;
	}

	*pLog << "Opening public folders..." << "\n";
	CORg(pSession->OpenMsgStore(NULL, publicEntryID.cb, (LPENTRYID)publicEntryID.lpb, NULL,
		MAPI_BEST_ACCESS | MDB_ONLINE, &lpMDB));

	CORg(HrGetOneProp(
			lpMDB,
			PR_HIERARCHY_SERVER,
			&lpServerName));

	*pLog << "Using public folder server/mailbox: " << lpServerName->Value.LPSZ << "\n";
	LPTSTR szServerName = lpServerName->Value.LPSZ;
	LPCTSTR szServerNamePTR = szServerName;
	CORg(BuildServerDN(
		szServerNamePTR,
				_T("/cn=Microsoft Public MDB"),
				&szServerDN));

	CORg(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(LPVOID*) &lpXManageStore));

#ifdef UNICODE
		{
			char *szAnsiMsgStoreDN = NULL;
			CORg(UnicodeToAnsi(szServerDN,&szAnsiMsgStoreDN,-1));

			CORg(lpXManageStore->CreateStoreEntryID(
				szAnsiMsgStoreDN,
				NULL,
				OPENSTORE_PUBLIC | OPENSTORE_USE_ADMIN_PRIVILEGE,
				&adminEntryID.cb,
				(LPENTRYID*) &adminEntryID.lpb));
			delete[] szAnsiMsgStoreDN;
		}
#else
		EC_MAPI(lpXManageStore->CreateStoreEntryID(
			(LPSTR) lpszMsgStoreDN,
			(LPSTR) lpszMailboxDN,
			ulFlags,
			&mailboxEID->cb,
			(LPENTRYID*) &mailboxEID->lpb));
#endif

	CORg(pSession->OpenMsgStore(NULL, adminEntryID.cb, (LPENTRYID)adminEntryID.lpb, NULL,
		MAPI_BEST_ACCESS | MDB_ONLINE, &lpAdminMDB));

	ULONG ulObjType = NULL;
	CORg(lpAdminMDB->OpenEntry(NULL, NULL, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *) &lpRoot));

Error:
	if (pmrows)
		FreeProws(pmrows);
	if (szServerDN)
		MAPIFreeBuffer(szServerDN);

	return lpRoot;
}

LPMAPIFOLDER OperationBase::GetStartingFolder(IMAPISession *pSession, tstring *calculatedFolderPath)
{
	SRowSet *pmrows = NULL;
	LPMAPITABLE hierarchyTable = NULL;
	HRESULT hr = NULL;
	ULONG objType = NULL;
	tstring userSpecifiedPath;
	LPMAPIFOLDER lpTopFolder = NULL; // will be IPM_SUBTREE or NON_IPM_SUBTREE, do NOT release
	LPMAPIFOLDER lpFolderToReturn = NULL;
	ULONG ulObjType = NULL;
	std::vector<tstring> splitPath;
	UINT startIndex = 1;

	if (this->strMailbox != NULL)
	{
		lpTopFolder = lpRootFolder;
		if (this->strBasePath == NULL)
		{
			LPSPropValue lpPropVal = NULL;
			CORg(HrGetOneProp(lpAdminMDB, PR_IPM_SUBTREE_ENTRYID, &lpPropVal));
			ULONG objType = NULL;
			LPMAPIFOLDER ipmSubtreeFolder = NULL;
			CORg(lpAdminMDB->OpenEntry(lpPropVal->Value.bin.cb, (LPENTRYID)lpPropVal->Value.bin.lpb, NULL, MAPI_BEST_ACCESS, &objType, (LPUNKNOWN *)&ipmSubtreeFolder));
			lpPropVal = NULL;
			CORg(HrGetOneProp(ipmSubtreeFolder, PR_DISPLAY_NAME, &lpPropVal));
			userSpecifiedPath = _T("\\");
			userSpecifiedPath.append(lpPropVal->Value.LPSZ);
		}
		else
			userSpecifiedPath = *this->strBasePath;

		splitPath = Split(userSpecifiedPath, (_T('\\')));
		startIndex = 1;
	}
	else
	{
		if (this->strBasePath == NULL)
		{
			// If user didn't specify a path, assume IPM_SUBTREE
			userSpecifiedPath = (_T("\\IPM_SUBTREE"));
		}
		else
		{
			userSpecifiedPath = *this->strBasePath;
		}

		// Split the specified path using backslash separator
		splitPath = Split(userSpecifiedPath, (_T('\\')));

		// We must have at least two items
		if (splitPath.size() < 2)
		{
			goto Error;
		}

		// The first item must be empty, which indicates we had a leading backslash
		if (splitPath.at(0).length() != 0)
		{
			goto Error;
		}

		// The second item could be IPM_SUBTREE, NON_IPM_SUBTREE, or a folder name
		// If a folder name, assume IPM_SUBTREE
		if (splitPath.at(1).length() > 0)
		{
			tstring returnedFolderName = _T("");
			if (_tstricmp(splitPath.at(1).c_str(), _T("non_ipm_subtree")) == 0)
			{
				CORg(GetSubfolderByName(lpRootFolder, _T("NON_IPM_SUBTREE"), &lpTopFolder, &returnedFolderName));
				calculatedFolderPath->append(_T("\\NON_IPM_SUBTREE"));
			}
			else
			{
				CORg(GetSubfolderByName(lpRootFolder, _T("IPM_SUBTREE"), &lpTopFolder, &returnedFolderName));
				calculatedFolderPath->append(_T("\\IPM_SUBTREE"));
			}
		}

		// At this point we should know where we're starting to look for the
		// specified folder.
		if (lpTopFolder == NULL)
		{
			goto Error;
		}

		// If the first item is a root folder, then skip it, because
		// we already got that far.
		if ((_tstricmp(splitPath.at(1).c_str(), _T("non_ipm_subtree")) == 0) ||
			(_tstricmp(splitPath.at(1).c_str(), _T("ipm_subtree")) == 0))
		{
			startIndex = 2;
		}
	}

	// If startIndex points to the 3rd item, but we only have 2 because
	// a root folder was specified as the starting folder, then we're
	// already done. Or if this is a mailbox and no path was specified, we're done.
	if (startIndex >= splitPath.size())
	{
		lpFolderToReturn = lpTopFolder;
	}
	else
	{
		// We're going to iterate through the path and open each folder.
		// Be sure to release each one when done with it
		LPMAPIFOLDER returnedFolder = NULL;
		tstring realFolderName = _T("");
		CORg(GetSubfolderByName(lpTopFolder, splitPath.at(startIndex), &returnedFolder, &realFolderName));
		if (returnedFolder == NULL)
		{
			*pLog << "Could not find folder: " << splitPath.at(startIndex) << "\n";
			goto Error;
		}

		calculatedFolderPath->append(_T("\\"));
		calculatedFolderPath->append(realFolderName);

		startIndex++;
		for (UINT x = startIndex; x < splitPath.size(); x++)
		{
			tstring returnedFolderName = _T("");
			LPMAPIFOLDER parentFolder = returnedFolder;
			CORg(GetSubfolderByName(parentFolder, splitPath.at(x), &returnedFolder, &returnedFolderName));
			if (returnedFolder == NULL || returnedFolderName.length() < 1)
			{
				*pLog << "Could not find folder: " << splitPath.at(x) << "\n";
				parentFolder->Release();
				goto Error;
			}
			else
			{
				calculatedFolderPath->append(_T("\\"));
				calculatedFolderPath->append(returnedFolderName);
				parentFolder->Release();
			}
		}

		lpFolderToReturn = returnedFolder;
	}

Cleanup:
	return lpFolderToReturn;
Error:
	lpFolderToReturn = NULL;
	goto Cleanup;
}

HRESULT OperationBase::GetSubfolderByName(LPMAPIFOLDER parentFolder, tstring folderNameToFind, LPMAPIFOLDER *returnedFolder, tstring *returnedFolderName)
{
	HRESULT hr = S_OK;
	LPMAPITABLE hierarchyTable = NULL;
	SRowSet *pmrows = NULL;
	ULONG ulObjType = NULL;
	returnedFolderName->clear();

	enum MAPIColumns
	{
		COL_ENTRYID = 0,
		COL_DISPLAYNAME_W,
		cCols				// End marker
	};

	static const SizedSPropTagArray(cCols, mcols) = 
	{	
		cCols,
		{
			PR_ENTRYID,
			PR_DISPLAY_NAME_W
		}
	};

	CORg(parentFolder->GetHierarchyTable(NULL, &hierarchyTable));
	CORg(hierarchyTable->SetColumns((LPSPropTagArray)&mcols, TBL_BATCH));
	CORg(hierarchyTable->SeekRow(BOOKMARK_BEGINNING, 0, 0));
	CORg(hierarchyTable->QueryRows(10000, 0, &pmrows));
	
	for (UINT i=0; i != pmrows->cRows; i++)
	{
		SRow *prow = pmrows->aRow + i;
		LPCWSTR pwz = NULL;
		if (PR_DISPLAY_NAME_W == prow->lpProps[COL_DISPLAYNAME_W].ulPropTag)
		{
			pwz = prow->lpProps[COL_DISPLAYNAME_W].Value.lpszW;
			if (_tstricmp(folderNameToFind.c_str(), pwz) == 0)
			{
				CORg(this->lpAdminMDB->OpenEntry(prow->lpProps[COL_ENTRYID].Value.bin.cb, (LPENTRYID)prow->lpProps[COL_ENTRYID].Value.bin.lpb, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *) returnedFolder));
				returnedFolderName->append(pwz);
				break;
			}
		}
	}

Cleanup:
	if (pmrows)
		FreeProws(pmrows);
	if (hierarchyTable)
		hierarchyTable->Release();
	// The caller should free the folders

	return hr;
Error:
	goto Cleanup;
}

void OperationBase::TraverseFolders(CComPtr<IMAPISession> session, LPMAPIFOLDER baseFolder, tstring parentPath)
{
	SRowSet *pmrows = NULL;
	LPMAPITABLE hierarchyTable = NULL;
	HRESULT hr = NULL;
	ULONG objType = NULL;

	if ((this->nScope == UserArgs::ActionScope::ONELEVEL) ||
		(this->nScope == UserArgs::ActionScope::SUBTREE))
	{
		enum MAPIColumns
		{
			COL_ENTRYID = 0,
			COL_DISPLAYNAME,
			cCols				// End marker
		};

		static const SizedSPropTagArray(cCols, mcols) = 
		{	
			cCols,
			{
				PR_ENTRYID,
				PR_DISPLAY_NAME
			}
		};

RetryGetHierarchyTable:
		hr = baseFolder->GetHierarchyTable(NULL, &hierarchyTable);
		if (FAILED(hr))
		{
			if (hr == MAPI_E_TIMEOUT)
			{
				*pLog << "     Encountered a timeout in GetHierarchyTable. Retrying in 5 seconds..." << "\n";
				Sleep(5000);
				goto RetryGetHierarchyTable;
			}
			else if (hr == MAPI_E_NETWORK_ERROR)
			{
				*pLog << "     Encountered a network error in GetHierarchyTable. Retrying in 5 seconds..." << "\n";
				Sleep(5000);
				goto RetryGetHierarchyTable;
			}
			else
			{
				*pLog << "     GetHierarchyTable returned an error. hr = " << std::hex << hr << "\n";
				goto Error;
			}
		}

		CORg(hierarchyTable->SetColumns((LPSPropTagArray)&mcols, TBL_BATCH));
		CORg(hierarchyTable->SeekRow(BOOKMARK_BEGINNING, 0, 0));
	
RetryQueryRows:
		hr = hierarchyTable->QueryRows(10000, 0, &pmrows);
		if (FAILED(hr))
		{
			if (hr == MAPI_E_TIMEOUT)
			{
				*pLog << "     Encountered a timeout in QueryRows. Retrying in 5 seconds..." << "\n";
				pmrows = NULL;
				Sleep(5000);
				goto RetryQueryRows;
			}
			else if (hr == MAPI_E_NETWORK_ERROR)
			{
				*pLog << "     Encountered a network error in QueryRows. Retrying in 5 seconds..." << "\n";
				pmrows = NULL;
				Sleep(5000);
				goto RetryQueryRows;
			}
			else
			{
				*pLog << "     QueryRows returned an error. hr = " << std::hex << hr << "\n";
				goto Error;
			}
		}

		for (UINT i=0; i != pmrows->cRows; i++)
		{
			SRow *prow = pmrows->aRow + i;
			LPCWSTR pwz = NULL;
			if (PR_DISPLAY_NAME == prow->lpProps[COL_DISPLAYNAME].ulPropTag)
				pwz = prow->lpProps[COL_DISPLAYNAME].Value.LPSZ;

			tstring thisPath(parentPath.c_str());
			thisPath.append(_T("\\"));
			thisPath.append(pwz);

			ULONG ulObjType = NULL;
			LPMAPIFOLDER lpSubfolder = NULL;

RetryOpenEntry:
			hr = lpAdminMDB->OpenEntry(prow->lpProps[COL_ENTRYID].Value.bin.cb, (LPENTRYID)prow->lpProps[COL_ENTRYID].Value.bin.lpb, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *) &lpSubfolder);
			if (FAILED(hr))
			{
				if (hr == MAPI_E_TIMEOUT)
				{
					*pLog << "     Encountered a timeout in OpenEntry. Retrying in 5 seconds..." << "\n";
					Sleep(5000);
					goto RetryOpenEntry;
				}
				else if (hr == MAPI_E_NETWORK_ERROR)
				{
					*pLog << "     Encountered a network error in OpenEntry. Retrying in 5 seconds..." << "\n";
					Sleep(5000);
					goto RetryOpenEntry;
				}
				else
				{
					*pLog << "     OpenEntry returned an error. hr = " << std::hex << hr << "\n";
				}
			}

			if (lpSubfolder)
			{
				this->ProcessFolder(lpSubfolder, thisPath);

				if (this->nScope == UserArgs::ActionScope::SUBTREE)
				{
					TraverseFolders(session, lpSubfolder, thisPath);
				}

				lpSubfolder->Release();
			}
		}
	}
	else
	{
		// Scope is Base
		this->ProcessFolder(baseFolder, parentPath);
	}

Error:
	if (pmrows)
		FreeProws(pmrows);
	if (hierarchyTable)
		hierarchyTable->Release();

	return;
}

LPSBinary OperationBase::ResolveNameToEID(tstring *pstrName)
{
	HRESULT hr = S_OK;
	LPADRBOOK lpAdrBook = NULL;
	LPADRLIST lpAdrList = NULL;
	LPSBinary lpEID = NULL;

	enum
	{
		NAME,
		NUM_RECIP_PROPS
	};

	CORg(lpSession->OpenAddressBook(NULL, NULL, NULL, &lpAdrBook));
	HrAllocAdrList(NUM_RECIP_PROPS,&lpAdrList);
	if (lpAdrList)
	{
		// Setup the One Time recipient by indicating how many recipients
		// and how many properties will be set on each recipient.
		lpAdrList->cEntries = 1;	// How many recipients.
		lpAdrList->aEntries[0].cValues = NUM_RECIP_PROPS; // How many properties per recipient

		// Set the SPropValue members == the desired values.
		lpAdrList->aEntries[0].rgPropVals[NAME].ulPropTag = PR_DISPLAY_NAME;
		lpAdrList->aEntries[0].rgPropVals[NAME].Value.LPSZ = (LPTSTR) pstrName->c_str();

		CORg(lpAdrBook->ResolveName(
			0L,
			fMapiUnicode,
			NULL,
			lpAdrList));

		CORg(MAPIAllocateBuffer(sizeof(SBinary), (LPVOID*) &lpEID));
		ZeroMemory(lpEID, sizeof(SBinary));
		bool foundEID = false;
		for (UINT x = 0; x < lpAdrList->aEntries[0].cValues; x++)
		{
			if (lpAdrList->aEntries[0].rgPropVals[x].ulPropTag == PR_ENTRYID)
			{
				CORg(this->CopySBinary(lpEID, &lpAdrList->aEntries[0].rgPropVals[x].Value.bin, NULL));
				foundEID = true;
				break;
			}
		}

		if (!foundEID)
		{
			hr = MAPI_E_NOT_FOUND;
			goto Error;
		}
	}

Cleanup:
	if (lpAdrList)
		FreePadrlist(lpAdrList);
	if (lpAdrBook)
		lpAdrBook->Release();
	return lpEID;
Error:
	goto Cleanup;
}

// Borrowed from MfcMapi
// Build a server DN. Allocates memory. Free with MAPIFreeBuffer.
HRESULT OperationBase::BuildServerDN(
					  _In_z_ LPCTSTR szServerName,
					  _In_z_ LPCTSTR szPost,
					  _Deref_out_z_ LPTSTR* lpszServerDN)
{
	HRESULT hr = S_OK;
	if (!lpszServerDN) return MAPI_E_INVALID_PARAMETER;

	static LPCTSTR szPre = _T("/cn=Configuration/cn=Servers/cn="); // STRING_OK
	size_t cbPreLen = 0;
	size_t cbServerLen = 0;
	size_t cbPostLen = 0;
	size_t cbServerDN = 0;

	CORg(StringCbLength(szPre,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbPreLen));
	CORg(StringCbLength(szServerName,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbServerLen));
	CORg(StringCbLength(szPost,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbPostLen));

	cbServerDN = cbPreLen + cbServerLen + cbPostLen + sizeof(TCHAR);

	CORg(MAPIAllocateBuffer(
		(ULONG) cbServerDN,
		(LPVOID*)lpszServerDN));

	CORg(StringCbPrintf(
		*lpszServerDN,
		cbServerDN,
		_T("%s%s%s"), // STRING_OK
		szPre,
		szServerName,
		szPost));
Error:

	return hr;
} // BuildServerDN

std::vector<tstring> OperationBase::Split(const tstring &s, TCHAR delim)
{
    std::vector<tstring> elems;
    Split(s, delim, elems);
    return elems;
}

std::vector<tstring> &OperationBase::Split(const tstring &s, TCHAR delim, std::vector<tstring> &elems)
{
    std::wstringstream ss(s);
    tstring item;
    while (std::getline(ss, item, delim)) {
        elems.push_back(item);
    }
    return elems;
}

// From MFCMapi MAPIFunctions.cpp
///////////////////////////////////////////////////////////////////////////////
//	CopySBinary()
//
//	Parameters
//		psbDest - Address of the destination binary
//		psbSrc  - Address of the source binary
//		lpParent - Pointer to parent object (not, however, pointer to pointer!)
//
//	Purpose
//		Allocates a new SBinary and copies psbSrc into it
//
HRESULT OperationBase::CopySBinary(_Out_ LPSBinary psbDest, _In_ const LPSBinary psbSrc, _In_ LPVOID lpParent)
{
	HRESULT hr = S_OK;

	if (!psbDest || !psbSrc) return MAPI_E_INVALID_PARAMETER;

	psbDest->cb = psbSrc->cb;

	if (psbSrc->cb)
	{
		if (lpParent)
		{
			CORg(MAPIAllocateMore(
				psbSrc->cb,
				lpParent,
				(LPVOID *) &psbDest->lpb));
		}
		else
		{
			CORg(MAPIAllocateBuffer(
				psbSrc->cb,
				(LPVOID *) &psbDest->lpb));
		}
		if (S_OK == hr)
			CopyMemory(psbDest->lpb,psbSrc->lpb,psbSrc->cb);
	}
Error:

	return hr;
} // CopySBinary

// From MAPIABFunctions.cpp in MFCMapi
HRESULT OperationBase::HrAllocAdrList(ULONG ulNumProps, _Deref_out_opt_ LPADRLIST* lpAdrList)
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

void OperationBase::OutputSBinary(SBinary lpsbin)
{
	for (UINT x = 0; x < lpsbin.cb; x++)
	{
		*pLog << std::hex << static_cast<int>(lpsbin.lpb[x]);
	}
}

// From MapiStoreFunctions.ccp in MFCMapi
HRESULT OperationBase::OpenDefaultMessageStore(
								_In_ LPMAPISESSION lpMAPISession,
								_Deref_out_ LPMDB* lppDefaultMDB)
{
	HRESULT				hr = S_OK;
	LPMAPITABLE			pStoresTbl = NULL;
	LPSRowSet			pRow = NULL;
	static SRestriction	sres;
	SPropValue			spv;

	enum
	{
		EID,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS,sptEIDCol) =
	{
		NUM_COLS,
		PR_ENTRYID,
	};
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	CORg(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

	// set up restriction for the default store
	sres.rt = RES_PROPERTY; // gonna compare a property
	sres.res.resProperty.relop = RELOP_EQ; // gonna test equality
	sres.res.resProperty.ulPropTag = PR_DEFAULT_STORE; // tag to compare
	sres.res.resProperty.lpProp = &spv; // prop tag to compare against

	spv.ulPropTag = PR_DEFAULT_STORE; // tag type
	spv.Value.b	= true; // tag value

	CORg(HrQueryAllRows(
		pStoresTbl,						// table to query
		(LPSPropTagArray) &sptEIDCol,	// columns to get
		&sres,							// restriction to use
		NULL,							// sort order
		0,								// max number of rows - 0 means ALL
		&pRow));

	if (pRow && pRow->cRows)
	{
		CORg(CallOpenMsgStore(
			lpMAPISession,
			NULL,
			&pRow->aRow[0].lpProps[EID].Value.bin,
			MDB_WRITE,
			lppDefaultMDB));
	}
	else hr = MAPI_E_NOT_FOUND;

Cleanup:

	if (pRow) FreeProws(pRow);
	if (pStoresTbl) pStoresTbl->Release();
	return hr;
Error:
	goto Cleanup;
} // OpenDefaultMessageStore

HRESULT OperationBase::CallOpenMsgStore(
						 _In_ LPMAPISESSION	lpSession,
						 _In_ ULONG_PTR		ulUIParam,
						 _In_ LPSBinary		lpEID,
						 ULONG			ulFlags,
						 _Deref_out_ LPMDB*			lpMDB)
{
	if (!lpSession || !lpMDB || !lpEID) return MAPI_E_INVALID_PARAMETER;

	HRESULT hr = S_OK;

	// MAPIFolders always uses MDB_ONLINE
	ulFlags |= MDB_ONLINE;

	CORg(lpSession->OpenMsgStore(
		ulUIParam,
		lpEID->cb,
		(LPENTRYID) lpEID->lpb,
		NULL,
		ulFlags,
		(LPMDB*) lpMDB));

	// MAPIFolders will never hit this statement due to the CORg macro
	// This is intentional
	if (MAPI_E_UNKNOWN_FLAGS == hr && (ulFlags & MDB_ONLINE))
	{
		hr = S_OK;
		// perhaps this store doesn't know the MDB_ONLINE flag - remove and retry
		ulFlags = ulFlags & ~MDB_ONLINE;
		CORg(lpSession->OpenMsgStore(
			ulUIParam,
			lpEID->cb,
			(LPENTRYID) lpEID->lpb,
			NULL,
			ulFlags,
			(LPMDB*) lpMDB));
	}

Cleanup:
	return hr;
Error:
	goto Cleanup;
} // CallOpenMsgStore

// From MFCMapi MapiFunctions.cpp
// if cchszW == -1, WideCharToMultiByte will compute the length
// Delete with delete[]
HRESULT OperationBase::UnicodeToAnsi(_In_z_ LPCWSTR pszW, _Out_z_cap_(cchszW) LPSTR* ppszA, size_t cchszW)
{
	HRESULT hr = S_OK;
	if (!ppszA) return MAPI_E_INVALID_PARAMETER;
	*ppszA = NULL;
	if (NULL == pszW) return S_OK;

	// Get our buffer size
	int iRet = 0;
	iRet = WideCharToMultiByte(
		CP_ACP,
		0,
		pszW,
		(int) cchszW,
		NULL,
		NULL,
		NULL,
		NULL);

	if (SUCCEEDED(hr) && 0 != iRet)
	{
		// WideCharToMultiByte returns num of bytes
		LPSTR pszA = (LPSTR) new BYTE[iRet];

		iRet = WideCharToMultiByte(
			CP_ACP,
			0,
			pszW,
			(int) cchszW,
			pszA,
			iRet,
			NULL,
			NULL);
		if (SUCCEEDED(hr))
		{
			*ppszA = pszA;
		}
		else
		{
			delete[] pszA;
		}
	}
	return hr;
} // UnicodeToAnsi

// if cchszA == -1, MultiByteToWideChar will compute the length
// Delete with delete[]
HRESULT OperationBase::AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, size_t cchszA)
{
	HRESULT hr = S_OK;
	if (!ppszW) return MAPI_E_INVALID_PARAMETER;
	*ppszW = NULL;
	if (NULL == pszA) return S_OK;
	if (!cchszA) return S_OK;

	// Get our buffer size
	int iRet = 0;
	iRet = MultiByteToWideChar(
		CP_ACP,
		0,
		pszA,
		(int) cchszA,
		NULL,
		NULL);

	if (SUCCEEDED(hr) && 0 != iRet)
	{
		// MultiByteToWideChar returns num of chars
		LPWSTR pszW = new WCHAR[iRet];

		iRet = MultiByteToWideChar(
			CP_ACP,
			0,
			pszA,
			(int) cchszA,
			pszW,
			iRet);
		if (SUCCEEDED(hr))
		{
			*ppszW = pszW;
		}
		else
		{
			delete[] pszW;
		}
	}
	return hr;
} // AnsiToUnicode

HRESULT OperationBase::GetServerName(_In_ LPMAPISESSION lpSession, _Deref_out_opt_z_ LPTSTR* szServerName)
{
	HRESULT			hr = S_OK;
	LPSERVICEADMIN	pSvcAdmin = NULL;
	LPPROFSECT		pGlobalProfSect = NULL;
	LPSPropValue	lpServerName	= NULL;

	if (!lpSession) return MAPI_E_INVALID_PARAMETER;

	*szServerName = NULL;

	CORg(lpSession->AdminServices(
		0,
		&pSvcAdmin));

	CORg(pSvcAdmin->OpenProfileSection(
		(LPMAPIUID)pbGlobalProfileSectionGuid,
		NULL,
		0,
		&pGlobalProfSect));

	CORg(HrGetOneProp(pGlobalProfSect,
		PR_PROFILE_HOME_SERVER,
		&lpServerName));

	if (CheckStringProp(lpServerName,PT_STRING8)) // profiles are ASCII only
	{
#ifdef UNICODE
		LPWSTR	szWideServer = NULL;
		CORg(AnsiToUnicode(
			lpServerName->Value.lpszA,
			&szWideServer,-1));
		CORg(CopyStringW(szServerName,szWideServer,NULL));
		delete[] szWideServer;
#else
		CORg(CopyStringA(szServerName,lpServerName->Value.lpszA,NULL));
#endif
	}

	MAPIFreeBuffer(lpServerName);
	if (pGlobalProfSect) pGlobalProfSect->Release();
	if (pSvcAdmin) pSvcAdmin->Release();
Error:

	return hr;
} // GetServerName

bool OperationBase::CheckStringProp(_In_opt_ LPSPropValue lpProp, ULONG ulPropType)
{
	if (PT_STRING8 != ulPropType && PT_UNICODE != ulPropType)
	{
		return false;
	}
	if (!lpProp)
	{
		return false;
	}

	if (PT_ERROR == PROP_TYPE(lpProp->ulPropTag))
	{
		return false;
	}

	if (ulPropType != PROP_TYPE(lpProp->ulPropTag))
	{
		return false;
	}

	if (NULL == lpProp->Value.LPSZ)
	{
		return false;
	}

	if (PT_STRING8 == ulPropType && NULL == lpProp->Value.lpszA[0])
	{
		return false;
	}

	if (PT_UNICODE == ulPropType && NULL == lpProp->Value.lpszW[0])
	{
		return false;
	}

	return true;
} // CheckStringProp

HRESULT OperationBase::CopyStringW(_Deref_out_z_ LPWSTR* lpszDestination, _In_z_ LPCWSTR szSource, _In_opt_ LPVOID pParent)
{
	HRESULT	hr = S_OK;
	size_t	cbSource = 0;

	if (!szSource)
	{
		*lpszDestination = NULL;
		return hr;
	}

	CORg(StringCbLengthW(szSource,STRSAFE_MAX_CCH * sizeof(WCHAR),&cbSource));
	cbSource += sizeof(WCHAR);

	if (pParent)
	{
		CORg(MAPIAllocateMore(
			(ULONG) cbSource,
			pParent,
			(LPVOID*) lpszDestination));
	}
	else
	{
		CORg(MAPIAllocateBuffer(
			(ULONG) cbSource,
			(LPVOID*) lpszDestination));
	}
	CORg(StringCbCopyW(*lpszDestination, cbSource, szSource));
Error:

	return hr;
} // CopyStringW
