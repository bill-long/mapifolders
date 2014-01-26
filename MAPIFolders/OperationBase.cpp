#include "stdafx.h"
#include "OperationBase.h"
#include "MAPIFolders.h"
#include <algorithm>

OperationBase::OperationBase(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope)
{
	this->lpAdminMDB = NULL;
	this->strBasePath = pstrBasePath;
	this->strMailbox = pstrMailbox;
	this->nScope = nScope;
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


void OperationBase::DoOperation()
{
	HRESULT hr = S_OK;
	lpAdminMDB = NULL;
	lpRootFolder = NULL;
	lpStartingFolder = NULL;
	lpSession = NULL;
	tstring startingPath(_T(""));

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
		goto Error;

	lpStartingFolder = GetStartingFolder(lpSession, &startingPath);
	if (lpStartingFolder == NULL)
	{
		tcout << "Failed to find the specified folder" << std::endl;
		goto Error;
	}

	TraverseFolders(lpSession, lpStartingFolder, startingPath);

Cleanup:
	if (lpRootFolder)
		lpRootFolder->Release();
	if (lpAdminMDB)
		lpAdminMDB->Release();
	if (lpSession)
		lpSession->Release();
	return;
Error:
	goto Cleanup;
}

LPMAPIFOLDER OperationBase::GetMailboxRoot(IMAPISession *pSession)
{
	HRESULT hr = S_OK;

	return NULL;
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
	LPSTR	szServerDN = NULL;

	enum MAPIColumns
	{
		COL_ENTRYID = 0,
		COL_DISPLAYNAME_W,
		COL_DISPLAYNAME_A,
		COL_MDB_PROVIDER,
		cCols				// End marker
	};

	static const SizedSPropTagArray(cCols, mcols) = 
	{	
		cCols,
		{
			PR_ENTRYID,
			PR_DISPLAY_NAME_W,
			PR_DISPLAY_NAME_A,
			PR_MDB_PROVIDER
		}
	};

	CORg(pSession->GetMsgStoresTable(0, &spTable));

	CORg(spTable->SetColumns((LPSPropTagArray)&mcols, TBL_BATCH));
	CORg(spTable->SeekRow(BOOKMARK_BEGINNING, 0, 0));

	CORg(spTable->QueryRows(50, 0, &pmrows));

	tcout << "Found " << pmrows->cRows << " stores in MAPI profile:" << std::endl;
	for (UINT i=0; i != pmrows->cRows; i++)
	{
		SRow *prow = pmrows->aRow + i;
		LPCWSTR pwz = NULL;
		LPCSTR pwzA = NULL;
		if (PR_DISPLAY_NAME_W == prow->lpProps[COL_DISPLAYNAME_W].ulPropTag)
			pwz = prow->lpProps[COL_DISPLAYNAME_W].Value.lpszW;
		if (PR_DISPLAY_NAME_A == prow->lpProps[COL_DISPLAYNAME_A].ulPropTag)
			pwzA = prow->lpProps[COL_DISPLAYNAME_A].Value.lpszA;
		if (pwz)
			std::wcout << L" " << std::setw(4) << i << L": " << (pwz ? pwz : L"<NULL>") << std::endl;
		else
			tcout << " " << std::setw(4) << i << ": " << (pwzA ? pwzA : "<NULL>") << std::endl;

		if (IsEqualMAPIUID(prow->lpProps[COL_MDB_PROVIDER].Value.bin.lpb,pbExchangeProviderPublicGuid))
		{
			publicEntryID = prow->lpProps[COL_ENTRYID].Value.bin;
			foundPFStore = TRUE;
		}
	}

	if (!foundPFStore)
	{
		tcout << "No public folder database in MAPI profile." << std::endl;
		goto Error;
	}

	if (publicEntryID.cb < 1)
	{
		tcout << "Could not get public folder store entry ID." << std::endl;
		goto Error;
	}

	tcout << "Opening public folders..." << std::endl;
	CORg(pSession->OpenMsgStore(NULL, publicEntryID.cb, (LPENTRYID)publicEntryID.lpb, NULL,
		MAPI_BEST_ACCESS | MDB_ONLINE, &lpMDB));

	CORg(HrGetOneProp(
			lpMDB,
			PR_HIERARCHY_SERVER,
			&lpServerName));

	tcout << "Using public folder server/mailbox: " << lpServerName->Value.lpszW << std::endl;

	CORg(BuildServerDN(
				(LPCSTR)lpServerName->Value.lpszA,
				("/cn=Microsoft Public MDB"),
				&szServerDN));

	CORg(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(LPVOID*) &lpXManageStore));

	LPSTR lpszMailboxDN = NULL;
	ULONG flags = OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_PUBLIC;
	CORg(lpXManageStore->CreateStoreEntryID((LPSTR)szServerDN, lpszMailboxDN,
		flags,
		&adminEntryID.cb, (LPENTRYID *)&adminEntryID.lpb));

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

	if (this->strBasePath == NULL)
	{
		// If user didn't specify a path, assume IPM_SUBTREE
		userSpecifiedPath = (_T("\\IPM_SUBTREE"));
	}
	else
	{
		userSpecifiedPath = *this->strBasePath;
	}

	// Split the specified pathing using backslash separator
	std::vector<tstring> splitPath = Split(userSpecifiedPath, (_T('\\')));

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
	UINT startIndex = 1;
	if ((_tstricmp(splitPath.at(1).c_str(), _T("non_ipm_subtree")) == 0) ||
		(_tstricmp(splitPath.at(1).c_str(), _T("ipm_subtree")) == 0))
	{
		startIndex = 2;
	}

	// If startIndex points to the 3rd item, but we only have 2 because
	// a root folder was specified as the starting folder, then we're
	// already done
	if (startIndex == 2 && splitPath.size() == 2)
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
			tcout << "Could not find folder: " << splitPath.at(startIndex) << std::endl;
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
				tcout << "Could not find folder: " << splitPath.at(x) << std::endl;
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

RetryGetHierarchyTable:
		hr = baseFolder->GetHierarchyTable(NULL, &hierarchyTable);
		if (FAILED(hr))
		{
			if (hr == MAPI_E_TIMEOUT)
			{
				tcout << "     Encountered a timeout in GetHierarchyTable. Retrying in 5 seconds..." << std::endl;
				Sleep(5000);
				goto RetryGetHierarchyTable;
			}
			else if (hr == MAPI_E_NETWORK_ERROR)
			{
				tcout << "     Encountered a network error in GetHierarchyTable. Retrying in 5 seconds..." << std::endl;
				Sleep(5000);
				goto RetryGetHierarchyTable;
			}
			else
			{
				tcout << "     GetHierarchyTable returned an error. hr = " << std::hex << hr << std::endl;
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
				tcout << "     Encountered a timeout in QueryRows. Retrying in 5 seconds..." << std::endl;
				pmrows = NULL;
				Sleep(5000);
				goto RetryQueryRows;
			}
			else if (hr == MAPI_E_NETWORK_ERROR)
			{
				tcout << "     Encountered a network error in QueryRows. Retrying in 5 seconds..." << std::endl;
				pmrows = NULL;
				Sleep(5000);
				goto RetryQueryRows;
			}
			else
			{
				tcout << "     QueryRows returned an error. hr = " << std::hex << hr << std::endl;
				goto Error;
			}
		}

		for (UINT i=0; i != pmrows->cRows; i++)
		{
			SRow *prow = pmrows->aRow + i;
			LPCWSTR pwz = NULL;
			if (PR_DISPLAY_NAME_W == prow->lpProps[COL_DISPLAYNAME_W].ulPropTag)
				pwz = prow->lpProps[COL_DISPLAYNAME_W].Value.lpszW;

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
					tcout << "     Encountered a timeout in OpenEntry. Retrying in 5 seconds..." << std::endl;
					Sleep(5000);
					goto RetryOpenEntry;
				}
				else if (hr == MAPI_E_NETWORK_ERROR)
				{
					tcout << "     Encountered a network error in OpenEntry. Retrying in 5 seconds..." << std::endl;
					Sleep(5000);
					goto RetryOpenEntry;
				}
				else
				{
					tcout << "     OpenEntry returned an error. hr = " << std::hex << hr << std::endl;
					goto Error;
				}
			}

			if (lpSubfolder)
			{
				this->ProcessFolder(lpSubfolder, thisPath);
			}

			if (this->nScope == UserArgs::ActionScope::SUBTREE)
			{
				TraverseFolders(session, lpSubfolder, thisPath);
			}

			lpSubfolder->Release();
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

// Borrowed from MfcMapi mostly
// CreateStoreEntryID doesn't seem to support Unicode, so changed
// this to be hard-coded to ASCII.
// Build a server DN. Allocates memory. Free with MAPIFreeBuffer.
_Check_return_ HRESULT OperationBase::BuildServerDN(
					  _In_z_ LPCSTR szServerName,
					  _In_z_ LPCSTR szPost,
					  _Deref_out_z_ LPSTR* lpszServerDN)
{
	HRESULT hr = S_OK;
	if (!lpszServerDN) return MAPI_E_INVALID_PARAMETER;

	static LPCSTR szPre = ("/cn=Configuration/cn=Servers/cn=");
	size_t cbPreLen = 0;
	size_t cbServerLen = 0;
	size_t cbPostLen = 0;
	size_t cbServerDN = 0;

	CORg(StringCbLengthA(szPre,STRSAFE_MAX_CCH * sizeof(CHAR),&cbPreLen));
	CORg(StringCbLengthA(szServerName,STRSAFE_MAX_CCH * sizeof(CHAR),&cbServerLen));
	CORg(StringCbLengthA(szPost,STRSAFE_MAX_CCH * sizeof(CHAR),&cbPostLen));

	cbServerDN = cbPreLen + cbServerLen + cbPostLen + sizeof(CHAR);

	CORg(MAPIAllocateBuffer(
		(ULONG) cbServerDN,
		(LPVOID*)lpszServerDN));

	CORg(StringCbPrintfA(
		*lpszServerDN,
		cbServerDN,
		("%s%s%s"),
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

bool OperationBase::IsEntryIdEqual(SBinary a, SBinary b)
{
	if (a.cb != b.cb)
		return false;

	return (!memcmp(a.lpb, b.lpb, a.cb));
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