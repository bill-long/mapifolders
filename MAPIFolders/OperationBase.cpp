#include "stdafx.h"
#include "OperationBase.h"
#include "MAPIFolders.h"
#include <algorithm>


OperationBase::OperationBase()
{
	this->lpAdminMDB = NULL;
}


OperationBase::~OperationBase(void)
{
}


void OperationBase::ProcessFolder(LPMAPIFOLDER folder, std::wstring folderPath)
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
	lpPFRoot = NULL;
	lpSession = NULL;
	std::wstring rootPath(L"");

	MAPIINIT_0  MAPIINIT= { 0, MAPI_MULTITHREAD_NOTIFICATIONS};

	CORg(MAPIInitialize (&MAPIINIT));

	CORg(MAPILogonEx(0, NULL, NULL, MAPI_UNICODE | MAPI_LOGON_UI | MAPI_EXTENDED | MAPI_ALLOW_OTHERS | MAPI_LOGON_UI | MAPI_EXPLICIT_PROFILE, &lpSession));
		
	lpPFRoot = GetPFRoot(lpSession);
	if (lpPFRoot == NULL)
		goto Error;

	TraverseFolders(lpSession, lpPFRoot, rootPath);

Cleanup:
	if (lpPFRoot)
		lpPFRoot->Release();
	if (lpAdminMDB)
		lpAdminMDB->Release();
	if (lpSession)
		lpSession->Release();
	return;
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

	std::wcout << L"Found " << pmrows->cRows << L" stores in MAPI profile:" << std::endl;
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
			std::wcout << L" " << std::setw(4) << i << ": " << (pwz ? pwz : L"<NULL>") << std::endl;
		else
			std::wcout << L" " << std::setw(4) << i << ": " << (pwzA ? pwzA : "<NULL>") << std::endl;

		if (IsEqualMAPIUID(prow->lpProps[COL_MDB_PROVIDER].Value.bin.lpb,pbExchangeProviderPublicGuid))
		{
			publicEntryID = prow->lpProps[COL_ENTRYID].Value.bin;
			foundPFStore = TRUE;
		}
	}

	if (!foundPFStore)
	{
		std::wcout << "No public folder database in MAPI profile." << std::endl;
		goto Error;
	}

	if (publicEntryID.cb < 1)
	{
		std::wcout << "Could not get public folder store entry ID." << std::endl;
		goto Error;
	}

	std::wcout << "Opening public folders..." << std::endl;
	CORg(pSession->OpenMsgStore(NULL, publicEntryID.cb, (LPENTRYID)publicEntryID.lpb, NULL,
		MAPI_BEST_ACCESS | 0x00000100, &lpMDB));

	CORg(HrGetOneProp(
			lpMDB,
			PR_HIERARCHY_SERVER,
			&lpServerName));

	std::cout << "Using public folder server/mailbox: " << lpServerName->Value.lpszA << std::endl;

	CORg(BuildServerDN(
				(LPCTSTR)lpServerName->Value.lpszA,
				_T("/cn=Microsoft Public MDB"), // STRING_OK
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
		MAPI_BEST_ACCESS | 0x00000100, &lpAdminMDB)); // MDB_ONLINE is 0x100

	ULONG ulObjType = NULL;
	CORg(lpAdminMDB->OpenEntry(NULL, NULL, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *) &lpRoot));

Error:
	if (pmrows)
		FreeProws(pmrows);
	if (szServerDN)
		MAPIFreeBuffer(szServerDN);

	return lpRoot;
}

void OperationBase::TraverseFolders(CComPtr<IMAPISession> session, LPMAPIFOLDER baseFolder, std::wstring parentPath)
{
	SRowSet *pmrows = NULL;
	LPMAPITABLE hierarchyTable = NULL;
	HRESULT hr = NULL;
	ULONG objType = NULL;

	enum MAPIColumns
	{
		COL_ENTRYID = 0,
		COL_DISPLAYNAME_W,
		COL_DISPLAYNAME_A,
		cCols				// End marker
	};

	static const SizedSPropTagArray(cCols, mcols) = 
	{	
		cCols,
		{
			PR_ENTRYID,
			PR_DISPLAY_NAME_W,
			PR_DISPLAY_NAME_A,
		}
	};

RetryGetHierarchyTable:
	hr = baseFolder->GetHierarchyTable(NULL, &hierarchyTable);
	if (FAILED(hr))
	{
		if (hr == MAPI_E_TIMEOUT)
		{
			std::wcout << L"     Encountered a timeout in GetHierarchyTable. Retrying in 5 seconds..." << std::endl;
			Sleep(5000);
			goto RetryGetHierarchyTable;
		}
		else if (hr == MAPI_E_NETWORK_ERROR)
		{
			std::wcout << L"     Encountered a network error in GetHierarchyTable. Retrying in 5 seconds..." << std::endl;
			Sleep(5000);
			goto RetryGetHierarchyTable;
		}
		else
		{
			std::wcout << L"     GetHierarchyTable returned an error. hr = " << std::hex << hr << std::endl;
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
			std::wcout << L"     Encountered a timeout in QueryRows. Retrying in 5 seconds..." << std::endl;
			pmrows = NULL;
			Sleep(5000);
			goto RetryQueryRows;
		}
		else if (hr == MAPI_E_NETWORK_ERROR)
		{
			std::wcout << L"     Encountered a network error in QueryRows. Retrying in 5 seconds..." << std::endl;
			pmrows = NULL;
			Sleep(5000);
			goto RetryQueryRows;
		}
		else
		{
			std::wcout << L"     QueryRows returned an error. hr = " << std::hex << hr << std::endl;
			goto Error;
		}
	}

	for (UINT i=0; i != pmrows->cRows; i++)
	{
		SRow *prow = pmrows->aRow + i;
		LPCWSTR pwz = NULL;
		if (PR_DISPLAY_NAME_W == prow->lpProps[COL_DISPLAYNAME_W].ulPropTag)
			pwz = prow->lpProps[COL_DISPLAYNAME_W].Value.lpszW;

		std::wstring thisPath(parentPath.c_str());
		thisPath.append(L"\\");
		thisPath.append(pwz);

		ULONG ulObjType = NULL;
		LPMAPIFOLDER lpSubfolder = NULL;

RetryOpenEntry:
		hr = lpAdminMDB->OpenEntry(prow->lpProps[COL_ENTRYID].Value.bin.cb, (LPENTRYID)prow->lpProps[COL_ENTRYID].Value.bin.lpb, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *) &lpSubfolder);
		if (FAILED(hr))
		{
			if (hr == MAPI_E_TIMEOUT)
			{
				std::wcout << L"     Encountered a timeout in OpenEntry. Retrying in 5 seconds..." << std::endl;
				Sleep(5000);
				goto RetryOpenEntry;
			}
			else if (hr == MAPI_E_NETWORK_ERROR)
			{
				std::wcout << L"     Encountered a network error in OpenEntry. Retrying in 5 seconds..." << std::endl;
				Sleep(5000);
				goto RetryOpenEntry;
			}
			else
			{
				std::wcout << L"     OpenEntry returned an error. hr = " << std::hex << hr << std::endl;
				goto Error;
			}
		}

		if (lpSubfolder)
		{
			this->ProcessFolder(lpSubfolder, thisPath);
		}

		TraverseFolders(session, lpSubfolder, thisPath);

		lpSubfolder->Release();
	}

Error:
	if (pmrows)
		FreeProws(pmrows);
	if (hierarchyTable)
		hierarchyTable->Release();

	return;
}

// Borrowed from MfcMapi
// Build a server DN. Allocates memory. Free with MAPIFreeBuffer.
_Check_return_ HRESULT OperationBase::BuildServerDN(
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
