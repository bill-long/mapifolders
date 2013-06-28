// FixPFItems.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "MAPIFolders.h"
#include "ValidateFolderACL.h"

#include <windows.h>

#include <mapix.h>
#include <mapiutil.h>
#include <mapi.h>
#include <atlbase.h>

#include <iostream>
#include <iomanip>
#include <initguid.h>
#include <edkguid.h>
#include <edkmdb.h>
#include <strsafe.h>
#include <algorithm>

#pragma warning( disable : 4127)

// Globals
	LPMDB lpAdminMDB = NULL;

// Borrowed from MfcMapi
// Build a server DN. Allocates memory. Free with MAPIFreeBuffer.
_Check_return_ HRESULT BuildServerDN(
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

LPMAPIFOLDER GetPFRoot(IMAPISession *pSession)
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

void TraverseFolders(CComPtr<IMAPISession> session, LPMAPIFOLDER baseFolder, std::wstring parentPath)
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

	CORg(baseFolder->GetHierarchyTable(NULL, &hierarchyTable));
	CORg(hierarchyTable->SetColumns((LPSPropTagArray)&mcols, TBL_BATCH));
	CORg(hierarchyTable->SeekRow(BOOKMARK_BEGINNING, 0, 0));
	CORg(hierarchyTable->QueryRows(1000000, 0, &pmrows));

	ValidateFolderACL* opValidateFolderACL = new ValidateFolderACL(session);

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
		CORg(lpAdminMDB->OpenEntry(prow->lpProps[COL_ENTRYID].Value.bin.cb, (LPENTRYID)prow->lpProps[COL_ENTRYID].Value.bin.lpb, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *) &lpSubfolder));
		if (lpSubfolder)
		{
			opValidateFolderACL->ProcessFolder(lpSubfolder, thisPath);
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

int _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hr = S_OK;
	LPMDB lpMDB = NULL;
	LPMAPIFOLDER lpPFRoot = NULL;

	MAPIINIT_0  MAPIINIT= { 0, MAPI_MULTITHREAD_NOTIFICATIONS};

	CORg(MAPIInitialize (&MAPIINIT));

	{	// IMAPISession Smart Pointer Context
		CComPtr<IMAPISession> spSession;
		CORg(MAPILogonEx(0, NULL, NULL, MAPI_UNICODE | MAPI_LOGON_UI | MAPI_EXTENDED | MAPI_ALLOW_OTHERS | MAPI_LOGON_UI | MAPI_EXPLICIT_PROFILE, &spSession));
		
		lpPFRoot = GetPFRoot(spSession);
		if (lpPFRoot == NULL)
			goto Error;

		std::wstring rootPath(L"");
		TraverseFolders(spSession, lpPFRoot, rootPath);

	}

Cleanup:
	if (lpPFRoot)
		lpPFRoot->Release();
	if (lpMDB)
		lpMDB->Release();

	return 0;

Error:
	goto Cleanup;
}

