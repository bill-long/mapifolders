#pragma once
#include "stdafx.h"

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
#include <vector>
#include <sstream>
#include <string>

#include "UserArgs.h"

#define MDB_ONLINE 0x100

class OperationBase
{
public:
	OperationBase(tstring *basePath, tstring *mailbox, UserArgs::ActionScope scope);
	~OperationBase(void);
	HRESULT OperationBase::Initialize(void);
	void DoOperation();
	virtual void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);
	std::string GetStringFromFolderPath(LPMAPIFOLDER folder);
	HRESULT OperationBase::CopySBinary(_Out_ LPSBinary psbDest, _In_ const LPSBinary psbSrc, _In_ LPVOID lpParent);
	bool OperationBase::IsEntryIdEqual(SBinary a, SBinary b);
	LPSBinary OperationBase::ResolveNameToEID(tstring *pstrName);
	HRESULT OperationBase::HrAllocAdrList(ULONG ulNumProps, _Deref_out_opt_ LPADRLIST* lpAdrList);
	LPMAPISESSION lpSession;

private:
	LPMAPIFOLDER GetPFRoot(IMAPISession *pSession);
	LPMAPIFOLDER GetMailboxRoot(IMAPISession *pSession);
	LPMAPIFOLDER OperationBase::GetStartingFolder(IMAPISession *pSession, tstring *calculatedFolderPath);
	HRESULT OperationBase::GetSubfolderByName(LPMAPIFOLDER parentFolder, tstring folderNameToFind, LPMAPIFOLDER *returnedFolder, tstring *returnedFolderName);
	void TraverseFolders(CComPtr<IMAPISession> session, LPMAPIFOLDER baseFolder, tstring parentPath);
	HRESULT BuildServerDN(LPCSTR szServerName, LPCSTR szPost, LPSTR* lpszServerDN);
	std::vector<tstring> OperationBase::Split(const tstring &s, TCHAR delim);
	std::vector<tstring> &Split(const tstring &s, TCHAR delim, std::vector<tstring> &elems);
	LPMAPIFOLDER lpRootFolder;
	LPMAPIFOLDER lpStartingFolder;
	tstring startingPath;
	LPMDB lpAdminMDB;
	tstring *strBasePath;
	tstring *strMailbox;
	UserArgs::ActionScope nScope;
};

