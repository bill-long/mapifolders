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
	OperationBase(tstring *basePath, UserArgs::ActionScope scope);
	~OperationBase(void);
	void DoOperation();
	virtual void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);
	std::string GetStringFromFolderPath(LPMAPIFOLDER folder);

private:
	LPMAPIFOLDER GetPFRoot(IMAPISession *pSession);
	LPMAPIFOLDER OperationBase::GetStartingFolder(IMAPISession *pSession, tstring *calculatedFolderPath);
	HRESULT OperationBase::GetSubfolderByName(LPMAPIFOLDER parentFolder, tstring folderNameToFind, LPMAPIFOLDER *returnedFolder, tstring *returnedFolderName);
	void TraverseFolders(CComPtr<IMAPISession> session, LPMAPIFOLDER baseFolder, tstring parentPath);
	HRESULT BuildServerDN(LPCSTR szServerName, LPCSTR szPost, LPSTR* lpszServerDN);
	std::vector<tstring> OperationBase::Split(const tstring &s, TCHAR delim);
	std::vector<tstring> &Split(const tstring &s, TCHAR delim, std::vector<tstring> &elems);
	LPMAPIFOLDER lpPFRoot;
	LPMAPIFOLDER lpStartingFolder;
	LPMDB lpAdminMDB;
	LPMAPISESSION lpSession;
	tstring *strBasePath;
	UserArgs::ActionScope nScope;
};

