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
	LPMAPIFOLDER GetStartingFolder(IMAPISession *pSession);
	void TraverseFolders(CComPtr<IMAPISession> session, LPMAPIFOLDER baseFolder, tstring parentPath);
	HRESULT BuildServerDN(LPCTSTR szServerName, LPCTSTR szPost, LPTSTR* lpszServerDN);
	LPMAPIFOLDER lpPFRoot;
	LPMAPIFOLDER lpStartingFolder;
	LPMDB lpAdminMDB;
	LPMAPISESSION lpSession;
	tstring *strBasePath;
	UserArgs::ActionScope nScope;
};

