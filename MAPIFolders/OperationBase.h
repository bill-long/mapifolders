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

class OperationBase
{
public:
	OperationBase();
	~OperationBase(void);
	void DoOperation();
	virtual void ProcessFolder(LPMAPIFOLDER folder, std::wstring folderPath);
	std::string GetStringFromFolderPath(LPMAPIFOLDER folder);

private:
	LPMAPIFOLDER GetPFRoot(IMAPISession *pSession);
	void TraverseFolders(CComPtr<IMAPISession> session, LPMAPIFOLDER baseFolder, std::wstring parentPath);
	HRESULT BuildServerDN(LPCTSTR szServerName, LPCTSTR szPost, LPTSTR* lpszServerDN);
	LPMDB lpAdminMDB;
};

