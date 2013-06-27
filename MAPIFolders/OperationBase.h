#pragma once
#include "stdafx.h"
#include "MAPIFolders.h"

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
	OperationBase(CComPtr<IMAPISession> session);
	~OperationBase(void);
	virtual void ProcessFolder(LPMAPIFOLDER folder);
	std::string GetStringFromFolderPath(LPMAPIFOLDER folder);

private:
	CComPtr<IMAPISession> spSession;
};

