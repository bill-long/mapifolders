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
#include "Log.h"

#define MDB_ONLINE 0x100

class OperationBase
{
public:
	OperationBase(tstring *basePath, tstring *mailbox, UserArgs::ActionScope scope, Log *log);
	~OperationBase(void);
	HRESULT OperationBase::Initialize(void);
	void DoOperation();
	virtual void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);
	std::string GetStringFromFolderPath(LPMAPIFOLDER folder);
	HRESULT OperationBase::CopySBinary(_Out_ LPSBinary psbDest, _In_ const LPSBinary psbSrc, _In_ LPVOID lpParent);
	bool OperationBase::IsEntryIdEqual(SBinary a, SBinary b);
	LPSBinary OperationBase::ResolveNameToEID(tstring *pstrName);
	void OperationBase::OutputSBinary(SBinary lpsbin);
	HRESULT OperationBase::HrAllocAdrList(ULONG ulNumProps, _Deref_out_opt_ LPADRLIST* lpAdrList);
	LPMAPISESSION lpSession;
	Log *pLog;

private:
	LPMAPIFOLDER GetPFRoot(IMAPISession *pSession);
	LPMAPIFOLDER GetMailboxRoot(IMAPISession *pSession);
	LPMAPIFOLDER OperationBase::GetStartingFolder(IMAPISession *pSession, tstring *calculatedFolderPath);
	HRESULT OperationBase::GetSubfolderByName(LPMAPIFOLDER parentFolder, tstring folderNameToFind, LPMAPIFOLDER *returnedFolder, tstring *returnedFolderName);
	void TraverseFolders(CComPtr<IMAPISession> session, LPMAPIFOLDER baseFolder, tstring parentPath);
	HRESULT BuildServerDN(
					  _In_z_ LPCTSTR szServerName,
					  _In_z_ LPCTSTR szPost,
					  _Deref_out_z_ LPTSTR* lpszServerDN);
	std::vector<tstring> OperationBase::Split(const tstring &s, TCHAR delim);
	std::vector<tstring> &Split(const tstring &s, TCHAR delim, std::vector<tstring> &elems);
	HRESULT OperationBase::OpenDefaultMessageStore(
								_In_ LPMAPISESSION lpMAPISession,
								_Deref_out_ LPMDB* lppDefaultMDB);
	HRESULT OperationBase::CallOpenMsgStore(
						 _In_ LPMAPISESSION	lpSession,
						 _In_ ULONG_PTR		ulUIParam,
						 _In_ LPSBinary		lpEID,
						 ULONG			ulFlags,
						 _Deref_out_ LPMDB*			lpMDB);
	HRESULT OperationBase::UnicodeToAnsi(_In_z_ LPCWSTR pszW, _Out_z_cap_(cchszW) LPSTR* ppszA, size_t cchszW);
	HRESULT OperationBase::AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, size_t cchszA);
	HRESULT OperationBase::GetServerName(_In_ LPMAPISESSION lpSession, _Deref_out_opt_z_ LPTSTR* szServerName);
	bool OperationBase::CheckStringProp(_In_opt_ LPSPropValue lpProp, ULONG ulPropType);
	HRESULT OperationBase::CopyStringW(_Deref_out_z_ LPWSTR* lpszDestination, _In_z_ LPCWSTR szSource, _In_opt_ LPVOID pParent);
	LPMAPIFOLDER lpRootFolder;
	LPMAPIFOLDER lpStartingFolder;
	tstring startingPath;
	LPMDB lpAdminMDB;
	tstring *strBasePath;
	tstring *strMailbox;
	UserArgs::ActionScope nScope;
};

