#include "stdafx.h"
#include "OperationBase.h"
#include "MAPIFolders.h"
#include <algorithm>


OperationBase::OperationBase(CComPtr<IMAPISession> session)
{
	this->spSession = session;
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
