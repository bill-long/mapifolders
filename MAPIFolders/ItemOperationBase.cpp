#include "ItemOperationBase.h"
#include "MAPIFolders.h"
#include <MAPIDefS.h>


ItemOperationBase::ItemOperationBase(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, bool useAdmin)
	:OperationBase(pstrBasePath, pstrMailbox, nScope, log, useAdmin)
{
}


ItemOperationBase::~ItemOperationBase()
{
}


HRESULT ItemOperationBase::Initialize()
{
	HRESULT hr = S_OK;
	CORg(OperationBase::Initialize());

Cleanup:
	return hr;
Error:
	goto Cleanup;
}


void ItemOperationBase::ProcessItem(LPMAPIPROP item, LPCTSTR itemSubject)
{
	return;
}


void ItemOperationBase::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
{
	HRESULT hr;
	ULONG ulFlags = MAPI_UNICODE; // Might need to change this later for MAPI_ASSOCIATED and/or SHOW_SOFT_DELETES
	LPMAPITABLE lpContentsTable = NULL;
	LPSRowSet pRows = NULL;
	LPSPropTagArray lpColSet = NULL;
	LPMAPIPROP lpMessage = NULL;

	LPSPropValue lpFolderType = NULL;
	CORg(HrGetOneProp(folder, PR_FOLDER_TYPE, &lpFolderType));
	if (lpFolderType->Value.l == FOLDER_SEARCH)
	{
		*pLog << "Skipping search folder: " << folderPath << "\n";
		goto Cleanup;
	}

	*pLog << "Checking items in folder: " << folderPath.c_str() << "\n";

	if (this->strMailbox == NULL)
	{
		LPSPropValue lpReplicaServer = NULL;
		CORg(HrGetOneProp(folder, PR_REPLICA_SERVER, &lpReplicaServer));
		*pLog << "    Replica server/mailbox: " << lpReplicaServer->Value.LPSZ << "\n";
		MAPIFreeBuffer(lpReplicaServer);
	}

	CORg(folder->GetContentsTable(ulFlags, &lpContentsTable));
	CORg(lpContentsTable->QueryColumns(0, &lpColSet));
	CORg(lpContentsTable->SeekRow(BOOKMARK_BEGINNING, 0, NULL));

	ULONG rowsProcessed = 0;
	if (!FAILED(hr)) for (;;)
	{
		hr = S_OK;
		if (pRows) FreeProws(pRows);
		pRows = NULL;

		CORg(lpContentsTable->QueryRows(1000, NULL, &pRows));
		if (FAILED(hr) || !pRows || !pRows->cRows) break;

		for (ULONG currentRow = 0; currentRow < pRows->cRows; currentRow++)
		{
			if (lpMessage) lpMessage->Release();
			lpMessage = NULL;

			SRow prow = pRows->aRow[currentRow];
			LPCTSTR pwzSubject = NULL;
			SBinary eid = { 0 };
			for (ULONG x = 0; x < prow.cValues; x++)
			{
				if (prow.lpProps[x].ulPropTag == PR_SUBJECT)
				{
					pwzSubject = prow.lpProps[x].Value.LPSZ;
				}
				else if (prow.lpProps[x].ulPropTag == PR_ENTRYID)
				{
					eid = prow.lpProps[x].Value.bin;
				}
			}

			ULONG ulObjType = NULL;
			hr = folder->OpenEntry(eid.cb, (LPENTRYID)eid.lpb, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *)&lpMessage);
			if (hr != S_OK)
			{
				*pLog << "        Error opening item: " << std::hex << hr << "\n";
				continue;
			}

			this->ProcessItem(lpMessage, pwzSubject);

		}
	}

Cleanup:
	MAPIFreeBuffer(lpFolderType);
	if (pRows) FreeProws(pRows);
	if (lpColSet) MAPIFreeBuffer(lpColSet);
	if (lpContentsTable) lpContentsTable->Release();
	if (lpMessage) lpMessage->Release();
	return;

Error:
	goto Cleanup;
}