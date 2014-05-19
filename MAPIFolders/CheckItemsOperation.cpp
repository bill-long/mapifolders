#include "CheckItemsOperation.h"
#include "MAPIFolders.h"
#include <MAPIDefS.h>

CheckItemsOperation::CheckItemsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, bool fix)
	:OperationBase(pstrBasePath, pstrMailbox, nScope, log)
{
	this->fix = fix;
}

CheckItemsOperation::~CheckItemsOperation()
{
}

HRESULT CheckItemsOperation::Initialize()
{
	HRESULT hr = S_OK;
	CORg(OperationBase::Initialize());

Cleanup:
	return hr;
Error:
	goto Cleanup;
}

void CheckItemsOperation::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
{
	HRESULT hr;
	ULONG ulFlags = MAPI_UNICODE; // Might need to change this later for MAPI_ASSOCIATED and/or SHOW_SOFT_DELETES
	LPMAPITABLE lpContentsTable = NULL;
	LPSRowSet pRows = NULL;
	LPSPropTagArray lpColSet = NULL;

	LPSPropValue lpFolderType = NULL;
	CORg(HrGetOneProp(folder, PR_FOLDER_TYPE, &lpFolderType));
	if (lpFolderType->Value.l == FOLDER_SEARCH)
	{
		*pLog << "Skipping search folder: " << folderPath << "\n";
		goto Cleanup;
	}

	// This initialization can't go in Initialize() because MAPI isn't
	// initialized at that point. Is there something we can do so we
	// don't have to do this over and over?
	ULONG cPropsToRemove = 2;
	ULONG propsToRemove[2] = { 0x12041002, 0x12051002 };
	this->lpPropsToRemove = NULL;
	CORg(MAPIAllocateBuffer(CbNewSPropTagArray(cPropsToRemove),
		(LPVOID*)&lpPropsToRemove));
	lpPropsToRemove->cValues = cPropsToRemove;
	for (INT x = 0; x < cPropsToRemove; x++)
	{
		lpPropsToRemove->aulPropTag[x] = propsToRemove[x];
	}
	// End of initialization

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

			*pLog << "    Checking item: " << (pwzSubject ? pwzSubject : _T("<NULL>")) << "\n";

			ULONG ulObjType = NULL;
			LPMAPIPROP lpMessage = NULL;
			folder->OpenEntry(eid.cb, (LPENTRYID)eid.lpb, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *)&lpMessage);

			ULONG cCount = 0;
			SPropValue *rgprops = NULL;
			CORg(lpMessage->GetProps(lpPropsToRemove, MAPI_UNICODE, &cCount, &rgprops));
			bool foundSomeProps = false;
			for (ULONG y = 0; y < cCount; y++)
			{
				if (PROP_TYPE(rgprops[y].ulPropTag) != PT_ERROR)
				{
					*pLog << "        Found property: " << std::hex << rgprops[y].ulPropTag << "\n";
					foundSomeProps = true;
				}
			}

			if (foundSomeProps && this->fix)
			{
				*pLog << "    Deleting these properties... ";
				// If we found any one of the specified props, just send a delete for all of them
				LPSPropProblemArray problemArray = NULL;
				CORg(lpMessage->DeleteProps(lpPropsToRemove, &problemArray));
				CORg(lpMessage->SaveChanges(NULL));
				*pLog << "Done.\n";
			}
		}
	}

Cleanup:
	MAPIFreeBuffer(lpFolderType);
	if (pRows) FreeProws(pRows);
	if (lpColSet) MAPIFreeBuffer(lpColSet);
	if (lpContentsTable) lpContentsTable->Release();
	return;

Error:
	goto Cleanup;
}