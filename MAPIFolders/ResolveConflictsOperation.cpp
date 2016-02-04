#include "ResolveConflictsOperation.h"
#include "MAPIFolders.h"
#include <MAPIDefS.h>

ResolveConflictsOperation::ResolveConflictsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, bool useAdmin)
	:ItemOperationBase(pstrBasePath, pstrMailbox, nScope, log, useAdmin)
{
}


ResolveConflictsOperation::~ResolveConflictsOperation()
{
}

HRESULT ResolveConflictsOperation::Initialize()
{
	HRESULT hr = S_OK;
	std::vector<tstring> splitPropStrings;

	CORg(OperationBase::Initialize());

	ULONG cMsgProps = 2;
	ULONG msgProps[2] = { PR_MSG_STATUS, PR_PARENT_ENTRYID };
	this->lpMsgProps = NULL;
	CORg(MAPIAllocateBuffer(CbNewSPropTagArray(cMsgProps),
		(LPVOID*)&lpMsgProps));
	lpMsgProps->cValues = cMsgProps;
	for (INT x = 0; x < cMsgProps; x++)
	{
		lpMsgProps->aulPropTag[x] = msgProps[x];
	}

	ULONG cAttachTableProps = 4;
	ULONG attachTableProps[4] = { PR_ATTACH_NUM, PR_ATTACH_FILENAME, PR_ATTACH_METHOD, PR_IN_CONFLICT };
	this->lpAttachProps = NULL;
	CORg(MAPIAllocateBuffer(CbNewSPropTagArray(cAttachTableProps),
		(LPVOID*)&lpAttachProps));
	lpAttachProps->cValues = cAttachTableProps;
	for (INT x = 0; x < cAttachTableProps; x++)
	{
		lpAttachProps->aulPropTag[x] = attachTableProps[x];
	}

Cleanup:
	return hr;
Error:
	goto Cleanup;
}

void ResolveConflictsOperation::ProcessItem(LPMAPIPROP lpMessage, LPCTSTR pwzSubject)
{
	HRESULT hr = S_OK;
	LPMAPITABLE lpAttachTable = NULL;
	LPSPropTagArray lpColSet = NULL;
	LPSRowSet pRows = NULL;
	LPATTACH lpAttach = NULL;
	LPMESSAGE lpEmbeddedMessage = NULL;

	*pLog << "Checking item: " << (pwzSubject ? pwzSubject : _T("<NULL>")) << "\n";

	ULONG cCount = 0;
	SPropValue *rgprops = NULL;
	CORg(lpMessage->GetProps(lpMsgProps, MAPI_UNICODE, &cCount, &rgprops));

	// Are we in conflict status?
	if (PROP_TYPE(rgprops[0].ulPropTag) != PT_ERROR && (rgprops[0].Value.l & 0x800) == 0x800)
	{
		*pLog << "    This message is in a conflict state." << "\n";
		CORg(((LPMESSAGE)lpMessage)->GetAttachmentTable(0, &lpAttachTable));
		CORg(lpAttachTable->QueryColumns(0, &lpColSet));
		CORg(lpAttachTable->SeekRow(BOOKMARK_BEGINNING, 0, NULL));
		CORg(lpAttachTable->QueryRows(1000, NULL, &pRows));
		if (FAILED(hr) || !pRows || !pRows->cRows) goto Cleanup;
		for (ULONG currentRow = 0; currentRow < pRows->cRows; currentRow++)
		{
			if (lpAttach) lpAttach->Release();
			SRow prow = pRows->aRow[currentRow];

			CORg(((LPMESSAGE)lpMessage)->OpenAttach(currentRow, NULL, MAPI_BEST_ACCESS, &lpAttach));
			ULONG cCount = 0;
			SPropValue *rgAttachProps = NULL;
			CORg(lpAttach->GetProps(lpAttachProps, MAPI_UNICODE, &cCount, &rgAttachProps));

			bool isConflictAttach = false;
			for (ULONG y = 0; y < cCount; y++)
			{
				if (PROP_TYPE(rgAttachProps[y].ulPropTag) != PT_ERROR)
				{
					if (rgAttachProps[y].ulPropTag == PR_IN_CONFLICT)
					{
						if (rgAttachProps[y].Value.b == true)
						{
							isConflictAttach = true;
							break;
						}
					}
				}
			}

			MAPIFreeBuffer(rgAttachProps);

			if (isConflictAttach)
			{
				CORg(lpAttach->OpenProperty(PR_ATTACH_DATA_OBJ, (LPIID)&IID_IMessage, 0, MAPI_MODIFY, (LPUNKNOWN *)&lpEmbeddedMessage));
				CORg(lpEmbeddedMessage->CopyTo(0, NULL, NULL, NULL, NULL, (LPGUID)&IID_IMessage, lpMessage, 0, NULL));
				CORg(lpMessage->SaveChanges(0));
				lpEmbeddedMessage->Release();

				*pLog << "    Conflict resolved." << "\n";
				break;
			}
		}
	}

	MAPIFreeBuffer(rgprops);

Cleanup:
	if (pRows) FreeProws(pRows);
	if (lpColSet) MAPIFreeBuffer(lpColSet);
	if (lpAttachTable) lpAttachTable->Release();
	if (lpAttach) lpAttach->Release();
	return;

Error:
	goto Cleanup;
}