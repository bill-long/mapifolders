#include "ResolveConflictsOperation.h"
#include "MAPIFolders.h"
#include <MAPIDefS.h>

ResolveConflictsOperation::ResolveConflictsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log)
	:ItemOperationBase(pstrBasePath, pstrMailbox, nScope, log)
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

	ULONG cAttachTableProps = 3;
	ULONG attachTableProps[3] = { PR_ATTACH_NUM, PR_ATTACH_FILENAME, PR_ATTACH_METHOD };
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
	ULONG objType = NULL;
	LPMAPIFOLDER lpParentFolder = NULL;
	LPMESSAGE lpNewMessage = NULL;

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
		CORg(lpAttachTable->QueryRows(1, NULL, &pRows));
		if (FAILED(hr) || !pRows || !pRows->cRows) goto Cleanup;

		CORg(((LPMESSAGE)lpMessage)->OpenAttach(0, NULL, MAPI_BEST_ACCESS, &lpAttach));
		CORg(lpAttach->OpenProperty(PR_ATTACH_DATA_OBJ, (LPIID)&IID_IMessage, 0, MAPI_MODIFY, (LPUNKNOWN *)&lpEmbeddedMessage));
		CORg(lpEmbeddedMessage->CopyTo(0, NULL, NULL, NULL, NULL, (LPGUID)&IID_IMessage, lpMessage, 0, NULL));
		CORg(lpMessage->SaveChanges(0));
		
		*pLog << "    Conflict resolved." << "\n";
	}

Cleanup:
	if (pRows) FreeProws(pRows);
	if (lpColSet) MAPIFreeBuffer(lpColSet);
	if (lpAttachTable) lpAttachTable->Release();
	if (lpAttach) lpAttach->Release();
	if (lpNewMessage) lpNewMessage->Release();
	if (lpParentFolder) lpParentFolder->Release();
	return;

Error:
	goto Cleanup;
}