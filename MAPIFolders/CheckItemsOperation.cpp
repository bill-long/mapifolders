#include "CheckItemsOperation.h"
#include "MAPIFolders.h"
#include <MAPIDefS.h>

DEFINE_OLEGUID(IID_IMAPITable, 0x00020301, 0, 0);
DEFINE_OLEGUID(IID_IMessage, 0x00020307, 0, 0);

CheckItemsOperation::CheckItemsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, bool fix)
	:ItemOperationBase(pstrBasePath, pstrMailbox, nScope, log)
{
	this->fix = fix;
}

CheckItemsOperation::~CheckItemsOperation()
{
}

HRESULT CheckItemsOperation::Initialize()
{
	HRESULT hr = S_OK;
	CORg(ItemOperationBase::Initialize());

	// Base class Initialize() MUST be called before we use
	// MAPIAllocateBuffer below, so that MAPI is initialized.
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

	ULONG cAttachProps = 1;
	ULONG attachProps[1] = { PR_HASATTACH };
	this->lpAttachProps = NULL;
	CORg(MAPIAllocateBuffer(CbNewSPropTagArray(cAttachProps), 
		(LPVOID*)&lpAttachProps));
	lpAttachProps->cValues = cAttachProps;
	for (INT x = 0; x < cAttachProps; x++) 
	{
		lpAttachProps->aulPropTag[x] = attachProps[x];
	}

	ULONG cAttachTableProps = 3;
	ULONG attachTableProps[3] = { PR_ATTACH_NUM, PR_ATTACH_FILENAME, PR_ATTACH_METHOD };
	this->lpAttachTableProps = NULL;
	CORg(MAPIAllocateBuffer(CbNewSPropTagArray(cAttachTableProps),
		(LPVOID*)&lpAttachTableProps));
	lpAttachTableProps->cValues = cAttachTableProps;
	for (INT x = 0; x < cAttachTableProps; x++)
	{
		lpAttachTableProps->aulPropTag[x] = attachTableProps[x];
	}
	// End of initialization

Cleanup:
	return hr;
Error:
	goto Cleanup;
}


void CheckItemsOperation::ProcessItem(LPMAPIPROP lpMessage, LPCTSTR pwzSubject)
{
	HRESULT hr = S_OK;

	*pLog << "    Checking item: " << (pwzSubject ? pwzSubject : _T("<NULL>")) << "\n";

	ULONG cCount = 0;
	SPropValue *rgAttachProps = NULL;
	CORg(lpMessage->GetProps(lpAttachProps, MAPI_UNICODE, &cCount, &rgAttachProps));
	if (PROP_TYPE(rgAttachProps[0].ulPropTag) != PT_ERROR)
	{
		if (rgAttachProps[0].Value.b)
		{
			this->ProcessAttachmentsRecursive(lpMessage);
		}
	}

	cCount = 0;
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

	MAPIFreeBuffer(rgprops);
	MAPIFreeBuffer(rgAttachProps);

	if (foundSomeProps && this->fix)
	{
		*pLog << "        Deleting these properties... ";
		// If we found any one of the specified props, just send a delete for all of them
		LPSPropProblemArray problemArray = NULL;
		CORg(lpMessage->DeleteProps(lpPropsToRemove, &problemArray));
		CORg(lpMessage->SaveChanges(NULL));
		MAPIFreeBuffer(problemArray);
		*pLog << "Done.\n";
	}

Cleanup:
	return;

Error:
	goto Cleanup;
}

bool CheckItemsOperation::ProcessAttachmentsRecursive(LPMAPIPROP lpMAPIProp)
{
	HRESULT hr = S_OK;
	LPMAPITABLE lpTable = NULL;
	LPSPropTagArray lpColSet = NULL;
	LPSRowSet pRows = NULL;
	LPATTACH lpAttach = NULL;
	LPMESSAGE lpEmbeddedMessage = NULL;
	bool bUpdatedItem = false;

	CORg(((LPMESSAGE)lpMAPIProp)->GetAttachmentTable(0, &lpTable));
	CORg(lpTable->QueryColumns(0, &lpColSet));
	CORg(lpTable->SeekRow(BOOKMARK_BEGINNING, 0, NULL));

	ULONG rowsProcessed = 0;
	if (!FAILED(hr)) for (;;)
	{
		hr = S_OK;
		if (pRows) FreeProws(pRows);
		pRows = NULL;

		CORg(lpTable->QueryRows(1000, NULL, &pRows));
		if (FAILED(hr) || !pRows || !pRows->cRows) break;

		for (ULONG currentRow = 0; currentRow < pRows->cRows; currentRow++)
		{
			lpAttach = NULL;

			SRow prow = pRows->aRow[currentRow];
			LPCTSTR pwzAttachName = NULL;
			ULONG ulAttachMethod = NULL;
			ULONG ulAttachNum = NULL;

			CORg(((LPMESSAGE)lpMAPIProp)->OpenAttach(currentRow, NULL, MAPI_BEST_ACCESS, &lpAttach));
			ULONG cCount = 0;
			SPropValue *rgprops = NULL;
			CORg(lpAttach->GetProps(lpAttachTableProps, MAPI_UNICODE, &cCount, &rgprops));

			for (ULONG y = 0; y < cCount; y++)
			{
				if (PROP_TYPE(rgprops[y].ulPropTag) != PT_ERROR)
				{
					if (rgprops[y].ulPropTag == PR_ATTACH_FILENAME)
					{
						pwzAttachName = rgprops[y].Value.LPSZ;
					}
					else if (rgprops[y].ulPropTag == PR_ATTACH_NUM)
					{
						ulAttachNum = rgprops[y].Value.l;
					}
					else if (rgprops[y].ulPropTag == PR_ATTACH_METHOD)
					{
						ulAttachMethod = rgprops[y].Value.l;
					}
				}
			}

			if (ulAttachMethod == ATTACH_EMBEDDED_MSG)
			{
				CORg(lpAttach->OpenProperty(PR_ATTACH_DATA_OBJ, (LPIID)&IID_IMessage, 0, MAPI_MODIFY, (LPUNKNOWN *)&lpEmbeddedMessage));

				// Recurse first
				bUpdatedItem = this->ProcessAttachmentsRecursive(lpEmbeddedMessage);

				// Now check this one
				*pLog << "        Checking attachment: " << (pwzAttachName ? pwzAttachName : _T("<NULL>")) << "\n";
				cCount = 0;
				SPropValue *rgprops = NULL;
				CORg(lpEmbeddedMessage->GetProps(lpPropsToRemove, MAPI_UNICODE, &cCount, &rgprops));
				bool foundSomeProps = false;
				for (ULONG y = 0; y < cCount; y++)
				{
					if ((rgprops[y].ulPropTag == lpPropsToRemove->aulPropTag[0]) || (rgprops[y].ulPropTag == lpPropsToRemove->aulPropTag[1]))
					{
						*pLog << "            Found property: " << std::hex << rgprops[y].ulPropTag << "\n";
						foundSomeProps = true;
					}
				}

				MAPIFreeBuffer(rgprops);

				if (foundSomeProps && this->fix)
				{
					*pLog << "            Deleting these properties... ";
					// If we found any one of the specified props, just send a delete for all of them
					LPSPropProblemArray problemArray = NULL;
					CORg(lpEmbeddedMessage->DeleteProps(lpPropsToRemove, &problemArray));
					bUpdatedItem = true;
					MAPIFreeBuffer(problemArray);
					*pLog << "Done.\n";
				}

				if (bUpdatedItem) 
				{
					CORg(lpEmbeddedMessage->SaveChanges(NULL));
					CORg(lpAttach->SaveChanges(NULL));
					CORg(lpMAPIProp->SaveChanges(NULL));
				}

				lpEmbeddedMessage->Release();
				lpEmbeddedMessage = NULL;
			}

			lpAttach->Release();
			lpAttach = NULL;
		}
	}

Cleanup:
	if (pRows) FreeProws(pRows);
	if (lpColSet) MAPIFreeBuffer(lpColSet);
	if (lpTable) lpTable->Release();
	if (lpAttach) lpAttach->Release();
	return bUpdatedItem;

Error:
	goto Cleanup;
}