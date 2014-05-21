#include "CheckItemsOperation.h"
#include "MAPIFolders.h"
#include <MAPIDefS.h>

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