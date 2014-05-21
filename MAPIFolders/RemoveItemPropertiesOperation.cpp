#include "RemoveItemPropertiesOperation.h"
#include "MAPIFolders.h"
#include <MAPIDefS.h>

RemoveItemPropertiesOperation::RemoveItemPropertiesOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, tstring *pstrPropTags)
	:ItemOperationBase(pstrBasePath, pstrMailbox, nScope, log)
{
	this->pstrPropTags = pstrPropTags;
}


RemoveItemPropertiesOperation::~RemoveItemPropertiesOperation()
{
}

HRESULT RemoveItemPropertiesOperation::Initialize()
{
	HRESULT hr = S_OK;
	std::vector<tstring> splitPropStrings;

	CORg(OperationBase::Initialize());

	splitPropStrings = Split(this->pstrPropTags->c_str(), (_T(',')));
	lpPropsToGet = NULL;
	CORg(MAPIAllocateBuffer(CbNewSPropTagArray(splitPropStrings.size()),
		(LPVOID*)&lpPropsToGet));
	this->lpPropsToGet->cValues = splitPropStrings.size();
	for (INT x = 0; x < this->lpPropsToGet->cValues; x++)
	{
		ULONG ptag = std::wcstoul(splitPropStrings.at(x).c_str(), NULL, 16);
		if (ptag == 0xffffffff)
		{
			hr = E_FAIL;
			*pLog << "Failed to parse property tag: " << splitPropStrings.at(x);
			goto Error;
		}
		else
		{
			lpPropsToGet->aulPropTag[x] = ptag;
		}
	}

Cleanup:
	return hr;
Error:
	goto Cleanup;
}

void RemoveItemPropertiesOperation::ProcessItem(LPMAPIPROP lpMessage, LPCTSTR pwzSubject)
{
	HRESULT hr = S_OK;

			*pLog << "Checking item: " << (pwzSubject ? pwzSubject : _T("<NULL>")) << "\n";

			ULONG cCount = 0;
			SPropValue *rgprops = NULL;
			CORg(lpMessage->GetProps(lpPropsToGet, MAPI_UNICODE, &cCount, &rgprops));
			for (ULONG y = 0; y < cCount; y++)
			{
				if (PROP_TYPE(rgprops[y].ulPropTag) != PT_ERROR)
				{
					*pLog << "    Found one or more of the specified props. Deleting these props..." << "\n";
					// If we found any one of the specified props, just send a delete for all of them
					// Then break out of the loop
					LPSPropProblemArray problemArray = NULL;
					CORg(lpMessage->DeleteProps(lpPropsToGet, &problemArray));
					CORg(lpMessage->SaveChanges(NULL));
					break;
				}
			}


Cleanup:
	return;

Error:
	goto Cleanup;
}