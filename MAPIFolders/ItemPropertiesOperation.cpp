#include "ItemPropertiesOperation.h"
#include "MAPIFolders.h"
#include <MAPIDefS.h>

ItemPropertiesOperation::ItemPropertiesOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, tstring *pstrPropTags)
	:OperationBase(pstrBasePath, pstrMailbox, nScope, log)
{
	this->pstrPropTags = pstrPropTags;
}


ItemPropertiesOperation::~ItemPropertiesOperation()
{
}

HRESULT ItemPropertiesOperation::Initialize()
{
	HRESULT hr = S_OK;
	std::vector<tstring> splitPropStrings;

	CORg(OperationBase::Initialize());

	splitPropStrings = Split(this->pstrPropTags->c_str(), (_T(',')));
	this->cproptags = splitPropStrings.size();
	this->proptags = new ULONG[this->cproptags];
	for (INT x = 0; x < this->cproptags; x++)
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
			this->proptags[x] = ptag;
		}
	}

Cleanup:
	return hr;
Error:
	goto Cleanup;
}

void ItemPropertiesOperation::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
{
	HRESULT hr;
	ULONG ulFlags = MAPI_UNICODE; // Might need to change this later for MAPI_ASSOCIATED and/or SHOW_SOFT_DELETES
	LPMAPITABLE lpContentsTable = NULL;
	LPSRowSet pRows = NULL;
	LPSPropTagArray lpColSet = NULL;

	LPSPropTagArray lpPropsToGet = NULL;
	CORg(MAPIAllocateBuffer(CbNewSPropTagArray(this->cproptags),
		(LPVOID*)&lpPropsToGet));
	lpPropsToGet->cValues = this->cproptags;
	for (INT x = 0; x < this->cproptags; x++)
	{
		lpPropsToGet->aulPropTag[x] = this->proptags[x];
	}

	*pLog << "Checking items in folder: " << folderPath.c_str() << "\n";

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

			*pLog << "Checking item: " << (pwzSubject ? pwzSubject : _T("<NULL>")) << "\n";

			ULONG ulObjType = NULL;
			LPMAPIPROP lpMessage = NULL;
			folder->OpenEntry(eid.cb, (LPENTRYID)eid.lpb, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *) &lpMessage);

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
		}
	}

Cleanup:
	if (pRows) FreeProws(pRows);
	return;

Error:
	goto Cleanup;
}