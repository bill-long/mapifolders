#include "SearchTestOperation.h"
#include "MAPIFolders.h"

SearchTestOperation::SearchTestOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log)
	:OperationBase(pstrBasePath, pstrMailbox, nScope, log)
{
}


SearchTestOperation::~SearchTestOperation()
{
}


HRESULT SearchTestOperation::Initialize(void)
{
	HRESULT hr = S_OK;

	tstring searchFolderName = tstring(_T("TestSearchFolder"));
	tstring ipmSubtreeRootName = tstring(_T("Top of Information Store"));
	tstring inboxName = tstring(_T("Calendar"));
	tstring returnedIpmSubtreeName = tstring(_T(""));
	tstring returnedSearchFolderName = tstring(_T(""));
	tstring returnedInboxName = tstring(_T(""));

	CORg(OperationBase::Initialize());

	LPMAPIFOLDER lpIpmSubtreeRoot = NULL;
	CORg(this->GetSubfolderByName(this->lpRootFolder, ipmSubtreeRootName, &lpIpmSubtreeRoot, &returnedIpmSubtreeName));
	LPMAPIFOLDER lpInboxFolder = NULL;
	CORg(this->GetSubfolderByName(lpIpmSubtreeRoot, inboxName, &lpInboxFolder, &returnedInboxName));
	LPMAPIFOLDER lpSearchFolder = NULL;
	CORg(this->GetSubfolderByName(this->lpRootFolder, searchFolderName, &lpSearchFolder, &returnedSearchFolderName));

	// Set up restriction

	SRestriction sres;
	sres.rt = RES_AND;

	SRestriction srNotFoo1;
	srNotFoo1.rt = RES_PROPERTY;
	srNotFoo1.res.resProperty.relop = RELOP_NE;
	_PV pvFoo1;
	pvFoo1.lpszW = L"IPM.Foo1";
	SPropValue sPropValFoo1;
	sPropValFoo1.dwAlignPad = 0;
	sPropValFoo1.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo1.Value = pvFoo1;
	srNotFoo1.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo1;
	srNotFoo1.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo2;
	srNotFoo2.rt = RES_PROPERTY;
	srNotFoo2.res.resProperty.relop = RELOP_NE;
	_PV pvFoo2;
	pvFoo2.lpszW = L"IPM.Foo2";
	SPropValue sPropValFoo2;
	sPropValFoo2.dwAlignPad = 0;
	sPropValFoo2.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo2.Value = pvFoo2;
	srNotFoo2.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo2;
	srNotFoo2.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo3;
	srNotFoo3.rt = RES_PROPERTY;
	srNotFoo3.res.resProperty.relop = RELOP_NE;
	_PV pvFoo3;
	pvFoo3.lpszW = L"IPM.Foo3";
	SPropValue sPropValFoo3;
	sPropValFoo3.dwAlignPad = 0;
	sPropValFoo3.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo3.Value = pvFoo3;
	srNotFoo3.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo3;
	srNotFoo3.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo4;
	srNotFoo4.rt = RES_PROPERTY;
	srNotFoo4.res.resProperty.relop = RELOP_NE;
	_PV pvFoo4;
	pvFoo4.lpszW = L"IPM.Foo4";
	SPropValue sPropValFoo4;
	sPropValFoo4.dwAlignPad = 0;
	sPropValFoo4.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo4.Value = pvFoo4;
	srNotFoo4.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo4;
	srNotFoo4.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo5;
	srNotFoo5.rt = RES_PROPERTY;
	srNotFoo5.res.resProperty.relop = RELOP_NE;
	_PV pvFoo5;
	pvFoo5.lpszW = L"IPM.Foo5";
	SPropValue sPropValFoo5;
	sPropValFoo5.dwAlignPad = 0;
	sPropValFoo5.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo5.Value = pvFoo5;
	srNotFoo5.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo5;
	srNotFoo5.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo6;
	srNotFoo6.rt = RES_PROPERTY;
	srNotFoo6.res.resProperty.relop = RELOP_NE;
	_PV pvFoo6;
	pvFoo6.lpszW = L"IPM.Foo6";
	SPropValue sPropValFoo6;
	sPropValFoo6.dwAlignPad = 0;
	sPropValFoo6.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo6.Value = pvFoo6;
	srNotFoo6.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo6;
	srNotFoo6.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo7;
	srNotFoo7.rt = RES_PROPERTY;
	srNotFoo7.res.resProperty.relop = RELOP_NE;
	_PV pvFoo7;
	pvFoo7.lpszW = L"IPM.Foo7";
	SPropValue sPropValFoo7;
	sPropValFoo7.dwAlignPad = 0;
	sPropValFoo7.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo7.Value = pvFoo7;
	srNotFoo7.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo7;
	srNotFoo7.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo8;
	srNotFoo8.rt = RES_PROPERTY;
	srNotFoo8.res.resProperty.relop = RELOP_NE;
	_PV pvFoo8;
	pvFoo8.lpszW = L"IPM.Foo8";
	SPropValue sPropValFoo8;
	sPropValFoo8.dwAlignPad = 0;
	sPropValFoo8.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo8.Value = pvFoo8;
	srNotFoo8.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo8;
	srNotFoo8.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo9;
	srNotFoo9.rt = RES_PROPERTY;
	srNotFoo9.res.resProperty.relop = RELOP_NE;
	_PV pvFoo9;
	pvFoo9.lpszW = L"IPM.Foo9";
	SPropValue sPropValFoo9;
	sPropValFoo9.dwAlignPad = 0;
	sPropValFoo9.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo9.Value = pvFoo9;
	srNotFoo9.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo9;
	srNotFoo9.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo10;
	srNotFoo10.rt = RES_PROPERTY;
	srNotFoo10.res.resProperty.relop = RELOP_NE;
	_PV pvFoo10;
	pvFoo10.lpszW = L"IPM.Foo10";
	SPropValue sPropValFoo10;
	sPropValFoo10.dwAlignPad = 0;
	sPropValFoo10.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo10.Value = pvFoo10;
	srNotFoo10.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo10;
	srNotFoo10.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo11;
	srNotFoo11.rt = RES_PROPERTY;
	srNotFoo11.res.resProperty.relop = RELOP_NE;
	_PV pvFoo11;
	pvFoo11.lpszW = L"IPM.Foo11";
	SPropValue sPropValFoo11;
	sPropValFoo11.dwAlignPad = 0;
	sPropValFoo11.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo11.Value = pvFoo11;
	srNotFoo11.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo11;
	srNotFoo11.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction srNotFoo12;
	srNotFoo12.rt = RES_PROPERTY;
	srNotFoo12.res.resProperty.relop = RELOP_NE;
	_PV pvFoo12;
	pvFoo12.lpszW = L"IPM.Foo12";
	SPropValue sPropValFoo12;
	sPropValFoo12.dwAlignPad = 0;
	sPropValFoo12.ulPropTag = PR_MESSAGE_CLASS_W;
	sPropValFoo12.Value = pvFoo12;
	srNotFoo12.res.resProperty.lpProp = (LPSPropValue)&sPropValFoo12;
	srNotFoo12.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_W;

	SRestriction* restrictions = new SRestriction[12];
	restrictions[0] = srNotFoo1;
	restrictions[1] = srNotFoo2;
	restrictions[2] = srNotFoo3;
	restrictions[3] = srNotFoo4;
	restrictions[4] = srNotFoo5;
	restrictions[5] = srNotFoo6;
	restrictions[6] = srNotFoo7;
	restrictions[7] = srNotFoo8;
	restrictions[8] = srNotFoo9;
	restrictions[9] = srNotFoo10;
	restrictions[10] = srNotFoo11;
	restrictions[11] = srNotFoo12;

	sres.res.resOr.cRes = 12;
	sres.res.resOr.lpRes = restrictions;

	// Does the search folder exist?
	if (lpSearchFolder == NULL) {
		LPTSTR lpFolderName = L"TestSearchFolder";
		CORg(this->lpRootFolder->CreateFolder(FOLDER_SEARCH, lpFolderName, NULL, NULL, MAPI_UNICODE, &lpSearchFolder));

				SPropValue * rgprops = NULL;
		ULONG rgTags[] = { 1, PR_ENTRYID };
		ULONG cCount = 0;
		lpIpmSubtreeRoot->GetProps((LPSPropTagArray)rgTags, NULL, &cCount, &rgprops);
		LPENTRYLIST lpEntryList = NULL;
		MAPIAllocateBuffer(sizeof(ENTRYLIST), (void**)&lpEntryList);
		lpEntryList->cValues = 1;
		lpEntryList->lpbin = &rgprops->Value.bin;
		lpSearchFolder->SetSearchCriteria((LPSRestriction)&sres, lpEntryList, BACKGROUND_SEARCH | RECURSIVE_SEARCH | 0x00020000); // NON_CONTENT_INDEXED_SEARCH
	}

	SYSTEMTIME stStart;
	GetLocalTime(&stStart);
	tcout << "\n";
	tcout << "Retrieve items from search folder.\n";
	tcout << "Start: " << stStart.wHour << ":" << stStart.wMinute << ":" << stStart.wSecond << "." << stStart.wMilliseconds << "\n";

	ULONG ulFlags = MAPI_UNICODE;
	LPMAPITABLE lpContentsTable = NULL;
	LPSPropTagArray lpColSet = NULL;
	CORg(lpSearchFolder->GetContentsTable(ulFlags, &lpContentsTable));
	CORg(lpContentsTable->QueryColumns(0, &lpColSet));
	CORg(lpContentsTable->SeekRow(BOOKMARK_BEGINNING, 0, NULL));

	LPSRowSet pRows = NULL;
	ULONG rowsProcessed = 0;
	if (!FAILED(hr)) for (;;)
	{
		hr = S_OK;
		if (pRows) FreeProws(pRows);
		pRows = NULL;

		CORg(lpContentsTable->QueryRows(1000, NULL, &pRows));
		if (FAILED(hr) || !pRows || !pRows->cRows) break;

		rowsProcessed += pRows->cRows;
	}

	SYSTEMTIME stEnd;
	GetLocalTime(&stEnd);
	tcout << "End: " << stEnd.wHour << ":" << stEnd.wMinute << ":" << stEnd.wSecond << "." << stEnd.wMilliseconds << "\n";
	tcout << "Rows processed: " << rowsProcessed << "\n";

	if (pRows) FreeProws(pRows);
	if (lpColSet) MAPIFreeBuffer(lpColSet);
	if (lpContentsTable) lpContentsTable->Release();

	// Now try a restrict instead

	GetLocalTime(&stStart);
	tcout << "\n";
	tcout << "Retrieve items by restriction.\n";
	tcout << "Start: " << stStart.wHour << ":" << stStart.wMinute << ":" << stStart.wSecond << "." << stStart.wMilliseconds << "\n";

	lpContentsTable = NULL;
	lpColSet = NULL;
	CORg(lpInboxFolder->GetContentsTable(ulFlags, &lpContentsTable));
	CORg(lpContentsTable->Restrict((LPSRestriction)&sres, NULL));
	CORg(lpContentsTable->QueryColumns(0, &lpColSet));
	CORg(lpContentsTable->SeekRow(BOOKMARK_BEGINNING, 0, NULL));

	SYSTEMTIME stSeekRow;
	GetLocalTime(&stSeekRow);
	tcout << "SeekRow complete: " << stSeekRow.wHour << ":" << stSeekRow.wMinute << ":" << stSeekRow.wSecond << "." << stSeekRow.wMilliseconds << "\n";

	pRows = NULL;
	rowsProcessed = 0;
	if (!FAILED(hr)) for (;;)
	{
		hr = S_OK;
		if (pRows) FreeProws(pRows);
		pRows = NULL;

		CORg(lpContentsTable->QueryRows(1000, NULL, &pRows));
		if (FAILED(hr) || !pRows || !pRows->cRows) break;

		rowsProcessed += pRows->cRows;
	}

	SYSTEMTIME stRestrictEnd;
	GetLocalTime(&stRestrictEnd);
	tcout << "End: " << stRestrictEnd.wHour << ":" << stRestrictEnd.wMinute << ":" << stRestrictEnd.wSecond << "." << stRestrictEnd.wMilliseconds << "\n";
	tcout << "Rows processed: " << rowsProcessed << "\n";

	if (pRows) FreeProws(pRows);
	if (lpColSet) MAPIFreeBuffer(lpColSet);
	if (lpContentsTable) lpContentsTable->Release();

	// Now just read the items from the folder

	GetLocalTime(&stStart);
	tcout << "\n";
	tcout << "Retrieve items from folder.\n";
	tcout << "Start: " << stStart.wHour << ":" << stStart.wMinute << ":" << stStart.wSecond << "." << stStart.wMilliseconds << "\n";

	lpContentsTable = NULL;
	lpColSet = NULL;
	CORg(lpInboxFolder->GetContentsTable(ulFlags, &lpContentsTable));
	CORg(lpContentsTable->QueryColumns(0, &lpColSet));
	CORg(lpContentsTable->SeekRow(BOOKMARK_BEGINNING, 0, NULL));

	GetLocalTime(&stSeekRow);
	tcout << "SeekRow complete: " << stSeekRow.wHour << ":" << stSeekRow.wMinute << ":" << stSeekRow.wSecond << "." << stSeekRow.wMilliseconds << "\n";

	pRows = NULL;
	rowsProcessed = 0;
	if (!FAILED(hr)) for (;;)
	{
		hr = S_OK;
		if (pRows) FreeProws(pRows);
		pRows = NULL;

		CORg(lpContentsTable->QueryRows(1000, NULL, &pRows));
		if (FAILED(hr) || !pRows || !pRows->cRows) break;

		rowsProcessed += pRows->cRows;
	}

	SYSTEMTIME stReadContentsEnd;
	GetLocalTime(&stReadContentsEnd);
	tcout << "End: " << stReadContentsEnd.wHour << ":" << stReadContentsEnd.wMinute << ":" << stReadContentsEnd.wSecond << "." << stReadContentsEnd.wMilliseconds << "\n";
	tcout << "Rows processed: " << rowsProcessed << "\n";

Error:
	return hr;
}


void SearchTestOperation::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
{
}