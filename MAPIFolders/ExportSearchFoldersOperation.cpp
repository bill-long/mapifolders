#include "ExportSearchFoldersOperation.h"
#include "MAPIFolders.h"


ExportSearchFoldersOperation::ExportSearchFoldersOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, Log *exportFile, bool useAdmin)
	:OperationBase(pstrBasePath, pstrMailbox, nScope, log, useAdmin)
{
	this->exportFile = exportFile;
}


ExportSearchFoldersOperation::~ExportSearchFoldersOperation()
{
}


HRESULT ExportSearchFoldersOperation::Initialize(void)
{
	HRESULT hr = S_OK;
	std::vector<tstring> splitPropStrings;

	CORg(OperationBase::Initialize());

	*exportFile << _T("Folder Path,Search Criteria\n");

Error:
	return hr;
}


void ExportSearchFoldersOperation::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
{
	HRESULT hr = S_OK;
	LPSPropValue folderType = NULL;
	tstring restString;

	*pLog << "Checking folder type of folder: " << folderPath.c_str() << " ... ";

	hr = HrGetOneProp(folder, PR_FOLDER_TYPE, &folderType);
	if (hr != S_OK)
	{
		*pLog << _T("Couldn't get PR_FOLDER_TYPE. Skipping this folder.\n");
		goto Cleanup;
	}

	if (folderType->Value.l != FOLDER_SEARCH)
	{
		*pLog << _T("Not a search folder.\n");
		goto Cleanup;
	}

	*pLog << _T("This is a search folder. Exporting...\n");

	LPSRestriction lpRes = NULL;
	LPENTRYLIST lpEntryList = NULL;
	ULONG ulSearchState = 0;

	CORg(folder->GetSearchCriteria(
		fMapiUnicode,
		&lpRes,
		&lpEntryList,
		&ulSearchState));

	if (MAPI_E_NOT_INITIALIZED == hr)
	{
		*exportFile << folderPath.c_str() << _T(",");
		hr = S_OK;
	}

	restString = RestrictionToString(lpRes, folder);

	*exportFile << folderPath.c_str() << _T(",");
	*exportFile << restString.c_str();
	*exportFile << "\n";

Cleanup:
	return;

Error:
	goto Cleanup;
}

tstring ExportSearchFoldersOperation::GetPropValueString(SPropValue thisProp)
{
	std::wstringstream wss;
	if (PROP_TYPE(thisProp.ulPropTag) == PT_UNICODE)
	{
		wss << thisProp.Value.lpszW;
	}
	else if (PROP_TYPE(thisProp.ulPropTag) == PT_LONG)
	{
		wss << thisProp.Value.l;
	}
	else if (PROP_TYPE(thisProp.ulPropTag) == PT_BOOLEAN)
	{
		wss << thisProp.Value.b;
	}
	else if (PROP_TYPE(thisProp.ulPropTag) == PT_BINARY)
	{
		for (ULONG i = 0; i < thisProp.Value.bin.cb; i++)
		{
			BYTE oneByte = thisProp.Value.bin.lpb[i];
			if (oneByte < 0x10)
			{
				wss << _T("0");
			}

			wss << std::hex << oneByte;
		}
	}
	else
	{
		wss << "MAPIFolders doesn't support this prop type yet";
	}
	
	tstring stringToReturn = _T("");
	stringToReturn.append(wss.str());
	return stringToReturn;
}

//Based on MFCMapi, but heavily modified
void ExportSearchFoldersOperation::RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj, _In_ tstring *PropString)
{
	if (!PropString) return;
	*PropString = _T("("); // STRING_OK

	std::wstringstream oss;

	ULONG i = 0;
	if (!lpRes)
	{
		PropString->append(_T("NULL)"));
		return;
	}

	tstring szTmp;
	tstring szProp;
	tstring szAltProp;

	LPTSTR szFlags = NULL;
	LPWSTR szPropNum = NULL;

	switch (lpRes->rt)
	{
	case RES_COMPAREPROPS:
		oss << _T("RES_COMPAREPROPS ");
		oss << std::hex << lpRes->res.resCompareProps.ulPropTag1;
		oss << _T(" ");
		oss << std::hex << lpRes->res.resCompareProps.ulPropTag2;
		break;
	case RES_AND:
		oss << _T("RES_AND ");
		if (lpRes->res.resAnd.lpRes)
		{
			for (i = 0; i< lpRes->res.resAnd.cRes; i++)
			{
				RestrictionToString(&lpRes->res.resAnd.lpRes[i], lpObj, &szTmp);
				oss << szTmp.c_str();
			}
		}
		break;
	case RES_OR:
		oss << _T("RES_OR ");
		if (lpRes->res.resOr.lpRes)
		{
			for (i = 0; i < lpRes->res.resOr.cRes; i++)
			{
				RestrictionToString(&lpRes->res.resOr.lpRes[i], lpObj, &szTmp);
				oss << szTmp.c_str();
			}
		}
		break;
	case RES_NOT:
		oss << _T("RES_NOT ");
		RestrictionToString(lpRes->res.resNot.lpRes, lpObj, &szTmp);
		oss << szTmp.c_str();
		break;
	case RES_COUNT:
		oss << _T("RES_COUNT ");
		RestrictionToString(lpRes->res.resNot.lpRes, lpObj, &szTmp);
		oss << szTmp.c_str();
		break;
	case RES_CONTENT:
		oss << _T("RES_CONTENT ");
		oss << lpRes->res.resContent.ulFuzzyLevel << _T(" ");
		oss << std::hex << lpRes->res.resContent.ulPropTag;
		oss << " ";
		if (lpRes->res.resContent.lpProp)
		{
			tstring propValueString = GetPropValueString(*lpRes->res.resContent.lpProp);
			oss << propValueString.c_str();
		}
		break;
	case RES_PROPERTY:
		oss << _T("RES_PROPERTY ");
		oss << lpRes->res.resProperty.relop << " ";
		oss << std::hex << lpRes->res.resProperty.ulPropTag;
		oss << " ";
		if (lpRes->res.resProperty.lpProp)
		{
			if (lpRes->res.resProperty.lpProp)
			{
				tstring propValueString = GetPropValueString(*lpRes->res.resProperty.lpProp);
				oss << propValueString.c_str();
			}
		}
		break;
	case RES_BITMASK:
		oss << _T("RES_BITMASK ");
		oss << std::hex << lpRes->res.resBitMask.relBMR;
		oss << " ";
		oss << std::hex << lpRes->res.resBitMask.ulMask;
		oss << " ";
		oss << std::hex << lpRes->res.resBitMask.ulPropTag;
		break;
	case RES_SIZE:
		oss << _T("RES_SIZE ");
		oss << std::hex << lpRes->res.resSize.relop;
		oss << " ";
		oss << std::hex << lpRes->res.resSize.cb;
		oss << " ";
		oss << std::hex << lpRes->res.resSize.ulPropTag;
		break;
	case RES_EXIST:
		oss << _T("RES_EXIST ");
		oss << std::hex << lpRes->res.resExist.ulPropTag;
		oss << " ";
		oss << std::hex << lpRes->res.resExist.ulReserved1;
		oss << " ";
		oss << std::hex << lpRes->res.resExist.ulReserved2;
		break;
	case RES_SUBRESTRICTION:
		oss << _T("RES_SUBRESTRICTION ");
		oss << std::hex << lpRes->res.resSub.ulSubObject;
		break;
	case RES_COMMENT:
		oss << _T("RES_COMMENT ");
		oss << std::hex << lpRes->res.resComment.cValues;
		oss << " ";
		if (lpRes->res.resComment.lpProp)
		{
			for (i = 0; i< lpRes->res.resComment.cValues; i++)
			{
				oss << std::hex << lpRes->res.resComment.lpProp[i].ulPropTag;
			}

			oss << " ";
		}

		RestrictionToString(lpRes->res.resComment.lpRes, lpObj, &szTmp);
		oss << szTmp.c_str();
		break;
	case RES_ANNOTATION:
		oss << _T("RES_ANNOTATION ");
		oss << std::hex << lpRes->res.resAnnotation.cValues;
		oss << " ";
		if (lpRes->res.resComment.lpProp)
		{
			for (i = 0; i< lpRes->res.resComment.cValues; i++)
			{
				oss << std::hex << lpRes->res.resAnnotation.lpProp[i].ulPropTag;
				oss << " ";
			}
		}

		RestrictionToString(lpRes->res.resAnnotation.lpRes, lpObj, &szTmp);
		*PropString += szTmp;
		break;
	}

	oss << ")";
	PropString->append(oss.str());
	return;
} // RestrictionToString

_Check_return_ tstring ExportSearchFoldersOperation::RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj)
{
	tstring szRes;
	RestrictionToString(lpRes, lpObj, &szRes);
	return szRes;
} // RestrictionToString
