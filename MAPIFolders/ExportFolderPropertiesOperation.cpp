#include "ExportFolderPropertiesOperation.h"
#include "MAPIFolders.h"

struct ID
{
	WORD replid;
	BYTE globcnt[6];
};

ExportFolderPropertiesOperation::ExportFolderPropertiesOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, tstring *pstrPropTags, Log *exportFile, bool useAdmin)
	:OperationBase(pstrBasePath, pstrMailbox, nScope, log, useAdmin)
{
	this->exportFile = exportFile;
	this->pstrPropTags = pstrPropTags;
}


ExportFolderPropertiesOperation::~ExportFolderPropertiesOperation()
{
}

HRESULT ExportFolderPropertiesOperation::Initialize(void)
{
	HRESULT hr = S_OK;
	std::vector<tstring> splitPropStrings;

	CORg(OperationBase::Initialize());

	splitPropStrings = Split(this->pstrPropTags->c_str(), (_T(',')));
	lpPropsToExport = NULL;
	CORg(MAPIAllocateBuffer(CbNewSPropTagArray(splitPropStrings.size()),
		(LPVOID*)&lpPropsToExport));
	this->lpPropsToExport->cValues = splitPropStrings.size();
	for (INT x = 0; x < this->lpPropsToExport->cValues; x++)
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
			lpPropsToExport->aulPropTag[x] = ptag;
		}
	}

	*exportFile << _T("Folder Path,");
	for (INT x = 0; x < this->lpPropsToExport->cValues; x++)
	{
		*exportFile << std::hex << this->lpPropsToExport->aulPropTag[x];
		if (x + 1 < this->lpPropsToExport->cValues)
		{
			*exportFile << _T(",");
		}
	}

	*exportFile << "\n";

Error:
	return hr;
}

void ExportFolderPropertiesOperation::ProcessFolder(LPMAPIFOLDER folder, tstring folderPath)
{
	HRESULT hr = S_OK;

	*pLog << "Exporting properties for folder: " << folderPath.c_str() << "\n";

	ULONG cCount = 0;
	SPropValue *rgprops = NULL;
	CORg(folder->GetProps(lpPropsToExport, MAPI_UNICODE, &cCount, &rgprops));
	*exportFile << folderPath.c_str() << _T(",");
	for (ULONG y = 0; y < cCount; y++)
	{
		SPropValue thisProp = rgprops[y];
		if (PROP_TYPE(thisProp.ulPropTag) == PT_STRING8)
		{
			*exportFile << thisProp.Value.lpszA;
		}
		else if (PROP_TYPE(thisProp.ulPropTag) == PT_UNICODE)
		{
			*exportFile << thisProp.Value.lpszW;
		}
		else if (PROP_TYPE(thisProp.ulPropTag) == PT_I8)
		{
			if (thisProp.ulPropTag == 0x67480014)
			{
				ID* pid = (ID*)&thisProp.Value.li.QuadPart;
				ULONG ul = 0;
				for (int i = 0; i < 6; ++i)
				{
					ul <<= 8;
					ul += pid->globcnt[i];
				}

				*exportFile << std::hex << pid->replid << _T("-") << ul;

			}
			else
			{
				*exportFile << thisProp.Value.li.QuadPart;
			}
		}
		else if (PROP_TYPE(thisProp.ulPropTag) == PT_LONG)
		{
			*exportFile << thisProp.Value.l;
		}
		else if (PROP_TYPE(thisProp.ulPropTag) == PT_BOOLEAN)
		{
			if (thisProp.Value.b)
			{
				*exportFile << "True";
			}
			else
			{
				*exportFile << "False";
			}
		}
		else if (PROP_TYPE(thisProp.ulPropTag) == PT_BINARY)
		{
			for (ULONG i = 0; i < thisProp.Value.bin.cb; i++)
			{
				BYTE oneByte = thisProp.Value.bin.lpb[i];
				if (oneByte < 0x10)
				{
					*exportFile << _T("0");
				}

				*exportFile << std::hex << oneByte;
			}
		}
		else
		{
			*exportFile << _T("MAPIFolders cannot export this property type");
		}

		if (y + 1 < cCount)
		{
			*exportFile << _T(",");
		}
	}

	*exportFile << "\n";

	MAPIFreeBuffer(rgprops);

Cleanup:
	return;

Error:
	goto Cleanup;
}