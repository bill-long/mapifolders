#pragma once
#include "OperationBase.h"
class ExportSearchFoldersOperation :
	public OperationBase
{
public:
	ExportSearchFoldersOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, Log *exportFile, bool useAdmin);
	~ExportSearchFoldersOperation();
	HRESULT Initialize(void) override;
	void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);
private:
	tstring ExportSearchFoldersOperation::GetPropValueString(SPropValue thisProp);
	void ExportSearchFoldersOperation::RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj, _In_ tstring *PropString);
	_Check_return_ tstring ExportSearchFoldersOperation::RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj);
	Log *exportFile;
};

