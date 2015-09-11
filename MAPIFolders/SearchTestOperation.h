#pragma once
#include "OperationBase.h"
class SearchTestOperation :
	public OperationBase
{
public:
	SearchTestOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log);
	~SearchTestOperation();
	HRESULT Initialize(void) override;
	void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);
private:
	tstring SearchTestOperation::GetPropValueString(SPropValue thisProp);
	void SearchTestOperation::RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj, _In_ tstring *PropString);
	_Check_return_ tstring SearchTestOperation::RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj);
	Log *exportFile;
};
