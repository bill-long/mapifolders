#pragma once
#include "OperationBase.h"
class CheckItemsOperation :
	public OperationBase
{
public:
	CheckItemsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, bool fix);
	~CheckItemsOperation();
	HRESULT Initialize();
	void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);

private:
	bool fix;
	LPSPropTagArray lpPropsToRemove;
};

