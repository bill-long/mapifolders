#pragma once
#include "ItemOperationBase.h"
class CheckItemsOperation :
	public ItemOperationBase
{
public:
	CheckItemsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, bool fix);
	~CheckItemsOperation();
	HRESULT Initialize();
	void ProcessItem(LPMAPIPROP item, LPCTSTR itemSubject);

private:
	bool fix;
	LPSPropTagArray lpPropsToRemove;
};

