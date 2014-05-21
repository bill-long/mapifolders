#pragma once
#include "ItemOperationBase.h"

class RemoveItemPropertiesOperation :
	public ItemOperationBase
{
public:
	RemoveItemPropertiesOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, tstring *pstrPropTags);
	~RemoveItemPropertiesOperation();
	HRESULT Initialize();
	void ProcessItem(LPMAPIPROP lpMessage, LPCTSTR itemSubject);

private:
	tstring *pstrPropTags;
	LPSPropTagArray lpPropsToGet;
};

