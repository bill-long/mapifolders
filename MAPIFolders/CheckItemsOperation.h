#pragma once
#include "ItemOperationBase.h"
class CheckItemsOperation :
	public ItemOperationBase
{
public:
	CheckItemsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, bool fix);
	~CheckItemsOperation();
	HRESULT Initialize() override;
	void ProcessItem(LPMAPIPROP item, LPCTSTR itemSubject);
	bool CheckItemsOperation::ProcessAttachmentsRecursive(LPMAPIPROP lpMessage);

private:
	bool fix;
	LPSPropTagArray lpPropsToRemove;
	LPSPropTagArray lpAttachProps;
	LPSPropTagArray lpAttachTableProps;
};
