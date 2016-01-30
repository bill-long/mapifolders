#pragma once
#include "ItemOperationBase.h"

class ResolveConflictsOperation :
	public ItemOperationBase
{
public:
	ResolveConflictsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log);
	~ResolveConflictsOperation();
	HRESULT Initialize() override;
	void ProcessItem(LPMAPIPROP lpMessage, LPCTSTR itemSubject);

private:
	LPSPropTagArray lpMsgProps;
	LPSPropTagArray lpAttachProps;
};

