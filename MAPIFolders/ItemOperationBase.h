#pragma once
#include "OperationBase.h"
class ItemOperationBase :
	public OperationBase
{
public:
	ItemOperationBase(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log);
	~ItemOperationBase();
	HRESULT Initialize();
	void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);
	virtual void ProcessItem(LPMAPIPROP item, LPCTSTR itemSubject);
};

