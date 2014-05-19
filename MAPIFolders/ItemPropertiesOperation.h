#pragma once
#include "OperationBase.h"

class ItemPropertiesOperation :
	public OperationBase
{
public:
	ItemPropertiesOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, tstring *pstrPropTags);
	~ItemPropertiesOperation();
	HRESULT Initialize();
	void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);

private:
	tstring *pstrPropTags;
	ULONG *proptags;
	INT cproptags;
};

