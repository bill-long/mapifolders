#pragma once
#include "OperationBase.h"
class ExportFolderPropertiesOperation :
	public OperationBase
{
public:
	ExportFolderPropertiesOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, tstring *pstrPropTags, Log *exportFile, bool useAdmin);
	~ExportFolderPropertiesOperation();
	HRESULT Initialize(void) override;
	void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);
private:
	Log *exportFile;
	tstring *pstrPropTags;
	LPSPropTagArray lpPropsToExport;
};

