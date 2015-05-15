#pragma once
#include "OperationBase.h"
class ExportFolderPermissionsOperation :
	public OperationBase
{
public:
	ExportFolderPermissionsOperation(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, Log *log, Log *exportFile);
	~ExportFolderPermissionsOperation();
	HRESULT Initialize(void) override;
	void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);
private:
	Log *exportFile;
};

