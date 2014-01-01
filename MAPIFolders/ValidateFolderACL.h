#pragma once
#include "operationbase.h"

#define PR_NT_SECURITY_DESCRIPTOR PROP_TAG(PT_BINARY, 0x0E27) // What header file am I missing?

class ValidateFolderACL :
	public OperationBase
{
public:
	ValidateFolderACL(tstring *pstrBasePath, UserArgs::ActionScope nScope, bool fixBadACLs);
	~ValidateFolderACL(void);
	void ProcessFolder(LPMAPIFOLDER folder, std::wstring folderPath);

private:
	bool FixBadACLs;
	void FixACL(LPMAPIFOLDER folder);
};

