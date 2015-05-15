#pragma once
#include "OperationBase.h"

#define PR_NT_SECURITY_DESCRIPTOR PROP_TAG(PT_BINARY, 0x0E27) // What header file am I missing?

class ValidateFolderACL :
	public OperationBase
{
public:
	ValidateFolderACL(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, bool fixBadACLs, Log *log);
	~ValidateFolderACL(void);
	HRESULT Initialize(void) override;
	void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);

private:
	bool FixBadACLs;
	bool IsInitialized;
	PSID psidAnonymous;
	void FixACL(LPMAPIFOLDER folder);
	HRESULT CheckACLTable(LPMAPIFOLDER folder, bool &aclTableIsGood);
};

