#pragma once
#include "operationbase.h"

typedef
enum ACLRIGHTS
{ 
	 RIGHTS_EDIT_OWN = 0x8,
	 RIGHTS_EDIT_ALL = 0x20,
	 RIGHTS_DELETE_OWN = 0x10,
	 RIGHTS_DELETE_ALL = 0x40,
	 RIGHTS_READ_ITEMS = 0x1,
	 RIGHTS_CREATE_ITEMS = 0x2,
	 RIGHTS_CREATE_SUBFOLDERS = 0x80,
	 RIGHTS_FOLDER_OWNER = 0x100,
	 RIGHTS_FOLDER_CONTACT = 0x200,
	 RIGHTS_FOLDER_VISIBLE = 0x400,
	 RIGHTS_NONE = 0,
	 ROLE_OWNER = 0x5e3,
	 ROLE_PUBLISH_EDITOR = 0x4e3,
	 ROLE_EDITOR = 0x463,
	 ROLE_PUBLISH_AUTHOR = 0x49b,
	 ROLE_AUTHOR = 0x41b,
	 ROLE_NONEDITING_AUTHOR = 0x413,
	 ROLE_REVIEWER = 0x401,
	 ROLE_CONTRIBUTOR = 0x402,
	 ROLE_NONE = 0x400
} ACLRIGHTS;

class ModifyFolderPermissions :
	public OperationBase
{
public:
	ModifyFolderPermissions(tstring *pstrBasePath, tstring *pstrMailbox, UserArgs::ActionScope nScope, bool remove, tstring *pstrUserString, tstring *pstrRightsString, Log *log);
	HRESULT Initialize(void) override;
	void ProcessFolder(LPMAPIFOLDER folder, tstring folderPath);
	~ModifyFolderPermissions(void);

private:
	bool remove;
	tstring *pstrUserString;
	tstring *pstrRightsString;
	LPSBinary resolvedUserEID;
	LONG rights;
};

