#if defined(UNICODE) || defined(_UNICODE)
#define tcout std::wcout
#define _tstricmp _wcsicmp
#else
#define tcout std::cout
#define _tstricmp _stricmp
#endif

#define CORg(_hr) \
	do { \
		hr = _hr; \
		if (FAILED(hr)) \
		{ \
			*pLog << "FAILED! hr = " << std::hex << hr << ".  LINE = " << std::dec << __LINE__ << "\n"; \
			*pLog << " >>> " << (wchar_t*) L#_hr <<  "\n"; \
			goto Error; \
		} \
	} while (0)

#include "OperationBase.h"
#include "ValidateFolderACL.h"
#include "ModifyFolderPermissions.h"
#include "ItemPropertiesOperation.h"
#include "CheckItemsOperation.h"
#include "Log.h"

enum {
    ePR_MEMBER_ENTRYID, 
    ePR_MEMBER_RIGHTS,  
    ePR_MEMBER_ID, 
    ePR_MEMBER_NAME, 
    NUM_COLS
};

static const SizedSPropTagArray(NUM_COLS, rgAclTablePropTags) =
{
	NUM_COLS,
	{
		PR_MEMBER_ENTRYID,  // Unique across directory.
		PR_MEMBER_RIGHTS,  
		PR_MEMBER_ID,       // Unique within ACL table. 
		PR_MEMBER_NAME,     // Display name.
	}
};

static const tstring strAnonymous = _T("Anonymous");
static const tstring strDefault = _T("Default");