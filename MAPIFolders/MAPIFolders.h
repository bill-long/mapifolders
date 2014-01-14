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
			tcout << "FAILED! hr = " << std::hex << hr << ".  LINE = " << std::dec << __LINE__ << std::endl; \
			tcout << " >>> " << (wchar_t*) L#_hr <<  std::endl; \
			goto Error; \
		} \
	} while (0)

#include "OperationBase.h"
#include "ValidateFolderACL.h"