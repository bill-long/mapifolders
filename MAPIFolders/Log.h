#pragma once
#include "stdafx.h"
#include <shlobj.h>
#include <KnownFolders.h>
#include <ctime>
#include <iostream>
#include <fstream>

#if defined(UNICODE) || defined(_UNICODE)
#define tcout std::wcout
#define _tstricmp _wcsicmp
#else
#define tcout std::cout
#define _tstricmp _stricmp
#endif

class Log
{
public:
	Log(tstring *logFolderPath, bool consoleOutput);
	~Log(void);
	HRESULT Initialize(void);
	HRESULT Initialize(tstring *filename);
	tstring *pstrLogFolderPath;
	tstring *pstrLogFilePath;
	tstring *pstrTimeString;

	template <class T>
	Log& operator<<(const T& out)
	{
		if (consoleOutput)
		{
			tcout << out;
			tcout.flush();
		}
		*logstream << out;
		logstream->flush();
		return *this;
	}

private:
	std::wofstream *logstream;
	HANDLE hLogFile;
	bool consoleOutput;

};

