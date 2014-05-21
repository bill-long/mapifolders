#include "Log.h"

#define CORg(_hr) \
	do { \
		hr = _hr; \
		if (FAILED(hr)) \
		{ \
			tcout << "FAILED! hr = " << std::hex << hr << ".  LINE = " << std::dec << __LINE__ << "\n"; \
			tcout << " >>> " << (wchar_t*) L#_hr <<  "\n"; \
			goto Error; \
		} \
	} while (0)

Log::Log(tstring *logFolderPath, bool consoleOutput)
{
	this->pstrLogFolderPath = logFolderPath;
	this->consoleOutput = consoleOutput;
	this->pstrLogFilePath = NULL;
	this->hLogFile = NULL;
	this->pstrTimeString = NULL;
	this->logstream = NULL;

	// Generate a time string like "20140209114527"
	this->pstrTimeString = new tstring(_T(""));
	time_t t = time(0);
	struct tm *now = localtime(&t);
	now->tm_year += 1900;
	now->tm_mon += 1;
	wchar_t buffer[8];
	int r = 10;
	_itow_s(now->tm_year, buffer, r);
	pstrTimeString->append(buffer);
	_itow_s(now->tm_mon, buffer, r);
	if (wcslen(buffer) < 2)
		pstrTimeString->append(_T("0"));
	pstrTimeString->append(buffer);
	_itow_s(now->tm_mday, buffer, r);
	if (wcslen(buffer) < 2)
		pstrTimeString->append(_T("0"));
	pstrTimeString->append(buffer);
	_itow_s(now->tm_hour, buffer, r);
	if (wcslen(buffer) < 2)
		pstrTimeString->append(_T("0"));
	pstrTimeString->append(buffer);
	_itow_s(now->tm_min, buffer, r);
	if (wcslen(buffer) < 2)
		pstrTimeString->append(_T("0"));
	pstrTimeString->append(buffer);
	_itow_s(now->tm_sec, buffer, r);
	if (wcslen(buffer) < 2)
		pstrTimeString->append(_T("0"));
	pstrTimeString->append(buffer);
}

HRESULT Log::Initialize(tstring *filename)
{
	HRESULT hr = S_OK;

	if (!this->pstrLogFolderPath)
	{
		PWSTR pstrDocumentsPath = NULL;
		CORg(SHGetKnownFolderPath(FOLDERID_Documents, 0, NULL, &pstrDocumentsPath));
		pstrLogFolderPath = new tstring(pstrDocumentsPath);
		pstrLogFolderPath->append(_T("\\MAPIFolders\\"));
	}

	// The parent path must already exist, or this will fail.
	if (!CreateDirectory(pstrLogFolderPath->c_str(), NULL))
	{
		DWORD lastError = GetLastError();
		if (lastError == ERROR_ALREADY_EXISTS)
		{
			// That's fine
		}
		else
		{
			// Any other error is not.
			tcout << "Error initializing log path: " << lastError << std::endl;
			hr = HRESULT_FROM_WIN32(lastError);
		}
	}

	this->pstrLogFilePath = new tstring(pstrLogFolderPath->c_str());
	this->pstrLogFilePath->append(filename->c_str());

	hLogFile = CreateFile(this->pstrLogFilePath->c_str(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hLogFile == INVALID_HANDLE_VALUE)
	{
		DWORD lastError = GetLastError();
		tcout << "Error creating log file: " << lastError;
		hr = HRESULT_FROM_WIN32(lastError);
		goto Error;
	}

	// If we're here, then we have a handle to a uniquely named log file. Yay!
	// Now get rid of the handle and reopen the file as a standard stream.
	CloseHandle(hLogFile);
	this->logstream = new std::wofstream(this->pstrLogFilePath->c_str());
	if (logstream->fail() || logstream->bad() || !logstream->is_open())
	{
		tcout << "Failed to open log file stream." << std::endl;
		hr = S_FALSE;
	}

Error:
	return hr;
}

HRESULT Log::Initialize(void)
{
	HRESULT hr = S_OK;

	tstring *logFilePath = new tstring(_T("MAPIFoldersLog"));
	logFilePath->append(this->pstrTimeString->c_str());
	logFilePath->append(_T(".txt"));

	CORg(this->Initialize(logFilePath));

Error:
	return hr;
}

Log::~Log(void)
{
	if (logstream)
	{
		logstream->flush();
		logstream->close();
	}
}
