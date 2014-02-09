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

Log::Log(tstring *logFolderPath)
{
	this->pstrLogFolderPath = logFolderPath;
	this->pstrLogFilePath = NULL;
	this->hLogFile = NULL;
	this->pstrTimeString = NULL;
	this->logstream = NULL;
}

HRESULT Log::Initialize(void)
{
	HRESULT hr = S_OK;

	if (!this->pstrLogFolderPath)
	{
		PWSTR pstrDocumentsPath = NULL;
		CORg(SHGetKnownFolderPath(FOLDERID_Documents, 0, NULL, &pstrDocumentsPath));
		pstrLogFolderPath = new tstring(pstrDocumentsPath);
		pstrLogFolderPath->append(_T("\\MAPIFolders"));
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

	// Genereate a log file name like "20140209114527.log"
	// We also save this time string and make it public so 
	// operations that generate their own specific logs can use it.
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

	this->pstrLogFilePath = new tstring(pstrLogFolderPath->c_str());
	this->pstrLogFilePath->append(_T("\\MAPIFoldersLog"));
	this->pstrLogFilePath->append(this->pstrTimeString->c_str());
	this->pstrLogFilePath->append(_T(".txt"));

	hLogFile = CreateFile(this->pstrLogFilePath->c_str(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hLogFile == INVALID_HANDLE_VALUE)
	{
		DWORD lastError = GetLastError();
		if (lastError == ERROR_FILE_EXISTS)
		{
			// Append numbers until we find one that doesn't already exist or we get a different error
			// Try 100 times
			bool succeeded = false;
			for (UINT x = 1; x <= 100; x++)
			{
				_itow_s(x, buffer, r);
				tstring *newFileName = new tstring(pstrLogFolderPath->c_str());
				newFileName->append(_T("\\MAPIFoldersLog"));
				newFileName->append(this->pstrTimeString->c_str());
				newFileName->append(buffer);
				newFileName->append(_T(".txt"));

				hLogFile = CreateFile(newFileName->c_str(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
				if (hLogFile != INVALID_HANDLE_VALUE)
				{
					delete this->pstrLogFilePath;
					this->pstrLogFilePath = NULL;
					this->pstrLogFilePath = newFileName;
					succeeded = true;
					break;
				}
				else
				{
					lastError = GetLastError();
					if (lastError != ERROR_FILE_EXISTS)
					{
						tcout << "Error creating log file: " << lastError;
						hr = HRESULT_FROM_WIN32(lastError);
						goto Error;
					}
				}

				// Didn't work; clean up and try again
				delete newFileName;
				newFileName = NULL;
			}
		}
		else
		{
			tcout << "Error creating log file: " << lastError;
			hr = HRESULT_FROM_WIN32(lastError);
			goto Error;
		}
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

Log::~Log(void)
{
	if (logstream)
	{
		logstream->flush();
		logstream->close();
	}
}
