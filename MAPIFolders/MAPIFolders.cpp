// FixPFItems.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "fcntl.h"
#include "io.h"
#include "MAPIFolders.h"

#pragma warning( disable : 4127)

int _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hr = S_OK;

#if defined(UNICODE) || defined(_UNICODE)
	_setmode(_fileno(stdout), _O_U16TEXT);
#endif

	UserArgs ua;
	if (ua.Parse(argc, argv))
	{

#ifdef DEBUG
		// Dump parsed args
		tcout << "Parsed successfully.\n";
		tcout << "fCheckFolderAcl: " << ua.fCheckFolderAcl() << "\n";
		tcout << "fCheckItems: " << ua.fCheckItems() << "\n";
		tcout << "fDisplayHelp: " << ua.fDisplayHelp() << "\n";
		tcout << "fFixFolderAcl: " << ua.fFixFolderAcl() << "\n";
		tcout << "fFixItems: " << ua.fFixItems() << "\n";
		tcout << "nScope: " << (short int)ua.nScope() << "\n";
		tcout << "pstrMailbox: " ;
		if (ua.pstrMailbox())
			tcout << ua.pstrMailbox()->c_str() << "\n";
		else
			tcout << "Null" << "\n";
		tcout << "pstrFolderPath: ";
		if (ua.pstrFolderPath())
			tcout << ua.pstrFolderPath()->c_str() << "\n";
		else
			tcout << "Null" << "\n";
#endif

		if(ua.fDisplayHelp())
		{
			ua.ShowHelp();
		}
		else
		{
			Log *pLog = new Log(NULL, true);
			CORg(pLog->Initialize());
			LPTSTR lpTimeStr = NULL;
			int cchTime = GetTimeFormat(LOCALE_SYSTEM_DEFAULT, 0, NULL, NULL, NULL, 0);
			lpTimeStr = new TCHAR[cchTime];
			if (!GetTimeFormat(LOCALE_SYSTEM_DEFAULT, 0, NULL, NULL, lpTimeStr, cchTime))
			{
				DWORD lastError = GetLastError();
				*pLog << "Error getting system time: " << lastError << "\n";
				hr = HRESULT_FROM_WIN32(lastError);
				goto Error;
			}
			*pLog << "MAPIFolders started at " << lpTimeStr << "\n";

			OperationBase *op = NULL;

			if(ua.fCheckFolderAcl())
			{
				op = new ValidateFolderACL(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), false, pLog);
			}

			if (ua.fFixFolderAcl())
			{
				op = new ValidateFolderACL(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), true, pLog);
			}

			if (ua.fAddFolderPermission())
			{
				op = new ModifyFolderPermissions(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), false, ua.pstrUser(), ua.pstrRights(), pLog);
			}

			if (ua.fRemoveFolderPermission())
			{
				op = new ModifyFolderPermissions(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), true, ua.pstrUser(), ua.pstrRights(), pLog);
			}

			if (ua.fRemoveItemProperties())
			{
				op = new RemoveItemPropertiesOperation(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), pLog, ua.pstrProplist());
			}

			if(ua.fCheckItems() || ua.fFixItems())
			{
				op = new CheckItemsOperation(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), pLog, ua.fFixItems());
			}

			if (ua.fResolveConflicts())
			{
				op = new ResolveConflictsOperation(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), pLog);
			}

			if (ua.fExportFolderProperties())
			{
				Log *exportFile = new Log(NULL, false);
				tstring *exportFileName = new tstring(_T("ExportFolderProperties"));
				exportFileName->append(pLog->pstrTimeString->c_str());
				exportFileName->append(_T(".csv"));
				CORg(exportFile->Initialize(exportFileName));
				op = new ExportFolderPropertiesOperation(
					ua.pstrFolderPath(),
					ua.pstrMailbox(),
					ua.nScope(),
					pLog,
					ua.pstrProplist(),
					exportFile);
			}

			if (ua.fExportFolderPermissions())
			{
				Log *exportFile = new Log(NULL, false);
				tstring *exportFileName = new tstring(_T("ExportFolderPermissions"));
				exportFileName->append(pLog->pstrTimeString->c_str());
				CORg(exportFile->Initialize(exportFileName));
				op = new ExportFolderPermissionsOperation(
					ua.pstrFolderPath(),
					ua.pstrMailbox(),
					ua.nScope(),
					pLog,
					exportFile);
			}

			if (ua.fExportSearchFolders())
			{
				Log *exportFile = new Log(NULL, false);
				tstring *exportFileName = new tstring(_T("ExportSearchFolders"));
				exportFileName->append(pLog->pstrTimeString->c_str());
				CORg(exportFile->Initialize(exportFileName));
				op = new ExportSearchFoldersOperation(
					ua.pstrFolderPath(),
					ua.pstrMailbox(),
					ua.nScope(),
					pLog,
					exportFile);
			}

			if (op) {
				CORg(op->Initialize());
				op->DoOperation();
			}

			delete lpTimeStr;
			lpTimeStr = NULL;
			cchTime = GetTimeFormat(LOCALE_SYSTEM_DEFAULT, 0, NULL, NULL, NULL, 0);
			lpTimeStr = new TCHAR[cchTime];
			if (!GetTimeFormat(LOCALE_SYSTEM_DEFAULT, 0, NULL, NULL, lpTimeStr, cchTime))
			{
				// We're done anyway, so who cares
			}
			else
			{
				*pLog << "MAPIFolders finished at " << lpTimeStr << "\n";
			}
		}
	}

	else
	{
		// Parsing error
		std::cerr << "Error parsing command line." << "\n";
		ua.ShowHelp();
	}

Error:
;
}

