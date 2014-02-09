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
			Log *pLog = new Log(NULL);
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

			if(ua.fCheckFolderAcl())
			{
				ValidateFolderACL *checkACLOp = new ValidateFolderACL(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), false, pLog);
				CORg(checkACLOp->Initialize());
				OperationBase *op = checkACLOp;
				op->DoOperation();
			}

			if (ua.fFixFolderAcl())
			{
				ValidateFolderACL *checkACLOp = new ValidateFolderACL(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), true, pLog);
				CORg(checkACLOp->Initialize());
				OperationBase *op = checkACLOp;
				op->DoOperation();
			}

			if (ua.fAddFolderPermission())
			{
				ModifyFolderPermissions *modifyPermsOp = new ModifyFolderPermissions(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), false, ua.pstrUser(), ua.pstrRights(), pLog);
				CORg(modifyPermsOp->Initialize());
				OperationBase *op = modifyPermsOp;
				op->DoOperation();
			}

			if (ua.fRemoveFolderPermission())
			{
				ModifyFolderPermissions *modifyPermsOp = new ModifyFolderPermissions(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), true, ua.pstrUser(), ua.pstrRights(), pLog);
				CORg(modifyPermsOp->Initialize());
				OperationBase *op = modifyPermsOp;
				op->DoOperation();
			}

			if(ua.fCheckItems() || ua.fFixItems())
			{
				std::cerr << "unimplemented feature." << "\n";
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

