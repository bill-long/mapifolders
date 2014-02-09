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
		tcout << "Parsed successfully." << std::endl;
		tcout << "fCheckFolderAcl: " << ua.fCheckFolderAcl() << std::endl;
		tcout << "fCheckItems: " << ua.fCheckItems() << std::endl;
		tcout << "fDisplayHelp: " << ua.fDisplayHelp() << std::endl;
		tcout << "fFixFolderAcl: " << ua.fFixFolderAcl() << std::endl;
		tcout << "fFixItems: " << ua.fFixItems() << std::endl;
		tcout << "nScope: " << (short int)ua.nScope() << std::endl;
		tcout << "pstrMailbox: " ;
		if (ua.pstrMailbox())
			tcout << ua.pstrMailbox()->c_str() << std::endl;
		else
			tcout << "Null" << std::endl;
		tcout << "pstrFolderPath: ";
		if (ua.pstrFolderPath())
			tcout << ua.pstrFolderPath()->c_str() << std::endl;
		else
			tcout << "Null" << std::endl;
#endif

		if(ua.fDisplayHelp())
		{
			ua.ShowHelp();
		}
		else
		{
			if(ua.fCheckFolderAcl())
			{
				ValidateFolderACL *checkACLOp = new ValidateFolderACL(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), false);
				CORg(checkACLOp->Initialize());
				OperationBase *op = checkACLOp;
				op->DoOperation();
			}

			if (ua.fFixFolderAcl())
			{
				ValidateFolderACL *checkACLOp = new ValidateFolderACL(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), true);
				CORg(checkACLOp->Initialize());
				OperationBase *op = checkACLOp;
				op->DoOperation();
			}

			if (ua.fAddFolderPermission())
			{
				ModifyFolderPermissions *modifyPermsOp = new ModifyFolderPermissions(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), false, ua.pstrUser(), ua.pstrRights());
				CORg(modifyPermsOp->Initialize());
				OperationBase *op = modifyPermsOp;
				op->DoOperation();
			}

			if (ua.fRemoveFolderPermission())
			{
				ModifyFolderPermissions *modifyPermsOp = new ModifyFolderPermissions(ua.pstrFolderPath(), ua.pstrMailbox(), ua.nScope(), true, ua.pstrUser(), ua.pstrRights());
				CORg(modifyPermsOp->Initialize());
				OperationBase *op = modifyPermsOp;
				op->DoOperation();
			}

			if(ua.fCheckItems() || ua.fFixItems())
			{
				std::cerr << "unimplemented feature." << std::endl;
			}
		}
	}

	else
	{
		// Parsing error
		std::cerr << "Error parsing command line." << std::endl;
		ua.ShowHelp();
	}

Error:
	tcout << "Finished." << std::endl;
}

