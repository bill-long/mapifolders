// FixPFItems.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "MAPIFolders.h"

#pragma warning( disable : 4127)

int _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hr = S_OK;
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

