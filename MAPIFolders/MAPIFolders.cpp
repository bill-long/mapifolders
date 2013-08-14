// FixPFItems.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "MAPIFolders.h"

#pragma warning( disable : 4127)

int _tmain(int argc, _TCHAR* argv[])
{
	UserArgs ua;
	if (ua.Parse(argc, argv))
	{

#ifdef DEBUG
		// Dump parsed args
		std::cout << "Parsed successfully.";
		std::cout << "fCheckFolderAcl: " << ua.fCheckFolderAcl() << std::endl;
		std::cout << "fCheckItems: " << ua.fCheckItems() << std::endl;
		std::cout << "fDisplayHelp: " << ua.fDisplayHelp() << std::endl;
		std::cout << "fFixFolderAcl: " << ua.fFixFolderAcl() << std::endl;
		std::cout << "fFixItems: " << ua.fFixItems() << std::endl;
#endif

		if(ua.fDisplayHelp())
		{
			ua.ShowHelp();
		}
		else
		{
			if(ua.fCheckFolderAcl())
			{
				ValidateFolderACL *checkACLOp = new ValidateFolderACL(false);
				OperationBase *op = checkACLOp;
				op->DoOperation();
			}
			if(ua.fCheckItems() || ua.fFixFolderAcl() || ua.fFixItems())
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

}

