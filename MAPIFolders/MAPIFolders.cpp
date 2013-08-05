// FixPFItems.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "MAPIFolders.h"

#pragma warning( disable : 4127)

int _tmain(int argc, _TCHAR* argv[])
{
	UserArgs ua;
	ua.Parse(argc, argv);
	ValidateFolderACL *checkACLOp = new ValidateFolderACL(false);
	OperationBase *op = checkACLOp;
	op->DoOperation();
}

