#include "stdafx.h"
#include "CppUnitTest.h"
#include "UserArgs.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;
/*

*/
namespace UnitTests
{		
	TEST_CLASS(UnitTest1)
	{
	public:
		/*
		** HOW TO MAKE A TEST METHOD
		** 1. Create a  array of TCHAR strings and its count (argc) to mimic the C runtime's parsed argv
		**    A. For example,  TCHAR *argv[] = {_T("ArgTest")};
		**    B. As in C/C++, argv[0] is the application name
		** 2. Pass your argc and the argv to an instance of UserArgs' Parse() or ParseGetActions()
		**    A. Parse() just returns success
		**    B. ParseGetActions() returns the _actions code or throws an exception if parsing fails
		**       i. _actions is the internal code to flag the switches selected
		**       ii. CHECKFOLDERACLS=0b0001; FIXFOLDERACLS=0b0010; CHECKITEMS=0b0100; FIXITEMS=0b1000; SCOPE=0b00010000; DISPLAYHELP=0x80000000;
		** 3. Use Assert::xxx() to check the method results of UserArgs
		**    A. You may want to check the Parse() or ParseGetActions() results in 2
		**       i. Parse() just returns whether parsing failed
		**       ii. ParseGetActions() returns which switches were parsed
		**    B. You may want to check the method results, e.g. pstrPublicFolder()
		**       i. After Parse() or ParseGetActions(), you can check the result of the folder passed in
		**       ii. Currently the public folder is the only non-switch argument passed to this utility
		*/
		TEST_METHOD(Nothing)
		{
			 TCHAR *argv[] = {_T("ArgTest")};

			UserArgs ua;
			bool retVal = ua.Parse(1, argv);

			Assert::AreEqual(retVal, false);
		}
		
		TEST_METHOD(Help)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?")};

			UserArgs ua;
			unsigned long retVal = ua.ParseGetActions(2, argv);
			
			Assert::AreEqual<unsigned long>(retVal, 0x80000000);
		}

		TEST_METHOD(CheckFolderACL)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv), 0x00000001);
		}

		TEST_METHOD(FixFolderACL)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-fixFolderACL")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv), 0x00000002);
		}

		TEST_METHOD(CheckItems)
		{
			 TCHAR *argv[] = {_T("ArgTest)"), _T("-checkItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv), 0x00000004);
		}

		TEST_METHOD(FixItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-fixItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv), 0x00000008);
		}
		
		// Test that a call with an action/switch does not interfere with subsequent switches
		TEST_METHOD(RunTwice)
		{
			 TCHAR *argv1[] = {_T("ArgTest"), _T("-fixItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv1), 0x00000008);
			 TCHAR *argv2[] = {_T("ArgTest"), _T("-fixFolderACL")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv2), 0x00000002);
		}
		
		// Result should be 0b1111=0xF
		TEST_METHOD(Everything)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL"), _T("-fixFolderACL"), _T("-checkItems"), _T("-fixItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(5, argv), 0xF);
		}

		// Try all parameters, NOT including help, and a folder
		TEST_METHOD(EverythingAndFolder)
		{
			 TCHAR *pwszMyFolder = _T("MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL"), _T("-fixFolderACL"), _T("-checkItems"), _T("-fixItems"), pwszMyFolder};
			tstring *pstrArgVal;
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(6, argv), 0xF);
			pstrArgVal = pua->pstrPublicFolder();
			Assert::AreEqual(pwszMyFolder, pstrArgVal->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelp)
		{
			 TCHAR *pwszMyFolder = _T("MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL"), _T("-fixFolderACL"), _T("-checkItems"), _T("-fixItems"), _T("-?"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpSlashes)
		{
			 TCHAR *pwszMyFolder = _T("MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("/checkFolderACL"), _T("/fixFolderACL"), _T("/checkItems"), _T("/fixItems"), _T("/?"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes1)
		{
			 TCHAR *pwszMyFolder = _T("MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("/checkFolderACL"), _T("-fixFolderACL"), _T("/checkItems"), _T("-fixItems"), _T("/?"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes2)
		{
			 TCHAR *pwszMyFolder = _T("MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL"), _T("/fixFolderACL"), _T("-checkItems"), _T("/fixItems"), _T("-?"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Run the pair CheckFolderACL, CheckItem
		TEST_METHOD(CheckFolderACLCheckItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL"), _T("-checkItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0x5);
		}
		
		// Run the pair CheckFolderACL, FixFolderACL
		TEST_METHOD(CheckFolderACLFixFolderACL)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL"), _T("-fixFolderACL")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0x3);
		}
		
		// Run the pair CheckFolderACL, FixItems
		TEST_METHOD(CheckFolderACLFixItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL"), _T("-fixItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0x9);
		}
		
		// Run the pair FixFolderACL, CheckItem
		TEST_METHOD(FixFolderACLCheckItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-fixFolderACL"), _T("-checkItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0x6);
		}
		
		// Run the pair FixFolderACL, FixItems
		TEST_METHOD(FixFolderACLFixItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-fixFolderACL"), _T("-fixItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0xA);
		}
		
		// Run the triple CheckFolderACL, FixItems, CheckItems
		TEST_METHOD(CheckFolderACLFixItemsCheckItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL"), _T("-fixItems"), _T("-checkItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0xd);
		}
		
		// Run the triple FixFolderACL, FixItems, CheckItems
		TEST_METHOD(FixFolderACLFixItemsCheckItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-fixFolderACL"), _T("-fixItems"), _T("-checkItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0xe);
		}
		
		// Run the triple FixFolderACL, CheckFolderACL, CheckItems
		TEST_METHOD(FixFolderACLCheckFolderACLCheckItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-fixFolderACL"), _T("-checkFolderACL"), _T("-checkItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0x7);
		}
		
		// Run the triple FixFolderACL, CheckFolderACL, FixItems
		TEST_METHOD(FixFolderACLCheckFolderACLFixItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-fixFolderACL"), _T("-checkFolderACL"), _T("-fixItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0xb);
		}
		
		//////////////////////////////////////////
		// Mix up args a little

		// Run the triple CheckFolderACL, FixFolderACL, FixItems
		TEST_METHOD(CheckFolderACLFixFolderACLFixItems)
		{
			 TCHAR *argv[] = {_T("ArgTest"), _T("-checkFolderACL"), _T("-fixFolderACL"), _T("-fixItems")};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0xb);
		}
		
		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes3)
		{
			 TCHAR *pwszMyFolder = _T("MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("/fixFolderACL"), _T("-checkItems"), _T("/fixItems"), _T("-?"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes4)
		{
			 TCHAR *pwszMyFolder = _T("MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes4NewPath)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(BadPrams1)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACLBAD"), _T("-checkItems"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual(pua->Parse(7, argv), false);
			delete pua;
		}

		TEST_METHOD(BadPrams2)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-BADcheckItems"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual(pua->Parse(7, argv), false);
			delete pua;
		}

		TEST_METHOD(MultipleFolders)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), pwszMyFolder, _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual(pua->Parse(8, argv), false);
			delete pua;
		}

		TEST_METHOD(RunTwiceWithFolders)
		{
			 TCHAR *pwszMyFolder1 = _T("\\MyPath\\MyFolder");
			 TCHAR *pwszMyFolder2 = _T("\\MyPath\\MyFolder2");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder1};
			 TCHAR *argv2[] = {_T("ArgTest"), _T("-?"), pwszMyFolder2, _T("/fixFolderACL"), _T("-checkItems"), _T("/fixItems"), _T("-checkFolderACL")};
			UserArgs *pua = new UserArgs();
			pua->Parse(7, argv);
			pua->Parse(7,argv2);
			Assert::AreEqual(pwszMyFolder2, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		TEST_METHOD(ScopeBaseWithColon)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("-scope:Base"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<unsigned long>(retVal, 0x8000001f);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::BASE));
		}

		TEST_METHOD(ScopeBaseWithSpace)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("-scope"), _T("Base"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			unsigned long ulExpectedScope = (unsigned long)(UserArgs::ActionScope::BASE);
			Assert::AreEqual<unsigned long>(retVal, 0x8000001f);
			Assert::AreEqual<unsigned long>(ulReturnedScope, ulExpectedScope);
		}


		TEST_METHOD(ScopeSubtreeWithColon)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("-scope:Subtree"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<unsigned long>(retVal, 0x8000001f);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::SUBTREE));
		}

		TEST_METHOD(ScopeMissingWithColonBad)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("-scope:"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder};
			UserArgs ua; //*pua = new UserArgs();
			bool retVal = ua.Parse(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<bool>(retVal, false);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::NONE));
		}

		TEST_METHOD(ScopeMissingWithColonAtEnd)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder, _T("-scope:")};
			UserArgs ua; //*pua = new UserArgs();
			bool retVal = ua.Parse(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<bool>(retVal, false);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::NONE));
		}

		TEST_METHOD(ScopeWithColonFixFolderACL)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("/fixFolderACL"), _T("-scope:OneLevel")};
			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(3, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<unsigned long>(retVal, 0x12); // Verify actions
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::ONELEVEL));
		}

		TEST_METHOD(ScopeWithSpaceFixFolderACL)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("/fixFolderACL"), _T("-scope"), _T("OneLevel")};

			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(4, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			
			Assert::AreEqual<unsigned long>(retVal, 0x12); // Verify actions
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::ONELEVEL));
		}

		TEST_METHOD(ScopeBadWithColonFixFolderACL)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("/fixFolderACL"), _T("-scope:OneLevelBad")};

			UserArgs ua;
			bool retVal = ua.Parse(3, argv);

			Assert::AreEqual<bool>(retVal, false); // Verify actions
		}

		TEST_METHOD(ScopeBadWithSpaceFixFolderACL)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("/fixFolderACL"), _T("-scope"), _T("NoneBad")};

			UserArgs ua; //*pua = new UserArgs();
			bool retVal = ua.Parse(4, argv);
			
			Assert::AreEqual<bool>(retVal, false); // Verify actions
		}

		TEST_METHOD(ScopeMissingWithSpaceAtEnd)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-?"), _T("/fixFolderACL"), _T("-checkItems"), _T("/fixItems"), _T("-checkFolderACL"), pwszMyFolder, _T("-scope")};
			UserArgs ua; //*pua = new UserArgs();
			bool retVal = ua.Parse(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<bool>(retVal, false);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::NONE));
		}

		TEST_METHOD(ScopeMissingWithSpaceAtEndNoOtherArgs)
		{
			 TCHAR *pwszMyFolder = _T("\\MyPath\\MyFolder");
			 TCHAR *argv[] = {_T("ArgTest"), _T("-scope")};
			UserArgs ua; //*pua = new UserArgs();
			bool retVal = ua.Parse(2, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<bool>(retVal, false);
		}

		//TEST_METHOD(UnparsedAccess)
		//{
		//	UserArgs ua;
		//	Assert::ExpectException<>(ua.nScope()); 
		//}



	};
}