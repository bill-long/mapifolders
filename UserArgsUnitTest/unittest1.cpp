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
		** 1. Create a  array of wchar_t strings and its count (argc) to mimic the C runtime's parsed argv
		**    A. For example,  wchar_t *argv[] = {L"ArgTest"};
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
			 wchar_t *argv[] = {L"ArgTest"};

			UserArgs ua;
			bool retVal = ua.Parse(1, argv);

			Assert::AreEqual(retVal, false);
		}
		
		TEST_METHOD(Help)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-?"};

			UserArgs ua;
			unsigned long retVal = ua.ParseGetActions(2, argv);
			
			Assert::AreEqual<unsigned long>(retVal, 0x80000000);
		}

		TEST_METHOD(CheckFolderACL)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv), 0x00000001);
		}

		TEST_METHOD(FixFolderACL)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-fixFolderACL"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv), 0x00000002);
		}

		TEST_METHOD(CheckItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-checkItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv), 0x00000004);
		}

		TEST_METHOD(FixItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-fixItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv), 0x00000008);
		}
		
		// Test that a call with an action/switch does not interfere with subsequent switches
		TEST_METHOD(RunTwice)
		{
			 wchar_t *argv1[] = {L"ArgTest", L"-fixItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv1), 0x00000008);
			 wchar_t *argv2[] = {L"ArgTest", L"-fixFolderACL"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(2, argv2), 0x00000002);
		}
		
		// Result should be 0b1111=0xF
		TEST_METHOD(Everything)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL", L"-fixFolderACL", L"-checkItems", L"-fixItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(5, argv), 0xF);
		}

		// Try all parameters, NOT including help, and a folder
		TEST_METHOD(EverythingAndFolder)
		{
			 wchar_t *pwszMyFolder = L"MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL", L"-fixFolderACL", L"-checkItems", L"-fixItems", pwszMyFolder};
			std::wstring *pstrArgVal;
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(6, argv), 0xF);
			pstrArgVal = pua->pstrPublicFolder();
			Assert::AreEqual(pwszMyFolder, pstrArgVal->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelp)
		{
			 wchar_t *pwszMyFolder = L"MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL", L"-fixFolderACL", L"-checkItems", L"-fixItems", L"-?", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpSlashes)
		{
			 wchar_t *pwszMyFolder = L"MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"/checkFolderACL", L"/fixFolderACL", L"/checkItems", L"/fixItems", L"/?", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes1)
		{
			 wchar_t *pwszMyFolder = L"MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"/checkFolderACL", L"-fixFolderACL", L"/checkItems", L"-fixItems", L"/?", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes2)
		{
			 wchar_t *pwszMyFolder = L"MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL", L"/fixFolderACL", L"-checkItems", L"/fixItems", L"-?", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Run the pair CheckFolderACL, CheckItem
		TEST_METHOD(CheckFolderACLCheckItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL", L"-checkItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0x5);
		}
		
		// Run the pair CheckFolderACL, FixFolderACL
		TEST_METHOD(CheckFolderACLFixFolderACL)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL", L"-fixFolderACL"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0x3);
		}
		
		// Run the pair CheckFolderACL, FixItems
		TEST_METHOD(CheckFolderACLFixItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL", L"-fixItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0x9);
		}
		
		// Run the pair FixFolderACL, CheckItem
		TEST_METHOD(FixFolderACLCheckItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-fixFolderACL", L"-checkItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0x6);
		}
		
		// Run the pair FixFolderACL, FixItems
		TEST_METHOD(FixFolderACLFixItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-fixFolderACL", L"-fixItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(3, argv), 0xA);
		}
		
		// Run the triple CheckFolderACL, FixItems, CheckItems
		TEST_METHOD(CheckFolderACLFixItemsCheckItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL", L"-fixItems", L"-checkItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0xd);
		}
		
		// Run the triple FixFolderACL, FixItems, CheckItems
		TEST_METHOD(FixFolderACLFixItemsCheckItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-fixFolderACL", L"-fixItems", L"-checkItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0xe);
		}
		
		// Run the triple FixFolderACL, CheckFolderACL, CheckItems
		TEST_METHOD(FixFolderACLCheckFolderACLCheckItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-fixFolderACL", L"-checkFolderACL", L"-checkItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0x7);
		}
		
		// Run the triple FixFolderACL, CheckFolderACL, FixItems
		TEST_METHOD(FixFolderACLCheckFolderACLFixItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-fixFolderACL", L"-checkFolderACL", L"-fixItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0xb);
		}
		
		//////////////////////////////////////////
		// Mix up args a little

		// Run the triple CheckFolderACL, FixFolderACL, FixItems
		TEST_METHOD(CheckFolderACLFixFolderACLFixItems)
		{
			 wchar_t *argv[] = {L"ArgTest", L"-checkFolderACL", L"-fixFolderACL", L"-fixItems"};
			Assert::AreEqual<unsigned long>((new UserArgs())->ParseGetActions(4, argv), 0xb);
		}
		
		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes3)
		{
			 wchar_t *pwszMyFolder = L"MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"/fixFolderACL", L"-checkItems", L"/fixItems", L"-?", L"-checkFolderACL", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes4)
		{
			 wchar_t *pwszMyFolder = L"MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-checkItems", L"/fixItems", L"-checkFolderACL", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(EverythingFolderAndHelpMixedSlashes4NewPath)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-checkItems", L"/fixItems", L"-checkFolderACL", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual<unsigned long>(pua->ParseGetActions(7, argv), 0x8000000F);
			Assert::AreEqual(pwszMyFolder, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		// Try all parameters, including help, and a folder
		TEST_METHOD(BadPrams1)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACLBAD", L"-checkItems", L"/fixItems", L"-checkFolderACL", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual(pua->Parse(7, argv), false);
			delete pua;
		}

		TEST_METHOD(BadPrams2)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-BADcheckItems", L"/fixItems", L"-checkFolderACL", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual(pua->Parse(7, argv), false);
			delete pua;
		}

		TEST_METHOD(MultipleFolders)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", pwszMyFolder, L"-?", L"/fixFolderACL", L"-checkItems", L"/fixItems", L"-checkFolderACL", pwszMyFolder};
			UserArgs *pua = new UserArgs();
			Assert::AreEqual(pua->Parse(8, argv), false);
			delete pua;
		}

		TEST_METHOD(RunTwiceWithFolders)
		{
			 wchar_t *pwszMyFolder1 = L"\\MyPath\\MyFolder";
			 wchar_t *pwszMyFolder2 = L"\\MyPath\\MyFolder2";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-checkItems", L"/fixItems", L"-checkFolderACL", pwszMyFolder1};
			 wchar_t *argv2[] = {L"ArgTest", L"-?", pwszMyFolder2, L"/fixFolderACL", L"-checkItems", L"/fixItems", L"-checkFolderACL"};
			UserArgs *pua = new UserArgs();
			pua->Parse(7, argv);
			pua->Parse(7,argv2);
			Assert::AreEqual(pwszMyFolder2, pua->pstrPublicFolder()->c_str(), true);
			delete pua;
		}

		TEST_METHOD(ScopeBaseWithColon)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-checkItems", L"-scope:Base", L"/fixItems", L"-checkFolderACL", pwszMyFolder};
			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<unsigned long>(retVal, 0x8000001f);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::BASE));
		}

		TEST_METHOD(ScopeBaseWithSpace)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-checkItems", L"-scope", L"Base", L"/fixItems", L"-checkFolderACL", pwszMyFolder};
			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			unsigned long ulExpectedScope = (unsigned long)(UserArgs::ActionScope::BASE);
			Assert::AreEqual<unsigned long>(retVal, 0x8000001f);
			Assert::AreEqual<unsigned long>(ulReturnedScope, ulExpectedScope);
		}


		TEST_METHOD(ScopeSubtreeWithColon)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-checkItems", L"-scope:Subtree", L"/fixItems", L"-checkFolderACL", pwszMyFolder};
			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<unsigned long>(retVal, 0x8000001f);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::SUBTREE));
		}

		TEST_METHOD(ScopeMissingWithColonBad)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-checkItems", L"-scope:", L"/fixItems", L"-checkFolderACL", pwszMyFolder};
			UserArgs ua; //*pua = new UserArgs();
			bool retVal = ua.Parse(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<bool>(retVal, false);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::NONE));
		}

		TEST_METHOD(ScopeMissingWithColonAtEnd)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-checkItems", L"/fixItems", L"-checkFolderACL", pwszMyFolder, L"-scope:"};
			UserArgs ua; //*pua = new UserArgs();
			bool retVal = ua.Parse(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<bool>(retVal, false);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::NONE));
		}

		TEST_METHOD(ScopeWithColonFixFolderACL)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"/fixFolderACL", L"-scope:OneLevel"};
			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(3, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<unsigned long>(retVal, 0x12); // Verify actions
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::ONELEVEL));
		}

		TEST_METHOD(ScopeWithSpaceFixFolderACL)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"/fixFolderACL", L"-scope", L"OneLevel"};

			UserArgs ua; //*pua = new UserArgs();
			unsigned long retVal = ua.ParseGetActions(4, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			
			Assert::AreEqual<unsigned long>(retVal, 0x12); // Verify actions
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::ONELEVEL));
		}

		TEST_METHOD(ScopeBadWithColonFixFolderACL)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"/fixFolderACL", L"-scope:OneLevelBad"};

			UserArgs ua;
			bool retVal = ua.Parse(3, argv);

			Assert::AreEqual<bool>(retVal, false); // Verify actions
		}

		TEST_METHOD(ScopeBadWithSpaceFixFolderACL)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"/fixFolderACL", L"-scope", L"NoneBad"};

			UserArgs ua; //*pua = new UserArgs();
			bool retVal = ua.Parse(4, argv);
			
			Assert::AreEqual<bool>(retVal, false); // Verify actions
		}

		TEST_METHOD(ScopeMissingWithSpaceAtEnd)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-?", L"/fixFolderACL", L"-checkItems", L"/fixItems", L"-checkFolderACL", pwszMyFolder, L"-scope"};
			UserArgs ua; //*pua = new UserArgs();
			bool retVal = ua.Parse(8, argv);
			unsigned long ulReturnedScope = (unsigned long)(ua.nScope()); 
			Assert::AreEqual<bool>(retVal, false);
			Assert::AreEqual<unsigned long>(ulReturnedScope, (unsigned long)(UserArgs::ActionScope::NONE));
		}

		TEST_METHOD(ScopeMissingWithSpaceAtEndNoOtherArgs)
		{
			 wchar_t *pwszMyFolder = L"\\MyPath\\MyFolder";
			 wchar_t *argv[] = {L"ArgTest", L"-scope"};
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