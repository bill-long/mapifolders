#include "stdafx.h"
#include "UserArgs.h"


// Initialize static const's
//The array of switch structures
//Update this if new actions are added as well as #define _COUNTOFACCEPTEDSWITCHES
const UserArgs::ArgSwitch UserArgs::rgArgSwitches[_COUNTOFACCEPTEDSWITCHES] = 
{
	{L"CheckFolderACL", L"Check Folder ACL", false, CHECKFOLDERACLS},
	{L"FixFolderACL", L"Fix Folder ACL", false, FIXFOLDERACLS},
	{L"CheckItems", L"Check Items", false, CHECKITEMS},
	{L"FixItems", L"Fix Items", false, FIXITEMS},
	{L"Scope", L"Action Scope", true, SCOPE},
	{L"?", L"Help", false, DISPLAYHELP},
};

// The array of scopes
const UserArgs::ScopeValue UserArgs::rgScopeValues[_COUNTOFSCOPES] =
{
	{UserArgs::ActionScope::BASE, L"Base"},
	{UserArgs::ActionScope::ONELEVEL, L"OneLevel"},
	{UserArgs::ActionScope::SUBTREE, L"SubTree"},
};

//The string of characters to accept as switches
const std::wstring UserArgs::_wstrSwitchChars = std::wstring(L"-/");


UserArgs::~UserArgs(void)
{
	if(_pstrPublicFolder)
	{
		delete _pstrPublicFolder;
		_pstrPublicFolder=NULL;
	}
}

UserArgs::UserArgs(void)
{
	// Initialize _pstrPublicFolder as it is a pointer
	 _pstrPublicFolder=NULL;
	// Default _actions
	init();
}

void UserArgs::init(void)
{
	 _actions=0;
	 if(_pstrPublicFolder!=NULL)
		 delete _pstrPublicFolder;
	 _pstrPublicFolder=NULL;
	 _scope = ActionScope::NONE;
}

// Error logging
void UserArgs::logError(int argNum, const wchar_t *strArg, unsigned long err)
{
	wprintf(L"Error parsing argument %d ('%s').\n", argNum, strArg);
	switch(err)
	{
		case ERR_NONE:
			wprintf(L"Success\n");	// Should not log success error
			break;
		case ERR_INVALIDSWITCH:
			wprintf(L"Invalid switch\n");	// Should not log success error
			break;
		case ERR_DUPLICATEFOLDER:
			wprintf(L"Folder already specified\n");	// Should not log success error
			break;
		case ERR_EXPECTEDSWITCHVALUE:
			wprintf(L"Expected a value\n");
			break;
		case ERR_INVALIDSTATE:
			wprintf(L"Unexpected state during parsing\n");
			break;
		default:
			wprintf(L"Unknown error logged\n");
	}
}

bool UserArgs::Parse(int argc, wchar_t* argv[])
{
	init();	// Reset from any previous run
	int iCurrentArg=1, iCurrChar=0; // indexes of current arg and character withing argv[iCurrentArg]
	const wchar_t *pchCurrent = NULL;	// Pointer to the current character in argv[iCurrentArg]

	bool retVal = false;
	
	// If we have at least one arg (aside from arg[0])
	// CONSIDER: Error message or error code describing what was wrong with args?
	if(argc>1)
	{
		// At this point, we have a starting token.
		PARSESTATE state = STATE_NEWTOKEN;
		bool foundSimpleSwitch, foundValueSwitch, foundScope;
		unsigned long ulCurrentValueSwitch=NULL;	// used when a valueswitch is found and we are looking for value
		do
		{
			pchCurrent=&argv[iCurrentArg][iCurrChar];
			switch(state)
			{
			case STATE_NEWTOKEN:
				/* See if current character is an accepted switch character
				(e.g. '-' or '/'); if so, advance a character and set state to 
				SWITCH; else do not advance but set state to VALUE */
				if(std::string::npos != _wstrSwitchChars.find(pchCurrent[0]))
				{
					iCurrChar++;	// could be '\0' here
					state=STATE_SWITCH;
				}
				else
					state=STATE_VALUE;
				break;
			case STATE_SWITCH:
				foundSimpleSwitch=foundValueSwitch=false; // set when we find a switch
				for(int i=0; i<_COUNTOFACCEPTEDSWITCHES; i++)
				{
					// Is this a simple- or a value-switch?
					if(!rgArgSwitches[i].fHasValue)
					{
						// Search for an exact match for non-value (i.e. simple) switches
						if(0==_wcsnicmp(pchCurrent, rgArgSwitches[i].pszSwitch, FILENAME_MAX))	// arbitrary max length
						{
							// matched a switch
							// Add to actions
							_actions |= rgArgSwitches[i].flagAction;
							foundSimpleSwitch=true;
							break;
						} // if (0==_wcsnicmp(pchCurrent, rgArgSwitches[i].pszSwitch, ...
					} // if !rgArgSwitches[i].fHasValue
					else
					{
						// Then this is a value switch
						// Set iCurrChar character after switch's name (will be '\0' or ':' else fail), set ulCurrentSwitch, set state to find value
						size_t nSwitchLength = wcsnlen(rgArgSwitches[i].pszSwitch, FILENAME_MAX);
						size_t nCurrentSubstrLength = wcsnlen(pchCurrent, FILENAME_MAX);
						if(0==_wcsnicmp(pchCurrent, 
							rgArgSwitches[i].pszSwitch, 
							nSwitchLength)) // See if pchCurrent starts with this accepted value switch
						{
							// Found a matching valueswitch
							_actions |= rgArgSwitches[i].flagAction; // Though this is not an *action*, _actions are to verify accepted options
							ulCurrentValueSwitch = rgArgSwitches[i].flagAction;
							foundValueSwitch=true;
							if(pchCurrent[nSwitchLength]==L':')
								iCurrChar+=nSwitchLength+1;	// set to just after ':'
							else if(pchCurrent[nSwitchLength]==L'\0')
							{
								if(++iCurrentArg >= argc)
								{
									logError(--iCurrentArg, argv[iCurrentArg], ERR_EXPECTEDSWITCHVALUE);
									foundSimpleSwitch=foundSimpleSwitch=foundValueSwitch=false;	// We found nothing good
								}
								// else leave iCurrentArg at next and...
								iCurrChar=0;
							}
							break;
						}
					} // if rgArgSwitches[i].fHasValue
				} // for

				// when done, see if we found the switch
				if(foundSimpleSwitch)
					state=STATE_ADVANCEARG;
				else if(foundValueSwitch)
					state=STATE_VALUESWITCH; // cursor is left at end of arg[current]+len(switch)+1 ('\0' or ':') or start of next arg
				else
				{
					logError(iCurrentArg, argv[iCurrentArg], ERR_INVALIDSWITCH);
					state=STATE_FAIL;
				}
				break;
			case STATE_VALUESWITCH:
				// Here we have pchCurrent, iCurrChar, and iCurrentArg all set correctly for start of value for this valueswitch
				// There is only one value switch for this utility, but for expandability, pretend there are many
				switch(ulCurrentValueSwitch)
				{
					case SCOPE:
						foundScope=false;
						for(int i=0; i<_COUNTOFSCOPES && !foundScope; i++)
						{
							if(0==_wcsnicmp(pchCurrent, rgScopeValues[i].pszScope, FILENAME_MAX))
							{
								this->_scope= rgScopeValues[i].nScopeCode;
								foundScope=true;
							}
						}
						break;
					default:
						logError(iCurrentArg, argv[iCurrentArg], ERR_INVALIDSTATE);
						break;
				}
				if(foundScope)
				{
					state=STATE_ADVANCEARG;
				}
				else
				{
					logError(iCurrentArg, argv[iCurrentArg], ERR_EXPECTEDSWITCHVALUE);
					state=STATE_FAIL;
				}
				break;
			case STATE_VALUE:
				// TODO: accept a value, set state to ADVANCEARG
				if(_pstrPublicFolder!=NULL)
				{
					// Not allowing multiple values
					logError(iCurrentArg, argv[iCurrentArg], ERR_DUPLICATEFOLDER);
					state=STATE_FAIL;
				}
				else
				{
					_pstrPublicFolder = new std::wstring(pchCurrent);
					state=STATE_ADVANCEARG;
				}
				break;
			case STATE_ADVANCEARG:
				iCurrChar=0;
				if(++iCurrentArg>=argc)
				{
					state=STATE_DONE;
					retVal = true;	// If we get to the end with no errors, we are successful
				}
				else
					state = STATE_NEWTOKEN;
				break;
			case STATE_DONE:
				break;
			case STATE_FAIL:
				// TODO: Log msg (index of arg, error, etc.) everywhere state is set to STATE_FAIL
				retVal=false;
				state=STATE_DONE;
				break;
			default:
				state=STATE_FAIL;
				break;
			}
		} while(state!=STATE_DONE);
	}
	return retVal;
}

// Parses arguments and returns the _actions code
unsigned long UserArgs::ParseGetActions(int argc, wchar_t *argv[])
{
	bool result = Parse(argc, argv);
	if(result)
		return _actions;
	else
		throw new std::exception; //was return NULL; 
}
