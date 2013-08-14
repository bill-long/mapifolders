#pragma once
#include <string>
#include <iostream>
#include <tchar.h>

#define _COUNTOFACCEPTEDSWITCHES (6)
#define _COUNTOFSCOPES (3)

typedef std::basic_string<TCHAR> tstring;

// Internal structure for accepted switches

class UserArgs
{
public:
	UserArgs();
	~UserArgs();

	// Action accessor methods; use these to determine user selected actions after parsing
	// Update this list to add actions to the application
	bool fCheckFolderAcl() {return (NULL!=(m_actions & CHECKFOLDERACLS));}
	bool fFixFolderAcl() {return (NULL!=(m_actions & FIXFOLDERACLS));}
	bool fCheckItems() {return (NULL!=(m_actions & CHECKITEMS));}
	bool fFixItems() {return (NULL!=(m_actions & FIXITEMS));}
	bool fDisplayHelp() {return (NULL!=(m_actions & DISPLAYHELP));}

	// Argument value (i.e. public folder, the only value argument) accessor methods
	tstring *pstrPublicFolder() {return m_pstrPublicFolder;}

	// Enum for the possible scopes for operations
	enum ActionScope : short int
	{
		NONE=0,	// Default, uninitialized; not a selection for user
		BASE,
		ONELEVEL,
		SUBTREE
	};

	// Access the parsed scope
	ActionScope nScope() {return m_scope;}

	// Parse the command line and return whether successful
	bool Parse(int argc,  TCHAR *argv[]);

	// Parse the command line and return the actions selected as a code; for testing; throws exception on error
	unsigned long ParseGetActions(int argc, TCHAR  *argv[]);

	// Display syntax help
	void ShowHelp(TCHAR *lptMessage=NULL);

private:
	// Structure to hold acceptable arguments for parsing
	struct ArgSwitch
	{
		TCHAR* pszSwitch;	// actual switch text
		TCHAR* pszDescription;	// description of switch
		bool fHasValue;	// whether is expects a switch value
		unsigned long flagAction;	// The flag for this action
	};

	// Structure to hold acceptable scope values for parsing
	struct ScopeValue
	{
		UserArgs::ActionScope nScopeCode;
		TCHAR* pszScope;
	};

	// Internal state enum used in Parse()
	enum PARSESTATE
	{
		STATE_NEWTOKEN,		// current position is on a new token
		STATE_SWITCH,		// current position is on the first character of a switch
		STATE_VALUE,		// current position is on the first character of a non-switch parameter
		STATE_ADVANCEARG,	// need to move the current position to the next token
		STATE_ADVANCECHAR,	// need to move the current position to the next character in current token
		STATE_VALUESWITCH,	// current position is on the first character of a switch's value (e.g. the M in "-switchname:MyValue"); switchname is stored
		STATE_DONE,			// finished parsing
		STATE_FAIL,			// could not parse the arguments
	};

	// Action to be populated when parsing; accessed indirectly via fXxx() methods
	unsigned long m_actions;

	// Flags for actions; these are private, accessed indirectly through accessor methods below
	// Update this list to add actions (i.e. switches) to the application
	static const unsigned long CHECKFOLDERACLS=1<<0;
	static const unsigned long FIXFOLDERACLS=1<<1;
	static const unsigned long CHECKITEMS=1<<2;
	static const unsigned long FIXITEMS=1<<3;
	static const unsigned long SCOPE=1<<4;	// Not an action, do not add to _actions
	static const unsigned long DISPLAYHELP=1<<31;	// User requested help in a valid syntax

	// Error codes
	static const unsigned long ERR_NONE=0;
	static const unsigned long ERR_INVALIDSWITCH=1<<0;
	static const unsigned long ERR_DUPLICATEFOLDER=1<<1;
	static const unsigned long ERR_EXPECTEDSWITCHVALUE=1<<2;
	static const unsigned long ERR_INVALIDSTATE=1<<3;

	// public folder, scope of action
	tstring *m_pstrPublicFolder;
	ActionScope m_scope;
	

	// An array of switch structures defining the acceptable switches
	static const ArgSwitch rgArgSwitches[_COUNTOFACCEPTEDSWITCHES];
	
	// A string of characters we accept as switch escapes; typically "-/"
	static const tstring *m_pstrSwitchChars;	

	// An array of Scope structures defining acceptable scopes
	static const ScopeValue rgScopeValues[_COUNTOFSCOPES];

	// Initialize internal data structures for each round of parsing
	void init();

	// Log an error about an argument
	void logError(int iArg, const TCHAR *arg, unsigned long err);

	// Validate the parsed parameters
	bool validate();

};



