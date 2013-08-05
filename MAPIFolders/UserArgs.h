#pragma once
#include <string>

#define _COUNTOFACCEPTEDSWITCHES (6)
#define _COUNTOFSCOPES (3)

// Internal structure for accepted switches

class UserArgs
{
public:
	UserArgs();
	~UserArgs();

	// Action accessor methods; use these to determine user selected actions after parsing
	// Update this list to add actions to the application
	bool fCheckFolderAcl() {return (NULL!=(_actions & CHECKFOLDERACLS));}
	bool fFixFolderAcl() {return (NULL!=(_actions & FIXFOLDERACLS));}
	bool fCheckItems() {return (NULL!=(_actions & CHECKITEMS));}
	bool fFixItems() {return (NULL!=(_actions & FIXITEMS));}
	bool fDisplayHelp() {return (NULL!=(_actions & DISPLAYHELP));}

	// Argument value accessor methods
	std::wstring *pstrPublicFolder() {return _pstrPublicFolder;}

	// Enum for the possible scopes for operations
	enum ActionScope : short int
	{
		NONE=0,	// Default, uninitialized; not a selection for user
		BASE,
		ONELEVEL,
		SUBTREE
	};

	// Get the parsed scope
	ActionScope nScope() {return _scope;}

	// Parse the command line and return whether successful
	bool Parse(int argc,  wchar_t *argv[]);

	// Parse the command line and return the actions selected as a code; for testing; throws exception on error
	unsigned long ParseGetActions(int argc, wchar_t  *argv[]);

private:
	// Structure to hold acceptable arguments for parsing
	struct ArgSwitch
	{
		wchar_t* pszSwitch;	// actual switch text
		wchar_t* pszDescription;	// description of switch
		bool fHasValue;	// whether is expects a switch value
		unsigned long flagAction;	// The flag for this action
	};

	// Structure to hold acceptable scope values for parsing
	struct ScopeValue
	{
		UserArgs::ActionScope nScopeCode;
		wchar_t* pszScope;
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
	unsigned long _actions;

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
	std::wstring *_pstrPublicFolder;
	ActionScope _scope;
	

	// An array of switch structures defining the acceptable switches
	static const ArgSwitch rgArgSwitches[_COUNTOFACCEPTEDSWITCHES];
	
	// A string of characters we accept as switch escapes; typically "-/"
	static const std::wstring _wstrSwitchChars;	

	// An array of Scope structures defining acceptable scopes
	static const ScopeValue rgScopeValues[_COUNTOFSCOPES];

	// Initialize internal data structures for each round of parsing
	void init();

	// Log an error about an argument
	void logError(int iArg, const wchar_t *arg, unsigned long err);

};



