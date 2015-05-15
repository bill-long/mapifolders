# MAPIFolders

## What's it for?

MAPIFolders does a few specific things:

* Adds or removes permissions on public folders quickly. For example, in my lab it can add permissions to 10,000 public folders in less than 3 minutes.
* Checks for and fixes certain ACL problems.
For more information on using MAPIFolders, see the [wiki](https://github.com/bill-long/mapifolders/wiki).

## What's the difference between MAPIFolders and MFCMapi?

They're completely different. For starters:

* MFCMapi is a GUI tool. MAPIFolders is a command line tool.
* MFCMapi is good example of how to properly write MAPI code in C++. MAPIFolders is not. I have little experience writing C++, and I certainly don't recommend anyone use this as an example of the right way to do things. Look to MFCMapi for that.
* MFCMapi can do almost anything MAPIFolders can do, but it does it one folder at a time. It's not intended for bulk modifications of thousands of folders. By contrast, MAPIFolders exposes only a narrow set of MAPI functionality, but it is intended to perform those operations across thousands of folders at a time.

## What's the difference between MAPIFolders and ExFolders or PFDAVAdmin?

* MAPIFolders uses MAPI, which means it can be run from any machine with Outlook - no ancient .NET Framework 1.1 dependencies (PFDAVAdmin) or having to be run from an Exchange server (ExFolders).
* MAPIFolders will theoretically work against Exchange 2013, 2010, 2007, and 2003.
* MAPIFolders has no GUI. Anyone want to write one?
* MAPIFolders currently does not do as much as ExFolders or PFDAVAdmin, but hopefully that will change over time.
