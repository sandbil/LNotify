==============================================================================
Lotus Notes object layer
Version 3.11 beta 0
Created by:
   Sergey Kolchin (Russia) aka kol skolchin@yahoo.com
   Sergey Kucherov (Russia)                                 
   Sergey Okorochkov (Russia)
  Last Modified:
    17.11.2011, sandbil
==============================================================================


==============================================================================
Description
==============================================================================
This class library allows to work with the Lotus Notes from Delphi programs. It covers almost all aspects of Notes functionality:
- databases (open/create, search)
- mail (send/receive)
- documents (read/create, send)
- items (read/create, RTF)
- database directories (browse)
- hierarchical names (parsing/combining, lookups)
- access control list support;
- password interception;
The library supports Delphi 3, 4, 5, 7 and Notes 4.5+, 5, 8. Complete help file and example project are available

==============================================================================
New in this version
==============================================================================
Added the library supports for Delphi 7 and Notes 8.

- function UnreadDocumentsUserName
- function TimeDatetoStr 
- Function UserNameFromID
- REGIDGetxxx - Information type codes for ID files.
- function REGGetIDInfo
- function SECKFMSwitchToIDFile
- function NSFDbGetUnreadNoteTable2



The latest original version is always available at http://www.geocities.com/skolchin
Please, send bugs, comments and suggestions (except as mentioned before) to Sergey Kolchin (skolchin@yahoo.com).

The latest modification's version is available at https://github.com/sandbil/LNotify


==============================================================================
Installation
==============================================================================
- Unzip all *.pas and *.dfm files to any directory and include it into Delphi search path
- Make sure that Lotus Notes directory (where NLNOTES.DLL resides, like C:\Notes) is in PATH
- Once again: this is NOT a component library, so DO NOT try to install it as a package

To install the help file, use freeware utility HLPREG (written by Moskw@), which is
available from the library site (see N/D Library downloads page).

To install the help file manually do the following (N is Delphi version 3-5):
- go to Delphi help dir (C:\Program Files\Borland\DelphiN\Help)
- put LN.HLP and LN.CNT files there
- open DelphiN.cnt with Notepad
- add a line :Include LN.cnt to the end of the file
- save and close the file
- delete DelphiN.gid file

==============================================================================
Examples
==============================================================================
In this version, example program is in a separate zip file LNHELP.ZIP. Unzip all files from this archive to any 
directory, except TEST.NSF, which must be in Notes' data directory (like C:\Notes\Data). Main project file is 
LNHELP.DPR.
Added source of example program LNotify. Main project file is 
LNotify.DPR

==============================================================================
Disclaimer
==============================================================================
THIS LIBRARY IS PROVIDED "AS-IS". NO WARRANTIES OF ANY KIND, EXPRESSED OR IMPLIED, ARE MADE AS TO IT OR 
ANY MEDIUM IT MAY BE ON. NO REMEDY WILL BE PROVIDED FOR INDIRECT, CONSEQUENTIAL, PUNITIVE OR
INCIDENTAL DAMAGES ARISING FROM IT, INCLUDING SUCH FROM NEGLIGENCE, STRICT LIABILITY, OR BREACH OF 
WARRANTY OR CONTRACT, EVEN AFTER NOTICE OF THE POSSIBILITY OF SUCH DAMAGES.

The library is distributed as an open source project under Mozilla public license. See MPL10.TXT file for details.

==============================================================================
History
==============================================================================
Version 3.10 beta 7 (16.12.2003)

- property TNotesACL.MaximumInternetAccess
- TNotesServer class: support for server statistics retrieval (Daniel)
- bugfix in TNotesDocument.CreateProfile
- TNotesDBScript class for Database Scripts support (Daniel)
- TNotesDesignNote.IsXXX properties: easy note type access

==============================================================================
Version 3.10 beta 6 (29.05.2003)

- small bugfix in GetUserActivity (Daniel)
- TNotesRichTextItem in case of multi-items loads attachment data incorrectly so its BLOCKid moved to last item in sequence. This results in that the item sequence cannot be correctly processed (CreateNext or LoadNextItem fail with 'No such item' error)
6a update: Conditional symbol NOTES_R4 which allows to exclude all functions new to Notes R5. This affects Class_LotusNotes, Class_NotesDesign and Util_LnApi units
- bug fix: TNotesItem.GetAsString: #0#0 now treated as empty string, not a text end. This was done in order to support string lists
- bug fix: TNotesItem.GetAsString doesn't reflect text item changes. This was happening due to the fact that ValueLength property was not changed when AsString is used.
- new function TNotesItem.Refresh - re-reads item properties and data from note
- new property TNotesItem.AsShortString - differs from AsString in that it DOESN'T support huge items (> 64KB) because it uses basic API functions. This function is useful when operating with pre-defined Notes fields which cannot be large and don't need extra care. As a side effect - changes made throught this property are immediatelly available to read using it or AsString
- new constructor TNotesNote.CreateFromHandle which is helpful when a note handle is obtained elsewhere (from Menu add-in for example). The handle supplied is never released!
- new function TNotesDatabase.OpenByHandle - similar to previous
- properties TNotesDatabase.UserActivity, ActivityInfo - return User- and DB-activity (summary and detailed access log) - by Daniel
- bug fix: TNotesItem.GetAsList incorrectly processes TIME_RANGE items (by Daniel)
- new function TNotesRichTextItem.AddFormulaButton (by Andreas Pape)
- full support for OLE attachments in TNotesRichTextItem: counted in AttachCount, extracted with Detach(). Only one problem - I don't know yet how to get original file name from OLE attachment (actually I suppose there's no such way), so OLE attachments receive some cryptic names and have to be renamed when extracted
- new functionality in TNotesRichTextItem.AttachOleObject - if aHint = '', the function marks attachment with application icon (the same way as Attach with fIcon = true)
- bug fix: TNotesRichTextItem.LoadAttachInfo now correctly handles multi-item attachments
- rewritten Agent and ScriptLib support classes (NotesDesign unit) in attempt to simplify the code and class structure. Several properties, functions were added, my only hope I didn't break it functionality - please test!
- LnHelp project updated to support ScriptLib and Agents creation testing

Version 3.10 beta 5 (11.11.2002)
- full support of Notes ScriptLibs and Agents: Java code, compiled Java agents, ability to change code, code compilation etc. Since this was done solely by Daniel Lehtihet

Version 3.10 beta 4 (30.08.2002)

- NotesItem.GetAsString reverted back to original version dure to incompability with
  Eastern code pages
- NotesInitExt now works! When a program is started with '=<ini path>' command line switch,
  this INI file will be used to initialize Notes. Define EXTENDED_INIT symbol in order to use
  this function
- New function TNotesRichTextItem.AddButton (by Andreas Pape)
- New functions: TNotesView.CreateSearch, .GetAllDocumentsByKey for quick view lookups
- New property: TNotesNote.Created

Version 3.10 beta 2-3 (24.07.2002)

- Bug fixing in TNotesScriptLib.SourceCode (returns incomplete source code if it is >64K). This was again multi-item issue, but I found that standard functionality doesn't work and LoadNextItem returns already processed item. I had to use NSFItemScan in order to get all code items. This could probably affect other library parts as well - if anybody have had such problems, please respond
- Changes in TNotesAgent - properties LastRun, DocumentsProcessed and ExitCode are now typed as TDateTime, DWORD, DWORD corresp. and return correct values
- Memory leak fixed in TNotesAgent.GetAgentRunInfo
- New functions in TNotesDocument: ResponseCount, ParentDocumentID, Responses by Matt Saint
- attempt to fix minor bug in TNotesItem.GetAsString (invalid chars appear at the end of text when the function is used on RTF items)

Version 3.10 beta 1 (02.07.2002)

- New class TNotesNote, which is now a parent of TNotesDocument and contains methods common to all notes, not only documents. Just to remind - Notes note is a basic object, which is everything - design objects, agents, documents, mail etc.
- New unit Class_NotesDesign.pas, which is supposed to contain design-related classes. New classes TNotesAgent and TNotesScriptLib written by Daniel Lehtihet are now in this unit, along with some others, which provide basic functionality for these two.
- New unit Utils_NotesAddin.pas written by Daniel Lehtihet adding support for server add-ins (these are console applications running directly inside Domino server )
- Bug fixing: "Error: document deleted when opening from collection" by Matt Saint. New property TNotesNote.IsDeleted added to support this
- Bug fixing: "Error when importing Rich Text files" found by Andreas Menge. Actually this is just a workaround, but at least it works
- Bug fixing: "Cannot send mail when only Internet location is active" found by Alex SW. New class property TNotesDatabase.MailType added to support this
- Other changes/fixes collected for long period, I just don't remember them :(


Version 3.9.1 (25.08.01)
- Fixed bug in TNotesRichTextItem.CreateFromFile: the function was assuming that file passed in is always RTF ignoring extension
- New functions in TNotesView: OpenViewExt, CreateExt, which allow to limit returned documents by key (value of 1st column) or count - by piton
- New unit: Util_LnFaxApi.pas containing translation of Notes FAX Api (no backend classes, it's very easy)
- New function: TNotesRichTextItem.AddPassthroughHtml

Version 3.9 (29.11.00)
- support for multiple items with the same name: TNotesItem.CreateNextItem
- fixed bug in TNotesView.GetViewID (unable to open a view with alias)
- conversion between Notes and Delphi-style text lists (#0 <-> #13#10). This also fixes problems in reading/writing lists to text items: NotesToString, StringToNotes
- color conversion functions: NotesToColor and ColorToNotes
- fixed bug in TNotesItem.SetAsStrings (text list contains unreadable chars);
- fixed bugs in TNotesDatabase.MarkRead, TNotesDatabase.MarkAllRead (mark read doesn't work);
- new unit Notes_AddInMenu: support for Notes UI add-in menus
- fixed LMBCS problems in different functions (by tty)

Version 3.8 (29.02.00)
- Notes password prompt replacement (LNPASS) is included in the library: (see SetPassword function) 
- Fixed RTF table creation bug reported (InnerCellSpace value); 
- Fixed minor TNotesItem.GetAsReference bug; 
- New unit: Class_NotesRepl - database replication support by Dmitry Mokrushin; 
- New function HugeNsfRtfItemAppend by Olaf Hahl is now used in RTF support class. The function splits huge RTF item to smaller parts (required by Notes 4) 
- New parameter FDelmiters in UNIDToStr function. For Delphi 3 users: change all references to the function, please 
- Fixed TNotesDocument.Send 'CC' and 'Bcc' bugs; 
- Fixed TNotesDatabase.UnreadDocuments bug; 
- New functions for RTI enumeration: TNotesRichTextItem.ReadItem and others; 
- New properties: TNotesRichTextItem.LinkCount - hotspot link enumeration; 
- New feature: TNotesRichTextItem.Attach can now display attached file as an icon as normal Notes client; 
- View formatting constants and types from VIEWFMT.H are now included in Util_LnApi.pas; 
- Delphi 5 support now works properly

Version 3.7 (25.10.99) 
- New functions: TNotesRichTextItem.StartSection, TNotesRichTextItem.EndSection 
- Fixed bug with GetExeFileName function 
- Changed terms of distribution

Version 3.6 (20.09.99)
- Control-break handling
- New functions: GetNotesExeDir, GetNotesIniFile, GetNotesDataDir

Version 3.5 (30.07.99)
- RTI classes are now in a separate unit Class_NotesRTF.pas
- fixed bug in a TNotesFolder.AddDocuments
- rewritten RTI import mechanism (thanks, Andy);
- new method: TNotesRichTextItem.AddUrl for adding URL hotspots
- new methods: TNotesRichTextItem.AddRtfFile (replaces LoadRtfFile), TNotesRichTextItem.AddJpgFile, 
    TNotesRichTextItem.AddFile
- new property: TNotesDatabase.QuotaInfo

Version 3.4 (17.06.99)
- Numerious bug fixes (most serious: font table overflow);
- RT Items bigger than 64K are supported now;
- form attachments: new TNotesDocument.AttachForm method, new parameter to TNotesDocument.Send method;
- TNotesDocument.CreateResponseByUNID constructor
- TNotesDocument.CopyItem method for copying items between documents

Version 3.3d (06.05.99)
- bug fixes: TNotesItem.Is... properties now work with text and text lists.

Version 3.3c (30.03.99)
- RTF tables support

Version 3.3b (09.03.99)
- small bug fixes

Version 3.3a (09.03.99)
- compatible with Delphi 3
- Anchor links in Notes rich-text items
- Summary buffers support by the Search function of TNotesDocument

Version 3.3 (25.02.99)
- TNotesACL, TNotesACLEntry classes
- Export/import to Windows RTF via TNotesRichTextItem class
- Access to view columns using TNotesDocument.SummaryValues property
- Document/View/Database links in TNotesRichTextItem

Version 3.2f (10.11.98)
- access to Notes folders
- access to profile documents

Version 3.2e (04.11.98)
- Note ID functions and properties
- responses and replies
- view access

Version 3.2c (15.10.98)
- TNotesDocument.Evaluate function

Version 3.2 (11.08.98)
- TNotesName class: hierarchical names, name lookups
- unread documents handling: TNotesDatabase.UnreadDocuments
- support for national texts: Lmbcs2Native, Native2Lmbcs functions

Version 3.1 (26.05.98)
- new TNotesDirectory class and LnBrowse function for browsing any server for a database
- ReloadFields method of TNotesDocument class
- memory leaks and other bugs catched (thanks for Igor Naumov)
- new functions in Util_NotesAPI unit

Version 3.0 (17.02.98) - I (Sergey Kolchin) got full responsibility of the library, so I changed a lot
- new design of TNotesDocument class (now working through Items property)
- Sign and GetSignature methods of TNotesDocument class
- significant changes of TNotesItem class (As... properties)
- new methods and properties of TNotesDatabase class: CreateNew, CopyRecords, MailServer, MailFileName, UserName etc
- new TNotesRichTextItem class
- bugs found and killed
- new functions in Util_NotesAPI unit

Version 2.0 (August 97)
- completelly new design similar to Lotus Script classes set
- new TNotesDatabase, TNotesDocument, TNotesItem, TNotesDocumentCollection classes
- Util_LnApi unit re-created from Notes API using C2Pas utility (thanks for Anatoly Ivkov)

Version 1.0 (January 97)
- TLotusNotes component allows to send mail and read databases
- Basic API functions in Util_LnApi unit
