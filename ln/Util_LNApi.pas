{==============================================================================|
| Project : Notes/Delphi class library                           | 3.10        |
|==============================================================================|
| Content:                                                                     |
|==============================================================================|
| The contents of this file are subject to the Mozilla Public License Ver. 1.0 |
| (the "License"); you may not use this file except in compliance with the     |
| License. You may obtain a copy of the License at http://www.mozilla.org/MPL/ |
|                                                                              |
| Software distributed under the License is distributed on an "AS IS" basis,   |
| WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for |
| the specific language governing rights and limitations under the License.    |
|==============================================================================|
| Initial Developers of the Original Code are:                                 |
|   Sergey Kolchin (Russia) skolchin@usa.net ICQ#2292387                       |
|   Sergey Kucherov (Russia)                                                   |
|   Sergey Okorochkov (Russia)                                                 |
| All Rights Reserved.                                                         |
|   Last Modified:                                                             |
|     23.05.03, Sergey Kolchin                                                 |
|==============================================================================|
| Contributors and Bug Corrections:                                            |
|   Fujio Kurose                                                               |
|   Noah Silva                                                                 |
|   Tibor Egressi                                                              |
|   Andreas Pape                                                               |
|   Anatoly Ivkov                                                              |
|   Winalot                                                                    |
|   Daniel Lehtihet                                                            |
|     and others...                                                            |
|==============================================================================|
| History: see README.TXT                                                      |
|==============================================================================|
  Notes API functions
  Transfered from Lotus Notes C-API using:
    C2PAS
    HeadConv 3.25 (c) 1998 by Bob Swart (aka Dr.Bob - www.drbob42.com)
    HeadConv 4.0 (c) 2000 by Bob Swart (aka Dr.Bob - www.drbob42.com)     

    nssdata.h
    nsfnote.h
    nsfsearch.h
    editods.h
    easycd.h
    osmisc.h
    kfm.h
    osmem.h
    osenv.h
    osFile.H
    Nif.h
    OsTime.h
    TextList.h
    nsfdb.h
    event.h
    mailserv.h
    ft.h
    idtable.h
    ods.h
    misc.h
    fontid.h
    ossignal.h
    ns.h
    lookup.h
    dname.h
    foldman.g
    acl.h
    ixedit.h
    ixport.h
    repl.h
    colorid.h
    viewfmt.h
    dbdrv.h
    nsfole.h
    oleods.h

    Enable NOTES_R4 conditional symbol to maintain Notes R4 compatibility
|==============================================================================|}
unit Util_LNApi;

{$WEAKPACKAGEUNIT ON}
{$ALIGN OFF}
{.$DEFINE NOTES_R4}

interface

Uses Windows, ActiveX, Util_LnApiErr;

const
  NOTES_DLL_NAME = 'NNOTES.DLL';

(******************************************************************************)
{ Note storage file data definition}
{ from NSFData.H }
(******************************************************************************)

type
  USHORT = WORD;
  PUSHORT = ^USHORT;
  STATUS = word;
  PSTATUS = ^STATUS;

const NOTEID_RESERVED = longint($80000000); { Reserved Note ID, used for}
      NOTEID_ADD = longint(0);
      NOTEID_ADD_OR_REPLACE = longint($80000000);
      NOTEID_ADD_UNID = longint($80000001);

{ An RRV "file position" is defined to be a DWORD, 4 bytes long. }

const RRV_ALIGNMENT = 4; { most typical RRV alignment (DBTABLE.C) }
const RRV_DELETED = NOTEID_RESERVED; { indicates a deleted note (DBTABLE.C) }

const NOTEID_NO_PARENT = $00000000; { Reserved Note ID, used to indicate}
{This is the structure that identifies a database. It is used for both }
{the creation date/time and the originator date/time. }


{ This is the structure that identifies a note within a database. It is }
{simply a file position (RRV) that is guaranteed never to change WITHIN }
{this file. A replicated note, however, may have a different note id }
{in two separate files. }

type NOTEID = LongInt;

{ This is the structure that identifies ALL replicas of the same note. The }
{"File" member contains a totally unique (random) number, generated at }
{the time the note is created. The "Note" member contains the date/time }
{when the very first copy of the note was stored into the first NSF. The }
{"Sequence" member is a sequence number used to keep track of the most }
{recent version of the note for replicated data purposes. The }
{"SequenceTime" member is a sequence number qualifier, that allows the }
{replicator to determine which note is later given identical Sequence's. }
{Both are required for the following reason. The sequence number is needed }
{to prevent someone from locking out future edits by setting the time/date }
{to the future. The sequence time qualifies the sequence number for two }
{reasons: 1) It prevents two concurrent updates from looking like }
{no update at all and 2) it forces all systems to reach the same decision }
{as to which update is the "latest" version. }

{Time/dates associated with notes: }

{OID.Note Timedate when the note was created }
{Obtained by NSFNoteGetInfo(_NOTE_OID) or }
{OID in SEARCH_MATCH. }
{OID.SequenceTime Timedate of last revision }
{Obtained by NSFNoteGetInfo(_NOTE_OID) or }
{OID in SEARCH_MATCH. }
{NOTE.EditModified Timedate when added to (or last updated in) }
{this database. }
{(Obtained by NSFNoteGetInfo(_NOTE_MODIFIED) or }
{ID.Note in SEARCH_MATCH. }


{ }

const OID_SEQNO_MASK = $0000FFFF; { Mask used to extract sequence # }
const OID_NO_REPLICATE = $80000000; { Never replicate outward, currently used ONLY for deleted stubs }

{ Replication flags }

{NOTE: Please note the distinction between REPLFLG_DISABLE and }
{REPLFLG_NEVER_REPLICATE. The former is used to temporarily disable }
{replication. The latter is used to indicate that this database should }
{NEVER be replicated. The former may be set and cleared by the Notes }
{user interface. The latter is intended to be set programmatically }
{and SHOULD NEVER be able to be cleared by the user interface. }

{The latter was invented to avoid having to set the replica ID to }
{the known value of REPLICA_ID_NEVERREPLICATE. This latter method has }
{the failing that DBs that use it cannot have DocLinks to them. }

{ 0x0001 spare was COPY_ACL }
{ 0x0002 spare }
const
   REPLFLG_DISABLE = $0004; { Disable replication }
   REPLFLG_IGNORE_DELETES = $0010; { Don't propagate deleted notes when}
   REPLFLG_DO_NOT_CATALOG = $0040; { Do not list in catalog }
   REPLFLG_CUTOFF_DELETE = $0080; { Auto-Delete documents prior to cutoff date }
   REPLFLG_NEVER_REPLICATE = $0100; { DB is not to be replicated at all }
   REPLFLG_ABSTRACT = $0200; { Abstract during replication }
   REPLFLG_DO_NOT_BROWSE = $0400; { Do not list in database add }
   REPLFLG_NO_CHRONOS = $0800; { Do not run chronos on database }
   REPLFLG_IGNORE_DEST_DELETES = $1000; { Don't replicate deleted notes}
   REPLFLG_PRIORITY_MED = $0000; { Medium priority }
   REPLFLG_PRIORITY_HI = $4000; { High priority }
   REPLFLG_PRIORITY_SHIFT = 14; { Shift count for priority field }
   REPLFLG_PRIORITY_MASK = $0003; { Mask for priority field after shifting}
   REPLFLG_PRIORITY_INVMASK = $3fff; { Mask for clearing the field }
   REPLFLG_USED_MASK = ($4 or $8 or $10 or $40 or $80 or $100 or $200 or $C000 or $1000 or $4000);


{ Replication priority values are stored in the high bits of the }
{replication flags. The stored value is biased by -1 so that }
{an encoding of 0 represents medium priority (-1 is low and +1 is high). }
{The following macros make getting and setting the priority easy. }
{They return and accept normalized values of 0 - 2. }


//const REPL_GET_PRIORITY(Flags) = \;

//function T(var REPLFLG_PRIORITY_MASK): ((FLAGS >> REPLFLG_PRIORITY_SHIFT)+1); { }
//           REPL_SET_PRIORITY(Pri: CONST): T; stdcall; far;

//function T(1: ((PRI -): T; stdcall; far;
{ Reserved ReplicaID.Date. Used in ID.Date field in ReplicaID to escape }
{to reserved REPLICA_ID_xxx }

{ }

function REPL_GET_PRIORITY (Flags: word): word;
function REPL_SET_PRIORITY (Pri: word): word;

const
  REPLICA_DATE_RESERVED = 0; { If used, see REPLICA_ID_xxx Known Replica IDs. Used in ID.Time field in ReplicaID. Date
                              subfield must be REPLICA_DATE_RESERVED). NOTE: If you add to this list,
                              you should check the code in \catalog\search.c to see if the new one(s)
                              need to be added to that code (probably not - but worth checking).
                              The format is as follows. Least sig. byte is version number. 2nd
                              byte represents package code but is hard-coded to protect against
                              changes in the package code. Most sig. 2 bytes are reserved for future
                              use. }
  REPLICA_ID_UNINITIALIZED = $00000000; { Uninitialized ID }
  REPLICA_ID_CATALOG = $00003301; { Database Catalog (Version 2) }
  REPLICA_ID_EVENT = $00003302; { Stats & Events Config DB }
{ The following known replica ID is now obsolete. Although the replicator
  still supports it, the problem is that DBs that use it cannot have
  DocLinks to them. Instead use the replica flag REPLFLG_NEVER_REPLICATE. }
  REPLICA_ID_NEVERREPLICATE = $00001601; { Do not allow replicas }

{Number of times within cutoff interval that we purge deleted stubs.
For example, if the cutoff interval is 90 days, we purge every 30
days.}

  CUTOFF_CHANGES_DURING_INTERVAL = 3;

{This is the structure that identifies a replica database. }

type
  TIMEDATE = packed record
    T1,T2: longint;
  end;
  DBReplicaInfo = packed record
    ID: TIMEDATE; { ID that is same for all replica files }
    Flags: Word;  { Replication flags }
    CutoffInterval: Word; { Automatic Replication Cutoff Interval (Days) }
    Cutoff: TIMEDATE; { Replication cutoff date }
  end;
  PDBReplicaInfo = ^DBReplicaInfo;
  DBID = TimeDate;
    {This is the structure that globally identifies an INSTANCE of a note,
    that is, if we are doing a SEARCH_ALL_VERSIONS, the one with the
    latest modification date is the one that is the "most recent" instance. }
  PDBID = ^DBID;
  GLOBALINSTANCEID = packed record
    aFile: DBID; { database Creation time/date }
    Note: TIMEDATE; { note Modification time/date }
    NoteID: NOTEID; { note ID within database }
  end;

    { This is the structure that universally (across all servers) describes
    a note (ALL INSTANCES of the same note), but without the information
    necessary to directly access the note in a given database. It is used
    for referencing a specific note from another note (response notes and
    hot buttons are examples of its use) by storing this structure in the
    referencing note itself. It is intended to work properly on any server,
    and even if the note being referenced is updated. Matching of notes
    to other notes is done via the NIF machinery. }
  UNIVERSALNOTEID = packed record
    aFile: DBID; {Unique (random) number (Even though this field is called "File",}
                 {it doesn't have anything to do with the file!) }
    Note: TIMEDATE; { Original Note Creation time/date }
  end;

  UNID = UNIVERSALNOTEID;

const
  BlankUNID: UNID = (aFile: (T1: 0; T2: 0); Note: (T1: 0; T2: 0));

type
  {This is the structure that universally (across all servers) describes }
  {a note LINK. }
  NOTELINK = packed record
    aFile: TIMEDATE;{ File's replica ID }
    View: UNID; { View's Note Creation TIMEDATE }
    Note: UNID; { Note's Creation TIMEDATE }
  end {_5};
  PNOTELINK = ^NOTELINK;

{ Data type Definitions. }

{ Class definitions. Classes are defined to be the }
{"generic" classes of data type that the internal formula computation }
{mechanism recognizes when doing recalcs. }

const
 CLASS_NOCOMPUTE = (0 shl 8);
 CLASS_ERROR = (1 shl 8);
 CLASS_UNAVAILABLE = (2 shl 8);
 CLASS_NUMBER = (3 shl 8);
 CLASS_TIME = (4 shl 8);
 CLASS_TEXT = (5 shl 8);
 CLASS_FORMULA = (6 shl 8);
 CLASS_USERID = (7 shl 8);

 CLASS_MASK = $ff00;

    {All datatypes below are passed to NSF in either host (machine-specific
    byte ordering and padding) or canonical form (Intel 86 packed form).
    The format of each datatype, as it is passed to and from NSF functions,
    is listed below in the comment field next to each of the data types.
    (This host/canonical issue is NOT applicable to Intel86 machines,
    because on that machine, they are the same and no conversion is required).
    On all other machines, use the ODS subroutine package to perform
    conversions of those datatypes in canonical format before they can
    be interpreted. }

   // "Computable" Data Types
   TYPE_ERROR = 0 + CLASS_ERROR; { Host form }
   TYPE_UNAVAILABLE = 0 + CLASS_UNAVAILABLE; { Host form }
   TYPE_TEXT = 0 + CLASS_TEXT; { Host form }
   TYPE_TEXT_LIST = 1 + CLASS_TEXT; { Host form }
   TYPE_NUMBER = 0 + CLASS_NUMBER; { Host form }
   TYPE_NUMBER_RANGE = 1 + CLASS_NUMBER; { Host form }
   TYPE_TIME = 0 + CLASS_TIME; { Host form }
   TYPE_TIME_RANGE = 1 + CLASS_TIME; { Host form }
   TYPE_FORMULA = 0 + CLASS_FORMULA; { Canonical form }
   TYPE_USERID = 0 + CLASS_USERID; { Host form }

   { "Non-Computable" Data Types }
   TYPE_INVALID_OR_UNKNOWN = 0 + CLASS_NOCOMPUTE; { Host form }
   TYPE_COMPOSITE = 1 + CLASS_NOCOMPUTE; { Canonical form, >64K handled by more than one item of same name concatenated }
   TYPE_COLLATION = 2 + CLASS_NOCOMPUTE; { Canonical form }
   TYPE_OBJECT = 3 + CLASS_NOCOMPUTE; { Canonical form }
   TYPE_NOTEREF_LIST = 4 + CLASS_NOCOMPUTE; { Host form }
   TYPE_VIEW_FORMAT = 5 + CLASS_NOCOMPUTE; { Canonical form }
   TYPE_ICON = 6 + CLASS_NOCOMPUTE; { Canonical form }
   TYPE_NOTELINK_LIST = 7 + CLASS_NOCOMPUTE; { Host form }
   TYPE_SIGNATURE = 8 + CLASS_NOCOMPUTE; { Canonical form }
   TYPE_SEAL = 9 + CLASS_NOCOMPUTE; { Canonical form }
   TYPE_SEALDATA = 10 + CLASS_NOCOMPUTE; { Canonical form }
   TYPE_SEAL_LIST = 11 + CLASS_NOCOMPUTE; { Canonical form }
   TYPE_HIGHLIGHTS = 12 + CLASS_NOCOMPUTE; { Host form }
   TYPE_WORKSHEET_DATA = 13 + CLASS_NOCOMPUTE; { Used ONLY by Chronicle product }
  { Canonical form }
   TYPE_USERDATA = 14 + CLASS_NOCOMPUTE; { Arbitrary format data - see format below }
  { Canonical form }
   TYPE_QUERY = 15 + CLASS_NOCOMPUTE; { Saved query CD records; Canonical form }
   TYPE_ACTION = 16 + CLASS_NOCOMPUTE; { Saved action CD records; Canonical form }
   TYPE_ASSISTANT_INFO = 17 + CLASS_NOCOMPUTE; { Saved assistant info }
   TYPE_VIEWMAP_DATASET = 18 + CLASS_NOCOMPUTE; { Saved ViewMap dataset; Canonical form }
   TYPE_VIEWMAP_LAYOUT = 19 + CLASS_NOCOMPUTE; { Saved ViewMap layout; Canonical form }
   TYPE_LSOBJECT = 20 + CLASS_NOCOMPUTE; { Saved LS Object code for an agent. }
   TYPE_HTML = 21 + CLASS_NOCOMPUTE; { LMBCS-encoded HTML, >64K handled by more than one item of same name concatenated }


{ This is the structure used for summary buffers }


type
  ITEM_TABLE = packed record
    Length: USHORT;{ total length of this buffer }
    Items: USHORT; { number of items in the table now come the ITEMs now comes the packed text }
  end;

  ITEM = packed record { used for item names and values }
    NameLength: USHORT;    { length of the item's name }
    ValueLength: USHORT;   { length of the value field }
  end;

  ITEM_NAME_TABLE = packed record
    Length: USHORT; {total length of this buffer }
    Items: USHORT;{number of items in the table now comes an array of WORDS representing }
                  {the lengths of the item names now comes the item names as packed text}
  end;

  ITEM_VALUE_TABLE = packed record
    Length: USHORT;{total length of this buffer }
    Items: USHORT; {number of items in the table now comes an array of WORDS representing }
                  {the lengths of the item values.now comes the item values as packed bytes }
  end;

  Object_Descriptor = packed record
   ObjectType: Word; { type of object (OBJECT_xxx) }
   RRV: LongInt;     { Object ID of the object in THIS FILE }
  end {_10};

  ITEM_DEFINITION_TABLE = packed record
    Length: Word;  {total length of this buffer }
    Items: Word;   {number of items in the table now come the ITEM_DEFINITION structures
                   now comes the packed text }
  end;
  pITEM_DEFINITION_TABLE = ^ITEM_DEFINITION_TABLE;

  ITEM_DEFINITION = packed record
    Spare: Word;     {unused }
    ItemType: Word;  {default data type of the item }
    NameLength: Word;{ length of the item's name }
  end;
  pITEM_DEFINITION = ^ITEM_DEFINITION;

const
  OBJECT_NO_COPY = $8000; { do not copy object when updating to new note or database }
  OBJECT_PRESERVE = $4000; { keep object around even if hNote doesn't have it when NoteUpdating }

{ Object Types, a sub-category of TYPE_OBJECT }

   OBJECT_FILE = 0; { File Attachment }
   OBJECT_FILTER_LEFTTODO = 3; { IDTable of 'done' docs attached to filter }
   OBJECT_UNKNOWN = $ffff; { Used as input to NSFDbCopyObject, }
{ NSFDbGetObjectInfo and NSFDbGetObjectSize. }
{ File Attachment definitions }

   HOST_MASK = $0f00; { used for NSFNoteAttachFile Encoding arg }
   HOST_MSDOS = (0 shl 8);{ CRNL at EOL, optional ^Z at EOF }
   HOST_OLE = (1 shl 8);{ unknown internal representation, up to app }
   HOST_MAC = (2 shl 8);{ potentially has resource forks, etc. }
   HOST_UNKNOWN = (3 shl 8);{ came inbound thru a gateway }
   HOST_HPFS = (4 shl 8);{ HPFS. Contains EAs and long filenames }
   HOST_OLELIB = (5 shl 8);{ OLE 1 Library encapsulation }
   HOST_BYTEARRAY_EXT = (6 shl 8);{ OLE 2 ILockBytes byte array extent table }
   HOST_BYTEARRAY_PAGE = (7 shl 8);{ OLE 2 ILockBytes byte array page }
   HOST_LOCAL = $0f00; { ONLY used as argument to NSFNoteAttachFile }
{ means "use MY os's HOST_ type }

   EFLAGS_MASK = $f000; { used for NSFNoteAttachFile encoding arg }
   EFLAGS_INDOC = $1000; { used to pass FILEFLAG_INDOC flag to NSFNoteAttachFile }

{ changed below from 0x00ff to 0x000f to make room for flags defined below }
   COMPRESS_MASK = $000; { used for NSFNoteAttachFile Encoding arg }
   COMPRESS_NONE = 0; { no compression }
   COMPRESS_HUFF = 1; { huffman encoding for compression }

   NTATT_FTYPE_MASK = $0070; { File type mask }
   NTATT_FTYPE_FLAT = $0000; { Normal one fork file }
   NTATT_FTYPE_MACBIN = $0010; { MacBinaryII file }
   NTATT_NODEALLOC = $0080; { Don't deallocate object when item is deleted }

   ATTRIB_READONLY = $0001; { file was read-only }
   ATTRIB_PRIVATE = $0002; { file was private or public }

   FILEFLAG_SIGN = $0001; { file object has signature appended }
   FILEFLAG_INDOC = $0002; { file is represented by an editor run in the document }


type
  FileObject = packed record
   Header: OBJECT_DESCRIPTOR;{ object header }
   FileNameLength: Word; { length of file name }
   HostType: Word; { identifies type of text file delimeters (HOST_) }
   CompressionType: Word; { compression technique used (COMPRESS_) }
   FileAttributes: Word;  { original file attributes (ATTRIB_) }
   Flags: Word; { miscellaneous flags (FILEFLAG_) }
   FileSize: LongInt; { original file size }
   FileCreated: TIMEDATE; { original file date/time of creation, 0 if unknown }
   FileModified: TIMEDATE; { original file date/time of modification }
{ Now comes the file name... It is the original }
{ RELATIVE file path with no device specifiers }
end {_11};


type
  FileObject_MACExt = packed record
             FileCreator: Array[0..4-1] of Char; { application that created the file }
             FileType: Array[0..4-1] of Char; { type of file }
             ResourcesStart: LongInt; { offset into the object at which resources begin }
             ResourcesLen: LongInt; { length of the resources section in bytes }
             CompressionType: Word;{ compression used for Mac resources }
             Spare: LongInt;{ 0 }
end;


type
  FileObject_HPFS = packed record
             EAStart: LongInt;{ offset into the object at which EAs begin }
             EALen: LongInt;{ length of EA section }
             Spare: LongInt; { 0 }
           end ;


{ @SPECIAL Escape Codes }

const ESCBEGIN = $7;
const ESCEND = $ff;

{ Index information structure passed into NSFTranslateSpecial to provide }
{index-related information for certain @INDEX functions, if specified. }



type
  IndexSpecialInfo = packed record
    IndexSiblings,  { # siblings of entry }
    IndexChildren,   { # direct children of entry }
    IndexDescendants: dword; { # descendants of entry }
    IndexAnyUnread: Word; { TRUE if entry "unread, or any descendants "unread" }
  end {_14};
  PIndexSpecialInfo = ^IndexSpecialInfo;
  {Function templates}
  function NSFTranslateSpecial(InputString: Pointer;
                             InputStringLength: Word;
                             OutputString: Pointer;
                             OutputStringBufferLength: Word;
                             NoteID: NOTEID;
                             IndexPosition: Pointer;
                             IndexInfo: PINDEXSPECIALINFO;
                             hUnreadList: THandle;
                             hCollapsedList: THandle;
                             FileTitle: Pchar;
                             ViewTitle: Pchar;
                             var RetLength: word): STATUS; stdcall; far;


(******************************************************************************)
{ Some definitions from old LN API unit}
(******************************************************************************)

const
    {Consts to find out from header C files}
    STRINGLEN=255;
    NOERROR=0;
    MAXWORD=65535;
    BodySize=32000;
  {$IFNDEF WIN32}
    LNTrue=1;
    LNFalse=0;
  {$else}
    LNTrue=True;
    LNFalse=False;
  {$ENDIF}

    USER_BREAK = 100; {Error code for Brake NSFSearch operation}
    USER_CANCEL = 38696;

    MAXUSERNAME=256;
    MAXENVVALUE=256;

    MAIL_SENDTO_ITEM_NUM =0;
    MAIL_COPYTO_ITEM_NUM= 1;
    MAIL_FROM_ITEM_NUM= 2;
    MAIL_SUBJECT_ITEM_NUM= 3;
    MAIL_COMPOSEDDATE_ITEM_NUM= 4;
    MAIL_POSTEDDATE_ITEM_NUM= 5;
    MAIL_BODY_ITEM_NUM= 6;
    MAIL_INTENDEDRECIPIENT_ITEM_NUM=7;
    MAIL_FAILUREREASON_ITEM_NUM= 8;
    MAIL_RECIPIENTS_ITEM_NUM= 9;
    MAIL_ROUTINGSTATE_ITEM_NUM= 10;
    MAIL_FORM_ITEM_NUM= 11;
    MAIL_SAVED_FORM_ITEM_NUM= 12;
    MAIL_BLINDCOPYTO_ITEM_NUM= 13;
    MAIL_DELIVERYPRIORITY_ITEM_NUM= 14;
    MAIL_DELIVERYREPORT_ITEM_NUM= 15;
    MAIL_DELIVEREDDATE_ITEM_NUM= 16;
    MAIL_DELIVERYDATE_ITEM_NUM= 17;
    MAIL_CATEGORIES_ITEM_NUM= 18;
    MAIL_FROMDOMAIN_ITEM_NUM= 19 ;
    MAIL_SENDTO_LIST_ITEM_NUM= 20;
    MAIL_RECIPIENTS_LIST_ITEM_NUM= 21;
    MAIL_RECIP_GROUPS_EXP_ITEM_NUM= 22;
    MAIL_RETURNRECEIPT_ITEM_NUM= 23;
    MAIL_ROUTE_HOPS_ITEM_NUM= 24;
    MAIL_CORRELATION_ITEM_NUM= 25;
    MAIL_ORIGINALPATH_ITEM_NUM= 26;
    MAIL_DELIVER_LOOPS_ITEM_NUM= 27;

    MAIL_MAILSERVER_ITEM = 'MailServer';
    MAXPATH_OLE = 260;
    MAXPATH = 100;
    MAILBOX_NAME = 'mail.box';
    SMTPBOX_NAME = 'smtp.box';
    MAIL_MAILFILE_ITEM = 'MailFile';
    MAIL_MEMO_FORM = 'Memo';
    MAIL_LOCATION_ITEM = 'Location';

   LNoTrue = True ;
   LNoFalse = False;


type
  RANGE = packed record
    ListEntries: USHORT;
    RangeEntries: USHORT;
    //now come the list entries
    //now come the range entries
    Data: array[0..0] of pchar;
  end;
  PRANGE = ^RANGE;

type
  PTIMEDATE = ^TIMEDATE;

const ISTRMAX = 5;
const YTSTRMAX = 32;

// Added 2003-08-27 - Daniel Lehtihet
type
  INTLFORMAT = packed record
        Flags: Word;
        CurrencyDigits: BYTE;
        Length: BYTE;
        TimeZone: integer;
        AMString: array[0..ISTRMAX-1] of char;       
        PMString: array[0..ISTRMAX-1] of char;
        CurrencyString: array[0..ISTRMAX-1] of char;
        ThousandString: array[0..ISTRMAX-1] of char;
        DecimalString: array[0..ISTRMAX-1] of char;
        DateString: array[0..ISTRMAX-1] of char;
        TimeString: array[0..ISTRMAX-1] of char;
        YesterdayString: array[0..YTSTRMAX-1] of char;
        TodayString: array[0..YTSTRMAX-1] of char;
        TomorrowString: array[0..YTSTRMAX-1] of char;
  end;
  PINTLFORMAT = ^INTLFORMAT;

  TLNPriorites =  ( High, Normal, Low ); {by first letter}
  TLNReports   =  ( NoReport, Basic, Confirm );
  TLNQueryType =  (Document, Form, View);

  pBYTE= ^byte;
  NUMBER = double;
  HANDLE = integer;
  LHandle = HANDLE;
  pHANDLE = ^LHandle;
  ITEMDEFTABLEHANDLE = HANDLE;
  PITEMDEFTABLEHANDLE = ^ITEMDEFTABLEHANDLE;
  FORMULAHANDLE = HANDLE;
  PFormulaHandle = ^FormulaHandle;

  {$IFNDEF WIN32}
    BOOL= wordbool;
  {$else}
    ArgumentsArray = array[0..255] of PChar;
    PArgument = ^ArgumentsArray;
    BOOL= longbool;
  {$ENDIF}

  pNOTEID = ^NOTEID;
  pUNID = ^UNID;
  pITEM_TABLE=^ITEM_TABLE;
  pITEM = ^ITEM;
  pITEM_NAME_TABLE = ^ITEM_NAME_TABLE;
  pITEM_VALUE_TABLE = ^ITEM_VALUE_TABLE;

  OriginatiorID=record
     FileNum: DBID;  { this is number. It probably is integer}
     Note: TIMEDATE;
     Sequence: longint;
     SequenceTime: TIMEDATE;
  end;

  OID=OriginatiorID;
  OIDcl=class
    OIDItem: OID;
  end;

  LIST = packed record
    ListEntries: USHORT ;      {* list entries following }
  end;        {* now come the list entries }
  ptrLIST= ^LIST;
  pLIST= ^LIST;

  pOID= ^OID;
  pWORD=^word;
  plongInt=^longint;
  idList=array [0..255]of NOTEID;
  pIdList=^idList;

  DARRAY = packed record
    ObjectSize: WORD ;             { Total array object size }
    ElementsUsed: WORD ;           { Elements in use }
    ElementsFree: WORD ;           { Free elements }
    ElementsFreeMax: WORD ;        { Maximum free elements }
    ElementsFreeExtra: WORD ;      { Extra free elements to maintain }
    ElementSize: WORD ;            { Element size in bytes }
    ElementStrings: WORD ;         { Number packed string descriptors
                                        in each element }
    StringStorageOffset: WORD ;    { Offset to packed string storage }
    StringStorageUsed: WORD ;      { In use bytes of string storage }
    StringStorageFree: WORD ;      { Free bytes of string storage }
    StringStorageFreeMax: WORD ;   { Maximum free storage }
    StringStorageFreeExtra: WORD ; { Extra free storage to maintain }

    { First array element follows here.  First byte of packed string
      storage follows last allocated array element. */
     }
  end;
  pDARRAY=^DARRAY;
  ppDARRAY=^pDARRAY;

  DWORD = longint;
  pDWORD = ^DWORD;
  BLOCK = word;

  BLOCKID = packed record
    pool: integer;  { pool handle }
    block: BLOCK;   { block handle }
  end;
  pBLOCKID=^BLOCKID;

const
  NullBid: BLOCKID = (pool: 0; block: 0);

type
  BLOCKIDcl = class
    blockitem: BLOCKID;
  end;

  TFMT=record
   Date: BYTE;        { Date Display Format }
   Time: BYTE;        { Time Display Format }
   Zone:  BYTE;       { Time Zone Display Format }
   Structure: BYTE;   { Overall Date/Time Structure }
  end;
  pTFMT=^TFMT;

  PTIMESTRUCT = ^TIMESTRUCT;
  TIMESTRUCT = packed record
    year,          // 1-32767
    month,         //* 1-12 */
    day,           //* 1-31 */
    weekday,       //* 1-7, Sunday is 1 */
    hour,          //* 0-23 */
    minute,        //* 0-59 */
    second,        //* 0-59 */
    hundredth,     //* 0-99 */
    dst,           //* FALSE or TRUE */
    zone: integer; // -11 to +11 */
    GM: TIMEDATE;
 end;
 pPChar=^PChar;

(******************************************************************************)
{Definitions was made automatically}
(******************************************************************************)

type
 PBool     = ^WordBool;
 PBoolean  = ^Boolean;
 PShortInt = ^ShortInt;
 PInteger  = ^Integer;
 PSingle   = ^Single;
 PDouble   = ^Double;

 HGlobal                 =  THandle;
 PRGBTriple              = ^TRGBTriple;
 PRGBQuad                = ^TRGBQuad;
 PMenuItemTemplateHeader = ^TMenuItemTemplateHeader;
 PMenuItemTemplate       = ^TMenuItemTemplate;
 PMultiKeyHelp           = ^TMultiKeyHelp;

const
 ITEM_SIGN = $0001; { This field will be signed if requested }
 ITEM_SEAL = $0002; { This field will be encrypted if requested }
 ITEM_SUMMARY = $0004; { This field can be referenced in a formula }
 ITEM_READWRITERS = $0020; { This field identifies subset of users that have read/write access }
 ITEM_NAMES = $0040; { This field contains user/group names }
 ITEM_PLACEHOLDER = $0100; { Simply add this item to 'item name table', but do not store }
 ITEM_PROTECTED = $0200; { This field cannot be modified except by 'owner' }
 ITEM_READERS = $0400; { This field identifies subset of users that have read access }

{ If the following is ORed in with a note class, the resultant note ID }
{may be passed into NSFNoteOpen and may be treated as though you first }
{did an NSFGetSpecialNoteID followed by an NSFNoteOpen, all in a single }
{transaction. }

 NOTE_ID_SPECIAL = $FFFF0000;

{ Note Classifications }
{ If NOTE_CLASS_DEFAULT is ORed with another note class, it is in }
{essence specifying that this is the default item in this class. There }
{should only be one DEFAULT note of each class that is ever updated, }
{although nothing in the NSF machinery prevents the caller from adding }
{more than one. The file header contains a table of the note IDs of }
{the default notes (for efficient access to them). Whenever a note }
{is updated that has the default bit set, the reference in the file }
{header is updated to reflect that fact. }
{WARNING: NOTE_CLASS_DOCUMENT CANNOT have a "default". This is precluded }
{by code in NSFNoteOpen to make it fast for data notes. }
{ }

   NOTE_CLASS_DOCUMENT = $0001; { document note }
   NOTE_CLASS_DATA = NOTE_CLASS_DOCUMENT; { old name for document note }
   NOTE_CLASS_INFO = $0002; { notefile info (help-about) note }
   NOTE_CLASS_FORM = $0004; { form note }
   NOTE_CLASS_VIEW = $0008; { view note }
   NOTE_CLASS_ICON = $0010; { icon note }
   NOTE_CLASS_DESIGN = $0020; { design note collection }
   NOTE_CLASS_ACL = $0040; { acl note }
   NOTE_CLASS_HELP_INDEX = $0080; { Notes product help index note }
   NOTE_CLASS_HELP = $0100; { designer's help note }
   NOTE_CLASS_FILTER = $0200; { filter note }
   NOTE_CLASS_FIELD = $0400; { field note }
   NOTE_CLASS_REPLFORMULA = $0800; { replication formula }
   NOTE_CLASS_PRIVATE = $1000; { Private design note, use $PrivateDesign view to locate/classify }


   NOTE_CLASS_DEFAULT = $8000; { MODIFIER - default version of each }

   NOTE_CLASS_NOTIFYDELETION = NOTE_CLASS_DEFAULT; { see SEARCH_NOTIFYDELETIONS }
   NOTE_CLASS_ALL = $7fff; { all note types }
   NOTE_CLASS_ALLNONDATA = $7ffe; { all non-data notes }
   NOTE_CLASS_NONE = $0000; { no notes }


{ Define symbol for those note classes that allow only one such in a file }

         NOTE_CLASS_SINGLE_INSTANCE = (NOTE_CLASS_DESIGN or
                                    NOTE_CLASS_ACL or
                                    NOTE_CLASS_INFO or
                                    NOTE_CLASS_ICON or
                                    NOTE_CLASS_HELP_INDEX or 0);
{ Note flag definitions }

         NOTE_SIGNED = $0001; { signed }
         NOTE_ENCRYPTED = $0002; { encrypted }

{ Open Flag Definitions. These flags are passed to NSFNoteOpen. }

         OPEN_SUMMARY = $0001; { open only summary info }
         OPEN_NOVERIFYDEFAULT = $0002; { don't bother verifying default bit }
         OPEN_EXPAND = $0004; { expand data while opening }
         OPEN_NOOBJECTS = $0008; { don't include any objects }
         OPEN_SHARE = $0020; { open in a 'shared' memory mode }
         OPEN_MARK_READ = $0100; { Mark unread if unread list is currently associated }
         OPEN_ABSTRACT = $0200; { Only open an abstract of large documents }
         OPEN_RESPONSE_ID_TABLE = $1000; { Return response ID table }

{ Update Flag Definitions. These flags are passed to NSFNoteUpdate and }
{NSFNoteDelete. See also NOTEID_xxx special definitions in nsfdata.h. }

         UPDATE_FORCE = $0001; { update even if ERR_CONFLICT }
         UPDATE_NAME_KEY_WARNING = $0002; { give error if new field name defined }
         UPDATE_NOCOMMIT = $0004; { do NOT do a database commit after update }
         UPDATE_NOREVISION = $0100; { do NOT maintain revision history }
         UPDATE_NOSTUB = $0200; { update body but leave no trace of note in file if deleted }
         UPDATE_INCREMENTAL = $4000; { Compute incremental note info }
         UPDATE_DELETED = $8000; { update body DELETED }

         UPDATE_DUPLICATES = 0; { Obsolete; but in SDK }

{ Conflict Handler defines }
         CONFLICT_ACTION_MERGE = 1;
         CONFLICT_ACTION_HANDLED = 2;

         UPDATE_SHARE_SECOND = $00200000; { Split the second update of this note with the object store }
         UPDATE_SHARE_OBJECTS = $00400000; { Share objects only, not non-summary items, with the object store }

{ Structure returned from NSFNoteDecrypt which can be used to decrypt }
{file attachment objects, which are not decrypted until necessary. }


type
  Encryption_Key = packed record
    Byte1: BYTE;
    Word1: Word;
    Text: Array[0..16-1] of BYTE;
  end;
  PEncryption_Key = ^Encryption_Key;


{ Flags returned (beginning in V3) in the _NOTE_FLAGS }

const NOTE_FLAG_READONLY = $0001; { TRUE if document cannot be updated }
         NOTE_FLAG_ABSTRACTED = $0002; { missing some data }
         NOTE_FLAG_LINKED = $0020; { Note contains linked items or linked objects }

{ Note structure member IDs for NSFNoteGet&SetInfo. }

         _NOTE_DB = 0; { IDs for NSFNoteGet&SetInfo }
         _NOTE_ID = 1; { (When adding new values, see the }
         _NOTE_OID = 2; { table in NTINFO.C }
         _NOTE_CLASS = 3;
         _NOTE_MODIFIED = 4;
         _NOTE_PRIVILEGES = 5; { For pre-V3 compatibility. Should use $Readers item }
         _NOTE_FLAGS = 7;
         _NOTE_ACCESSED = 8;
         _NOTE_PARENT_NOTEID = 10; { For response hierarchy }
         _NOTE_RESPONSE_COUNT = 11; { For response hierarchy }
         _NOTE_RESPONSES = 12; { For response hierarchy }
         _NOTE_ADDED_TO_FILE = 13; { For AddedToFile time }
         _NOTE_OBJSTORE_DB = 14; { DBHANDLE of object store used by linked items }


{ EncryptFlags used in NSFNoteCopyAndEncrypt }

         ENCRYPT_WITH_USER_PUBLIC_KEY = $0001;

{ DecryptFlags used in NSFNoteDecrypt }

         DECRYPT_ATTACHMENTS_IN_PLACE = $0001;

{ Flags used for NSFNoteExtractFileExt }

         NTEXT_RESONLY = $0001; { If a Mac attachment, extract resource fork only. }
         NTEXT_FTYPE_MASK = $0070; { File type mask }
         NTEXT_FTYPE_FLAT = $0000; { Normal one fork file }
         NTEXT_FTYPE_MACBIN = $0010; { MacBinaryII file }

{ Possible return values from the callback routine specified in }
{NSFNoteComputeWithForm() }

         CWF_ABORT = 1;
         CWF_NEXT_FIELD = 2;
         CWF_RECHECK_FIELD = 3;
         CWF_CONVERT = 4;

{ Possible validation phases for NSFNoteComputeWithForm() }

         CWF_DV_FORMULA = 1;
         CWF_IT_FORMULA = 2;
         CWF_IV_FORMULA = 3;
         CWF_COMPUTED_FORMULA = 4;
         CWF_DATATYPE_CONVERSION = 5;

{ Function pointer type for NSFNoteComputeWithForm() callback }

(******************************************************************************)
{Hand-made definition}
(******************************************************************************)

type
  NoteHandle = Handle;
  PNoteHandle = ^PNoteHandle;
  PNUMBER    = ^Number;
  DbHandle   = Handle;
  PDbHandle  = ^DBHandle;
  HModule    = Handle;
  PHModule   = ^HModule;
  OriginatorId = packed record
    dFile: DBID; { Unique (random) number (Even though this field is called "File", }
                 { it doesn't have anything to do with the file!) }
    Note: TIMEDATE; { Original Note Creation time/date  (THE ABOVE 2 FIELDS MUST BE FIRST - UNID }
                      { COPIED FROM HERE ASSUMED AT OFFSET 0) }
    Sequence: LongInt; { LOW ORDER: sequence number, 1 for first version }
                        { HIGH ORDER WORD: flags, as above }
    SequenceTime: TIMEDATE; { time/date when sequence number was bumped }
   end;
   LicenseId = packed record
     ID     : Array[0..5-1] of BYTE;{license number }
     Product: BYTE; {product code, mfgr-specific }
     Check  : Array[0..2-1] of BYTE; {validity check field, mfgr-specific }
   end {_7};
   PLicenseId = ^LicenseId;
const
  NullHandle: Handle = 0;
function TimeGMToLocal (var aTime: TimeStruct): bool; stdcall; far;
function TimeGMToLocalZone (var aTime: TimeStruct): bool; stdcall; far;
function TimeLocalToGM (var aTime: TimeStruct): bool; stdcall; far;

function NotesInitIni(pConfigFileName: PChar): STATUS; stdcall; far;

function NotesInit: STATUS; stdcall; far;

function NotesInitExtended(argc: Integer;
                           argv: PPChar): STATUS; stdcall; far;

procedure NotesTerm; stdcall; far;

procedure NotesInitModule(rethModule: PHMODULE;
                          rethInstance: PHMODULE;
                          rethPrevInstance: PHMODULE); stdcall; far;
{$IFDEF NLM}
// type  = VOID EXPORTED_LIBRARY_PROC(VOID);

function NotesLibraryMain(argc: Integer;
                          argv: PPChar;
                          initproc: EXPORTED_LIBRARY_PROC): STATUS; stdcall; far;
{$ENDIF /* NLM }

function NotesInitThread: STATUS; stdcall; far;

procedure NotesTermThread; stdcall; far;


(******************************************************************************)
{NsfNOte.H}
(******************************************************************************)

type
  CWF_ERROR_PROC  = function (const PCDFIELD: pointer; PHASE:word; ERROR: STATUS; ERRORTEXT: HANDLE;  WERRORTEXTSIZE:WORD; CTX: pointer): word; stdcall;
  NSFItemScanProc = function (Spare, ItemFlags: word;
                              Name: PChar;
                              NameLength: word;
                              Value: pointer;
                              ValueLength: dword;
                              RoutineParameter: pointer): Status;  stdcall;

const CWF_CONTINUE_ON_ERROR = $0001;

{ function templates }

function NSFItemAppend(hNote: NOTEHANDLE;
                       ItemFlags: Word;
                       Name: PChar;
                       NameLength: Word;
                       DataType: Word;
                       Value: Pointer;
                       ValueLength: LongInt): STATUS; stdcall; far;

function NSFItemAppendByBLOCKID(hNote: NOTEHANDLE;
                                ItemFlags: Word;
                                Name: PChar;
                                NameLength: Word;
                                bhValue: BLOCKID;
                                ValueLength: LongInt;
                                retbhItem: PBLOCKID): STATUS; stdcall; far;


function NSFItemAppendObject(hNote: NOTEHANDLE;
                             ItemFlags: Word;
                             Name: PChar;
                             NameLength: Word;
                             bhValue: BLOCKID;
                             ValueLength: LongInt;
                             fDealloc: Bool): STATUS; stdcall; far;

function NSFItemDelete(hNote: NOTEHANDLE;
                       Name: PChar;
                       NameLength: Word): STATUS; stdcall; far;

function NSFItemDeleteByBLOCKID(hNote: NOTEHANDLE;
                                bhItem: BLOCKID): STATUS; stdcall; far;


function NSFItemRealloc(bhItem: BLOCKID;
                        bhValue: PBLOCKID;
                        ValueLength: LongInt): STATUS; stdcall; far;


function NSFItemCopy(hNote: NOTEHANDLE;
                     bhItem: BLOCKID): STATUS; stdcall; far;

function NSFItemInfo(hNote: NOTEHANDLE;
                     Name: PChar;
                     NameLength: Word;
                     retbhItem: PBLOCKID;
                     retDataType: PWord;
                     retbhValue: PBLOCKID;
                     retValueLength: PLongInt): STATUS; stdcall; far;

// const NSFItemIsPresent(hNote, = Name, NameLength) \;

function NSFItemIsPresent (hNote: NoteHandle; Name: pchar; NameLength: word): boolean;

function NSFItemInfoNext(hNote: NOTEHANDLE;
                         PrevItem: BLOCKID;
                         Name: PChar;
                         NameLength: Word;
                         retbhItem: PBLOCKID;
                         retDataType: PWord;
                         retbhValue: PBLOCKID;
                         retValueLength: PDWORD): STATUS; stdcall; far;


procedure NSFItemQuery(hNote: NOTEHANDLE;
                       bhItem: BLOCKID;
                       retItemName: PChar;
                       ItemNameBufferLength: Word;
                       retItemNameLength: PWord;
                       retItemFlags: PWord;
                       retDataType: PWord;
                       retbhValue: PBLOCKID;
                       retValueLength: PLongInt); stdcall; far;


function NSFItemGetText(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        retBuffer: PChar;
                        BufferLength: Word): Word; stdcall; far;

function NSFItemGetModifiedTime (hNote: NOTEHANDLE;
                                 ItemName: PChar;
                                 ItemNameLength: word;
                                 Flags: dword;
                                 retTime: PTIMEDATE): Status; stdcall; far;

function NSFItemGetTime(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        retTime: PTIMEDATE): Bool; stdcall; far;

function NSFItemGetNumber(hNote: NOTEHANDLE;
                          ItemName: PChar;
                          retNumber: PNUMBER): Bool; stdcall; far;

function NSFItemGetLong(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        DefaultNumber: LongInt): LongInt; stdcall; far;


function NSFItemSetText(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        Text: PChar;
                        TextLength: Word): STATUS; stdcall; far;

function NSFItemSetTextSummary(hNote: NOTEHANDLE;
                               ItemName: PChar;
                               Text: PChar;
                               TextLength: Word;
                               Summary: Bool): STATUS; stdcall; far;

function NSFItemSetTime(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        Time: PTIMEDATE): STATUS; stdcall; far;

function NSFItemSetNumber(hNote: NOTEHANDLE;
                          ItemName: PChar;
                          Number: PNUMBER): STATUS; stdcall; far;


function NSFItemGetTextListEntries(hNote: NOTEHANDLE;
                                   ItemName: PChar): Word; stdcall; far;

function NSFItemGetTextListEntry(hNote: NOTEHANDLE;
                                 ItemName: PChar;
                                 EntryPos: Word;
                                 retBuffer: PChar;
                                 BufferLength: Word): Word; stdcall; far;

function NSFItemCreateTextList(hNote: NOTEHANDLE;
                               ItemName: PChar;
                               Text: PChar;
                               TextLength: Word): STATUS; stdcall; far;

function NSFItemAppendTextList(hNote: NOTEHANDLE;
                               ItemName: PChar;
                               Text: PChar;
                               TextLength: Word;
                               fAllowDuplicates: Bool): STATUS; stdcall; far;


function NSFItemTextEqual(hNote: NOTEHANDLE;
                          ItemName: PChar;
                          Text: PChar;
                          TextLength: Word;
                          fCaseSensitive: Bool): Bool; stdcall; far;

function NSFItemTimeCompare(hNote: NOTEHANDLE;
                            ItemName: PChar;
                            Time: PTIMEDATE;
                            retVal: PInteger): Bool; stdcall; far;

function NSFItemLongCompare(hNote: NOTEHANDLE;
                            ItemName: PChar;
                            Value: LongInt;
                            retVal: PInteger): Bool; stdcall; far;


function NSFItemConvertValueToText(DataType: Word;
                                   bhValue: BLOCKID;
                                   ValueLength: LongInt;
                                   retBuffer: PChar;
                                   BufferLength: Word;
                                   SepChar: Char): Word; stdcall; far;

function NSFItemConvertToText(hNote: NOTEHANDLE;
                              ItemName: PChar;
                              retBuffer: PChar;
                              BufferLength: Word;
                              SepChar: Char): Word; stdcall; far;


function NSFGetSummaryValue(SummaryBuffer: Pointer;
                            Name: PChar;
                            retValue: PChar;
                            ValueBufferLength: Word): Bool; stdcall; far;

function NSFLocateSummaryValue(SummaryBuffer: Pointer;
                               Name: PChar;
                               retValuePointer: Pointer;
                               retValueLength: PWord;
                               retDataType: PWord): Bool; stdcall; far;
function NSFItemScan(hNote: NOTEHANDLE;
                     ActionRoutine: NSFITEMSCANPROC;
                     RoutineParameter: Pointer): STATUS; stdcall; far;


procedure NSFNoteGetInfo(hNote: NOTEHANDLE;
                         wType: Word;
                         Value: Pointer); stdcall; far;

procedure NSFNoteSetInfo(hNote: NOTEHANDLE;
                         wType: Word;
                         Value: Pointer); stdcall; far;
function NSFNoteClose(hNote: NOTEHANDLE): STATUS; stdcall; far;
function NSFNoteCreate(hDB: DBHANDLE;
                       rethNote: PNOTEHANDLE): STATUS; stdcall; far;
function NSFNoteDelete(hDB: DBHANDLE;
                       NoteID: NOTEID;
                       UpdateFlags: Word): STATUS; stdcall; far;
function NSFNoteOpen(hDB: DBHANDLE; NoteID: NOTEID; OpenFlags: Word; rethNote: PNOTEHANDLE): STATUS; stdcall; far;
function NSFNoteOpenByUNID(hDB: THandle; pUNID: PUNID; flags: Word; rtn: PHandle): STATUS; stdcall; far;
function NSFNoteUpdate(hNote: NOTEHANDLE; UpdateFlags: Word): STATUS; stdcall; far;
function NSFNoteUpdateExtended(hNote: NOTEHANDLE; UpdateFlags: LongInt): STATUS; stdcall; far;
function NSFNoteComputeWithForm(hNote: NOTEHANDLE; hFormNote: NOTEHANDLE; dwFlags: LongInt;
                                ErrorRoutine: CWF_ERROR_PROC;
                                CallersContext: Pointer): STATUS; stdcall; far;
function NSFNoteAttachFile(hNOTE: NOTEHANDLE;
                           ItemName: PChar;
                           ItemNameLength: Word;
                           PathName: PChar;
                           OriginalPathName: PChar;
                           Encoding: Word): STATUS; stdcall; far;
function NSFNoteExtractFile(hNote: NOTEHANDLE;
                            bhItem: BLOCKID;
                            FileName: PChar;
                            DecryptionKey: PENCRYPTION_KEY): STATUS; stdcall; far;
function NSFNoteExtractFileExt(hNote: NOTEHANDLE;
                               bhItem: BLOCKID;
                               FileName: PChar;
                               DecryptionKey: PENCRYPTION_KEY;
                               wFlags: Word): STATUS; stdcall; far;
function NSFNoteDetachFile(hNote: NOTEHANDLE;
                           bhItem: BLOCKID): STATUS; stdcall; far;
function NSFNoteHasObjects(hNote: NOTEHANDLE;
                           bhFirstObjectItem: PBLOCKID): Bool; stdcall; far;
function NSFNoteGetAuthor(hNote: NOTEHANDLE;
                          retName: PChar;
                          retNameLength: PWord;
                          retIsItMe: PBool): STATUS; stdcall; far;
function NSFNoteCopy(hSrcNote: NOTEHANDLE;
                     rethDstNote: PNOTEHANDLE): STATUS; stdcall; far;
function NSFNoteSignExt(hNote: NOTEHANDLE;
                        SignatureItemName: PChar;
                        ItemCount: Word;
                        hItemIDs: THandle): STATUS; stdcall; far;

function NSFNoteSign(hNote: NOTEHANDLE): STATUS; stdcall; far;
function NSFNoteVerifySignature(hNote: NOTEHANDLE;
                                Reserved: PChar;
                                retWhenSigned: PTIMEDATE;
                                retSigner: PChar;
                                retCertifier: PChar): STATUS; stdcall; far;
function NSFVerifyFileObjSignature(hDB: DBHANDLE;
                                   bhItem: BLOCKID): STATUS; stdcall; far;
function NSFNoteUnsign(hNote: NOTEHANDLE): STATUS; stdcall; far;
function NSFNoteCopyAndEncrypt(hSrcNote: NOTEHANDLE;
                               EncryptFlags: Word;
                               rethDstNote: PNOTEHANDLE): STATUS; stdcall; far;
function NSFNoteDecrypt(hNote: NOTEHANDLE;
                        DecryptFlags: Word;
                        retKeyForAttachments: PENCRYPTION_KEY): STATUS; stdcall; far;
function NSFNoteIsSignedOrSealed(hNote: NOTEHANDLE;
                                 retfSigned: PBool;
                                 retfSealed: PBool): Bool; stdcall; far;
function NSFNoteCheck (hNote: THandle): STATUS; stdcall; far;

{ External (text) link routines }


const
  LINKFLAG_ADD_TEMPORARY = $00000002;
  LINKFLAG_NO_REPL_SEARCH = $00000004;

function NSFNoteLinkFromText(hLinkText: THandle;
                             LinkTextLength: Word;
                             NoteLink: PNOTELINK;
                             ServerHint: PChar;
                             LinkText: PChar;
                             MaxLinkText: Word;
                             retFlags: PLongInt): STATUS; stdcall; far;

function NSFNoteLinkToText(Title: PChar;
                           NoteLink: PNOTELINK;
                           ServerHint: PChar;
                           LinkText: PChar;
                           phLinkText: PHandle;
                           pLinkTextLength: PWord;
                           Flags: LongInt): STATUS; stdcall; far;

function NSFProfileOpen(
  hDB: LHandle;
  ProfileName: pchar;
  ProfileNameLength: word;
  UserName: pchar;
  UserNameLength: word;
  CopyProfile: boolean;
  rethProfileNote: PHandle): STATUS; stdcall; far;

type
  NSFPROFILEENUMPROC = function(
   hDB: LHandle;
   Ctx: pointer;
   ProfileName: pchar;
   ProfileNameLength: word;
   UserName: pchar;
   UserNameLength: word;
   ProfileNoteID: NOTEID): STATUS; stdcall;

function NSFProfileEnum(
  hDB: LHandle;
  ProfileName: pchar;
  ProfileNameLength: word;
  Callback: NSFPROFILEENUMPROC;
  CallbackCtx: pointer;
  Flags: DWORD): STATUS; stdcall; far;

function NSFProfileGetField(
  hDB: LHandle;
  ProfileName: pchar;
  ProfileNameLength: word;
  UserName: pchar;
  UserNameLength: word;
  FieldName: pchar;
  FieldNameLength: word;
  var retDatatype: word;
  var retbhValue: BLOCKID;
  var retValueLength: DWORD): STATUS; stdcall; far;

function NSFProfileUpdate(
  hProfile: LHandle;
  ProfileName: pchar;
  ProfileNameLength: word;
  UserName: pchar;
  UserNameLength: word): STATUS; stdcall; far;

function NSFProfileSetField(
  hDB: LHandle;
  ProfileName: pchar;
  ProfileNameLength: word;
  UserName: pchar;
  UserNameLength: word;
  FieldName: pchar;
  FieldNameLength: word;
  Datatype: word;
  Value: pointer;
  ValueLength: dword): STATUS; stdcall; far;

function NSFProfileDelete(
  hDB: LHandle;
  ProfileName: pchar;
  ProfileNameLength: word;
  UserName: pchar;
  UserNameLength: word): STATUS; stdcall; far;

(******************************************************************************)
{NsfSearch.h}
(******************************************************************************)
{ Note Storage File Search Package Definitions }

{ Search Flag Definitions }

const SEARCH_ALL_VERSIONS = $0001; { Include deleted and non-matching notes in search }
{ (ALWAYS "ON" in partial searches!) }
const SEARCH_SUMMARY = $0002; { TRUE to return summary buffer with each match }
const SEARCH_FILETYPE = $0004; { For directory mode file type filtering }
{ If set, "NoteClassMask" is treated }
{ as a FILE_xxx mask for directory filtering }
const SEARCH_NOTIFYDELETIONS = $0010; { Set NOTE_CLASS_NOTIFYDELETION bit of NoteClass for deleted notes }
const SEARCH_ALLPRIVS = $0040; { return error if we don't have full privileges }
const SEARCH_SESSION_USERNAME = $0400; { Use current session's user name, not server's }
const SEARCH_NOABSTRACTS = $1000; { Filter out 'Truncated' documents }
const SEARCH_DATAONLY_FORMULA = $4000; { Search formula applies only to}
{ This descriptor is embedded in the search queue entry. Note: The }
{information returned in the "summary" field is always returned in }
{machine-independent canonical form. }

{Note: In DIRECTORY searches, the following information is returned }
{in the SEARCH_MATCH structure (build 86 & later only): }

{OriginatorID.File NSF modified time (later of data & non-data modified time) }
{OriginatorID.Note 0 (unused) }
{OriginatorID.SequenceTime NSF's Replica ID (Used by NSFMakeReplicaFormula) }
{ID.Note NSF's Replica ID }
{ID.File NSF's DBID }
{ }

{ SERetFlags values (bit-field) }

const SE_FNOMATCH = $00; { does not match formula (deleted or updated) }
const SE_FMATCH = $01; { matches formula }
const SE_FTRUNCATED = $02; { document truncated }
const SE_FPURGED = $04; { note has been purged. Returned only when SEARCH_INCLUDE_PURGED is used }

{ If recompiling a V3 API application and you used the MatchesFormula field }
{the following code change should be made: }

{For V3: }

{1) if (SearchMatch.MatchesFormula == SE_FMATCH) }
{2) if (SearchMatch.MatchesFormula == SE_FNOMATCH) }
{3) if (SearchMatch.MatchesFormula != SE_FMATCH) is equivalent to 2) }
{4) if (SearchMatch.MatchesFormula != SE_FNOMATCH) is equivalent to 1) }

{For V4 }

{1) if (SearchMatch.SERetFlags & SE_FMATCH) }
{2) if (!(SearchMatch.SERetFlags & SE_FMATCH)) }
{ }


type
  Search_Match = packed record
    ID: GLOBALINSTANCEID; { identity of the note within the file }
    OriginatorID: ORIGINATORID; { identity of the note in the universe }
    NoteClass: Word; { class of the note }
    SERetFlags: BYTE; { MUST check for SE_FMATCH! }
    Privileges: BYTE; { note privileges }
    SummaryLength: Word;
{length of the summary information }
{54 bytes to here }
{now comes an ITEM_TABLE with Summary Info }
end;
pSEARCH_MATCH=^SEARCH_MATCH;

NSFSEARCHPROC=function(EnumRoutineParameter:pointer; SearchMatch: pSEARCH_MATCH;
               SummaryBuffer: pITEM_TABLE):word; stdcall;


function NSFSearch(hDB: DBHANDLE;
                   hFormula: FORMULAHANDLE;
                   ViewTitle: PChar;
                   SearchFlags: Word;
                   NoteClassMask: Word;
                   Since: PTIMEDATE;
                   EnumRoutine: NSFSEARCHPROC;
                   EnumRoutineParameter: Pointer;
                   retUntil: PTIMEDATE): STATUS; stdcall; far;

{ Formula compilation functions }


function NSFFormulaCompile(FormulaName: PChar;
                           FormulaNameLength: Word;
                           FormulaText: PChar;
                           FormulaTextLength: Word;
                           rethFormula: PFORMULAHANDLE;
                           retFormulaLength: PWord;
                           retCompileError: PSTATUS;
                           retCompileErrorLine,
                           retCompileErrorColumn,
                           retCompileErrorOffSet,
                           retCompileErrorLength: pWord): Status; stdcall; far;

function NSFFormulaDecompile(FormulaBuffer: PChar;
                           fSelectionFormula: Boolean;
                           rethFormulaText: PHandle;
                           retFormulaTextLength: PWord): STATUS; stdcall; far;

function NSFFormulaMerge(hSrcFormula: FORMULAHANDLE;
                         hDestFormula: FORMULAHANDLE): STATUS; stdcall; far;

function NSFFormulaSummaryItem(hFormula: FORMULAHANDLE;
                               ItemName: PChar;
                               ItemNameLength: Word): STATUS; stdcall; far;

function NSFFormulaGetSize(hFormula: FORMULAHANDLE;
                           retFormulaLength: PWord): STATUS; stdcall; far;


{ Formula computation (evaluation) functions }

type HCOMPUTE = LHANDLE;
     PHCOMPUTE = ^HCOMPUTE;


function NSFComputeStart(Flags: Word;
                         pCompiledFormula: pointer;
                         rethCompute: PHCOMPUTE): STATUS; stdcall; far;

function NSFComputeStop(hCompute: HCOMPUTE): STATUS; stdcall; far;

function NSFComputeEvaluate(hCompute: HCOMPUTE;
                            hNote: NOTEHANDLE;
                            rethResult: PHandle;
                            retResultLength: PWord;
                            retNoteMatchesFormula: PBool;
                            retNoteShouldBeDeleted: PBool;
                            retNoteModified: PBool): STATUS; stdcall; far;

{ End of Note Storage File Search Package Definitions }



(******************************************************************************)
{Notes Rich Text On-Disk Structure Definitions}
{Record format used in the NSF data type TYPE_COMPOSITE. }
{From EDITODS.H}
(******************************************************************************)
{Paragraph Record - Defines the start of a new paragraph }


type
  BSIG = packed record
    Signature: byte;
    Length: byte;
  end;
  WSIG = packed record
    Signature: word;
    Length: word;
  end;
  SWORD = smallint;
  CDPARAGRAPH = packed record
     Header: BSIG;
  end;

const
  maxfacesize = 32;

{Paragraph Attribute Block Definition Record }
const MAXTABS = 20; { maximum number of stops in tables }
const MAX_STYLE_NAME = 35;
const MAX_STYLE_USERNAME = 128;


type
  CDPABDEFINITION = packed record
    Header: WSIG;
    PABID: Word;{ ID of this PAB }
    JustifyMode: Word;{ paragraph justification type }
    LineSpacing: Word;{ (2 * (Line Spacing - 1)) (0:1,1:1.5,2:2,etc) }
    ParagraphSpacingBefore: Word; { no. of LineSpacing units above paragraph }
    ParagraphSpacingAfter: Word; { no. of LineSpacing units below paragraph }
    LeftMargin: Word; {leftmost margin, twips rel to abs left (16 bits = about 44") }
    RightMargin: Word;
{rightmost margin, twips rel to abs right (16 bits = about 44") }
{Special value "0" means right margin will be placed 1" from right edge of
paper, regardless of paper size.}
    FirstLineLeftMargin: Word; {leftmost margin on first line (16 bits = about 44") }
    Tabs: Word; { number of tab stops in table }
    Tab: Array[0..MAXTABS-1] of SWORD;
{ table of tab stop positions, negative }
{value means decimal tab }
{ (15 bits = about 22") }
    Flags: Word; { paragraph attribute flags }
    TabTypes: LongInt; { 2 bits per tab }
    Spare: Array[0..1-1] of Word;
  end {_2};


{New PAB record for V4 -hide when formula}
  CDPABHIDE = packed record
    Header: WSIG;
    PABID: Word;
    RESERVERD: Array[0..8-1] of BYTE ;
  end {_3};
  {Follows is the actual formula}
  {PAB Reference Record - }
  {This record is output in two situations: First, at the start of every }
  {item of type Composite. Second, at the start of every paragraph. If, }
  {when reading this record during a note READ operation, the paragraph }
  {in question has NO runs in it, this defines the PAB to use for the }
  {paragraph. If the paragraph has runs already in it, ignore this record. }
  CDPABREFERENCE = packed record
    Header: BSIG;
    PABID: Word; {ID number of the PAB being referenced }
end {_4};

const PABFLAG_PAGINATE_BEFORE = $0001;  //* start new page with this par */
const PABFLAG_KEEP_WITH_NEXT  = $0002;  //* don't separate this and next par */
const PABFLAG_KEEP_TOGETHER = $0004;  //* don't split lines in paragraph */
const PABFLAG_PROPAGATE   = $0008;  //* propagate even PAGINATE_BEFORE and KEEP_WITH_NEXT */
const PABFLAG_HIDE_RO     = $0010;  //* hide paragraph in R/O mode */
const PABFLAG_HIDE_RW     = $0020;  //* hide paragraph in R/W mode */
const PABFLAG_HIDE_PR     = $0040;  //* hide paragraph when printing */
const PABFLAG_DISPLAY_RM    = $0080;  //* honor right margin when displaying to a window */
const PABFLAG_HIDE_CO     = $0200;  //* hide paragraph when copying/forwarding */
const PABFLAG_BULLET      = $0400;  //* display paragraph with bullet */
const PABFLAG_HIDE_IF     = $0800;  //*  use the hide when formula even if there is one.    */
const PABFLAG_NUMBEREDLIST  = $1000;  //* display paragraph with number */
const PABFLAG_HIDE_PV     = $2000;  //* hide paragraph when previewing*/
const PABFLAG_HIDE_PVE    = $4000;  //* hide paragraph when editing in the preview pane.    */
const PABFLAG_HIDE_NOTES    = $8000;  //* hide paragraph from Notes clients */

{This record is similar to a pab reference record but applies }
{only to the pab's hide when formula which is new for V4 }
{so we make a new type that v3 can safely ignore. }
{}


type
  CDPABFORMULAREF = packed record
            Header: BSIG;
            SourcePABID: Word;{ID number of the source PAB ontaining the formula. }
            DestPABID: Word; {ID number of the dest PAB }
end {_5};


{ Style Name Record }

const STYLE_FLAG_FONTID = $00000001;
const STYLE_FLAG_INCYCLE = $00000002;
const STYLE_FLAG_PERMANENT = $00000004;

type
  CDSTYLENAME = packed record
    Header: BSIG;
    Flags: LongInt;{ Currently unused, but reserve some flags }
    PABID: Word; { ID number of the PAB being named }
    StyleName: Array[0..MAX_STYLE_NAME+1-1] of Char;{ The style name. }
            { If STYLE_FLAG_FONTID, a FONTID follows this structure. }
end {_6};

{ Begin Table Record - }

{This record specifies the beginning of a table. It contains interesting }
{information about the format and size of the table. }


type
  CDTABLEBEGIN = packed record
            Header: BSIG;
            LeftMargin: Word;{ TWIPS }
            HorizInterCellSpace: Word; { TWIPS }
            VertInterCellSpace: Word;  { TWIPS }
{ NOTE! all items below this comment are NOT guaranteed to have been zeroed if }
{created in V2; all items are zeroed before use in V4 }
            V4HorizInterCellSpace: Word;{TWIPS -- this field was spare in v3 }
            V4VertInterCellSpace: Word; {TWIPS -- this field was spare in v3 }
            Flags: Word; { Flags (CDTABLE_xxx) }
          end {_7};

const CDTABLE_AUTO_CELL_WIDTH = $0001; { True if automatic cell width calculation }
const CDTABLE_V4_BORDERS = $0002; { True if the table was created in v4 }
const CDTABLE_3D_BORDER_EMBOSS = $0004;
const CDTABLE_3D_BORDER_EXTRUDE = $0008;


type CDTABLECELL = packed record
            Header: BSIG;
            Row: BYTE;{ Row number (0 based) }
            Column: BYTE; { Column number (0 based) }
            LeftMargin: Word;{ Twips }
            RightMargin: Word; { Twips }
            FractionalWidth: Word; { 20" (in twips) * CellWidth / TableWidth Used only if AutoCellWidth is }
                                    {specified in the TABLEBEGIN. }
            Border: BYTE;{ 4 cell borders, each 2 bits wide }
                          { (see shift and mask CDTC_xxx values) }
                          { Value of each cell border is one of }
                          { TABLE_BORDER_xxx. }
            Flags: BYTE;
            v42Border: WORD;
            RowSpan: BYTE;
            ColumnSpan: BYTE;
            BackgroundColor: WORD;
          end {_8};

const CDTC_S_Left = 0;
const CDTC_M_Left = $0003;
const CDTC_S_Right = 2;
const CDTC_M_Right = $000c;
const CDTC_S_Top = 4;
const CDTC_M_Top = $0030;
const CDTC_S_Bottom = 6;
const CDTC_M_Bottom = $00c0;
const TABLE_BORDER_NONE = 0;
const TABLE_BORDER_SINGLE = 1;
const TABLE_BORDER_DOUBLE = 2;

const CDTC_S_V42_Left   = 0;
const CDTC_M_V42_Left   = $000f;
const CDTC_S_V42_Right  = 4;
const CDTC_M_V42_Right  = $00f0;
const CDTC_S_V42_Top    = 8;
const CDTC_M_V42_Top    = $0f00;
const CDTC_S_V42_Bottom = 12;
const CDTC_M_V42_Bottom  = $f000;

const CDTABLECELL_USE_BKGCOLOR  = $01;  //* True if background color */
const CDTABLECELL_USE_V42BORDERS  = $02;  //* True if version 4.2 or after */
const CDTABLECELL_INVISIBLEH    = $04;  //* True if cell is not spanned */
const CDTABLECELL_INVISIBLEV    = $08;  //* True if cell is not spanned */

type
  Fontid = dword;
  CDTABLEEND = packed record
    Header: BSIG;
    SpareWORD: Word;
    SpareFlags: Word;
  end {_9};
  NFMT = packed record
    Digits,
    Format,
    Attributes,
    Unused: byte;
  end;
  LSIG = packed record
    Signature: word;
    Length: dword;
  end;
  RectSize = packed record
    width,
    height:word;
  end;

// CDFIELD - Field Reference Record, used in forms ($BODY) to define a field. }
type
  CDFIELD = packed record
   Header: WSIG;
   Flags: Word; { Field Flags (see Fxxx flags below) }
   DataType: Word;{ Alleged NSF Data type }
   ListDelim: Word;{ List Delimiters (LDELIM_xxx and LDDELIM_xxx) }
   NumberFormat: NFMT; { Number format, if applicable }
   TimeFormat: TFMT; { Time format, if applicable }
   FontID: FONTID; { displayed font }
   DVLength: Word; { Default Value Formula }
   ITLength: Word; { Input Translation Formula }
   Unused1: Word; {Unused }
   IVLength: Word; { Input Validity Check Formula }
   NameLength: Word; { NSF Item Name }
   DescLength: Word; { Description of the item }
   TextValueLength: Word; {(Text List) List of valid text values. Now comes the variable part of the struct...}
 end {_10};

{ CDFIELD List Delimeters (ListDelim) }

const LDELIM_SPACE = $0001; { low three nibbles contain delim flags }
const LDELIM_COMMA = $0002;
const LDELIM_SEMICOLON = $0004;
const LDELIM_NEWLINE = $0008;
const LDELIM_BLANKLINE = $0010;
const LD_MASK = $0fff;

const LDDELIM_SPACE = $1000; { high nibble contains the display type }
const LDDELIM_COMMA = $2000;
const LDDELIM_SEMICOLON = $3000;
const LDDELIM_NEWLINE = $4000;
const LDDELIM_BLANKLINE = $5000;
const LDD_MASK = $f000;

{ CDFIELD Flags Definitions }

const V3SPARESTOCLEAR = $0075; { Clear these if FOCLEARSPARES is TRUE }

const FREADWRITERS = $0001; { Field contains read/writers }
const FEDITABLE = $0002; { Field is editable, not read only }
const FNAMES = $0004; { Field contains distinguished names }
const FSTOREDV = $0008; { Store DV, even if not spec'ed by user }
const FREADERS = $0010; { Field contains document readers }
const FSECTION = $0020; { Field contains a section }
const FSPARE3 = $0040; { can be assumed to be clear in memory, V3 & later }
const FV3FAB = $0080; { IF CLEAR, CLEAR AS ABOVE }
const FCOMPUTED = $0100; { Field is a computed field }
const FKEYWORDS = $0200; { Field is a keywords field }
const FPROTECTED = $0400; { Field is protected }
const FREFERENCE = $0800; { Field name is simply a reference to a shared field note }
const FSIGN = $1000; { sign field }
const FSEAL = $2000; { seal field }
const FKEYWORDS_UI_STANDARD = $0000; { standard UI }
const FKEYWORDS_UI_CHECKBOX = $4000; { checkbox UI }
const FKEYWORDS_UI_RADIOBUTTON = $8000; { radiobutton UI }
const FKEYWORDS_UI_ALLOW_NEW = $c000; { allow doc editor to add new values }

{ CDEXTFIELD - Extended Field Reference Record, used in forms ($BODY) to define a field. }


type
  CDEXTFIELD = packed record
     Header: WSIG;
     Flags1: LongInt;{ Field Flags (see FEXT_xxx flags below) }
     wFlags2: Word;
     Spare: Array[0..2-1] of BYTE;
     EntryHelper: Word;{ Field entry helper type (see FIELD_HELPER_XXX below) }
     EntryDBNameLen: Word;{ Entry helper DB name length }
     EntryViewNameLen: Word;{ Entry helper View name length }
     EntryColumnNumber: Word;{ Entry helper column number }
{ Now comes the variable part of the struct... }
end {_11};

{ Flags for CDEXTFIELD Flags1. Note that the low word in Flags1 is not used. }

const FEXT_LOOKUP_EACHCHAR = $00010000; { lookup name as each char typed }
const FEXT_KWSELRECALC = $00020000; { recalc on new keyword selection }
const FEXT_KWHINKYMINKY = $00040000; { suppress showing field hinky minky }
const FEXT_AFTERVALIDATION = $00080000; { recalc after validation }
const FEXT_ACCEPT_CARET = $00100000; { the first field with this bit set will accept the caret }
{ These bits are in use by the 0x02000000L }
{column value. The result of 0x04000000L }
{the shifted bits is (cols - 1) 0x08000000L }
const FEXT_KEYWORD_COLS_SHIFT = 25;
const FEXT_KEYWORD_COLS_MASK = $0E000000;
const FEXT_KEYWORD_FRAME_3D = $00000000;
const FEXT_KEYWORD_FRAME_STANDARD = $10000000;
const FEXT_KEYWORD_FRAME_NONE = $20000000;
const FEXT_KEYWORD_FRAME_MASK = $30000000;
const FEXT_KEYWORD_FRAME_SHIFT = 28;
const FEXT_KEYWORDS_UI_COMBO = $40000000;
const FEXT_KEYWORDS_UI_LIST = $80000000;

{ The following identifiers indicate the type of helper in use by the }
{Keyword and the Name helper/pickers }
{ these define the VarDataFlags signifying variable length data following struct }
const CDEXTFIELD_KEYWORDHELPER = $0001;
const CDEXTFIELD_NAMEHELPER = $0002;
const FIELD_HELPER_NONE = 0;
const FIELD_HELPER_ADDRDLG = 1;
const FIELD_HELPER_ACLDLG = 2;
const FIELD_HELPER_VIEWDLG = 3;

{ CDTEXT - 8-bit text string record }


type CDTEXT = packed record
             Header: WSIG;{ Tag and length }
             FontID: FONTID;{ Font ID  The 8-bit text string follows... }
           end {_13};
{CDLINK2 - Link record }


type CDLINK2 = packed record
             Header: WSIG;
             LinkID: Word;{Index into array in $LINKS/$FORMLINKS field of this document }
            { Now comes the display comment... }
           end {_14};


{ CDLINKEXPORT - This record is used in the case of exporting }
{a note to the clipboard, where the NSF item describing the links }
{cannot be generated. }



type CDLINKEXPORT2 = packed record
             Header: WSIG;
             NoteLink: NOTELINK;
{ Now comes the display comment... }
           end {_15};


{ CDKEYWORD - Keyword Record }

const CDKEYWORD_RADIO = $0001;
{ These bits are in use by the 0x0002 }
{column value. The result of 0x0004 }
{the shifted bits is (cols - 1) 0x0008 }
const CDKEYWORD_COLS_SHIFT = 1;
const CDKEYWORD_COLS_MASK = $000E;
const CDKEYWORD_FRAME_3D = $0000;
const CDKEYWORD_FRAME_STANDARD = $0010;
const CDKEYWORD_FRAME_NONE = $0020;
const CDKEYWORD_FRAME_MASK = $0030;
const CDKEYWORD_FRAME_SHIFT = 4;


type
  CDKEYWORD = packed record
     Header: WSIG;  {Tag and length }
     FontID: FONTID;{ Font ID }
     Keywords: Word;{ number of keywords }
     Flags: Word;   {char OnOff[]; array of '1' or '0' indicating state
                    char TextValues[]; packed buffer of keyword text, fab->pTextValues format}
end;

{ Here is a description of Notes bitmap encoding. }
{}
{COLOR FORMATS: }
{}
{Notes displays 3 types of bitmaps: monochrome, color and grey scale. }
{All monochome bitmaps are one bit per pixel. Color bitmaps can be either 8 bits per Pel (color }
{mapped) or 16 bits per pel "quasi true" color. Grey scale bitmaps }
{are simply treated as "color" bitmaps, using the 8 bits per Pel format }
{with a color table whose RGB tuples range from [0,0,0] through }
{[255,255,255]. }
{}
{GEOMETRY: }
{}
{All bitmaps are single plane encoding. 8 bit color/grey scale must have }
{color tables provided. Monochome bitmaps and 16 bit "quasi true" color }
{bitmaps don't need a color table. }
{}
{RASTER LINE ENCODING: }
{}
{For those programmers using Notes API, raster lines are encoded using }
{a simple run-length encoding format, where each raster line of the }
{bitmap is encoded separately (i.e. run length won't exceed length of }
{a raster line). Also, each raster line is NOT padded to any particular }
{boundary; each scanline ends on the byte boundary which is defined by the }
{width of the bitmap. The following section describes how each raster line }
{is encoded using a simple run-length encoding scheme. }
{}
{}
{Notes bitmap compression scheme description }
{------------------------------------------- }
{We have devised a scheme which does a good job compressing }
{monochrome, color (both mapped and RGB) and gray scale }
{bitmaps, and a secondary encoding to allow "raw" uncompressed }
{scanlines for those scanlines which actually "expand" by using }
{the run-length scheme (this happens in dithered images). }
{In the run length encoding scheme, there an escape codes followed by either }
{a run length byte which is then followed by the }
{byte(s) to repeat. Note that the byte(s) to repeat may be either }
{one or two bytes depending on the color format. For monochrome, }
{8 bit color and 8 bit grey scale, use one byte. For 16 bit color }
{the PEL is two bytes long, so the two byte quantity is repeated. }
{}
{MSB<---------->LSB }
{+-----------------------------------+ }
{(1) | 1 1 c c c c c c | r r r r r r r r | }
{+-----------------------------------+ }
{cccccc = six bit repeat count }
{rrrrrrrr = PELS to repeat }
{}
{}
{In the following non-compressed encoding, the escape code is }
{followed by a 6 bit repeat count of raw PELs (one or two byte). }
{}
{}
{MSB<---------->LSB }
{+------------------------------------------------------ }
{(2) | 0 0 c c c c c c | r r r r r r r r |[r r r r r r r r]|... }
{+------------------------------------------------------ }
{cccccc = 6 bit repeat count }
{r[cccccc] = 1 or more raw uncompressed PELS }
{}
{}


{ A color table (used in CDBITMAPHEADER) is an array of packed colors. }
{Each color is stored in 3 bytes (Red,Green,Blue), packed without any }
{intervening pad bytes. }

const CT_ENTRY_SIZE = 3; { Always 3 bytes, packed }
const CT_RED_OFFSET = 0;
const CT_GREEN_OFFSET = 1;
const CT_BLUE_OFFSET = 2;

{const CT_REDVALUE(x) = (x[CT_RED_OFFSET]);
const CT_GREENVALUE(x) = (x[CT_GREEN_OFFSET]);
const CT_BLUEVALUE(x) = (x[CT_BLUE_OFFSET]);

const CT_NEXT(x) = (x+=CT_ENTRY_SIZE);
const CT_ENTRY_PTR(x,ElmNum) = (&x[CT_ENTRY_SIZE*ElmNum]);}

{ A pattern table is a fixed-size color table used for patterns by }
{CDBITMAPHEADER (patterns are used to compress the bitmap). }
{A entry in the pattern table is 8 (PELS_PER_PATTERN) packed colors }
{(3 bytes per color as above). }

const PELS_PER_PATTERN = 8;

{ Maximum number of patterns we will ever store in a CDBITMAPHEADER. }

const
  MAXPATTERNS = 64;{Maximum number of colors in a color table (8 bit mapped color) }
  MAXCOLORS = 256;  // The CDBITMAPHEADER record must be present for all bitmaps. It must
                   //follow the CDGRAPHIC record, but come before any of the other bitmap
                   //CD records.

type
  CDBITMAPHEADER = packed record
             Header: LSIG; { Signature and Length }
             Dest: RECTSIZE; { dest bitmap height and width in PELS }
             Crop: RECTSIZE; { crop destination dimensions in PELS (UNUSED) }
             Flags: Word;    { CDBITMAP_FLAGS Valid only in CDGRAPHIC_VERSION2 and later }
             wReserved: Word; { Reserved for future use }
             lReserved: LongInt; { Reserved for future use }
             Width: Word; { Width of bitmap in PELS }
             Height: Word; { Height " " }
             BitsPerPixel: Word; { Bits per PEL, must be 1,8 or 16 }
             SamplesPerPixel: Word; { For 1 or 8 BPP, set to 1. For 16 BBP, set to 3 }
             BitsPerSample: Word; { For 1 BPP, set to 1. For 8 BPP, set to 8. For 16 BPP, set to 5 }
             SegmentCount: Word; { Number of CDBITMAPSEGMENTS }
             ColorCount: Word; { Number of entries in CDCOLORTABLE. 0-256 }
             PatternCount: Word; { Number of entries in CDPATTERNTABLE. Set to 0 if using Notes API. }
           end {_17};

{ Bitmap Uses > 16 colors or > 4 grey scale levels }

const CDBITMAP_FLAG_REQUIRES_PALETTE = 1;

{ Initialized by import code for "first time" importing of bitmaps }
{from clipboard or file, to tell Notes that it should compute whether }
{or not to use a color palette or not. All imports and API programs }
{should initially set this bit to let the Editor compute whether it }
{needs the palette or not. }

const CDBITMAP_FLAG_COMPUTE_PALETTE = 2;


{ Each of the following CDBITMAP segments contains the compressed raster }
{data of the bitmap. It is recommended that each segment be no larger }
{than 10K for optimal use within Notes, but try to keep the segments as }
{large as possible to increase painting speed. A scanline must not }
{span a segment. A bitmap must contain at least one segment, but may have }
{many segments. }


type CDBITMAPSEGMENT = packed record
             Header: LSIG; { Signature and Length }
             Reserved: Array[0..2-1] of LongInt; { Reserved for future use }
             ScanlineCount: Word; { Number of compressed scanlines in seg }
             DataSize: Word; { Size, in bytes, of compressed data Comressed raster data for the segment follows right here }
           end {_18};

{ Bitmap Color Table. If the bitmap is 8 bit color or grey scale, you }
{must have a color table. However, you only need as many entries as }
{you have colors, i.e. if a 16 color bitmap was converted to 8 bit }
{form for Notes, the color table would only require 16 entries even }
{though 8 bit color implies 256 entries. The number of entries must }
{match that specified in the CDBITMAPHEADER ColorCount. }


type CDCOLORTABLE = packed record
             Header: LSIG;
{ One or more color table entries go here }
           end {_19};

{ Bitmap Pattern Table (optionally one per bitmap) }


type CDPATTERNTABLE = packed record
             Header: LSIG; { One or more pattern table entries }
           end {_20};


{ Crop rectangle used in graphic run }


type
  CropRect = packed record
             left: Word;
             top: Word;
             right: Word;
             bottom: Word;
end {_21};

{ The Graphic combination record is used to store one or more graphic objects. }
{This record marks the beginning of a graphic composite item, and MUST }
{be present for any graphic object to be loaded/displayed. A }
{graphic composite item can be one or more of the following CD }
{record types: BITMAPHEADER, BITMAPSEGMENT, COLORTABLE, CGMMETA, }
{WINMETA,WINMETASEG,PMMETAHEADER,PMMETASEG,MACMETAHEADER,MACMETASEG. If }
{there is more than one graphic object, Notes will display only one object }
{using the following order: CGM Metafile, Native Metafile (i.e. Windows, }
{PM,Mac),Bitmap. }


type CDGRAPHIC = packed record
             Header: LSIG;
{ Signature and Length }
             DestSize: RECTSIZE;
{ Destination Display size in twips (1/1440 inch) }
             CropSize: RECTSIZE;
{ Width and Height of crop rect in twips. Currently unused }
             CropOffset: CROPRECT;
{ Crop rectangle offset from bottom left of Dest (in twips).Currently unused }
             fResize: Word;
{ Set to true if object has been resized by user. }
             Version: BYTE;
{ CDGRAPHIC_VERSION }
             bReserved: BYTE;
             wReserved: Word;
           end {_22};

{ Version control of graphic header }
const CDGRAPHIC_VERSION1 = 0; { Created by Notes version 2 }
const CDGRAPHIC_VERSION2 = 1; { Created by Notes version 3 }

{ CGM Metafile Record. This record follows the CDGRAPHIC record. It can }
{contain the entire contents of a CGM metafile, and must be <= 64K Bytes }
{in length. }


type CDCGMMETA = packed record
             Header: LSIG; { Signature and Length }
             mm: SWORD; { see above CGM_MAPMODE_??? }
             xExt,yExt: SWORD; { Extents of drawing in world coordinates }
             OriginalSize: RECTSIZE; { Original display size of metafile in twips }
                                    { CGM Metafile Bits Follow, must be <= 64K bytes total }
           end {_23};

const CGM_MAPMODE_ABSTRACT = 0; { Virtual coordinate system. This is default }
const CGM_MAPMODE_METRIC = 1; { Currently unsupported }


{ Windows Metafile Record. This record follows the CDGRAPHIC record and }
{contains the entire contents of a Windows GDI metafile. Since these }
{metafiles tend to be large, they may be segmented in chunks of any }
{arbitrary size, as long as each segment is <= 64K bytes. }


type
  CDWINMETAHEADER = packed record
   Header: LSIG; { Signature and Length }
   mm: SWORD; { Windows mapping mode }
   xExt,yExt: SWORD; { size in mapping mode units }
   OriginalDisplaySize: RECTSIZE; { Original display size of metafile in twips }
   MetafileSize: LongInt; { Total size of metafile raw data in bytes }
   SegCount: Word; { Number of CDWINMETASEG records Metafile segments Follow }
end {_24};


type
  CDWINMETASEG = packed record
             Header: LSIG; { Signature and Length }
             DataSize: Word; { Actual Size of metafile bits in bytes, ignoring any filler }
             SegSize: Word; { Size of segment, is equal to or larger than DataSize }
{if filler byte added to maintain word boundary }
{ Windows Metafile Bits for this segment. Each segment must be }
{<= 64K bytes. }
end {_25};


{ PM Metafile Record. This record follows the CDGRAPHIC record and }
{contains the entire contents of a PM GPI metafile. Since these }
{metafiles tend to be large, they may be segmented in chunks of any }
{arbitrary size, as long as each segment is <= 64K bytes. }


type CDPMMETAHEADER = packed record
             Header: LSIG;
{ Signature and Length of this record }
             mm: SWORD;
{ PM mapping mode, i.e. PU_??? }
             xExt,yExt: SWORD;
{ size in mapping mode units }
             OriginalDisplaySize: RECTSIZE;
{ Original display size of metafile in twips }
             MetafileSize: LongInt;
{ Total size of metafile raw data in bytes }
             SegCount: Word;
{ Number of CDPMMETASEG records }
           end {_26};


type
  CDPMMETASEG = packed record
   Header: LSIG; { Signature and Length }
   DataSize: Word; { Actual Size of metafile bits in bytes, ignoring any filler }
   SegSize: Word; { Size of segment, is equal to or larger than DataSize
                  if filler byte added to maintain word boundary
                  PM Metafile Bits for this segment. Must be <= 64K bytes. }
end {_27};

{ Document Record stored in $INFO field of a document. This contains }
{document-wide attributes. }
{ for FormFlags }
const TPL_FLAG_REFERENCE = $0001; { Use Reference Note }
const TPL_FLAG_MAIL = $0002; { Mail during DocSave }
const TPL_FLAG_NOTEREF = $0004; { Add note ref. to 'reference note' }
const TPL_FLAG_NOTEREF_MAIN = $0008; { Add note ref. to main parent of 'reference note' }
const TPL_FLAG_RECALC = $0010; { Recalc when leaving fields }
const TPL_FLAG_BOILERPLATE = $0020; { Store form item in with note }
const TPL_FLAG_FGCOLOR = $0040; { Use foreground color to paint }
const TPL_FLAG_SPARESOK = $0080; { Spare DWORDs have been zeroed }
const TPL_FLAG_ACTIVATE_OBJECT_COMP = $0100; { Activate OLE objects when composing a new doc }
const TPL_FLAG_ACTIVATE_OBJECT_EDIT = $0200; { Activate OLE objects when editing an existing doc }
const TPL_FLAG_ACTIVATE_OBJECT_READ = $0400; { Activate OLE objects when reading an existing doc }
const TPL_FLAG_SHOW_WINDOW_COMPOSE = $0800; { Show Editor window if TPL_FLAG_ACTIVATE_OBJECT_COMPOSE }
const TPL_FLAG_SHOW_WINDOW_EDIT = $1000; { Show Editor window if TPL_FLAG_ACTIVATE_OBJECT_EDIT }
const TPL_FLAG_SHOW_WINDOW_READ = $2000; { Show Editor window if TPL_FLAG_ACTIVATE_OBJECT_READ }
const TPL_FLAG_UPDATE_RESPONSE = $4000; { V3 Updates become responses }
const TPL_FLAG_UPDATE_PARENT = $8000; { V3 Updates become parents  for FormFlags2 }
const TPL_FLAG_INCLUDEREF = $0001; { insert copy of ref note }
const TPL_FLAG_RENDERREF = $0002; { render ref (else it's a doclink) }
const TPL_FLAG_RENDCOLLAPSE = $0004; { render it collapsed? }
const TPL_FLAG_EDITONOPEN = $0008; { edit mode on open }
const TPL_FLAG_OPENCNTXT = $0010; { open context panes }
const TPL_FLAG_CNTXTPARENT = $0020; { context pane is parent }
const TPL_FLAG_MANVCREATE = $0040; { manual versioning }
const TPL_FLAG_UPDATE_SIBLING = $0080; { V4 versioning - updates are sibblings }
const TPL_FLAG_ANONYMOUS = $0100; { V4 Anonymous form }
const TPL_FLAG_NAVIG_DOCLINK_IN_PLACE = $0200; { Doclink dive into same window }
const TPL_FLAG_INTERNOTES = $0400; { InterNotes special form }
const TPL_FLAG_DISABLE_FX = $0800; { Disable FX for this doc}




type
  CDDOCUMENT = packed record
     Header: BSIG;
     PaperColor: Word; { Color of the paper being used }
     FormFlags: Word; { Form Flags }
     NotePrivileges: Word; { Privs for notes created when using form }
{WARNING!!! Fields below this comment were not zeroed in builds }
{prior to 100. A mechanism has been set up to use them however. }
{dload checks the TPL_FLAG_SPARESOK bit in the flags word. If it }
{is not set, all of the storage after this comment is zeroed. On }
{save, dsave makes sure the unused storage is zero and sets the bit. }
     FormFlags2: Word; { more Form Flags }
     InherFieldNameLength: Word; { Length of the name, which follows this struct }
     PaperColorExt: Word; { Palette Color of the paper being used. New in V4. }
     Spare: Array[0..5-1] of Word;
      { ... now the Inherit Field Name string }
      { ... now the Text Field Name string indicating }
      {which field to append version number to }
   end {_30};

const ODS_COLOR_MASK = $00F; { Palette color is an index into a 240 entry table }
{ Header/Footer Record, stored in $HEADER and $FOOTER fields of a }
{document. This contains the header and footer used in the document. }


type CDHEADER = packed record
             Header: WSIG;
             FontPitchAndFamily: BYTE;
             FontName: Array[0..MAXFACESIZE-1] of Char;
             Font: FONTID;
             HeadLength: Word; {total header string length  ... now comes the string }
           end {_31};

{ Font Table Record, stored in the $FONTS field of a document. }
{This contains the list of "non-standard" fonts used in the }
{document. }


type
  CDFONTTABLE = packed record
             Header: WSIG; { Tag and length }
             Fonts: Word;  { Number of CDFACEs following }
           end {_32};
{Now come the CDFACE records... }
  CDFACE = packed record
             Face: BYTE;{ ID number of face }
             Family: BYTE;{ Font Family }
             Name: Array[0..MAXFACESIZE-1] of Char;
           end {_33};


{Print settings data structure - (stored in desktop file per icon) }
type
  PRINTNEW_SETTINGS = packed record
     Flags: Word; { PS_ flags below }
     StartingPageNum: Word; { Starting page number }
     TopMargin: Word; { Height between main body & top of page (TWIPS) }
     BottomMargin: Word; { Height between main body & bottom of page (TWIPS) }
     ExtraLeftMargin: Word; { Extra left margin width (TWIPS) (beyond whats already specified in document) }
     ExtraRightMargin: Word; { Extra right margin width (TWIPS) (beyond whats already specified in document) }
     HeaderMargin: Word; { Height between header & top of page (TWIPS) }
     FooterMargin: Word; { Height between footer & bottom of page (TWIPS) }
     PageWidth: Word; { Page width override (TWIPS) (0 = "use printer's page width") }
     PageHeight: Word; { Page height override (TWIPS) (0 = "use printer's page height") }
     BinFirstPage: Word; { Index of bin for 1st page }
     BinOtherPage: Word; { Index of bin for other pages }
     spare: Array[0..3-1] of LongInt; { (spare words) }
  end {_34};
  PRINT_SETTINGS = packed record
   Flags: Word;
   StartingPageNum: Word; { Starting page number }
   TopMargin: Word; { Height between main body & top of page (TWIPS) }
   BottomMargin: Word; { Height between main body & bottom of page (TWIPS) }
   ExtraLeftMargin: Word; { Extra left margin width (TWIPS) beyond whats already specified in document) }
   ExtraRightMargin: Word; { Extra right margin width (TWIPS) (beyond whats already specified in document) }
   HeaderMargin: Word; { Height between header & top of page (TWIPS) }
   FooterMargin: Word; { Height between footer & bottom of page (TWIPS) }
   PageWidth: Word; { Page width override (TWIPS) (0 = "use printer's page width") }
   PageHeight: Word; {Page height override (TWIPS) (0 = "use printer's page height") }
   BinFirstPage: Word;
   BinOtherPage: Word;
   spare: Array[0..3-1] of LongInt; {(spare words) }
end {_35};
const PS_Initialized = $0001; { Print settings have been initialized }
const PS_HeaderFooterOnFirst = $0002; { Print header/footer on first page }
const PS_CropMarks = $0004; { Print crop marks }
const PS_ChangeBin = $0008; { Paper source should be set for 1st & Other Pg. }




{ Header/Footer data structure - passed into import/export modules }


type HEAD_DESC = packed record
             FontPitchAndFamily: BYTE;
             FontName: Array[0..MAXFACESIZE-1] of Char;
             Font: FONTID;
             HeadLength: Word;
{ string length not including '\0' }
{ Header string (ASCIIZ) follows }
           end;

const MAXHEADERSTRING = 256; { maximum header string size }

type
  HEAD_DESC_BUFFER = packed record
{ used for stack-local ones }
   Desc: HEAD_DESC;
   aString: Array[0..MAXHEADERSTRING-1] of Char; { Must be terminated by '\0' }
end {_37};

{ DDE composite data On Disk structures }

const DDESERVERNAMEMAX = 32;
const DDEITEMNAMEMAX = 64;
const DDESERVERCOMMANDMAX = 256;


type
  CDDDEBEGIN = packed record
             Header: WSIG; { Signature and length of this record }
             ServerName: Array[0..DDESERVERNAMEMAX-1] of Char; { Null terminated server name }
             TopicName: Array[0..100-1] of Char; { Null terminated DDE Topic (usually a file name) }
             ItemName: Array[0..DDEITEMNAMEMAX-1] of Char; { Null terminated Place reference string }
             Flags: LongInt; { See DDEFLAGS_xxx flag definitions below }
             PasteEmbedDocName: Array[0..80-1] of Char; { only used on when making new link during Paste Special }
             EmbeddedDocCount: Word; { Number of embedded docs for this link (MUST BE 0 or 1) }
             ClipFormat: Word; {Clipboard format with which data should be rendered
                                (DDEFORMAT_xxx defined below) Null terminated embedded document name which is attached to the note follows.. }
           end {_38};

{ CDDDEBEGIN flags }

const DDEFLAGS_AUTOLINK = $01; { Link type == Automatic (hot) }
const DDEFLAGS_MANUALLINK = $02; { Link type == Manual (warm) }
const DDEFLAGS_EMBEDDED = $04; { Embedded document exists }
const DDEFLAGS_INITIATE = $08; { Used on paste to indicate not to}
const DDEFLAGS_CONV_ACTIVE = $40; { Used on non-CDP paste/load to indicate that}
const DDEFLAGS_NEWOBJECT = $100; { Set if this DDE Range is a new}
{ These remappings of Native clipboard formats are used because we can't }
{use Windows or PM constants because they are different }

const DDEFORMAT_TEXT = $01; { CF_TEXT }
const DDEFORMAT_METAFILE = $02; { CF_METAFILE or CF_METAFILEPICT }
const DDEFORMAT_BITMAP = $03; { CF_BITMAP }
const DDEFORMAT_RTF = $04; { Rich Text Format }
const DDEFORMAT_OWNERLINK = $06; { OLE Ownerlink (never saved in CD_DDE or CD_OLE: used at run time) }
const DDEFORMAT_OBJECTLINK = $07; { OLE Objectlink (never saved in CD_DDE or CD_OLE: used at run time) }
const DDEFORMAT_NATIVE = $08; { OLE Native (never saved in CD_DDE or CD_OLE: used at run time) }
const DDEFORMAT_ICON = $09; { Program Icon for embedded object }

{ Total number of DDE format types supported. Increment this if }
{one is added above }

const DDEFORMAT_TYPES = 5;



type
  CDDDEEND = packed record
             Header: WSIG;{ Signature and length of this record }
             Flags: LongInt;{ Currently unused, but reserve some flags }
end {_39};


{ On-disk format for an OLE object. Both Links and }
{embedded objects actually have an attached $FILE "object" }
{which is the variable length portion of the data which follows }
{the CDOLEBEGIN record. }


type CDOLEBEGIN = packed record
             Header: WSIG; {Signature and length of this record }
             Version: Word; {Notes OLE implementation version }
             Flags: LongInt; {See OLEREC_FLAG_xxx flag definitions below }
             ClipFormat: Word; {Clipboard format with which data should be rendered
                               (DDEFORMAT_xxx defined above)}
             AttachNameLength: Word; {Attached file name length }
             ClassNameLength: Word; {Used during Insert New Object, but never saved to disk }
             TemplateNameLength: Word; {User during Insert New Object, but never saved to disk }
{ The Attachment Name (length "AttachNameLength") always follows... }
{ The Classname, optional, then follows... }
{ The Template Name, optional, then follows... }
           end {_40};


type CDOLEEND = packed record
   Header: WSIG; {Signature and length of this record }
   Flags: LongInt; {Currently unused, but reserve some flags }
 end {_41};

{ Current OLE Version }

const NOTES_OLEVERSION1 = 1;
const NOTES_OLEVERSION2 = 2;

const OLEREC_FLAG_OBJECT = $01; { The data is an OLE embedded OBJECT }
const OLEREC_FLAG_LINK = $02; { The data is an OLE Link }
const OLEREC_FLAG_AUTOLINK = $04; { If link, Link type == Automatic (hot) }
const OLEREC_FLAG_MANUALLINK = $08; { If link, Link type == Manual (warm) }
const OLEREC_FLAG_NEWOBJECT = $10; { New object, just inserted }
const OLEREC_FLAG_PASTED = $20; { New object, just pasted }
const OLEREC_FLAG_SAVEOBJWHENCHANGED = $40; { Object came from form and should be saved}

{ On-disk format for HotSpots.}

{ HOTSPOT_RUN Types }

const HOTSPOTREC_TYPE_POPUP = 1;
const HOTSPOTREC_TYPE_HOTREGION = 2;
const HOTSPOTREC_TYPE_BUTTON = 3;
const HOTSPOTREC_TYPE_FILE = 4;
const HOTSPOTREC_TYPE_SECTION = 7;
const HOTSPOTREC_TYPE_ANY = 8;

const HOTSPOTREC_TYPE_HOTLINK = 11;
const HOTSPOTREC_TYPE_BUNDLE = 12;
const HOTSPOTREC_TYPE_V4_SECTION = 13;
const HOTSPOTREC_TYPE_SUBFORM = 14;

{ HOTSPOT_RUN Flags }

const HOTSPOTREC_RUNFLAG_BEGIN = $00000001;
const HOTSPOTREC_RUNFLAG_END = $00000002;
const HOTSPOTREC_RUNFLAG_BOX = $00000004;
const HOTSPOTREC_RUNFLAG_NOBORDER = $00000008;
const HOTSPOTREC_RUNFLAG_FORMULA = $00000010; { Popup is a formula, not text. }
{ Also defined in edit\hmem.h }
const HOTSPOTREC_RUNFLAG_INOTES = $00001000;
const HOTSPOTREC_RUNFLAG_ISMAP = $00002000;
const HOTSPOTREC_RUNFLAG_INOTES_AUTO = $00004000;
const HOTSPOTREC_RUNFLAG_ISMAP_INPUT = $00008000;


type
  CDHOTSPOTBEGIN = packed record
             Header: WSIG; { Signature and length of this record }
             aType: Word;
             Flags: DWORD;
             DataLength: Word; { Data Follows. }
  end {_42};
  CDHOTSPOTEND = packed record
    Header: BSIG; { Signature and length of this record }
  end {_43};

{On-disk format for Buttons}
const BUTTONREC_IS_DOWN = $0002;
const BUTTONREC_IS_EDITABLE = $0004;
const BUTTONREC_FLAG_CARET_ON = $0008;
const BUTTONREC_FLAG_RESIZE_ON = $0010;
const BUTTONREC_FLAG_DISABLED = $0020;


type
  CDBUTTON = packed record
     Header: WSIG; { Signature and length of this record. }
     Flags: Word;
     Width: Word;
     Height: Word;
     Lines: Word;
     FontID: FONTID; { Button Text Follows}
   end {_44};


{ On-disk format for Object Bars. }

const BARREC_DISABLED_FOR_NON_EDITORS   = 1;
const BARREC_EXPANDED                   = 2;
const BARREC_PREVIEW                    = 4;

const BARREC_BORDER_INVISIBLE           = $1000;
const BARREC_ISFORMULA                  = $2000;
const BARREC_HIDE_EXPANDED              = $4000;

//* Auto expand/collapse properties.  */

const BARREC_AUTO_EXP_READ  = $10;
const BARREC_AUTO_EXP_PRE   = $20;
const BARREC_AUTO_EXP_EDIT  = $40;
const BARREC_AUTO_EXP_PRINT = $80;

const BARREC_AUTO_EXP_MASK  = $f0;

const BARREC_AUTO_COL_READ  = $100;
const BARREC_AUTO_COL_PRE   = $200;
const BARREC_AUTO_COL_EDIT  = $400;
const BARREC_AUTO_COL_PRINT = $800;
const BARREC_AUTO_COL_MASK  = $F00;

const BARREC_AUTO_PRE_MASK  = (BARREC_AUTO_COL_PRE or BARREC_AUTO_EXP_PRE);
const BARREC_AUTO_READ_MASK = (BARREC_AUTO_COL_READ or BARREC_AUTO_EXP_READ);
const BARREC_AUTO_EDIT_MASK = (BARREC_AUTO_COL_EDIT or BARREC_AUTO_EXP_EDIT);
const BARREC_AUTO_PRINT_MASK = (BARREC_AUTO_COL_PRINT or BARREC_AUTO_EXP_PRINT);

{/* We will make use (in the code) of the fact that the auto expand/collapse
  flags for editors are simply shifted left twelve bits from the normal
  expand/collapse flags.
*/}

const BARREC_AUTO_ED_SHIFT    = $12;

const BARREC_AUTO_ED_EXP_READ   = $10000;
const BARREC_AUTO_ED_EXP_PRE    = $20000;
const BARREC_AUTO_ED_EXP_EDIT   = $40000;
const BARREC_AUTO_ED_EXP_PRINT  = $80000;

const BARREC_AUTO_ED_EXP_MASK   = $f00000;

const BARREC_AUTO_ED_COL_READ   = $100000;
const BARREC_AUTO_ED_COL_PRE    = $200000;
const BARREC_AUTO_ED_COL_EDIT   = $400000;
const BARREC_AUTO_ED_COL_PRINT  = $800000;
const BARREC_AUTO_ED_COL_MASK   = $F00000;

const BARREC_AUTO_ED_PRE_MASK   = (BARREC_AUTO_ED_COL_PRE or BARREC_AUTO_ED_EXP_PRE);
const BARREC_AUTO_ED_READ_MASK  = (BARREC_AUTO_ED_COL_READ or BARREC_AUTO_ED_EXP_READ);
const BARREC_AUTO_ED_EDIT_MASK  = (BARREC_AUTO_ED_COL_EDIT or BARREC_AUTO_ED_EXP_EDIT);
const BARREC_AUTO_ED_PRINT_MASK = (BARREC_AUTO_ED_COL_PRINT or BARREC_AUTO_ED_EXP_PRINT);

const BARREC_INTENDED       = $1000000;
const BARREC_HAS_COLOR      = $4000000;

const BARREC_BORDER_MASK      = $70000000;

function GetBorderType (const a: DWORD): DWORD;
procedure SetBorderType (var a: DWORD; const b: DWORD);

const BARREC_BORDER_SHADOW    = $0;
const BARREC_BORDER_NONE      = $1;
const BARREC_BORDER_SINGLE    = $2;
const BARREC_BORDER_DOUBLE    = $3;
const BARREC_BORDER_TRIPLE    = $4;
const BARREC_BORDER_TWOLINE   = $5;

const BARREC_INDENTED   = $80000000;

{/* Indicate explicitly those bits that we want to save on-disk
  so that we insure that the others are zero when we save to
  disk so that we can use later.
*/}

const BARREC_ODS_MASK = $F4FF6FF7;

//* On-disk format for Object Bars. */

const BARREC_IS_EXPANDED = $0001;
const BARREC_IS_DISABLED = $0002;



type
  CDBAR = packed record
   Header: WSIG;
   Flags: LongInt;
   FontID: FONTID;{ Caption and name follow }
end {_45};

{On-disk format for form layout objects }

const LAYOUT_FLAG_SHOWBORDER = $00000001;
const LAYOUT_FLAG_SHOWGRID = $00000002;
const LAYOUT_FLAG_SNAPTOGRID = $00000004;
const LAYOUT_FLAG_3DSTYLE = $00000008;


type
  CDLAYOUT = packed record
     Header: BSIG;
     wLeft: Word;
     wWidth: Word;
     wHeight: Word;
     Flags: LongInt;
     wGridSize: Word;
     Reserved: Array[0..14-1] of BYTE;
   end {_46};
  RElementHeader= packed record
     wLeft: Word;
     wTop: Word;
     wWidth: Word;
     wHeight: Word;
     FontID: FONTID; { Font ID }
     byBackColor: BYTE; { Background color }
     bSpare: Array[0..7-1] of BYTE;
   end {_47};

{ The following flags must be the same as LAYOUT_FIELD_FLAG_ equiv's. }
const LAYOUT_TEXT_FLAG_TRANS = $10000000;
const LAYOUT_TEXT_FLAG_LEFT = $00000000;
const LAYOUT_TEXT_FLAG_CENTER = $20000000;
const LAYOUT_TEXT_FLAG_RIGHT = $40000000;
const LAYOUT_TEXT_FLAG_ALIGN_MASK = $60000000;
const LAYOUT_TEXT_FLAG_VCENTER = $80000000;
const LAYOUT_TEXT_FLAGS_MASK = $F0000000;


type
  CDLAYOUTTEXT = packed record
   Header: BSIG;
   ElementHeader: RElementHeader;
   Flags: LongInt;
   Reserved: Array[0..16-1] of BYTE; { For records save with builds prior to 134 the 8-bit text string follows... }
end {_48};

const LAYOUT_FIELD_TYPE_TEXT = 0;
const LAYOUT_FIELD_TYPE_CHECK = 1;
const LAYOUT_FIELD_TYPE_RADIO = 2;
const LAYOUT_FIELD_TYPE_LIST = 3;
const LAYOUT_FIELD_TYPE_COMBO = 4;

const LAYOUT_FIELD_FLAG_SINGLELINE = $00000001;
const LAYOUT_FIELD_FLAG_VSCROLL = $00000002;
{The following flag must not be sampled by any design mode code. It is, in effect, "write only" for design elements.
Play mode elements, on the other hand, can rely on its value. }
const LAYOUT_FIELD_FLAG_MULTISEL = $00000004;
const LAYOUT_FIELD_FLAG_STATIC = $00000008;
const LAYOUT_FIELD_FLAG_NOBORDER = $00000010;
const LAYOUT_FIELD_FLAG_IMAGE = $00000020;
{The following flags must be the same as LAYOUT_TEXT_FLAG_ equiv's. }
const LAYOUT_FIELD_FLAG_TRANS = $10000000;
const LAYOUT_FIELD_FLAG_LEFT = $00000000;
const LAYOUT_FIELD_FLAG_CENTER = $20000000;
const LAYOUT_FIELD_FLAG_RIGHT = $40000000;
const LAYOUT_FIELD_FLAG_VCENTER = $80000000;
const LAYOUT_GRAPHIC_FLAG_BUTTON = $00000001;

type
  CDLAYOUTFIELD = packed record
     Header: BSIG;
     ElementHeader: RELEMENTHEADER;
     Flags: LongInt;
     bFieldType: BYTE;
     Reserved: Array[0..15-1] of BYTE;
   end {_49};
  CDLAYOUTGRAPHIC = packed record
   Header: BSIG;
   ElementHeader: rELEMENTHEADER;
   Flags: LongInt;
   Reserved: Array[0..16-1] of BYTE;
 end {_50};
  CDLAYOUTBUTTON = packed record
     Header: BSIG;
     ElementHeader: RELEMENTHEADER;
     Flags: LongInt;
     Reserved: Array[0..16-1] of BYTE;
  end {_51};
  CDLAYOUTEND = packed record
   Header: BSIG;
   Reserved: Array[0..16-1] of BYTE;
end {_52};

const ONEINCH = (20*72);      //* One inch worth of TWIPS */
const JUSTIFY_LEFT    = 0;  //* flush left, ragged right */
const JUSTIFY_RIGHT   = 1;  //* flush right, ragged left */
const JUSTIFY_BLOCK   = 2;  //* full block justification */
const JUSTIFY_CENTER    = 3;  //* centered */
const JUSTIFY_NONE    = 4;  //* no line wrapping AT ALL (except hard CRs) */

const
  DEFAULT_JUSTIFICATION     = JUSTIFY_LEFT;
  DEFAULT_LINE_SPACING      = 0;
  DEFAULT_ABOVE_PAR_SPACING = 0;
  DEFAULT_BELOW_PAR_SPACING = 0;
  DEFAULT_LEFT_MARGIN       = ONEINCH;
  DEFAULT_FIRST_LEFT_MARGIN = DEFAULT_LEFT_MARGIN;
  DEFAULT_RIGHT_MARGIN      = 0;

//* Note: Right Margin = "0" means [DEFAULT_RIGHT_GUTTER] inches from */
//* right edge of paper, regardless of paper width. */
  DEFAULT_RIGHT_GUTTER    = ONEINCH;
  DEFAULT_PAGINATION      = 0;
  DEFAULT_FLAGS2        = 0;

//* Note: tabs are relative to the absolute left edge of the paper. */
  DEFAULT_TABS        = 0;
  DEFAULT_TAB_INTERVAL    = (ONEINCH div 2);
  DEFAULT_TABLE_HCELLSPACE  = 0;
  DEFAULT_TABLE_VCELLSPACE  = 0;

  DEFAULT_LAYOUT_LEFT     = DEFAULT_LEFT_MARGIN;
  DEFAULT_LAYOUT_WIDTH    = (ONEINCH * 6);
  DEFAULT_LAYOUT_HEIGHT   = (3 * ONEINCH / 2);
  MIN_LAYOUT_WIDTH      = (ONEINCH / 4);
  MIN_LAYOUT_HEIGHT     = (ONEINCH / 4);

  DEFAULT_LAYOUT_ELEM_WIDTH = (4 * ONEINCH / 3);  //* 1.333 inch */
  DEFAULT_LAYOUT_ELEM_HEIGHT  = (ONEINCH / 5);
  MIN_ELEMENT_WIDTH     = (ONEINCH / 8);
  MIN_ELEMENT_HEIGHT      = (ONEINCH / 8);

//* Horizontal Rule Defaults        */

  DEFAULTHRULEHEIGHT  = 7;
  DEFAULTHRULEWIDTH = 720;

type
  CDHTMLBEGIN = packed record
     Header: WSIG;
     Spares: DWORD;
  end;
  CDHTMLEND = packed record
     Header: WSIG;
     Spares: DWORD;
  end;

(******************************************************************************)
{COMPOUND TEXT FUNCTIONS}
{FROM EASYCD.H}
(******************************************************************************)
type
  CompoundStyle = packed record
    JustifyMode: Word;
    LineSpacing: Word;
    ParagraphSpacingBefore: Word;
    ParagraphSpacingAfter: Word;
    LeftMargin: Word;
    RightMargin: Word;
    FirstLineLeftMargin: Word;
    Tabs: Word;
    Tab: Array [0..MAXTABS-1] of SWORD;
    Flags: Word;
  end;
  PCompoundStyle = ^CompoundStyle;

{ Flags for CompoundTextAddText. }

const COMP_FROM_FILE = $00000001;
const COMP_PRESERVE_LINES = $00000002;
const COMP_PARA_LINE = $00000004;
const COMP_PARA_BLANK_LINE = $00000008;
const COMP_SERVER_HINT_FOLLOWS = $00000010;

{ Use this style ID in CompoundTextAddText to continue using the }
{same paragraph style as the previous paragraph. }
const STYLE_ID_SAMEASPREV = $FFFFFFF;

{ Font IDs for SetFaceID }
const
  FONT_FACE_ROMAN = 0;      //Tms Roman
  FONT_FACE_SWISS = 1;      //Helv
  FONT_FACE_TYPEWRITER = 4; //Courier
  STATIC_FONT_FACES = 5;    //??

// Font components mask
const
  FONT_SIZE_SHIFT   = 24;
  FONT_COLOR_SHIFT  = 16;
  FONT_STYLE_SHIFT  = 8;
  FONT_FACE_SHIFT   = 0;

{ Font styles for SetStyle. Use OR combination }
const
  CF_ISBOLD   = $01;
  CF_ISITALIC = $02;
  CF_ISUNDERLINE  = $04;
  CF_ISSTRIKEOUT  = $08;
  CF_ISSUPER    = $10;
  CF_ISSUB    = $20;
  CF_ISEFFECT = $80;
  CF_ISSHADOW = $80;
  CF_ISEMBOSS = $90;
  CF_ISEXTRUDE  = $a0;

procedure FontIDSetSize (var fontid: dword; size: integer);
procedure FontIDSetFaceID (var fontid: dword; faceId: dword);
procedure FontIDSetColor (var fontid: dword; colorId: dword);
procedure FontIDSetStyle (var fontid: dword; styleId: dword);
function FontIDIsUnderline(const fontid: dword): boolean;
function FontIDIsBold(const fontid: dword): boolean;
function FontIDIsItalic(const fontid: dword): boolean;
function FontIDIsStrikeout(const fontid: dword): boolean;
function FontIDIsSuperscript(const fontid: dword): boolean;
function FontIDIsSubscript(const fontid: dword): boolean;
function FontIDIsShadow(const fontid: dword): boolean;
function FontIDIsExtrude(const fontid: dword): boolean;

function DEFAULT_FONT_ID: dword;

// New functions by Winalot
function FontIDGetSize (const fontid: dword): integer;
 //Maps to NOTES_COLOR_XXX constants...
function FontIDGetColor (const fontid: dword): integer;
 //Maps to FONT_FACE_XXX constants...
function FontIDGetFace (const fontid: dword): integer;

{ Function prototypes. }
function CompoundTextCreate(hNote: NOTEHANDLE;
                            pszItemName: PChar;
                            phCompound: PHandle): STATUS; stdcall; far;

function CompoundTextClose(hCompound: THandle;
                           phReturnBuffer: PHandle;
                           pdwReturnBufferSize: PLongInt;
                           pchReturnFile: PChar;
                           wReturnFileSize: Word): STATUS; stdcall; far;

procedure CompoundTextDiscard(hCompound: THandle); stdcall; far;

function CompoundTextDefineStyle(hCompound: THandle;
                                 pszStyleName: PChar;
                                 pDefinition: PCOMPOUNDSTYLE;
                                 pdwStyleId: PLongInt): STATUS; stdcall; far;

function CompoundTextAssimilateItem(hCompound: THandle;
                                    hNote: NOTEHANDLE;
                                    pszItemName: PChar;
                                    dwFlags: LongInt): STATUS; stdcall; far;

function CompoundTextAssimilateFile(hCompound: THandle;
                                    pszFileSpec: PChar;
                                    dwFlags: LongInt): STATUS; stdcall; far;

function CompoundTextAddParagraph(hCompound: THandle;
                                  dwStyleId: LongInt;
                                  FontID: FONTID;
                                  pchText: PChar;
                                  dwTextLen: LongInt;
                                  hCLSTable: THandle): STATUS; stdcall; far;

function CompoundTextAddText(hCompound: THandle;
                             dwStyleId: LongInt;
                             FontID: FONTID;
                             pchText: PChar;
                             dwTextLen: LongInt;
                             pszLineDelim: PChar;
                             dwFlags: LongInt;
                             hCLSTable: THandle): STATUS; stdcall; far;

procedure CompoundTextInitStyle(pStyle: PCOMPOUNDSTYLE); stdcall; far;


function CompoundTextAddDocLink(hCompound: THandle;
                                DBReplicaID: TIMEDATE;
                                ViewUNID: UNID;
                                NoteUNID: UNID;
                                pszComment: PChar;
                                dwFlags: LongInt): STATUS; stdcall; far;


function CompoundTextAddRenderedNote(hCompound: THandle;
                                     hNote: NOTEHANDLE;
                                     hFormNote: NOTEHANDLE;
                                     dwFlags: LongInt): STATUS; stdcall; far;
(******************************************************************************)
{OsMisc.H}
(******************************************************************************)
type
  Nls_PInfo = pointer;

function OSLoadString(hModule: HMODULE;
                      StringCode: STATUS;
                      retBuffer: PChar;
                      BufferLength: Word): Word; stdcall; far;
                      
{ Charsets used with OSTranslate }

const OS_TRANSLATE_NATIVE_TO_LMBCS = 0; { Translate platform-specific to LMBCS }
const OS_TRANSLATE_LMBCS_TO_NATIVE = 1; { Translate LMBCS to platform-specific }
const OS_TRANSLATE_LOWER_TO_UPPER = 3; { current int'l case table }
const OS_TRANSLATE_UPPER_TO_LOWER = 4; { current int'l case table }
const OS_TRANSLATE_UNACCENT = 5; { int'l unaccenting table }

{$IFDEF DOS}
const OS_TRANSLATE_OSNATIVE_TO_LMBCS = 7; { Used in DOS (codepage) }
const OS_TRANSLATE_LMBCS_TO_OSNATIVE = 8; { Used in DOS }
{$ELSE defined (OS2)}
const OS_TRANSLATE_OSNATIVE_TO_LMBCS = OS_TRANSLATE_NATIVE_TO_LMBCS;
const OS_TRANSLATE_LMBCS_TO_OSNATIVE = OS_TRANSLATE_LMBCS_TO_NATIVE;
{$ELSE}
const OS_TRANSLATE_OSNATIVE_TO_LMBCS = OS_TRANSLATE_NATIVE_TO_LMBCS;
const OS_TRANSLATE_LMBCS_TO_OSNATIVE = OS_TRANSLATE_LMBCS_TO_NATIVE;
{$ENDIF}

{$IFDEF DOS || defined(OS2)}
const OS_TRANSLATE_LMBCS_TO_ASCII = 13;
{$ELSE}
const OS_TRANSLATE_LMBCS_TO_ASCII = 11;
{$ENDIF}

{ Character Set Translation Routines }


function OSTranslate(TranslateMode: Word;
                     sIn: PChar;
                     InLength: Word;
                     Out: PChar;
                     OutLength: Word): Word; stdcall; far;

{ Dynamic link library portable load routines }


function OSLoadLibrary(LibraryName: PChar;
                       Flags: LongInt;
                       var rethModule: HMODULE;
                       var retEntryPoint: FARPROC): STATUS; stdcall; far;

procedure OSFreeLibrary(_1: HMODULE); stdcall; far;

{ Routine used in non-premptive platforms to simulate it. }


procedure OSPreemptOccasionally; stdcall; far;

function OSGetLMBCSCLS: NLS_PINFO; stdcall; far;

function OSGetNativeCLS: NLS_PINFO; stdcall; far;

(******************************************************************************)
{reg.h}
(******************************************************************************)
 {REGIDGetxxx - Information type codes for ID files.}
const REGIDGetUSAFlag		      =1;    { Data structure returned is BOOL }
const REGIDGetHierarchicalFlag	=2;    { Data structure returned is BOOL }
const REGIDGetSafeFlag			    =3;    { Data structure returned is BOOL }
const REGIDGetCertifierFlag		=4;    { Data structure returned is BOOL }
const REGIDGetNotesExpressFlag	=5;    { Data structure returned is BOOL }
const REGIDGetDesktopFlag			=6;    { Data structure returned is BOOL }
const REGIDGetName				      =7;    { Data structure returned is char xx[MAXUSERNAME] }
const REGIDGetPublicKey			  =8;    { Data structure returned is char xx[xx] }
const REGIDGetPrivateKey			  =9;    { Data structure returned is char xx[xx] }
const REGIDGetIntlPublicKey		=10;   { Data structure returned is char xx[xx] }
const REGIDGetIntlPrivateKey		=11;   { Data structure returned is char xx[xx] }

function REGGetIDInfo(
  IDFileName: pchar;
  InfoType: WORD;
  OutBufr: pointer;
  OutBufrLen: WORD;
  ActualLen: PWORD): STATUS; stdcall; far;

(******************************************************************************)
{kfm.h}
(******************************************************************************)

{ Structure returned by KFMCreatePassword to encode a password securely }
{in memory to avoid scavenging. }


type
  KFM_PASSWORD = packed record
            aType: BYTE; {type as shown is "0". This field is a hook for future compatibility. }
            HashedPassword: Array[0..48-1] of BYTE; { Hashed password }
  end {_1};
  PKFM_PASSWORD = ^KFM_PASSWORD;
  HCERTIFIER = THandle;
  PHCERTIFIER = ^HCERTIFIER;
const
  NULLHCERTIFIER: HCERTIFIER=0;

{ Aliases for public routines }
{ }

{$IFNDEF SEMAPHORES}
{$DEFINE _NOSEMS_OR_BSAFE_INTERNAL_}
{$ENDIF}



{ Function codes for routine SECKFMUserInfo }
{ }

const KFM_ui_GetUserInfo = 1;

{ Function codes for routine SECKFMGetPublicKey }
{ }

const KFM_pubkey_Primary = 0;
const KFM_pubkey_International = 1;

{ Public Routines }
{ }

function SECKFMSwitchToIDFile( pIDFileName:PChar;
                               pPassword:PChar;
                               pUserName:PChar;
                               MaxUserNameLength:WORD;
                               Flags:DWORD;
                               pReserved:PChar): STATUS; stdcall; far;

function SECKFMUserInfo(aFunction: Word;
                        lpName: PChar;
                        var lpLicense: LICENSEID): STATUS; stdcall; far;

function SECKFMGetUserName(retUserName: PChar): STATUS; stdcall; far;



function SECKFMGetCertifierCtx(pCertFile: PChar;
                               pKfmPW: PKFM_PASSWORD;
                               pLogFile: PChar;
                               pExpDate: PTIMEDATE;
                               retCertName: PChar;
                               rethKfmCertCtx: PHCERTIFIER;
                               retfIsHierarchical: PBool;
                               retwFileVersion: PWord): STATUS; stdcall; far;

function SECKFMSetCertifierExpiration(hKfmCertCtx: HCERTIFIER;
                                      pExpirationDate: PTIMEDATE): STATUS; stdcall; far;

function SECKFMGetPublicKey(pName: PChar;
                            aFunction: Word;
                            Flags: Word;
                            rethPubKey: PHandle): STATUS; stdcall; far;

{ Constants used to indicate various types of IDs that can be created. }

const KFM_IDFILE_TYPE_FLAT = 0; { Flat name space (V2 compatible) }
const KFM_IDFILE_TYPE_STD = 1; { Standard (user/server hierarchical) }
const KFM_IDFILE_TYPE_ORG = 2; { Organization certifier }
const KFM_IDFILE_TYPE_ORGUNIT = 3; { Organizational unit certifier }
const KFM_IDFILE_TYPE_DERIVED = 4; { Derived from certifer context. }
{ hierarchical => STD }
{ non-hierarchical => FLAT }


(******************************************************************************)
{OsMem.h}
(******************************************************************************)

{Memory manager package}


function OSMemAlloc (BlkType: Word;
                     dwSize: LongInt;
                     retHandle: PHandle): STATUS; stdcall; far;

function OSMemFree (Handle: THandle): STATUS; stdcall; far;

function OSMemGetSize (Handle: THandle;
                       retSize: PLongInt): STATUS; stdcall; far;

function OSMemRealloc (Handle: THandle;
                       NewSize: LongInt): STATUS; stdcall; far;

function OSLockObject (Handle: THandle): pointer; stdcall; far;

function OSLockBlock (BlckId: BLOCKID): pointer;

procedure OSUnlockBlock(BlckId: BLOCKID);

//const OSLock(blocktype,handle) = ((blocktype far * ) OSLockObject(handle));
//function OSLOCK ();
//function OSLock(blocktype,handle)
//const OSUnlock(handle) = OSUnlockObject(handle);
//function OSUnlock(handle)

function OSUnlockObject(Handle: THandle): Bool; stdcall; far;

{ Define the maximum single-segment memory object size, because OSMem needs }
{space for overhead. Also, Windows has a restriction that it also adds }
{another 16 bytes of overhead to the request (arena header), and if that }
{causes the object to grow into a "huge" object (more than one segment), }
{most caller's will crash because Windows will actually change a segment's }
{address when a huge object gets reallocated. So, for Windows only, }
{we restrict the maximum size of a segment to allow for both our overhead }
{and the Windows arena header overhead to avoid huge allocations. }
{Note that we are subtracting odd numbers from MAXWORD. This is because }
{the result needs to be an even number. Otherwise, if the requested size }
{were MAXONESEGSIZE, memalloc.c would bump the size up over MAXONESEGSIZE }
{in order to keep it even. }

{NOTE: Beginning in V3.2, we define MAXONESEGSIZE to be the MIN of }
{the required MAXONESEGSIZE on all platforms, because, for example, }
{if a server (such as OS/2) allocates an object that IT THINKS is }
{MAXONESEGSIZE and then sends it to a client, he'll choke on it. }

(******************************************************************************)
{OsEnv.H}
(******************************************************************************)
{Size of the buffer used to hold the environment variable values (i.e., it }
{excludes the variable name) but including the trailing null terminator. }

{NOTE: The largest known example of an environment variable value is a }
{max'ed out COMx=... (the modem init strings can be large, and }
{there's plenty of them). }

{ Used to preface ini variables that are different between OSs which may }
{ share the same INI file. }

const OS_PREFIX = 'WIN';

function OSGetEnvironmentString(VariableName: PChar;
                                retValueBuffer: PChar;
                                BufferLength: Word): Bool; stdcall; far;

function OSGetEnvironmentLong(VariableName: PChar): LongInt; stdcall; far;

procedure OSSetEnvironmentVariable(VariableName: PChar;
                                   Value: PChar); stdcall; far;

procedure OSSetEnvironmentInt(VariableName: PChar;
                              Value: Integer); stdcall; far;

(******************************************************************************)
{OsFile.H}
(******************************************************************************)
{ File system interface }
{ File type flags (used with NSFSearch directory searching). }


const FILE_ANY = 0; { Any file type }
const FILE_DBREPL = 1; { Starting in V3, any DB that is a candidate for replication }
const FILE_DBDESIGN = 2; { Databases that can be templates }
{ 3 unused }
const FILE_DBANY = 4; { NS?, any NSF version }
const FILE_FTANY = 5; { NT?, any NTF version }
const FILE_MDMTYPE = 6; { MDM - modem command file }
const FILE_DIRSONLY = 7; { directories only }
const FILE_VPCTYPE = 8; { VPC - virtual port command file }
const FILE_SCRTYPE = 9; { SCR - comm port script files }

const FILE_TYPEMASK = $00ff; { File type mask (for FILE_xxx codes above) }
const FILE_DIRS = $8000; { List subdirectories as well as normal files }
const FILE_NOUPDIRS = $4000; { Do NOT return ..'s }
const FILE_RECURSE = $2000; { Recurse into subdirectories }

function OSPathNetConstruct(PortName: PChar;
                            ServerName: PChar;
                            FileName: PChar;
                            retPathName: PChar): STATUS; stdcall; far;
function OSPathNetParse(PathName: PChar;
                        retPortName: PChar;
                        retServerName: PChar;
                        retFileName: PChar): STATUS; stdcall; far;
function OSGetDataDirectory(retPathName: PChar): Word; stdcall; far;

(******************************************************************************)
{Nif.H}
(******************************************************************************)
{ NIF manipulation routines & basic datatypes }
{ Collection handle }

type HCOLLECTION = Word;
{ Handle to NIF collection }
const NULLHCOLLECTION:HCollection=0;


{ NIFOpenCollection "open" flags }

const OPEN_REBUILD_INDEX = $0001; { Throw away existing index and }
{ rebuild it from scratch }
const OPEN_NOUPDATE = $0002; { Do not update index or unread }
{ list as part of open (usually }
{ set by server when it does it }
{ incrementally instead). }
const OPEN_DO_NOT_CREATE = $0004; { If collection object has not yet }
{ been created, do NOT create it }
{ automatically, but instead return }
{ a special internal error called }
{ ERR_COLLECTION_NOT_CREATED }
const OPEN_SHARED_VIEW_NOTE = $0010; { Tells NIF to 'own' the view note }
{ (which gets read while opening the }
{ collection) in memory, rather than }
{ the caller "owning" the view note }
{ by default. If this flag is specified }
{ on subsequent opens, and NIF currently }
{ owns a copy of the view note, it }
{ will just pass back the view note }
{ handle rather than re-reading it }
{ from disk/network. If specified, }
{ the the caller does NOT have to }
{ close the handle. If not specified, }
{ the caller gets a separate copy, }
{ and has to NSFNoteClose the }
{ handle when its done with it. }
const OPEN_REOPEN_COLLECTION = $0020; { Force re-open of collection and }
{ thus, re-read of view note. }
{ Also implicitly prevents sharing }
{ of collection handle, and thus }
{ prevents any sharing of associated }
{ structures such as unread lists, etc }


{ Collection navigation directives }

const NAVIGATE_CURRENT = 0; { Remain at current position }
{ (reset position & return data) }
const NAVIGATE_PARENT = 3; { Up 1 level }
const NAVIGATE_CHILD = 4; { Down 1 level to first child }
const NAVIGATE_NEXT_PEER = 5; { Next node at our level }
const NAVIGATE_PREV_PEER = 6; { Prev node at our level }
const NAVIGATE_FIRST_PEER = 7; { First node at our level }
const NAVIGATE_LAST_PEER = 8; { Last node at our level }
const NAVIGATE_CURRENT_MAIN = 11; { Highest level non-category entry }
const NAVIGATE_NEXT_MAIN = 12; { CURRENT_MAIN, then NEXT_PEER }
const NAVIGATE_PREV_MAIN = 13; { CURRENT_MAIN, then PREV_PEER only if already there }
const NAVIGATE_NEXT_PARENT = 19; { PARENT, then NEXT_PEER }
const NAVIGATE_PREV_PARENT = 20; { PARENT, then PREV_PEER }

const NAVIGATE_NEXT = 1; { Next entry over entire tree }
{ (parent first, then children,...) }
const NAVIGATE_PREV = 9; { Previous entry over entire tree }
{ (opposite order of PREORDER) }
const NAVIGATE_ALL_DESCENDANTS = 17; { NEXT, but only descendants }
{ below NIFReadEntries StartPos }
const NAVIGATE_NEXT_UNREAD = 10; { NEXT, but only 'unread' entries }
const NAVIGATE_NEXT_UNREAD_MAIN = 18; { NEXT_UNREAD, but stop at main note also }
const NAVIGATE_PREV_UNREAD_MAIN = 34; { Previous unread main. }
const NAVIGATE_PREV_UNREAD = 21; { PREV, but only 'unread' entries }
const NAVIGATE_NEXT_SELECTED = 14; { NEXT, but only 'selected' entries }
const NAVIGATE_PREV_SELECTED = 22; { PREV, but only 'selected' entries }
const NAVIGATE_NEXT_SELECTED_MAIN = 32; { Next selected main. (Next unread }
{ main can be found above.) }
const NAVIGATE_PREV_SELECTED_MAIN = 33; { Previous selected main. }
const NAVIGATE_NEXT_EXPANDED = 15; { NEXT, but only 'expanded' entries }
const NAVIGATE_PREV_EXPANDED = 16; { PREV, but only 'expanded' entries }
const NAVIGATE_NEXT_EXPANDED_UNREAD = 23; { NEXT, but only 'expanded' AND 'unread' entries }
const NAVIGATE_PREV_EXPANDED_UNREAD = 24; { PREV, but only 'expanded' AND 'unread' entries }
const NAVIGATE_NEXT_EXPANDED_SELECTED = 25; { NEXT, but only 'expanded' AND 'selected' entries }
const NAVIGATE_PREV_EXPANDED_SELECTED = 26; { PREV, but only 'expanded' AND 'selected' entries }
const NAVIGATE_NEXT_EXPANDED_CATEGORY = 27; { NEXT, but only 'expanded' AND 'category' entries }
const NAVIGATE_PREV_EXPANDED_CATEGORY = 28; { PREV, but only 'expanded' AND 'category' entries }
const NAVIGATE_NEXT_EXP_NONCATEGORY = 39; { NEXT, but only 'expanded' 'non-category' entries }
const NAVIGATE_PREV_EXP_NONCATEGORY = 40; { PREV, but only 'expanded' 'non-category' entries }
const NAVIGATE_NEXT_HIT = 29; { NEXT, but only FTSearch 'hit' entries }
{ (in the SAME ORDER as the hit's relevance ranking) }
const NAVIGATE_PREV_HIT = 30; { PREV, but only FTSearch 'hit' entries }
{ (in the SAME ORDER as the hit's relevance ranking) }
const NAVIGATE_CURRENT_HIT = 31; { Remain at current position in hit's relevance rank array }
{ (in the order of the hit's relevance ranking) }
const NAVIGATE_NEXT_SELECTED_HIT = 35; { NEXT, but only 'selected' and FTSearch 'hit' entries }
{ (in the SAME ORDER as the hit's relevance ranking) }
const NAVIGATE_PREV_SELECTED_HIT = 36; { PREV, but only 'selected' and FTSearch 'hit' entries }
{ (in the SAME ORDER as the hit's relevance ranking) }
const NAVIGATE_NEXT_UNREAD_HIT = 37; { NEXT, but only 'unread' and FTSearch 'hit' entries }
{ (in the SAME ORDER as the hit's relevance ranking) }
const NAVIGATE_PREV_UNREAD_HIT = 38; { PREV, but only 'unread' and FTSearch 'hit' entries }
{ (in the SAME ORDER as the hit's relevance ranking) }
const NAVIGATE_NEXT_CATEGORY = 41; { NEXT, but only 'category' entries }
const NAVIGATE_PREV_CATEGORY = 42; { PREV, but only 'category' entries }
const NAVIGATE_NEXT_NONCATEGORY = 43; { NEXT, but only 'non-category' entries }
const NAVIGATE_PREV_NONCATEGORY = 44; { PREV, but only 'non-category' entries }

const NAVIGATE_MASK = $007; { Navigator code (see above) }


{ Flag which can be used with ALL navigators which causes the navigation }
{to be limited to entries at a specific level (specified by the }
{field "MinLevel" in the collection position) or any higher levels }
{but never a level lower than the "MinLevel" level. Note that level 0 }
{means the top level of the index, so the term "minimum level" really }
{means the "highest level" the navigation can move to. }
{This can be used to find all entries below a specific position }
{in the index, limiting yourself only to that subindex, and yet be }
{able to use any of the navigators to move around within that subindex. }
{This feature was added in Version 4 of Notes, so it cannot be used }
{with earlier Notes Servers. }

const NAVIGATE_MINLEVEL = $0100; { Honor 'Minlevel' field in position }
const NAVIGATE_MAXLEVEL = $0200; { Honor 'Maxlevel' field in position }


{ Flag which can be combined with any navigation directive to prevent }
{having a navigation (Skip) failure abort the (ReadEntries) operation. }
{This is used by VIEW when getting the entries to display in the view, }
{so that if an attempt is made to skip past either end of the index }
{(e.g. using PageUp/PageDown), the skip will be left at the end of the }
{index, and the return will return whatever can be returned using the }
{separate return navigator. It is also used when VIEW attempts to get }
{the "last" N entries: it uses a SkipCount of MAXWORD, and a return }
{navigator of NAVIGATE_ALL_PREVEXPANDED and a return count of N. }

const NAVIGATE_CONTINUE = $8000; { 'Return' even if 'Skip' error }

{ Structure which describes statistics about the overall collection, }
{and can be requested using the READ_MASK_COLLECTIONSTATS flag. If }
{requested, this structure is returned at the beginning of the returned }
{ReadEntries buffer. }


type
  CollectionStats16 = packed record
    TopLevelEntries:word;
    spare: Array[0..3-1] of Word;
end {_1};


type
  CollectionStats = packed record
   TopLevelEntries: dword;
   spare: dword;
end;

{ Structure which specifies collection index position. }

const MAXTUMBLERLEVELS_V2 = 8; { Max. levels in hierarchy tree in V2 }
const MAXTUMBLERLEVELS = 32; { Max. levels in hierarchy tree }


type
  CollectionPosition16 = packed record
    Level: word; { (top level = 0) }
    Tumbler: Array[0..MAXTUMBLERLEVELS_V2-1] of Word;
  end {_3};
  CollectionPosition = packed record
    Level: word;
    MinLevel: BYTE; { MINIMUM level that this position is allowed to be nagivated to. }
          { This is useful to navigate a  subtree using all navigator codes. }
          { This field is IGNORED unless the NAVIGATE_MINLEVEL flag is  enabled (for backward compat) }
    MaxLevel: BYTE;
{ MAXIMUM level that this position }
{ is allowed to be nagivated to. }
{ This is useful to navigate a }
{ subtree using all navigator codes. }
{ This field is IGNORED unless }
{ the NAVIGATE_MAXLEVEL flag is }
{ enabled (for backward compat) }
    Tumbler: Array[0..MAXTUMBLERLEVELS-1] of LongInt;
{Current tumbler (1.2.3, etc) }
{ (an array of ordinal ranks) }
{ (0th entry = top level) }
{ Actual number of array entries }
{ is Level+1 }
end;
PCOLLECTIONPOSITION = ^COLLECTIONPOSITION;

{Macro which computes size of portion of COLLECTIONPOSITION structure }
{which is actually used. This is the size which is returned by }
{NIFReadEntries when READ_MASK_INDEXPOSITION is specified. }

//const COLLECTIONPOSITIONSIZE16(p) = (sizeof(WORD) * ((p)->Level+2));
//const COLLECTIONPOSITIONSIZE(p) = (sizeof(DWORD) * ((p)->Level+2));

{ NIFReadEntries return mask flags }

{These flags specified what information is returned in the return }
{buffer. With the exception of READ_MASK_COLLECTIONSTATS, the }
{information which corresponds to each of the flags in this mask }
{are returned in the buffer, repeated for each index entry, in the }
{order in which the bits are listed here. }

{The return buffer consists of: }

{1) COLLECTIONSTATS structure, if requested (READ_MASK_COLLECTIONSTATS). }
{This structure is returned only once at the beginning of the }
{buffer, and is not repeated for each index entry. }

{2) Information about each index entry. Each flag requested a different }
{bit of information about the index entry. If more than one flag }
{is defined, the values follow each other, in the order in which }
{the bits are listed here. This portion repeats for as many }
{index entries as are requested. }
{ }

{ Fixed length stuff }
const READ_MASK_NOTEID = $00000001; { NOTEID of entry }
const READ_MASK_NOTEUNID = $00000002; { UNID of entry }
const READ_MASK_NOTECLASS = $00000004; { WORD of 'note class' }
const READ_MASK_INDEXSIBLINGS = $00000008; { DWORD/WORD of # siblings of entry }
const READ_MASK_INDEXCHILDREN = $00000010; { DWORD/WORD of # direct children of entry }
const READ_MASK_INDEXDESCENDANTS = $00000020; { DWORD/WORD of # descendants below entry }
const READ_MASK_INDEXANYUNREAD = $00000040; { WORD of TRUE if 'unread' or }
{ "unread" descendants; else FALSE }
const READ_MASK_INDENTLEVELS = $00000080; { WORD of # levels that this }
{entry should be indented in }
{a formatted view. }
{For category entries: }
{# sub-levels that this }
{category entry is within its }
{Collation Descriptor. Used }
{for multiple-level category }
{columns (backslash-delimited). }
{"0" for 1st level in this column, etc. }
{For response entries: }
{# levels that this response }
{is below the "main note" level. }
{For normal entries: 0 }
const READ_MASK_SCORE = $00000200; { Relavence 'score' of entry }
{(only used with FTSearch). }
const READ_MASK_INDEXUNREAD = $00000400; { WORD of TRUE if this entry (only) 'unread' }


{ Stuff returned only once at beginning of return buffer }
const READ_MASK_COLLECTIONSTATS = $00000100; { Collection statistics (COLLECTIONSTATS/COLLECTIONSTATS16) }


{ Variable length stuff }
const READ_MASK_INDEXPOSITION = $00004000; { Truncated COLLECTIONPOSITION/COLLECTIONPOSITION16 }
const READ_MASK_SUMMARYVALUES = $00002000; { Summary buffer w/o item names }
const READ_MASK_SUMMARY = $00008000; { Summary buffer with item names }

{ Structures which are used by NIFGetCollectionData to return data }
{about the collection. NOTE: If the COLLECTIONDATA structure changes, }
{nifods.c must change as well. }

{Definitions which are used by NIFGetCollectionData to return data about the collection. }

const PERCENTILE_COUNT = 11;

const PERCENTILE_0 = 0;
const PERCENTILE_10 = 1;
const PERCENTILE_20 = 2;
const PERCENTILE_30 = 3;
const PERCENTILE_40 = 4;
const PERCENTILE_50 = 5;
const PERCENTILE_60 = 6;
const PERCENTILE_70 = 7;
const PERCENTILE_80 = 8;
const PERCENTILE_90 = 9;
const PERCENTILE_100 = 10;


type
  CollectionData = packed record
          DocCount: LongInt; {Total number of documents in the collection }
          DocTotalSize: LongInt; {Total number of bytes occupied by the documents in the collection }
          BTreeLeafNodes: LongInt; {Number of B-Tree leaf nodes for this index. }
          BTreeDepth: Word; {Number of B-tree levels for this index. }
          Spare: Word; {Unused }
          KeyOffset: Array[0..PERCENTILE_COUNT-1] of LongInt;
{Offset of ITEM_VALUE_TABLE for each 10th-percentile key value }
{A series of ITEM_VALUE_TABLEs follows this structure. }
end {_5};


{ NIFUpdateFilters "modified" flags }

const FILTER_UNREAD = $0001; { UnreadList has been modified }
const FILTER_COLLAPSED = $0002; { CollpasedList has been modified }
const FILTER_SELECTED = $0004; { SelectedList has been modified }
const FILTER_UNID_TABLE = $0008; { UNID table has been modified. }


{ Flag in index entry's NOTEID to indicate (ghost) "category entry" }
{ Note: this relies upon the fact that NOTEID_RESERVED is high bit! }

const NOTEID_CATEGORY = $80000000; { Bit 31 -> (ghost) 'category entry' }
const NOTEID_CATEGORY_TOTAL = $C0000000; { Bit 31+30 -> (ghost) 'grand total entry' }
const NOTEID_CATEGORY_INDENT = $3F000000; { Bits 24-29 -> category indent level within this column }
const NOTEID_CATEGORY_ID = $00FFFFFF; { Low 24 bits are unique category # }


{ SignalFlags word returned by NIFReadEntries and V4+ NIFFindByKey }

const SIGNAL_DEFN_ITEM_MODIFIED = $0001;
{At least one of the "definition" }
{view items ($FORMULA, $COLLATION, }
{or $FORMULACLASS) has been modified }
{by another user since last ReadEntries. }
{Upon receipt, you may wish to }
{re-read the view note if up-to-date }
{copies of these items are needed. }
{Upon receipt, you may also wish to }
{re-synchronize your index position }
{and re-read the rebuilt index. }
{Signal returned only ONCE per detection }
const SIGNAL_VIEW_ITEM_MODIFIED = $0002;
{At least one of the non-"definition" }
{view items ($TITLE,etc) has been }
{modified since last ReadEntries. }
{Upon receipt, you may wish to }
{re-read the view note if up-to-date }
{copies of these items are needed. }
{Signal returned only ONCE per detection }
const SIGNAL_INDEX_MODIFIED = $0004;
{Collection index has been modified }
{by another user since last ReadEntries. }
{Upon receipt, you may wish to }
{re-synchronize your index position }
{and re-read the modified index. }
{Signal returned only ONCE per detection }
const SIGNAL_UNREADLIST_MODIFIED = $0008;
{Unread list has been modified }
{by another window using the same }
{hCollection context }
{Upon receipt, you may wish to }
{repaint the window if the window }
{contains the state of unread flags }
{(This signal is never generated }
{by NIF - only unread list users) }
const SIGNAL_DATABASE_MODIFIED = $0010;
{Collection is not up to date }
const SIGNAL_MORE_TO_DO = $0020;
{ End of collection has not been reached }
{ due to buffer being too full. }
{ The ReadEntries should be repeated }
{ to continue reading the desired entries. }
const SIGNAL_VIEW_TIME_RELATIVE = $0040;
{ The view contains a time-relative formula }
{ (e.g., @Now). Use this flag to tell if the }
{ collection will EVER be up-to-date since }
{ time-relative views, by definition, are NEVER }
{ up-to-date. }
const SIGNAL_NOT_SUPPORTED = $0080;
{Returned if signal flags are not supported }
{This is used by NIFFindByKeyExtended when it }
{is talking to a pre-V4 server that does not }
{support signal flags for FindByKey }

{Mask that defines all "sharing conflicts", which are cases when }
{the database or collection has changed out from under the user. }

const SIGNAL_ANY_CONFLICT = (SIGNAL_DEFN_ITEM_MODIFIED Or
                             SIGNAL_VIEW_ITEM_MODIFIED OR
                             SIGNAL_INDEX_MODIFIED OR
                             SIGNAL_DATABASE_MODIFIED);
const SIGNAL_ANY_NONDATA_CONFLICT = (SIGNAL_DEFN_ITEM_MODIFIED or
                                     SIGNAL_VIEW_ITEM_MODIFIED OR
                                     SIGNAL_INDEX_MODIFIED OR
                                     SIGNAL_UNREADLIST_MODIFIED);


const FIND_PARTIAL = $0001; { Match only initial characters }
{ ("T" matches "Tim") }
const FIND_CASE_INSENSITIVE = $0002; { Case insensitive }
{ ("tim" matches "Tim") }
const FIND_RETURN_DWORD = $0004;        { Input/Output is DWORD COLLECTIONPOSITION }
const FIND_ACCENT_INSENSITIVE = $0008;  { Accent insensitive (ignore diacritical marks }
const FIND_UPDATE_IF_NOT_FOUND = $0020; { If key is not found, update collection }
{ and search again }

{ At most one of the following four flags should be specified }
const FIND_LESS_THAN = $0040;    { Find last entry less than the key value }
const FIND_FIRST_EQUAL = $0000;  { Find first entry equal to the key value (if more than one) }
const FIND_LAST_EQUAL = $0080;   { Find last entry equal to the key value (if more than one) }
const FIND_GREATER_THAN = $00C0; { Find first entry greater than the key value }
const FIND_COMPARE_MASK = $00C0; { Bitmask of the comparison flags defined above }


{NIF public entry points }


function NIFOpenCollection(hViewDB: DBHANDLE;
                           hDataDB: DBHANDLE;
                           ViewNoteID: NOTEID;
                           OpenFlags: Word;
                           hUnreadList: THandle;
                           var rethCollection: HCOLLECTION;
                           rethViewNote: PNOTEHANDLE;
                           retViewUNID: PUNID;
                           rethCollapsedList: PHandle;
                           rethSelectedList:  PHANDLE): STATUS; stdcall; far;

function NIFCloseCollection(hCollection: HCOLLECTION): STATUS; stdcall; far;
function NIFUpdateCollection(hCollection: HCOLLECTION): STATUS; stdcall; far;


function NIFReadEntries(hCollection: HCOLLECTION;
                        IndexPos: PCOLLECTIONPOSITION;
                        SkipNavigator: Word;
                        SkipCount: DWORD;
                        ReturnNavigator: Word;
                        ReturnCount: DWORD;
                        ReturnMask: DWORD;
                        rethBuffer: PHandle;
                        retBufferLength: PWord;
                        retNumEntriesSkiped: pdword;
                        var retNumEntriesReturned: dword;
                        var retSignalFlags: word): STATUS; stdcall; far;

function NIFSetCollation(hCollection: HCOLLECTION;
                         CollationNum: Word): STATUS; stdcall; far;


function NIFFindByKey(hCollection: HCOLLECTION;
                      KeyBuffer: Pointer;
                      FindFlags: Word;
                      retIndexPos: PCOLLECTIONPOSITION;
                      retNumMatches: PDWORD): STATUS; stdcall; far;

function NIFFindByName(hCollection: HCOLLECTION;
                       Name: PChar;
                       FindFlags: Word;
                       retIndexPos: PCOLLECTIONPOSITION;
                       retNumMatches: PLongInt): STATUS; stdcall; far;


function NIFFindDesignNote(hFile: DBHANDLE;
                           Name: PChar;
                           aClass: Word;
                           retNoteID: PNOTEID): STATUS; stdcall; far;

function NIFFindView(hFile:DBHANDLE; Name: PChar;retNoteID: PNoteId): Status;

function NIFFindDesignNoteByName (hFile: DBHandle; Name: PChar; retNoteID:PNoteId):Status;

function NIFFindPrivateDesignNote(hFile: DBHANDLE;
                                  Name: PChar;
                                  aClass: Word;
                                  retNoteID: PNOTEID): STATUS; stdcall; far;
function NIFFindPrivateView(hFile: DbHandle;Name: PChar; retNoteID: PNoteId): status;

function NIFGetCollectionData(hCollection: HCOLLECTION;
                              rethCollData: PHandle): STATUS; stdcall; far;

procedure NIFGetLastModifiedTime(hCollection: HCOLLECTION;
                                 retLastModifiedTime: PTIMEDATE); stdcall; far;


(******************************************************************************)
{OsTime.h}
(******************************************************************************)

{OS time/date package }

procedure OSCurrentTIMEDATE(retTimeDate: PTIMEDATE); stdcall; far;

procedure OSCurrentTimeZone(retZone: PInteger;
                            retDST: PInteger); stdcall; far;

(******************************************************************************)
{  TextList.h}
(******************************************************************************)
{ Text list functions. }

function ListAllocate(ListEntries: Word;
                      TextSize: Word;
                      fPrefixDataType: Bool;
                      rethList: PHandle;
                      retpList: Pointer;
                      retListSize: PWord): STATUS; stdcall; far;


function ListAddText(pList: Pointer;
                     fPrefixDataType: Bool;
                     EntryNumber: Word;
                     Text: PChar;
                     TextSize: Word): STATUS; stdcall; far;



function ListGetText(pList: Pointer;
                     fPrefixDataType: Bool;
                     EntryNumber: Word;
                     retTextPointer: PPChar;
                     retTextLength: PWord): STATUS; stdcall; far;




function ListRemoveEntry(hList: THandle;
                         fPrefixDataType: Bool;
                         pListSize: PWord;
                         EntryNumber: Word): STATUS; stdcall; far;


function ListRemoveAllEntries(hList: THandle;
                              fPrefixDataType: Bool;
                              pListSize: PWord): STATUS; stdcall; far;



function ListAddEntry(hList: THandle;
                      fPrefixDataType: Bool;
                      pListSize: PWord;
                      EntryNumber: Word;
                      Text: PChar;
                      TextSize: Word): STATUS; stdcall; far;



function ListGetSize(pList: Pointer;
                     fPrefixDataType: Bool): Word; stdcall; far;


function ListDuplicate(var pInList: LIST;
                       fNoteItem: Bool;
                       phOutList: PHandle): STATUS; stdcall; far;


function ListGetNumEntries(vList: Pointer;
                           NoteItem: Bool): Word; stdcall; far;

(******************************************************************************)
{nsfdb.h}
(******************************************************************************)

{Note Storage File Database Definitions }

{NSF File Information Buffer size. This buffer is defined to contain }
{Text (host format) that is NULL-TERMINATED. This is the ONLY null-terminated }
{field in all of NSF. }

const NSF_INFO_SIZE = 128;

{Define NSFDbOpenExtended option bits. These bits select individual }
{open options. }

const DBOPEN_WITH_SCAN_LOCK = $0001; { Open with scan lock to prevent}

      DBOPEN_PURGE = $0002;
      DBOPEN_NO_USERINFO = $0004;
      DBOPEN_FORCE_FIXUP = $0008;
      DBOPEN_FIXUP_FULL_NOTE_SCAN = $0010;
      DBOPEN_FIXUP_NO_NOTE_DELTE = $0020;
      DBOPEN_CLUSTER_FAILOVER = $0080;
      DBOPEN_CLOSE_SESS_ON_ERROR = $0100;
      DBOPEN_NOLOG = $0200;

const DBCLASS_BY_EXTENSION = 0; { automatically figure it out }

const DBCLASS_NSFTESTFILE = $ff00;
const DBCLASS_NOTEFILE = $ff01;
const DBCLASS_DESKTOP = $ff02;
const DBCLASS_NOTECLIPBOARD = $ff03;
const DBCLASS_TEMPLATEFILE = $ff04;
const DBCLASS_GIANTNOTEFILE = $ff05;
const DBCLASS_HUGENOTEFILE = $ff06;
const DBCLASS_ONEDOCFILE = $ff07; { Not a mail message }
const DBCLASS_V2NOTEFILE = $ff08;
const DBCLASS_ENCAPSMAILFILE = $ff09; { Specifically used by alt mail }
const DBCLASS_LRGENCAPSMAILFILE = $ff0a; { Specifically used by alt mail }
const DBCLASS_V3NOTEFILE = $ff0b;
const DBCLASS_OBJSTORE = $ff0c; { Object store }
const DBCLASS_V3ONEDOCFILE = $ff0d;

const DBCLASS_MASK = $00ff;
const DBCLASS_VALID_MASK = $ff00;

{Define NSF Special Note ID Indices. The first 16 of these are reserved }
{for "default notes" in each of the 16 note classes. In order to access }
{these, use SPECIAL_ID_NOTE+NOTE_CLASS_XXX. This is generally used }
{when calling NSFDbGetSpecialNoteID. NOTE: NSFNoteOpen, NSFDbReadObject }
{and NSFDbWriteObject support reading special notes or objects directly }
{(without calling NSFDbGetSpecialNoteID). They use a DIFFERENT flag }
{with a similar name: NOTE_ID_SPECIAL (see nsfnote.h). Remember this }
{rule: }

{SPECIAL_ID_NOTE is a 16 bit mask and is used as a NoteClass argument. }
{NOTE_ID_SPECIAL is a 32 bit mask and is used as a NoteID or RRV argument. }
{ }

const SPECIAL_ID_NOTE = $8000; { use in combination w/NOTE_CLASS}


function NSFDbGetOptions(hDB: DBHANDLE;
                         retDbOptions: PLongInt): STATUS; stdcall; far;

function NSFDbSetOptions(hDB: DBHANDLE;
                         DbOptions: LongInt;
                         Mask: LongInt): STATUS; stdcall; far;

const DBOPTION_FT_INDEX = $00000001; { Enable full text indexing }
const DBOPTION_IS_OBJSTORE = $00000002; { TRUE if database is being used}
const DBOPTION_UNIFORM_ACCESS = $00000020; { TRUE if uniform access control}
const DBOPTION_OUT_OF_SERVICE = $00000400; { TRUE is db is out-of-service, no new opens allowed,}
const DBOPTION_MARKED_FOR_DELETE = $00001000; { TRUE if db is marked for delete. no new opens allowed,}


(*type
  ITEM_DEFINITION_TABLE = packed record
    Length: Word; {total length of this buffer }
    Items: Word; {number of items in the table now
                  come the ITEM_DEFINITION structures
                  now comes the packed text }
  end;
  ITEM_DEFINITION = packed record
    Spare: Word; {unused }
    ItemType: Word; {default data type of the item }
    NameLength: Word;{ length of the item's name }
  end {_2};
*)

{ Define NSF DB open modes }

const DB_LOADED = 1; { hDB refers to a normal database file }
const DB_DIRECTORY = 2; { hDB refers to a 'directory' and not a file }

{ Define argument to NSFDbInfoParse/Modify to manipulate components from DbInfo }

const INFOPARSE_TITLE = 0;
const INFOPARSE_CATEGORIES = 1;
const INFOPARSE_CLASS = 2;
const INFOPARSE_DESIGN_CLASS = 3;

{ Option flags for NSFDbCreateExtended }

const DBCREATE_LOCALSECURITY = $0001;
const DBCREATE_OBJSTORE_NEVER = $0002;
const DBCREATE_MAX_SPECIFIED = $0004;

{ Values for EncryptStrength of NSFDbCreateExtended }

const DBCREATE_ENCRYPT_NONE = $00;
const DBCREATE_ENCRYPT_SIMPLE = $01;
const DBCREATE_ENCRYPT_MEDIUM = $02;
const DBCREATE_ENCRYPT_STRONG = $03;
const DBCOPY_REPLICA = $00000001;
const DBCOPY_SUBCLASS_TEMPLATE = $00000002;
const DBCOPY_DBINFO2 = $00000004;
const DBCOPY_SPECIAL_OBJECTS = $00000008;
const DBCOPY_NO_ACL = $00000010;
const DBCOPY_NO_FULLTEXT = $00000020;
const DBCOPY_ENCRYPT_SIMPLE = $00000040;
const DBCOPY_ENCRYPT_MEDIUM = $00000080;
const DBCOPY_ENCRYPT_STRONG = $00000100;
const DBCOPY_KEEP_NOTE_MODTIME = $00000200;


function NSFDbCopyExtended(hSrcDB: DBHANDLE;
                           hDstDB: DBHANDLE;
                           Since: TIMEDATE;
                           NoteClassMask: Word;
                           Flags: LongInt;
                           retUntil: PTIMEDATE): STATUS; stdcall; far;

function NSFDbOpen(PathName: PChar;
                   rethDB: PDBHANDLE): STATUS; stdcall; far;

function NSFDbOpenExtended(PathName: PChar;
                           Options: Word;
                           hNames: THandle;
                           ModifiedTime: PTIMEDATE;
                           rethDB: PDBHANDLE;
                           retDataModified: PTIMEDATE;
                           retNonDataModified: PTIMEDATE): STATUS; stdcall; far;

function NSFDbClose(hDB: DBHANDLE): STATUS; stdcall; far;

function NSFDbCreate(PathName: PChar;
                     DbClass: USHORT;
                     ForceCreation: Bool): STATUS; stdcall; far;

function NSFDbCreateObjectStore(PathName: PChar;
                                ForceCreation: Bool): STATUS; stdcall; far;

function NSFDbDelete(PathName: PChar): STATUS; stdcall; far;

function NSFDbCreateExtended(PathName: PChar;
                             DbClass: Word;
                             ForceCreation: Bool;
                             Options: Word;
                             EncryptStrength: BYTE;
                             MaxFileSize: LongInt): STATUS; stdcall; far;

function NSFDbCopy(hSrcDB: DBHANDLE;
                   hDstDB: DBHANDLE;
                   Since: TIMEDATE;
                   NoteClassMask: Word): STATUS; stdcall; far;

function NSFDbCopyNote(hSrcDB: DBHANDLE;
                       SrcDbID: PDBID;
                       SrcReplicaID: PDBID;
                       SrcNoteID: NOTEID;
                       hDstDB: DBHANDLE;
                       DstDbID: PDBID;
                       DstReplicaID: PDBID;
                       retDstNoteID: PNOTEID;
                       retNoteClass: PWord): STATUS; stdcall; far;


function NSFDbCreateAndCopy(srcDb: PChar;
                            dstDb: PChar;
                            NoteClass: Word;
                            limit: Word;
                            flags: LongInt;
                            retHandle: PDBHANDLE): STATUS; stdcall; far;

function NSFDbMarkForDelete(dbPathPtr: PChar): STATUS; stdcall; far;

function NSFDbMarkInService(dbPathPtr: PChar): STATUS; stdcall; far;

function NSFDbMarkOutOfService(dbPathPtr: PChar): STATUS; stdcall; far;


function NSFDbCopyACL(hSrcDB: DBHANDLE;
                      hDstDB: DBHANDLE): STATUS; stdcall; far;

function NSFDbCopyTemplateACL(hSrcDB: DBHANDLE;
                              hDstDB: DBHANDLE;
                              Manager: PChar;
                              DefaultAccessLevel: Word): STATUS; stdcall; far;

function NSFDbCreateACLFromTemplate(hNTF: DBHANDLE;
                                    hNSF: DBHANDLE;
                                    Manager: PChar;
                                    DefaultAccess: Word;
                                    rethACL: PHandle): STATUS; stdcall; far;

function NSFDbStoreACL(hDB: DBHANDLE;
                       hACL: THandle;
                       ObjectID: LongInt;
                       Method: Word): STATUS; stdcall; far;

function NSFDbReadACL(hDB: DBHANDLE;
                      rethACL: PHandle): STATUS; stdcall; far;

function NSFDbGenerateOID(hDB: DBHANDLE;
                          retOID: POID): STATUS; stdcall; far;

function NSFDbModifiedTime(hDB: DBHANDLE;
                           retDataModified: PTIMEDATE;
                           retNonDataModified: PTIMEDATE): STATUS; stdcall; far;

function NSFDbPathGet(hDB: DBHANDLE;
                      retCanonicalPathName: PChar;
                      retExpandedPathName: PChar): STATUS; stdcall; far;

function NSFDbInfoGet(hDB: DBHANDLE;
                      retBuffer: PChar): STATUS; stdcall; far;

function NSFDbInfoSet(hDB: DBHANDLE;
                      Buffer: PChar): STATUS; stdcall; far;

procedure NSFDbInfoParse(Info: PChar;
                         What: Word;
                         Buffer: PChar;
                         Length: Word); stdcall; far;

procedure NSFDbInfoModify(Info: PChar;
                          What: Word;
                          Buffer: PChar); stdcall; far;

function NSFDbGetSpecialNoteID(hDB: DBHANDLE;
                               Index: Word;
                               retNoteID: PNOTEID): STATUS; stdcall; far;

function NSFDbIDGet(hDB: DBHANDLE;
                    retDbID: PDBID): STATUS; stdcall; far;

function NSFDbReplicaInfoGet(hDB: DBHANDLE;
                             retReplicationInfo: PDBREPLICAINFO): STATUS; stdcall; far;

function NSFDbReplicaInfoSet(hDB: DBHANDLE;
                             ReplicationInfo: PDBREPLICAINFO): STATUS; stdcall; far;


function NSFDbGetNoteInfo(hDB: DBHANDLE;
                          NoteID: NOTEID;
                          retNoteOID: POID;
                          retModified: PTIMEDATE;
                          retNoteClass: PWord): STATUS; stdcall; far;

function NSFDbGetNoteInfoByUNID(hDB: THandle;
                                pUNID: PUNID;
                                retNoteID: PNOTEID;
                                retOID: POID;
                                retModTime: PTIMEDATE;
                                retClass: PWord): STATUS; stdcall; far;

function NSFDbGetModifiedNoteTable(hDB: DBHANDLE;
                                   NoteClassMask: Word;
                                   Since: TIMEDATE;
                                   retUntil: PTIMEDATE;
                                   rethTable: PHandle): STATUS; stdcall; far;

function NSFApplyModifiedNoteTable(hModifiedNotes: THandle;
                                   hTargetTable: THandle): STATUS; stdcall; far;

function NSFDbLocateByReplicaID(hDB: DBHANDLE;
                                ReplicaID: PDBID;
                                retPathName: PChar;
                                PathMaxLen: Word): STATUS; stdcall; far;


function NSFDbStampNotes(hDB: DBHANDLE;
                         hTable: THandle;
                         ItemName: PChar;
                         ItemNameLength: Word;
                         Data: Pointer;
                         DataLength: Word): STATUS; stdcall; far;

function NSFDbDeleteNotes(hDB: DBHANDLE;
                          hTable: THandle;
                          retUNIDArray: PUNID): STATUS; stdcall; far;


procedure NSFDbAccessGet(hDB: THandle;
                         retAccessLevel: PWord;
                         retAccessFlag: PWord); stdcall; far;

function NSFDbClassGet(hDB: DBHANDLE;
                       retClass: PWord): STATUS; stdcall; far;

function NSFDbModeGet(hDB: DBHANDLE;
                      retMode: PUSHORT): STATUS; stdcall; far;

function NSFDbCloseSession(hDB: DBHANDLE): STATUS; stdcall; far;

function NSFDbReopen(hDB: DBHANDLE;
                     rethDB: PDBHANDLE): STATUS; stdcall; far;

function NSFDbMajorMinorVersionGet(hDB: DBHANDLE;
                                   retMajorVersion: PWord;
                                   retMinorVersion: PWord): STATUS; stdcall; far;

function NSFDbItemDefTable(hDB: DBHANDLE;
                           retItemNameTable: PITEMDEFTABLEHANDLE): STATUS; stdcall; far;

function NSFDbGetBuildVersion(hDB: DBHANDLE;
                              retVersion: PWord): STATUS; stdcall; far;

function NSFDbSpaceUsage(hDB: DBHANDLE;
                         retAllocatedBytes: PLongInt;
                         retFreeBytes: PLongInt): STATUS; stdcall; far;

function NSFDbGetOpenDatabaseID(hDB: DBHANDLE): LongInt; stdcall; far;

function NSFGetServerStats(ServerName: PChar;
                           Facility: PChar;
                           StatName: PChar;
                           rethTable: PHandle;
                           retTableSize: PLongInt): STATUS; stdcall; far;

function NSFGetServerLatency(ServerName: PChar;
                             Timeout: LongInt;
                             retClientToServerMS: PLongInt;
                             retServerToClientMS: PLongInt;
                             ServerVersion: PWord): STATUS; stdcall; far;

function NSFRemoteConsole(ServerName: PChar;
                          ConsoleCommand: PChar;
                          hResponseText: PHandle): STATUS; stdcall; far;

function NSFDbUpdateUnread(hDataDB: DBHANDLE;
                           hUnreadList: THandle): STATUS; stdcall; far;


function NSFDbGetUnreadNoteTable(hDB: DBHANDLE;
                                 UserName: PChar;
                                 UserNameLength: Word;
                                 fCreateIfNotAvailable: Bool;
                                 rethUnreadList: PHandle): STATUS; stdcall; far;

function NSFDbGetUnreadNoteTable2(hDB: DBHANDLE;
                                 UserName: PChar;
                                 UserNameLength: Word;
                                 fCreateIfNotAvailable: Bool;
                                 fUpdateUnread: Bool;
                                 rethUnreadList: PHandle): STATUS; stdcall; far;


function NSFDbSetUnreadNoteTable(hDB: DBHANDLE;
                                 UserName: PChar;
                                 UserNameLength: Word;
                                 fFlushToDisk: Bool;
                                 hOriginalUnreadList: THandle;
                                 hUnreadList: THandle): STATUS; stdcall; far;

function NSFDbGetObjectStoreID(dbhandle: DBHANDLE;
                               Specified: PBool;
                               ObjStoreReplicaID: PDBID): STATUS; stdcall; far;

function NSFDbSetObjectStoreID(dbhandle: DBHANDLE;
                               ObjStoreReplicaID: PDBID): STATUS; stdcall; far;

function NSFDbFilter(hFilterDB: DBHANDLE;
                     hFilterNote: NOTEHANDLE;
                     hNotesToFilter: THandle;
                     fIncremental: Bool;
                     Reserved1: Pointer;
                     Reserved2: Pointer;
                     DbTitle: PChar;
                     ViewTitle: PChar;
                     Reserved3: Pointer;
                     Reserved4: Pointer;
                     hDeletedList,HSelectedList: THandle): STATUS; stdcall; far;

function NSFDbCompact (PathName: PChar; Options: word; var RetStats: dword):Status; stdcall; far;

type
  DBQUOTAINFO = packed record
    WarningThreshold: DWord; { Database size warning threshold in kbyte units }
    SizeLimit: DWord;        { Database size limit in kbyte units }
    CurrentDbSize: DWord;    { Current size of database (in kbyte units) }
    MaxDbSize: DWord;        { Max database file size possible (in kbyte units) }
  end;
  PDBQuotaInfo = ^DBQuotaInfo;

function NSFDbQuotaGet(Filename: PChar;
                            retQuotaInfo: PDBQUOTAINFO): STATUS; stdcall; far;

(******************************************************************************)
{event.h}
(******************************************************************************)

function EventQueueAlloc(QueueName: PChar): STATUS; stdcall; far;

{EventQueueAlloc - Create an event queue with the given name. If one }
{already exists, return an error. }

{Inputs: }
{QueueName - ASCIIZ name of queue to create (32 chars including NULL MAX) }

{Outputs: }
{If queue with that name does not already exist, creates the queue, }
{else returns error. }



{Each event consumer calls EventQueueAlloc at startup to create a event }
{queue with a specific name to receive events. }




procedure EventQueueFree(QueueName: PChar); stdcall; far;

{EventQueueFree - destroys the queue and deallocates the memory it used. }

{Inputs: }
{QueueName - ASCIIZ name of queue to destroy }

{Outputs: }
{none }


{Called at shutdown time by each event consumer. }




function EventQueuePut(QueueName: PChar;
                       OriginatingServer: PChar;
                       aType: Word;
                       Severity: Word;
                       EventTime: PTIMEDATE;
                       FormatSpecifier: Word;
                       EventDataLength: Word;
                       EventSpecificData: Pointer): STATUS; stdcall; far;

{EventQueuePut - puts a event into a queue. }


{Inputs: }
{QueueName - (ASCIIZ) name of queue to receive this event }
{OriginatingServer - (ASCIIZ) name of server where event }
{occured (if, NULL, uses the current server name) }
{type - one of: EVT_COMM }
{EVT_SECURITY }
{EVT_MAIL }
{EVT_RESOURCE }
{EVT_MISC }
{EVT_ALARM }
{EVT_SERVER }
{EVT_UNKNOWN }

{Severity - one of: }
{SEV_FATAL }
{SEV_FAILURE }
{SEV_WARNING1 }
{SEV_WARNING2 }
{SEV_NORMAL }
{SEV_UNKNOWN }

{EventTime - event's temporal locus }
{FormatSpecifier - format of data in EventSpecificData }
{EventDataLength - number of bytes in EventSpecificData }
{EventSpecificData - event info }

{Outputs: }
{Event is placed in the specified queue. }
{(routine status) }



{Event producers call this routine whenever an event occurs that }
{anyone may be interested in. If no event consumer has requested }
{notification of a particular event, the event is discarded. }
{ }


function EventQueueGet(QueueName: PChar;
                       rethEvent: PHandle): STATUS; stdcall; far;

{EventQueueGet - removes an event from a queue and returns the }
{handle to it's object. It is the caller's responsibility }
{to free it when through. }

{Inputs: }
{QueueName - name of queue to search for events }

{Outputs: }
{*hEvent - handle to event object. NULLHANDLE if queue is empty. }
{(retstatus) - ERR_EVTQUEUE_EMPTY if empty queue, else NOERROR }
{if something dequeued }


{Event consumers call this routine to dequeue any events }
{presently in their queue. If the queue is empty, the routine }
{returns ERR_EVTQUEUE_EMPTY. Else, it returns NOERROR, and }
{stuffs the output parameter with the handle to the dequeued event. }



function EventRegisterEventRequest(EventType: Word;
                                   EventSeverity: Word;
                                   QueueName: PChar;
                                   DestName: PChar): STATUS; stdcall; far;

{EventRegisterEventRequest }

{Inputs: }
{EventType - type of event to notify of }
{EventSeverity - severity of event to notify of }
{QueueName - name of queue that desires notification }
{DestName - name of person/database to address event to }

{Outputs: }
{(none) }
{routine = status }


{At registration time, an event consumer calls this routine }
{once for each class and type of event that it is interested in. }



function EventDeregisterEventRequest(EventType: Word;
                                     EventSeverity: Word;
                                     QueueName: PChar): STATUS; stdcall; far;

{EventDeregisterEventRequest }

{Inputs: }
{EventType - type of event to discontinue notification of }
{EventSeverity - severity of event to discontinue notification of }
{QueueName - name of queue that desires no longer desires notification }

{Outputs: }
{(none) }
{routine = status }

{Called by process to discontinue notification of particular events }



function EventGetDestName(EventType: Word;
                          Severity: Word;
                          QueueName: PChar;
                          DestName: PChar;
                          DestNameSize: Word): Bool; stdcall; far;
{Inputs: }
{EventType - type of event }
{EventSeverity - severity of event }
{QueueName - name of queue that desires information }
{DestName - buffer to receive name of person/database to address event to }
{DestNameSize - size of ret buffer }

{Outputs: }
{DestName - contains name of destination person/database }
{routine = TRUE if dest name is set }
{Called by a process to obtain the destination for these events for this }
{queue. For mail, would return a user or group name. For logging, would }
{return a database name, or a server and database name, etc. }

{flags for EventQueuePut }
const EVT_UNKNOWN = 0;
const EVT_COMM = 1;
const EVT_SECURITY = 2;
const EVT_MAIL = 3;
const EVT_REPLICA = 4;
const EVT_RESOURCE = 5;
const EVT_MISC = 6;
const EVT_SERVER = 7;
const EVT_ALARM = 8;
const EVT_UPDATE = 9;
const MAX_TYPE = 10;

{event type names }
const UNKNOWN_NAME = 'Unknown';
const COMM_NAME = 'Comm';
const SECURE_NAME = 'Security';
const MAIL_NAME = 'Mail';
const REPLICA_NAME = 'Replica';
const RESOURCE_NAME = 'Resource';
const MISC_NAME = 'Misc';
const SERVER_NAME = 'Server';
const ALARM_NAME = 'Statistic';
const UPDATE_NAME = 'Update';

{Severity FLAGS }
const SEV_UNKNOWN = 0;
const SEV_FATAL = 1;
const SEV_FAILURE = 2;
const SEV_WARNING1 = 3;
const SEV_WARNING2 = 4;
const SEV_NORMAL = 5;
const MAX_SEVERITY = 6;

const FATAL_NAME = 'Fatal';
const FAILURE_NAME = 'Failure';
const WARNING1_NAME = 'Warning (high)';
const WARNING2_NAME = 'Warning (low)';
const NORMAL_NAME = 'Normal';

{FormatSpecifier FLAGS }

const FMT_UNKNOWN = 0;
const FMT_TEXT = 1;
const FMT_ERROR_CODE = 2;
const FMT_ERROR_MSG = 3;

{Version field values specified in following structure }

const EVENT_VERSION = 1;

{Event structure }


type
  EVENT_DATA = packed record
    Links: Array[0..3-1] of LongInt; { Reserved - used to link this struct onto queues }
    OriginatingServerName: Array[0..MAXUSERNAME-1] of Char; { Server name (only if event relayed to another server) }
    Version: Word;      {EVENT_VERSION }
    Spare1: Word;       {Spare - Must be 0 }
    Spare2: Word;       {Spare - Must be 0 }
    Spare3: Word;       {Spare - Must be 0 }
    aType: Word;        {EVT_xxx }
    Severity: Word;     {SEV_xxx }
    EventTime: TIMEDATE;{Time/date event was generated }
    FormatSpecifier: Word; { FMT_xxx (format of event data which follows) }
    EventDataLength: Word; { Length of event data which follows }
    EventSpecificData: BYTE; { (First byte of) Event Data which follows... }
  end;


(******************************************************************************)
{mailserv.h}
(******************************************************************************)

{Mail delivery priorities. Note: order is assumed. }

const DELIVERY_PRIORITY_LOW = 0;
const DELIVERY_PRIORITY_NORMAL = 1;
const DELIVERY_PRIORITY_HIGH = 2;

{Mail delivery report requests. Note: order is assumed. }

const DELIVERY_REPORT_NONE = 0;
const DELIVERY_REPORT_BASIC = 1;
const DELIVERY_REPORT_CONFIRMED = 2;
const DELIVERY_REPORT_TRACE = 3;
const DELIVERY_REPORT_TRACE_NO_DELIVER = 4;
const DELIVERY_REPORT_CONFIRM_NO_DELIVER = 5;

{Mail delivery time constants. }

const DELIVERY_HOUR = 3600;
const DELIVERY_MINUTE = 60;
const DELIVERY_MESSAGE_SIZE = 1024;

{Message types - Returned by MailGetMessageType. Note: order is assumed. }

const MAIL_MESSAGE_UNKNOWN = 0;
const MAIL_MESSAGE_MEMO = 1;
const MAIL_MESSAGE_DELIVERYREPORT = 2;
const MAIL_MESSAGE_NONDELIVERYREPORT = 3;
const MAIL_MESSAGE_RETURNRECEIPT = 4;
const MAIL_MESSAGE_PHONEMESSAGE = 5;
const MAIL_MESSAGE_TRACEREPORT = 6;

{Address file functions }


function MailGetDomainName(Domain: PChar): STATUS; stdcall; far;


function MailLookupAddress(UserName: PChar;
                           MailAddress: PChar): STATUS; stdcall; far;

function MailLookupUser(UserName: PChar;
                        FullName: PChar;
                        MailServerName: PChar;
                        MailFileName: PChar;
                        MailAddress: PChar;
                        ShortName: PChar): STATUS; stdcall; far;

{Message mailing functions }


function MailGetMessageItem(hMessage: THandle;
                            ItemCode: Word;
                            retString: PChar;
                            StringSize: Word;
                            retStringLength: PWord): STATUS; stdcall; far;


function MailGetMessageItemHandle(hMessage: THandle;
                                  ItemCode: Word;
                                  retbhValue: PBLOCKID;
                                  retValueType: PWord;
                                  retValueLength: PLongInt): STATUS; stdcall; far;


function MailGetMessageItemTimeDate(hMessage: THandle;
                                    ItemCode: Word;
                                    retTimeDate: PTIMEDATE): STATUS; stdcall; far;


function MailCreateMessage(hFile: DBHANDLE;
                           rethMessage: PHandle): STATUS; stdcall; far;


function MailAddHeaderItem(hMessage: THandle;
                           ItemCode: Word;
                           Value: PChar;
                           ValueLength: Word): STATUS; stdcall; far;


function MailAddHeaderItemByHandle(hMessage: THandle;
                                   ItemCode: Word;
                                   hValue: THandle;
                                   ValueLength: Word;
                                   ItemFlags: Word): STATUS; stdcall; far;


function MailReplaceHeaderItem(hMessage: THandle;
                               ItemCode: Word;
                               Value: Pointer;
                               ValueLength: Word): STATUS; stdcall; far;


function MailCreateBodyItem(rethBodyItem: PHandle;
                            retBodyLength: PLongInt): STATUS; stdcall; far;

function MailAppendBodyItemLine(hBodyItem: THandle;
                                BodyLength: PLongInt;
                                Text: PChar;
                                TextLength: Word): STATUS; stdcall; far;

function MailAddBodyItem(hMessage: THandle;
                         hBodyItem: THandle;
                         BodyLength: LongInt;
                         CTFName: PChar): STATUS; stdcall; far;


function MailAddRecipientsItem(hMessage: THandle;
                               hRecipientsItem: THandle;
                               RecipientsLength: Word): STATUS; stdcall; far;



function MailTransferMessageLocal(hMessage: THandle): STATUS; stdcall; far;


function MailIsNonDeliveryReport(hMessage: THandle): Bool; stdcall; far;

function MailGetMessageType(hMessage: THandle): Word; stdcall; far;

function MailCloseMessage(hMessage: THandle): STATUS; stdcall; far;


function MailExpandNames(hWorkList: THandle;
                         WorkListSize: Word;
                         hOutputList: PHandle;
                         OutputListSize: PWord;
                         UseExpanded: Bool;
                         hRecipsExpanded: THandle): STATUS; stdcall; far;


function MailLogEvent(Flags: Word;
                      StringID: STATUS;
                      hModule: HMODULE;
                      AdditionalErrorCode: STATUS;
                      _5: dword {Undefined number of parametrs}): STATUS; stdcall; far;


function MailLogEventText(Flags: Word;
                          aString: PChar;
                          hModule: HMODULE;
                          AdditionalErrorCode: STATUS;
                          _5: dword {Undefined number of parametrs}): STATUS; stdcall; far;

{Mail event logging flags }
const MAIL_LOG_TO_MISCEVENTS = $0001; { Log message to Miscellaneuos Events view }
const MAIL_LOG_TO_MAILEVENTS = $0002; { Log message to Mail Events view }
const MAIL_LOG_TO_EVENTS_ONLY = $0004; { Don't log messages to console }
const MAIL_LOG_TO_BOTH = (MAIL_LOG_TO_MAILEVENTS or MAIL_LOG_TO_MISCEVENTS);


{Message attachment handling functions }


function MailGetMessageAttachmentInfo(hMessage: THandle;
                                      Num: Word;
                                      bhItem: PBLOCKID;
                                      FileName: PChar;
                                      FileSize: PLongInt;
                                      FileAttributes: PWord;
                                      FileHostType: PWord;
                                      FileCreated: PTIMEDATE;
                                      FileModified: PTIMEDATE): Bool; stdcall; far;


function MailExtractMessageAttachment(hMessage: THandle;
                                      bhItem: BLOCKID;
                                      FileName: PChar): STATUS; stdcall; far;

function MailAddMessageAttachment(hMessage: THandle;
                                  FileName: PChar;
                                  OriginalFileName: PChar): STATUS; stdcall; far;

{Message file functions }


function MailOpenMessageFile(FileName: PChar;
                             rethFile: PDBHANDLE): STATUS; stdcall; far;

function MailCreateMessageFile(FileName: PChar;
                               TemplateFileName: PChar;
                               Title: PChar;
                               rethFile: PDBHANDLE): STATUS; stdcall; far;

function MailPurgeMessageFile(hFile: DBHANDLE): STATUS; stdcall; far;

function MailCloseMessageFile(hFile: DBHANDLE): STATUS; stdcall; far;

function MailGetMessageFileModifiedTime(hFile: DBHANDLE;
                                        retModifiedTime: PTIMEDATE): STATUS; stdcall; far;

{Message list functions}


function MailCreateMessageList(hFile: DBHANDLE;
                               hMessageList: PHandle;
                               var MessageList: PDARRAY;
                               MessageCount: PWord): STATUS; stdcall; far;

function MailFreeMessageList(hMessageList: THandle;
                             MessageList: PDARRAY): STATUS; stdcall; far;

function MailGetMessageInfo(MessageList: PDARRAY;
                            aMessage: Word;
                            RecipientCount: PWord;
                            Priority: PWord;
                            Report: PWord): STATUS; stdcall; far;

function MailGetMessageSize(MessageList: PDARRAY;
                            aMessage: Word;
                            MessageSize: PLongInt): STATUS; stdcall; far;

function MailGetMessageRecipient(MessageList: PDARRAY;
                                 aMessage: Word;
                                 RecipientNum: Word;
                                 RecipientName: PChar;
                                 RecipientNameSize: Word;
                                 RecipientNameLength: PWord): STATUS; stdcall; far;


function MailDeleteMessageRecipient(MessageList: PDARRAY;
                                    aMessage: Word;
                                    RecipientNum: Word): STATUS; stdcall; far;


function MailGetMessageOriginator(MessageList: PDARRAY;
                                  aMessage: Word;
                                  OriginatorName: PChar;
                                  OriginatorNameSize: Word;
                                  OriginatorNameLength: PWord): STATUS; stdcall; far;


function MailGetMessageOriginatorDomain(MessageList: PDARRAY;
                                        aMessage: Word;
                                        OriginatorDomain: PChar;
                                        OriginatorDomainSize: Word;
                                        OriginatorNameLength: PWord): STATUS; stdcall; far;


function MailOpenMessage(MessageList: PDARRAY;
                         aMessage: Word;
                         hMessage: PHandle): STATUS; stdcall; far;


function MailGetMessageBody(hMessage: THandle;
                            hBody: PHandle;
                            BodyLength: PLongInt): STATUS; stdcall; far;


function MailGetMessageBodyText(hMessage: THandle;
                                ItemName: PChar;
                                LineDelims: PChar;
                                LineLength: Word;
                                ConvertTabs: Bool;
                                OutputFileName: PChar;
                                OutputFileSize: PLongInt): STATUS; stdcall; far;


function MailGetMessageBodyComposite(hMessage: THandle;
                                     ItemName: PChar;
                                     OutputFileName: PChar;
                                     OutputFileSize: PLongInt): STATUS; stdcall; far;


function MailAddMessageBodyText(hMessage: THandle;
                                ItemName: PChar;
                                InputFileName: PChar;
                                FontID: LongInt;
                                LineDelim: PChar;
                                ParaDelim: Word;
                                CTFName: PChar): STATUS; stdcall; far;


function MailAddMessageBodyComposite(hMessage: THandle;
                                     ItemName: PChar;
                                     InputFileName: PChar): STATUS; stdcall; far;


function MailSetMessageLastError(MessageList: PDARRAY;
                                 aMessage: Word;
                                 ErrorText: PChar): STATUS; stdcall; far;


function MailPurgeMessage(MessageList: PDARRAY;
                          aMessage: Word): STATUS; stdcall; far;


function MailSendNonDeliveryReport(MessageList: PDARRAY;
                                   aMessage: Word;
                                   RecipientNums: Word;
                                   RecipientNumList: PWord;
                                   ReasonText: PChar;
                                   ReasonTextLength: Word): STATUS; stdcall; far;


function MailSendDeliveryReport(MessageList: PDARRAY;
                                aMessage: Word;
                                RecipientNums: Word;
                                RecipientNumList: PWord): STATUS; stdcall; far;

{ Mail address to user and domain name parsing function }


function MailParseMailAddress(MailAddress: PChar;
                              MailAddressLength: Word;
                              UserName: PChar;
                              UserNameSize: Word;
                              UserNameLength: PWord;
                              DomainName: PChar;
                              DomainNameSize: Word;
                              DomainNameLength: PWord): STATUS; stdcall; far;


{ Broadcast newmail recieved message }


procedure MailBroadcastNewMail(MessageText: PChar); stdcall; far;
{ V2 Compatible, NETBIOS-ONLY }

{ Routing Table Services }


function MailLoadRoutingTables(hAddressBook: DBHANDLE;
                               LocalServerName: PChar;
                               LocalDomainDomain: PChar;
                               TaskName: PChar;
                               EnableTrace: Bool;
                               EnableDebug: Bool;
                               rethTables: PHandle): STATUS; stdcall; far;


function MailReloadRoutingTables(hTables: THandle;
                                 EnableTrace: Bool;
                                 EnableDebug: Bool;
                                 retAddressBookModified: PBool): STATUS; stdcall; far;


function MailUnloadRoutingTables(hTables: THandle): STATUS; stdcall; far;

{NextHopFlags for MailFindNextHopTo* routines }

const NEXTHOP_INTRANET = $00000001; { Next Hop is on same network }


function MailFindNextHopToDomain(hTables: THandle;
                                 OriginatorsDomain: PChar;
                                 DestDomain: PChar;
                                 NextHopServer: PChar;
                                 NextHopMailbox: PChar;
                                 NextHopFlags: PLongInt;
                                 ErrorServer: PChar): STATUS; stdcall; far;


function MailFindNextHopToServer(hTables: THandle;
                                 DestDomain: PChar;
                                 DestServer: PChar;
                                 NextHopServer: PChar;
                                 NextHopMailbox: PChar;
                                 NextHopFlags: PLongInt;
                                 ActualCost: PWord): STATUS; stdcall; far;

type
  Mail_Routing_Actions = (MAIL_ERROR,
           MAIL_TRANSFER,
           MAIL_DELIVER,
           MAIL_FORWARD );
function MailFindNextHopToRecipient(hTables: THandle;
                                    OriginatorsDomain: PChar;
                                    RecipientAddress: PChar;
                                    var Action: MAIL_ROUTING_ACTIONS;
                                    NextHopServer: PChar;
                                    NextHopMailbox: PChar;
                                    ForwardAddress: PChar;
                                    ErrorText: PChar;
                                    var NextHopFlags: dword): STATUS; stdcall; far;
function MailFindNextHopViaRules(hTables: THandle;
                                 RecipientAddress: PChar;
                                 retDestServer: PChar;
                                 retDestDomain: PChar): STATUS; stdcall; far;
function MailSetDynamicCost(hTables: THandle;
                            Server: PChar;
                            CostBias: SWORD): Bool; stdcall; far;
function MailResetAllDynamicCosts(hTables: THandle): Bool; stdcall; far;

(******************************************************************************)
{ft.h}
(******************************************************************************)
{Public Definitions for Full Text Package }

{Define Indexing options }

const FT_INDEX_REINDEX = $0002; { Re-index from scratch}
const FT_INDEX_CASE_SENS = $0004; { Build case sensitive index}
const FT_INDEX_STEM_INDEX = $0008; { Build stem index }
const FT_INDEX_PSW = $0010; { Index paragraph & sentence breaks}
const FT_INDEX_OPTIMIZE = $0020; { Optimize index (e.g. for CDROM) }
const FT_INDEX_ATT = $0040; { Index Attachments }
const FT_INDEX_ENCRYPTED_FIELDS = $0080; { Index Encrypted Fields }
const FT_INDEX_AUTOOPTIONS = $0100; { Get options from database }

{Define Search options }

const FT_SEARCH_SET_COLL = $00000001; { Store search results in NIF collections;}
const FT_SEARCH_REFINE = $00000004; { Refine the query using the IDTABLE }
const FT_SEARCH_SCORES = $00000008; { Return document scores (default sort) }
const FT_SEARCH_RET_IDTABLE = $00000010; { Return ID table }
const FT_SEARCH_SORT_DATE = $00000020; { Sort results by date }
const FT_SEARCH_SORT_ASCEND = $00000040; { Sort in ascending order }
const FT_SEARCH_TOP_SCORES = $00000080; { Use Limit arg. to return only top scores }
const FT_SEARCH_STEM_WORDS = $00000200; { Stem words in this query }
const FT_SEARCH_THESAURUS_WORDS = $00000400; { Thesaurus words in this query }

{Define search results data structure }

const FT_RESULTS_SCORES = $0001; { Array of scores follows }


type
  FT_SEARCH_RESULTS = packed record
    NumHits: LongInt; {Number of search hits following }
    Flags: Word; {Flags (FT_RESULTS_xxx) }
    Spare: Word; {Followed by an array of NoteIDs Followed by a BYTE array of scores (optional) }
  end;
  PFT_Index_Stats = ^FT_INDEX_STATS;
  FT_INDEX_STATS = packed record
    DocsAdded    :DWORD;
    DocsUpdated  :DWORD;
    DocsDeleted  :DWORD;
    BytesIndexed :DWORD;
  end;



function FTIndex(hDB: THandle;
                 Options: Word;
                 StopFile: PChar;
                 retStats: PFT_INDEX_STATS): STATUS; stdcall; far;

function FTDeleteIndex(hDB: THandle): STATUS; stdcall; far;

function FTGetLastIndexTime(hDB: THandle;
                            retTime: PTIMEDATE): STATUS; stdcall; far;


function FTOpenSearch (rethSearch: PHandle): STATUS; stdcall; far;

function FTSearch(hDB: THandle;
                  phSearch: PHandle;
                  hColl: HCOLLECTION;
                  Query: PChar;
                  Options: LongInt;
                  Limit: Word;
                  hIDTable: THandle;
                  retNumDocs: PLongInt;
                  Reserved: PHandle;
                  rethResults: PHandle): STATUS; stdcall; far;

function FTCloseSearch(hSearch: THandle): STATUS; stdcall; far;

(******************************************************************************)
{idtable.h}
(******************************************************************************)
{ID Table Routines }

{ This package is used to create and manipulate tables that contain }
{ compressed double-word values that typically represent IDs. The }
{primitives allow the caller to create an ID table, add or delete IDs, }
{and query for the presence of an ID. }
{}
{Compression of the table is achieved by virtue of the fact that it }
{is assumed that the ID space is relatively "regular", that is, that }
{ID values differ from each other by some regular value, say 4. }

{ID tables are always stored in Canonical format. }

{(This .H file is global so that the ODS routines can access it; all }
{access to the following structures should be via the programmatic }
{interfaces provided.) }

type
  IdTable = packed record
    Alignment: LongInt; {alignment factor (4 if IDs are 4 apart) }
    IdsPinnedAt64K: Word;
    Entries: word;
    Flags: Word;
{ flags }
    Time: TIMEDATE;
{ time - reserved for use by caller only }
end {_1};


type
  IDEntry = packed record { BYTE Repeat; /* # of IDs AFTER this one that match Alignment */ }
    aRepeat: byte;
    Value: LongInt; {Value of this ID }
end {_2};


const IDTABLE_MODIFIED = $0001; {modified - set by Insert/Delete}
                                { and can be cleared by caller if desired */ }
const IDTABLE_INVERTED = $0002; { sense of list inverted (reserved for use by caller only) */ }


function IDCreateTable(Alignment: LongInt;
                       rethTable: PHandle): STATUS; stdcall; far;

function IDDestroyTable(hTable: THandle): STATUS; stdcall; far;

function IDInsert(hTable: THandle;
                  id: LongInt;
                  retfInserted: PBool): STATUS; stdcall; far;

function IDDelete(hTable: THandle;
                  id: LongInt;
                  retfDeleted: PBool): STATUS; stdcall; far;

function IDDeleteAll(hTable: THandle): STATUS; stdcall; far;

function IDScan(hTable: THandle;
                fFirst: Bool;
                retID: PLongInt): Bool; stdcall; far;
type
  IDENUMERATEPROC = function (aParameter: pointer; anId: dword): status; far;
//  = STATUS (LNCALLBACKPTR IDENUMERATEPROC) (VOID *PARAMETER, DWORD ID);

function IDEnumerate(hTable: THandle;
                     Routine: IDENUMERATEPROC;
                     Parameter: Pointer): STATUS; stdcall; far;

function IDEntries(hTable: THandle): LongInt; stdcall; far;

function IDIsPresent(hTable: THandle;
                     id: LongInt): Bool; stdcall; far;

function IDTableSize(hTable: THandle): LongInt; stdcall; far;

function IDTableCopy(hTable: THandle;
                     rethTable: PHandle): STATUS; stdcall; far;

function IDTableSizeP(pIDTable: Pointer): LongInt; stdcall; far;

function IDTableFlags(pIDTable: Pointer): Word; stdcall; far;

function IDTableTime(pIDTable: Pointer): TIMEDATE; stdcall; far;

procedure IDTableSetFlags(pIDTable: Pointer;
                          Flags: Word); stdcall; far;

procedure IDTableSetTime(pIDTable: Pointer;
                         Time: TIMEDATE); stdcall; far;


// Names for design elements
const
  // common fields
  FIELD_TITLE = '$TITLE';
  FIELD_FORM=   'Form';
  FIELD_TYPE_TYPE=  'type';
  FIELD_LINK  = '$REF';
  FIELD_UPDATED_BY = '$UpdatedBy';
  FIELD_NAMELIST = '$NameList';
  FIELD_NAMED   = '$Name';
  ITEM_NAME_TEMPLATE = '$Body';     { form item to hold form CD }
  ITEM_NAME_DOCUMENT ='$Info';      { document header info }
  ITEM_NAME_TEMPLATE_NAME = FIELD_TITLE;  { form title item }
  ITEM_NAME_FORMLINK = '$FormLinks';    { form link table }
  ITEM_NAME_FIELDS = '$Fields';     { field name table }
  ITEM_NAME_FORMPRIVS = '$FormPrivs'; { form privileges }
  ITEM_NAME_FORMUSERS = '$FormUsers'; { text list of users allowed to use the form }

  // Design flags
  DESIGN_FLAGS = '$Flags';

  // Misc flags - Added by Daniel
  ASSIST_TRIGGER = '$AssistTrigger';
  ASSIST_FLAGS = '$AssistFlags';
  ASSIST_FLAG_ENABLED = 'E';
  ASSIST_FLAG_NEWCOPY = 'N';
  ASSIST_FLAG_HIDDEN =  'H';
  ASSIST_FLAG_PRIVATE = 'P';
  AGENT_MACHINE_NAME = '$MachineName';
{ Please keep these flags in alphabetic order (based on the flag itself) so that
  we can easily tell which flags to use next. Note that some of these flags apply
  to a particular NOTE_CLASS; others apply to all design elements. The comments
  indicate which is which. In theory, flags that apply to two different NOTE_CLASSes
  could overlap, but for now, try to make each flag unique. }

   DESIGN_FLAG_ADD =          'A';  { FORM: Indicates that a subform is in the add subform list }
   DESIGN_FLAG_BACKGROUND_FILTER  = 'B';  { FILTER: Indicates FILTER_TYPE_BACKGROUND is asserted }
   DESIGN_FLAG_NO_COMPOSE  =      'C';  { FORM: Indicates a form that is used only for }
                      {   query by form (not on compose menu). }
   DESIGN_FLAG_CALENDAR_VIEW =    'c';  { VIEW: Indicates a form is a calendar style view. }
   DESIGN_FLAG_NO_QUERY =       'D';  {   FORM: Indicates a form that should not be used in query by form }
   DESIGN_FLAG_DEFAULT_DESIGN =     'd';  {   ALL: Indicates the default design note for it's class (used for VIEW) }
   DESIGN_FLAG_MAIL_FILTER =    'E';  { FILTER: Indicates FILTER_TYPE_MAIL is asserted }
   DESIGN_FLAG_FOLDER_VIEW =      'F';  { VIEW: This is a V4 folder view. }
   DESIGN_FLAG_V4AGENT =      'f';  { FILTER: This is a V4 agent }
   DESIGN_FLAG_VIEWMAP =      'G';  { VIEW: This is ViewMap/GraphicView/Navigator }
   DESIGN_FLAG_OTHER_DLG =      'H';  { ALL: Indicates a form that is placed in Other... dialog }
   DESIGN_FLAG_V4PASTE_AGENT =    'I';  { FILTER: This is a V4 paste agent }
   DESIGN_FLAG_IMAGE_RESOURCE =		'i';  {	FORM: Note is a shared image resource } // New 2003-09-28/Daniel
   DESIGN_FLAG_JAVA_AGENT =       'J'; {  FILTER: If its Java }
   DESIGN_FLAG_JAVA_AGENT_WITH_SOURCE =       'j'; {  FILTER: If it is a java agent with java source code.  }  // Added by Daniel
   DESIGN_FLAG_LOTUSSCRIPT_AGENT =   'L'; {  FILTER: If its LOTUSSCRIPT }
   DESIGN_FLAG_QUERY_MACRO_FILTER =  'M';  { FILTER: Stored FT query AND macro }
   DESIGN_FLAG_SITEMAP =        	'm';  { FILTER: This is a site(m)ap. } // New 2003-09-28/Daniel
   DESIGN_FLAG_NEW =          'N';  {  FORM: Indicates that a subform is listed when making a new form.}
   DESIGN_FLAG_HIDE_FROM_NOTES =    'n'; {  ALL: notes stamped with this flag
                          will be hidden from Notes clients
                          We need a separate value here
                          because it is possible to be
                          hidden from V4 AND to be hidden
                          from Notes, and clearing one
                          should not clear the other }
   DESIGN_FLAG_QUERY_V4_OBJECT =   'O';  { FILTER: Indicates V4 search bar query object - used in addition to 'Q' }
   DESIGN_FLAG_PRIVATE_STOREDESK =   'o'; {  VIEW: If Private_1stUse, store the private view in desktop }
   DESIGN_FLAG_PRESERVE =      'P';  { ALL: related to data dictionary }
   DESIGN_FLAG_PRIVATE_1STUSE =     'p';  {   VIEW: This is a private copy of a private on first use view. }
   DESIGN_FLAG_QUERY_FILTER =    'Q';  { FILTER: Indicates full text query ONLY, no filter macro }
   DESIGN_FLAG_AGENT_SHOWINSEARCH = 'q'; {  FILTER: Search part of this agent should be shown in search bar }
   DESIGN_FLAG_REPLACE_SPECIAL =    'R';  { SPECIAL: this flag is the opposite of DESIGN_FLAG_PRESERVE, used
                        only for the 'About' and 'Using' notes + the icon bitmap in the icon note }
   DESIGN_FLAG_V4BACKGROUND_MACRO =   'S';  { FILTER: This is a V4 background agent }
   DESIGN_FLAG_SCRIPTLIB =      's';  { FILTER: A database global script library note }
   DESIGN_FLAG_VIEW_CATEGORIZED =   'T';  {   VIEW: Indicates a view that is categorized on the categories field }
   DESIGN_FLAG_DATABASESCRIPT =   't';  { FILTER: A database script note }
   DESIGN_FLAG_SUBFORM =      'U';  { FORM: Indicates that a form is a subform.}
   DESIGN_FLAG_AGENT_RUNASWEBUSER =  'u';  { FILTER: Indicates agent should run as effective user on web }
   DESIGN_FLAG_PRIVATE_IN_DB =    'V';  {   ALL: This is a private element stored in the database }
   DESIGN_FLAG_WEBPAGE = 	        'W';	{	FORM: Note is a WEBPAGE	}
   DESIGN_FLAG_HIDE_FROM_WEB =    'w'; {  ALL: notes stamped with this flag
                          will be hidden from WEB clients }
{ WARNING: A formula that build Design Collecion relies on the fact that Agent Data's
      $Flags is the only Desing Collection element whose $Flags='X' }
   DESIGN_FLAG_V4AGENT_DATA =   'X'; {  FILTER: This is a V4 agent data note }
   DESIGN_FLAG_SUBFORM_NORENDER = 'x';  { SUBFORM: indicates whether
                        we should render a subform in
                        the parent form         }
   DESIGN_FLAG_NO_MENU =      'Y';  { ALL: Indicates that folder/view/etc. should be hidden from menu. }
   DESIGN_FLAG_SACTIONS	=		'y';  {	Shared actions note  } // New 2003-09-28/Daniel
   DESIGN_FLAG_MULTILINGUAL_PRESERVE_HIDDEN = 'Z'; { ALL: Used to indicate design element was hidden }
                      { before the 'Notes Global Designer' modified it. }
                      { (used with the '!' flag) }
   DESIGN_FLAG_FRAMESET	=               '#';  {	FORM: Indicates that this is a frameset note } // New 2003-09-28/Daniel
   DESIGN_FLAG_MULTILINGUAL_ELEMENT = '!'; {  ALL: Indicates this design element supports the }
                      { 'Notes Global Designer' multilingual addin }
   DESIGN_FLAG_HIDE_FROM_V3 =   '3';  { ALL: notes stamped with this flag
                          will be hidden from V3 client }
   DESIGN_FLAG_HIDE_FROM_V4 =   '4';  { ALL: notes stamped with this flag
                          will be hidden from V4 client }
   DESIGN_FLAG_HIDE_FROM_V5 =   '5';  { ALL: notes stamped with this flag
                          will be hidden from V5 client }
   DESIGN_FLAG_HIDE_FROM_V6 =   '6';  { ALL: notes stamped with this flag
                          will be hidden from V6 client }
   DESIGN_FLAG_HIDE_FROM_V7 =   '7';  { ALL: notes stamped with this flag
                          will be hidden from V7 client }
   DESIGN_FLAG_HIDE_FROM_V8 =   '8';  { ALL: notes stamped with this flag
                          will be hidden from V8 client }
   DESIGN_FLAG_HIDE_FROM_V9 =   '9';  { ALL: notes stamped with this flag
                          will be hidden from V9 client }
   DESIGN_FLAG_MUTILINGUAL_HIDE = '0';  { ALL: notes stamped with this flag
                          will be hidden from the client
                          usage is for different language
                          versions of the design list to be
                          hidden completely       }


  { Special form flags }

  ITEM_NAME_KEEP_PRIVATE = '$KeepPrivate';
  PRIVATE_FLAG_YES = '1';       { $KeepPrivate = TRUE  force disabling of printing, mail forwarding and edit copy }
  PRIVATE_FLAG_YES_RESEND = '2';      { $KeepPrivate = TRUE  same as PRIVATE_FLAG_YES except allow resend }

  ITEM_NAME_BACKGROUNDGRAPHIC = '$Background';
  ITEM_NAME_PAPERCOLOR = '$PaperColor';
  ITEM_NAME_RESTRICTBKOVERRIDE = '$NoBackgroundOverride';
  RESTRICTBK_FLAG_NOOVERRIDE = '1';   { $NoBackgroundOverride = TRUE Don't allow user to override document background }

  ITEM_NAME_AUTO_EDIT_NOTE = '$AutoEditMode';
  AUTO_EDIT_FLAG_YES  = '1';        { $AutoEditMode = TRUE  force edit mode on open regardless of Form flag }

  ITEM_NAME_SHOW_NAVIGATIONBAR = '$ShowNavigationBar';  { Display the URL navigation Bar }
  ITEM_NAME_HIDE_SCROLL_BARS  = '$HideScrollBars';
  WINDOW_SCROLL_BARS_NONE   = '1';
  WINDOW_SCROLL_BARS_HORZ   = '2';
  WINDOW_SCROLL_BARS_VERT   = '3';



  ITEM_NAME_VERSION_OPT = '$VersionOpt';  { Over-ride the Form flags for versioning. }
  VERSION_FLAG_NONE = '0';        { $Version = 0, None }
  VERSION_FLAG_MURESP = '1';        { $Version = 1, Manual - Update becomes response }
  VERSION_FLAG_AURESP = '2';        { $Version = 2, Auto   - Update becomes response }
  VERSION_FLAG_MUPAR  = '3';        { $Version = 3, Manual - Update becomes parent }
  VERSION_FLAG_AUPAR  = '4';        { $Version = 4, Auto   - Update becomes parent }
  VERSION_FLAG_MUSIB  = '5';        { $Version = 5, Manual - Update becomes sibling }
  VERSION_FLAG_AUSIB  = '6';        { $Version = 6, Auto   - Update becomes sibling }


  { Document note item names }

  ITEM_NAME_TEMPLATE_USED = FIELD_FORM; { form name used to create note, user-visible }
  ITEM_NAME_NOTEREF = FIELD_LINK;   { optional reference to another note }
  ITEM_NAME_VERREF = '$VERREF';     { optional reference to master version note }
  ITEM_NAME_LINK = '$Links';        { note link table }
  ITEM_NAME_REVISIONS = '$Revisions'; { Revision history }
  ITEM_NAME_AUTHORS = '$Authors';   { text list of users allowed to modify document }

  { Document and form note item names, all items are optional }

  ITEM_NAME_FONTS = '$Fonts';     { font table }
  ITEM_NAME_HEADER = '$Header';     { print page header }
  ITEM_NAME_FOOTER = '$Footer';     { print page footer }
  ITEM_NAME_HFFLAGS = '$HFFlags';   { header/footer flags }
  HFFLAGS_NOPRINTONFIRSTPAGE  = '1';    { suppress printing header/footer on first page }
  ITEM_NAME_WINDOWTITLE = '$WindowTitle'; { window title }
  ITEM_NAME_ATTACHMENT = '$FILE';     { file attachment, MUST STAY UPPER-CASE BECAUSE IT'S SIGNED! }
  ITEM_NAME_HTMLBODYTAG = '$HTMLBodyTag'; { Override for HTML body tag }
  ITEM_NAME_WEBQUERYSAVE = '$WEBQuerySave'; {WebQuerySave formula }
  ITEM_NAME_WEBQUERYOPEN = '$WEBQueryOpen'; {WebQueryOpen formula }

  ITEM_NAME_WEBFLAGS  = '$WebFlags';    { Web related flags for form or document }
  WEBFLAG_NOTE_IS_HTML    = 'H';    { treat this document or form as plain HTML, do not convert styled text to HTML }
  WEBFLAG_NOTE_CONTAINS_VIEW  = 'V';    { optimization for web server: this note contains an embedded view }

  { Document note Sign/Seal item names }

  ITEM_NAME_NOTE_SIGNATURE = '$Signature';
  ITEM_NAME_NOTE_SIG_PREFIX = '$Sig_';  { Prefix for multiple signatures. }
  ITEM_NAME_NOTE_SEAL = '$Seal';
  ITEM_NAME_NOTE_SEALDATA = '$SealData';
  ITEM_NAME_NOTE_SEALNAMES = 'SecretEncryptionKeys';
  ITEM_NAME_NOTE_SEALUSERS = 'PublicEncryptionKeys';
  ITEM_NAME_NOTE_FORCESIGN = 'Sign';
  ITEM_NAME_NOTE_FORCESEAL = 'Encrypt';
  ITEM_NAME_NOTE_FORCEMAIL = 'MailOptions';
  ITEM_NAME_NOTE_FORCESAVE = 'SaveOptions';
  ITEM_NAME_NOTE_FORCESEALSAVED = 'EncryptSaved';
  ITEM_NAME_NOTE_MAILSAVE = 'MailSaveOptions';
  ITEM_NAME_NOTE_FOLDERADD = 'FolderOptions';

  { Group expansion item and legal values }

  ITEM_NAME_NOTE_GROUPEXP  = 'ExpandPersonalGroups';  { For backward compatibility }
  ITEM_NAME_NOTE_EXPANDGROUPS  = '$ExpandGroups';
  MAIL_DONT_EXPAND_GROUPS     = '0';
  MAIL_EXPAND_LOCAL_GROUPS    = '1';
  MAIL_EXPAND_PUBLIC_GROUPS   = '2';
  MAIL_EXPAND_LOCAL_AND_PUBLIC_GROUPS = '3';

  { Search term highlights item name prefix.  An item name is
  concatenated to this; e.g. $Highlights_Body.  }

  ITEM_NAME_HIGHLIGHTS  = '$Highlights_';

  { Import/Export document item names }

  IMPORT_BODY_ITEM_NAME = 'Body';
  IMPORT_FORM_ITEM_NAME = FIELD_FORM;
  NEW_FORM_ITEM_NAME = FIELD_FORM;


(******************************************************************************)
{from ODS.H}
(******************************************************************************)
type
  ActionRoutinePtr = function (RecordPtr: pchar; RecordType: word; RecordLength: dword; vContext: pointer): STATUS; stdcall;

function EnumCompositeBuffer (ItemValue: BLOCKID; ItemValueLength: DWORD; ActionRoutine: ActionRoutinePtr;
  vContext: pointer): STATUS; stdcall; far;

procedure ODSReadMemory(ppSrc: pointer; mtype: word; pDest: pointer; iterations: word); stdcall; far;
procedure ODSWriteMemory(ppDest: pointer; mtype: word; pSrc: pointer; iterations: word); stdcall; far;
function ODSLength(mtype: word): word; stdcall; far;

const
  LONGRECORDLENGTH = 0;
  WORDRECORDLENGTH = $ff00;
  BYTERECORDLENGTH = 0;   // High byte contains record length

// Base ODS types
const
  _SHORT          = 0;
  _USHORT         = _SHORT;
  _WORD         = _SHORT;
  _BOOL         = _SHORT;
  _STATUS         = _SHORT;
  _UNICODE        = _SHORT;
  _LONG         = 1;
  _FLOAT          = 2;
  _DWORD          = _LONG;
  _ULONG          = _LONG;

// ODS types that are the size of a base type

  _NUMBER         = _FLOAT;
  _NOTEID         = _LONG;

// ODS types as results of odsmacro
  _TIMEDATE = 10;
  _TIMEDATE_PAIR = 11;
  _NUMBER_PAIR = 12;
  _LIST = 13;
  _RANGE = 14;
  _DBID = 15;
  _ITEM = 17;
  _ITEM_TABLE = 18;
  _SEARCH_MATCH = 24;
  _ORIGINATORID = 26;
  _OID = _ORIGINATORID;
  _OBJECT_DESCRIPTOR = 27;
  _UNIVERSALNOTEID = 28;
  _UNID = _UNIVERSALNOTEID;
  _VIEW_TABLE_FORMAT = 29;
  _VIEW_COLUMN_FORMAT = 30;
  _NOTELINK = 33;
  _LICENSEID = 34;
  _VIEW_FORMAT_HEADER = 42;
  _VIEW_TABLE_FORMAT2 = 43;
  _DBREPLICAINFO = 56;
  _FILEOBJECT = 58;
  _COLLATION = 59;
  _COLLATE_DESCRIPTOR = 60;
  _CDKEYWORD = 68;
  _CDLINK2 = 72;
  _CDLINKEXPORT2 = 97;
  _CDPARAGRAPH = 109;
  _CDPABDEFINITION = 110;
  _CDPABREFERENCE = 111;
  _CDFIELD_PRE_36 = 112;
  _CDTEXT = 113;
  _CDDOCUMENT = 114;
  _CDMETAFILE = 115;
  _CDBITMAP = 116;
  _CDHEADER = 117;
  _CDFIELD = 118;
  _CDFONTTABLE = 119;
  _CDFACE = 120;
  _CDCGM = 156;
  _CDTIFF = 159;
  _CDBITMAPHEADER = 162;
  _CDBITMAPSEGMENT = 163;
  _CDCOLORTABLE = 164;
  _CDPATTERNTABLE = 165;
  _CDGRAPHIC = 166;
  _CDPMMETAHEADER = 167;
  _CDWINMETAHEADER = 168;
  _CDMACMETAHEADER = 169;
  _CDCGMMETA = 170;
  _CDPMMETASEG = 171;
  _CDWINMETASEG = 172;
  _CDMACMETASEG = 173;
  _CDDDEBEGIN = 174;
  _CDDDEEND = 175;
  _CDTABLEBEGIN = 176;
  _CDTABLECELL = 177;
  _CDTABLEEND = 178;
  _CDSTYLENAME = 188;
  _FILEOBJECT_MACEXT = 192;
  _FILEOBJECT_HPFSEXT = 193;
  _CDOLEBEGIN = 218;
  _CDOLEEND = 219;
  _CDHOTSPOTBEGIN = 230;
  _CDHOTSPOTEND = 231;
  _CDBUTTON = 237;
  _CDBAR = 308;
  _CDQUERYHEADER = 314;
  _CDQUERYTEXTTERM = 315;
  _CDACTIONHEADER = 316;
  _CDACTIONMODIFYFIELD = 317;
  _ODS_ASSISTSTRUCT = 318;
  _VIEWMAP_HEADER_RECORD = 319;
  _VIEWMAP_RECT_RECORD = 320;
  _VIEWMAP_BITMAP_RECORD = 321;
  _VIEWMAP_REGION_RECORD = 322;
  _VIEWMAP_POLYGON_RECORD_BYTE = 323;
  _VIEWMAP_POLYLINE_RECORD_BYTE = 324;
  _VIEWMAP_ACTION_RECORD = 325;
  _ODS_ASSISTRUNINFO = 326;
  _CDACTIONREPLY = 327;
  _CDACTIONFORMULA = 332;
  _CDACTIONLOTUSSCRIPT = 333;
  _CDQUERYBYFIELD = 334;
  _CDACTIONSENDMAIL = 335;
  _CDACTIONDBCOPY = 336;
  _CDACTIONDELETE = 337;
  _CDACTIONBYFORM = 338;
  _ODS_ASSISTFIELDSTRUCT = 339;
  _CDACTION = 340;
  _CDACTIONREADMARKS = 341;
  _CDEXTFIELD = 342;
  _CDLAYOUT = 343;
  _CDLAYOUTTEXT = 344;
  _CDLAYOUTEND = 345;
  _CDLAYOUTFIELD = 346;
  _VIEWMAP_DATASET_RECORD = 347;
  _CDDOCAUTOLAUNCH = 350;
  _CDPABHIDE = 358;
  _CDPABFORMULAREF = 359;
  _CDACTIONBAR = 360;
  _CDACTIONFOLDER = 361;
  _CDACTIONNEWSLETTER = 362;
  _CDACTIONRUNAGENT = 363;
  _CDACTIONSENDDOCUMENT = 364;
  _CDQUERYFORMULA = 365;
  _CDQUERYBYFORM = 373;
  _ODS_ASSISTRUNOBJECTHEADER = 374;
  _ODS_ASSISTRUNOBJECTENTRY = 375;
  _CDOLEOBJ_INFO = 379;
  _CDLAYOUTGRAPHIC = 407;
  _CDQUERYBYFOLDER = 413;
  _CDQUERYUSESFORM = 423;
  _VIEW_COLUMN_FORMAT2 = 428;
  _VIEWMAP_TEXT_RECORD = 464;
  _CDLAYOUTBUTTON = 466;
  _CDQUERYTOPIC = 471;
  _CDLSOBJECT = 482;
  _CDHTMLHEADER = 492;
  _CDHTMLSEGMENT = 493;
  _SCHED_LIST = 502;
  _SCHED_LIST_OBJ = _SCHED_LIST;
  _SCHED_ENTRY = 503;
  _SCHEDULE = 504;
  _CDTEXTEFFECT = 508;
  _CDSTORAGELINK = 515;
  _ACTIVEOBJECT = 516;
  _ACTIVEOBJECTPARAM = 517;
  _ACTIVEOBJECTSTORAGELINK = 518;
  _CDTRANSPARENTTABLE = 541;
  _VIEWMAP_POLYGON_RECORD = 551;
  _VIEWMAP_POLYLINE_RECORD = 552;
  _SCHED_ENTRY_DETAIL = 553;
  _CDALTERNATEBEGIN = 554;
  _CDALTERNATEEND = 555;
  _CDOLERTMARKER = 556;
  _HSOLERICHTEXT = 557;
  _CDANCHOR = 559;
  _CDHRULE = 560;
  _CDALTTEXT = 561;
  _CDACTIONJAVAAGENT = 562;
  _CDHTMLBEGIN = 564;
  _CDHTMLEND = 565;
  _CDHTMLFORMULA = 566;

{ Signatures for Composite Records in items of data type COMPOSITE }
const
  SIG_CD_PARAGRAPH  = (129 or BYTERECORDLENGTH);
  SIG_CD_PABDEFINITION= (130 or WORDRECORDLENGTH);
  SIG_CD_PABREFERENCE = (131 or BYTERECORDLENGTH);
  SIG_CD_TEXT     = (133 or WORDRECORDLENGTH);
  SIG_CD_HEADER   = (142 or WORDRECORDLENGTH);
  SIG_CD_LINKEXPORT2  = (146 or WORDRECORDLENGTH);
  SIG_CD_BITMAPHEADER = (149 or LONGRECORDLENGTH);
  SIG_CD_BITMAPSEGMENT   = (150 or LONGRECORDLENGTH);
  SIG_CD_COLORTABLE    = (151 or LONGRECORDLENGTH);
  SIG_CD_GRAPHIC       = (153 or LONGRECORDLENGTH);
  SIG_CD_PMMETASEG     = (154 or LONGRECORDLENGTH);
  SIG_CD_WINMETASEG    = (155 or LONGRECORDLENGTH);
  SIG_CD_MACMETASEG    = (156 or LONGRECORDLENGTH);
  SIG_CD_CGMMETA       = (157 or LONGRECORDLENGTH);
  SIG_CD_PMMETAHEADER    = (158 or LONGRECORDLENGTH);
  SIG_CD_WINMETAHEADER   = (159 or LONGRECORDLENGTH);
  SIG_CD_MACMETAHEADER   = (160 or LONGRECORDLENGTH);
  SIG_CD_TABLEBEGIN = (163 or BYTERECORDLENGTH);
  SIG_CD_TABLECELL  = (164 or BYTERECORDLENGTH);
  SIG_CD_TABLEEND   = (165 or BYTERECORDLENGTH);
  SIG_CD_STYLENAME  = (166 or BYTERECORDLENGTH);
  SIG_CD_STORAGELINK  = (196 or WORDRECORDLENGTH);
  SIG_CD_TRANSPARENTTABLE= (197 or LONGRECORDLENGTH);
  SIG_CD_HORIZONTALRULE=  (201 or WORDRECORDLENGTH);
  SIG_CD_ALTTEXT    = (202 or WORDRECORDLENGTH);
  SIG_CD_ANCHOR   = (203 or WORDRECORDLENGTH);
  SIG_CD_HTMLBEGIN  = (204 or WORDRECORDLENGTH);
  SIG_CD_HTMLEND    = (205 or WORDRECORDLENGTH);
  SIG_CD_HTMLFORMULA  = (206 or WORDRECORDLENGTH);

  { Signatures for Composite Records that are reserved internal records, }
  { whose format may change between releases. }

  SIG_CD_DOCUMENT_PRE_26= (128 or BYTERECORDLENGTH);
  SIG_CD_FIELD_PRE_36 = (132 or WORDRECORDLENGTH);
  SIG_CD_FIELD    = (138 or WORDRECORDLENGTH);
  SIG_CD_DOCUMENT   = (134 or BYTERECORDLENGTH);
  SIG_CD_METAFILE   = (135 or WORDRECORDLENGTH);
  SIG_CD_BITMAP   = (136 or WORDRECORDLENGTH);
  SIG_CD_FONTTABLE  = (139 or WORDRECORDLENGTH);
  SIG_CD_LINK     = (140 or BYTERECORDLENGTH);
  SIG_CD_LINKEXPORT = (141 or BYTERECORDLENGTH);
  SIG_CD_KEYWORD    = (143 or WORDRECORDLENGTH);
  SIG_CD_LINK2    = (145 or WORDRECORDLENGTH);
  SIG_CD_CGM      = (147 or WORDRECORDLENGTH);
  SIG_CD_TIFF     = (148 or LONGRECORDLENGTH);
  SIG_CD_PATTERNTABLE    = (152 or LONGRECORDLENGTH);
  SIG_CD_DDEBEGIN   = (161 or WORDRECORDLENGTH);
  SIG_CD_DDEEND   = (162 or WORDRECORDLENGTH);
  SIG_CD_OLEBEGIN   = (167 or WORDRECORDLENGTH);
  SIG_CD_OLEEND   = (168 or WORDRECORDLENGTH);
  SIG_CD_HOTSPOTBEGIN = (169 or WORDRECORDLENGTH);
  SIG_CD_HOTSPOTEND = (170 or BYTERECORDLENGTH);
  SIG_CD_BUTTON   = (171 or WORDRECORDLENGTH);
  SIG_CD_BAR      = (172 or WORDRECORDLENGTH);
  SIG_CD_V4HOTSPOTBEGIN=  (173 or WORDRECORDLENGTH);
  SIG_CD_V4HOTSPOTEND = (174 or BYTERECORDLENGTH);
  SIG_CD_EXT_FIELD  = (176 or WORDRECORDLENGTH);
  SIG_CD_LSOBJECT   = (177 or WORDRECORDLENGTH){ Compiled LS code};
  SIG_CD_HTMLHEADER = (178 or WORDRECORDLENGTH) { Raw HTML };
  SIG_CD_HTMLSEGMENT  = (179 or WORDRECORDLENGTH);
  SIG_CD_LAYOUT   = (183 or BYTERECORDLENGTH);
  SIG_CD_LAYOUTTEXT = (184 or BYTERECORDLENGTH);
  SIG_CD_LAYOUTEND  = (185 or BYTERECORDLENGTH);
  SIG_CD_LAYOUTFIELD  = (186 or BYTERECORDLENGTH);
  SIG_CD_PABHIDE    = (187 or WORDRECORDLENGTH);
  SIG_CD_PABFORMREF = (188 or BYTERECORDLENGTH);
  SIG_CD_ACTIONBAR  = (189 or BYTERECORDLENGTH);
  SIG_CD_ACTION   = (190 or WORDRECORDLENGTH);

  SIG_CD_DOCAUTOLAUNCH= (191 or WORDRECORDLENGTH);
  SIG_CD_LAYOUTGRAPHIC= (192 or BYTERECORDLENGTH);
  SIG_CD_OLEOBJINFO = (193 or WORDRECORDLENGTH);
  SIG_CD_LAYOUTBUTTON = (194 or BYTERECORDLENGTH);
  SIG_CD_TEXTEFFECT = (195 or WORDRECORDLENGTH);

  { Saved Query records for items of type TYPE_QUERY }

  SIG_QUERY_HEADER  = (129 or BYTERECORDLENGTH);
  SIG_QUERY_TEXTTERM  = (130 or WORDRECORDLENGTH);
  SIG_QUERY_BYFIELD = (131 or WORDRECORDLENGTH);
  SIG_QUERY_BYDATE  = (132 or WORDRECORDLENGTH);
  SIG_QUERY_BYAUTHOR  = (133 or WORDRECORDLENGTH);
  SIG_QUERY_FORMULA = (134 or WORDRECORDLENGTH);
  SIG_QUERY_BYFORM  = (135 or WORDRECORDLENGTH);
  SIG_QUERY_BYFOLDER  = (136 or WORDRECORDLENGTH);
  SIG_QUERY_USESFORM  = (137 or WORDRECORDLENGTH);
  SIG_QUERY_TOPIC   = (138 or WORDRECORDLENGTH);

  { Save Action records for items of type TYPE_ACTION }
  ASSIST_SIG_ACTION_NONE  = -1; {* No action defined *} // Added by Daniel

  SIG_ACTION_HEADER = (129 or BYTERECORDLENGTH);
  SIG_ACTION_MODIFYFIELD= (130 or WORDRECORDLENGTH);
  SIG_ACTION_REPLY  = (131 or WORDRECORDLENGTH);
  SIG_ACTION_FORMULA  = (132 or WORDRECORDLENGTH);
  SIG_ACTION_LOTUSSCRIPT= (133 or WORDRECORDLENGTH);
  SIG_ACTION_SENDMAIL = (134 or WORDRECORDLENGTH);
  SIG_ACTION_DBCOPY = (135 or WORDRECORDLENGTH);
  SIG_ACTION_DELETE = (136 or BYTERECORDLENGTH);
  SIG_ACTION_BYFORM = (137 or WORDRECORDLENGTH);
  SIG_ACTION_MARKREAD = (138 or BYTERECORDLENGTH);
  SIG_ACTION_MARKUNREAD=  (139 or BYTERECORDLENGTH);
  SIG_ACTION_MOVETOFOLDER=  (140 or WORDRECORDLENGTH);
  SIG_ACTION_COPYTOFOLDER=  (141 or WORDRECORDLENGTH);
  SIG_ACTION_REMOVEFROMFOLDER=  (142 or WORDRECORDLENGTH);
  SIG_ACTION_NEWSLETTER=  (143 or WORDRECORDLENGTH);
  SIG_ACTION_RUNAGENT = (144 or WORDRECORDLENGTH);
  SIG_ACTION_SENDDOCUMENT=  (145 or BYTERECORDLENGTH);
  SIG_ACTION_FORMULAONLY= (146 or WORDRECORDLENGTH);
  SIG_ACTION_JAVAAGENT= (147 or WORDRECORDLENGTH);


  { Signatures for items of type TYPE_VIEWMAP_DATASET }

  SIG_VIEWMAP_DATASET=  (87 or WORDRECORDLENGTH);

  { Signatures for items of type TYPE_VIEWMAP }

  SIG_CD_VMHEADER   = (175 or BYTERECORDLENGTH);
  SIG_CD_VMBITMAP   = (176 or BYTERECORDLENGTH);
  SIG_CD_VMRECT   = (177 or BYTERECORDLENGTH);
  SIG_CD_VMPOLYGON_BYTE=  (178 or BYTERECORDLENGTH);
  SIG_CD_VMPOLYLINE_BYTE= (179 or BYTERECORDLENGTH);
  SIG_CD_VMREGION   = (180 or BYTERECORDLENGTH);
  SIG_CD_VMACTION   = (181 or BYTERECORDLENGTH);
  SIG_CD_VMELLIPSE  = (182 or BYTERECORDLENGTH);
  SIG_CD_VMRNDRECT  = (184 or BYTERECORDLENGTH);
  SIG_CD_VMBUTTON   = (185 or BYTERECORDLENGTH);
  SIG_CD_VMACTION_2 = (186 or WORDRECORDLENGTH);
  SIG_CD_VMTEXTBOX  = (187 or WORDRECORDLENGTH);
  SIG_CD_VMPOLYGON  = (188 or WORDRECORDLENGTH);
  SIG_CD_VMPOLYLINE = (189 or WORDRECORDLENGTH);
  SIG_CD_VMPOLYRGN  = (190 or WORDRECORDLENGTH);
  SIG_CD_VMCIRCLE   = (191 or BYTERECORDLENGTH);
  SIG_CD_VMPOLYRGN_BYTE=  (192 or BYTERECORDLENGTH);

  { Signatures for alternate CD sequences}
  SIG_CD_ALTERNATEBEGIN=  (198 or WORDRECORDLENGTH);
  SIG_CD_ALTERNATEEND = (199 or BYTERECORDLENGTH);

  CD_BUFFER_LENGTH                      = 64000;  { max segment size }
  CD_HIGH_WATER_MARK                    = 40000;  { max item size }

(******************************************************************************)
{from MISC.H}
(******************************************************************************)

const
  TIMEDATE_MINIMUM  = 0;
  TIMEDATE_MAXIMUM  = 1;
  TIMEDATE_WILDCARD = 2;

procedure TimeConstant (TimeConstantType: WORD; var Value: TIMEDATE); far; stdcall; far;

// Added by Daniel 2002-10-28. Support constants for agents
const
        ALLDAY = $ffffffff; { put this in the TIME field }
        ANYDAY = $ffffffff; { put this in the DATE field }
        TICKS_IN_DAY = 8640000; { 10msec ticks in a day }
        TICKS_IN_HOUR = 360000; { 10msec ticks in an hour }
        TICKS_IN_MINUTE = 6000; { 10msec ticks in a minute }
        TICKS_IN_SECOND = 100;  { 10msec ticks in a second }
        SECS_IN_DAY = 86400;  { seconds in a day }
// End addition

(******************************************************************************)
{ from ossignal.h }
(******************************************************************************)
const
  OS_SIGNAL_MESSAGE = 3;  //Indirect way to call NEMMessageBox */
                          //STATUS = Proc(Message, OSMESSAGETYPE_xxx) */
type
  OSSIGMSGPROC = function (Message: pchar; wType: WORD): STATUS; stdcall;

const
  OS_SIGNAL_BUSY = 4;     //Paint busy indicator on screen
                          //STATUS = Proc(BUSY_xxx)
type
  OSSIGBUSYPROC = function (BusyType: word): STATUS; stdcall;

const
  OS_SIGNAL_CHECK_BREAK = 5;  //Called from NET to see if user cancelled I/O */
                              //STATUS = Proc(void) */
  OS_SIGNAL_BREAK = $0100 + 157;  //cancel code
type
  OSSIGBREAKPROC = function: STATUS; stdcall;

const
  OS_SIGNAL_DIAL = 10;  // Prompt to dial a remote system */
                        //pServer = Desired server name (or NULL) */
                        //pPort = Desired port name (or NULL) */
                        //pDialParams = Reserved */
                        //pRetServer = Actual server name to be called */
                        //  (or NULL if not desired) */
                        //pRetPort = Actual port name being used */
                        //  (or NULL if not desired) */

type
  OSSIGDIALPROC = function (pServer: pchar;
                         pPort: pchar;
                         pDialParams: pointer;
                         pRetServer: pchar;
                         pRetPort: pchar): STATUS; stdcall;


  OSSIGPROC = pointer;

function OSSetSignalHandler (wType: WORD; Proc: OSSIGPROC): OSSIGPROC; stdcall; far;
function OSGetSignalHandler (wType: WORD): OSSIGPROC;  stdcall; far;

//  Definitions specific to message signal handler */

const
  OSMESSAGETYPE_OK      = 0;
  OSMESSAGETYPE_OKCANCEL    = 1;
  OSMESSAGETYPE_YESNO     = 2;
  OSMESSAGETYPE_YESNOCANCEL = 3;
  OSMESSAGETYPE_RETRYCANCEL = 4;
  OSMESSAGETYPE_POST      = 5;
  OSMESSAGETYPE_POST_NOSERVER = 6;


// Definitions specific to busy signal handler */

const
  BUSY_SIGNAL_FILE_INACTIVE = 0;
  BUSY_SIGNAL_FILE_ACTIVE   = 1;
  BUSY_SIGNAL_NET_INACTIVE  = 2;
  BUSY_SIGNAL_NET_ACTIVE    = 3;
  BUSY_SIGNAL_POLL      = 4;
  BUSY_SIGNAL_WAN_SENDING   = 5;
  BUSY_SIGNAL_WAN_RECEIVING = 6;

(******************************************************************************)
{ from global.h }
(******************************************************************************)
type
  VARARG_PTR = pointer;
  function VARARG_GET (var AP: VARARG_PTR; TypeSz: word): pointer;

(******************************************************************************)
{ from extmgr.h }
(******************************************************************************)
type
  EID = WORD;
  HEMREGISTRATION = DWORD;
  PHEMREGISTRATION = ^HEMREGISTRATION;

  EMRECORD = packed record
    EId: EID;                      //* identifier */
    NotificationType: WORD;        //* EM_BEFORE or EM_AFTER */
    Status: STATUS;                //* core error code */
    Ap: VARARG_PTR;                //* ptr to args */
  end;
  PEMRECORD = ^EMRECORD;

//* the callback; takes one argument */

  EMHANDLER = function (aRecord: PEMRECORD): STATUS; stdcall;

//* prototypes */

  function EMRegister(EmID: EID; Flags: DWORD; Proc: EMHANDLER; RecursionID: WORD; rethRegistration: PHEMREGISTRATION): STATUS; stdcall; far;
  function EMDeregister(hRegistration: HEMREGISTRATION): STATUS; stdcall; far;
  function EMCreateRecursionID(retRecursionID: PWORD): STATUS; stdcall; far;

//* Constants used in NotificationType */

const
  EM_BEFORE = 0;
  EM_AFTER  = 1;

//* Flags which can be passed to EMRegister */

  EM_REG_BEFORE   = $0001;
  EM_REG_AFTER    = $0002;


//* Types of extension callbacks */

  EM_NSFDBCLOSESESSION        = 1;
  EM_NSFDBCLOSE           = 2;
  EM_NSFDBCREATE          =   3;
  EM_NSFDBDELETE          =   4;
  EM_NSFNOTEOPEN          =   5;
  EM_NSFNOTECLOSE         =   6;
  EM_NSFNOTECREATE        =   7;
  EM_NSFNOTEDELETE        =   8;
  EM_NSFNOTEOPENBYUNID    =     10;
  EM_FTGETLASTINDEXTIME   =     11;
  EM_FTINDEX              = 12;
  EM_FTSEARCH             = 13;
  EM_NIFFINDBYKEY         =   14;
  EM_NIFFINDBYNAME        =   15;
  EM_NIFREADENTRIES       =   18;
  EM_NIFUPDATECOLLECTION  =       20;
  EM_NSFDBALLOCOBJECT     =     22;
  EM_NSFDBCOMPACT         =   23;
  EM_NSFDBDELETENOTES     =     24;
  EM_NSFDBFREEOBJECT      =     25;
  EM_NSFDBGETMODIFIEDNOTETABLE =    26;
  EM_NSFDBGETNOTEINFO     =     29;
  EM_NSFDBGETNOTEINFOBYUNID =     30;
  EM_NSFDBGETOBJECTSIZE     =   31;
  EM_NSFDBGETSPECIALNOTEID  =     32;
  EM_NSFDBINFOGET           = 33;
  EM_NSFDBINFOSET           = 34;
  EM_NSFDBLOCATEBYREPLICAID =     35;
  EM_NSFDBMODIFIEDTIME      =   36;
  EM_NSFDBREADOBJECT        =   37;
  EM_NSFDBREALLOCOBJECT     =   39;
  EM_NSFDBREPLICAINFOGET    =     40;
  EM_NSFDBREPLICAINFOSET    =     41;
  EM_NSFDBSPACEUSAGE        =   42;
  EM_NSFDBSTAMPNOTES        =   43;
  EM_NSFDBWRITEOBJECT       =   45;
  EM_NSFNOTEUPDATE          = 47;
  EM_NIFOPENCOLLECTION      =   50;
  EM_NIFCLOSECOLLECTION     =   51;
  EM_NSFDBGETBUILDVERSION   =     52;
  EM_NSFDBITEMDEFTABLE      =   56;
  EM_NSFDBREOPEN            = 59;
  EM_NSFDBOPENEXTENDED      =   63;
  EM_NSFNOTEDECRYPT         = 70;
  EM_GETPASSWORD            = 73;
  EM_SETPASSWORD            = 74;
  EM_NSFCONFLICTHANDLER     =   75;
  EM_CLEARPASSWORD          = 90;
  EM_SCHFREETIMESEARCH      =   105;
  EM_SCHRETRIEVE            = 106;
  EM_SCHSRVRETRIEVE         = 107;

// Error codes (see Util_LnApiErr)

(******************************************************************************)
{ ns.h }
(******************************************************************************)

(* function templates *)
function NSGetServerList (pPortName: pchar; retServerTextList: PHandle): STATUS; stdcall; far;

const CLUSTER_LOOKUP_NOCACHE        = $00000001;   (* don't use cluster name cache *)
const CLUSTER_LOOKUP_CACHEONLY      = $00000002;   (* only use cluster name cache *)

function NSGetServerClusterMates (pServerName: pchar; dwFlags: DWORD; var phList: THandle): STATUS; stdcall; far;
function NSPingServer (pServerName: pchar; pdwIndex: PDWORD; var phList: THandle): STATUS; stdcall; far;


(******************************************************************************)
{ from lookup.h }
(******************************************************************************)
{+// Name & Address Book lookup package definitions*/ }

const
  NAME_GET_AB_TITLES = $0001;
const
  NAME_DEFAULT_TITLES = $0002;
const
  NAME_GET_AB_FIRSTONLY = $0004;


function NAMEGetAddressBooks(pszServer: PChar;
                             wOptions: Word;
                             var pwReturnCount: Word;
                             var pwReturnLength: Word;
                             var phReturn: Handle): STATUS; stdcall; far;

procedure NAMEGetModifiedTime(var retModified: TIMEDATE); stdcall; far;
function NAMELookup(ServerName: PChar;
                    Flags: Word;
                    NumNameSpaces: Word;
                    NameSpaces: PChar;
                    NumNames: Word;
                    Names: PChar;
                    NumItems: Word;
                    Items: PChar;
                    var rethBuffer: Handle): STATUS; stdcall; far;

function NAMELocateNextName(pLookup: Pointer;
                             pName: Pointer;
                             retNumMatches: PWord): Pointer; stdcall; far;

function NAMELocateNextMatch(pLookup: Pointer;
                              pName: Pointer;
                              pMatch: Pointer): Pointer; stdcall; far;

function NAMELocateItem(pMatch: Pointer;
                         Item: Word;
                         var retDataType: Word;
                         retSize: PWord): Pointer; stdcall; far;

function NAMEGetTextItem(pMatch: Pointer;
                         Item: Word;
                         Member: Word;
                         Buffer: PChar;
                         BufLen: Word): STATUS; stdcall; far;

function NAMELocateMatchAndItem(pLookup: Pointer;
                                MatchNum: Word;
                                Item: Word;
                                var retDataType: Word;
                                retpMatch: Pointer;
                                retpItem: Pointer;
                                var retSize: Word): STATUS; stdcall; far;


{+// NAMELookup flags*/ }

const
  NAME_LOOKUP_ALL = $0001; {/* Return all entries in the view*/}
{+// (Note: a Names value of "" must also be specified)*/ }

const
  NAME_LOOKUP_NOSEARCHING = $0002; {/* Only look in first names database containing*/}

{+// desired namespace (view) for specified names*/ }
{+// rather than searching other names databases*/ }
{+// if name was not found. Note that this may not*/ }
{+// necessarily be the first names database in the*/ }
{+// search path - just the first one containing*/ }
{+// the desired view.*/ }
const
  NAME_LOOKUP_EXHAUSTIVE = $0020; {/* Do not stop searching when the first*/}
{+// matching entry is found.*/ }

{+// NAMELookup programming notes: }

{-NAMELookup offers the capability to lookup an arbitrary number of }
{-"items" of information for an arbitrary number of "names" in a }
{-single procedure call. Furthermore, NAMELookup may return }
{-multiple "matches" for each name, each with the selected items of }
{-information. Finally, this lookup is performed in one or more }
{-"name spaces" which really refer to the names of views in the }
{-name & address book(s) on a specified server. }

{-Note the terminology: "names", "matches", and "items". These }
{-concepts relate directly to the API calls. }

{-An example will help to illustrate this. Suppose a piece of mail }
{-is being sent to 3 names: Bill Smith, Ted Jones, and Marketing }
{-(Marketing is the name of a group). }

{-For each user name, the mailer needs the mail domain, the mail server, }
{-and the public key. For the group, the mailer needs the members of }
{-the group. }

{-So the "names" are Bill Smith, Ted Jones, and Marketing. }

{-The "items" are domain, server, public key, and members. }

{-However, if there are two individuals named "Bill Smith" NAMELookup }
{-will return two matches for the name "Bill Smith". The mailer can }
{-then use the returned information (e.g. domain) to display a dialog }
{-box to ask the user which Bill Smith to send to (e.g. Bill Smith @ Iris }
{-or Bill Smith @ Lotus). }

{-In this example, the mailer doesn't know prior to the NAMELookup call }
{-which names refer to individuals and which names refer to groups. }
{-The results from the NAMELookup can be used to determine this. }
{-For example, if for a given name, the member item is not returned, }
{-the name can be assumed to be an individual. If the member item }
{-is returned, then it is a group. The routine NAMELocateItem }
{-returns a null pointer and a zero size if the item is not present }
{-in this match. }

const
  USER_NAMESSPACE = '$Users';

const
  ITEM_DOMAIN = 0;
const
  ITEM_SERVER = 1;
const
  ITEM_PUBLICKEY = 2;
const
  ITEM_MEMBERS = 3;
const
  ITEM_CERTIFICATE = 4;

{+// Structure of the header of the return buffer.*/ }

type
  LOOKUP_HEADER = record
    Length: Word;
  end {LOOKUP_HEADER};


{+// Structure which is returned for each name to be looked up, in the }
{=same order as the names were provided in the request. }

type
  LOOKUP_INFO = record
{ WORD NumMatches; /* # records which match the name*/ }
  end {LOOKUP_INFO};


{+// Structure which is returned for every matching record of a name.*/ }

type
  LOOKUP_MATCH = record
    Length: Word;
  end {LOOKUP_MATCH};


(******************************************************************************)
{ dname.h }
(******************************************************************************)
const
  DN_DELIM = '/'; {/* Component delimiter between RDN's*/}
const
  DN_DELIM_ALT1 = ','; {/* Component delimiter between RDN's*/}
const
  DN_DELIM_ALT2 = ''; {/* Component delimiter between RDN's*/}
const
  DN_DELIM_RDN = '+'; {/* Component delimiter within an RDN*/}
const
  DN_DELIM_RDN_ABBREV = '+'; {/* Display component delimiter within an RDN (when abbreviating)*/}
const
  DN_TYPE_DELIM = '='; {/* type name delimiter*/}
const
  DN_OUNITS = 4; {/* Maximum number org units*/}

const
  DN_NONSTANDARD = $0001; {/* Name includes non-standard components*/}
{+// Ie., contains unrecognized labels*/ }
const
  DN_NONDISTINGUISHED = $0002; {/* Non-distinguished name*/}
{+// Ie., contains no delimiters or labeled attributes*/ }
const
  DN_CN_OU_RDN = $0008; {/* CN plus OU are relative distinguished name*/}
const
  DN_O_C_RDN = $0010; {/* O plus C are relative distinguished name*/}
const
  DN_NONABBREV = $0020; {/* Name includes components that cannot be abbreviated*/}
{+// E.g., G, I, S, Q, P, A*/ }


{+// Distinguished name parsing result data structure*/ }

type
  DN_COMPONENTS = packed record
    Flags: LongInt;
{= Parsing flags }
    CLength: Word;
{= Country name length }
    C: PChar;
{= Country name pointer }
    OLength: Word;
{= Organization name length }
    O: PChar;
{= Organization name pointer }
    OULength: Array[0..DN_OUNITS-1] of Word;
{= Org Unit name lengths }
{+// OULength[0] is rightmost org unit*/ }
    OU: Array[0..DN_OUNITS-1] of PChar;
{= Org unit name pointers }
{+// OU[0] is rightmost org unit*/ }
    CNLength: Word;
{= Common name length }
    CN: PChar;
{= Common name pointer }
    DomainLength: Word;
{= Domain name length }
    Domain: PChar;
{= Domain name pointer }

{+// Original V3 structure ended here. The following fields were added in V4*/ }

    PRMDLength: Word;
{= Private management domain name length }
    PRMD: PChar;
{= Private management domain name pointer }
    ADMDLength: Word;
{= Administration management domain name length }
    ADMD: PChar;
{= Administration management domain name pointer }
    GLength: Word;
{= Given name length }
    G: PChar;
{= Given name name pointer }
    SLength: Word;
{= Surname length }
    S: PChar;
{= Surname pointer }
    ILength: Word;
{= Initials length }
    I: PChar;
{= Initials pointer }
    QLength: Word;
{= Generational qualifier (e.g., Jr) length }
    Q: PChar;
{= Generational qualifier (e.g., Jr) pointer }
  end {DN_COMPONENTS};

const
  DN_ABBREV_INCLUDEALL = $00000001; {/* Include all component types, even when same as template*/}

function DNAbbreviate(Flags: LongInt;
                      TemplateName: PChar;
                      InName: PChar;
                      OutName: PChar;
                      OutSize: Word;
                      var OutLength: Word): STATUS; stdcall; far;

function DNCanonicalize(Flags: LongInt;
                        TemplateName: PChar;
                        InName: PChar;
                        OutName: PChar;
                        OutSize: Word;
                        var OutLength: Word): STATUS; stdcall; far;


function DNParse(Flags: LongInt;
                 TemplateName: PChar;
                 InName: PChar;
                 var Comp: DN_COMPONENTS;
                 CompSize: Word): STATUS; stdcall; far;



(******************************************************************************)
{ foldman.h }
(******************************************************************************)
type
  DESIGN_TYPE = DWORD;

const
  DESIGN_TYPE_SHARED        = 0;      //* Note is shared (always located in the database) */
  DESIGN_TYPE_PRIVATE_DATABASE  = 1;  //* Note is private and is located in the database */

function FolderCreate
      (
      hDataDB: DBHANDLE;
      hFolderDB: DBHANDLE;
      FormatNoteID: NOTEID;
      hFormatDB: DBHANDLE;
      pszName: pchar;
      wNameLen: WORD;
      FolderType: DESIGN_TYPE;
      dwFlags: DWORD;
      var pNoteID: NOTEID
      ): STATUS; stdcall; far;

function FolderCopy
      (
      hDataDB: DBHANDLE;
      hFolderDB: DBHANDLE;
      FolderNoteID: NOTEID;
      pszName: pchar;
      wNameLen: WORD;
      dwFlags: DWORD;
      var pNoteID: NOTEID
      ): STATUS; stdcall; far;

function FolderDocRemove
      (
      hDataDB: DBHANDLE;
      hFolderDB: DBHANDLE;
      FolderNoteID: NOTEID;
      hTable: LHandle;
      dwFlags: DWORD
      ): STATUS; stdcall; far;

function FolderDocAdd
      (
      hDataDB: DBHANDLE;
      hFolderDB: DBHANDLE;
      FolderNoteID: NOTEID;
      hTable: LHandle;
      dwFlags: DWORD
      ): STATUS; stdcall; far;

function FolderDocRemoveAll
      (
      hDataDB: DBHANDLE;
      hFolderDB: DBHANDLE;
      FolderNoteID: NOTEID;
      dwFlags: DWORD
      ): STATUS; stdcall; far;

function FolderDocCount
      (
      hDataDB: DBHANDLE;
      hFolderDB: DBHANDLE;
      FolderNoteID: NOTEID;
      dwFlags: DWORD;
      var pdwNumDocs: DWORD
      ): STATUS; stdcall; far;

function FolderDelete
      (
      hDataDB: DBHANDLE;
      hFolderDB: DBHANDLE;
      FolderNoteID: NOTEID;
      dwFlags: DWORD
      ): STATUS; stdcall; far;

function FolderMove
      (
      hDataDB: DBHANDLE;
      hFolderDB: DBHANDLE;
      FolderNoteID: NOTEID;
      hParentDB: DBHANDLE;
      ParentNoteID: NOTEID;
      dwFlags: DWORD
      ): STATUS; stdcall; far;


function FolderRename
      (
      hDataDB: DBHANDLE;
      hFolderDB: DBHANDLE;
      FolderNoteID: NOTEID;
      pszName: pchar;
      wNameLen: WORD;
      dwFlags: DWORD
      ): STATUS; stdcall; far;


(******************************************************************************)
{ acl.h }
(******************************************************************************)
const
  ACL_UNIFORM_ACCESS = $00000001;       { Require same ACL in ALL replicas of database }

{ Access Levels  }
const
  ACL_LEVEL_NOACCESS = 0;
  ACL_LEVEL_DEPOSITOR = 1;
  ACL_LEVEL_READER = 2;
  ACL_LEVEL_AUTHOR = 3;
  ACL_LEVEL_EDITOR = 4;
  ACL_LEVEL_DESIGNER = 5;
  ACL_LEVEL_MANAGER = 6;
  ACL_LEVEL_HIGHEST = 6;                { Highest access level }
  ACL_LEVEL_COUNT = 7;                  { Number of access levels }
  ACL_LEVEL_STRINGMAX = 128;            { size to allocate for access descriptors }

{ Named privilege parameters }

const
  ACL_PRIVCOUNT = 80;                   { Number of privilege bits (10 bytes) }
  ACL_PRIVNAMEMAX = 16;                 { Privilege name max (including null) }
  ACL_PRIVSTRINGMAX = 16 + 2;           { Privilege string max }
  ACL_BITPRIVCOUNT = 5;                 { Original "bit" privileges count }
  ACL_BITPRIVS = $1f;                   { Original "bit" privileges mask }
  ACL_BITPRIV_LEFT_PAREN = '(';         { Original "bit" privilege name syntax }
  ACL_BITPRIV_RIGHT_PAREN = ')';

  ACL_SUBGROUP_LEFT_PAREN = '[';        // Subgroup name syntax */
  ACL_SUBGROUP_RIGHT_PAREN = ']';       // Subgroup name syntax */

{  Access level modifier flags }
const
  ACL_FLAG_AUTHOR_NOCREATE = $0001;     { Authors can't create new notes (only edit existing ones) }
  ACL_FLAG_SERVER = $0002;              { Entry represents a Server (V4) }
  ACL_FLAG_NODELETE = $0004;            { User cannot delete notes }
  ACL_FLAG_CREATE_PRAGENT = $0008;      { User can create personal agents (V4) }
  ACL_FLAG_CREATE_PRFOLDER = $0010;     { User can create personal folders (V4) }
  ACL_FLAG_PERSON = $0020;              { Entry represents a Person (V4) }
  ACL_FLAG_GROUP = $0040;               { Entry represents a group (V4) }
  ACL_FLAG_CREATE_FOLDER = $0080;       { User can create and update shared views & folders (V4)
                          This allows an Editor to assume some Designer-level access }
  ACL_FLAG_CREATE_LOTUSSCRIPT = $0100;  { User can create LotusScript }
  ACL_FLAG_PUBLICREADER = $0200;        { User can read public notes }
  ACL_FLAG_PUBLICWRITER = $0400;        { User can write public notes }

{ free bits are here }
const
  ACL_FLAG_ADMIN_READERAUTHOR = $4000;  { Admin server can modify reader and author fields in db }
  ACL_FLAG_ADMIN_SERVER = $8000;        { Entry is administration server (V4) }

{ ACLUpdateEntry flags - Set flag if parameter is being modified }
const
  ACL_UPDATE_NAME = $01;
  ACL_UPDATE_LEVEL = $02;
  ACL_UPDATE_PRIVILEGES = $04;
  ACL_UPDATE_FLAGS = $08;

{ Usernames list structure }
type
  NAMES_LIST = record
      NumNames : WORD;                 {  Number of names in list }
      License : LICENSEID;             {  User's license - now obsolete MUST BE ZERO.}
      Authenticated : DWORD;           {  Authentication flags }
                                       {  Names follow as packed ASCIZ strings }
                                       {  First name is Username. }
                                       {  Subsequent names are ALL the group }
                                       {  names that User is a member of }
                                       {  (directly or indirectly). }
    end;

{ Defines for Authentication flags }
const
  NAMES_LIST_AUTHENTICATED = $0001;     {   Set if names list has been authenticated via Notes }
  NAMES_LIST_PASSWORD_AUTHENTICATED = $0002;
                                       {  Set if names list has been  }
                                       {  authenticated using external }
                                       {  password -- Triggers "maximum }
                                       {  password access allowed" feature }

{ Privileges bitmap structure }
{ WARNING! Privileges 0..4 do not map to roles, you need to subtract 4 in calling to Acl... functions}
type
  ACL_PRIVILEGES = packed record
    BitMask: array[0..9] of byte;
  end;
  PACL_PRIVILEGES = ^ACL_PRIVILEGES;

  TAclEnumEntriesProc = procedure (Param: pointer; Name: pchar;
    AccessLevel: word; Privileges: PACL_PRIVILEGES; AccessFlags: WORD); stdcall;

function ACLIsPrivSet(var privs: ACL_PRIVILEGES; num: integer): boolean;
procedure ACLSetPriv(var privs: ACL_PRIVILEGES; num: integer);
procedure ACLClearPriv(var privs: ACL_PRIVILEGES; num: integer);
procedure ACLInvertPriv(var privs: ACL_PRIVILEGES; num: integer);

function ACLLookupAccess(hACL: LHANDLE;
                         var pNamesList: NAMES_LIST;
                         var retAccessLevel: Word;
                         var retPrivileges: ACL_PRIVILEGES;
                         var retAccessFlags: Word;
                         var rethPrivNames: Handle): STATUS; stdcall; far;

function ACLCreate(var rethACL: Handle): STATUS; stdcall; far;

function ACLAddEntry(hACL: LHANDLE;
                     Name: PChar;
                     AccessLevel: Word;
                     var Privileges: ACL_PRIVILEGES;
                     AccessFlags: Word): STATUS; stdcall; far;

function ACLDeleteEntry(hACL: LHANDLE;
                        Name: PChar): STATUS; stdcall; far;

function ACLUpdateEntry(hACL: LHANDLE;
                        Name: PChar;
                        UpdateFlags: Word;
                        NewName: PChar;
                        NewAccessLevel: Word;
                        var NewPrivileges: ACL_PRIVILEGES;
                        NewAccessFlags: Word): STATUS; stdcall; far;

function ACLEnumEntries(hACL: LHANDLE;
                        EnumFunc: TAclEnumEntriesProc;
                        EnumFuncParam: Pointer): STATUS; stdcall; far;

function ACLGetPrivName(hACL: LHANDLE;
                        PrivNum: Word;
                        retPrivName: PChar): STATUS; stdcall; far;

function ACLSetPrivName(hACL: LHANDLE;
                        PrivNum: Word;
                        PrivName: PChar): STATUS; stdcall; far;


function ACLGetHistory(hACL: LHANDLE;
                       var hHistory: Handle;
                       var HistoryCount: Word): STATUS; stdcall; far;

function ACLGetFlags(hACL: LHANDLE;
                     var Flags: LongInt): STATUS; stdcall; far;

function ACLSetFlags(hACL: LHANDLE;
                     Flags: LongInt): STATUS; stdcall; far;

function ACLGetAdminServer(hList: LHANDLE;
                           ServerName: PChar): STATUS; stdcall; far;

function ACLSetAdminServer(hList: LHANDLE;
                           ServerName: PChar): STATUS; stdcall; far;


(******************************************************************************)
{ from ixedit.h, ixport.h }
(******************************************************************************)
const
  IXFLAG_FIRST    = $01;    //* First time thru flag */
  IXFLAG_LAST     = $02;    //* Last time thru flag */
  IXFLAG_APPEND   = $04;    //* For exports, Append to output file */
  OLDMAXPATH      = 100;    //* Maximum pathname */

type
  TEDITIMPORTDATA = packed record
    OutputFileName: array [0..OLDMAXPATH-1] of char;  //* File to be filled by import with CD records */
    FontID: FONTID;         //* font used at the current caret position */
  end;
  PTEDITIMPORTDATA = ^TEDITIMPORTDATA;

  TEDITEXPORTDATA = packed record
    InputFileName: array [0..OLDMAXPATH-1] of char; //* File to be read by export containing CD records */
    hCompBuffer: LHandle;       //* Handle to composite buffer (V1 Exports) */
    CompLength: DWORD;        //* Length of composite buffer (V1 Exports) */
    HeaderBuffer: HEAD_DESC_BUFFER;
    FooterBuffer: HEAD_DESC_BUFFER;
    PrintSettings: PRINT_SETTINGS;
  end;
  PTEDITEXPORTDATA = ^TEDITEXPORTDATA;

  { defined in IXPORT.H }
  IXENTRYPROC     = function ( var  IXContext       : TEDITIMPORTDATA;
                                    Flags           : WORD;
                                    hModule         : HMODULE;
                                    AltLibraryName  : PChar;
                                    FileName        : PChar ): STATUS; stdcall;

type
  TLnImportProc = function (
    IXContext: PTEDITIMPORTDATA;
    Flags: word;
    hModule: THandle;
    AltLibraryName: pchar;
    FileName: pchar
  ): STATUS; stdcall;

  TLnExportProc = function (
    IXContext: PTEDITEXPORTDATA;
    Flags: WORD;
    hModule: THandle;
    AltLibraryName: pchar;
    FileName: pchar
  ): STATUS; stdcall;

function ConvertItemToText( ItemValue: BLOCKID; ItemValueLength: DWORD;
  LineDelimiter: pChar; CharsPerLine: word; rethBuffer: pHandle; retBufferLength:
  pDWORD; fStripTabs: BOOL ): STATUS; stdcall; far;

(******************************************************************************)
{ Repl.h}
(******************************************************************************)

const
 REPL_OPTION_RCV_NOTES=$00000001;{ Receive notes from server (pull)}
 REPL_OPTION_SEND_NOTES=$00000002;{ Send notes to server (push) }
 REPL_OPTION_CLOSE_SESS=$00000040;{ Close sessions when done }
 REPL_OPTION_PRI_LOW=$00000000;{ Low, Medium, & High priority databases }
 REPL_OPTION_PRI_MED=$00004000;{ Medium & High priority databases only }
 REPL_OPTION_PRI_HI=$00008000;{ High priority databases only }

type
 REPLFILESTATS=Record
  TotalFiles:LongInt;
  FilesCompleted:LongInt;
  NotesAdded:LongInt;
  NotesDeleted:LongInt;
  NotesUpdated:LongInt;
  Successful:LongInt;
  Failed:LongInt;
  NumberErrors:LongInt;
 End;

 REPLSERVSTATS=Record
  Pull:REPLFILESTATS;
  Push:REPLFILESTATS;
  StubsInitialized:LongInt;
  TotalUnreadExchanges:LongInt;
  NumberErrors:LongInt;
  LastError:STATUS;
 End;

Function ReplicateWithServer(PortName:PChar;
 ServerName:PChar;
 Options:Word;
 NumFiles:Word;
 FileList:PChar;
 Var retStats:REPLSERVSTATS):Status;stdcall;far;

(******************************************************************************)
{ colorid.h }
(******************************************************************************)
{ Maximum number of colors that can be handled by Notes. }
const
  MAX_NOTES_COLORS = 240;

{   Number of colors for V3 form background compatablilty }
const
  V3_FORMCOLORS = 21;

{ Standard colors -- so useful they're available by name. }
const
  MAX_NOTES_SOLIDCOLORS = 16;
  NOTES_COLOR_BLACK = 0;
  NOTES_COLOR_WHITE = 1;
  NOTES_COLOR_RED = 2;
  NOTES_COLOR_GREEN = 3;
  NOTES_COLOR_BLUE = 4;
  NOTES_COLOR_MAGENTA = 5;
  NOTES_COLOR_YELLOW = 6;
  NOTES_COLOR_CYAN = 7;
  NOTES_COLOR_DKRED = 8;
  NOTES_COLOR_DKGREEN = 9;
  NOTES_COLOR_DKBLUE = 10;
  NOTES_COLOR_DKMAGENTA = 11;
  NOTES_COLOR_DKYELLOW = 12;
  NOTES_COLOR_DKCYAN = 13;
  NOTES_COLOR_GRAY = 14;
  NOTES_COLOR_LTGRAY = 15;

{ The following FONT_COLOR_XXX are for compatibility with earlier
  revs of the SDK.  New code should use NOTES_COLOR_XXX }
const
  FONT_COLOR_BLACK = NOTES_COLOR_BLACK;
  FONT_COLOR_WHITE = NOTES_COLOR_WHITE;
  FONT_COLOR_RED = NOTES_COLOR_RED;
  FONT_COLOR_GREEN = NOTES_COLOR_GREEN;
  FONT_COLOR_BLUE = NOTES_COLOR_BLUE;
  FONT_COLOR_CYAN = NOTES_COLOR_CYAN;
  FONT_COLOR_YELLOW = NOTES_COLOR_YELLOW;
  FONT_COLOR_MAGENTA = NOTES_COLOR_MAGENTA;


(******************************************************************************)
{ undocumented - by Winalot }
(******************************************************************************)
{$IFNDEF NOTES_R4}
function OSGetIniFileName(retIniName: PChar): Word; stdcall; far;
function OSGetExecutableDirectory(retExeDir: PChar): Word; stdcall; far;
{$ENDIF}

(******************************************************************************)
const ASSISTODS_FLAG_HIDDEN =                   $00000001;  {*  TRUE if manual assistant is hidden *}
const ASSISTODS_FLAG_NOWEEKENDS =   $00000002;  {*  Do not run on weekends *}
const ASSISTODS_FLAG_STOREHIGHLIGHTS =          $00000004;  {*  TRUE if storing highlights *}
const ASSISTODS_FLAG_MAILANDPASTE =   $00000008;  {*  TRUE if this is the V3-style mail and paste macro *}
const ASSISTODS_FLAG_CHOOSEWHENENABLED =  $00000010;  {*

(******************************************************************************)
{ Agents.h - by Daniel }
(******************************************************************************)
const OBJECT_ASSIST_RUNDATA = 8;    {* Assistant run data object *}

// Agent support
const AGENT_LOTUSSCRIPT = 0;
const AGENT_JAVA = 1;
const AGENT_COMPILED_JAVA = 2;
const AGENT_MACRO         = 3;
const AGENT_UNKNOWN       = -1;

//scriptlib support
const LOTUSSCRIPT_LIB = 0;
const JAVA_LIB = 1;
const UNKNOWN_LIB = -1;

const MEM_GROWABLE    = $4000;              // Object may be OSMemRealloc'ed LARGER 

const ASSISTTRIGGER_TYPE_NONE           = 0;  //  Unknown or unavailable
const ASSISTTRIGGER_TYPE_SCHEDULED      = 1;  //  According to time schedule
const ASSISTTRIGGER_TYPE_NEWMAIL        = 2;  //  When new mail delivered
const ASSISTTRIGGER_TYPE_PASTED         = 3;  //  When documents pasted into database
const ASSISTTRIGGER_TYPE_MANUAL         = 4;  //  Manually executed
const ASSISTTRIGGER_TYPE_DOCUPDATE      = 5;  //  When doc is updated
const ASSISTTRIGGER_TYPE_SYNCHNEWMAIL   = 6;  //  Synchronous new mail agent executed by router

const ASSISTSEARCH_TYPE_NONE    = 0;    //  Unknown or unavailable
const ASSISTSEARCH_TYPE_ALL   = 1;  //  All documents in database
const ASSISTSEARCH_TYPE_NEW   = 2;  //  New documents since last run
const ASSISTSEARCH_TYPE_MODIFIED  = 3;  //  New or modified docs since last run
const ASSISTSEARCH_TYPE_SELECTED  = 4;  //  Selected documents
const ASSISTSEARCH_TYPE_VIEW    = 5;  //  All documents in view
const ASSISTSEARCH_TYPE_UNREAD    = 6;  //  All unread documents
const ASSISTSEARCH_TYPE_PROMPT    = 7;  //  Prompt user
const ASSISTSEARCH_TYPE_UI    = 8;  //  Works on the selectable object

const ASSISTSEARCH_TYPE_COUNT   = 9;  //  Total number of search types

const ASSISTINTERVAL_TYPE_NONE    = 0;  //  Unknown
const ASSISTINTERVAL_TYPE_MINUTES = 1;
const ASSISTINTERVAL_TYPE_DAYS    = 2;
const ASSISTINTERVAL_TYPE_WEEK    = 3;
const ASSISTINTERVAL_TYPE_MONTH   = 4;

const DESIGNER_VERSION_ITEM = '$DesignerVersion';
const ASSIST_RUNINFO_ITEM = '$AssistRunInfo'; //  Run information object
const ASSIST_ACTION = '$AssistAction';
const ASSIST_TYPE_ITEM = '$AssistType';
const ASSIST_VERSION_ITEM = '$AssistVersion';
const ASSIST_LASTRUN_ITEM = '$AssistLastRun';
const ASSIST_DOCCOUNT_ITEM = '$AssistDocCount';
const ASSIST_FLAGS_ITEM = '$AssistFlags';
const ASSIST_TRIGGER_ITEM = '$AssistTrigger';
const ASSIST_INFO_ITEM = '$AssistInfo';
const ASSIST_QUERY_ITEM = '$AssistQuery';
const ASSIST_ACTION_ITEM = '$AssistAction';
const ASSIST_EXACTION_ITEM = '$AssistAction_Ex';
const MAXALPHATIMEDATE  =  80;
const DESIGN_FLAGS_MAX = 32;
const AGENT_LAST_RUN = 0;
const AGENT_DOCUMENTS_PROCESSED = 1;
const AGENT_EXIT_CODE = 2;
const AGENT_LOG = 3;
const FIELD_PUBLICACCESS = '$PublicAccess'; // from stdNames.h

type Design_max_array = array [0..DESIGN_FLAGS_MAX] of char;

type CDACTIONHEADER = packed record
  Header: BSIG;
end;

type CDACTIONLOTUSSCRIPT = packed record
  Header: WSIG;
  dwFlags:DWORD;
  dwScriptLen:DWORD;
  {*  Script follows *}
end;

type CDACTIONFORMULA = packed record
  Header: WSIG;
  dwFlags:DWORD;
        wFormulaLen:WORD;
  {*  Formula follows *}
end;
// Java Action
type CDACTIONJAVAAGENT = packed record
  Header: WSIG;
  wClassNameLen: WORD;       // Agent name length
  wCodePathLen: WORD;
        wFileListBytes: WORD;
  wLibraryListBytes: WORD;
  wSpare: array[0..1] of WORD;
  dwSpare: array[0..1] of DWORD;
        {*  Strings follows *}
end;

type CDQUERYHEADER = packed record
        Header: BSIG;
  dwFlags: DWord;       //  Flags for query
end;

type ODS_ASSISTSTRUCT = packed record
        wVersion: WORD;       //  Structure version

  wTriggerType:WORD;      //  Type of trigger
  wSearchType:WORD;     //  Type of search
  wIntervalType:WORD;     //  Type of interval
  wInterval:WORD;       //  Interval
  dwTime1:DWORD;        //  depends on interval type
  dwTime2:DWORD;        //  depends on interval type

  StartTime:TIMEDATE;     //  Agent does not run before this time
  EndTime:TIMEDATE;     //  Agent does not run after this time

  dwFlags:DWORD;

  dwSpare: array [0..15] of DWORD;
end;

type ODS_ASSISTRUNINFO = packed record
  LastRun: TIMEDATE;
  dwProcessed: DWORD;
  AssistMod: TIMEDATE;
  DbID: TIMEDATE;
  dwExitCode: longint;
  dwSpare: packed array[0..4] of Dword;
end;

ODS_ASSISTRUNOBJECTHEADER = packed record
  dwFlags: DWORD;
  wEntries: WORD;
  wSpare: WORD;
end;

ODS_ASSISTRUNOBJECTENTRY = packed record
  dwLength: DWORD;
  dwFlags: DWORD;
end;

type
       HAGENT=HANDLE;
       HAGENTCTX= HANDLE;


function AgentIsEnabled(hAgent: HAGENT): BOOL; stdcall; far;
function AgentOpen (hSrcDB: DBHANDLE; AgentNoteID: NOTEID; var hAgent: HANDLE): STATUS; stdcall; far;
function AgentCreateRunContext(hAgent: HANDLE; pReserved: PChar; dwFlags: dword; var rethContext: HANDLE): STATUS; stdcall; far;
function AgentRun(hAgent: HAGENT; hAgentCtx: HAGENTCTX; hSelection: HANDLE; dwFlags: dword): STATUS; stdcall; far;
function AgentDestroyRunContext(hAgentCtx: HAGENTCTX): STATUS;stdcall; far;
function AgentSetDocumentContext(hAgentCtx: HAGENTCTX; hNote: HANDLE): STATUS ;stdcall; far;
function AgentClose(hAgent: HANDLE): STATUS;stdcall; far;

{$IFNDEF NOTES_R4}
function NSFNoteLSCompile(hDb: HANDLE; hNote: HANDLE; dwFlags: DWord): STATUS; stdcall; far;
function AgentLSTextFormat(hSrc: Handle;var hDest;var hErrs: HANDLE; dwFlags: dWord; var phData: HANDLE): STATUS; stdcall; far;
function AgentDelete(hSrc: LHandle): STATUS; stdcall; far;
{$ENDIF}

function NSFDbGetObjectSize(
  hDB:DBHANDLE;
  ObjectID:Integer;
  ObjectType:WORD;
  var retSize:DWORD;
  var retClass:WORD;
  var retPrivileges:WORD): STATUS;stdcall; far;

function NSFDbReadObject(
  hDB: DBHANDLE;
  ObjectID: Integer;
  Offset,Length:Dword;
  var rethBuffer: HANDLE): STATUS;stdcall; far;

function ConvertTIMEDATEToText(
  IntlFormat: pchar;
  var TextFormat: TFMT;
  var InputTime: TIMEDATE;
  var retTextBuffer: Char;
  TextBufferLength: Integer;
  var retTextLength:WORD): STATUS; stdcall; far;

function ConvertTextToTIMEDATE(
	IntlFormat: pchar;
	TextFormat: pTFMT;
	Text: pchar;
	MaxLength: Word;
	var retTIMEDATE: TIMEDATE): STATUS; stdcall; far;

function NSFDbAllocObject(
  hDB:DBHANDLE;
  dwSize:DWORD;
  aClass:WORD;
  Privileges:WORD;
  var retObjectID:DWORD): STATUS;stdcall; far;

function NSFDbFreeObject(
  hDB:DBHANDLE;
  ObjectID:DWORD): STATUS;stdcall; far;

function NSFDbWriteObject(
  hDB: DBHANDLE;
  ObjectID: DWord;
        hBuffer:HANDLE;
      Offset:DWORD;
  Length:DWORD): STATUS;stdcall; far;

const STATPKG_OS = 'OS';
const STATPKG_STATS = 'Stats';
const STATPKG_OSMEM = 'Mem';
const STATPKG_OSSEM = 'Sem';
const STATPKG_OSSPIN = 'Spin';
const STATPKG_OSFILE = 'Disk';
const STATPKG_SERVER = 'Server';
const STATPKG_REPLICA = 'Replica';
const STATPKG_MAIL = 'Mail';
const STATPKG_MAILBYDEST = 'MailByDest';
const STATPKG_COMM = 'Comm';
const STATPKG_NSF = 'Database';
const STATPKG_NIF = 'Database';
const STATPKG_TESTNSF = 'Testnsf';
const STATPKG_OSIO = 'IO';
const STATPKG_NET = 'NET';
const STATPKG_OBJSTORE = 'Object';
const STATPKG_AGENT = 'Agent';	        	// used by agent manager */
const STATPKG_WEB = 'Web';			// used by Web retriever */
const STATPKG_CAL = 'Calendar';		        // used by schedule manager */
const STATPKG_SMTP = 'SMTP';			// Used by SMTP listener */
const STATPKG_LDAP = 'LDAP';			// Used by the LDAP Server */
const STATPKG_NNTP = 'NNTP';			// Used by the NNTP Server */
const STATPKG_ICM = 'ICM';			// Used by the ICM Server */

const STATPKG_MONITOR = 'Monitor';
const STATPKG_POP3 = 'POP3';			// Used by the POP3 Server */

// Value type constants

const VT_LONG =	0;
const VT_TEXT	= 1;
const VT_TIMEDATE = 2;
const VT_NUMBER = 3;

type
STATTRAVPROC=function(
        EnumRoutineParameter:Pointer;
        Facility: PChar;
        StatName: PChar;
        ValueType: Word;
        Value: Pointer): STATUS; stdcall;

function StatTraverse(
	Facility: PChar;
        StatName: PChar;
	EnumRoutine: STATTRAVPROC;
	Context: Pointer): Status; stdcall; far;

function StatToText(
	Facility: PChar;
	StatName: PChar;
	ValueType: Word;
	Value: Pointer;
	NameBuffer: PChar;
	NameBufferLen: Word;
	ValueBuffer: PChar;
      	ValueBufferLen: Word): Status; stdcall; far;

// END added by Daniel
(******************************************************************************)


(******************************************************************************)
{ Internal link definition - by Olaf                                           }
(******************************************************************************)
type
  TLinkType = (rtlUnknown,
               rtlAnchorLink,
               rtlDocumentLink,
               rtlDatabaseLink,
               rtlViewLink,
               rtlHotSpotLink);
  LinkDef = packed record
    aFile : TIMEDATE; // File's replica ID
    View : UNID;      // View's Note Creation TIMEDATE
    Note : UNID;      // Note's Creation TIMEDATE
    Comment : string; // comment of doclink
    Hint : string;    // server
    Anchor : string;  // anchor text
    LinkType : TLinkType;
  end;
  PLinkDef = ^LinkDef;
  // Record LinkDefI in NotesRTF unit mirrors this one


(******************************************************************************)
{ viewfmt.h }
(******************************************************************************)
const
  VIEW_VIEW_FORMAT_ITEM  = '$ViewFormat';

{ View on-disk format definitions }
const
  VIEW_FORMAT_VERSION = 1;
  VIEW_COLUMN_FORMAT_SIGNATURE = $4356;
  VIEW_COLUMN_FORMAT_SIGNATURE2 = $4357;
  VIEW_CLASS_TABLE = (0 shl 4);
  VIEW_CLASS_CALENDAR = (1 shl 4);
  VIEW_CLASS_MASK = $F0;
  CALENDAR_TYPE_DAY = 0;
  CALENDAR_TYPE_WEEK = 1;
  CALENDAR_TYPE_MONTH = 2;
  VIEW_STYLE_TABLE = VIEW_CLASS_TABLE;
  VIEW_STYLE_DAY = (VIEW_CLASS_CALENDAR + 0);
  VIEW_STYLE_WEEK = (VIEW_CLASS_CALENDAR + 1);
  VIEW_STYLE_MONTH = (VIEW_CLASS_CALENDAR + 2);

{ View table format descriptor.  Followed by VIEW_COLUMN_FORMAT }
{ descriptors; one per column.  The column format descriptors are followed }
{ by the packed item name, title, formula, and constant values.  }
{ All of this is followed by a VIEW_TABLE_FORMAT2 data structure that }
{ is only present in views saved in V2 or later. }
{ All descriptors and values are packed into one item named $VIEWFORMAT. }
type
  VIEW_FORMAT_HEADER = packed record
    Version: BYTE;              //* Version number */
    ViewStyle: BYTE;            //* View Style - Table,Calendar */
  end;
  PVIEW_FORMAT_HEADER = ^VIEW_FORMAT_HEADER;

const
  VIEW_TABLE_FLAG_COLLAPSED = $0001;{ Default to fully collapsed }
  VIEW_TABLE_FLAG_FLATINDEX = $0002;{ Do not index hierarchically }
                                   { If FALSE, MUST have }
                                   { NSFFormulaSummaryItem($REF) }
                                   { as LAST item! }
  VIEW_TABLE_FLAG_DISP_ALLUNREAD = $0004;
                                   { Display unread flags in margin at ALL levels }
  VIEW_TABLE_FLAG_CONFLICT = $0008; { Display replication conflicts }
                                   { If TRUE, MUST have }
                                   { NSFFormulaSummaryItem($Conflict) }
                                   { as SECOND-TO-LAST item! }
  VIEW_TABLE_FLAG_DISP_UNREADDOCS = $0010;
                                   { Display unread flags in margin for documents only }
  VIEW_TABLE_GOTO_TOP_ON_OPEN = $0020;
                                   { Position to top when view is opened. }
  VIEW_TABLE_GOTO_BOTTOM_ON_OPEN = $0040;
                                   { Position to bottom when view is opened. }
  VIEW_TABLE_ALTERNATE_ROW_COLORING = $0080;
                                   { Color alternate rows. }
  VIEW_TABLE_HIDE_HEADINGS = $0100; { Hide headings. }
  VIEW_TABLE_HIDE_LEFT_MARGIN = $0200;
                                   { Hide left margin. }
  VIEW_TABLE_SIMPLE_HEADINGS = $0400;
                                   { Show simple (background color) headings. }
  VIEW_TABLE_VARIABLE_LINE_COUNT = $0800;
                                   { TRUE if LineCount is variable (can be reduced as needed). }

  { Refresh flags.

    When both flags are clear, automatic refresh of display on update
    notification is disabled.  In this case, the refresh indicator will
    be displayed.

    When VIEW_TABLE_GOTO_TOP_ON_REFRESH is set, the view will fe refreshed from
    the top row of the collection (as if the user pressed F9 and Ctrl-Home).

    When VIEW_TABLE_GOTO_BOTTOM_ON_REFRESH is set, the view will be refreshed
    so the bottom row of the collection is visible (as if the user pressed F9
    and Ctrl-End).

    When BOTH flags are set (done to avoid using another bit in the flags),
    the view will be refreshed from the current top row (as if the user
    pressed F9). }

  VIEW_TABLE_GOTO_TOP_ON_REFRESH = $1000;
                                   { Position to top when view is refreshed. }
  VIEW_TABLE_GOTO_BOTTOM_ON_REFRESH = $2000;
                                   { Position to bottom when view is refreshed. }
  { More flag(s). }
  VIEW_TABLE_EXTEND_LAST_COLUMN = $4000;
                                   { TRUE if last column should be extended to fit the window width. }

type
  VIEW_TABLE_FORMAT = packed record
      Header : VIEW_FORMAT_HEADER;
      Columns : WORD;
      ItemSequenceNumber : WORD;
      Flags : WORD;
      spare2 : WORD;
    end;
  PVIEW_TABLE_FORMAT = ^VIEW_TABLE_FORMAT;

{  Additional (since V2) format info.  This structure follows the
  variable length strings that follow the VIEW_COLUMN_FORMAT structres }
const
  VALID_VIEW_FORMAT_SIG = $2BAD;
  VIEW_TABLE_MAX_LINE_COUNT = 10;
  VIEW_TABLE_SINGLE_SPACE = 0;
  VIEW_TABLE_ONE_POINT_25_SPACE = 1;
  VIEW_TABLE_ONE_POINT_50_SPACE = 2;
  VIEW_TABLE_ONE_POINT_75_SPACE = 3;
  VIEW_TABLE_DOUBLE_SPACE = 4;
  VIEW_TABLE_COLOR_MASK = $00FF;     { color is index into 240 element array }
  VIEW_TABLE_HAS_LINK_COLUMN = $01; { TRUE if a link column has been specified for a web browser. }
  VIEW_TABLE_HTML_PASSTHRU = $02;   { TRUE if line entry text should be treated as HTML by a web browser. }

type
  VIEW_TABLE_FORMAT2 = packed record
      Length : WORD;
      BackgroundColor : WORD;           { Color of view's background. Pre-V4 compatible }
      V2BorderColor : WORD;             { Archaic! Color of view's border lines. }
      TitleFont : FONTID;               { Title and borders }
      UnreadFont : FONTID;              { Unread lines }
      TotalsFont : FONTID;              { Totals/Statistics }
      AutoUpdateSeconds : WORD;         { interval b/w auto updates (zero for no autoupdate) }
      AlternateBackgroundColor : WORD;  { Color of view's background for alternate rows. }
                                        { When wSig == VALID_VIEW_FORMAT_SIG, rest of struct is safe to use.  Bug
                                          in versions prior to V4 caused spare space in this structure to contain
                                          random stuff. }
      wSig : WORD;
      LineCount : BYTE;                 { Number of lines per row.  1, 2, etc. }
      Spacing : BYTE;                   { Spacing.  VIEW_TABLE_XXX_SPACE. }
      BackgroundColorExt : WORD;        { Palette Color of view's background. }
      HeaderLineCount : BYTE;           { Lines per header. }
      Flags1 : BYTE;                    { Spares.  Will be zero when wSig == VALID_VIEW_FORMAT_SIG. }
      Spare : array[0..4 - 1] of WORD;
    end;
    
type
  VIEW_DAY_FORMAT = packed record
    Header : VIEW_FORMAT_HEADER;
  end;

  VIEW_WEEK_FORMAT = packed record
      Header : VIEW_FORMAT_HEADER;
    end;

  VIEW_MONTH_FORMAT = packed record
      Header : VIEW_FORMAT_HEADER;
    end;

{   Calendar View Format Information.  Introduced in build 141 (for 4.2).
  This is in Calendar Style Views only. }

const
  VIEW_CALENDAR_FORMAT_VERSION = 1;

const
  VIEW_CAL_FORMAT_TWO_DAY = $01;
  VIEW_CAL_FORMAT_ONE_WEEK = $02;
  VIEW_CAL_FORMAT_TWO_WEEKS = $04;
  VIEW_CAL_FORMAT_ONE_MONTH = $08;
  VIEW_CAL_FORMAT_ONE_YEAR = $10;
  VIEW_CAL_FORMAT_ALL = $ff;
  CAL_DISPLAY_CONFLICTS = $0001;    { Display Conflict marks }
  CAL_ENABLE_TIMESLOTS = $0002;     { Disable Time Slots }
  CAL_DISPLAY_TIMESLOT_BMPS = $0004;{ Show Time Slot Bitmaps }

type
  VIEW_CALENDAR_FORMAT = packed record
      Version : BYTE;
                                       { Version Number }
      Formats : BYTE;
                                       { Formats supported by this view VIEW_CAL_FORMAT_XXX.}
      DayDateFont : FONTID;
                                       { Day and Date display }
      TimeSlotFont : FONTID;
                                       { Time Slot display }
      HeaderFont : FONTID;
                                       { Month Headers }
      DaySeparatorsColor : WORD;
                                       { Lines separating days }
      TodayColor : WORD;
                                       { Color Today is displayed in }
      wFlags : WORD;
                                       { Misc Flags }
      BusyColor : WORD;
                                       { Color busy times are displayed in }
      wTimeSlotStart : WORD;
                                       { TimeSlot start time (in minutes from midnight) }
      wTimeSlotEnd : WORD;
                                       { TimeSlot end time (in minutes from midnight) }
      wTimeSlotDuration : WORD;
                                       { TimeSlot duration (in minutes) }
      unused : WORD;
      Spare : array[0..7 - 1] of DWORD;
    end;

type
  PVIEW_CALENDAR_FORMAT = ^VIEW_CALENDAR_FORMAT;

{ View column format descriptor.  One per column. }
    const
      VCF1_S_Sort = 0                  { Add column to sort }
      ;
      VCF1_M_Sort = $0001;
      VCF1_S_SortCategorize = 1        { Make column a category }
      ;
      VCF1_M_SortCategorize = $0002;
      VCF1_S_SortDescending = 2        { Sort in descending order (ascending if FALSE) }
      ;
      VCF1_M_SortDescending = $0004;
      VCF1_S_Hidden = 3                { Hidden column }
      ;
      VCF1_M_Hidden = $0008;
      VCF1_S_Response = 4              { Response column }
      ;
      VCF1_M_Response = $0010;
      VCF1_S_HideDetail = 5            { Do not show detail on subtotalled columns }
      ;
      VCF1_M_HideDetail = $0020;
      VCF1_S_Icon = 6                  { Display icon instead of text }
      ;
      VCF1_M_Icon = $0040;
      VCF1_S_NoResize = 7              { Resizable at run time. }
      ;
      VCF1_M_NoResize = $0080;
      VCF1_S_ResortAscending = 8       { Resortable in ascending order. }
      ;
      VCF1_M_ResortAscending = $0100;
      VCF1_S_ResortDescending = 9      { Resortable in descending order. }
      ;
      VCF1_M_ResortDescending = $0200;
      VCF1_S_Twistie = 10              { Show twistie if expandable. }
      ;
      VCF1_M_Twistie = $0400;
      VCF1_S_ResortToView = 11         { Resort to a view. }
      ;
      VCF1_M_ResortToView = $0800;
      VCF1_S_SecondResort = 12         { Secondary resort column set. }
      ;
      VCF1_M_SecondResort = $1000;
      VCF1_S_SecondResortDescending = 13
                                       { Secondary column resort descending (ascending if clear). }
      ;
      VCF1_M_SecondResortDescending = $2000;
      VCF1_S_CaseInsensitiveSort = 14  { Case insensitive sorting. }
      ;
      VCF1_M_CaseInsensitiveSort = $4000;
      VCF1_S_AccentInsensitiveSort = 15{ Accent insensitive sorting. }
      ;
      VCF1_M_AccentInsensitiveSort = $8000;
      VCF1_M_spare = $c000             { Spare flags. }
      ;

      VCF2_S_DisplayAlignment = 0      { Display alignment - VIEW_COL_ALIGN_XXX }
      ;
      VCF2_M_DisplayAlignment = $0003;
      VCF2_S_SubtotalCode = 2          { Subtotal code (NIF_STAT_xxx) }
      ;
      VCF2_M_SubtotalCode = $003c;
      VCF2_S_HeaderAlignment = 6       { Header alignment - VIEW_COL_ALIGN_XXX }
      ;
      VCF2_M_HeaderAlignment = $00c0;
      VCF2_S_SortPermute = 8           { Make column permuted if multi-valued }
      ;
      VCF2_M_SortPermute = $0100;
      VCF2_S_SecondResortUniqueSort = 9{ Secondary resort column props different from column def.}
      ;
      VCF2_M_SecondResortUniqueSort = $0200;
      VCF2_S_SecondResortCategorized = 10
                                       { Secondary resort column categorized. }
      ;
      VCF2_M_SecondResortCategorized = $0400;
      VCF2_S_SecondResortPermute = 11  { Secondary resort column permuted. }
      ;
      VCF2_M_SecondResortPermute = $0800;
      VCF2_S_SecondResortPermutePair = 12
                                       { Secondary resort column pairwise permuted. }
      ;
      VCF2_M_SecondResortPermutePair = $1000;
      VCF2_S_ShowValuesAsLinks = 13    { Show values as links when viewed by web browsers. }
      ;
      VCF2_M_ShowValuesAsLinks = $2000;
      VCF2_S_Available2 = 14           {  }
      ;
      VCF2_M_Available2 = $4000;
      VCF2_S_Available3 = 15           {  }
      ;
      VCF2_M_Available3 = $8000;

type
  VIEW_COLUMN_FORMAT = packed record
      Signature : WORD;
                                       { VIEW_COLUMN_FORMAT_SIGNATURE }
      Flags1 : WORD;

      ItemNameSize : WORD;
                                       { Item name string size }
      TitleSize : WORD;
                                       { Title string size }
      FormulaSize : WORD;
                                       { Compiled formula size }
      ConstantValueSize : WORD;
                                       { Constant value size }
      DisplayWidth : WORD;
                                       { Display width - 1/8 ave. char width units }
      FontID : FONTID;
                                       { Display font ID }
      Flags2 : WORD;
      NumberFormat : NFMT;
                                       { Number format specification }
      TimeFormat : TFMT;
                                       { Time format specification }
      FormatDataType : WORD;
                                       { Last format data type }
      ListSep : WORD;
                                       { List Separator }
    end;
    PVIEW_COLUMN_FORMAT = ^VIEW_COLUMN_FORMAT;

{ View column display alignment.  }

{   Note: order and values are assumed in VIEW_ALIGN_XXX_ID's. }

const
  VIEW_COL_ALIGN_LEFT = 0              { Left justified }
  ;
  VIEW_COL_ALIGN_RIGHT = 1             { Right justified }
  ;
  VIEW_COL_ALIGN_CENTER = 2            { Centered }
  ;

{ Simple format data types, used to initialize dialog box to last "mode". }

const
  VIEW_COL_NUMBER = 0;
  VIEW_COL_TIMEDATE = 1;
  VIEW_COL_TEXT = 2;

{ Extended View column format descriptor.  One per column as of Notes V4.
  NOTE:  If you add variable data to this structure, store the packed,
  variable data AFTER the array of structures. }

type
  VIEW_COLUMN_FORMAT2 = packed record
      Signature : WORD;
                                       { VIEW_COLUMN_FORMAT_SIGNATURE2 }
      HeaderFontID : FONTID;
                                       { FontID of column header. }
      ResortToViewUNID : UNID;
                                       { UNID of view to switch to. }
      wSecondResortColumnIndex : WORD;
                                       { 0 based index of secondary resort column. }
      wSpare : WORD;
      dwSpare : array[0..5 - 1] of DWORD;
    end;

(******************************************************************************)
{ dbdrv.h }

const
  MAX_DBDRV_NAME = 5;                   { Max. length of database driver class name }
  DBDRV_PREFIX = 'DB';                  { pre-pended to database driver name }

{ Database function definitions common to all databases, although
  the actual arguments may differ for any given function. }

const
  DB_LOOKUP = 0;                        { look something up in a database }
  DB_COLUMN = 1;                        { return an entire column from a database }
  DB_DBEXT = 2;                         { Extended function }

{ Driver vectors }

type
  PPointer = ^Pointer;
  THDBDSESSION = Pointer;               { DBD session handle }
  
  TDBOPENBYIDPROC = function (ReplicaID : DBID; FileTitle : PChar;
    hNames : HANDLE; var rethDB : DBHANDLE) : STATUS;

  // KOL:
  // Due to Delphi rules, functions in the record cannot use it as a parameter
  // (type is not defined). Therefore all of them have vec: Pointer parameter
  // which should be casted to PDBVEC
  
  TDBVEC = packed record
    ClassName : array[0..MAX_DBDRV_NAME + 1 - 1] of Char;     { name of driver class + '\0' }
    hModule : HMODULE;                                        { hModule of loaded module }

      { Do per-process initialization routine.  This is called just after
        the LoadLibrary, and is the first entry point in the library.
        When this entry point is called, all of the other entry point
        vectors are filled in by this routine. }
    Init : function (vec: Pointer) : STATUS;

      { Per-process termination routine, ASSUMING that all open sessions
        for this context have been closed by the time this is called.  This
        is called just prior to the FreeLibrary. }
    Term : function (vec: Pointer) : STATUS;

      { Open a session.  Any databases opened as a side-effect of Functions
        performed on this session will gather up their context in the
        hSession returned by this routine. }
      Open : function (vec: Pointer; var rethSession : THDBDSESSION) : STATUS;

      { Close a session, and as a side-effect all databases whose context
        has been built up in hSession. }
      Close : function (vec: Pointer; hSession : THDBDSESSION) : STATUS;

      { Set auxiliary context, used principally when called from Desk }
      SetOpenContext : function (vec: Pointer; hSession : THDBDSESSION; DefaultDbName : PChar;
        Proc : TDBOPENBYIDPROC; hNames : HANDLE; hParentWnd : DWORD) : STATUS;

        { Perform a function on a session.  If any databases must be opened
        as a side-effect of this function, gather context into hSession
        so that it may be later deallocated/closed in Close. }
      DoFunction : function (vec: Pointer; hSession : THDBDSESSION; wFunction : WORD;
        argc : WORD; argl : PDWORD; argv : PPointer; rethResult : PHANDLE;
        retResultLength : PDWORD) : STATUS;

      { Flags }
      fUpdateIfModified : byte;     { TRUE if we want UpdateCollections if modified }
    end;
    PDBVEC = ^TDBVEC;

(******************************************************************************)
{ oleods.h }
{+// Name of a form autolaunch item. This optional item is created when }
{=designing a Notes form using the auto launch options. }

const
  FORM_AUTOLAUNCH_ITEM = '$AUTOLAUNCH';

{+// Name of an OLE object item. One of these is created for every }
{-OLE embedded object that exists in a Notes document. This item }
{-is used to access OLE objects witout having to parse the }
{=Rich Text item within the document to find an OLE CD record }

const
  OLE_OBJECT_ITEM = '$OLEOBJINFO';


{+// On-disk structure of an OLE GUID }

type
  OLE_GUID = TGUID;

{+// Format of an on-disk autolaunch item. Most of the info contained in }
{-this structure refer to OLE autolaunching behaviors, but there are }
{=some }

type
  CDDOCAUTOLAUNCH = packed record
    Header: WSIG;
{= Signature and length of this record }
    ObjectType: LongInt;
{= Type of object to launch, see OBJECT_TYPE_??? }
    HideWhenFlags: LongInt;
{= HIDE_ flags below }
    LaunchWhenFlags: LongInt;
{= LAUNCH_ flags below }
    OleFlags: LongInt;
{= OLE Flags below }
    CopyToFieldFlags: LongInt;
{= Field create flags below }
    Spare1: LongInt;
    Spare2: LongInt;
    FieldNameLength: Word;
{= If named field, length of field name }
    OleObjClass: OLE_GUID;
{= ClassID GUID of OLE object, if create new }
{+// Field Name, if used, goes here*/ }
  end {CDDOCAUTOLAUNCH};


{+// Autolaunch Object type flags*/ }

const
  AUTOLAUNCH_OBJTYPE_NONE = $00000000;
const
  AUTOLAUNCH_OBJTYPE_OLE_CLASS = $00000001; {/* OLE Class ID (GUID)*/}
const
  AUTOLAUNCH_OBJTYPE_OLEOBJ = $00000002; {/* First OLE Object*/}
const
  AUTOLAUNCH_OBJTYPE_DOCLINK = $00000004; {/* First Notes doclink*/}
const
  AUTOLAUNCH_OBJTYPE_ATTACH = $00000008; {/* First Attachment*/}
const
  AUTOLAUNCH_OBJTYPE_URL = $00000010; {/* AutoLaunch the url in the URL field*/}


{+// Hide-when flags*/ }

const
  HIDE_OPEN_CREATE = $00000001; {/* Hide when opening flags*/}
const
  HIDE_OPEN_EDIT = $00000002;
const
  HIDE_OPEN_READ = $00000004;
const
  HIDE_CLOSE_CREATE = $00000008; {/* Hide when closing flags*/}
const
  HIDE_CLOSE_EDIT = $00000010;
const
  HIDE_CLOSE_READ = $00000020;

{+// Launch-when flags*/ }

const
  LAUNCH_WHEN_CREATE = $00000001;
const
  LAUNCH_WHEN_EDIT = $00000002;
const
  LAUNCH_WHEN_READ = $00000004;

{+// OLE Flags*/ }

const
  OLE_EDIT_INPLACE = $00000001;
const
  OLE_MODAL_WINDOW = $00000002;
const
  OLE_ADV_OPTIONS = $00000004;

{+// Field Location Flags*/ }

const
  FIELD_COPY_NONE = $00000001; {/* Don't copy obj to any field (V3 compatabile)*/}
const
  FIELD_COPY_FIRST = $00000002; {/* Copy obj to first rich text field*/}
const
  FIELD_COPY_NAMED = $00000004; {/* Copy obj to named rich text field*/}


type
  CDOLEOBJ_INFO = packed record
    Header: WSIG;
{= Signature and length of this record }
    FileObjNameLength: Word;
{- Length of name of extendable $FILE object containing }
{=object data }
    DescriptionNameLength: Word;
{= Length of description of object }
    FieldNameLength: Word;
{= Length of field name in which object resides }
    TextIndexObjNameLength: Word;
{- Length of name of the $FILE object containing LMBCS text }
{=for object }
    OleObjClass: OLE_GUID;
{= OLE ClassID GUID of OLE object }
    StorageFormat: Word;
{= See below OLE_STG_FMT_??? }
    DisplayFormat: Word;
{= Object's display format within document, DDEFORMAT_??? }
    Flags: LongInt;
{= Object information flags, see OBJINFO_FLAGS_??? }
{ WORD StorageFormatAppearedIn; /* Version # of Notes, high byte=major, low byte=minor, }
    HTMLDataLength: WORD;
{= Length of HTML data for object }
    Reserved2: Word;
{= Unused, must be 0 }
    Reserved3: Word;
{= Unused, must be 0 }
    Reserved4: LongInt;
{= Unused, must be 0 }
{+// The variable length portions go here in the following order: }
{-FileObjectName }
{-DescriptionName }
{-Field Name in Document in which this object resides }
{-Full Text index $FILE object name }
{-HTML Data }
{= }
  end {CDOLEOBJ_INFO};

const
  OBJINFO_FLAGS_SCRIPTED = $00000001; {/* Object is scripted*/}
const
  OBJINFO_FLAGS_RUNREADONLY = $00000002; {/* Object is run in read-only mode*/}
const
  OBJINFO_FLAGS_CONTROL = $00000004; {/* Object is a control*/}
const
  OBJINFO_FLAGS_FITTOWINDOW = $00000008; {/* Object is sized to fit to window*/}
const
  OBJINFO_FLAGS_FITBELOWFIELDS = $00000010; {/* Object is sized to fit below fields*/}
const
  OBJINFO_FLAGS_UPDATEFROMDOCUMENT = $00000020; {/* Object is to be updated from document*/}
const
  OBJINFO_FLAGS_INCLUDERICHTEXT = $00000040; {/* Object is to be updated from document*/}
const
  OBJINFO_FLAGS_ISTORAGE_ISTREAM = $00000080; {/* Object is stored in IStorage/IStream}

const
  OLE_STG_FMT_STRUCT_STORAGE = 1; {/* OLE 'Docfile' structured storage format}

{$IFNDEF NOTES_R4}

{+// HTML OBJECT Event Entry -----------------------------------------------------*/ }
type
  OLEOBJHTMLEVENT = packed record
    wLength: Word;
{- Size of this structure including both fixed and }
{=variable sections }
    wsNameLength: Word;
{= Length of Name }
    wsScriptLength: Word;
{= Length of Script }
    wReserved1: Word;
{= Reserved }
    wReserved2: Word;
{= Reserved }
{+// The variable length portions go here in the following order: }
{-Name }
{-Script }
{= }
  end {OLEOBJHTMLEVENT};

{+// HTML OBJECT Param Entry -----------------------------------------------------*/ }
type
  OLEOBJHTMLPARAM = packed record
    wLength: Word;
{- Size of this structure including both fixed and }
{=variable sections }
    wsDataFldLength: Word;
{= Length of Data Field }
    wsDataFmtsLength: Word;
{= Length of Data Formats }
    wsDataSrcLength: Word;
{= Length of Data Sourcet }
    wsNameLength: Word;
{= Length of Name }
    wsTypeLength: Word;
{= Length of Type }
    wsValueLength: Word;
{= Length of Value }
    wsValueTypeLength: Word;
{= Length of Value Type }
    wReserved1: Word;
{= Reserved }
    wReserved2: Word;
{= Reserved }
{+// The variable length portions go here in the following order: }
{-Data Field - column name ffrom the data source object }
{-Data Formats - indicates whether bound data is plain text or HTML }
{ Data Source - "#ID" of the data source object }
{-Name - name of this parameter }
{-Type - internal media type }
{-Value - value associated with parameter }
{-Value Type - type of value (data, ref, object) }
{= }
  end {OLEOBJHTMLPARAM};


{+// OLE Object HTML Data -------------------------------------------------------------*/ }
type
  OLEOBJHTMLDATA = packed record
    wLength: Word;
{- Size of this structure including both fixed and }
{=variable sections }
    wsURLBaseLength: Word;
{= Length of Base URL }
    wsURLCodeBaseLength: Word;
{= Length of CodeBase URL }
    wsMIMETypeCodeLength: Word;
{= Length of MIME CodeType }
    wsURLDataLength: Word;
{= Length of Data URL }
    wsDataFldLength: Word;
{= Length of Data Field name }
    wsDataSrcLength: Word;
{= Length of Data Source ID }
    dwFlags: LongInt;
{= Flags }
    wsLangLength: Word;
{= Length of Language }
    wsNameLength: Word;
{= Length of Name }
    wsMIMETypeDataLength: Word;
{= Length of MIME Type }
    wcEvents: Word;
{= Number of events }
    wcParams: Word;
{= Number of params }
    wHeight: Word;
{= Height of object }
    wWidth: Word;
{= Width of object }
    wReserved1: Word;
{= Reserved }
    wReserved2: Word;
{= Reserved }
    wReserved3: Word;
{= Reserved }
    wReserved4: Word;
{= Reserved }
    wReserved5: Word;
{= Reserved }
    wReserved6: Word;
{= Reserved }
{+// The variable length portions go here in the following order: }
{-URLBase - Base URL }
{-CodeBase - URL that references where to find implementation of the object. }
{-CodeType - MIME type of the code referenced by CLSID }
{-Data - URL of the data to be loaded }
{-Data Field - column name from the data source object }
{ Data Source - "#ID" of the data source object }
{-Lang - ISO standard language abbreviation }
{-Name - variable name }
{-Type - MIME type of Data attribute. }
{-Events - array or list of events (OLEOBJHTMLEVENT structures) }
{-Params - array or list of params (OLEOBJHTMLPARAM structures) }
{= }
  end {OLEOBJHTMLDATA};

{$ENDIF}

(******************************************************************************)
{ nsfole.h }
const
  OBJINFO_HTMLFLAGS_DECLARE = $00000001; {/* Declare - download and install object's code}

const
  OLE_ROOTISTORAGE_SUPPORTED = 1;
const
  OLE_ISTORAGE_SUPPORTED = 2;

{+// }
{-API Utility functions used to extract, create and delete OLE2 object storage }
{-blobs from Notes documents. These functions only deal with the }
{-storage-related objects that comprise OLE object storage in Notes }
{-documents-and do NOT address any Rich-text references or representations of }
{-the OLE objects. It is assumed that the Rich Text (CDOLEBEGIN/CDOLEEND) }
{-objects are dealt with externally and separately from these functions. }
{= }

{+// NSFNoteExtractOLE2Object }

{-Create a copy of an OLE2 object structured storage file, serialized to }
{-an On-disk file. }

{-Note: It is assumed that the caller of this function has access to the OLE }
{-structured storage type (in the $OLEOBJINFO item) for this OLE object }
{-and handles setting or passing this information on as appropriate. }

{-Input: }

{-hNote }
{-Note handle of open Note }

{-pszObjectName }
{-Name of the OLE $FILE object which is the "master" extendable file object }
{-("ext***") in the set of $FILE objects that comprise this OLE object. }

{-char*pszFileName }
{-The file name, including path, into which the Storage file will be dumped. }

{-pEncryptionKey }
{-The Note's bulk data encryption key, acquired from NSFNoteDecrypt(...) }

{-fOverwrite }
{-TRUE to overwrite the file, if it already exists }

{-dwFlags }
{-Additional flags, unused at this time, must be 0 }

{= }


function NSFNoteExtractOLE2Object(hNote: NOTEHANDLE;
                                  pszObjectName: PChar; 
                                  pszFileName: PChar; 
                                  pEncryptionKey: PENCRYPTION_KEY;
                                  fOverwrite: Bool; 
                                  dwFlags: LongInt): STATUS cdecl  {$IFDEF WIN32} stdcall {$ENDIF};

{+// NSFNoteDeleteOLE2Object }

{-Delete an OLE2 structured storage object, which in Notes is stored as }
{-an extendable file object. Also, optionally delete the associated }
{-$OLEOBJINFO item }

{-Input: }

{-hNote }
{-Note handle of open Note }

{-pszObjectName }
{-Name of the OLE $FILE object which is the "master" extendable file object }
{-("ext***") in the set of $FILE objects that comprise this OLE object. }

{-fDeleteObjInfo }
{-True if the associated $OLEOBJINFO with this object should also be deleted }

{-pEncryptionKey }
{-The Note's bulk data encryption key, acquired from NSFNoteDecrypt(...). }
{-Necessary to decrypt the $FILE extent table. }

{-dwFlags }
{-Additional flags, unused at this time, must be 0 }

{= }


function NSFNoteDeleteOLE2Object(hNote: NOTEHANDLE; 
                                 pszObjectName: PChar; 
                                 fDeleteObjInfo: Bool;
                                 var pEncryptionKey: ENCRYPTION_KEY; 
                                 dwFlags: LongInt): STATUS cdecl  {$IFDEF WIN32} stdcall {$ENDIF}; 


{+// }
{-NSFNoteAttachOLE2Object }

{-Attach an OLE structured storage object to the specified Note, creating the }
{-OLE $FILE objects using the specified storage file. Also, create an }
{-$OLEOBJINFO note item using the specified parameters. }

{-OLE structured storage objects can have several forms. The forms supported are: }
{-- Notes OLE structured storage object - has the form RootIStorage, IStorage, }
{-IStream. }
{-- IStorage, IStream - has the form IStorage, IStream. }

{-Input: }

{-hNote }
{-NSF Note Handle to open note }

{-pszFileName }
{-Input file name containing the OLE structured storage file }

{-pszObjectName }
{-The name of the NSF $FILE extendable file object to create. This MUST }
{-match the one created in the CDOLEBEGIN record in the body of the }
{-document. }

{-fCreateInfoItem }
{-Obsolete, no longer used. }

{-pszObjDescription }
{-Object user description, i.e. "My Worksheet" }

{-pszFieldName }
{-Field name in which the OLE object resides within the document. }

{-pOleObjClassID (see OLEODS.H, optional) }
{-The OLE object's GUID }

{-wDisplayFormat }
{-Visual rendering format, like bitmap, metafile, DDEFORMAT_* from editods.h }

{-fScripted }
{-True if object is scripted (for ActiveX). If this is true, it's up to the }
{-caller to attach the associated Lotus script source and object code to this }
{-note, using the following naming convention <xxxx>.lso for the Lotus Script }
{-compiled object code and <xxxx>.lss for the Lotus Script source, where }
{-"xxxx" is identical to the the pszObjectName used above. If Object name }
{-is "foo" then "foo.lss" has the lotus script source, and "foo.lso" has the }
{-Lotus script object code. }

{-fOleControl }
{-True if this object is registered an OLE ActiveX. Optional, set to FALSE }
{-if unknown. Notes will determine setting when object is activated. }

{-dwFlags }
{-OLE storage structure of incoming OLE object }
{-=OLE_ROOTISTORAGE_SUPPORTED, incoming OLE object's storage structure }
{-is RootIStorage, IStorage, IStream (same as Notes OLE storage) }
{-=OLE_ISTORAGE_SUPPORTED, incoming OLE object's storage structure }
{-is IStorage, IStream }

{= }

function NSFNoteAttachOLE2Object(hNote: NOTEHANDLE;
                                 pszFileName: PChar;
                                 pszObjectName: PChar;
                                 fCreateInfoItem: Bool;
                                 pszObjDescription: PChar;
                                 pszFieldName: PChar;
                                 var pOleObjClassID: OLE_GUID;
                                 wDisplayFormat: Word;
                                 fScripted: Bool;
                                 fOleControl: Bool;
                                 dw: LongInt
                                 ) : STATUS stdcall; far;

(******************************************************************************)
{ <unknown> }

{ USER ACTIVITY RECORDS }

type
  DBACTIVITY_ENTRY = packed record
    Time: TIMEDATE;         // Time of record
    Reads: WORD;            // of data notes read
    Writes: WORD;           // of data notes written
    UserNameOffset:DWORD;   // Offset of the user name from the beginning of this memory block
                            // User names follow -- '\0' terminated
  end;
  PDBACTIVITY_ENTRY = ^DBACTIVITY_ENTRY;

  DBACTIVITY = packed record
    First: TIMEDATE;        // Beginning of reporting period */
    Last: TIMEDATE;         // End of reporting period */
    UsesPeriod: DWORD;      // # of uses in reporting period */
    Reads: DWORD;           // # of reads in reporting period */
    Writes: DWORD;          // # of writes in reporting period */
    PrevDayUses: DWORD;     // # of uses in previous 24 hours */
    PrevDayReads: DWORD;    // # of reads in previous 24 hours */
    PrevDayWrites: DWORD;   // # of writes in previous 24 hours */
    PrevWeekUses: DWORD;    // # of uses in previous week */
    PrevWeekReads: DWORD;   // # of reads in previous week */
    PrevWeekWrites: DWORD;  // # of writes in previous week */
    PrevMonthUses: DWORD;   // # of uses in previous month */
    PrevMonthReads: DWORD;  // # of reads in previous month */
    PrevMonthWrites: DWORD; // # of writes in previous month */
  end;
  PDBACTIVITY = ^DBACTIVITY;

  DBUserActivity = packed record
    Date: TDateTime;
    Reads: Word;
    Writes: Word;
    UserName: String;
  end;
  DBUserActivityArray = array of DBUserActivity;

function NSFDbGetUserActivity(
  hDB: DBHANDLE;
  Flags: DWord;
  retDbActivity: PDBACTIVITY;
  var rethUserInfo: HANDLE;
  var retUserCount: WORD): STATUS; stdcall; far;

function NSFGetMaxPasswordAccess(hDB: DBHANDLE; retLevel: Pword): STATUS; far;
function NSFSetMaxPasswordAccess(hDB: DBHANDLE; Level: word): STATUS; far; 

(******************************************************************************)
(******************************************************************************)
implementation

function TimeGMToLocal; external NOTES_DLL_NAME;

function TimeGMToLocalZone; external NOTES_DLL_NAME;

function TimeLocalToGM (var aTime: TimeStruct): bool; external NOTES_DLL_NAME;

function NotesInitIni(pConfigFileName: PChar): STATUS; external NOTES_DLL_NAME;

function NotesInit: STATUS; external NOTES_DLL_NAME;

function NotesInitExtended(argc: Integer;
                           argv: PPChar): STATUS; external NOTES_DLL_NAME;

procedure NotesTerm; external NOTES_DLL_NAME;

procedure NotesInitModule(rethModule: PHMODULE;
                          rethInstance: PHMODULE;
                          rethPrevInstance: PHMODULE); external NOTES_DLL_NAME;


function NotesInitThread: STATUS; external NOTES_DLL_NAME;

procedure NotesTermThread; external NOTES_DLL_NAME;


(******************************************************************************)
{NSFData.h}
(******************************************************************************)

function REPL_GET_PRIORITY (Flags: word): word;
begin
  result := ((flags shr ReplFlg_PRIORITY_SHIFT)+1) and REPLFLG_PRIORITY_MASK;
end;

(******************************************************************************)
function REPL_SET_PRIORITY (Pri: word): word;
begin
  result := (((Pri - 1) and REPLFLG_PRIORITY_MASK) shl REPLFLG_PRIORITY_SHIFT);
end;

(******************************************************************************)
function NSFTranslateSpecial(InputString: Pointer;
                             InputStringLength: Word;
                             OutputString: Pointer;
                             OutputStringBufferLength: Word;
                             NoteID: NOTEID;
                             IndexPosition: Pointer;
                             IndexInfo: PINDEXSPECIALINFO;
                             hUnreadList: THandle;
                             hCollapsedList: THandle;
                             FileTitle: Pchar;
                             ViewTitle: Pchar;
                             var RetLength: word): STATUS; external NOTES_DLL_NAME;

(******************************************************************************)
{NSFNote.h}
(******************************************************************************)
function NSFItemAppend(hNote: NOTEHANDLE;
                       ItemFlags: Word;
                       Name: PChar;
                       NameLength: Word;
                       DataType: Word;
                       Value: Pointer;
                       ValueLength: LongInt): STATUS; external NOTES_DLL_NAME;

function NSFItemAppendByBLOCKID(hNote: NOTEHANDLE;
                                ItemFlags: Word;
                                Name: PChar;
                                NameLength: Word;
                                bhValue: BLOCKID;
                                ValueLength: DWord;
                                retbhItem: PBLOCKID): STATUS; external NOTES_DLL_NAME;

function NSFItemAppendObject(hNote: NOTEHANDLE;
                             ItemFlags: Word;
                             Name: PChar;
                             NameLength: Word;
                             bhValue: BLOCKID;
                             ValueLength: DWord;
                             fDealloc: Bool): STATUS; external NOTES_DLL_NAME;
function NSFItemDelete(hNote: NOTEHANDLE;
                       Name: PChar;
                       NameLength: Word): STATUS; external NOTES_DLL_NAME;
function NSFItemDeleteByBLOCKID(hNote: NOTEHANDLE;
                                bhItem: BLOCKID): STATUS; external NOTES_DLL_NAME;
function NSFItemRealloc(bhItem: BLOCKID;
                        bhValue: PBLOCKID;
                        ValueLength: LongInt): STATUS; external NOTES_DLL_NAME;
function NSFItemCopy(hNote: NOTEHANDLE;
                     bhItem: BLOCKID): STATUS; external NOTES_DLL_NAME;
function NSFItemInfo(hNote: NOTEHANDLE;
                     Name: PChar;
                     NameLength: Word;
                     retbhItem: PBLOCKID;
                     retDataType: PWord;
                     retbhValue: PBLOCKID;
                     retValueLength: PLongInt): STATUS; external NOTES_DLL_NAME;

function NSFItemIsPresent (hNote: NoteHandle; Name: pchar; NameLength: word): boolean;
begin
  result := NSFItemInfo(hNote,NAME,NAMELENGTH,nil,nil,nil,nil) = NOERROR
end;

function NSFItemInfoNext(hNote: NOTEHANDLE;
                         PrevItem: BLOCKID;
                         Name: PChar;
                         NameLength: Word;
                         retbhItem: PBLOCKID;
                         retDataType: PWord;
                         retbhValue: PBLOCKID;
                         retValueLength: PDWORD): STATUS; external NOTES_DLL_NAME;

procedure NSFItemQuery(hNote: NOTEHANDLE;
                       bhItem: BLOCKID;
                       retItemName: PChar;
                       ItemNameBufferLength: Word;
                       retItemNameLength: PWord;
                       retItemFlags: PWord;
                       retDataType: PWord;
                       retbhValue: PBLOCKID;
                       retValueLength: PLongInt); external NOTES_DLL_NAME;
function NSFItemGetText(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        retBuffer: PChar;
                        BufferLength: Word): Word; external NOTES_DLL_NAME;

function NSFItemGetTime(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        retTime: PTIMEDATE): Bool; external NOTES_DLL_NAME;

function NSFItemGetModifiedTime (hNote: NOTEHANDLE;
                                 ItemName: PChar;
                                 ItemNameLength: word;
                                 Flags: dword;
                                 retTime: PTIMEDATE): Status; external NOTES_DLL_NAME;

function NSFItemGetNumber(hNote: NOTEHANDLE;
                          ItemName: PChar;
                          retNumber: PNUMBER): Bool; external NOTES_DLL_NAME;
function NSFItemGetLong(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        DefaultNumber: LongInt): LongInt; external NOTES_DLL_NAME;
function NSFItemSetText(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        Text: PChar;
                        TextLength: Word): STATUS; external NOTES_DLL_NAME;
function NSFItemSetTextSummary(hNote: NOTEHANDLE;
                               ItemName: PChar;
                               Text: PChar;
                               TextLength: Word;
                               Summary: Bool): STATUS; external NOTES_DLL_NAME;
function NSFItemSetTime(hNote: NOTEHANDLE;
                        ItemName: PChar;
                        Time: PTIMEDATE): STATUS; external NOTES_DLL_NAME;
function NSFItemSetNumber(hNote: NOTEHANDLE;
                          ItemName: PChar;
                          Number: PNUMBER): STATUS; external NOTES_DLL_NAME;
function NSFItemGetTextListEntries(hNote: NOTEHANDLE;
                                   ItemName: PChar): Word; external NOTES_DLL_NAME;
function NSFItemGetTextListEntry(hNote: NOTEHANDLE;
                                 ItemName: PChar;
                                 EntryPos: Word;
                                 retBuffer: PChar;
                                 BufferLength: Word): Word; external NOTES_DLL_NAME;

function NSFItemCreateTextList(hNote: NOTEHANDLE;
                               ItemName: PChar;
                               Text: PChar;
                               TextLength: Word): STATUS; external NOTES_DLL_NAME;
function NSFItemAppendTextList(hNote: NOTEHANDLE;
                               ItemName: PChar;
                               Text: PChar;
                               TextLength: Word;
                               fAllowDuplicates: Bool): STATUS; external NOTES_DLL_NAME;
function NSFItemTextEqual(hNote: NOTEHANDLE;
                          ItemName: PChar;
                          Text: PChar;
                          TextLength: Word;
                          fCaseSensitive: Bool): Bool; external NOTES_DLL_NAME;
function NSFItemTimeCompare(hNote: NOTEHANDLE;
                            ItemName: PChar;
                            Time: PTIMEDATE;
                            retVal: PInteger): Bool; external NOTES_DLL_NAME;
function NSFItemLongCompare(hNote: NOTEHANDLE;
                            ItemName: PChar;
                            Value: LongInt;
                            retVal: PInteger): Bool; external NOTES_DLL_NAME;
function NSFItemConvertValueToText(DataType: Word;
                                   bhValue: BLOCKID;
                                   ValueLength: LongInt;
                                   retBuffer: PChar;
                                   BufferLength: Word;
                                   SepChar: Char): Word; external NOTES_DLL_NAME;
function NSFItemConvertToText(hNote: NOTEHANDLE;
                              ItemName: PChar;
                              retBuffer: PChar;
                              BufferLength: Word;
                              SepChar: Char): Word; external NOTES_DLL_NAME;

function NSFGetSummaryValue(SummaryBuffer: Pointer;
                            Name: PChar;
                            retValue: PChar;
                            ValueBufferLength: Word): Bool; external NOTES_DLL_NAME;

function NSFLocateSummaryValue(SummaryBuffer: Pointer;
                               Name: PChar;
                               retValuePointer: Pointer;
                               retValueLength: PWord;
                               retDataType: PWord): Bool; external NOTES_DLL_NAME;
function NNOTESLinkFromText(hLinkText: THandle;
                             LinkTextLength: Word;
                             NoteLink: PNOTELINK;
                             ServerHint: PChar;
                             LinkText: PChar;
                             MaxLinkText: Word;
                             retFlags: PLongInt): STATUS; external NOTES_DLL_NAME;
function NNOTESLinkToText(Title: PChar;
                           NoteLink: PNOTELINK;
                           ServerHint: PChar;
                           LinkText: PChar;
                           phLinkText: PHandle;
                           pLinkTextLength: PWord;
                           Flags: LongInt): STATUS; external NOTES_DLL_NAME;
function NSFItemScan(hNote: NOTEHANDLE;
                     ActionRoutine: NSFITEMSCANPROC;
                     RoutineParameter: Pointer): STATUS; external NOTES_DLL_NAME;


procedure NNOTESGetInfo(hNote: NOTEHANDLE;
                         wType: Word;
                         Value: Pointer); external NOTES_DLL_NAME;

procedure NNOTESSetInfo(hNote: NOTEHANDLE;
                         wType: Word;
                         Value: Pointer); external NOTES_DLL_NAME;

function NNOTESClose(hNote: NOTEHANDLE): STATUS; external NOTES_DLL_NAME;

function NNOTESCreate(hDB: DBHANDLE;
                       rethNote: PNOTEHANDLE): STATUS; external NOTES_DLL_NAME;

function NNOTESDelete(hDB: DBHANDLE;
                       NoteID: NOTEID;
                       UpdateFlags: Word): STATUS; external NOTES_DLL_NAME;

function NNOTESOpen(hDB: DBHANDLE;
                     NoteID: NOTEID;
                     OpenFlags: Word;
                     rethNote: PNOTEHANDLE): STATUS; external NOTES_DLL_NAME;

function NNOTESOpenByUNID(hDB: THandle;
                           pUNID: PUNID;
                           flags: Word;
                           rtn: PHandle): STATUS; external NOTES_DLL_NAME;

function NNOTESUpdate(hNote: NOTEHANDLE;
                       UpdateFlags: Word): STATUS; external NOTES_DLL_NAME;

function NNOTESUpdateExtended(hNote: NOTEHANDLE;
                               UpdateFlags: LongInt): STATUS; external NOTES_DLL_NAME;

function NNOTESComputeWithForm(hNote: NOTEHANDLE;
                                hFormNote: NOTEHANDLE;
                                dwFlags: LongInt;
                                ErrorRoutine: CWF_ERROR_PROC;
                                CallersContext: Pointer): STATUS; external NOTES_DLL_NAME;
function NNOTESAttachFile(hNOTE: NOTEHANDLE;
                           ItemName: PChar;
                           ItemNameLength: Word;
                           PathName: PChar;
                           OriginalPathName: PChar;
                           Encoding: Word): STATUS; external NOTES_DLL_NAME;
function NNOTESExtractFile(hNote: NOTEHANDLE;
                            bhItem: BLOCKID;
                            FileName: PChar;
                            DecryptionKey: PENCRYPTION_KEY): STATUS; external NOTES_DLL_NAME;
function NNOTESExtractFileExt(hNote: NOTEHANDLE;
                               bhItem: BLOCKID;
                               FileName: PChar;
                               DecryptionKey: PENCRYPTION_KEY;
                               wFlags: Word): STATUS; external NOTES_DLL_NAME;
function NNOTESDetachFile(hNote: NOTEHANDLE;
                           bhItem: BLOCKID): STATUS; external NOTES_DLL_NAME;
function NNOTESHasObjects(hNote: NOTEHANDLE;
                           bhFirstObjectItem: PBLOCKID): Bool; external NOTES_DLL_NAME;
function NNOTESGetAuthor(hNote: NOTEHANDLE;
                          retName: PChar;
                          retNameLength: PWord;
                          retIsItMe: PBool): STATUS; external NOTES_DLL_NAME;
function NNOTESCopy(hSrcNote: NOTEHANDLE;
                     rethDstNote: PNOTEHANDLE): STATUS; external NOTES_DLL_NAME;
function NNOTESSignExt(hNote: NOTEHANDLE;
                        SignatureItemName: PChar;
                        ItemCount: Word;
                        hItemIDs: THandle): STATUS; external NOTES_DLL_NAME;
function NNOTESSign(hNote: NOTEHANDLE): STATUS; external NOTES_DLL_NAME;
function NNOTESVerifySignature(hNote: NOTEHANDLE;
                                Reserved: PChar;
                                retWhenSigned: PTIMEDATE;
                                retSigner: PChar;
                                retCertifier: PChar): STATUS; external NOTES_DLL_NAME;
function NSFVerifyFileObjSignature(hDB: DBHANDLE;
                                   bhItem: BLOCKID): STATUS; external NOTES_DLL_NAME;
function NNOTESUnsign(hNote: NOTEHANDLE): STATUS; external NOTES_DLL_NAME;
function NNOTESCopyAndEncrypt(hSrcNote: NOTEHANDLE;
                               EncryptFlags: Word;
                               rethDstNote: PNOTEHANDLE): STATUS; external NOTES_DLL_NAME;
function NNOTESDecrypt(hNote: NOTEHANDLE;
                        DecryptFlags: Word;
                        retKeyForAttachments: PENCRYPTION_KEY): STATUS; external NOTES_DLL_NAME;
function NNOTESIsSignedOrSealed(hNote: NOTEHANDLE;
                                 retfSigned: PBool;
                                 retfSealed: PBool): Bool; external NOTES_DLL_NAME;
function NNOTESCheck (hNote: THandle): STATUS; external NOTES_DLL_NAME;
function NSFNoteLinkFromText(hLinkText: THandle;
                             LinkTextLength: Word;
                             NoteLink: PNOTELINK;
                             ServerHint: PChar;
                             LinkText: PChar;
                             MaxLinkText: Word;
                             retFlags: PLongInt): STATUS; external NOTES_DLL_NAME;
function NSFNoteLinkToText(Title: PChar;
                           NoteLink: PNOTELINK;
                           ServerHint: PChar;
                           LinkText: PChar;
                           phLinkText: PHandle;
                           pLinkTextLength: PWord;
                           Flags: LongInt): STATUS; external NOTES_DLL_NAME;
procedure NSFNoteGetInfo(hNote: NOTEHANDLE;
                         wType: Word;
                         Value: Pointer); external  NOTES_DLL_NAME;
procedure NSFNoteSetInfo(hNote: NOTEHANDLE;
                         wType: Word;
                         Value: Pointer); external  NOTES_DLL_NAME;
function NSFNoteClose(hNote: NOTEHANDLE): STATUS; external  NOTES_DLL_NAME;
function NSFNoteCreate(hDB: DBHANDLE;
                       rethNote: PNOTEHANDLE): STATUS; external  NOTES_DLL_NAME;
function NSFNoteDelete(hDB: DBHANDLE;
                       NoteID: NOTEID;
                       UpdateFlags: Word): STATUS; external  NOTES_DLL_NAME;
function NSFNoteOpen(hDB: DBHANDLE;
                     NoteID: NOTEID;
                     OpenFlags: Word;
                     rethNote: PNOTEHANDLE): STATUS; external  NOTES_DLL_NAME;
function NSFNoteOpenByUNID(hDB: THandle;
                           pUNID: PUNID;
                           flags: Word;
                           rtn: PHandle): STATUS; external  NOTES_DLL_NAME;
function NSFNoteUpdate(hNote: NOTEHANDLE;
                       UpdateFlags: Word): STATUS; external  NOTES_DLL_NAME;
function NSFNoteUpdateExtended(hNote: NOTEHANDLE;
                               UpdateFlags: LongInt): STATUS; external  NOTES_DLL_NAME;
function NSFNoteComputeWithForm(hNote: NOTEHANDLE;
                                hFormNote: NOTEHANDLE;
                                dwFlags: LongInt;
                                ErrorRoutine: CWF_ERROR_PROC;
                                CallersContext: Pointer): STATUS; external  NOTES_DLL_NAME;
function NSFNoteAttachFile; external  NOTES_DLL_NAME;
function NSFNoteExtractFile(hNote: NOTEHANDLE;
                            bhItem: BLOCKID;
                            FileName: PChar;
                            DecryptionKey: PENCRYPTION_KEY): STATUS; external  NOTES_DLL_NAME;
function NSFNoteExtractFileExt(hNote: NOTEHANDLE;
                               bhItem: BLOCKID;
                               FileName: PChar;
                               DecryptionKey: PENCRYPTION_KEY;
                               wFlags: Word): STATUS; external  NOTES_DLL_NAME;
function NSFNoteDetachFile(hNote: NOTEHANDLE;
                           bhItem: BLOCKID): STATUS; external  NOTES_DLL_NAME;
function NSFNoteHasObjects(hNote: NOTEHANDLE;
                           bhFirstObjectItem: PBLOCKID): Bool; external  NOTES_DLL_NAME;
function NSFNoteGetAuthor(hNote: NOTEHANDLE;
                          retName: PChar;
                          retNameLength: PWord;
                          retIsItMe: PBool): STATUS; external  NOTES_DLL_NAME;
function NSFNoteCopy(hSrcNote: NOTEHANDLE;
                     rethDstNote: PNOTEHANDLE): STATUS; external  NOTES_DLL_NAME;
function NSFNoteSignExt(hNote: NOTEHANDLE;
                        SignatureItemName: PChar;
                        ItemCount: Word;
                        hItemIDs: THandle): STATUS; external  NOTES_DLL_NAME;
function NSFNoteSign(hNote: NOTEHANDLE): STATUS; external  NOTES_DLL_NAME;
function NSFNoteVerifySignature(hNote: NOTEHANDLE;
                                Reserved: PChar;
                                retWhenSigned: PTIMEDATE;
                                retSigner: PChar;
                                retCertifier: PChar): STATUS; external  NOTES_DLL_NAME;
function NSFNoteUnsign(hNote: NOTEHANDLE): STATUS; external  NOTES_DLL_NAME;
function NSFNoteCopyAndEncrypt(hSrcNote: NOTEHANDLE;
                               EncryptFlags: Word;
                               rethDstNote: PNOTEHANDLE): STATUS; external  NOTES_DLL_NAME;
function NSFNoteDecrypt(hNote: NOTEHANDLE;
                        DecryptFlags: Word;
                        retKeyForAttachments: PENCRYPTION_KEY): STATUS; external  NOTES_DLL_NAME;
function NSFNoteIsSignedOrSealed(hNote: NOTEHANDLE;
                                 retfSigned: PBool;
                                 retfSealed: PBool): Bool; external  NOTES_DLL_NAME;
function NSFNoteCheck (hNote: THandle): STATUS; external  NOTES_DLL_NAME;

function NSFProfileOpen; external NOTES_DLL_NAME;
function NSFProfileEnum; external NOTES_DLL_NAME;
function NSFProfileGetField; external NOTES_DLL_NAME;
function NSFProfileUpdate; external NOTES_DLL_NAME;
function NSFProfileSetField; external NOTES_DLL_NAME;
function NSFProfileDelete; external NOTES_DLL_NAME;


(******************************************************************************)
{COMPOUND TEXT FUNCTIONS}
(******************************************************************************)
function CompoundTextCreate(hNote: NOTEHANDLE;
                            pszItemName: PChar;
                            phCompound: PHandle): STATUS; external NOTES_DLL_NAME;
function CompoundTextClose(hCompound: THandle;
                           phReturnBuffer: PHandle;
                           pdwReturnBufferSize: PLongInt;
                           pchReturnFile: PChar;
                           wReturnFileSize: Word): STATUS; external NOTES_DLL_NAME;

procedure CompoundTextDiscard(hCompound: THandle); external NOTES_DLL_NAME;

function CompoundTextDefineStyle(hCompound: THandle;
                                 pszStyleName: PChar;
                                 pDefinition: PCOMPOUNDSTYLE;
                                 pdwStyleId: PLongInt): STATUS; external NOTES_DLL_NAME;
function CompoundTextAssimilateItem(hCompound: THandle;
                                    hNote: NOTEHANDLE;
                                    pszItemName: PChar;
                                    dwFlags: LongInt): STATUS; external NOTES_DLL_NAME;

function CompoundTextAssimilateFile(hCompound: THandle;
                                    pszFileSpec: PChar;
                                    dwFlags: LongInt): STATUS; external NOTES_DLL_NAME;

function CompoundTextAddParagraph(hCompound: THandle;
                                  dwStyleId: LongInt;
                                  FontID: FONTID;
                                  pchText: PChar;
                                  dwTextLen: LongInt;
                                  hCLSTable: THandle): STATUS; external NOTES_DLL_NAME;

function CompoundTextAddText(hCompound: THandle;
                             dwStyleId: LongInt;
                             FontID: FONTID;
                             pchText: PChar;
                             dwTextLen: LongInt;
                             pszLineDelim: PChar;
                             dwFlags: LongInt;
                             hCLSTable: THandle): STATUS; external NOTES_DLL_NAME;

procedure CompoundTextInitStyle(pStyle: PCOMPOUNDSTYLE); external NOTES_DLL_NAME;

function CompoundTextAddDocLink(hCompound: THandle;
                                DBReplicaID: TIMEDATE;
                                ViewUNID: UNID;
                                NoteUNID: UNID;
                                pszComment: PChar;
                                dwFlags: LongInt): STATUS; external NOTES_DLL_NAME;

function CompoundTextAddRenderedNote(hCompound: THandle;
                                     hNote: NOTEHANDLE;
                                     hFormNote: NOTEHANDLE;
                                     dwFlags: LongInt): STATUS; external NOTES_DLL_NAME;

(****************************************************************************)
{OsMisc.H}
(******************************************************************************)
function OSLoadString(hModule: HMODULE;
                      StringCode: STATUS;
                      retBuffer: PChar;
                      BufferLength: Word): Word; external NOTES_DLL_NAME;

function OSTranslate(TranslateMode: Word;
                     sIn: PChar;
                     InLength: Word;
                     Out: PChar;
                     OutLength: Word): Word; external NOTES_DLL_NAME;

{ Dynamic link library portable load routines }


function OSLoadLibrary; external NOTES_DLL_NAME

procedure OSFreeLibrary(_1: HMODULE); external NOTES_DLL_NAME;

{ Routine used in non-premptive platforms to simulate it. }


procedure OSPreemptOccasionally; external NOTES_DLL_NAME;

function OSGetLMBCSCLS: NLS_PINFO; external NOTES_DLL_NAME;

function OSGetNativeCLS: NLS_PINFO; external NOTES_DLL_NAME;


(******************************************************************************)
{reg.h}
(******************************************************************************)
function REGGetIDInfo(
  IDFileName: pchar;
  InfoType: WORD;
  OutBufr: pointer;
  OutBufrLen: WORD;
  ActualLen: PWORD): STATUS; external 'NNOTES.DLL';
(******************************************************************************)
{kfm.h}
(******************************************************************************)

function SECKFMSwitchToIDFile( pIDFileName:PChar;
                               pPassword:PChar;
                               pUserName:PChar;
                               MaxUserNameLength:WORD;
                               Flags:DWORD;
                               pReserved:PChar): STATUS; external NOTES_DLL_NAME;

function SECKFMUserInfo(aFunction: Word;
                        lpName: PChar;
                        var lpLicense: LICENSEID): STATUS; external NOTES_DLL_NAME;

function SECKFMGetUserName(retUserName: PChar): STATUS; external NOTES_DLL_NAME;

function SECKFMGetCertifierCtx(pCertFile: PChar;
                               pKfmPW: PKFM_PASSWORD;
                               pLogFile: PChar;
                               pExpDate: PTIMEDATE;
                               retCertName: PChar;
                               rethKfmCertCtx: PHCERTIFIER;
                               retfIsHierarchical: PBool;
                               retwFileVersion: PWord): STATUS; external NOTES_DLL_NAME;

function SECKFMSetCertifierExpiration(hKfmCertCtx: HCERTIFIER;
                                      pExpirationDate: PTIMEDATE): STATUS; external NOTES_DLL_NAME;

function SECKFMGetPublicKey(pName: PChar;
                            aFunction: Word;
                            Flags: Word;
                            rethPubKey: PHandle): STATUS; external NOTES_DLL_NAME;

(******************************************************************************)
{OsMem.H}
(******************************************************************************)

function OSMemAlloc (BlkType: Word;
                     dwSize: LongInt;
                     retHandle: PHandle): STATUS; stdcall; far; external NOTES_DLL_NAME;

function OSMemFree (Handle: THandle): STATUS; stdcall; far; external NOTES_DLL_NAME;

function OSMemGetSize (Handle: THandle;
                       retSize: PLongInt): STATUS; stdcall; far; external NOTES_DLL_NAME;

function OSMemRealloc (Handle: THandle;
                       NewSize: LongInt): STATUS; stdcall; far; external NOTES_DLL_NAME;

function OSLockObject (Handle: THandle): pointer; stdcall; far; external NOTES_DLL_NAME;

function OSUnlockObject(Handle: THandle): Bool; stdcall; far; external NOTES_DLL_NAME;

function OSLockBlock (BlckId: BLOCKID): pointer;
begin
  Result := pointer (longint (OSLockObject (BlckId.pool)) + BlckId.block);
end;

procedure OSUnlockBlock(BlckId: BLOCKID);
begin
  OSUnlockObject(blckid.pool);
end;

(******************************************************************************)
{OsEnv.H}
(******************************************************************************)
function OSGetEnvironmentString(VariableName: PChar;
                                retValueBuffer: PChar;
                                BufferLength: Word): Bool; external NOTES_DLL_NAME;
function OSGetEnvironmentLong(VariableName: PChar): LongInt; external NOTES_DLL_NAME;
procedure OSSetEnvironmentVariable(VariableName: PChar;
                                   Value: PChar); external NOTES_DLL_NAME;
procedure OSSetEnvironmentInt(VariableName: PChar;
                              Value: Integer); external NOTES_DLL_NAME;

(******************************************************************************)
{OsFile.H}
(******************************************************************************)

function OSPathNetConstruct(PortName: PChar;
                            ServerName: PChar;
                            FileName: PChar;
                            retPathName: PChar): STATUS; external NOTES_DLL_NAME;

function OSPathNetParse(PathName: PChar;
                        retPortName: PChar;
                        retServerName: PChar;
                        retFileName: PChar): STATUS; external NOTES_DLL_NAME;

function OSGetDataDirectory(retPathName: PChar): Word; external NOTES_DLL_NAME;

(******************************************************************************)
{NsfSearch.H}
(******************************************************************************)

function NSFFormulaCompile(FormulaName: PChar;
                           FormulaNameLength: Word;
                           FormulaText: PChar;
                           FormulaTextLength: Word;
                           rethFormula: PFORMULAHANDLE;
                           retFormulaLength: PWord;
                           retCompileError: PSTATUS;
                           retCompileErrorLine,
                           retCompileErrorColumn,
                           retCompileErrorOffSet,
                           retCompileErrorLength: pWord): Status; external NOTES_DLL_NAME;

function NSFFormulaDecompile(FormulaBuffer: PChar;
                           fSelectionFormula: Boolean;
                           rethFormulaText: PHandle;
                           retFormulaTextLength: PWord): STATUS; external NOTES_DLL_NAME;


function NSFFormulaMerge(hSrcFormula: FORMULAHANDLE;
                         hDestFormula: FORMULAHANDLE): STATUS; external NOTES_DLL_NAME;

function NSFFormulaSummaryItem(hFormula: FORMULAHANDLE;
                               ItemName: PChar;
                               ItemNameLength: Word): STATUS; external NOTES_DLL_NAME;

function NSFFormulaGetSize(hFormula: FORMULAHANDLE;
                           retFormulaLength: PWord): STATUS; external NOTES_DLL_NAME;

function NSFComputeEvaluate(hCompute: HCOMPUTE;
                            hNote: NOTEHANDLE;
                            rethResult: PHandle;
                            retResultLength: PWord;
                            retNoteMatchesFormula: PBool;
                            retNoteShouldBeDeleted: PBool;
                            retNoteModified: PBool): STATUS; external NOTES_DLL_NAME;


function NSFComputeStart; external NOTES_DLL_NAME;

function NSFComputeStop(hCompute: HCOMPUTE): STATUS; external NOTES_DLL_NAME;

function NSFSearch(hDB: DBHANDLE;
                   hFormula: FORMULAHANDLE;
                   ViewTitle: PChar;
                   SearchFlags: Word;
                   NoteClassMask: Word;
                   Since: PTIMEDATE;
                   EnumRoutine: NSFSEARCHPROC;
                   EnumRoutineParameter: Pointer;
                   retUntil: PTIMEDATE): STATUS; external NOTES_DLL_NAME;

(******************************************************************************)
{NIF.H}
(******************************************************************************)


function NIFOpenCollection(hViewDB: DBHANDLE;
                           hDataDB: DBHANDLE;
                           ViewNoteID: NOTEID;
                           OpenFlags: Word;
                           hUnreadList: THandle;
                           var rethCollection: HCOLLECTION;
                           rethViewNote: PNOTEHANDLE;
                           retViewUNID: PUNID;
                           rethCollapsedList: PHandle;
                           rethSelectedList:  PHANDLE): STATUS; external NOTES_DLL_NAME;

function NIFCloseCollection(hCollection: HCOLLECTION): STATUS; external NOTES_DLL_NAME;
function NIFUpdateCollection(hCollection: HCOLLECTION): STATUS; external NOTES_DLL_NAME;

function NIFReadEntries(hCollection: HCOLLECTION;
                        IndexPos: PCOLLECTIONPOSITION;
                        SkipNavigator: Word;
                        SkipCount: DWORD;
                        ReturnNavigator: Word;
                        ReturnCount: DWORD;
                        ReturnMask: DWORD;
                        rethBuffer: PHandle;
                        retBufferLength: PWord;
                        retNumEntriesSkiped: pdword;
                        var retNumEntriesReturned: dword;
                        var retSignalFlags: word): STATUS; external NOTES_DLL_NAME;

function NIFSetCollation(hCollection: HCOLLECTION;
                         CollationNum: Word): STATUS; external NOTES_DLL_NAME;

function NIFFindByKey(hCollection: HCOLLECTION;
                      KeyBuffer: Pointer;
                      FindFlags: Word;
                      retIndexPos: PCOLLECTIONPOSITION;
                      retNumMatches: PDWORD): STATUS; external NOTES_DLL_NAME;

function NIFFindByName(hCollection: HCOLLECTION;
                       Name: PChar;
                       FindFlags: Word;
                       retIndexPos: PCOLLECTIONPOSITION;
                       retNumMatches: PLongInt): STATUS; external NOTES_DLL_NAME;

function NIFFindDesignNote(hFile: DBHANDLE;
                           Name: PChar;
                           aClass: Word;
                           retNoteID: PNOTEID): STATUS; external NOTES_DLL_NAME;

function NIFFindPrivateDesignNote(hFile: DBHANDLE;
                                  Name: PChar;
                                  aClass: Word;
                                  retNoteID: PNOTEID): STATUS; external NOTES_DLL_NAME;

function NIFGetCollectionData(hCollection: HCOLLECTION;
                              rethCollData: PHandle): STATUS; external NOTES_DLL_NAME;

procedure NIFGetLastModifiedTime(hCollection: HCOLLECTION;
                                 retLastModifiedTime: PTIMEDATE); external NOTES_DLL_NAME;

function NIFFindView(hFile:DBHANDLE; Name: PChar;retNoteID: PNoteId): Status;
begin
 result := NIFFindDesignNote(hFile,Name,NOTE_CLASS_VIEW,retNoteID);
end;

function NIFFindDesignNoteByName (hFile: DBHandle; Name: PChar; retNoteID:PNoteId):Status;
begin
 result := NIFFindDesignNote(hFile,Name,NOTE_CLASS_ALL,retNoteID); { Only for V2 backward compatibility }
end;

function NIFFindPrivateView(hFile: DbHandle;Name: PChar; retNoteID: PNoteId): status;
begin
   result := NIFFindPrivateDesignNote(hFile,Name,NOTE_CLASS_VIEW,retNoteID);
end;

(******************************************************************************)
{OsTime.h}
(******************************************************************************)

procedure OSCurrentTIMEDATE(retTimeDate: PTIMEDATE); external NOTES_DLL_NAME;

procedure OSCurrentTimeZone(retZone: PInteger;
                            retDST: PInteger); external NOTES_DLL_NAME;

(******************************************************************************)
{TextList.h}
(******************************************************************************)

function ListAllocate(ListEntries: Word;
                      TextSize: Word;
                      fPrefixDataType: Bool;
                      rethList: PHandle;
                      retpList: Pointer;
                      retListSize: PWord): STATUS; external NOTES_DLL_NAME;

function ListAddText(pList: Pointer;
                     fPrefixDataType: Bool;
                     EntryNumber: Word;
                     Text: PChar;
                     TextSize: Word): STATUS; external NOTES_DLL_NAME;

function ListGetText(pList: Pointer;
                     fPrefixDataType: Bool;
                     EntryNumber: Word;
                     retTextPointer: PPChar;
                     retTextLength: PWord): STATUS; external NOTES_DLL_NAME;

function ListRemoveEntry(hList: THandle;
                         fPrefixDataType: Bool;
                         pListSize: PWord;
                         EntryNumber: Word): STATUS; external NOTES_DLL_NAME;

function ListRemoveAllEntries(hList: THandle;
                              fPrefixDataType: Bool;
                              pListSize: PWord): STATUS; external NOTES_DLL_NAME;

function ListAddEntry(hList: THandle;
                      fPrefixDataType: Bool;
                      pListSize: PWord;
                      EntryNumber: Word;
                      Text: PChar;
                      TextSize: Word): STATUS; external NOTES_DLL_NAME;

function ListGetSize(pList: Pointer;
                     fPrefixDataType: Bool): Word; external NOTES_DLL_NAME;

function ListDuplicate(var pInList: LIST;
                       fNoteItem: Bool;
                       phOutList: PHandle): STATUS; external NOTES_DLL_NAME;

function ListGetNumEntries(vList: Pointer;
                           NoteItem: Bool): Word; external NOTES_DLL_NAME;

(******************************************************************************)
{nsfdb.h}
(******************************************************************************)

function NSFDbGetOptions(hDB: DBHANDLE;
                         retDbOptions: PLongInt): STATUS; external NOTES_DLL_NAME;

function NSFDbSetOptions(hDB: DBHANDLE;
                         DbOptions: LongInt;
                         Mask: LongInt): STATUS; external NOTES_DLL_NAME;

function NSFDbCopyExtended(hSrcDB: DBHANDLE;
                           hDstDB: DBHANDLE;
                           Since: TIMEDATE;
                           NoteClassMask: Word;
                           Flags: LongInt;
                           retUntil: PTIMEDATE): STATUS; external NOTES_DLL_NAME;

function NSFDbOpen(PathName: PChar;
                   rethDB: PDBHANDLE): STATUS; external NOTES_DLL_NAME;

function NSFDbOpenExtended(PathName: PChar;
                           Options: Word;
                           hNames: THandle;
                           ModifiedTime: PTIMEDATE;
                           rethDB: PDBHANDLE;
                           retDataModified: PTIMEDATE;
                           retNonDataModified: PTIMEDATE): STATUS; external NOTES_DLL_NAME;

function NSFDbClose(hDB: DBHANDLE): STATUS; external NOTES_DLL_NAME;

function NSFDbCreate(PathName: PChar;
                     DbClass: USHORT;
                     ForceCreation: Bool): STATUS; external NOTES_DLL_NAME;

function NSFDbCreateObjectStore(PathName: PChar;
                                ForceCreation: Bool): STATUS; external NOTES_DLL_NAME;

function NSFDbDelete(PathName: PChar): STATUS; external NOTES_DLL_NAME;

function NSFDbCreateExtended(PathName: PChar;
                             DbClass: Word;
                             ForceCreation: Bool;
                             Options: Word;
                             EncryptStrength: BYTE;
                             MaxFileSize: LongInt): STATUS; external NOTES_DLL_NAME;

function NSFDbCopy(hSrcDB: DBHANDLE;
                   hDstDB: DBHANDLE;
                   Since: TIMEDATE;
                   NoteClassMask: Word): STATUS; external NOTES_DLL_NAME;

function NSFDbCopyNote(hSrcDB: DBHANDLE;
                       SrcDbID: PDBID;
                       SrcReplicaID: PDBID;
                       SrcNoteID: NOTEID;
                       hDstDB: DBHANDLE;
                       DstDbID: PDBID;
                       DstReplicaID: PDBID;
                       retDstNoteID: PNOTEID;
                       retNoteClass: PWord): STATUS; external NOTES_DLL_NAME;

function NSFDbCreateAndCopy(srcDb: PChar;
                            dstDb: PChar;
                            NoteClass: Word;
                            limit: Word;
                            flags: LongInt;
                            retHandle: PDBHANDLE): STATUS; external NOTES_DLL_NAME;

function NSFDbMarkForDelete(dbPathPtr: PChar): STATUS; external NOTES_DLL_NAME;

function NSFDbMarkInService(dbPathPtr: PChar): STATUS; external NOTES_DLL_NAME;

function NSFDbMarkOutOfService(dbPathPtr: PChar): STATUS; external NOTES_DLL_NAME;

function NSFDbCopyACL(hSrcDB: DBHANDLE;
                      hDstDB: DBHANDLE): STATUS; external NOTES_DLL_NAME;

function NSFDbCopyTemplateACL(hSrcDB: DBHANDLE;
                              hDstDB: DBHANDLE;
                              Manager: PChar;
                              DefaultAccessLevel: Word): STATUS; external NOTES_DLL_NAME;

function NSFDbCreateACLFromTemplate(hNTF: DBHANDLE;
                                    hNSF: DBHANDLE;
                                    Manager: PChar;
                                    DefaultAccess: Word;
                                    rethACL: PHandle): STATUS; external NOTES_DLL_NAME;

function NSFDbStoreACL(hDB: DBHANDLE;
                       hACL: THandle;
                       ObjectID: LongInt;
                       Method: Word): STATUS; external NOTES_DLL_NAME;

function NSFDbReadACL(hDB: DBHANDLE;
                      rethACL: PHandle): STATUS; external NOTES_DLL_NAME;

function NSFDbGenerateOID(hDB: DBHANDLE;
                          retOID: POID): STATUS; external NOTES_DLL_NAME;

function NSFDbModifiedTime(hDB: DBHANDLE;
                           retDataModified: PTIMEDATE;
                           retNonDataModified: PTIMEDATE): STATUS; external NOTES_DLL_NAME;

function NSFDbPathGet(hDB: DBHANDLE;
                      retCanonicalPathName: PChar;
                      retExpandedPathName: PChar): STATUS; external NOTES_DLL_NAME;

function NSFDbInfoGet(hDB: DBHANDLE;
                      retBuffer: PChar): STATUS; external NOTES_DLL_NAME;

function NSFDbInfoSet(hDB: DBHANDLE;
                      Buffer: PChar): STATUS; external NOTES_DLL_NAME;

procedure NSFDbInfoParse(Info: PChar;
                         What: Word;
                         Buffer: PChar;
                         Length: Word); external NOTES_DLL_NAME;

procedure NSFDbInfoModify(Info: PChar;
                          What: Word;
                          Buffer: PChar); external NOTES_DLL_NAME;

function NSFDbGetSpecialNoteID(hDB: DBHANDLE;
                               Index: Word;
                               retNoteID: PNOTEID): STATUS; external NOTES_DLL_NAME;

function NSFDbIDGet(hDB: DBHANDLE;
                    retDbID: PDBID): STATUS; external NOTES_DLL_NAME;

function NSFDbReplicaInfoGet(hDB: DBHANDLE;
                             retReplicationInfo: PDBREPLICAINFO): STATUS; external NOTES_DLL_NAME;

function NSFDbReplicaInfoSet(hDB: DBHANDLE;
                             ReplicationInfo: PDBREPLICAINFO): STATUS; external NOTES_DLL_NAME;

function NSFDbGetNoteInfo(hDB: DBHANDLE;
                          NoteID: NOTEID;
                          retNoteOID: POID;
                          retModified: PTIMEDATE;
                          retNoteClass: PWord): STATUS; external NOTES_DLL_NAME;

function NSFDbGetNoteInfoByUNID(hDB: THandle;
                                pUNID: PUNID;
                                retNoteID: PNOTEID;
                                retOID: POID;
                                retModTime: PTIMEDATE;
                                retClass: PWord): STATUS; external NOTES_DLL_NAME;

function NSFDbGetModifiedNoteTable(hDB: DBHANDLE;
                                   NoteClassMask: Word;
                                   Since: TIMEDATE;
                                   retUntil: PTIMEDATE;
                                   rethTable: PHandle): STATUS; external NOTES_DLL_NAME;

function NSFApplyModifiedNoteTable(hModifiedNotes: THandle;
                                   hTargetTable: THandle): STATUS; external NOTES_DLL_NAME;

function NSFDbLocateByReplicaID(hDB: DBHANDLE;
                                ReplicaID: PDBID;
                                retPathName: PChar;
                                PathMaxLen: Word): STATUS; external NOTES_DLL_NAME;

function NSFDbStampNotes(hDB: DBHANDLE;
                         hTable: THandle;
                         ItemName: PChar;
                         ItemNameLength: Word;
                         Data: Pointer;
                         DataLength: Word): STATUS; external NOTES_DLL_NAME;

function NSFDbDeleteNotes(hDB: DBHANDLE;
                          hTable: THandle;
                          retUNIDArray: PUNID): STATUS; external NOTES_DLL_NAME;

procedure NSFDbAccessGet(hDB: THandle;
                         retAccessLevel: PWord;
                         retAccessFlag: PWord); external NOTES_DLL_NAME;

function NSFDbClassGet(hDB: DBHANDLE;
                       retClass: PWord): STATUS; external NOTES_DLL_NAME;

function NSFDbModeGet(hDB: DBHANDLE;
                      retMode: PUSHORT): STATUS; external NOTES_DLL_NAME;

function NSFDbCloseSession(hDB: DBHANDLE): STATUS; external NOTES_DLL_NAME;

function NSFDbReopen(hDB: DBHANDLE;
                     rethDB: PDBHANDLE): STATUS; external NOTES_DLL_NAME;

function NSFDbMajorMinorVersionGet(hDB: DBHANDLE;
                                   retMajorVersion: PWord;
                                   retMinorVersion: PWord): STATUS; external NOTES_DLL_NAME;

function NSFDbItemDefTable(hDB: DBHANDLE;
                           retItemNameTable: PITEMDEFTABLEHANDLE): STATUS; external NOTES_DLL_NAME;

function NSFDbGetBuildVersion(hDB: DBHANDLE;
                              retVersion: PWord): STATUS; external NOTES_DLL_NAME;

function NSFDbSpaceUsage(hDB: DBHANDLE;
                         retAllocatedBytes: PLongInt;
                         retFreeBytes: PLongInt): STATUS; external NOTES_DLL_NAME;

function NSFDbGetOpenDatabaseID(hDB: DBHANDLE): LongInt; external NOTES_DLL_NAME;

function NSFGetServerStats(ServerName: PChar;
                           Facility: PChar;
                           StatName: PChar;
                           rethTable: PHandle;
                           retTableSize: PLongInt): STATUS; external NOTES_DLL_NAME;

function NSFGetServerLatency(ServerName: PChar;
                             Timeout: LongInt;
                             retClientToServerMS: PLongInt;
                             retServerToClientMS: PLongInt;
                             ServerVersion: PWord): STATUS; external NOTES_DLL_NAME;

function NSFRemoteConsole(ServerName: PChar;
                          ConsoleCommand: PChar;
                          hResponseText: PHandle): STATUS; external NOTES_DLL_NAME;

function NSFDbUpdateUnread(hDataDB: DBHANDLE;
                           hUnreadList: THandle): STATUS; external NOTES_DLL_NAME;

function NSFDbGetUnreadNoteTable(hDB: DBHANDLE;
                                 UserName: PChar;
                                 UserNameLength: Word;
                                 fCreateIfNotAvailable: Bool;
                                 rethUnreadList: PHandle): STATUS; external NOTES_DLL_NAME;

function NSFDbGetUnreadNoteTable2(hDB: DBHANDLE;
                                 UserName: PChar;
                                 UserNameLength: Word;
                                 fCreateIfNotAvailable: Bool;
                                 fUpdateUnread: Bool;                                 
                                 rethUnreadList: PHandle): STATUS; external NOTES_DLL_NAME;

function NSFDbSetUnreadNoteTable(hDB: DBHANDLE;
                                 UserName: PChar;
                                 UserNameLength: Word;
                                 fFlushToDisk: Bool;
                                 hOriginalUnreadList: THandle;
                                 hUnreadList: THandle): STATUS; external NOTES_DLL_NAME;

function NSFDbGetObjectStoreID(dbhandle: DBHANDLE;
                               Specified: PBool;
                               ObjStoreReplicaID: PDBID): STATUS; external NOTES_DLL_NAME;

function NSFDbSetObjectStoreID(dbhandle: DBHANDLE;
                               ObjStoreReplicaID: PDBID): STATUS; external NOTES_DLL_NAME;

function NSFDbFilter(hFilterDB: DBHANDLE;
                     hFilterNote: NOTEHANDLE;
                     hNotesToFilter: THandle;
                     fIncremental: Bool;
                     Reserved1: Pointer;
                     Reserved2: Pointer;
                     DbTitle: PChar;
                     ViewTitle: PChar;
                     Reserved3: Pointer;
                     Reserved4: Pointer;
                     hDeletedList,HSelectedList: THandle): STATUS; external NOTES_DLL_NAME;

function NSFDbCompact (PathName: PChar;
                       Options: word;
                       var RetStats: dword):Status; external NOTES_DLL_NAME;

function NSFDbQuotaGet(Filename: PChar;
                            retQuotaInfo: PDBQUOTAINFO): STATUS; external NOTES_DLL_NAME;


(******************************************************************************)
{event.h}
(******************************************************************************)

function EventQueueAlloc(QueueName: PChar): STATUS; external NOTES_DLL_NAME;

procedure EventQueueFree(QueueName: PChar); external NOTES_DLL_NAME;

function EventQueuePut(QueueName: PChar;
                       OriginatingServer: PChar;
                       aType: Word;
                       Severity: Word;
                       EventTime: PTIMEDATE;
                       FormatSpecifier: Word;
                       EventDataLength: Word;
                       EventSpecificData: Pointer): STATUS; external NOTES_DLL_NAME;

function EventQueueGet(QueueName: PChar;
                       rethEvent: PHandle): STATUS; external NOTES_DLL_NAME;

function EventRegisterEventRequest(EventType: Word;
                                   EventSeverity: Word;
                                   QueueName: PChar;
                                   DestName: PChar): STATUS; external NOTES_DLL_NAME;

function EventDeregisterEventRequest(EventType: Word;
                                     EventSeverity: Word;
                                     QueueName: PChar): STATUS; external NOTES_DLL_NAME;

function EventGetDestName(EventType: Word;
                          Severity: Word;
                          QueueName: PChar;
                          DestName: PChar;
                          DestNameSize: Word): Bool; external NOTES_DLL_NAME;
(******************************************************************************)
{mailserv.h}
(******************************************************************************)


function MailGetDomainName(Domain: PChar): STATUS; external NOTES_DLL_NAME;

function MailLookupAddress(UserName: PChar;
                           MailAddress: PChar): STATUS; external NOTES_DLL_NAME;

function MailLookupUser(UserName: PChar;
                        FullName: PChar;
                        MailServerName: PChar;
                        MailFileName: PChar;
                        MailAddress: PChar;
                        ShortName: PChar): STATUS; external NOTES_DLL_NAME;

function MailGetMessageItem(hMessage: THandle;
                            ItemCode: Word;
                            retString: PChar;
                            StringSize: Word;
                            retStringLength: PWord): STATUS; external NOTES_DLL_NAME;

function MailGetMessageItemHandle(hMessage: THandle;
                                  ItemCode: Word;
                                  retbhValue: PBLOCKID;
                                  retValueType: PWord;
                                  retValueLength: PLongInt): STATUS; external NOTES_DLL_NAME;

function MailGetMessageItemTimeDate(hMessage: THandle;
                                    ItemCode: Word;
                                    retTimeDate: PTIMEDATE): STATUS; external NOTES_DLL_NAME;

function MailCreateMessage(hFile: DBHANDLE;
                           rethMessage: PHandle): STATUS; external NOTES_DLL_NAME;

function MailAddHeaderItem(hMessage: THandle;
                           ItemCode: Word;
                           Value: PChar;
                           ValueLength: Word): STATUS; external NOTES_DLL_NAME;

function MailAddHeaderItemByHandle(hMessage: THandle;
                                   ItemCode: Word;
                                   hValue: THandle;
                                   ValueLength: Word;
                                   ItemFlags: Word): STATUS; external NOTES_DLL_NAME;

function MailReplaceHeaderItem(hMessage: THandle;
                               ItemCode: Word;
                               Value: Pointer;
                               ValueLength: Word): STATUS; external NOTES_DLL_NAME;

function MailCreateBodyItem(rethBodyItem: PHandle;
                            retBodyLength: PLongInt): STATUS; external NOTES_DLL_NAME;

function MailAppendBodyItemLine(hBodyItem: THandle;
                                BodyLength: PLongInt;
                                Text: PChar;
                                TextLength: Word): STATUS; external NOTES_DLL_NAME;

function MailAddBodyItem(hMessage: THandle;
                         hBodyItem: THandle;
                         BodyLength: LongInt;
                         CTFName: PChar): STATUS; external NOTES_DLL_NAME;

function MailAddRecipientsItem(hMessage: THandle;
                               hRecipientsItem: THandle;
                               RecipientsLength: Word): STATUS; external NOTES_DLL_NAME;

function MailTransferMessageLocal(hMessage: THandle): STATUS; external NOTES_DLL_NAME;

function MailIsNonDeliveryReport(hMessage: THandle): Bool; external NOTES_DLL_NAME;

function MailGetMessageType(hMessage: THandle): Word; external NOTES_DLL_NAME;

function MailCloseMessage(hMessage: THandle): STATUS; external NOTES_DLL_NAME;

function MailExpandNames(hWorkList: THandle;
                         WorkListSize: Word;
                         hOutputList: PHandle;
                         OutputListSize: PWord;
                         UseExpanded: Bool;
                         hRecipsExpanded: THandle): STATUS; external NOTES_DLL_NAME;

function MailLogEvent(Flags: Word;
                      StringID: STATUS;
                      hModule: HMODULE;
                      AdditionalErrorCode: STATUS;
                      _5: dword {Undefined number of parameters in C function was here}): STATUS; external NOTES_DLL_NAME;

function MailLogEventText(Flags: Word;
                          aString: PChar;
                          hModule: HMODULE;
                          AdditionalErrorCode: STATUS;
                          _5: dword {Undefined number of parameters in C function was here}): STATUS; external NOTES_DLL_NAME;

function MailGetMessageAttachmentInfo(hMessage: THandle;
                                      Num: Word;
                                      bhItem: PBLOCKID;
                                      FileName: PChar;
                                      FileSize: PLongInt;
                                      FileAttributes: PWord;
                                      FileHostType: PWord;
                                      FileCreated: PTIMEDATE;
                                      FileModified: PTIMEDATE): Bool; external NOTES_DLL_NAME;

function MailExtractMessageAttachment(hMessage: THandle;
                                      bhItem: BLOCKID;
                                      FileName: PChar): STATUS; external NOTES_DLL_NAME;

function MailAddMessageAttachment(hMessage: THandle;
                                  FileName: PChar;
                                  OriginalFileName: PChar): STATUS; external NOTES_DLL_NAME;

function MailOpenMessageFile(FileName: PChar;
                             rethFile: PDBHANDLE): STATUS; external NOTES_DLL_NAME;

function MailCreateMessageFile(FileName: PChar;
                               TemplateFileName: PChar;
                               Title: PChar;
                               rethFile: PDBHANDLE): STATUS; external NOTES_DLL_NAME;

function MailPurgeMessageFile(hFile: DBHANDLE): STATUS; external NOTES_DLL_NAME;

function MailCloseMessageFile(hFile: DBHANDLE): STATUS; external NOTES_DLL_NAME;

function MailGetMessageFileModifiedTime(hFile: DBHANDLE;
                                        retModifiedTime: PTIMEDATE): STATUS; external NOTES_DLL_NAME;

function MailCreateMessageList(hFile: DBHANDLE;
                               hMessageList: PHandle;
                               var MessageList: PDARRAY;
                               MessageCount: PWord): STATUS; external NOTES_DLL_NAME;

function MailFreeMessageList(hMessageList: THandle;
                             MessageList: PDARRAY): STATUS; external NOTES_DLL_NAME;

function MailGetMessageInfo(MessageList: PDARRAY;
                            aMessage: Word;
                            RecipientCount: PWord;
                            Priority: PWord;
                            Report: PWord): STATUS; external NOTES_DLL_NAME;

function MailGetMessageSize(MessageList: PDARRAY;
                            aMessage: Word;
                            MessageSize: PLongInt): STATUS; external NOTES_DLL_NAME;

function MailGetMessageRecipient(MessageList: PDARRAY;
                                 aMessage: Word;
                                 RecipientNum: Word;
                                 RecipientName: PChar;
                                 RecipientNameSize: Word;
                                 RecipientNameLength: PWord): STATUS; external NOTES_DLL_NAME;

function MailDeleteMessageRecipient(MessageList: PDARRAY;
                                    aMessage: Word;
                                    RecipientNum: Word): STATUS; external NOTES_DLL_NAME;

function MailGetMessageOriginator(MessageList: PDARRAY;
                                  aMessage: Word;
                                  OriginatorName: PChar;
                                  OriginatorNameSize: Word;
                                  OriginatorNameLength: PWord): STATUS; external NOTES_DLL_NAME;

function MailGetMessageOriginatorDomain(MessageList: PDARRAY;
                                        aMessage: Word;
                                        OriginatorDomain: PChar;
                                        OriginatorDomainSize: Word;
                                        OriginatorNameLength: PWord): STATUS; external NOTES_DLL_NAME;

function MailOpenMessage(MessageList: PDARRAY;
                         aMessage: Word;
                         hMessage: PHandle): STATUS; external NOTES_DLL_NAME;

function MailGetMessageBody(hMessage: THandle;
                            hBody: PHandle;
                            BodyLength: PLongInt): STATUS; external NOTES_DLL_NAME;

function MailGetMessageBodyText(hMessage: THandle;
                                ItemName: PChar;
                                LineDelims: PChar;
                                LineLength: Word;
                                ConvertTabs: Bool;
                                OutputFileName: PChar;
                                OutputFileSize: PLongInt): STATUS; external NOTES_DLL_NAME;

function MailGetMessageBodyComposite(hMessage: THandle;
                                     ItemName: PChar;
                                     OutputFileName: PChar;
                                     OutputFileSize: PLongInt): STATUS; external NOTES_DLL_NAME;

function MailAddMessageBodyText(hMessage: THandle;
                                ItemName: PChar;
                                InputFileName: PChar;
                                FontID: LongInt;
                                LineDelim: PChar;
                                ParaDelim: Word;
                                CTFName: PChar): STATUS; external NOTES_DLL_NAME;

function MailAddMessageBodyComposite(hMessage: THandle;
                                     ItemName: PChar;
                                     InputFileName: PChar): STATUS; external NOTES_DLL_NAME;

function MailSetMessageLastError(MessageList: PDARRAY;
                                 aMessage: Word;
                                 ErrorText: PChar): STATUS; external NOTES_DLL_NAME;

function MailPurgeMessage(MessageList: PDARRAY;
                          aMessage: Word): STATUS; external NOTES_DLL_NAME;

function MailSendNonDeliveryReport(MessageList: PDARRAY;
                                   aMessage: Word;
                                   RecipientNums: Word;
                                   RecipientNumList: PWord;
                                   ReasonText: PChar;
                                   ReasonTextLength: Word): STATUS; external NOTES_DLL_NAME;

function MailSendDeliveryReport(MessageList: PDARRAY;
                                aMessage: Word;
                                RecipientNums: Word;
                                RecipientNumList: PWord): STATUS; external NOTES_DLL_NAME;

function MailParseMailAddress(MailAddress: PChar;
                              MailAddressLength: Word;
                              UserName: PChar;
                              UserNameSize: Word;
                              UserNameLength: PWord;
                              DomainName: PChar;
                              DomainNameSize: Word;
                              DomainNameLength: PWord): STATUS; external NOTES_DLL_NAME;

procedure MailBroadcastNewMail(MessageText: PChar); external NOTES_DLL_NAME;

function MailLoadRoutingTables(hAddressBook: DBHANDLE;
                               LocalServerName: PChar;
                               LocalDomainDomain: PChar;
                               TaskName: PChar;
                               EnableTrace: Bool;
                               EnableDebug: Bool;
                               rethTables: PHandle): STATUS; external NOTES_DLL_NAME;

function MailReloadRoutingTables(hTables: THandle;
                                 EnableTrace: Bool;
                                 EnableDebug: Bool;
                                 retAddressBookModified: PBool): STATUS; external NOTES_DLL_NAME;

function MailUnloadRoutingTables(hTables: THandle): STATUS; external NOTES_DLL_NAME;

function MailFindNextHopToDomain(hTables: THandle;
                                 OriginatorsDomain: PChar;
                                 DestDomain: PChar;
                                 NextHopServer: PChar;
                                 NextHopMailbox: PChar;
                                 NextHopFlags: PLongInt;
                                 ErrorServer: PChar): STATUS; external NOTES_DLL_NAME;

function MailFindNextHopToServer(hTables: THandle;
                                 DestDomain: PChar;
                                 DestServer: PChar;
                                 NextHopServer: PChar;
                                 NextHopMailbox: PChar;
                                 NextHopFlags: PLongInt;
                                 ActualCost: PWord): STATUS; external NOTES_DLL_NAME;

function MailFindNextHopToRecipient(hTables: THandle;
                                    OriginatorsDomain: PChar;
                                    RecipientAddress: PChar;
                                    var Action: MAIL_ROUTING_ACTIONS;
                                    NextHopServer: PChar;
                                    NextHopMailbox: PChar;
                                    ForwardAddress: PChar;
                                    ErrorText: PChar;
                                    var NextHopFlags: dword): STATUS; external NOTES_DLL_NAME;

function MailFindNextHopViaRules(hTables: THandle;
                                 RecipientAddress: PChar;
                                 retDestServer: PChar;
                                 retDestDomain: PChar): STATUS; external NOTES_DLL_NAME;

function MailSetDynamicCost(hTables: THandle;
                            Server: PChar;
                            CostBias: SWORD): Bool; external NOTES_DLL_NAME;

function MailResetAllDynamicCosts(hTables: THandle): Bool; external NOTES_DLL_NAME;

(******************************************************************************)
{ft.h}
(******************************************************************************)

function FTIndex(hDB: THandle;
                 Options: Word;
                 StopFile: PChar;
                 retStats: PFT_INDEX_STATS): STATUS; external NOTES_DLL_NAME;

function FTDeleteIndex(hDB: THandle): STATUS; external NOTES_DLL_NAME;

function FTGetLastIndexTime(hDB: THandle;
                            retTime: PTIMEDATE): STATUS; external NOTES_DLL_NAME;

function FTOpenSearch(rethSearch: PHandle): STATUS; external NOTES_DLL_NAME;

function FTSearch(hDB: THandle;
                  phSearch: PHandle;
                  hColl: HCOLLECTION;
                  Query: PChar;
                  Options: LongInt;
                  Limit: Word;
                  hIDTable: THandle;
                  retNumDocs: PLongInt;
                  Reserved: PHandle;
                  rethResults: PHandle): STATUS; external NOTES_DLL_NAME;

function FTCloseSearch(hSearch: THandle): STATUS; external NOTES_DLL_NAME;

(******************************************************************************)
{idtable.h}
(******************************************************************************)

function IDCreateTable(Alignment: LongInt;
                       rethTable: PHandle): STATUS; external NOTES_DLL_NAME;

function IDDestroyTable(hTable: THandle): STATUS; external NOTES_DLL_NAME;

function IDInsert(hTable: THandle;
                  id: LongInt;
                  retfInserted: PBool): STATUS; external NOTES_DLL_NAME;

function IDDelete(hTable: THandle;
                  id: LongInt;
                  retfDeleted: PBool): STATUS; external NOTES_DLL_NAME;

function IDDeleteAll(hTable: THandle): STATUS; external NOTES_DLL_NAME;

function IDScan(hTable: THandle;
                fFirst: Bool;
                retID: PLongInt): Bool; external NOTES_DLL_NAME;

function IDEnumerate(hTable: THandle;
                     Routine: IDENUMERATEPROC;
                     Parameter: Pointer): STATUS; external NOTES_DLL_NAME;

function IDEntries(hTable: THandle): LongInt; external NOTES_DLL_NAME;

function IDIsPresent(hTable: THandle;
                     id: LongInt): Bool; external NOTES_DLL_NAME;

function IDTableSize(hTable: THandle): LongInt; external NOTES_DLL_NAME;

function IDTableCopy(hTable: THandle;
                     rethTable: PHandle): STATUS; external NOTES_DLL_NAME;

function IDTableSizeP(pIDTable: Pointer): LongInt; external NOTES_DLL_NAME;

function IDTableFlags(pIDTable: Pointer): Word; external NOTES_DLL_NAME;

function IDTableTime(pIDTable: Pointer): TIMEDATE; external NOTES_DLL_NAME;

procedure IDTableSetFlags(pIDTable: Pointer;
                          Flags: Word); external NOTES_DLL_NAME;

procedure IDTableSetTime(pIDTable: Pointer;
                         Time: TIMEDATE); external NOTES_DLL_NAME;


(******************************************************************************)
{from ODS.H}
(******************************************************************************)

function EnumCompositeBuffer (ItemValue: BLOCKID; ItemValueLength: DWORD; ActionRoutine: ActionRoutinePtr;
  vContext: pointer): STATUS; external NOTES_DLL_NAME;
procedure ODSReadMemory; external NOTES_DLL_NAME;
procedure ODSWriteMemory; external NOTES_DLL_NAME;
function ODSLength; external NOTES_DLL_NAME;

(******************************************************************************)
{from MISC.H}
(******************************************************************************)
procedure TimeConstant (TimeConstantType: WORD; var Value: TIMEDATE); external NOTES_DLL_NAME;

(******************************************************************************)
{from OSSIGNAL.H}
(******************************************************************************)
function OSSetSignalHandler (wType: WORD; Proc: OSSIGPROC): OSSIGPROC; external NOTES_DLL_NAME;
function OSGetSignalHandler (wType: WORD): OSSIGPROC;  external NOTES_DLL_NAME;

(******************************************************************************)
{ from FONTID.H}
(******************************************************************************)

function BYTEMASK (const LeftShift: dword): dword;
begin
  Result := $000000ff shl leftshift;
end;

procedure FontIDSetSize (var fontid: dword; size: integer);
begin
  fontid := (fontid and (not BYTEMASK(FONT_SIZE_SHIFT))) or (size shl FONT_SIZE_SHIFT);
end;

procedure FontIDSetFaceID (var fontid: dword; faceId: dword);
begin
  fontid := (fontid and (not BYTEMASK(FONT_FACE_SHIFT))) or (faceID shl FONT_FACE_SHIFT);
end;

procedure FontIDSetColor (var fontid: dword; colorId: dword);
begin
  fontid := (fontid and (not BYTEMASK(FONT_COLOR_SHIFT))) or (colorID shl FONT_COLOR_SHIFT);
end;

procedure FontIDSetStyle (var fontid: dword; styleId: dword);
begin
  fontid := (fontid and (not BYTEMASK(FONT_STYLE_SHIFT))) or (styleID shl FONT_STYLE_SHIFT);
end;

function FontIDIsUnderline(const fontid: dword): boolean;
begin
  Result := (fontid and (CF_ISUNDERLINE shl FONT_STYLE_SHIFT) <> 0);
end;

function FontIDIsBold(const fontid: dword): boolean;
begin
  Result := (fontid and (CF_ISBOLD shl FONT_STYLE_SHIFT) <> 0);
end;

function FontIDIsItalic(const fontid: dword): boolean;
begin
  Result := (fontid and (CF_ISITALIC shl FONT_STYLE_SHIFT) <> 0);
end;

function FontIDIsStrikeout(const fontid: dword): boolean;
begin
  Result := (fontid and (CF_ISSTRIKEOUT shl FONT_STYLE_SHIFT) <> 0);
end;

function FontIDIsSuperscript(const fontid: dword): boolean;
begin
  Result := (fontid and (CF_ISSUPER shl FONT_STYLE_SHIFT) <> 0);
end;

function FontIDIsSubscript(const fontid: dword): boolean;
begin
  Result := (fontid and (CF_ISSUB shl FONT_STYLE_SHIFT) <> 0);
end;

function FontIDIsShadow(const fontid: dword): boolean;
begin
  Result := (fontid and (CF_ISSHADOW shl FONT_STYLE_SHIFT) <> 0);
end;

function FontIDIsExtrude(const fontid: dword): boolean;
begin
  Result := (fontid and (CF_ISEXTRUDE shl FONT_STYLE_SHIFT) <> 0);
end;

function FontIDGetSize (const fontid: dword): integer;
begin
 result := ((fontid shr FONT_SIZE_SHIFT) and $FF);
end;

function FontIDGetColor (const fontid: dword): integer;
begin
 result := ((fontid shr FONT_COLOR_SHIFT) and $FF);
end;

function FontIDGetFace (const fontid: dword): integer;
begin
 result := ((fontid shr FONT_FACE_SHIFT) and $FF);
end;

function DEFAULT_FONT_ID: dword;
begin
  Result := 0;
  FontIDSetSize (Result, 10);
  FontIDSetFaceID (Result, FONT_FACE_SWISS);
end;

(******************************************************************************)
{ from global.h }
(******************************************************************************)
function VARARG_GET (var AP: VARARG_PTR; TypeSz: word): pointer;
begin
  Result := AP;
  Ap := pointer (longint(AP) + TypeSz);
end;

(******************************************************************************)
{ from extmgr.h }
(******************************************************************************)
function EMRegister; stdcall; far; external NOTES_DLL_NAME;
function EMDeregister; stdcall; far; external NOTES_DLL_NAME;
function EMCreateRecursionID; stdcall; far; external NOTES_DLL_NAME;

(******************************************************************************)
{ ns.h }
(******************************************************************************)
function NSGetServerList (pPortName: pchar; retServerTextList: PHandle): STATUS; stdcall; far; external NOTES_DLL_NAME;
function NSGetServerClusterMates (pServerName: pchar; dwFlags: DWORD; var phList: THandle): STATUS; stdcall; far; external NOTES_DLL_NAME;
function NSPingServer (pServerName: pchar; pdwIndex: PDWORD; var phList: THandle): STATUS; stdcall; far; external NOTES_DLL_NAME;

(******************************************************************************)
{ lookup.h }
(******************************************************************************)
function NAMEGetAddressBooks(pszServer: PChar;
                             wOptions: Word;
                             var pwReturnCount: Word;
                             var pwReturnLength: Word;
                             var phReturn: Handle): STATUS; stdcall; far; external NOTES_DLL_NAME;

procedure NAMEGetModifiedTime(var retModified: TIMEDATE); stdcall; far; external NOTES_DLL_NAME;

function NAMELookup(ServerName: PChar;
                    Flags: Word;
                    NumNameSpaces: Word;
                    NameSpaces: PChar;
                    NumNames: Word;
                    Names: PChar;
                    NumItems: Word;
                    Items: PChar;
                    var rethBuffer: Handle): STATUS; stdcall; far; external NOTES_DLL_NAME;

function NAMELocateNextName(pLookup: Pointer;
                             pName: Pointer;
                             retNumMatches: PWord): Pointer; stdcall; far; external NOTES_DLL_NAME;

function NAMELocateNextMatch(pLookup: Pointer;
                              pName: Pointer;
                              pMatch: Pointer): Pointer; stdcall; far; external NOTES_DLL_NAME;

function NAMELocateItem(pMatch: Pointer;
                         Item: Word;
                         var retDataType: Word;
                         retSize: PWord): Pointer; stdcall; far; external NOTES_DLL_NAME;

function NAMEGetTextItem(pMatch: Pointer;
                         Item: Word;
                         Member: Word;
                         Buffer: PChar;
                         BufLen: Word): STATUS; stdcall; far; external NOTES_DLL_NAME;

function NAMELocateMatchAndItem(pLookup: Pointer;
                                MatchNum: Word;
                                Item: Word;
                                var retDataType: Word;
                                retpMatch: Pointer;
                                retpItem: Pointer;
                                var retSize: Word): STATUS; stdcall; far; external NOTES_DLL_NAME;

function ConvertItemToText; external NOTES_DLL_NAME;

(******************************************************************************)
{ dname.h }
(******************************************************************************)
function DNAbbreviate(Flags: LongInt;
                      TemplateName: PChar;
                      InName: PChar;
                      OutName: PChar;
                      OutSize: Word;
                      var OutLength: Word): STATUS; stdcall; far; external NOTES_DLL_NAME;

function DNCanonicalize(Flags: LongInt;
                        TemplateName: PChar;
                        InName: PChar;
                        OutName: PChar;
                        OutSize: Word;
                        var OutLength: Word): STATUS; stdcall; far; external NOTES_DLL_NAME;


function DNParse(Flags: LongInt;
                 TemplateName: PChar;
                 InName: PChar;
                 var Comp: DN_COMPONENTS;
                 CompSize: Word): STATUS; stdcall; far; external NOTES_DLL_NAME;



(******************************************************************************)
{ foldman.h }
(******************************************************************************)
function FolderCreate; stdcall; far; external NOTES_DLL_NAME;
function FolderCopy; stdcall; far; external NOTES_DLL_NAME;
function FolderDocRemove; stdcall; far; external NOTES_DLL_NAME;
function FolderDocAdd; stdcall; far; external NOTES_DLL_NAME;
function FolderDocRemoveAll; stdcall; far; external NOTES_DLL_NAME;
function FolderDocCount; stdcall; far; external NOTES_DLL_NAME;
function FolderDelete; stdcall; far; external NOTES_DLL_NAME;
function FolderMove; stdcall; far; external NOTES_DLL_NAME;
function FolderRename; stdcall; far; external NOTES_DLL_NAME;

(******************************************************************************)
{ acl.h }
(******************************************************************************)
function ACLLookupAccess; external NOTES_DLL_NAME;
function ACLCreate; external NOTES_DLL_NAME;
function ACLAddEntry; external NOTES_DLL_NAME;
function ACLDeleteEntry; external NOTES_DLL_NAME;
function ACLUpdateEntry; external NOTES_DLL_NAME;
function ACLEnumEntries; external NOTES_DLL_NAME;
function ACLGetPrivName; external NOTES_DLL_NAME;
function ACLSetPrivName; external NOTES_DLL_NAME;
function ACLGetHistory; external NOTES_DLL_NAME;
function ACLGetFlags; external NOTES_DLL_NAME;
function ACLSetFlags; external NOTES_DLL_NAME;
function ACLGetAdminServer; external NOTES_DLL_NAME;
function ACLSetAdminServer; external NOTES_DLL_NAME;

{
#define ACLIsPrivSet(privs, num)  ((privs).BitMask[num / 8] &   (1 << (num % 8)))
#define ACLSetPriv(privs, num)    ((privs).BitMask[num / 8] |=  (1 << (num % 8)))
#define ACLClearPriv(privs, num)  ((privs).BitMask[num / 8] &= ~(1 << (num % 8)))
#define ACLInvertPriv(privs, num) ((privs).BitMask[num / 8] ^=  (1 << (num % 8)))
}

function mask (n: integer): word;
begin
  Result := 1 shl (n mod 8);
end;

function ACLIsPrivSet;
begin
  Result := (privs.BitMask[num div 8] and mask(num)) <> 0;
end;

procedure ACLSetPriv;
begin
  privs.BitMask[num div 8] := privs.BitMask[num div 8] or mask(num);
end;

procedure ACLClearPriv;
begin
  privs.BitMask[num div 8] := privs.BitMask[num div 8] and (not (mask(num)));
end;

procedure ACLInvertPriv;
begin
  privs.BitMask[num div 8] := privs.BitMask[num div 8] xor mask(num);
end;

(******************************************************************************)
{ Repl.h }
(******************************************************************************)
function ReplicateWithServer; external NOTES_DLL_NAME;

(******************************************************************************)
{ Edtods.h }
(******************************************************************************)
function GetBorderType (const a: DWORD): DWORD;
begin
//GETBORDERTYPE(a) ((DWORD)((a) & BARREC_BORDER_MASK) >> 28)
  Result := (a and BARREC_BORDER_MASK) shr 28;
end;

procedure SetBorderType (var a: DWORD; const b: DWORD);
begin
//SETBORDERTYPE(a,b) a = ((DWORD)((a) & ~BARREC_BORDER_MASK) | ((DWORD)(b) << 28))
  a := (a and (not BARREC_BORDER_MASK)) or (b shl 28);
end;


(******************************************************************************)
{ undocumented }
(******************************************************************************)
{$IFNDEF NOTES_R4}
function OSGetIniFileName; external NOTES_DLL_NAME;
function OSGetExecutableDirectory; external NOTES_DLL_NAME;
{$ENDIF}

(******************************************************************************)
{ nsfole.h }
(******************************************************************************)
function NSFNoteExtractOLE2Object; external NOTES_DLL_NAME;
function NSFNoteDeleteOLE2Object; external NOTES_DLL_NAME;
function NSFNoteAttachOLE2Object; external NOTES_DLL_NAME;

(******************************************************************************)
{ Agents - Added by Daniel}
(******************************************************************************)
function AgentIsEnabled(hAgent: HAGENT): BOOL; stdcall; far; external NOTES_DLL_NAME;
function AgentOpen (hSrcDB: DBHANDLE; AgentNoteID: NOTEID; var hAgent: HANDLE): STATUS; stdcall; far; external NOTES_DLL_NAME;
function AgentCreateRunContext(hAgent: HANDLE; pReserved: PChar; dwFlags: dword; var rethContext: HANDLE): STATUS; stdcall; far; external NOTES_DLL_NAME;
function AgentRun(hAgent: HAGENT; hAgentCtx: HAGENTCTX; hSelection: HANDLE; dwFlags: dword): STATUS; stdcall; far; external NOTES_DLL_NAME;
function AgentDestroyRunContext(hAgentCtx: HAGENTCTX): STATUS;stdcall; far; external NOTES_DLL_NAME;
function AgentSetDocumentContext(hAgentCtx: HAGENTCTX; hNote: HANDLE): STATUS ;stdcall; far; external NOTES_DLL_NAME;
function AgentClose(hAgent: HANDLE): STATUS;stdcall; far; external NOTES_DLL_NAME;

{$IFNDEF NOTES_R4}
function AgentLSTextFormat(hSrc: Handle;var hDest;var hErrs: HANDLE; dwFlags: dWord; var phData: HANDLE): STATUS; stdcall; far;  external NOTES_DLL_NAME;
function NSFNoteLSCompile(hDb: HANDLE; hNote: HANDLE; dwFlags: DWord): STATUS; stdcall; far; external NOTES_DLL_NAME;
function AgentDelete(hSrc: LHandle): STATUS; stdcall; far; external NOTES_DLL_NAME;
{$ENDIF}

function NSFDbGetObjectSize(
  hDB:DBHANDLE;
  ObjectID:Integer;
  ObjectType:WORD;
  var retSize:DWORD;
  var retClass:WORD;
  var retPrivileges:WORD): STATUS; stdcall; far; external NOTES_DLL_NAME;

function NSFDbReadObject(
  hDB: DBHANDLE;
  ObjectID: Integer;
  Offset,Length:Dword;
  var rethBuffer: HANDLE): STATUS;stdcall; far; external NOTES_DLL_NAME;

function ConvertTIMEDATEToText(
  IntlFormat: pchar;
  var TextFormat: TFMT;
  var InputTime: TIMEDATE;
  var retTextBuffer: Char;
  TextBufferLength: Integer;
  var retTextLength:WORD): STATUS; stdcall; far; external NOTES_DLL_NAME;

// New 2003-09-28/Daniel
function ConvertTextToTIMEDATE(
	IntlFormat: pchar;
	TextFormat: pTFMT;
	Text: pchar;
	MaxLength: Word;
	var retTIMEDATE: TIMEDATE): STATUS; stdcall; far; external NOTES_DLL_NAME;

function NSFDbAllocObject(
  hDB:DBHANDLE;
  dwSize:DWORD;
  aClass:WORD;
  Privileges:WORD;
  var retObjectID:DWORD): STATUS;stdcall; far; external NOTES_DLL_NAME;

function NSFDbFreeObject(
  hDB:DBHANDLE;
  ObjectID:DWORD): STATUS;stdcall; far; external NOTES_DLL_NAME;

function NSFDbWriteObject(
  hDB: DBHANDLE;
  ObjectID: DWord;
        hBuffer:HANDLE;
      Offset:DWORD;
  Length:DWORD): STATUS;stdcall; far; external NOTES_DLL_NAME;

(******************************************************************************)
{ Unknown }
(******************************************************************************)
function NSFDbGetUserActivity(
  hDB: DBHANDLE;
  Flags: DWord;
  retDbActivity: PDBACTIVITY;
  var rethUserInfo: HANDLE;
  var retUserCount: WORD): STATUS; stdcall; far; external NOTES_DLL_NAME;

function NSFGetMaxPasswordAccess(hDB: DBHANDLE; retLevel: Pword): STATUS; external NOTES_DLL_NAME;
function NSFSetMaxPasswordAccess(hDB: DBHANDLE; Level: word): STATUS; external NOTES_DLL_NAME;

(******************************************************************************)
{ Stats.h }
(******************************************************************************)

function StatTraverse(
	Facility: PChar;
        StatName: PChar;
	EnumRoutine: STATTRAVPROC;
 	Context: Pointer): Status; stdcall; far; external NOTES_DLL_NAME;

function StatToText(
	Facility: PChar;
	StatName: PChar;
	ValueType: Word;
	Value: Pointer;
	NameBuffer: PChar;
	NameBufferLen: Word;
	ValueBuffer: PChar;
      	ValueBufferLen: Word): Status; stdcall; far; external NOTES_DLL_NAME;
end.

