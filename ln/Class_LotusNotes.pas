{==============================================================================|
| Project : Notes/Delphi class library                           | 3.11        |
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
|   Sergey Kolchin (Russia) skolchin@yahoo.com                                 |
|   Sergey Kucherov (Russia)                                                   |
|   Sergey Okorochkov (Russia)                                                 |
| All Rights Reserved.                                                         |
|   Last Modified:                                                             |
|     17.11.2011, Sandbil                                                     |
|==============================================================================|
| Contributors and Bug Corrections:                                            |
|   Fujio Kurose                                                               |
|   Noah Silva                                                                 |
|   Tibor Egressi                                                              |
|   Andreas Pape                                                               |
|   Anatoly Ivkov                                                              |
|   Winalot                                                                    |
|   Dmitry Mokrushin                                                           |
|   Olaf Hahnl                                                                 |
|   Daniel Lehtihet                                                            |
|     and others...                                                            |
|==============================================================================|
| History: see README.TXT                                                      |
|==============================================================================|
| Main library unit                                                            |
|                                                                              |
| Use NO_NATIONAL_TEXT define to enable/disable national text processing       |
| Use NO_CTRL_BREAK to disable Control-Break handling                          |
| Use EXTENDED_INIT to pass program arguments to Notes runtime                 |
| Use NOTES_R4 to maintain compatibility with Notes R4                         |
|                                                                              |
| Latest versions are available at                                             |
|   http://www.geocities.com/skolchin/                                         |
|   https://github.com/sandbil/LNotify                                         |
|==============================================================================|}
unit Class_LotusNotes;

{.$DEFINE NO_NATIONAL_TEXT}
{.$DEFINE NO_CTRL_BREAK}
{.$DEFINE NO_INIT_SECTION}
{$DEFINE EXTENDED_INIT}

{$RANGECHECKS OFF}
{$ALIGN OFF}

{$IFDEF NO_INIT_SECTION}
{.$WEAKPACKAGEUNIT ON}
{$ENDIF}

{$INCLUDE Util_LnVersion.inc}

interface

uses
  Messages, SysUtils, Classes, Windows, Util_LNApi {$IFDEF D6}, Variants {$ENDIF};

const
  MAX_STR_LEN = 4096;
  ITEM_VALUE_SEPARATOR = ';';
  LOOKUP_VALUE_END = #9'==='#9;

type
  TNotesItem = class;
  TNotesDocumentCollection = class;
  TNotesNote = class;
  TNotesDocument = class;
  TNotesDatabase = class;
  TNotesACL = class;
  TNotesACLEntry = class;
  PTDateTime = ^TDateTime;

  TNotesItemClass = class of TNotesItem;
  TNotesNoteClass = class of TNotesNote;
  TNotesDocumentClass = class of TNotesDocument;

  ELotusNotes = class(Exception)
    // All LN exceptions
    ErrorCode: dword;
    constructor CreateErr (aCode: dword; aMsg: string);
  end;

  ELnFormulaCompile = class(ELotusNotes)
    // Formula compilation
    CompileErrCode: dword;
    CompileErrMsg: string;
    CompileErrOffset, CompileErrLength: word;
    constructor CreateErr (aCode: dword; aMsg: string; aErrCode: dword; aErrOff,aErrLen: word);
  end;

  ELnInvalidFolder = class(ELotusNotes)
    // Attempt to open a view as a folder
  end;

  ELnUserCancel = class(ELotusNotes)
    // User cancelled password prompt
  end;

  // Document collection
  TNotesDocumentCollection = class(TList)
  private
    Db: TNotesDatabase;
    fUnreadDocs: boolean;
    FDocClass: TNotesDocumentClass;
    FNoteClass: TNotesNoteClass;

    function  GetDocumentId(Idx: integer): longint;
    function  GetDocument(Idx: integer): TNotesDocument;
    function  GetNote(Idx: integer): TNotesNote;
    procedure GetCollectionByHandle(hBuffer: LHANDLE; NotesFound: word; fSummary: boolean);
    function GetSummaryValues (Idx: integer): string;
  public
    constructor Create(notesDatabase: TNotesDatabase);
    destructor Destroy; override;

    {$IFDEF D4}
    procedure Clear; override;
    {$ELSE}
    procedure Clear;
    {$ENDIF}

    procedure MarkAllRead(fRead: boolean);
      // Sets all documents read/unread

    property Database: TNotesDatabase read Db;
    property DocumentID[Ind: integer]: longint read GetDocumentId; default;
    property SummaryValues[Idx: integer]: string read GetSummaryValues;

    property DocumentClass: TNotesDocumentClass read FDocClass write FDocClass;
    property Document[Ind: integer]: TNotesDocument read GetDocument;
      // Returns NEW document object of DocumentClass based on stored ID.
      // Don't forget to free it!

    property NoteClass: TNotesNoteClass read FNoteClass write FNoteClass;
    property Note[Ind: integer]: TNotesNote read GetNote;
      // Returns NEW note object of NoteClass based on stored ID.
      // Don't forget to free it!

    // internal use
    function AddDocumentID(ID: longint; Summary: string): integer;
    procedure DeleteDocumentID(Index: integer);
  end;

  // Notes View (basic class)
  TNotesView = class(TNotesDocumentCollection)
  private
    FName: string;
    FNoteID: NOTEID;
    Flags: string;
    FKey: string;
    FMaxDocs: integer;
    FhColl: HCOLLECTION;

    function GetIsFolder: boolean;
    function GetIsShared: boolean;
    function GetUniversalID: UNID;
    function GetOriginatorID: OID;
  protected
    constructor CreateEmpty (notesDatabase: TNotesDatabase; aName: string);
    class function LoadFlags (Database: TNotesDatabase; anID: NOTEID): string;
    procedure Load(aKey: string; nMaxDocs: integer);
    procedure LoadByID(anID: NOTEID; aKey: string; nMaxDocs: integer; FreeCollectionOnClose:Boolean);
    procedure LoadColl(hcoll: LHandle; aKey: string; nMaxDocs: integer);
  public
    // Normal constructors/destructor
    constructor Create(notesDatabase: TNotesDatabase; aName: string);
    constructor CreateExt(notesDatabase: TNotesDatabase; aName: string; aKey: string; nMaxDocs: integer);
    destructor Destroy; override;

    // Open view and hold collection handle to perform multiple searches (by Matt Saint)
    constructor CreateSearch(notesDatabase: TNotesDatabase; aName: string);

    class function OpenView (notesDatabase: TNotesDatabase; aName: string): TNotesView;
      // use this function to create object of valid view/folder class

    class function OpenViewExt(notesDatabase: TNotesDatabase; aName: string; aKey: string; nMaxDocs: integer): TNotesView;
      // the same as OpenView, but allows to limit view selection by the key or documents count

    procedure Update;
      // Update (refresh) the view

    class function GetViewID(notesDatabase: TNotesDatabase; aName: string): NOTEID;
      // Get note ID of the view

    function GetAllDocumentByKey(akey: string; nmaxdocs:Integer): TNotesDocumentCollection;
      // Quick view search (by Matt Saint)
      // Works only if created with CreateForSearch!

    property Name: string read FName;
    property ID: NOTEID read FNoteID;
    property UniversalID: UNID read GetUniversalID;
    property OriginatorID: OID read GetOriginatorID;
    property IsFolder: boolean read GetIsFolder;
    property IsShared: boolean read GetIsShared;
    property Key: string read FKey;
    property MaxDocs: integer read FMaxDocs;
  end;

  // Notes folder
  TNotesFolder = class(TNotesView)
  private
    procedure SetName(Value: string);
  public
    constructor CreateNew (notesDatabase: TNotesDatabase; aName: string;
      fShared: boolean; FormatFolder: TNotesFolder);

    procedure AddDocument (Doc: TNotesDocument);
    procedure AddDocuments (DocList: TList);

    // replaces old Clear method
    procedure RemoveDocuments;

    function  Copy (NewName: string): TNotesFolder;
    procedure Delete;
    procedure DeleteDocument (Doc: TNotesDocument);
    procedure DeleteDocuments (DocList: TList);
      //WARNING! Free the folder object right after calling this function
    procedure Move (ParentFolder: TNotesFolder);

    property Name: string read FName write SetName;
  end;

  // Document item
  TNotesItem = class(TPersistent)
  private
    fDoc: TNotesNote;
    fName: string;
    fItemBid: BlockId;
    fValueBid: BlockId;
    fItemflags: word;
    fDataType: word;
    fSaveToDisk: boolean;
    fValueLength: DWORD;
    fIsNew: boolean;
    fCreated: boolean;
    FStringsValue: TStrings;
    FSeqNo: integer;
    function GetAsFloat : extended;
    function GetAsList : variant;
    function GetAsNumbers : variant;
    function GetAsString : string;
    procedure SetItemFlags (anItemFlag: integer; anValue: boolean);
    function  GetItemFlags (anItemFlag: integer): boolean;
    function  GetLastModifed: TDateTime;
    function  GetAsDateTime: TDateTime;
    function GetAsReference: UNID;
    function GetAsStrings : TStrings;
    function GetAsTimes : variant;
    procedure SetAsDateTime(Value: TDateTime);
    function  GetItemValue: variant;
    procedure SetAsFloat (Value: extended);
    procedure SetAsList (Value: variant);
    procedure SetAsNumbers (Value: variant);
    procedure SetAsReference(Value: UNID);
    procedure SetAsPChar(buffer: PChar);
    procedure SetAsString (Value: string);
    procedure SetAsStrings (Value: TStrings);
    procedure SetAsTimes (Value: variant);
    procedure SetItemValue (aValue: variant);
    procedure SetCreated;
    function GetDoc: TNotesDocument;
    procedure SetAsString2(const Value: string);
  protected
    function GetRichText: TStrings; virtual;
    procedure SetRichText (Value: TStrings); virtual;
    procedure InitItemInfo(PrevItem: BLOCKID); virtual;
  public
    constructor Create(notesDocument: TNotesNote; aName: string); virtual;
    constructor CreateNew (notesDocument: TNotesNote; aName: string); virtual;
    constructor CreateNext (notesItem: TNotesItem); virtual;
    destructor Destroy; override;

    // Update item definition from a note (only for existing items)
    procedure Refresh;

    // Item flags
    property IsAuthors    : boolean index ITEM_READWRITERS  read GetItemFlags write SetItemFlags;
    property IsEncrypted  : boolean index ITEM_SEAL         read GetItemFlags write SetItemFlags;
    property IsNames      : boolean index ITEM_NAMES        read GetItemFlags write SetItemFlags;
    property IsProtected  : boolean index ITEM_PROTECTED    read GetItemFlags write SetItemFlags;
    property IsReaders    : boolean index ITEM_READERS      read GetItemFlags write SetItemFlags;
    property IsSigned     : boolean index ITEM_SIGN         read GetItemFlags write SetItemFlags;
    property IsSummary    : boolean index ITEM_SUMMARY      read GetItemFlags write SetItemFlags;
    property IsNewItem    : boolean read fCreated;
    property ItemFlags    : word read fItemflags write fItemFlags;

    property LastModified : TDateTime read GetLastModifed;
    property Name: string read fName;
    property SaveToDisk: boolean read fSaveToDisk write fSaveToDisk;
    property Note: TNotesNote read fDoc;
    property Document: TNotesDocument read GetDoc;

    property ItemType: word read fDataType write fDataType; //see TYPE_... in Util_NotesAPI
    property ValueLength: DWORD read fValueLength;
    property Value: Variant read GetItemValue write SetItemValue;

    property AsPChar: PChar write SetAsPChar;
    property AsString: string read GetAsString write SetAsString;
    property AsDateTime: TDateTime read GetAsDateTime write SetAsDateTime;
    property AsNumber: extended read GetAsFloat write SetAsFloat;
    property AsList: variant read GetAsList write SetAsList;
    property AsNumbers: variant read GetAsNumbers write SetAsNumbers;
    property AsStrings: TStrings read GetAsStrings write SetAsStrings;
    property AsTimes: variant read GetAsTimes write SetAsTimes;
    property AsRichText: TStrings read GetRichText write SetRichText;
    property AsReference: UNID read GetAsReference write SetAsReference;
    property AsShortString: string read GetAsString write SetAsString2;

    // Low-level value access- don't use
    property ItemBid: BLOCKID read fItemBid;
    property ValueBid: BLOCKID read fValueBid;
    procedure GetValueBuffer(var Buffer: pointer; var BufSize: dword);
    procedure SetValueBuffer(iType,wFlags: word;  Buffer: pointer;  BufSize: dword);

    // Multiple items support
    function CreateNextItem: TNotesItem;
    function NextItemExists: boolean;
    function LoadNextItem: boolean;
    property SeqNo: integer read FSeqNo;
  end;

  // Notes note
  TNotesNote = class(TPersistent)
  private
    fHandle: LHandle;
    FId: NOTEID;
    FFields: TStringList;
    FDatabase: TNotesDatabase;
    FIsDeleted: boolean;
    FKeepHandle: boolean;
    function GetFieldCount : longint;
    function GetFieldName (Index: integer): string;
    function GetItemByName (ItemName: string): TNotesItem;
    function GetLastModified: TDateTime;
    function GetLastAccessed: TDateTime;
    function GetUniversalID: UNID;
    function GetOriginatorID: OID;
    function GetCreated: TDateTime;
    function GetSize: longint;
  protected
    constructor CreateEmpty(notesDatabase: TNotesDatabase);

    procedure InitDocument(notesDatabase: TNotesDatabase; anID: dword); virtual;
    procedure CreateDocument(notesDatabase: TNotesDatabase); virtual;
  public
    constructor CreateNew(notesDatabase: TNotesDatabase);
      // create new DOCUMENT in the database (see also TNotesDatabase.CreateDocument)
    constructor Create(notesDatabase: TNotesDatabase; anId: NOTEID);
      // open a document by its NoteID
    constructor CreateByUNID(notesDatabase: TNotesDatabase; anUNID: UNID);
      // open a document by its Universal ID
    constructor CreateFromHandle(notesDatabase: TNotesDatabase; aHandle: LHandle);
      // open a document using its handle
    destructor Destroy; override;

    {$IFDEF D4}
    procedure Save(force: boolean = False; createResponse: boolean = False; markRead: boolean = False); virtual;
    {$ELSE}
    procedure Save(force, createResponse, markRead: boolean); virtual;
    {$ENDIF}

    function GetSignature (var SignedBy: string; var CertifiedBy: string; pTime: PTDateTime): boolean;
      // Get signature information. SignedBy is user name, CertifiedBy - domain name. pTime can be nil
      // Returns True if document is signed and signature is valid
    procedure Sign; //signs a document

    function CopyToDatabase(DestDB:TNotesDatabase): TNotesDocument;
    function Evaluate(aFormula: string): variant;
      // Evaluates a formula on the document

    // Note properties
    property DocID: NOTEID read fID;
    property Handle: LHandle read fHandle;
    property Database: TNotesDatabase read fDatabase;
    property LastModified: TDateTime read GetLastModified;
    property LastAccessed: TDateTime read GetLastAccessed;
    property UniversalID: UNID read GetUniversalID;
    property OriginatorID: OID read GetOriginatorID;
    property Created: TDateTime read GetCreated;
    property Size: longint read GetSize;

    // Items
    // Use Items[Name].As... to read/write document properties
    property FieldCount: longint read GetFieldCount;
    property FieldName[Index: integer]: string read GetFieldName;
    property Items[ItemName: string]: TNotesItem read GetItemByName; default;

    procedure DeleteItem(ItemName: string);
    function IsItemExists(ItemName: string): boolean;
    function ReplaceItemValue(itemName: string; value: Variant): TNotesItem;
    procedure CopyItem (Source: TNotesDocument; itemName: string);

    // Multiple items support
    function CountMultipleItems(aName: string): integer;
        //returns a number of items with given name
    function LoadMultipleItems(aName: string): TList;
        //loads all items with given name and returns a list of them
        //the resulting list contains item objects, which must be free'd
        //manually, along with the list itself

    // Clears fields list and reloads it from the document
    // Call Save before to commit changes
    procedure ReloadFields;

    // By Matt Saint - Is document deleted?
    // This property is TRUE when a note opened from collection has already
    // been deleted from database
    property IsDeleted: boolean read FIsDeleted;
  end;

  // Notes document
  TNotesDocument = class(TNotesNote)
  private
    FAttach: TStrings;
    FFontTable: TStrings;
    FProfileName: string;
    FSummary: TStrings;
    FIsRead: boolean;
    function GetAttachment(Index: integer): string;
    function GetAttachmentCount: integer;
    function GetRecipients: string;
    procedure SetRecipients(Value: string);
    function GetBodyAsString: string;
    procedure SetBodyAsString(Value: string);
    function GetBodyAsMemo: TStrings;
    procedure SetBodyAsMemo(Value: TStrings);
    procedure SetIsRead (Value: boolean);
    function GetResponsesCount: DWORD;
  protected
    function GetItemByNum(ItemNum: integer): string;
    procedure SetItemByNum(ItemNum: integer; Value: string);

    procedure InitDocument(notesDatabase: TNotesDatabase; anID: dword); override;
    procedure CreateDocument(notesDatabase: TNotesDatabase); override;
  public
    FMaxAttachment: word; // for RTI, don't use
    FMaxFontID: word;

    constructor CreateProfile (notesDatabase: TNotesDatabase; aProfileName: string);
      // opens a profile document
    constructor CreateResponse(notesDatabase: TNotesDatabase;  MainDoc: TNotesDocument);
      // creates a new response document
    constructor CreateResponseByUNID(notesDatabase: TNotesDatabase; const anUNID: UNID);
      // creates a new response document with UNID only - Subject is not set!
    destructor Destroy; override;

    // These are Document attachment handling
    property AttachmentCount: integer read GetAttachmentCount;  //number of file attachments
    property Attachment[Index: integer]: string read GetAttachment;
    {$IFDEF D4}
    function Attach(FileName: string; DisplayName: string = ''): integer; //attach a file by its name
    {$ELSE}
    function Attach(FileName: string; DisplayName: string): integer; //attach a file by its name
    {$ENDIF}
    procedure Detach(Index: integer; FileName: string);
      //detach a file attached to a form with a given name
    function FindAttachment(aName: string): integer;  //return index in Attachment or -1
    procedure DeleteAttachment(Index: integer);
      //delete an attachment

    {$IFDEF D4}
    procedure Send(fAttachForm: boolean = False; ARecipients: string = '');
      // sends a document. if ARecipients <> '', overrides previously assigned addressees
      // if Database.SaveMail = True, also saves a document
    procedure Save(force: boolean = False; createResponse: boolean = False; markRead: boolean = False); override;
    {$ELSE}
    procedure Send(fAttachForm: boolean; ARecipients: string);
    procedure Save(force, createResponse, markRead: boolean); override;
    {$ENDIF}

    procedure CheckAddress;
      // Checks names assigned to SendTo and Recipients fields agains server address book

    procedure AttachForm (aForm: string);
    function ComputeWithForm (doDataTypes, raiseError : boolean) : boolean;

    property ProfileName: string read FProfileName;
      // blank for non-profile documents

    // Special fields for mail documents
    // Note that using this fields do not change doc items obtained from Items property
    property Form: string index MAIL_FORM_ITEM_NUM read GetItemByNum write SetItemByNum; //form name
    property Subject: string index MAIL_SUBJECT_ITEM_NUM read GetItemByNum write SetItemByNum; //subject field
    property SendTo: string index MAIL_SENDTO_ITEM_NUM read GetItemByNum write SetItemByNum;
    property CopyTo: string index MAIL_COPYTO_ITEM_NUM read GetItemByNum write SetItemByNum;
    property BlindCopyTo: string index MAIL_BLINDCOPYTO_ITEM_NUM read GetItemByNum write SetItemByNum;
      // Addresses, multiple ones are separated by ','.
    property MailFrom: string index MAIL_FROM_ITEM_NUM read GetItemByNum write SetItemByNum; // sender
    property Recipients: string read GetRecipients write SetRecipients;
      // List of recipients. Don't change manually, it's generated from SendTo + CopyTo + BlindCopyTo
      // If set, these properties are ignored
    property Body: TStrings read GetBodyAsMemo write SetBodyAsMemo;   // body field
    property BodyAsString: string read GetBodyAsString write SetBodyAsString;

    // Summary values are available only when a document was read from view (TNotesView class)
    // This class contains text representations of all values placed in summary buffer upon save
    // It's not updated during document saving
    property SummaryValues: TStrings read FSummary;

    // Font table access, internal use
    property FontTable: TStrings read FFontTable;
    procedure LoadFontTable;
    procedure SaveFontTable;

    // Update unread documents list
    procedure UpdateUnread;

    // Is the document read?
    property IsRead: boolean read FIsRead write SetIsRead;

    // By Matt Saint - hierarchy management
    property ResponsesCount: DWORD read GetResponsesCount;
    function GetParentDocumentID: NOTEID;
    function Responses: TNotesDocumentCollection;
      // returns NEW object, has to be deleted by the caller
  end;

  // Notes database
  TNotesDatabase = class
  private
    FACL: TNotesACL;
    FHandle: integer;
    FFileName: string;
    FServerName: string;
    FSaveMail: boolean;
    FKeepHandle: boolean;
    FViews: TStrings;
    function GetACL: TNotesACL;
    function GetFullName: string;
    function GetActive: boolean;
    function GetDatabaseID: DBID;
    function GetReplicaInfo: DBReplicaInfo;
    procedure SetActive(Value: boolean);
    function GetInfo(Index: integer): string;
    function GetViewByIndex (Index: integer): TNotesView;
    function GetViews (ViewName: string): TNotesView;
    function GetViewCount: integer;
    function GetAllViewNames: TStrings;
    procedure SetInfo(Index: integer; Value: string);
    procedure SetFileName(Value: string);
    procedure SetReplicaInfo(Value: DBReplicaInfo);
    procedure SetServerName(Value: string);
    function GetQuotaInfo: DBQUOTAINFO;
    function GetDBActivityInfo: DBACTIVITY;
    function GetUserActivity: DBUserActivityArray;
  protected
    procedure CheckViews;
    procedure UpdateViews;

    function IntFindNotes (formula: string; notesDateTime: TDateTime; noteClass: word;
      proc: NSFSEARCHPROC; fSummary: boolean): TNotesDocumentCollection;
  public
    constructor Create; // this doesn't open or create a database!
    destructor Destroy; override;

    function CreateDocument: TNotesDocument;
      // creates new empty document in the database
      // use Form property to assign the form
    function CreateResponseDocument (const aResUNID: UNID; aSubject: string): TNotesDocument;
      // creates a new response for given document. See also CreateResponse constructors

    // Find a note (generic)
    function FindNotes(formula: string; notesDateTime: TDateTime; noteClass: word;
      fSummary: boolean): TNotesDocumentCollection;

    function Open(aServer, dbFile: string): boolean;
      // open specified database. Use blank server name to open from disk
    procedure OpenMail;
      // Open mailbox. Equal to Open (MailServer, MailFile);
    function OpenByHandle(aHandle: LHandle): boolean;
      // attach to a database handle

    procedure Close;
      // close the database
    procedure CloseSession;
      // close the database and break a connection
    procedure CopyRecords(SourceDB: TNotesDatabase);
      // copies all records from SourceDB
    procedure CreateNew(aServer, dbFile: string; TemplateDB: string);
      // creates a new database with given name. TemplateDB is full name of template or ''
    procedure SendMail (Address: string; Subject: string; Body: string);
      // Simple mail sending
    procedure ReplyMail(Mail: TNotesDocument;  Body: string);
      // Simple reply
    procedure Delete(DocID: integer); //delete a document from database
    procedure DeleteDocument(Doc: TNotesDocument);
      // Delete an open document. Use this function if a document is akredy opened
      // WARNING! After deleting, you cannot access document properties!

    procedure ListAddressBooks (aServer: string; List: TStrings; fGetTitles: boolean);
      // list AB for specific server or for local system if aServer=""
      // if fGetTitles=False, List will contain only paths to address books
      // otherwise, it will be <path>=<title>
    procedure OpenPrivateAddressBook;
      // open local address book
    function GetLocationDocument: TNotesDocument;
      // open a location document
      // the database must be opened using OpenPrivateAddressBook

    function FTSearch(query: string; maxDocs, sortOptions, otherOptions: integer): TNotesDocumentCollection;
      // Search by Full-text index
      // maxDocs limits returned docs number (now unused)
      // for sortOptions and otherOptions look in Notes documentation
    function FindDocument(Formula: string): TNotesDocument;
      // finds one document

    {$IFDEF D4}
    function Search(formula: string; notesDateTime: TDateTime = 0; maxDocs: integer = 0): TNotesDocumentCollection;
      // searches for documents by given formula.
      // notesDateTime limits documents by creation date (since ...)
      // maxDocs limits returned docs number (now unused)
    {$ELSE}
    function Search(formula: string; notesDateTime: TDateTime; maxDocs: integer): TNotesDocumentCollection;
    {$ENDIF}

    function UnreadDocuments : TNotesDocumentCollection;
    function UnreadDocumentsUserName(UName:string) : TNotesDocumentCollection;
      // Returns collection of all unread documents in a database
    procedure MarkRead (NoteID: dword; fRead: boolean);
      // Changes given note status. Use MarkAllRead to process many docs
    procedure MarkAllRead (Docs: TNotesDocumentCollection; fRead: boolean);
      // Marks all docs read/unread

    // Views
    property ViewCount: integer read GetViewCount;
    property ViewByIndex[Index: integer]: TNotesView read GetViewByIndex;
    property Views[ViewName: string]: TNotesView read GetViews;
      // Don't free objects!
    property ViewNames: TStrings read GetAllViewNames; //new Daniel

    function OpenView(AName: string; notesDateTime: TDateTime; maxDocs: integer): TNotesView;
      // obsolete - use Views property instead
      // must free returned object
      // notesDateTime and maxDocs are unused

    // ACL
    property ACL: TNotesACL read GetACL;
      // Returns ACL for this database
      // If ACL doesn't exist it will be created
    procedure CopyACL(SourceDb: TNotesDatabase);
      //copies an ACL of a database to this one

    property Handle: integer read fHandle; //handle for direct access
    property SaveMail: boolean read FSaveMail write FSaveMail; //if true, mailed docs are also saved in the database
    property FileName: string read FFileName write SetFileName; //db file name
    property Server: string read FServerName write SetServerName; //db server name
    property Active: boolean read GetActive write SetActive;
    property FullName: string read GetFullName;
    property Title: string index INFOPARSE_TITLE read GetInfo write SetInfo;  //database title
    property DesignClass: string index INFOPARSE_CLASS read GetInfo write SetInfo;  //template name for templates
    property DesignTemplate: string index INFOPARSE_DESIGN_CLASS read GetInfo write SetInfo;  //inherited template name

    class function UserName: string;        //name of logged user
    class function LmbcsUserName: string;   //name of logged user as LMBSC string
    class function MailFileName: string;    //name of mailbox database
    class function MailServer: string;      //mail server
    class function MailType: integer;       //mail type (0 - Domino, 1 - SMTP)

    function NotesVersion: word;
    property DatabaseID: DBID read GetDatabaseID;
    property ReplicaInfo: DBReplicaInfo read GetReplicaInfo write SetReplicaInfo;

    // Quotas
    property QuotaInfo: DBQuotaInfo read GetQuotaInfo;

    // User and DB Activity
    property ActivityInfo: DBACTIVITY read GetDBActivityInfo;
    property UserActivity: DBUserActivityArray read GetUserActivity;
  end;

  // One directory entry
  TNotesDirEntry = record
    FileName: shortString;
    EntryType: boolean;         //False for file, True for directory
    FileInfo: shortString;      //description of the database
  end;
  pNotesDirEntry = ^TNotesDirEntry;

  // Search options
  // nfoFiles and nfoTemplates are mutually exclusive. If both are set, nfoFiles is used
  TNotesFindOption = (nfoFiles, nfoTemplates, nfoSubDirs);
  TNotesFindOptions = set of TNotesFindOption;

  //This class allows to list servers/directories/databases}
  TNotesDirectory = class
  private
    FPorts: TStrings;
    hDirectory: THandle;
    SrcServer: string;
    SrcPath: string;
    SrcOptions: TNotesFindOptions;
    SrcTable: TList;
    SrcIndex: integer;

    function GetPorts: TStrings;
  public
    constructor Create;
    destructor Destroy; override;

    procedure FindClose;
    // closes search sequence

    function FindFirst (const Server, Path: string; Options: TNotesFindOptions; var Entry: TNotesDirEntry): boolean;
    function FindNext (var Entry: TNotesDirEntry): boolean;
      // use these functions to find files in the directory
      // they returns False if no more files can be found

    procedure ListServers (Port: string; List: TStrings);
      // lists all servers for specific port (if empty string - for all ports)
      // at least one entry always exists: empty string for local computer

    property Ports: TStrings read GetPorts;
      // available ports
  end;

  // Options for LookupName funcion
  TNotesLookupOption = (
    nloAll,        //returns all entries of given category
    nloNoSearch,   //look only in the first AB
    nloExhaustive  //search in all AB on the server
  );
  TNotesLookupOptions = set of TNotesLookupOption;

  { Lotus Notes name parsing }
  TNotesName = class
  private
    FName: string;  //canonical
    FComponents: DN_COMPONENTS;
    fParsed: boolean;
    FTemplateName: string;

    function GetAbbreviatedName: string;
    function GetKeyword : string;
    procedure SetAbbreviatedName (Value: string);
    procedure SetName (Value: string);
  public
    constructor Create (aName: string);
      // The name must be in either canonical or abbreviated format
      // You can use function IsCanonical to determine name type

    class function IsCanonical (aName: string): boolean;
    function IsHirerarchical : boolean;
      // Returns True if the name is in canonical distinguished format

    class function TranslateName (aName: string; fToCanonical: boolean; aTemplate: string): string;
      // Translation function. If flag=True, tries to canonilize, otherwise - abbreviate

    function GetNameComponent (Index: integer): string;
    procedure SetNameComponent (Index: integer; Value: string);
      // Common name parts access functions

    property Common:     string index 0 read GetNameComponent write SetNameComponent;
    property Given:      string index 1 read GetNameComponent write SetNameComponent;
    property Surname:    string index 2 read GetNameComponent write SetNameComponent;
    property Initials:   string index 3 read GetNameComponent write SetNameComponent;
    property Generation: string index 4 read GetNameComponent write SetNameComponent;
    property Country:    string index 5 read GetNameComponent write SetNameComponent;
    property OrgUnit1:   string index 6 read GetNameComponent write SetNameComponent;
    property OrgUnit2:   string index 7 read GetNameComponent write SetNameComponent;
    property OrgUnit3:   string index 8 read GetNameComponent write SetNameComponent;
    property OrgUnit4:   string index 9 read GetNameComponent write SetNameComponent;
    property ADMD:       string index 10 read GetNameComponent write SetNameComponent;
    property PRMD:       string index 11 read GetNameComponent write SetNameComponent;
      // these two are mutually exclusive
    property Organization: string index 12 read GetNameComponent write SetNameComponent;

    property Abbreviated: string read GetAbbreviatedName write SetAbbreviatedName;
    property Canonical: string read FName write SetName;
    property Keyword: string read GetKeyword;
    property TemplateName: string read FTemplateName write FTemplateName;
      // template name used if a name contains no parts except user name (like 'Sergey Kolchin/')

    // ****************
    // Lookup functions
    class function LookupName (ServerName: string; NameSpaces: array of string;
      Names: array of string; Items: array of string; Flags: TNotesLookupOptions;
      Values: TStrings): boolean;
      // Generic function for name lookup in the address book (AB)
      // ServerName defines location of AB ('' for local)
      // NameSpaces define view in AB. Use USER_NAMESPACE or '' for default
      // Names is an array of names to look up. Names may = [''] if Flags specifies NAME_LOOKUP_ALL
      // Items is an array of item names to return with each match (cannot be [''])
      // Flags defines options for lookup
      // The function returns True on success and fills Values with all matched values found
      // Note that the function can handle only text items

    class function LookupNameList (ServerName: string; NameSpaces: TStrings;
      Names: TStrings; Items: TStrings; Flags: TNotesLookupOptions;
      Values: TStrings): boolean;
      // The same as above, but takes lists as parameters

    class function CheckAddress (aServer, aName: string): boolean;
      // Simple way to check mail address
      // Checks for existence of the name in the AB on the server
  end;

  // ACL access levels and flags
  TNotesAclAccessLevel = (
    aclNoAccess,
    aclDepositor,
    aclReader,
    aclAuthor,
    aclEditor,
    aclDesigner,
    aclManager
  );

  TNotesAclFlag = (
    acfAuthorNoCreate,        //Authors can't create new notes (only edit existing ones) }
    acfServer,                //Entry represents a Server (V4) }
    acfNoDelete,              //User cannot delete notes }
    acfCreatePersonalAgents,  //User can create personal agents (V4) }
    acfCreatePersonalFolders, //User can create personal folders (V4) }
    acfPerson,                //Entry represents a Person (V4) }
    acfGroup,                 //Entry represents a group (V4) }
    acfCreateFolders,         //User can create and update shared views & folders (V4)
    acfCreateLotusScript,     //User can create LotusScript }
    acfPublicReader,          //User can read public notes }
    acfPublicWriter,          //User can write public notes }
    acfAdminReaderAuthor,     //Admin server can modify reader and author fields in db }
    acfAdminServer            //Entry is administration server (V4) }
  );
  TNotesAclFlags = set of TNotesAclFlag;

  // ACL
  TNotesACL = class
  private
    FDatabase: TNotesDatabase;
    FEntries: TList;
    FHandle: LHandle;
    FRoles: TStrings;
    function GetEntriesCount: integer;
    function GetEntry (aName: string): TNotesACLEntry;
    function GetEntryByIndex (Index: integer): TNotesACLEntry;
    function GetUniformAccess: boolean;
    procedure ReadEntries;
    procedure SetUniformAccess (Value: boolean);
    function GetMaxInternetAccess: TNotesAclAccessLevel;
    procedure SetMaxInternetAccess(const Value: TNotesAclAccessLevel);
  public
    constructor Create (aDatabase: TNotesDatabase);
    destructor Destroy; override;
    procedure Save;

    function CreateACLEntry (aName: string): TNotesACLEntry;
    procedure DeleteACLEntry (aName: string);
    property EntriesCount: integer read GetEntriesCount;
    property Entry[aName: string]: TNotesACLEntry read GetEntry; default;
    property EntryByIndex[Index: integer]: TNotesACLEntry read GetEntryByIndex;

    property Database: TNotesDatabase read FDatabase;
    property Handle: LHandle read FHandle;
    property Roles: TStrings read FRoles;
    property UniformAccess: boolean read GetUniformAccess write SetUniformAccess;
    property MaximumInternetAccess: TNotesAclAccessLevel read GetMaxInternetAccess write SetMaxInternetAccess;
  end;

  // ACL entry
  TNotesACLEntry = class
  private
    FAccessLevel: TNotesAclAccessLevel;
    FACL: TNotesACL;
    FFlags: TNotesAclFlags;
    FName, FOldName: string;
    FPrivileges: ACL_PRIVILEGES;
    FNew: boolean;
    FUpdateFlags: word;

    constructor Create(anACL: TNotesACL; aName: string; AccLevel: TNotesAclAccessLevel;
      const AclPrivs: ACL_PRIVILEGES; AccFlags: TNotesAclFlags);
    function GetRoles: string;
    procedure SetAccessLevel (Value: TNotesAclAccessLevel);
    procedure SetFlags (Value: TNotesAclFlags);
    procedure SetName (Value: string);
    procedure SetPrivileges (Value: ACL_PRIVILEGES);
    procedure SetRoles (Value: string);
  public
    procedure AddRole (aRole: string);
    constructor CreateNew(anACL: TNotesACL; aName: string; AccLevel: TNotesAclAccessLevel; aRoles: string);
    procedure DeleteRole (aRole: string);
    destructor Destroy; override;

    procedure Update;

    property AccessLevel: TNotesAclAccessLevel read FAccessLevel write SetAccessLevel;
    property ACL: TNotesACL read FACL;
    property Flags: TNotesAclFlags read FFlags write SetFlags;
    property Name: string read FName write SetName;
    property Privileges: ACL_PRIVILEGES read FPrivileges write SetPrivileges;
    property Roles: string read GetRoles write SetRoles;
  end;

  type
    PStatRecord = ^TStat;
    TStat = record
      Facility: String;
      StatName: String;
      NameBuffer: String;
      ValueBuffer: String;
    end;

  // Collection of statistical data. Used only with TNotesServer class
  TStatCollection = class
  private
    FCount: word;  // Current element
    FList: TList;
    function GetHasMoreElements: boolean;
    function GetMaxCount: word;
  public
    constructor Create;
    destructor Destroy; override;
    
    property HasMoreElements: boolean read GetHasMoreElements;
    property Count: word read GetMaxCount;
    function NextElement: TStat;
    procedure Add(aStatRec: TStat);
    procedure Clear;
  end;

  // NotesServer class
  TNotesServer = class
  private
    FServerName: String;
    StatCollection: TStatCollection;
    function GetHasMoreStatistics: boolean;
    procedure parseToList( sList: TStrings );
  public
    // Give ServerName parameter constructor when you want to remotely extract
    // and parse server stats.
    // Use empty ServerName constructor when you build a server add-in
    {$IFDEF D4}
    constructor Create(ServerName: string = '');
    {$ELSE}
    constructor Create(ServerName: string);
    {$ENDIF}
    destructor Destroy; override;

    function NextStat: TStat;
    function GetConsoleInfo(cmd: String): String;
    function QueryLocalStatistics(Facility, StatName: PChar): Word;
    function QueryRemoteStatistics: Word;
    property HasMoreStatistics: boolean read GetHasMoreStatistics;
  end;

// Thread initialization
// InitNotesThread must be called by the thread before using of any Notes functions
// CloseNotesThread must be called by the thread before thread terminates
procedure InitNotesThread;
procedure CloseNotesThread;

// Standard initialization function
procedure InitNotes;

// Extended initialization function
// Uses NotesInitExtended instead of old NotesInit
// converting program arguments to C-style argc, argv ones
// Might not be compatible with custom program arguments
// If EXTENDED_INIT is defined, is used in the library initialization
procedure InitNotesExt;

// Date time conversions
function DateTimeToNotes (DelphiTime: TDateTime): TIMEDATE;
function NotesToDateTime (NotesTime: TIMEDATE): TDateTime;
function NotesToDateTimeEx (NotesTime: TimeStruct): TDateTime;

// UNID conversion
{$IFDEF D3}
function UNIDtoStr (const anUNID: UNID; fDelimiters: boolean): string;
{$ELSE}
function UNIDtoStr (const anUNID: UNID; fDelimiters: boolean = True): string;
{$ENDIF}
function TimeDatetoStr (const anID: TIMEDATE): string;

function StrtoUNID (const IDStr: String): UNID;
// Text conversion
// Work only if LN_NATIONAL_TEXT is defined
// Written by Fujio Kurose(fujio.kurose@nifty.ne.jp)
function Lmbcs2Native (aString: string): string;
function Native2Lmbcs (aString: string): string;
Function UserNameFromID(IDfile:string):string;
// String conversions
//   Converts Notes-style string (where #0 marks line-feeds) to Delphi
function NotesToString(NotesStr: pchar; StrSize: integer): string;
//   Opposite conversion. NotesStr is allocated internally and MUST BE FREED by FreeMem()
//   Returns True if NotesStr has embedded zeroes
function StringToNotes(DelphiStr: string; var NotesStr: pchar; var StrSize: integer): boolean;

// Color mapping
// TColor is not used because it's declared in Graphics unit
function NotesToColor(NotesColor: integer): integer;
function ColorToNotes(DelphiColor: integer): integer;

// File paths...
function ConstructPath (const aServer, aPath: string): string;
procedure ParsePath (const aPath: string; var aServer, aFile: string);
function GetNotesDataDir: string;
{$IFNDEF NOTES_R4}
function GetNotesExeDir: string;
function GetNotesIniFile: string;
{$ENDIF}

// Adds a large item to note by separating it to multiple items with the same name
procedure HugeNSFItemAppend(hNote: NOTEHANDLE;
                          ItemFlags: Word;
                          Name: PChar;
                          NameLength: Word;
                          DataType: Word;
                          Value: Pointer;
                          ValueLength: LongInt);

// Control-break signal handler function (internal use)
function NDBreakProc: STATUS; far; stdcall;

// N/D library version
const NDLib_Version: string = '3.10';

implementation
uses Util_LNApiErr, Class_NotesRTF;

(******************************************************************************)
// Exceptions
(******************************************************************************)
constructor ELotusNotes.CreateErr (aCode: dword; aMsg: string);
begin
  inherited Create(aMsg);
  ErrorCode := aCode;
end;

(******************************************************************************)
constructor ELnFormulaCompile.CreateErr;
var
  buf: string;
begin
  setLength(buf, 255);
  OSLoadString(0, aErrCode, PChar(buf), 254);
  buf := Lmbcs2Native(StrPas(PChar(buf)));
  inherited CreateErr (aCode, format ('%s: %s (at %d:%d)',[aMsg, buf, aErrOff,aErrLen]));
  CompileErrCode := aErrCode;
  CompileErrMsg := buf;
  CompileErrOffset := aErrOff;
  CompileErrLength := aErrLen;
end;

(******************************************************************************)
// Init and callbacks
(******************************************************************************)
var
  InitDone: boolean;
  
procedure InitNotesThread;
begin
  CheckError (NotesInitThread);
end;

(******************************************************************************)
procedure CloseNotesThread;
begin
  NotesTermThread;
end;

(******************************************************************************)
procedure InitNotes;
var
  r: LResult;
begin
  r := NotesInit;
  if r <> ERROR_SUCCESS then raise ELotusNotes.CreateErr(r, 'Initialization error');
  InitDone := true;
end;

(******************************************************************************)
procedure InitNotesExt;
var
  i, argc: integer;
  alen, dlen: longint;
  s: string;
  argv, parg: ppchar;
  pmem, pdata: pchar;
  r: LResult;
begin
  // Get total length of arguments
  argc := ParamCount;
  dlen := 0;
  for i := 0 to argc do inc(dlen, length(ParamStr(i)) + 1);
  alen := (argc+1) * sizeof(pchar);
  inc(dlen);

  // Get block of memory
  GetMem(argv, alen);
  GetMem(pmem, dlen);
  try
    FillChar(argv^, alen, 0);
    FillChar(pmem^, dlen, 0);
    
    // Prepare params
    parg := argv;
    pdata := pmem;
    for i := 0 to argc do begin
      s := ParamStr(i);
      parg^ := pdata;
      pdata := StrECopy(pdata, pchar(s));
      inc(pdata);
      inc(parg);
    end;

    // Call init function
    r := NotesInitExtended(argc+1, argv);
    if r <> ERROR_SUCCESS then raise ELotusNotes.CreateErr(r, 'Initialization error');
    InitDone := true;
  finally
    FreeMem(argv);
    FreeMem(pmem);
  end;
end;

(******************************************************************************)
function  FieldsScanProc(Spare, ItemFlags: word;
                              Name: PChar;
                              NameLength: word;
                              Value: pointer;
                              ValueLength: dword;
                              RoutineParameter: pointer): STATUS;  stdcall;
begin
  TStringList(RoutineParameter).Add(
                           Lmbcs2Native(Copy (strPas(name), 1, NameLength)));
  //TStringList (RoutineParameter).Add(Copy (strPas(name), 1, NameLength));
  Result := NOERROR;
end;

(******************************************************************************)
// Conversion procedures
(******************************************************************************)
function DateTimeToNotes (DelphiTime: TDateTime): TIMEDATE;
var
  T: TimeStruct;
  x,y,z,hs: word;
begin
  //OSCurrentTimeDate (@T.GM);
  //TimeGMToLocal (T);
  FillChar (T, sizeOf(T), #0);
  DecodeDate (DelphiTime,x,y,z);
  T.Year := x;
  T.Month := y;
  T.Day := z;
  DecodeTime(DelphiTime,x,y,z,hs);
  T.hour := x;
  T.minute := y;
  T.second := z;
  T.hundredth := hs div 10;
  //T.weekday := DayOfWeek (DelphiTime);
  OSCurrentTimeZone (@T.zone, @T.dst);
  TimeLocalToGM(T);
  Result := T.GM;
end;

(******************************************************************************)
function NotesToDateTime (NotesTime: TIMEDATE): TDateTime;
var
  T: TimeStruct;
begin
  T.GM := NotesTime;
  OSCurrentTimeZone (@T.zone, @T.dst);
  Result := NotesToDateTimeEx (T);
end;

(******************************************************************************)
{function NotesToDateTimeEx (NotesTime: TimeStruct): TDateTime;
begin
  TimeGMToLocal(NotesTime);
  Result := 0;
  if NotesTime.year>0 then Result := EncodeDate(NotesTime.year,NotesTime.month,NotesTime.day);
  if NotesTime.hour>0 then Result := Result + EncodeTime(NotesTime.hour,NotesTime.minute,
    NotesTime.second,NotesTime.hundredth);
end;}

(******************************************************************************)
function NotesToDateTimeEx (NotesTime: TimeStruct): TDateTime;
var
  failed: boolean;
  tm: TimeStruct;
begin
  Result := 0;   //default

  // Try this first, but note that dates without a time are undefined
  //(from API reference)
  tm := NotesTime;
  failed := TimeGMToLocalZone(tm);

  if (not failed) and ((tm.year = -1) or (tm.month = -1) or (tm.day = -1)) then begin
    // Time-only field, only Local time conversion works - experimental
    failed := TimeGMToLocal(NotesTime);
  end
  else begin
    NotesTime := tm;
     //failed := TimeGMToLocalZone (NotesTime);

     //If TimeGMToLocalZone failed, use TimeGMToLocal instead
    if failed then failed := TimeGMToLocal(NotesTime);
  end;

   //One of the above was successful, so convert to TDateTime
   if not failed then try
      if (NotesTime.year = -1) or (NotesTime.month = -1) or (NotesTime.day = -1)
        then Result := 0
        else Result := EncodeDate (NotesTime.year, NotesTime.month, NotesTime.day);
      if NotesTime.hour = -1 then NotesTime.hour := 0;
      if NotesTime.minute = -1 then NotesTime.minute := 0;
      if NotesTime.second = -1 then NotesTime.second := 0;
      if NotesTime.hundredth = -1 then NotesTime.hundredth := 0;
      Result := Result + EncodeTime (NotesTime.hour, NotesTime.minute,
                                     NotesTime.second, NotesTime.hundredth);
   except
   end;
end;

var
  OldBreakProc: OSSIGBREAKPROC;

(******************************************************************************)
function NDBreakProc: STATUS; far; stdcall;
begin
  if (GetAsyncKeyState(VK_CANCEL) and 1) = 0
    then Result := NOERROR
    else Result := ERR_CANCEL;
end;

(******************************************************************************)
// Translation
(******************************************************************************)

Function UserNameFromID(IDfile:string):string;
var
sz:word;
Buffer:string;
begin
    setLength (Buffer, MAXUSERNAME);
    CheckError(REGGetIDInfo( pchar(Native2Lmbcs(IDfile)), REGIDGetName, pchar(Buffer), MAXUSERNAME, @sz));
    Buffer:= strPas(pchar(buffer));
    if TNotesName.IsCanonical (Buffer) then
    Buffer := Lmbcs2Native(TNotesName.TranslateName (Buffer, false, ''));
  result:= Buffer;
end;
//***************************************************
function Lmbcs2Native;
var
  n: dword;
begin
{$IFNDEF NO_NATIONAL_TEXT}
  n := length(aString);
  Result := '';
  SetLength(Result, n + 2);
  n := OSTranslate(
    OS_TRANSLATE_LMBCS_TO_NATIVE,
    PChar(AString),
    n,
    PChar(Result),
    n);
  Result[n+1] := #0;
  Result := strPas(pchar(Result));
{$ELSE}
  Result := aString;
{$ENDIF}
end;

(******************************************************************************)
function Native2Lmbcs;
var
  n: dword;
begin
{$IFNDEF NO_NATIONAL_TEXT}
  n := length(aString) * 3; //LMBCS is ~3 times bigger - Fujio Kurose
  Result := '';
  SetLength(Result, n + 2);
  n := OSTranslate(
    OS_TRANSLATE_NATIVE_TO_LMBCS,
    PChar(AString),
    length(aString),
    PChar(Result),
    n);
  Result[n+1] := #0;
  Result := strPas(pchar(Result));
{$ELSE}
  Result := aString;
{$ENDIF}
end;

(******************************************************************************)
// Misc
(******************************************************************************)
function ConstructPath (const aServer, aPath: string): string;
begin
  setLength (Result, 256);
  CheckError (OSPathNetConstruct('', pchar(aServer), pchar(Native2Lmbcs(aPath)), pchar(Result)));
  Result := StrPas(PChar(Lmbcs2Native(Result)));
end;

(******************************************************************************)
procedure ParsePath (const aPath: string; var aServer, aFile: string);
var
  port: string;
begin
  setLength (aServer, 256);
  setLength (aFile, 256);
  setLength (port, 256);
  CheckError (OSPathNetParse(pchar(Native2Lmbcs(aPath)),pchar(port),pchar(aServer),pchar(aFile)));
  aServer := StrPas(PChar(Lmbcs2Native(aServer)));
  aFile := StrPas(PChar(Lmbcs2Native(aFile)));
end;

(******************************************************************************)
{$IFNDEF NOTES_R4}
function GetNotesExeDir: string;
begin
  setlength(Result,MAXPATH+1);
  OSGetExecutableDirectory (pchar(Result));
  Result := StrPas(PChar(Result));
  if (Result <> '') and (Result[length(Result)] <> '\') then appendStr(Result,'\');
end;
{$ENDIF}

(******************************************************************************)
{$IFNDEF NOTES_R4}
function GetNotesIniFile: string;
begin
  setlength(Result,MAXPATH+1);
  OSGetIniFileName(pchar(Result));
  Result := StrPas(PChar(Result));
end;
{$ENDIF}

(******************************************************************************)
function GetNotesDataDir: string;
begin
  Result:= '';
  setLength(Result, MAXPATH+1);
  if OSGetEnvironmentString( 'Directory', pchar(Result), MAXPATH)
    then Result := strPas(pchar(Result))
    else Result := '';
  if (Result <> '') and (Result[length(Result)] <> '\') then appendStr(Result,'\');
end;


(******************************************************************************)
// Buffer-to-value conversion
(******************************************************************************)
function BufferToValue (Buffer: pointer; BufLen: dword): variant;
var
  ItemType, nEntry, i, wLen: word;
  str: string;
  num: NUMBER;
  tm: TIMEDATE;
  pc: pchar;
  PRnge: PRANGE;
  PNumValue: PNUMBER;
  PDtValue: PTIMEDATE;
  buf: pchar;
begin
  Result := NULL;
  ItemType := pword(Buffer)^;
  Buffer := pointer (dword(Buffer) + sizeOf(word));
  dec (BufLen, sizeOf(word));
  case ItemType of
    TYPE_TEXT: begin
      GetMem(buf, BufLen+1);
      try
        Move(Buffer^, buf^, BufLen);
        buf[BufLen] := #0;
        Result := Lmbcs2Native(NotesToString(buf, BufLen));
      finally
        FreeMem(buf);
      end;
    end;
    TYPE_TEXT_LIST: begin
      nEntry := ListGetNumEntries(Buffer, False);
      if nEntry > 1 then Result := VarArrayCreate ([0, nEntry-1], varOleStr);
      for i := 0 to nEntry-1 do begin
        ListGetText (Buffer, False, i, @pc, @wLen);
        {setLength (str, wLen + 2);
        strLCopy (pchar(str), pc, wLen);
        str[wLen+1] := #0;
        if nEntry > 1
          then Result[i] := Lmbcs2Native(strPas(pchar(str)))
          else Result := Lmbcs2Native(strPas(pchar(str)));}
        GetMem(buf, wLen+1);
        try
          Move(pc^, buf^, wLen);
          buf[wLen] := #0;
          str := Lmbcs2Native(NotesToString(buf, wLen));
          if nEntry > 1
            then Result[i] := Lmbcs2Native(str)
            else Result := Lmbcs2Native(str);
        finally
          FreeMem(buf);
        end;
      end;
    end;
    TYPE_NUMBER: begin
      num := PNUMBER(Buffer)^;
      Result := num;
    end;
    TYPE_NUMBER_RANGE: begin
      PRnge := PRANGE (Buffer);
      PNumValue := PNUMBER (dword(Buffer) + sizeOf(USHORT)*2);
      if PRnge^.ListEntries = 0 then Result := Null
      else if PRnge^.ListEntries = 1 then Result := PNumValue^
      else begin
        Result := VarArrayCreate ([0, PRnge^.ListEntries-1], varDouble);
        for i := 0 to PRnge^.ListEntries-1 do begin
          Result[i] := PNumValue^;
          PNumValue := PNUMBER (dword(PNumValue) + sizeof(NUMBER));
        end;
      end;
    end;

    TYPE_TIME: begin
      tm := PTIMEDATE(Buffer)^;
      Result := NotesToDateTime (tm);
    end;

    TYPE_TIME_RANGE: begin
      PRnge := PRANGE (Buffer);
      PDtValue := PTIMEDATE (dword(Buffer) + sizeOf(USHORT)*2);
      if PRnge^.ListEntries = 0 then Result := Null
      else if PRnge^.ListEntries = 1 then Result := VarFromDateTime (NotesToDateTime (PDtValue^))
      else begin
        Result := VarArrayCreate ([0, PRnge^.ListEntries-1], varDate);
        for i := 0 to PRnge^.ListEntries-1 do begin
          Result[i] := VarFromDateTime (NotesToDateTime (PDtValue^));
          PDtValue := PTIMEDATE (dword(PDtValue) + sizeof(TIMEDATE));
        end;
      end;
    end;
    else Result := UNASSIGNED;
  end;
end;

(******************************************************************************)
function UNIDtoStr;
begin
  If fDelimiters
    then Result := intToHex(anUNID.aFile.T2, 8) + ':' +
                   intToHex(anUNID.aFile.T1, 8) + '-' +
                   intToHex(anUNID.Note.T2, 8) + ':' +
                   intToHex(anUNID.Note.T1, 8)
    else Result := intToHex(anUNID.aFile.T2, 8) +
                   intToHex(anUNID.aFile.T1, 8) +
                   intToHex(anUNID.Note.T2, 8) +
                   intToHex(anUNID.Note.T1, 8);
end;
(******************************************************************************)
function TimeDatetoStr;
begin
    Result := intToHex(anID.T2, 8) +
                   intToHex(anID.T1, 8);
end;

(******************************************************************************)
// New 2003-09-28/Daniel
type
  THexValue=0..15;

function CharToHex(c:char):THexValue;
begin
  c:=upcase(c);
  Result:=0;
  if c in ['0'..'9']then
    Result:=ord(c)-ord('0')
  else if c in ['A'..'F']then
    Result:=ord(c)-ord('A')+10;
end;
// New 2003-09-28/Daniel
function HexToInt(hex:string):integer;
var
  i:byte;
  HexFactor:integer;
begin
  result:=0;
  HexFactor:=1;
  if length(hex)=0 then exit;
  for i:=1 to length(hex)do
    begin
      inc(result,CharToHex(hex[length(hex)-pred(i)])*HexFactor);
      HexFactor:=HexFactor*16;
    end;
end;

// New 2003-09-28/Daniel
// Converta string representation of a UniversalID into a
// UNID structure
function StrtoUNID;
begin
        if pos(':', IDStr) > 0 then
        begin   // no : och - in the string representation
               Result.aFile.T2 := HexToInt(Copy(IDStr, 0, 8));
               Result.aFile.T1 := HexToInt(Copy(IDStr, 10, 8));
               Result.Note.T2 := HexToInt(Copy(IDStr, 19, 8));
               Result.Note.T1 := HexToInt(Copy(IDStr, 28, 8));
        end else
        begin  // : and - in here
               Result.aFile.T2 := HexToInt(Copy(IDStr, 0, 8));
               Result.aFile.T1 := HexToInt(Copy(IDStr, 9, 8));
               Result.Note.T2 := HexToInt(Copy(IDStr, 17, 8));
               Result.Note.T1 := HexToInt(Copy(IDStr, 26, 8));
        end;
end;

{ Table conversion }
type
  TConvProc = procedure (Name: string; Value: variant; Data: pointer);

// This function loops throught the item table and calls Proc for each item found
// passing the item's name, it's value (usually string) and supplied context data
procedure ReadItemTable (pTable: PITEM_TABLE; Proc: TConvProc; Data: pointer);
var
  n: integer;
  ptItem: PITEM;
  pSummary: pchar;
  name_buf: string;
  val: variant;
  wSize: dword;
begin
  ptItem := PITEM(dword(pTable) + ODSLength(_ITEM_TABLE));
  wSize := ODSLength(_ITEM_TABLE) + ODSLength(_ITEM) * pTable^.Items;
  pSummary := pchar(dword(pTable) + wSize);
  for n := 1 to pTable^.Items do begin
    name_buf := '';
    setLength(name_buf, ptItem^.NameLength + 2);
    strLCopy(pchar(name_buf), pSummary, ptItem^.NameLength);
    name_buf[ptItem^.NameLength+1] := #0;
    name_buf := Lmbcs2Native(strPas(pchar(name_buf)));

    val := BufferToValue (pchar (dword(pSummary) + ptItem^.NameLength), ptItem^.ValueLength);
    Proc(name_buf, val, Data);
    inc (wSize, ptItem^.NameLength + ptItem^.ValueLength);
    if wSize > pTable^.Length then break;
    pSummary := pchar(dword(pSummary) + ptItem^.NameLength + ptItem^.ValueLength);
    ptItem := PITEM(dword(ptItem) + ODSLength(_ITEM));
  end;
end;

//****************************************************
// Item table callback procedure
// Stores all passed items and values to one string in form <Name>=<Val>
procedure ReadItemTableAsStringProc (Name: string; Value: variant; Data: pointer);
var
  buf: string;
  i: integer;
begin
  buf := '';
  if (not VarIsNull(Value)) and (not VarIsEmpty(Value)) then try
    if (VarType(Value) and varArray) = 0 then buf := VarAsType(Value, varString)
    else for i := varArrayLowBound(Value,1) to varArrayHighBound(Value,1) do
      appendStr(buf, VarAsType(Value[i], varString) + ';');
  except
    // Ignore conversion errors
    buf := '';
  end;
  pstring(Data)^ := pstring(Data)^ + #13#10 + Name + '=' + buf;
end;

//**********************************************
// Adds an item with unlimited length - by Andy
procedure HugeNSFItemAppend(hNote: NOTEHANDLE;
                          ItemFlags: Word;
                          Name: PChar;
                          NameLength: Word;
                          DataType: Word;
                          Value: Pointer;
                          ValueLength: LongInt);
const
  MaxBufLen = 65000;
var
  r: longint;
  b: longint;
  l: longint;
  p: pointer;
begin
  if ValueLength <= MaxBufLen then CheckError (NSFItemAppend(hNote, ItemFlags, Name, NameLength,
    DataType, Value, ValueLength))
  else begin
    r := ValueLength;   // the rest of the length
    b := 0;             // the startposition of the pointer in  FContext
    repeat
      if r > MaxBufLen
        then l := MaxBufLen // the length to be inserted now
        else l := r;        // the length to be inserted now
      p:= pointer (longint (Value) + b);
      Inc (b, l);
      CheckError (NSFItemAppend(hNote, ItemFlags, Name, NameLength,
      DataType, p, l ));
      Dec (r, l);
    until (r = 0);
  end;
end;

(******************************************************************************)
{ String conversion }
(******************************************************************************)
function NotesToString(NotesStr: pchar; StrSize: integer): string;
var
  p: pchar;
  s: string;
  n, ofs: integer;
begin
  // Notes uses #0 as CRLF
  p := NotesStr;
  Result := '';
  ofs := 0;
  if StrSize = 0 then exit;

  // Expand
  while True do begin
    s := strPas(p);
    n := length(s);
    if (ofs + n) >= StrSize then begin
      appendStr(Result,copy(s,1,StrSize-ofs));
      break;
    end;
    inc(ofs, n + 1);
    p := pchar(dword(NotesStr) + ofs);
    appendStr(Result, s + #13#10);
  end;
end;

(******************************************************************************)
function StringToNotes(DelphiStr: string; var NotesStr: pchar; var StrSize: integer): boolean;
var
  i, n: integer;
begin
  // Remove LFs leaving CRs
  n := Pos(#10, DelphiStr);
  while n > 0 do begin
    delete(DelphiStr,n,1);
    n := Pos(#10, DelphiStr);
  end;

  // Copy to buffer and replace CRs with #0
  StrSize := length(DelphiStr) + 2;
  GetMem(NotesStr,StrSize);
  Result := False;
  try
    strCopy(NotesStr,pchar(DelphiStr));
    for i := 0 to StrSize-3 do
      if NotesStr[i] = #13 then begin
        NotesStr[i] := #0;
        Result := True;
      end;
    if Result then NotesStr[StrSize-1] := #0   //double-zero string
    else begin
      // No internal zeros - normal string
      dec(StrSize);
      ReallocMem(NotesStr,StrSize);
    end;
  except
    FreeMem(NotesStr);
    raise;
  end;
end;

(******************************************************************************)
{ Color conversion }
(******************************************************************************)
const
  ColorMap: array[0..MAX_NOTES_SOLIDCOLORS-1] of integer = (
    $000000, $FFFFFF, $0000FF, $00FF00, $FF0000, $FF00FF, $00FFFF, $FFFF00,
    $000080, $008000, $800000, $800080, $008080, $808000, $808080, $C0C0C0
  );
const
  UNKNOWN_COLOR = $1FFFFFFF;

(******************************************************************************)
function NotesToColor(NotesColor: integer): integer;
begin
  if NotesColor < MAX_NOTES_SOLIDCOLORS
    then Result := ColorMap[NotesColor]
    else Result := UNKNOWN_COLOR;
end;

(******************************************************************************)
function ColorToNotes(DelphiColor: integer): integer;
begin
  for Result := System.Low(ColorMap) to System.High(ColorMap) do
    if ColorMap[Result] = DelphiColor then exit;
  Result := UNKNOWN_COLOR;
end;

(******************************************************************************)
// TNotesDocumentCollection
(******************************************************************************)
type
  TIDInfo = class
    ID: NOTEID;
    Summary: string;
  end;

(******************************************************************************)
procedure TNotesDocumentCollection.Clear;
var
  i: integer;
begin
  for i := count-1 downto 0 do begin
    TIDInfo(Items[i]).free;
    //FreeMem(PTIdInfo(Items[i]));
    Items[i] := nil;
  end;
  inherited Clear;
end;

(******************************************************************************)
constructor TNotesDocumentCollection.Create;
begin
  inherited Create;
  Db := notesDatabase;
  FDocClass := TNotesDocument;
  FNoteClass := TNotesDocument;
end;

(******************************************************************************)
procedure TNotesDocumentCollection.DeleteDocumentID;
begin
  TIDInfo(Items[Index]).free;
  Delete(Index);
end;

(******************************************************************************)
function TNotesDocumentCollection.GetDocumentId;
begin
  Result := TIdInfo(Items[Idx]).ID;
end;

(******************************************************************************)
function TNotesDocumentCollection.GetDocument;
begin
  Result := DocumentClass.Create(Db,TIdInfo(Items[Idx]).ID);
  Result.FIsRead := not FUnreadDocs;
  if (not Result.FIsDeleted) and (Result.FSummary <> nil) then 
    Result.FSummary.Text := TIdInfo(Items[Idx]).Summary;
end;

(******************************************************************************)
function TNotesDocumentCollection.GetNote;
begin
  Result := NoteClass.Create(Db,TIdInfo(Items[Idx]).ID);
end;

(******************************************************************************)
function TNotesDocumentCollection.AddDocumentId;
var
  pRec: TIdInfo;
begin
  Result := -1;
  if ID <> 0 then begin
    pRec := TIDInfo.create;
    pRec.ID := ID;
    pRec.Summary := Summary;
    Result := Add(pRec);
  end;
end;

(******************************************************************************)
destructor TNotesDocumentCollection.Destroy;
begin
  Clear;
  inherited Destroy;
end;

//****************************************************
procedure TNotesDocumentCollection.GetCollectionByHandle;
var
  i: integer;
  ptList: PNOTEID;
  ptItems: PITEM_TABLE;
  ID: NOTEID;
  Summary: string;
begin
  PtList := PNOTEID(OSLockObject(hBuffer));
  try
    for i := 1 to NotesFound do begin
      ID := PtList^;
      ptList := PNOTEID(dword(PtList) + ODSLength(_NOTEID));
      ptItems := PITEM_TABLE(PtList);
      Summary := '';
      if fSummary then begin
        ReadItemTable (ptItems, ReadItemTableAsStringProc, @Summary);
        PtList := PNOTEID(dword(PtList) + PITEM_TABLE(PtList)^.Length);
      end;
      if ((ID and NOTEID_CATEGORY) = 0) and ((ID and NOTEID_CATEGORY_TOTAL) = 0) then begin
        AddDocumentID (ID, Summary);
      end;
    end;
  finally
    OSUnLockObject(hBuffer);
  end;
end;

(******************************************************************************)
function TNotesDocumentCollection.GetSummaryValues;
begin
  Result := TIdInfo(Items[Idx]).Summary;
end;

//****************************************************
procedure TNotesDocumentCollection.MarkAllRead;
begin
  Database.MarkAllRead (Self, fRead);
end;

(******************************************************************************)
// TNotesDatabase
(******************************************************************************)
constructor TNotesDatabase.Create;
begin
  inherited Create;
end;

(******************************************************************************)
destructor TNotesDatabase.Destroy;
begin
  Close;
  inherited Destroy;
end;

(******************************************************************************)
function TNotesDatabase.UnreadDocuments;
var
  hTable: LHANDLE;
  id: dword;
  fFirst: boolean;
  uName: string;
begin
  // Get IDs
  uName := LmbcsUserName;
  if not TNotesName.IsCanonical (uName) then uName := TNotesName.TranslateName (uName, True, '');
  Result := TNotesDocumentCollection.create (self);
  Result.fUnreadDocs := True;
  try
    hTable := 0;
    CheckError (NSFDbGetUnreadNoteTable2(Handle, pchar(uName), length(uName), True, True, @hTable));
    if hTable <> 0 then try
      CheckError (NSFDbUpdateUnread(Handle, hTable));

      // Scan ID table
      id := 0;
      fFirst := True;
      while IDScan (hTable, fFirst, @id) do begin
        fFirst := False;
        Result.AddDocumentId (id, '');
      end;
    finally
      IDDestroyTable(hTable);
    end;
  except
    Result.free;
    raise;
  end;
end;
(******************************************************************************)
function TNotesDatabase.UnreadDocumentsUserName(uName:string) : TNotesDocumentCollection;
var
  hTable: LHANDLE;
  id: dword;
  fFirst: boolean;
  //uName: string;
begin
  // Get IDs
  //uName := LmbcsUserName;
  if not TNotesName.IsCanonical (uName) then uName := TNotesName.TranslateName (uName, True, '');
  Result := TNotesDocumentCollection.create (self);
  Result.fUnreadDocs := True;
  try
    hTable := 0;
    CheckError (NSFDbGetUnreadNoteTable2(Handle, pchar(uName), length(uName), True, True, @hTable));
    if hTable <> 0 then try
      //CheckError (NSFDbUpdateUnread(Handle, hTable));

      // Scan ID table
      id := 0;
      fFirst := True;
      while IDScan (hTable, fFirst, @id) do begin
        fFirst := False;
        Result.AddDocumentId (id, '');
      end;
    finally
      IDDestroyTable(hTable);
    end;
  except
    Result.free;
    raise;
  end;
end;

(******************************************************************************)
procedure TNotesDatabase.UpdateViews;
var
  i: integer;
begin
  if (FViews <> nil) then begin
    for i := 0 to FViews.count-1 do FViews.Objects[i].free;
    FViews.free;
    FViews := nil;
  end;
end;

(******************************************************************************)
class function TNotesDatabase.UserName;
begin
  Result := Lmbcs2Native(LmbcsUserName);
end;

(******************************************************************************)
class function TNotesDatabase.LmbcsUserName;
begin
  setlength(Result,MAX_STR_LEN);
  SECKFMGetUserName(pchar(Result));
  Result := StrPas(PChar(Result));
end;

(******************************************************************************)
class function TNotesDatabase.MailFileName;
begin
  setlength(Result,MAXENVVALUE+1);
  OSGetEnvironmentString (MAIL_MAILFILE_ITEM, pchar(Result), MAXENVVALUE);
  Result := StrPas(PChar(Result));
end;

(******************************************************************************)
class function TNotesDatabase.MailServer;
begin
  setlength(Result,MAXENVVALUE+1);
  OSGetEnvironmentString (MAIL_MAILSERVER_ITEM, pchar(Result), MAXENVVALUE);
  Result := StrPas(PChar(Result));
end;

//***************************************************
procedure TNotesDatabase.ListAddressBooks;
var
  flags,wCount,wLength, i: word;
  hReturn: LHANDLE;
  pBuf, pSrv: pchar;
  buf, title: string;
begin
  if not fGetTitles
    then flags := 0
    else flags := NAME_GET_AB_TITLES or NAME_DEFAULT_TITLES;
  if aServer = '' then pSrv := nil else pSrv := pchar(aServer);
  CheckError (NAMEGetAddressBooks (pSrv, flags, wCount, wLength, hReturn));
  if hReturn <> 0 then try
    pBuf := OsLockObject(hReturn);
    for i := 1 to wCount do begin
      buf := strPas(pBuf);
      pBuf := pchar(dword(pBuf) + length(buf) + 1);
      if fGetTitles then begin
        title := strPas(pBuf);
        appendStr (buf, '=' + title);
        pBuf := pchar(dword(pBuf) + length(title) + 1);
      end;
      List.add (Lmbcs2Native (buf));
    end;
  finally
    OsUnlockObject (hReturn);
    OsMemFree (hReturn);
  end;
end;

//***************************************************
function TNotesDatabase.Open;
var
  openingmail:boolean;
  Error: word;
  FilePath: string;
begin
  Close;
  openingmail := false;
  if UpperCase(aServer) = 'LOCAL' then aServer := '';
  if (aServer = '') and (dbFile = '') then begin
    aServer := Lmbcs2Native(MailServer);
    dbFile := Lmbcs2Native(MailFileName);
    openingmail := true;
  end;
  if aServer=''
    then FilePath := dbFile
    else FilePath := ConstructPath (TNotesName.TranslateName((aServer), False, ''), dbFile);

  FilePath := Native2Lmbcs(FilePath);
  error := NSFDbOpen(pchar(FilePath), @FHandle);
  if (error <> 0) and OpeningMail and (error <> USER_CANCEL) then begin
    { Try to open local mailbox }
    aServer := '';
    error := NSFDbOpen(pchar(dbfile), @FHandle);
  end;
  if ((Handle=0) and (error=0)) or (error=USER_CANCEL) then raise ELnUserCancel.CreateErr(-1,'User cancelled Notes session');
  CheckError (Error);

  FFileName := dbFile;
  FServerName := aServer;
  FKeepHandle := false;
  Result := true;
end;

(******************************************************************************)
procedure TNotesDatabase.OpenMail;
begin
  Open('','');
end;

//***************************************************
function TNotesDatabase.OpenByHandle;
var
  fn, port: string;
begin
  Close;
  FHandle := aHandle;
  FKeepHandle := true;

  // Get file name and path
  setLength(fn, MAXPATH);
  CheckError(NSFDbPathGet(FHandle, nil, pchar(fn)));

  setLength(port, MAXPATH);
  setLength(FServerName, MAXPATH);
  setLength(FFileName, MAXPATH);
  CheckError(OSPathNetParse(pchar(fn),pchar(port),pchar(FServerName),pchar(FFileName)));
  FServerName := strPas(pchar(FServerName));
  FFileName := strPas(pchar(FFileName));

  Result := true;
end;

(******************************************************************************)
procedure TNotesDatabase.Close;
var
  i: integer;
begin
  if not Active then exit;
  if not FKeepHandle then CheckError(NSFDbClose(FHandle));
  FHandle := 0;
  FAcl.free;
  FAcl := nil;
  if FViews <> nil then begin
    for i := 0 to FViews.count-1 do FViews.Objects[i].free;
    FViews.free;
    FViews := nil;
  end;
end;

(******************************************************************************)
procedure TNotesDatabase.CloseSession;
var
  i: integer;
begin
  if not Active then exit;
  if not FKeepHandle then begin
    CheckError(NSFDbCloseSession(FHandle));
    NSFDbClose(FHandle);
  end;
  FHandle := 0;
  FAcl.free;
  FAcl := nil;
  if FViews <> nil then begin
    for i := 0 to FViews.count-1 do FViews.Objects[i].free;
    FViews.free;
    FViews := nil;
  end;
end;

(******************************************************************************)
procedure TNotesDatabase.SetActive;
begin
  if Value <> Active then
    if Value then Open (Server, FileName) else Close;
end;

(******************************************************************************)
function TNotesDatabase.GetACL;
begin
  if FAcl = nil then FAcl := TNotesACL.create(self);
  Result := FACL;
end;

(******************************************************************************)
procedure TNotesDatabase.CopyACL;
begin
  // By Robbert de Zeeuw
  CheckError (NSFDbCopyACL (SourceDb.Handle, Handle));
end;

(******************************************************************************)
function TNotesDatabase.GetActive;
begin
  Result := Handle <> 0;
end;

//***************************************************
function TNotesDatabase.FindDocument;
var
  Coll: TNotesDocumentCollection;
begin
  Coll := Search (Formula, 0, 1);
  try
    if (Coll = nil) or (Coll.Count = 0)
      then Result := nil
      else Result := Coll.Document[0];
  finally
    Coll.free;
  end;
end;

(******************************************************************************)
function TNotesDatabase.FTSearch;
var
  HSearch: THandle;
  retNumDocsFound: dword;
  retHResults: THandle;
  //hIdTable: THandle;
  pSearchResults: ^FT_SEARCH_RESULTS;
  pNoteId: ^NoteId;
  i: integer;
begin
  Result := TNotesDocumentCollection.Create (self);
  try
    checkError(FTOpenSearch (@HSearch));
    Query := Native2Lmbcs(Query);
    try
      checkError(Util_LNApi.FTSearch(Handle,
                                    @HSearch,
                                    NullHandle,
                                    PChar(Query),
                                    SortOptions or
                                      OtherOptions or
                                      FT_SEARCH_SCORES or
                                      FT_SEARCH_STEM_WORDS,
                                    0,
                                    NullHandle,
                                    @retNumDocsFound,
                                    nil,
                                    @retHResults
                                    ));

      // Well, I don't know why they created an ID table...
      // KOL - 09.03.99
      //checkError(IdCreateTable (sizeOf(NoteId),@hIdTable));
      //try

      pSearchResults := OsLockObject (retHResults);
      try
        pNoteId := pointer(pSearchResults);
        pNoteId := pointer(dword(pNoteId) + sizeOf (FT_SEARCH_RESULTS));
        for i := 0 to pSearchResults^.NumHits-1 do begin
          Result.AddDocumentId(pNoteID^, '');
          inc(pNoteId);
        end;
      finally
        OsUnlockObject (retHResults);
        OsMemFree (retHResults);
      end;

      {finally
        OsMemFree (hIdTable);
      end;}

    finally
      FTCloseSearch (HSearch);
    end;
  except
    Result.free;
    raise;
  end;
end;

//****************************************************
procedure TNotesDatabase.ReplyMail;
var
  MailDoc: TNotesDocument;
begin
  MailDoc := TNotesDocument.CreateResponse(self, Mail);
  try
    MailDoc.Form := 'Reply';
    MailDoc.SendTo := Mail.MailFrom;
    MailDoc.Subject := 'Re: ' + Mail.Subject;
    MailDoc.BodyAsString := Body;
    MailDoc.Sign;
    MailDoc.Send (False, '');
  finally
    MailDoc.free;
  end;
end;

//***************************************************
function TNotesDatabase.Search;
begin
  Result := FindNotes (formula, notesDateTime, NOTE_CLASS_DATA, True);
end;

//****************************************************
function TNotesDatabase.GetViewByIndex;
begin
  CheckViews;
  Result := TNotesView(FViews.Objects[Index]);
  if Result = nil then begin
    Result := TNotesView.OpenView(self, FViews[Index]);
    FViews.Objects[Index] := Result;
  end;
end;

//****************************************************
procedure TNotesDatabase.CheckViews;
var
  Coll: TNotesDocumentCollection;
  i: integer;
  Doc: TNotesDocument;
begin
  if FViews = nil then begin
    FViews := TStringList.create;
    Coll := FindNotes('@All', 0, NOTE_CLASS_VIEW or NOTE_CLASS_PRIVATE, False);
    try
      for i := 0 to Coll.count-1 do begin
        Doc := Coll.Document[i];
        try
          FViews.Add (Doc['$TITLE'].AsString);
        finally
          Doc.free;
        end;
      end;
    finally
      Coll.free;
    end;
  end;
end;

//****************************************************
function TNotesDatabase.GetViews;
// Changed by ap@svd-online.com
const
  AliasSeparator = '|';
var
  i: integer;
  aPos: integer;
begin
  CheckViews;
//  i := FViews.indexOf(ViewName);
//  if i = -1 then Result := nil else Result := ViewByIndex[i];

  Result := nil;
  I      := FViews.IndexOf(ViewName);
  if (I >= 0) then Result := ViewByIndex[I]
  else for I := 0 to FViews.Count - 1 do begin
    aPos := Pos (AliasSeparator, FViews.Strings[I]);
    if (aPos > 0) then begin
      if (0 = SysUtils.CompareText(ViewName,
        trim(Copy(FViews.Strings[I], aPos + 1, System.Length(FViews.Strings[I]) - aPos))))
        then begin
          Result:= ViewByIndex[I];
          Break;
        end;
    end;{of if}
  end;{of for}
end;

//****************************************************
function TNotesDatabase.GetViewCount;
begin
  CheckViews;
  Result := FViews.count;
end;

//****************************************************
function TNotesDatabase.GetAllViewNames;
begin
  CheckViews;
  Result := FViews;
end;

(******************************************************************************)
function SearchProc(Obj: pointer; search_info: pSEARCH_MATCH; summary_info: pITEM_TABLE): STATUS; far; stdcall;
var
  summary: string;
begin
  if ((search_info.SERetFlags and SE_FMATCH) <> 0) then begin
    summary := '';
    if summary_info <> nil then begin
      ReadItemTable(summary_info, ReadItemTableAsStringProc, @summary);
    end;
    TNotesDocumentCollection(Obj).AddDocumentId(search_info^.ID.NoteID, summary);
  end;
  Result := NOERROR;
end;

//****************************************************
function TNotesDatabase.FindNotes;
begin
  Result := IntFindNotes(formula, notesDateTime, noteClass, SearchProc, fSummary);
end;

//****************************************************
function TNotesDatabase.IntFindNotes;
var
  wdc2: status;
  wdc1,wdc3,wdc4,wdc5,wdc6, flags: WORD;
  tdSince: TIMEDATE;
  pSince: PTIMEDATE;
  formula_handle: FORMULAHANDLE;
begin
  Result := TNotesDocumentCollection.Create(self);
  try
    if Formula='' then Formula := '@All';
    if fSummary then flags := SEARCH_SUMMARY else flags := 0;
    if notesDateTime = 0 then pSince := nil
    else begin
      tdSince := DateTimeToNotes(notesDateTime);
      pSince := @tdSince;
    end;

    Formula := Native2Lmbcs(Formula);
    try
      CheckError(NSFFormulaCompile(
                 nil,                   { name of formula (none) }
                 WORD(0),               { length of name }
                 PChar(Formula),        { the ASCII formula }
                 WORD(length(Formula)), { length of ASCII formula }
                 @formula_handle,       { handle to compiled formula }
                 @wdc1,                 { compiled formula length (don't care) }
                 @wdc2,                  { return code from compile (don't care) }
                 @wdc3, @wdc4, @wdc5, @wdc6)); { compile error info (don't care) }

      CheckError(NSFSearch(
                 Handle,         { database handle }
                 formula_handle, { selection formula }
                 nil,            { title of view in selection formula }
                 flags,          { search flags }
                 noteClass,      { note class to find }
                 pSince,         { starting date }
                 proc,           { call for each note found }
                 //pointer(self),  { argument to print_fields }
                 pointer(Result),  { argument to print_fields }
                 nil));          { returned ending date (unused) }
    finally
      OSMemFree (formula_handle);
    end;
  except
    Result.free;
    raise;
  end;
end;

//***************************************************
procedure TNotesDatabase.OpenPrivateAddressBook;
var
  BookName: string;
  n: integer;
begin
  BookName := '';
  SetLength (BookName, 255);
  if OsGetEnvironmentString ('NAMES', pchar(BookName), 254) then begin
    BookName := strPas (pchar (BookName));
    n := Pos (',', BookName);
    if n > 0 then System.delete (BookName, n, length(BookName)-n+1);
  end;
  BookName := Trim(BookName);
  if BookName = '' then BookName := 'NAMES.NSF';
  Close;
  Open ('', BookName);
end;

//***************************************************
function TNotesDatabase.GetLocationDocument: TNotesDocument;
var
  sid: string;
  id: NoteID;
  n: integer;
begin
  // Get NoteID
  Result := nil;
  setlength(sid,MAXENVVALUE+1);
  OSGetEnvironmentString (MAIL_LOCATION_ITEM, pchar(sid), MAXENVVALUE);
  sid := StrPas(PChar(sid));
  n := pos(',', sid);
  if n > 0 then System.delete(sid,1,n); //location name
  n := pos(',', sid);
  if n > 0 then System.delete(sid,n,length(sid)); //user name
  try
    id := strToInt('$' + trim(sid));
  except
    exit;
  end;

  // Open document
  Result := TNotesDocument.Create(self, id);
end;

//***************************************************
class function TNotesDatabase.MailType: integer;
var
  db: TNotesDatabase;
  loc: TNotesDocument;
  s: string;
begin
  Result := 0;
  db := TNotesDatabase.Create;
  loc := nil;
  try
    db.OpenPrivateAddressBook;
    loc := db.GetLocationDocument;
    if loc <> nil then begin
      s := loc['SMTPRoute'].AsString;
      if s <> '' then try
        Result := strToInt(s);
        if Result < 0 then Result := 0
        else if Result > 1 then Result := 1;
      except
      end;
    end;
  finally
    loc.free;
    db.free;
  end;
end;

//***************************************************
function TNotesDatabase.OpenView;
begin
  Result := TNotesView.OpenView (self, aName);
end;

//***************************************************
function TNotesDatabase.CreateDocument;
begin
  Result := TNotesDocument.CreateNew(self);
end;

//***************************************************
function TNotesDatabase.CreateResponseDocument;
begin
  Result := TNotesDocument.CreateResponseByUNID(self, aResUNID);
  Result.Subject := aSubject;
end;

//***************************************************
procedure TNotesDatabase.CreateNew;
var
  FilePath: string;
  error: STATUS;
begin
  Close;

  if UpperCase(aServer) = 'LOCAL' then aServer := '';
  if dbFile = '' then raise ELotusNotes.create ('Database Name is blank');

  if aServer = ''
    then FilePath := dbFile
    else FilePath := ConstructPath (aServer, dbFile);

  if TemplateDB <> '' then begin
    Error := NsfDbCreateAndCopy (pchar(Native2Lmbcs(TemplateDB)), pchar(FilePath), NOTE_CLASS_ALL, 0, 0, @FHandle);
  end
  else begin
    FilePath := Native2Lmbcs(FilePath);
    CheckError (NsfDbCreate (pchar(FilePath), DBCLASS_NOTEFILE, True));
    Error := NsfDbOpen (pchar(FilePath), @FHandle);
  end;
  if (Error=USER_CANCEL) or ((Handle=0) and (Error=0)) then raise ELnUserCancel.CreateErr(-1,'User cancelled Notes session');
  CheckError (Error);

  FFileName := dbFile;
  FServerName := aServer;
end;

//***************************************************
function TNotesDatabase.GetDatabaseID;
begin
  FillChar(Result, sizeOf(Result), 0);
  CheckError(NsfDBIdGet(Handle, @Result));
end;

//***************************************************
function TNotesDatabase.GetFullName;
begin
  if (UpperCase(Server) = 'LOCAL') or (Server = '')
    then Result := FileName
    else Result := ConstructPath (Server, FileName);
end;

//***************************************************
function TNotesDatabase.NotesVersion;
begin
  if not Active then Open ('', '');
  CheckError (NSFDbGetBuildVersion (Handle, @Result));
end;

//***************************************************
procedure TNotesDatabase.CopyRecords;
var
  Td: TimeDate;
begin
  TimeConstant(TIMEDATE_WILDCARD, Td);
  CheckError (NSFDbCopy (SourceDb.Handle, Handle, TD, NOTE_CLASS_DOCUMENT));
end;

//***************************************************
procedure TNotesDatabase.Delete;
begin
  CheckError (NSFNoteDelete (Handle, DocId, UPDATE_NOCOMMIT));
end;

//***************************************************
procedure TNotesDatabase.DeleteDocument;
var
  h: integer;
begin
  if Doc.fHandle <> 0 then begin
    h := Doc.DocID;
    NsfNoteClose(Doc.FHandle);
    Doc.fHandle := 0;
    Delete(h);
  end;
end;

//***************************************************
function TNotesDatabase.GetInfo;
var
  Buf: string;
begin
  Buf := '';
  SetLength(Buf, NSF_INFO_SIZE+1);
  CheckError(NSFDbInfoGet (Handle, pchar(Buf)));
  SetLength(Result, 256);
  NSFDbInfoParse(pchar(Buf), Index, pchar(Result), 255);
  Result := Lmbcs2Native(strPas(pchar(Result)));
end;

//***************************************************
procedure TNotesDatabase.SetInfo;
var
  Buf: string;
begin
  Buf := '';
  SetLength (Buf, NSF_INFO_SIZE+1);
  CheckError (NSFDbInfoGet(Handle, pchar(Buf)));
  Value := Native2Lmbcs(Value);
  NSFDbInfoModify(pchar(Buf), Index, pchar(Value));
  CheckError(NSFDbInfoSet (Handle, pchar(Buf)));
end;

//***************************************************
procedure TNotesDatabase.SetFileName(Value: string);
begin
  if Value <> FileName then begin
    Close;
    FFileName := Value;
  end;
end;

//***************************************************
procedure TNotesDatabase.MarkRead;
var
  uName: string;
  hTable, hOriginalTable: THandle;
begin
  // Get IDs
  hOriginalTable := 0;
  uName := LmbcsUserName;
  if not TNotesName.IsCanonical (uName) then uName := TNotesName.TranslateName (uName, True, '');
  CheckError (NSFDbGetUnreadNoteTable(Handle, pchar(UName), length(UName), True, @hTable));
  try
    // Bring table up to date
    CheckError (NSFDbUpdateUnread (Handle, hTable));

    // Notes requires original unread table to merge changes
    CheckError (IDTableCopy (hTable, @hOriginalTable));

    // Delete or insert ID
    if fRead
      then IDDelete (hTable, NoteID, nil)
      else IDInsert (hTable, NoteID, nil);

    // Merge and update table
    CheckError (NSFDbSetUnreadNoteTable (Handle, pchar(UName), length(UName),
      False, hOriginalTable, hTable));
  finally
    if hTable <> 0 then IDDestroyTable(hTable);
    if hOriginalTable <> 0 then IDDestroyTable(hOriginalTable);
  end;
end;

//***************************************************
procedure TNotesDatabase.MarkAllRead;
var
  uName: string;
  hTable, hOriginalTable: THandle;
  i: integer;
begin
  // Get IDs
  hOriginalTable := 0;
  uName := LmbcsUserName;
  if not TNotesName.IsCanonical (uName) then uName := TNotesName.TranslateName (uName, True, '');
  CheckError (NSFDbGetUnreadNoteTable(Handle, pchar(UName), length(UName), True, @hTable));
  try
    // Bring table up to date
    CheckError (NSFDbUpdateUnread (Handle, hTable));

    // Notes requires original unread table to merge changes
    CheckError (IDTableCopy (hTable, @hOriginalTable));

    // Delete or insert IDs
    for i := 0 to Docs.count-1 do begin
      if fRead
        then IDDelete (hTable, Docs.DocumentId[i], nil)
        else IDInsert (hTable, Docs.DocumentId[i], nil);
    end;

    // Merge and update table
    CheckError (NSFDbSetUnreadNoteTable (Handle, pchar(UName), length(UName),
      False, hOriginalTable, hTable));
  finally
    if hTable <> 0 then IDDestroyTable(hTable);
    if hOriginalTable <> 0 then IDDestroyTable(hOriginalTable);
  end;
end;

//***************************************************
procedure TNotesDatabase.SetReplicaInfo;
begin
  CheckError(NsfDbReplicaInfoSet(Handle, @Value));
end;

//***************************************************
function TNotesDatabase.GetReplicaInfo;
begin
  FillChar(Result, sizeOf(Result), 0);
  CheckError(NsfDbReplicaInfoGet(Handle, @Result));
end;

//***************************************************
procedure TNotesDatabase.SetServerName(Value: string);
begin
  if Value <> Server then begin
    Close;
    FServerName := Value;
  end;
end;

//**********************************************
procedure TNotesDatabase.SendMail;
var
  MailDoc: TNotesDocument;
begin
  MailDoc := CreateDocument;
  try
    MailDoc.Form := 'Memo';
    MailDoc.SendTo := Address;
    MailDoc.Subject := Subject;
    MailDoc.BodyAsString := Body;
    MailDoc.Sign;
    MailDoc.Send (False, '');
  finally
    MailDoc.free;
  end;
end;

//***************************************************
function TNotesDatabase.GetQuotaInfo;
begin
  FillChar(Result, sizeOf(Result), 0);
  CheckError(NSFDbQuotaGet(PChar(ConstructPath(Server,
    FileName)), @Result));
end;

//***************************************************
function TNotesDatabase.GetDBActivityInfo: DBACTIVITY;
var
  hDbUserActivity: LHandle;
  wUserCount: WORD;
begin
  FillChar(Result, sizeOf(Result), 0);
  hDbUserActivity := 0;
  CheckError(NSFDbGetUserActivity(Handle, 0, @Result, hDbUserActivity, wUserCount));
  if hDbUserActivity <> 0 then OSMemFree(hDbUserActivity);
end;

//***************************************************
function TNotesDatabase.GetUserActivity: DBUserActivityArray;
var
  dbact: DBACTIVITY;
  pDbUserActivity, dptr: PDBACTIVITY_ENTRY;
  wUserCount: WORD;
  pt: PChar;
  hDbUserActivity, i: integer;
begin
  FillChar(dbact, sizeof(dbact), 0);
  CheckError(NSFDbGetUserActivity(Handle, 0, @dbact,
                           hDbUserActivity, wUserCount));
  try
    pDbUserActivity := OSLockObject(hDbUserActivity);
    setLength(Result, wUserCount);
    for i := 0 to wUserCount -1 do begin
      dptr := pointer(dword(pDbUserActivity) + (sizeOf(DBACTIVITY_ENTRY) * i));
      pt := pointer(dword(pDbUserActivity) + dptr^.UserNameOffset);

      Result[i].Date := NotesToDateTime(dptr^.Time);
      Result[i].Reads := dptr^.Reads;
      Result[i].Writes := dptr^.Writes;
      Result[i].UserName := string(pt);
    end;
    OSUnlockObject(hDbUserActivity);
  finally
    OSMemFree(hDbUserActivity);
  end;
end;


//ClassMarker_Method(TNotesDatabase)

(******************************************************************************)
{ TNotesItem}
(******************************************************************************)
constructor TNotesItem.Create;
begin
  inherited Create;
  fDoc := notesDocument;
  fName:= aName;
  fSaveToDisk := True;
  SetCreated;
end;

(******************************************************************************)
constructor TNotesItem.CreateNext;
begin
  inherited Create;
  fDoc := notesItem.Note;
  fName := notesItem.Name;
  fSaveToDisk := True;
  fCreated := True;
  InitItemInfo(notesItem.ItemBid);
  FSeqNo := notesItem.SeqNo + 1;
end;

(******************************************************************************)
destructor TNotesItem.Destroy;
begin
  FStringsValue.free;
  inherited destroy;
end;

//****************************************************
procedure TNotesItem.GetValueBuffer;
var
  Ptr: pointer;
begin
  Ptr := OSLockBlock(fValueBid);
  try
    GetMem (Buffer, fValueLength);
    Move (Ptr^, Buffer^, fValueLength);
    BufSize := fValueLength;
  finally
    OSUnlockBlock (fValueBid);
  end;
end;

(******************************************************************************)
procedure TNotesItem.InitItemInfo;
var
  retName: string[255];
  retNameLength: word;
  Name2: string;
begin
  Name2 := Native2Lmbcs(Name);
  if (PrevItem.pool = 0) and (PrevItem.block = 0) then begin
    // First item in sequence
    CheckError (NSFItemInfo (Note.Handle,
                             PChar(Name2),
                             Length(Name2),
                             @fItemBId,
                             @fDataType,
                             nil,
                             nil
                             ));
  end
  else begin
    // Next item
    CheckError (NSFItemInfoNext(Note.Handle,
                             PrevItem,
                             PChar(Name2),
                             Length(Name2),
                             @fItemBId,
                             @fDataType,
                             nil,
                             nil
                             ));
  end;
  NSFItemQuery (Note.Handle,
                fItemBID,
                @retName,
                255,
                @retNameLength,
                @fItemFlags,
                @fDataType,
                @fValueBID,
                @fValueLength
  );
end;

//**********************************************
procedure TNotesItem.Refresh;
begin
  if fCreated then InitItemInfo(NullBid);
end;

//**********************************************
procedure TNotesItem.SetAsTimes;
var
  i: integer;
  v: variant;
begin
  if (VarType(Value) and varArray) = 0 then exit;
  if (VarType(Value) and varDate) <> 0 then AsList := Value
  else begin
    v := VarArrayCreate ([VarArrayLowBound(Value,1), VarArrayHighBound(Value,1)], varDate);
    for i := VarArrayLowBound(Value,1) to VarArrayHighBound(Value,1) do
      v[i] := VarAsType (Value[i], varDate);
    AsList := v;
  end;
  SetCreated;
end;

(******************************************************************************)
procedure TNotesItem.SetItemFlags (anItemFlag: integer; anValue: boolean);
begin
  if anValue
    then fItemFlags := fItemflags or anItemFlag
    else fItemFlags := fItemflags xor anItemFlag;
end;

//**********************************************
function TNotesItem.GetRichText;
begin
  if FStringsValue = nil then FStringsValue := TStringList.create;
  if fCreated then FStringsValue.Text := AsString;
  Result := FStringsValue;
end;

//**********************************************
function TNotesItem.GetAsTimes;
begin
  if ItemType = TYPE_TIME_RANGE then Result := AsList else Result := NULL;
end;

(******************************************************************************)
function TNotesItem.GetItemFlags (anItemFlag: integer): boolean;
begin
  result := anItemFlag = (fItemFlags and anItemFlag);
end;

(******************************************************************************)
function TNotesItem.GetLastModifed: TDateTime;
var
  TM: TIMEDATE;
  Name2: string;
begin
  Result := 0;
  if fIsNew then exit;
  Name2 := Native2Lmbcs(Name);
  CheckError (NsfItemGetModifiedTime(Note.Handle,
                                     PChar(Name2),
                                     Length(Name2),
                                     0,
                                     @TM
                                     ));
  Result := NotesToDateTime (TM);
end;

(******************************************************************************)
function TNotesItem.GetAsDateTime: TDateTime;
var
  TM: TIMEDATE;
begin
  Result := 0;
  if not fCreated then exit;
  if NSFItemGetTime(Note.Handle,pchar(Native2Lmbcs(Name)), @TM) then Result := NotesToDateTime (TM);
end;

//**********************************************
function TNotesItem.GetAsFloat;
var
  ldb: double;
begin
  Result := 0;
  if not fCreated then exit;
  if NSFItemGetNumber(Note.Handle,pchar(Native2Lmbcs(Name)),@ldb) then Result := ldb;
end;

//**********************************************
function TNotesItem.GetAsList;
var
  Ptr: pointer;
  PRnge: PRANGE;
  PNumValue: PNUMBER;
  PDtValue: PTIMEDATE;
  i, rCount: integer;
  Str, Name2: string;
begin
  Result := NULL;
  if not fCreated then exit;
  Name2 := Native2Lmbcs(Name);
  case ItemType of
    TYPE_TEXT_LIST: begin
      rCount := NSFItemGetTextListEntries(Note.Handle,pchar(Name2));
      if rcount > 0 then begin
        Result := VarArrayCreate ([0, rCount-1], varOleStr);
        for i := 0 to rCount-1 do begin
          SetLength (Str,2024);
          NSFItemGetTextListEntry (Note.Handle,pchar(Name2),i,pchar (Str),2023);
          Result[i] := Lmbcs2Native(strPas (pchar (Str)));
        end;
      end;
    end;

    TYPE_NUMBER_RANGE: begin
      Ptr := OSLockBlock (fValueBid);
      try
        PRnge := PRANGE (dword(Ptr) + sizeof(WORD));
        if PRnge^.ListEntries > 0 then begin
          Result := VarArrayCreate ([0, PRnge^.ListEntries-1], varDouble);
          PNumValue := PNUMBER (dword(PRnge) + sizeOf(USHORT)*2);
          for i := 0 to PRnge^.ListEntries-1 do begin
            Result[i] := PNumValue^;
            PNumValue := PNUMBER (dword(PNumValue) + sizeof(NUMBER));
          end;
        end;
      finally
        OSUnlockBlock (fValueBid);
      end;
    end;

    TYPE_TIME_RANGE: begin
      Ptr := OSLockBlock (fValueBid);
      try
        PRnge := PRANGE(dword(Ptr) + sizeof(WORD));
        if PRnge^.ListEntries > 0
          then rCount := PRnge^.ListEntries
          else rCount := PRnge^.RangeEntries;
          
        if rCount > 0 then begin
          Result := VarArrayCreate ([0, rCount-1], varDate);
          PDtValue := PTIMEDATE(dword(PRnge) + sizeOf(USHORT)*2);
          for i := 0 to rCount-1 do begin
            Result[i] := VarFromDateTime (NotesToDateTime(PDtValue^));
            PDtValue := PTIMEDATE(dword(PDtValue) + sizeof(TIMEDATE));
          end;
        end;
      finally
        OSUnlockBlock (fValueBid);
      end;
    end;
  end;
end;

//**********************************************
function TNotesItem.GetAsNumbers;
begin
  Result := 0;
  if not fCreated then exit;
  if ItemType = TYPE_NUMBER_RANGE then Result := GetAsList;
end;

//****************************************************
function TNotesItem.GetAsReference;
var
  Buf: pLIST;
  sz: dword;
begin
  FillChar (Result, sizeOf(Result), 0);
  if not fCreated then exit;
  GetValueBuffer (pointer(Buf), sz);
  if Buf = nil then exit;

  // Buffer points to list of references now - but I never saw more than 1 ref
  if Buf^.ListEntries >= 1 then
//    CopyMemory (@Result, pointer(dword(Buf) + ODSLength(_LIST)), sizeOf(UNID));
      CopyMemory(@Result, pointer(dword(Buf) + ODSLength(_LIST) * 2), sizeOf(UNID));
end;

//**********************************************
function TNotesItem.GetAsString;
var
  hBuffer: LHandle;
  retLen: DWORD;
  pBuf: pchar;
begin
  Result := '';
  if not fCreated then exit;
  if ItemType = TYPE_TEXT then begin
    retLen := ValueLength+1;
    if retLen < 1024 then retLen := 1024;
    setLength (Result,retLen);
    retLen := NSFItemGetText (Note.Handle,pchar(Native2Lmbcs(Name)),PChar(Result),retLen);
    Result := NotesToString(pchar(Result),retLen);
  end
  else if ItemType = TYPE_COMPOSITE then begin
    // by Andy
    retLen := 0;
    CheckError(ConvertItemToText(fValueBID,ValueLength,#13#10,65535, @hBuffer, @retLen, false));
    if hBuffer <> 0 then try
      pBuf := OsLockObject(hBuffer);
      if retLen = 0 then Result := ''
      else begin
        setLength(Result, retLen+1);
        strLCopy(pchar(Result), pBuf, retLen);
        (pchar(Result))[retLen] := #0;
        Result := strPas(pchar(Result));
      end;
    finally
      OsUnlockObject(hBuffer);
      OsMemFree (hBuffer);
    end
  end
  else begin
    SetLength (Result,60000);
    NSFItemConvertToText(Note.Handle,
                         PChar(Native2Lmbcs(Name)),
                         PChar(Result),
                         60000,
                         ITEM_VALUE_SEPARATOR
    );
    Result := StrPas(PChar(Result));
  end;
  Result := Lmbcs2Native(Result);
end;

//**********************************************
procedure TNotesItem.SetAsFloat;
var
  v: NUMBER;
begin
  v := Value;
  IsSummary := True;
  CheckError(NSFItemSetNumber(Note.Handle,pchar(Native2Lmbcs(Name)),@v));
  SetCreated;
end;

//**********************************************
procedure TNotesItem.SetAsList;
var
  MemPtr: PRANGE;
  ValPtr: PNUMBER;
  DtPtr: PTIMEDATE;
  ArrSz, ValSz, ListSz: dword;
  i, Lob, Hib: integer;
  str, Name2: string;
  hList: LHANDLE;
  ListPtr: pointer;
begin
  if (VarType(Value) and varArray) = 0 then SetItemValue (Value)
  else begin
    Name2 := Native2Lmbcs(Name);
    NsfItemDelete(Note.Handle, pchar(Name2), length(Name2));
    HiB := VarArrayHighBound (Value, 1);
    LoB := VarArrayLowBound (Value, 1);
    ArrSz := HiB-LoB + 1;
    case (VarType(Value) and varTypeMask) of
      varString, varOleStr: begin
        ListSz := 0;
        CheckError(ListAllocate(0, 0, False, @hList, @ListPtr, @ListSz));
        try
          OsUnlockObject(hList);
          for i := LoB to HiB do begin
            str := Native2Lmbcs(Value[i]);
            CheckError(ListAddEntry(hList,False,@ListSz,i,pchar(str),length(str)));
          end;
          ListPtr := OsLockObject(hList);
          CheckError(NSFItemAppend(Note.Handle,ItemFlags,pchar(Name2),length(Name2),TYPE_TEXT_LIST,
            ListPtr,ListSz));
        finally
          OSUnlockObject(hList);
          OSMemFree(hList);
        end;
        {CheckError (NSFItemCreateTextList (Note.Handle,pchar(Name2), pchar(str), length(str)));
        str := Native2Lmbcs(Value[LoB]);
        for i := Lob+1 to Hib do begin
          str := Native2Lmbcs(Value[i]);
          CheckError (NSFItemAppendTextList (Note.Handle,pchar(Name2), pchar(str),MAXWORD,True));
        end;}
      end;
      varSmallint, varInteger, varSingle, varDouble, varCurrency, varBoolean: begin
        ValSz := 2*sizeof(USHORT) + (ArrSz * sizeOf(NUMBER));
        GetMem (MemPtr, ValSz);
        try
          MemPtr^.ListEntries := ArrSz;
          MemPtr^.RangeEntries := 0;
          ValPtr := PNUMBER (dword(MemPtr) + 2*sizeOf(USHORT));
          for i := Lob to Hib do begin
            ValPtr^ := NUMBER (Value[i]);
            ValPtr := PNUMBER (dword(ValPtr) + sizeOf(NUMBER));
          end;
          CheckError (NSFItemAppend (Note.Handle, ITEM_SUMMARY, pchar(Name2), length(Name2),
            TYPE_NUMBER_RANGE, MemPtr, ValSz));
        finally
          FreeMem (MemPtr);
        end;
      end;
      varDate: begin
        ValSz := 2*sizeof(USHORT) + (ArrSz * sizeOf(TIMEDATE));
        GetMem (MemPtr, ValSz);
        try
          MemPtr^.ListEntries := ArrSz;
          MemPtr^.RangeEntries := 0;
          DtPtr := PTIMEDATE (dword(MemPtr) + 2*sizeOf(USHORT));
          for i := Lob to Hib do begin
            DtPtr^ := DateTimeToNotes(VarToDateTime (Value[i]));
            DtPtr := PTIMEDATE (dword(DtPtr) + sizeOf(TIMEDATE));
          end;
          CheckError (NSFItemAppend (Note.Handle, ITEM_SUMMARY, pchar(Name2), length(Name2),
            TYPE_TIME_RANGE, MemPtr, ValSz));
        finally
          FreeMem (MemPtr);
        end;
      end;
    end;
  end;
  SetCreated;
end;

//**********************************************
procedure TNotesItem.SetAsNumbers;
var
  i: integer;
  v: variant;
begin
  if (VarType(Value) and varArray) = 0 then exit;
  if (VarType(Value) and varDouble) <> 0 then AsList := Value
  else begin
    v := VarArrayCreate ([VarArrayLowBound(Value,1), VarArrayHighBound(Value,1)], varDouble);
    for i := VarArrayLowBound(Value,1) to VarArrayHighBound(Value,1) do
      v[i] := VarAsType (Value[i], varDouble);
    AsList := v;
  end;
  SetCreated;
end;

//****************************************************
procedure TNotesItem.SetAsReference;
var
  sz: dword;
  buf: pLIST;
begin
  sz := ODSLength(_LIST) + sizeOf(UNID);
  GetMem (Buf, sz);
  try
    Buf^.ListEntries := 1;
    CopyMemory (pointer(dword(Buf) + ODSLength(_LIST)), @Value, sizeOf(Value));
    SetValueBuffer (TYPE_NOTEREF_LIST, ItemFlags, Buf, sz);
  finally
    FreeMem(buf);
  end;
end;

//**********************************************
procedure TNotesItem.SetAsString;
var
  buf: pchar;
  bufsz: integer;
begin
  Value := Native2Lmbcs(Value);
  StringToNotes(Value, buf, bufsz);
  try
    dec(bufsz);   //trailing zero is not counted in
    SetValueBuffer(TYPE_TEXT, ItemFlags, buf, bufsz);
    //SetValueBuffer(TYPE_TEXT, ItemFlags, pchar(Value), length(Value));
    //CheckError(NSFItemSetText(Note.Handle,pchar(Native2Lmbcs(Name)),PChar(Value),length(Value)));
    SetCreated;
    FValueLength := bufsz;
  finally
    FreeMem(buf);
  end;
end;

(******************************************************************************)
procedure TNotesItem.SetAsString2(const Value: string);
var
  name2: string;
begin
  name2 := Native2Lmbcs(name);
  CheckError(NSFItemSetText(Note.Handle, pchar(name2), pchar(Value), MAXWORD));
  SetCreated;
  ItemType := TYPE_TEXT;
  FValueLength := length(Value);
end;

(******************************************************************************)
procedure TNotesItem.SetAsDateTime (Value: TDateTime);
var
  T: TIMEDATE;
begin
  if value = 0 then AsString := ''
  else begin
    T := DateTimeToNotes (Value);
    CheckError(NSFItemSetTime(Note.Handle, pchar(Native2Lmbcs(Name)), @T));
    SetCreated;
    ItemType := TYPE_TIME;
  end;
end;

(******************************************************************************)
function TNotesItem.GetItemValue: variant;
var
  str: string;
begin
  Result := NULL;
  if not fCreated then exit;
  case ItemType of
    TYPE_TEXT:          Result := AsString;
    TYPE_TEXT_LIST:     Result := AsList;
    TYPE_NUMBER:        Result := AsNumber;
    TYPE_NUMBER_RANGE:  Result := AsNumbers;
    TYPE_TIME:          Result := VarFromDateTime (AsDateTime);
    TYPE_TIME_RANGE:    Result := AsTimes;
    TYPE_NOTEREF_LIST:  Result := UNIDToStr(AsReference,True);
    else begin
      SetLength (str, 62001);
      NSFItemConvertValueToText (ItemType, fValueBid, ValueLength, pchar(str), 62000, #0);
      Result := Lmbcs2Native(strPas(pchar(str)));
    end;
  end;
end;

(******************************************************************************)
procedure TNotesItem.SetItemValue (aValue: variant);
begin
  if (VarType(aValue) and varArray) <> 0 then AsList := aValue
  else case (VarType(aValue) and varTypeMask) of
    varEmpty, varNull:   ;
    varSmallint, varInteger:
      AsNumber := int(integer(aValue));
    varSingle, varDouble, varCurrency:
      AsNumber := extended(aValue);
    varDate:
      AsDateTime := VarToDateTime (aValue);
    varBoolean:
      AsNumber := int (ord(boolean (aValue)));
    varString, varOleStr:
      AsString := aValue;
  end;
  SetCreated;
end;

//**********************************************
constructor TNotesItem.CreateNew;
begin
  inherited Create;
  fDoc := notesDocument;
  fName := aName;
  fSaveToDisk := true;
  IsSummary := True;
  fCreated := False;
end;

//**********************************************
procedure TNotesItem.SetCreated;
begin
  if not fCreated then begin
    InitItemInfo(NullBid);
    fCreated := True;
  end;
end;

//**********************************************
function TNotesItem.GetAsStrings;
var
  Val: variant;
  i, retLen: integer;
  s: string;
begin
  if FStringsValue = nil then FStringsValue := TStringList.create;
  if fCreated and ((ItemType = TYPE_TEXT_LIST) or (ItemType = TYPE_NUMBER_RANGE)
  or (ItemType = TYPE_TIME_RANGE)) then begin
    Val := AsList;
    for i := VarArrayLowBound (Val,1) to VarArrayHighBound(Val,1) do FStringsValue.add (string(Val[i]));
  end
  // by Kristjan Bjarni Gudmundsson - handling TYPE_TEXT
  else if fCreated and (ItemType = TYPE_TEXT) then begin
    setLength (s,ValueLength+1);
    retLen := NSFItemGetText(Note.Handle,pchar(Native2Lmbcs(Name)),PChar(s),ValueLength);
    s := NotesToString(pchar(s),retLen);
    FStringsValue.Add(Lmbcs2Native(s));
  end;
  Result := FStringsValue;
end;

//**********************************************
procedure TNotesItem.SetAsStrings;
var
  Val: variant;
  i: integer;
begin
  if Value = nil then exit;
  if FStringsValue = nil then FStringsValue := TStringList.create;
  FStringsValue.Assign (Value);
  if FStringsValue.count = 0 then AsString := ''
  else begin
    Val := VarArrayCreate([0, FStringsValue.count-1], varOLEStr);
    for i := 0 to FStringsValue.count-1 do Val[i] := FStringsValue[i];
    AsList := Val;
    ItemType := TYPE_TEXT_LIST;
  end;
end;

//**********************************************
procedure TNotesItem.SetRichText;
var
  cHandle: THandle;
  Name2, buf: string;
begin
  buf := Native2Lmbcs(Value.text);
  Name2 := Native2Lmbcs(Name);
  NsfItemDelete(Note.Handle, pchar(Name2), length(Name2));
  CheckError (CompoundTextCreate(Note.Handle, pchar(Name2), @cHandle));
  try
    CheckError (CompoundTextAddText (cHandle, STYLE_ID_SAMEASPREV, Default_Font_ID,
      pchar(buf), length(buf), #13#10, COMP_PRESERVE_LINES, NULLHANDLE));
  finally
    CompoundTextClose (cHandle, nil, nil, nil, 0);
  end;
end;

//****************************************************
procedure TNotesItem.SetValueBuffer;
var
  Name2: string;
begin
  Name2 := pchar(Native2Lmbcs(Name));
  NsfItemDelete(Note.Handle, pchar(Name2), length(Name2));
  HugeNsfItemAppend(Note.Handle,wFlags,pchar(Name2),length(Name2),
    iType,Buffer,BufSize);
  ItemType := iType;
end;

(******************************************************************************)
procedure TNotesItem.SetAsPChar(buffer: pchar);
begin
  if buffer <> nil then begin
    SetValueBuffer(TYPE_TEXT, ItemFlags, buffer, strlen(buffer));
    SetCreated;
  end;
end;

(******************************************************************************)
function TNotesItem.CreateNextItem: TNotesItem;
var
  cls: TNotesItemClass;
begin
  if not NextItemExists then Result := nil
  else begin
    cls := TNotesItemClass(self.classType);
    Result := cls.createNext(self);
  end;
end;

(******************************************************************************)
function TNotesItem.NextItemExists: boolean;
var
  Name2: string;
  tDataType: WORD;
  tItemBid: BLOCKID;
begin
  Name2 := Native2Lmbcs(Name);
  Result := NSFItemInfoNext(Note.Handle,
                             ItemBid,
                             PChar(Name2),
                             Length(Name2),
                             @tItemBid,
                             @tDataType,
                             nil,
                             nil
                             ) = 0;
end;

(******************************************************************************)
function TNotesItem.LoadNextItem: boolean;
begin
  Result := NextItemExists;
  if Result then begin
    // Trying to load next item info
    InitItemInfo(ItemBid);
    inc(FSeqNo);
    FStringsValue.free;
    FStringsValue := nil;
  end;
end;

(******************************************************************************)
function TNotesItem.GetDoc: TNotesDocument;
begin
  if fDoc is TNotesDocument
    then Result := TNotesDocument(fDoc)
    else Result := nil;
end;


//ClassMarker_Method(TNotesItem)

(******************************************************************************)
{ TNotesNote }
(******************************************************************************)
constructor TNotesNote.CreateByUNID;
var
  id: dword;
begin
  inherited Create;
  CheckError(NSFNoteOpenByUNID(notesDatabase.Handle, @anUNID, 0, @FHandle));
  FKeepHandle := false;
  NSFNoteGetInfo(FHandle, _NOTE_ID, @id);
  InitDocument(NotesDatabase, id);
end;

//****************************************************
constructor TNotesNote.CreateEmpty;
begin
  inherited Create;
  FDatabase := notesDatabase;
  FKeepHandle := false;
end;

//****************************************************
constructor TNotesNote.CreateNew;
begin
  inherited Create;
  FDatabase := notesDatabase;
  FKeepHandle := false;
  CheckError (NsfNoteCreate(NotesDatabase.Handle, @FHandle));
  CreateDocument(notesDatabase);
end;

(******************************************************************************)
constructor TNotesNote.Create;
var
  err: DWORD;
begin
  inherited Create;
  FKeepHandle := false;
  err := NSFNoteOpen(notesDatabase.Handle, anId, 0, @FHandle);
  // Added by Matt Saint - handling situations where a note
  // been opened from collection has already been deleted
  if ((err and $0225) <> NOERROR) then FIsDeleted := true
  else begin
    CheckError(err);
    InitDocument(NotesDatabase, anID);
  end;
end;

(******************************************************************************)
constructor TNotesNote.CreateFromHandle(notesDatabase: TNotesDatabase; aHandle: LHandle);
var
  id: NOTEID;
begin
  inherited Create;
  FHandle := aHandle;
  FKeepHandle := true;
  NSFNoteGetInfo(Handle, _NOTE_ID, @id);
  InitDocument(notesDatabase, id);
end;

(******************************************************************************)
destructor TNotesNote.Destroy;
begin
  if (not FKeepHandle) and (FHandle <> 0) then NsfNoteClose(FHandle);
  FFields.free;
  inherited Destroy;
end;

(******************************************************************************)
procedure TNotesNote.CopyItem;
var
  itSource, itTarget: TNotesItem;
  buf, bufptr: pointer;
  bufsz: dword;
begin
  if IsItemExists(itemName) then DeleteItem(itemName);
  itSource := Source.Items[itemName];
  if itSource.ItemType = TYPE_COMPOSITE
    then itTarget := TNotesRichTextItem.CreateNew(self, itemName)
    else itTarget := TNotesItem.CreateNew (self, itemName);
  itSource.GetValueBuffer(buf, bufsz);
  if buf <> nil then try
    bufptr := pointer(dword(buf) + sizeof(WORD));
    itTarget.SetValueBuffer(itSource.ItemType, itSource.ItemFlags, bufptr, bufsz - sizeOf(WORD));
  finally
    FreeMem(buf);
  end;
end;

(******************************************************************************)
function TNotesNote.CopyToDatabase;
var
  Ddid:NoteId;
begin
  CheckError (NSFDbCopyNote(Database.Handle,nil,nil,DocId,DestDB.Handle,nil,nil,@Ddid,nil));
  Result := TNotesDocument.Create(DestDB,ddid);
end;

//***************************************************
procedure TNotesNote.ReloadFields;
var
  TempList: TStringList;
begin
  TempList := TStringList.create;
  try
    CheckError(NSFItemScan(Handle, FieldsScanProc, TempList));
    FFields.Assign(TempList);
  finally
    TempList.Free;
  end;
end;

(******************************************************************************)
function TNotesNote.ReplaceItemValue;
begin
  Result := nil;
  try
    DeleteItem(ItemName);
  except
  end;
  Items[ItemName].Value := Value;
end;

(******************************************************************************)
procedure TNotesNote.Save;
begin
  CheckError(NSFNoteUpdate(Handle,UPDATE_NOCOMMIT));
  if DocId = 0 then NSFNoteGetInfo(Handle, _NOTE_ID, @FId);
end;

//**********************************************
function TNotesNote.GetFieldCount;
begin
  Result := FFields.count;
end;

//**********************************************
function TNotesNote.GetFieldName;
begin
  Result := FFields[Index];
end;

//**********************************************
function TNotesNote.GetItemByName;
var
  Index: integer;
  name2: string;
begin
  Index := FFields.indexOf (ItemName);
  if Index < 0 then begin
    Result := TNotesItem.createNew (self, ItemName);
    FFields.addObject (ItemName, Result);
  end
  else begin
    if FFields.Objects[Index] = nil then begin
      name2 := Native2Lmbcs(ItemName);
      if NSFItemIsPresent (Handle, pchar(name2), length(name2)) then
        FFields.Objects[Index] := TNotesItem.create(self, ItemName)
      else begin
        FFields.Objects[Index] := TNotesItem.createNew(self, ItemName);
      end;
    end;
    Result := FFields.Objects[Index] as TNotesItem;
  end;
end;

//**********************************************
procedure TNotesNote.DeleteItem;
var
  Index: integer;
  name2: string;
begin
  name2 := Native2Lmbcs(ItemName);
  while NsfItemDelete (Handle, pchar(name2), length(name2)) = 0 do;
  Index := FFields.indexOf (ItemName);
  if Index >= 0 then begin
    FFields.Objects[Index].free;
    FFields.delete (Index);
  end;
end;

//****************************************************
function TNotesNote.GetOriginatorID;
var
  td: TIMEDATE;
  cls: word;
begin
  CheckError(NSFDbGetNoteInfo(Database.Handle, DocID, @Result, @td, @cls));
end;

//****************************************************
function TNotesNote.GetUniversalID;
var
  id: OID;
begin
  id := GetOriginatorID;
  Result.aFile := id.FileNum;
  Result.Note := id.Note;
end;

(******************************************************************************)
function TNotesNote.GetCreated: TDateTime;
var
  td: TIMEDATE;
  cls: word;
  OID:OriginatiorID;
begin
  CheckError(NSFDbGetNoteInfo(database.handle, fID, @OID, @td, @cls));
  result:=NotesToDateTime(OID.Note);
end;

(******************************************************************************)
function TNotesNote.GetSize: longint;
var
  v: variant;
begin
  Result := 0;
  if Handle <> 0 then begin
    v := Evaluate('@DocLength');
    if not VarIsEmpty(v) and not VarIsNull(v) then try
      Result := VarAsType(v, varInteger);
    except
    end;
  end;
end;

(******************************************************************************)
procedure TNotesNote.CreateDocument(notesDatabase: TNotesDatabase);
begin
  FFields := TStringList.Create;
end;

//***************************************************
procedure TNotesNote.InitDocument;
begin
  FFields := TStringList.create;
  FDatabase := notesDatabase;
  fId := anId;
  ReloadFields;
end;

//****************************************************
function TNotesNote.IsItemExists;
var
  name2: string;
begin
  name2 := Native2Lmbcs(ItemName);
  Result := NSFItemIsPresent (Handle, pchar(Name2), length(Name2));
end;

//***************************************************
function TNotesNote.GetLastModified: TDateTime;
var
  T: TIMEDATE;
begin
  NsfNoteGetInfo (Handle, _NOTE_MODIFIED, @T);
  Result := NotesToDateTime (T);
end;

//***************************************************
function TNotesNote.GetLastAccessed: TDateTime;
var
  T: TIMEDATE;
begin
  NsfNoteGetInfo (Handle, _NOTE_ACCESSED, @T);
  Result := NotesToDateTime (T);
end;

//****************************************************
function TNotesNote.Evaluate;
var
  hFCompiled: FormulaHandle;
  hFComputed: HCOMPUTE;
  wFormulaErr: STATUS;
  wFormulaLen, wFormulaErrLine,
  wFormulaErrCol, wFormulaErrOffset, wFormulaErrLen: word;
  hResult: LHANDLE;
  wResultLen: word;
  fNoteMatchesFormula, fNoteToDelete, fNoteModified: boolean;
  phFCompiled, phResult: pointer;
begin
  // Formula compilation
  hFCompiled := 0;
  try
    CheckError (NSFFormulaCompile(nil,0,pchar(aFormula),length(aFormula),@hFCompiled,
      @wFormulaLen, @wFormulaErr, @wFormulaErrLine, @wFormulaErrCol, @wFormulaErrOffset,
      @wFormulaErrLen));
  except
    on E: ELotusNotes do begin
      raise ELnFormulaCompile.CreateErr (E.ErrorCode, E.Message,
        wFormulaErr, wFormulaErrOffset, wFormulaErrLen);
    end
    else raise;
  end;

  // Evaluation
  hFComputed := 0;
  hResult := 0;
  try
    phFCompiled := OSLockObject (hFCompiled);
    CheckError (NSFComputeStart (0, phFCompiled, @hFComputed));
    CheckError (NSFComputeEvaluate (hFComputed, self.Handle, @hResult, @wResultLen,
      @fNoteMatchesFormula, @fNoteToDelete, @fNoteModified));
    phResult := OsLockObject (hResult);
    Result := BufferToValue(phResult, wResultLen);
  finally
    if hResult <> 0 then begin
      OsUnlockObject (hResult);
      OsMemFree (hResult);
    end;
    if hFCompiled <> 0 then begin
      OsUnlockObject (hFCompiled);
      OsMemFree (hFCompiled);
    end;
    if hFComputed <> 0 then NSFComputeStop (hFComputed);
  end;
end;

//**********************************************
procedure TNotesNote.Sign;
begin
  CheckError(NsfNoteSign(Handle));
end;

//**********************************************
function TNotesNote.GetSignature;
var
  d: TIMEDATE;
begin
  SetLength (SignedBy, MAXUSERNAME+1);
  SetLength (CertifiedBy, MAXUSERNAME+1);
  Result := NsfNoteVerifySignature (Handle, nil, @d, pchar(SignedBy), pchar(CertifiedBy)) = NOERROR;
  if Result then begin
    SignedBy := strPas (pchar(SignedBy));
    CertifiedBy := strPas (pchar(CertifiedBy));
    if pTime <> nil then pTime^ := NotesToDateTime (d);
  end;
end;

//***************************************************
function TNotesNote.CountMultipleItems(aName: string): integer;
var
  item_new, item_cur: TNotesItem;
  first_item: boolean;
begin
  Result := 0;
  first_item := True;
  item_cur := Items[aName];
  if not item_cur.IsNewItem then exit;
  inc(Result);

  while item_cur.NextItemExists do begin
    try
      item_new := item_cur.CreateNextItem;
    finally
      if not first_item then item_cur.free;
    end;
    inc(Result);
    item_cur := item_new;
    first_item := False;
  end;
end;

//***************************************************
function TNotesNote.LoadMultipleItems(aName: string): TList;
var
  item_new, item_cur: TNotesItem;
begin
  Result := TList.create;
  try
    item_cur := TNotesItem.Create(self, aName);
    Result.add(item_cur);

    while item_cur.NextItemExists do begin
      item_new := item_cur.CreateNextItem;
      Result.add(item_new);
      item_cur := item_new;
    end;
  except
    Result.free;
    raise;
  end;
end;


(******************************************************************************)
{ TNotesDocument }
(******************************************************************************)
destructor TNotesDocument.Destroy;
var
  i: integer;
begin
  FFontTable.free;
  FSummary.free;
  if FAttach <> nil then for i := 0 to FAttach.count-1 do FAttach.Objects[i].free;
  if FFields <> nil then for i := 0 to FFields.count-1 do FFields.Objects[i].Free;
  FAttach.free;
  inherited;
end;

(******************************************************************************)
procedure TNotesDocument.AttachForm;
var
  Forms: TNotesDocumentCollection;
  FormDoc: TNotesDocument;
  i: integer;
  itname: string;
begin
  // Find form note document
  FormDoc := nil;
  Forms := Database.FindNotes('$TITLE="' + aForm + '"', 0, NOTE_CLASS_FORM, False);
  if Forms.count = 0 then
    raise ELotusNotes.CreateErr(-1, 'Form ' + aForm + ' was not found');
  try
    // Copy form items to this document - $Body, $Info, $Title, $Script...
    FormDoc := Forms.Document[0];
    for i := 0 to FormDoc.FieldCount-1 do begin
      itname := upperCase(FormDoc.FieldName[i]);
      //if (itname = '$BODY') or (itname = '$INFO') or (itname = '$TITLE') or
      //(pos (itname, '$$Script') = 1)
      if (itname <> '') and (itname[1] = '$') and (itname <> '$FIELDS') and
      (itname <> '$SIGNATURE') and (itname <> '$UPDATEDBY') and (itname <> '$FLAGS') and
      (itname <> '$REVISIONS') and (pos('COPYTO',itname) = 0) and (pos('SENDTO',itname) = 0)
      then begin
        CopyItem (FormDoc, FormDoc.FieldName[i]);
      end;
    end;
    if IsItemExists('Form') then deleteItem('Form');
  finally
    FormDoc.free;
    Forms.free;
  end;
end;

(******************************************************************************)
function TNotesDocument.ComputeWithForm;
var
  res : STATUS;
begin
  res := NSFNoteComputeWithForm(Handle,0,0,nil,nil);
  if raiseerror then CheckError(res);
  result := (res = 0)
end;

(******************************************************************************)
procedure TNotesDocument.CheckAddress;

procedure DoCheck (Address: ansistring);
var
  n: integer;
  s: string;
begin
  while Address <> '' do begin
    n := Pos (',', Address);
    if n = 0 then n := length(Address) + 1;
    s := Native2Lmbcs(trim(copy (Address, 1, n-1)));
    delete (Address, 1, n);

    if not TNotesName.CheckAddress (Database.MailServer, s) then
      raise ELotusNotes.createErr (-1, 'Cannot find address "' + s + '" on server ' + Database.MailServer);
  end;
end;

begin
  DoCheck (SendTo);
  DoCheck (Recipients);
end;

(******************************************************************************)
function TNotesDocument.GetBodyAsString;
begin
  Result := Items['Body'].AsString;
end;

(******************************************************************************)
procedure TNotesDocument.SetBodyAsString;
var
  dwBodyItemLen: DWORD;
  hBodyItem: integer;
  i:Integer;
  TempList:TStringList;
begin
  if Value <> '' then begin
    //Value := Native2Lmbcs(Value);
    // MailAddBodyItem uses this function automatically - Fujio Kurose
    CheckError (MailCreateBodyItem(@hBodyItem, @dwBodyItemLen));
    TempList:=TStringList.Create;
    TempList.Text:=Value;
    try
      for i:=0 to TempList.Count-1 do begin
        CheckError (MailAppendBodyItemLine(hBodyItem, @dwBodyItemLen,
                    pchar(TempList[i]), word(length(TempList[i]))));
      end;
      CheckError (MailAddBodyItem(Handle, hBodyItem, dwBodyItemLen, nil));
    finally
      OsMemFree (hBodyItem);
      TempList.Free;
    end;
  end;
end;

(******************************************************************************)
function TNotesDocument.GetBodyAsMemo;
begin
  Result := Items['Body'].AsStrings;
end;

(******************************************************************************)
procedure TNotesDocument.SetBodyAsMemo;
var
  wBodyCount: word;
  hBodyItem: integer;
  dwBodyItemLen: dword;
  itemName: string;
begin
  itemName := 'Body';
  if Value.Count > 0 then begin
    CheckError (MailCreateBodyItem (@hBodyItem, @dwBodyItemLen));
    try
      for wBodyCount := 0 to Value.Count-1 do begin
        CheckError (MailAppendBodyItemLine (hBodyItem,@dwBodyItemLen,pchar(Value[wBodyCount]),length(Value[wBodyCount])));
      end;
      CheckError (MailAddBodyItem (Handle,hBodyItem,dwBodyItemLen, nil));
    finally
      OsMemFree (hBodyItem);
    end;
  end;
end;

(******************************************************************************)
function TNotesDocument.GetItemByNum;
var
  error: integer;
  fldLength: word;
begin
  Result := '';
  setLength(Result,MAX_STR_LEN);
  error := MailGetMessageItem (Handle, ItemNum, pchar(Result), MAX_STR_LEN, @fldLength);
  Result := strPas (pchar (Result));
  if error<>0 then Result := '' else Result := Lmbcs2Native(Result);
  if error <> 546 then CheckError(error); //546 - item not found
end;

(******************************************************************************)
procedure TNotesDocument.SetItemByNum;
begin
  Value := Native2Lmbcs(Value);
  CheckError(MailAddHeaderItem(Handle, ItemNum, pchar(Value), WORD(length(Value))));
end;

(******************************************************************************)
function TNotesDocument.GetRecipients;
var
  error: integer;
  fldLength: word;
begin
  Result := '';
  setLength(Result,MAX_STR_LEN);
  error := MailGetMessageItem(Handle, MAIL_RECIPIENTS_ITEM_NUM, pchar(Result), MAX_STR_LEN,@fldLength);
  Result := StrPas(PChar(Result));
  if error<>0 then Result := '';
end;

(******************************************************************************)
procedure TNotesDocument.SetRecipients;
var
  hRecipientsList: integer;
  plistRecipients: ptrLIST;
  wRecipientsSize: word;
  n: integer;
  s: string;
begin
  if Value = '' then exit;
  CheckError (ListAllocate(0,0,True,@hRecipientsList,@plistRecipients,@wRecipientsSize));
  try
    OSUnlockObject(hRecipientsList);
    n := Pos (',', Value);
    while n > 0 do begin
      s := Native2Lmbcs(Trim (copy (Value, 1, n-1)));
      Delete (Value, 1, n);
      n := Pos (',', Value);
      CheckError (ListAddEntry(hRecipientsList, True, @wRecipientsSize, 0,
        pchar(s),WORD(length(s))));
    end;
    Value:=Native2Lmbcs(Value);
    if Value <> '' then CheckError (ListAddEntry(hRecipientsList, True, @wRecipientsSize, 0,
      pchar(Value),WORD(length(Value))));
    CheckError (MailAddRecipientsItem (Handle, hRecipientsList, wRecipientsSize));
  except
    OsMemFree (hRecipientsList);
    raise;
  end;
end;

(******************************************************************************)
procedure TNotesDocument.Save;
var
  pname, uname: string;
begin
  if ProfileName = '' then inherited Save(force, createResponse, markRead)
  else begin
    // Profile document
    pname := Native2LMBCS(ProfileName);
    uname := Database.LmbcsUserName;
    if not TNotesName.IsCanonical (uName) then uName := TNotesName.TranslateName (uName, True, '');
    CheckError(NSFProfileUpdate(Handle,pchar(pname),length(pname),
      pchar(uname),length(uname)));
  end
end;

(******************************************************************************)
procedure TNotesDocument.Send;
var
  hMailBox, hOrigDB: integer;
  sMailbox, sServer: string;
  szMailBoxPath: array [0..256] of char;
  OrigNoteID:longint;
  OrigNoteOID, NewNoteOID: OID;
  tdDate: TIMEDATE;
  sSendTo, sCopyTo, sBCopyTo, sRecip: string;
begin
  if fAttachForm then begin
    if Form = '' then attachForm(MAIL_MEMO_FORM) else attachForm(Form);
  end
  else if Form = '' then Form := MAIL_MEMO_FORM;
  if MailFrom = '' then MailFrom := Database.UserName;
  if aRecipients <> '' then Recipients := aRecipients
  else if Recipients = '' then begin
    // Construct recipients list
    sSendTo := SendTo;
    sCopyTo := CopyTo;
    sBCopyTo := BlindCopyTo;
    if sSendTo = '' then sSendTo := Items['SendTo'].AsString;
    if sCopyTo = '' then sCopyTo := Items['CopyTo'].AsString;
    if sBCopyTo = '' then sBCopyTo := Items['BlindCopyTo'].AsString;
    sRecip := sSendTo;
    if sCopyTo <> '' then sRecip := sRecip+',' + sCopyTo;
    if sBCopyTo <> '' then sRecip := sRecip+',' + sBCopyTo;
    Recipients := sRecip;
  end;

  OSCurrentTIMEDATE(@tdDate);
  CheckError (MailAddHeaderItem(Handle, MAIL_POSTEDDATE_ITEM_NUM,
    PChar(@tdDate), WORD(sizeof(TIMEDATE))));
  Items['Principal'].asString := MailFrom;

  // Open mailbox
  hMailBox := 0;
  if Database.MailType = 1 then begin
    // SMTP mail - only local databases
    sMailbox := SMTPBOX_NAME;
    sServer := '';
  end
  else begin
    sMailbox := MAILBOX_NAME;
    sServer := Database.MailServer;
  end;
  CheckError (OSPathNetConstruct('', pchar(sServer), pchar(sMailbox), szMailBoxPath));
  CheckError (NSFDbOpen(szMailBoxPath, @hMailBox));

  if (hMailBox=0) then raise ELnUserCancel.CreateErr(-1,'User cancelled Notes session');
  try
    // Take note params
    NSFNoteGetInfo(Handle, _NOTE_DB,  @hOrigDB);
    NSFNoteGetInfo(Handle, _NOTE_ID,  @OrigNoteID);
    NSFNoteGetInfo(Handle, _NOTE_OID, @OrigNoteOID);

    // Set the message's OID database ID to match the mail box */}
    CheckError (NSFDbGenerateOID (hMailBox, @NewNoteOID));
    NSFNoteSetInfo(Handle, _NOTE_DB,  @hMailBox);
    NSFNoteSetInfo(Handle, _NOTE_ID, nil{0});
    NSFNoteSetInfo(Handle, _NOTE_OID, @NewNoteOID);

    // Update message into MAIL.BOX on mail server.}
    CheckError (NSFNoteUpdate(Handle, UPDATE_NOCOMMIT));

    // restore msg to user's mail file and Update to save it there.*/}
    if Database.SaveMail then begin
      NSFNoteSetInfo(Handle, _NOTE_DB, @hOrigDB);
      NSFNoteSetInfo(Handle, _NOTE_ID, @OrigNoteID);
      NSFNoteSetInfo(Handle, _NOTE_OID, @OrigNoteOID);
      CheckError(NSFNoteUpdate(Handle, UPDATE_NOCOMMIT));
    end;
  finally
    NSFDbClose (hMailBox);
  end;
end;

(******************************************************************************)
function TNotesDocument.Attach;
begin
  if not fileExists(FileName) then raise ELotusNotes.Create('File to be attached does not exist: ' + FileName);
  if DisplayName = '' then DisplayName := extractFileName(FileName);
  CheckError(NSFNoteAttachFile(Handle, ITEM_NAME_ATTACHMENT, length(ITEM_NAME_ATTACHMENT),
    pchar(Native2Lmbcs(FileName)), pchar(Native2Lmbcs(DisplayName)), COMPRESS_HUFF));
  Result := 0;
end;

(******************************************************************************)
procedure TNotesDocument.Detach;
var
  SourceName: string;
  atcItemcl: BLOCKIDcl;
begin
  SourceName := FAttach[Index];
  atcItemcl := BLOCKIDcl(FAttach.Objects[Index]);
  if FileName = '' then FileName := SourceName;
  CheckError(NSFNoteExtractFile(Handle, atcItemcl.BlockItem, pchar(Native2Lmbcs(FileName)), nil));
end;

//****************************************************
constructor TNotesDocument.CreateProfile;
var
  anID: NOTEID;
  uname, pname: string;
  puname: pchar;
  luname: integer;
begin
  CreateEmpty(notesDatabase);
  pname := Native2LMBCS(aProfileName);
  if uname = '' then begin
    puname := nil;
    luname := 0;
  end
  else begin
    uname := notesDatabase.LmbcsUserName;
    if not TNotesName.IsCanonical (uName) then uName := TNotesName.TranslateName(uName, True, '');
    puname := pchar(uname);
    luname := length(uname);
  end;
  CheckError(NSFProfileOpen(notesDatabase.Handle,pchar(pname),length(pname),
    puname,luname, TRUE, @FHandle));
  NSFNoteGetInfo(FHandle, _NOTE_ID, @anID);
  FProfileName := aProfileName;
  InitDocument(NotesDatabase, anID);
end;

//****************************************************
constructor TNotesDocument.CreateResponse;
begin
  CreateNew(notesDatabase);
  Items['$REF'].AsReference := MainDoc.UniversalID;
  if MainDoc.IsItemExists('Subject') then Subject := 'Re: ' + MainDoc.Subject;
end;

//**********************************************
constructor TNotesDocument.CreateResponseByUNID;
begin
  CreateNew (notesDatabase);
  Items['$REF'].AsReference := anUNID;
end;

//***************************************************
procedure TNotesDocument.InitDocument;
var
  atcNum: integer;
  atcItem: BLOCKID;
  atcItemcl: BLOCKIDcl;
  atcFileName: string;
begin
  inherited InitDocument(notesDatabase,anID);
  FAttach := TStringList.Create;
  FSummary := TStringList.create;

  // Load attachments list
  atcNum := 0;
  while True do begin
    setLength (atcFileName, 256);
    if not MailGetMessageAttachmentInfo(Handle, atcNum, @atcItem, pchar(atcFileName),
      nil,nil,nil,nil,nil) then break;
    atcFileName := strPas(pchar(atcFileName));
    if atcFileName <> '' then begin
      atcItemcl := BLOCKIDcl.Create;
      atcItemcl.BlockItem := atcItem;
      FAttach.AddObject(AnsiUpperCase(Lmbcs2Native(atcFileName)), atcItemcl);
      inc(atcNum);
    end;
  end;
end;

//****************************************************
procedure TNotesDocument.CreateDocument;
var
  tdDate: TIMEDATE;
begin
  inherited CreateDocument(notesDatabase);
  FFields := TStringList.create;
  FAttach := TStringList.Create;

  // Set "ComposedDate" to the current time/date right now
  OSCurrentTIMEDATE(@tdDate);
  CheckError (MailAddHeaderItem(Handle, MAIL_COMPOSEDDATE_ITEM_NUM,
    PChar(@tdDate), WORD(sizeof(TIMEDATE))));
end;

//***************************************************
function TNotesDocument.GetAttachmentCount;
begin
  Result := FAttach.count;
end;

//***************************************************
function TNotesDocument.GetAttachment;
begin
  Result := FAttach[Index];
end;

//***************************************************
function TNotesDocument.FindAttachment;
begin
  Result := FAttach.indexOf(AnsiUpperCase(aName));
end;

//***************************************************
procedure TNotesDocument.DeleteAttachment;
var
  blid: BLOCKIDcl;
begin
  if (Index < 0) or (Index >= FAttach.count) then exit;
  blid := BLOCKIDcl(FAttach.Objects[Index]);
  CheckError(NSFNoteDetachFile(Handle, blid.blockItem));
  FAttach.Objects[Index].free;
  FAttach.delete(Index);
end;

//***************************************************
procedure TNotesDocument.LoadFontTable;
type
  PCDFONTTABLE = ^CDFONTTABLE;
  PCDFACE = ^CDFACE;
var
  bhValue: BlockID;
  szValue: dword;
  pTbl: PCDFONTTABLE;
  pFace: PCDFACE;
  n: integer;
  FontID: dword;
  ff: word;
begin
  if FFontTable <> nil then exit; //alredy loaded

  // Font table resides in $FONTS item of TYPE_COMPOSITE
  FFontTable := TStringList.create;
  FMaxFontID := STATIC_FONT_FACES-1;
  if NsfItemInfo (Handle, ITEM_NAME_FONTS, length(ITEM_NAME_FONTS), nil, nil,
    @bhValue, @szValue) <> NOERROR then exit; //no such item
  pTbl := PCDFONTTABLE(DWORD(OsLockBlock(bhValue)));
  pTbl := PCDFONTTABLE(dword(pTbl) + sizeOf(word));
  try
    pFace := PCDFACE(dword(pTbl) + sizeOf(CDFONTTABLE));
    for n := 1 to pTbl^.Fonts do begin
      FontID := MakeLong (pFace^.Face, pFace^.Family);
      FFontTable.AddObject (Lmbcs2Native(strPas (pFace^.Name)), pointer(FontID));
      ff := pFace^.Face and $FF;
      if ff > FMaxFontId then FMaxFontId := ff;
      pFace := PCDFACE (dword(pFace) + ODSLength(_CDFACE));
    end;
  finally
    OsUnlockBlock(bhValue);
  end;
end;

//***************************************************
procedure TNotesDocument.SaveFontTable;
type
  PCDFONTTABLE = ^CDFONTTABLE;
  PCDFACE = ^CDFACE;
var
  MemSz: dword;
  pBuf: pointer;
  pTbl: PCDFONTTABLE;
  pFace: PCDFACE;
  FontId: dword;
  i: integer;
begin
  if (FFontTable = nil) or (FFontTable.count = 0) then exit;

  // Generate CD buffer
  MemSz := sizeOf(CDFONTTABLE) + sizeOf(CDFACE) * FFontTable.count;
  if (MemSz mod 2) <> 0 then inc(MemSz);
  GetMem (pBuf, MemSz);
  try
    // Fill header
    pTbl := PCDFONTTABLE(pBuf);
    pTbl^.Header.Signature := SIG_CD_FONTTABLE;
    pTbl^.Header.Length := MemSz;
    pTbl^.Fonts := FFontTable.count;

    // Add faces
    pFace := PCDFACE(dword(pBuf) + sizeOf(CDFONTTABLE));
    for i := 0 to FFontTable.count-1 do begin
      strCopy (pFace^.Name, pchar(Native2Lmbcs(FFontTable[i])));
      FontId := dword(FFontTable.Objects[i]);
      pFace^.Face := LoWord (FontId);
      pFace^.Family := HiWord (FontId);
      pFace := PCDFACE(dword(pFace) + sizeOf(CDFACE));
    end;

    // Add item
    NsfItemDelete (Handle, ITEM_NAME_FONTS, length(ITEM_NAME_FONTS));
    CheckError (NSFItemAppend (Handle, 0, ITEM_NAME_FONTS, length(ITEM_NAME_FONTS),
      TYPE_COMPOSITE, pBuf, MemSz));
    finally
      FreeMem (pBuf);
    end;
end;

//***************************************************
procedure TNotesDocument.UpdateUnread;
var
  uName: string;
  hTable: THandle;
begin
  // Get IDs
  uName := Database.LmbcsUserName;
  if not TNotesName.IsCanonical (uName) then uName := TNotesName.TranslateName (uName, True, '');
  hTable := 0;
  CheckError(NSFDbGetUnreadNoteTable(Database.Handle, pchar(UName), length(UName), True, @hTable));
  try
    // Check ID
    fIsRead := not (IDIsPresent(hTable, DocID));
  finally
    if hTable <> 0 then OsMemFree (hTable);
  end;
end;

(******************************************************************************)
procedure TNotesDocument.SetIsRead;
begin
  if (FIsRead <> Value) then begin
    FIsRead := Value;
    Database.MarkRead(DocId, Value);
  end;
end;

(******************************************************************************)
function TNotesDocument.GetParentDocumentID: NOTEID;
begin
  NSFNoteGetInfo(FHandle, _NOTE_PARENT_NOTEID, @Result);
end;

(******************************************************************************)
function TNotesDocument.GetResponsesCount: DWORD;
begin
  NSFNoteGetInfo(FHandle,_NOTE_RESPONSE_COUNT, @Result);
end;

(******************************************************************************)
function TNotesDocument.Responses: TNotesDocumentCollection;
var
  hTable,DocHandle: LHANDLE;
  id: dword;
  fFirst: boolean;
begin
  // OPEN A NEW DOCUMENT IN "OPEN ID TABLE OF RESPONSES" MODE
  Result := TNotesDocumentCollection.create(Database);
  Result.fUnreadDocs := false;
  CheckError(NSFNoteOpen(FDatabase.Handle, DocID, OPEN_RESPONSE_ID_TABLE, @DocHandle));
  try
    hTable := 0;
    NSFNoteGetInfo(DocHandle, _NOTE_RESPONSES , @hTable);
    if hTable <> 0 then try
      // Scan ID table
      id := 0;
      fFirst := True;
      while IDScan (hTable, fFirst, @id) do begin
        fFirst := False;
        Result.AddDocumentId (id, '');
      end;
    finally
      NsfNoteClose(DocHandle);
    end;
  except
    Result.free;
    raise;
  end;
end;

//ClassMarker_Method(TNotesDocument)


//***************************************************
// TNotesDirectory
//***************************************************
constructor TNotesDirectory.Create;
begin
  inherited Create;
  FPorts := TStringList.create;
end;

//***************************************************
destructor TNotesDirectory.Destroy;
begin
  FindClose;
  FPorts.free;
  inherited Destroy;
end;

//***************************************************
procedure TNotesDirectory.FindClose;
var
  i: integer;
begin
  if SrcTable <> nil then begin
    for i := 0 to SrcTable.count-1 do FreeMem(SrcTable[i]);
    SrcTable.free;
    SrcTable := nil;
  end;
  if hDirectory <> 0 then begin
    NSFDbClose(hDirectory);
    hDirectory := 0;
  end;
end;

(******************************************************************************)
function OnDirFound (Obj: pointer; search_info: pSEARCH_MATCH; summary_info: pITEM_TABLE): STATUS; far; stdcall;
var
  pItems: pITEM;
  pValues: pchar;
  i, n: integer;
  pEntry: PNotesDirEntry;
  sName, sValue: string;

  {pValues, valptr, pBuf: pchar;
  i, j, n: integer;
  pEntry: PNotesDirEntry;
  sName, sValue, buf: string;
  wType, bufSz: word;
  lst: pList;}
  val: variant;
begin
  if ((search_info.SERetFlags and SE_FMATCH) <> 0) then begin
    // summary_info contains a list of items describing directory entry found
    // First goes ITEM_TABLE, Items field contains total number of entries
    // Next goes array of ITEM
    pItems := pItem (dword (summary_info) + sizeOf(ITEM_TABLE));

    // Next goes array of records <name><type><value>
    // name is Item.NameLength, value is Item.ValueLength, type is WORD
    pValues := pchar (dword (summary_info) + sizeOf(ITEM_TABLE) + summary_info^.Items * sizeOf(ITEM));

    GetMem (pEntry, sizeOf(TNotesDirEntry));
    try
      FillChar (pEntry^, sizeOf(TNotesDirEntry), 0);

      // For each item...
      for i := 1 to summary_info^.Items do begin
        sValue := '';
        sName := '';
        setLength (sName, pItems^.NameLength+2);
        setLength (sValue, pItems^.ValueLength+2);

        // Get name
        strLCopy (pchar(sName), pValues, pItems^.NameLength);
        sName[pItems^.NameLength+1] := #0;
        sName := Lmbcs2Native (strPas(pchar(sName)));

        // Get value
        val := BufferToValue(pchar (dword(pValues) + pItems^.NameLength), pItems^.ValueLength);
        try
          sValue := VarAsType(val, varString);
        except
          sValue := '';
        end;

        if compareText(sName,'$title') = 0 then pEntry^.FileName := trim(sValue)
        else if compareText(sName,'$type') = 0 then pEntry^.EntryType := compareText(trim(sValue),'$dir') = 0
        else if compareText(sName,'$info') = 0 then begin
          n := Pos ('#', sValue); //template name follows #
          if n = 0
            then pEntry^.FileInfo := trim(sValue)
            else pEntry^.FileInfo := trim(copy (sValue, 1, n-2));
        end;

        pValues := pchar (dword(pValues) + pItems^.NameLength + pItems^.ValueLength);
        pItems := pItem (dword(pItems) + sizeOf(ITEM_TABLE));
      end;

      if pEntry^.FileName = ''
        then FreeMem(pEntry)
        else TList(Obj).Add(pEntry);
    except
      FreeMem (pEntry);
      raise;
    end;
  end;
  Result := NOERROR;
end;

//***************************************************
function CompareProc (Item1, Item2: Pointer): Integer;
begin
  Result := -1 * (ord(pNotesDirEntry(Item1)^.EntryType) - ord(pNotesDirEntry(Item2)^.EntryType));
  if Result = 0 then begin
    if pNotesDirEntry(Item1)^.EntryType
      then Result := compareText (pNotesDirEntry(Item1)^.FileName, pNotesDirEntry(Item2)^.FileName)
      else Result := compareText (pNotesDirEntry(Item1)^.FileInfo, pNotesDirEntry(Item2)^.FileInfo);
  end;
end;

//***************************************************
function TNotesDirectory.FindFirst;
var
  sPath: string;
  error: STATUS;
  wFileType: word;
begin
  FindClose;

  if (0 = compareText (Server, 'Local')) then SrcServer := '' else SrcServer := Server;
  SrcPath := Path;
  SrcOptions := Options;
  if SrcOptions = [] then SrcOptions := [nfoFiles];

  // Open a directory handle
  sPath := '';
  SetLength (sPath, 256);
  CheckError (OSPathNetConstruct('', pchar(SrcServer), pchar(SrcPath), pchar(sPath)));
  sPath := Native2Lmbcs (StrPas(PChar(sPath)));

  error := NSFDbOpen(pchar(sPath), @hDirectory);
  if ((hDirectory=0) and (Error=0)) or (Error=USER_CANCEL) then raise ELnUserCancel.CreateErr(-1,'User cancelled Notes session');
  CheckError (Error);

  // Enumeration
  wFileType := 0;
  if (nfoFiles in SrcOptions) then wFileType := FILE_DBANY
  else if (nfoTemplates in SrcOptions) then wFileType := FILE_FTANY;
  if (nfoSubDirs in SrcOptions) then wFileType := wFileType or FILE_DIRS;
  wFileType := wFileType or FILE_NOUPDIRS;
  SrcTable := TList.create;
  try
    // List all entries
    CheckError (NSFSearch (
      hDirectory,         //* directory handle           */
      0,                  //* selection formula          */
      nil,                //* title of view in formula   */
      SEARCH_FILETYPE +   //* search for files           */
      SEARCH_SUMMARY,     //* return a summary buffer    */
      wFileType,          // file type
      nil,                //* starting date              */
      OnDirFound,         //* call for each file found   */
      SrcTable,          //* argument to action routine */
      nil));              //* returned ending date (unused) */

    // Sort the list
    SrcTable.Sort (CompareProc);

    // Return first one
    SrcIndex := 0;
    Result := FindNext (Entry);
  except
    FindClose;
    raise;
  end;
end;

//***************************************************
function TNotesDirectory.FindNext;
begin
  if SrcIndex >= SrcTable.count then begin
    Result := False;
    FindClose;
  end
  else begin
    Result := True;
    Entry := PNotesDirEntry(SrcTable[SrcIndex])^;
    inc(SrcIndex);
  end;
end;

//***************************************************
function TNotesDirectory.GetPorts;
var
  s: string;
  n: integer;
begin
  if FPorts.count = 0 then begin
    // Read list of ports from NOTES.INI
    s := '';
    SetLength(s,256);
    OSGetEnvironmentString ('Ports', pchar(s), 255);
    s := strPas(pchar(s));
    // Entries are separated by comma
    n := Pos (',', s);
    while n <> 0 do begin
      FPorts.add (trim (copy (s, 1, n-1)));
      delete (s, 1, n);
      n := Pos (',', s);
    end;
    if s <> '' then FPorts.add (trim (copy (s, 1, n-1)));
  end;
  Result := FPorts;
end;

//***************************************************
procedure TNotesDirectory.ListServers;
var
  hList: HANDLE;
  pcName: pchar;
  pList: pointer;
  i, wCount: word;
  pwLengths: pword;
  pcNames: pchar;
  buf: string;
begin
  if Port = '' then pcName := nil else pcName := pchar(Port);
  CheckError (NSGetServerList(pcName,@hList));

  List.clear;
  List.add ('Local');
  pList := OSLockObject (hList);
  try
    // the list contains counter (word), list of name lengths and list of names
    wCount := pword(pList)^;
    pwLengths := pword (dword(pList) + sizeOf(wCount));
    pcNames := pchar (dword(pList) + sizeOf(wCount) + wCount*sizeOf(word));

    for i := 1 to wCount do begin
      buf := '';
      setLength (buf, pwLengths^ + 2);
      strLCopy (pchar (buf), pcNames, pwLengths^);
      buf[pwLengths^+1] := #0;
      buf := Lmbcs2Native (strPas(pchar(buf)));
      buf := TNotesName.TranslateName (buf, False, '');
      List.add (buf);
      pcNames := pchar (dword(pcNames) + pwLengths^);
      pwLengths := pword (dword(pwLengths) + sizeOf(word));
    end;
  finally
    OsUnlockObject (hList);
    OsMemFree (hList);
  end;
end;
//ClassMarker_Method(TNotesDirectory)

//***************************************************
// TNotesName
//***************************************************
const
  MAX_PARTS = 12;
  PartLabels: array [0..MAX_PARTS] of string = (
    'CN', 'G', 'S', 'I', 'Q', 'C', 'OU', 'OU', 'OU', 'OU', 'ADMD', 'PRMD', 'O'
  );


//***************************************************
class function TNotesName.CheckAddress;
var
  Values: TStringList;
begin
  Values := TStringList.create;
  try
    Result := LookupName (aServer, [USER_NAMESSPACE], [aName], ['FullName'], [nloExhaustive], Values);
  finally
    Values.free;
  end;
end;

//***************************************************
constructor TNotesName.Create;
begin
  inherited Create;
  SetName (aName);
end;

//***************************************************
function TNotesName.GetKeyword;
begin
  Result := format ('%s/%s/%s/%s/%s/%s', [Country, Organization, OrgUnit1, OrgUnit2, OrgUnit3, OrgUnit4]);
end;

//***************************************************
function TNotesName.GetNameComponent;

procedure getstr (var Str: string; const buf: pchar; const sz: word);
begin
  if buf = nil then Str := ''
  else begin
    SetLength (Str, sz + 1);
    strLCopy (pchar(Str), buf, sz);
    Str[sz+1] := #0;
    Str := strPas(pchar(Str));
  end;
end;

begin
  if not fParsed then begin
    CheckError (DnParse (0, nil, pchar(FName), FComponents, sizeOf(FComponents)));
    fParsed := True;
  end;
  case Index of
    0: begin
      getstr (Result, FComponents.CN, FComponents.CNLength);
      if Result = '' then Result := FName;
    end;
    1:  getstr (Result, FComponents.G, FComponents.GLength);
    2:  getstr (Result, FComponents.S, FComponents.SLength);
    3:  getstr (Result, FComponents.I, FComponents.ILength);
    4:  getstr (Result, FComponents.Q, FComponents.QLength);
    5:  getstr (Result, FComponents.C, FComponents.CLength);
    6:  getstr (Result, FComponents.OU[0], FComponents.OULength[0]);
    7:  getstr (Result, FComponents.OU[1], FComponents.OULength[1]);
    8:  getstr (Result, FComponents.OU[2], FComponents.OULength[2]);
    9: getstr (Result, FComponents.OU[3], FComponents.OULength[3]);
    10: getstr (Result, FComponents.ADMD, FComponents.ADMDLength);
    11: getstr (Result, FComponents.PRMD, FComponents.PRMDLength);
    12: getstr (Result, FComponents.O, FComponents.OLength);
  end;
end;

//***************************************************
class function TNotesName.IsCanonical;
var
  i: integer;
begin
  // Tries to find any of named part labels
  Result := True;
  aName := UpperCase(aName);
  for i := System.Low(PartLabels) to System.High(PartLabels) do
    if Pos ('/' + PartLabels[i] + '=', aName) > 0 then exit;
  Result := False;
end;

//***************************************************
function TNotesName.IsHirerarchical;
var
  n: integer;
begin
  n := Pos ('/', FName);
  Result := (n > 0) and (n < length(FName));
end;

//***************************************************
class function TNotesName.LookupName;
var
  lNameSpaces, lNames, lItems: TStringList;
  i: integer;
begin
  lNameSpaces := TStringList.create;
  lNames := TStringList.create;
  lItems := TStringList.create;
  try
    for i := System.Low(NameSpaces) to System.High(NameSpaces) do lNameSpaces.add(NameSpaces[i]);
    for i := System.Low(Names) to System.High(Names) do lNames.add(Names[i]);
    for i := System.Low(Items) to System.High(Items) do lItems.add(Items[i]);
    Result := LookupNameList (ServerName, lNameSpaces, lNames, lItems, Flags, Values);
  finally
    lNameSpaces.free;
    lNames.free;
    lItems.free;
  end;
end;

//***************************************************
class function TNotesName.LookupNameList;
const
  LookupMap: array [TNotesLookupOption] of word = (NAME_LOOKUP_ALL, NAME_LOOKUP_NOSEARCHING, NAME_LOOKUP_EXHAUSTIVE);

var
  pcServer: pchar;
  sNameSpaces, sNames, sItems: pchar;
  nNameSpaces, nNames, nItems: word;
  wFlags: word;
  iLookup: TNotesLookupOption;
  hBuffer: LHandle;
  pLookup, pName, pMatch, pItem: pchar;
  wDataType, wSize, nMatches: word;
  i, j, n, k, items_added: integer;
  buf: string;

procedure GetBuf (Arr: TStrings; var Buf: pchar; var nBuf: word);
var
  sz, n, i: integer;
begin
  sz := 0;
  Buf := nil;
  nBuf := 0;
  for i := 0 to Arr.count-1 do if Arr[i] <> '' then begin
    n := length(Arr[i]);
    ReallocMem (Buf, sz + n + 1);
    strCopy (pchar (dword(Buf) + sz), pchar(Arr[i]));
    //Buf[sz+n+1] := #0;
    inc (nBuf);
    inc (sz, n + 1);
  end;
end;

begin
  sNameSpaces := nil; sNames := nil; sItems := nil;
  try
    // Prepare params
    if (NameSpaces.count = 0) or (NameSpaces[0] = '') then NameSpaces[0] := USER_NAMESSPACE;
    GetBuf (NameSpaces, sNameSpaces, nNameSpaces);
    GetBuf (Names, sNames, nNames);
    GetBuf (Items, sItems, nItems);
    if ServerName = '' then pcServer := nil else pcServer := pchar(ServerName);
    wFlags := 0;
    for iLookup := System.Low(TNotesLookupOption) to System.High(TNotesLookupOption) do
      if iLookup in Flags then wFlags := wFlags or LookupMap[iLookup];
    if nNames = 0 then begin  //changed by Noah Silva
      nNames := 1;
      sNames := '\0';
    end;

    // Get data
    hBuffer := 0;
    CheckError (NAMELookup(pcServer,
      wFlags, nNameSpaces, sNameSpaces, nNames, sNames, nItems, sItems, hBuffer));

    // List names
    Values.clear;
    pLookup := OsLockObject (hBuffer);
    pName := nil;
    try
      for i := 1 to nNameSpaces * nNames do begin
        // Get name ptr
        pName := NAMELocateNextName (pLookup, pName, @nMatches);
        if pName = nil then break;

        pMatch := nil;
        for j := 0 to nMatches-1 do begin
          // Get match ptr
          items_added := 0;
          pMatch := NAMELocateNextMatch (pLookup, pName, pMatch);
          if pMatch = nil then break;

          for n := 0 to nItems-1 do begin
            // Get item ptr
            pItem := NAMELocateItem (pMatch, n, wDataType, @wSize);
            if pItem = nil then continue;   // if item not present in this match, go on to next match

            // Gets value
            k := 0;
            buf := '';
            setLength (buf, 255);
            while NAMEGetTextItem (pMatch, n, k, pchar(buf), 254) = NOERROR do begin
              Values.add (strPas(pchar(buf)));
              inc(k);
              inc(items_added);
            end;
          end;
          while items_added < nItems do begin
            Values.add('');
            inc(items_added);
          end;
          //Values.add(LOOKUP_VALUE_END);
        end;
      end;
    finally
      OsUnlockObject (hBuffer);
      OsMemFree (hBuffer);
    end;
  finally
    if sNameSpaces <> nil then FreeMem(sNameSpaces);
    if sNames <> nil then FreeMem(sNames);
    if sItems <> nil then FreeMem(sItems);
  end;
  Result := Values.count > 0;
end;

//***************************************************
procedure TNotesName.SetName (Value: string);
begin
  if Value <> fName then begin
    fParsed := False;
    if IsCanonical(Value)
      then fName := Value
      else fName := TranslateName (Value, True, '');
  end;
end;

//***************************************************
procedure TNotesName.SetNameComponent;
var
  parts: array [0..MAX_PARTS] of string;
  i: integer;
  Canon: string;
begin
  for i := System.Low(parts) to System.High(parts) do parts[i] := GetNameComponent(i);
  Canon := '';
  parts[Index] := Value;
  for i := System.Low(parts) to System.High(parts) do
    if parts[i] <> '' then appendStr (Canon, '/' + PartLabels[i] + '=' + parts[i]);
  Canonical := Canon;
end;

//***************************************************
function TNotesName.GetAbbreviatedName;
begin
  Result := TranslateName (Canonical, False, TemplateName);
end;

//***************************************************
procedure TNotesName.SetAbbreviatedName;
begin
  SetName (Value);
end;

//***************************************************
class function TNotesName.TranslateName;
var
  pct: pchar;
  sz: word;
begin
  sz := MAXUSERNAME;
  SetLength (Result, MAXUSERNAME + 1);
  if aTemplate = '' then pct := nil else pct := pchar(aTemplate);
  if fToCanonical
    then CheckError (DNCanonicalize(0, pct, pchar (aName), pchar(Result), 255, sz))
    else CheckError (DNAbbreviate(0, nil, pchar (aName), pchar(Result), 255, sz));
  Result := strPas(pchar(Result));
end;
//ClassMarker_Method(TNotesName)

//****************************************************
// TNotesView
//****************************************************
constructor TNotesView.Create;
begin
  inherited Create (notesDatabase);
  FName := aName;
  Load('', 0);
end;

constructor TNotesView.CreateExt(notesDatabase: TNotesDatabase; aName,
  aKey: string; nMaxDocs: integer);
begin
  inherited Create (notesDatabase);
  FName := aName;
  Load(aKey, nMaxDocs);
end;

//****************************************************
constructor TNotesView.CreateEmpty;
begin
  inherited Create (notesDatabase);
  FName := aName;
end;

//****************************************************
constructor TNotesView.CreateSearch(notesDatabase: TNotesDatabase; aName: string);
begin
  inherited Create (notesDatabase);
  FName := aName;
  LoadByID(GetViewID(Database,Name), '', 0, false);
end;

//****************************************************
destructor TNotesView.Destroy;
begin
  if (FHcoll<>0) then CheckError(NIFCloseCollection (fhColl));
  inherited Destroy;
end;

//****************************************************
function TNotesView.GetIsFolder;
begin
  Result := Pos(DESIGN_FLAG_FOLDER_VIEW,Flags) > 0;
end;

//****************************************************
function TNotesView.GetIsShared;
begin
  Result := Pos(DESIGN_FLAG_PRIVATE_IN_DB,Flags) = 0;
end;

//****************************************************
class function TNotesView.GetViewID;
var
  Name2: string;
  err: dword;
  n: integer;
begin
  Name2 := Native2Lmbcs(aName);
  err := NIFFindDesignNote (notesDatabase.Handle, PChar(Name2), NOTE_CLASS_VIEW, @Result);
  if err <> 0 then err := NIFFindPrivateDesignNote (notesDatabase.Handle, PChar(Name2),
    NOTE_CLASS_VIEW, @Result);
  if err <> 0 then begin
    // Do the same for view alias (if present)
    n := LastDelimiter('|',aName);
    if n > 0 then begin
      Name2 := Native2Lmbcs(copy(aName,n+1,length(aName)-n));
      err := NIFFindDesignNote (notesDatabase.Handle, PChar(Name2), NOTE_CLASS_VIEW, @Result);
      if err <> 0 then err := NIFFindPrivateDesignNote (notesDatabase.Handle, PChar(Name2),
        NOTE_CLASS_VIEW, @Result);
    end;
  end;
  CheckError(err);
end;

//****************************************************
function TNotesView.GetOriginatorID;
var
  td: TIMEDATE;
  cls: word;
begin
  CheckError(NSFDbGetNoteInfo(Database.Handle, ID, @Result, @td, @cls));
end;

//****************************************************
function TNotesView.GetUniversalID;
var
  id: OID;
begin
  id := GetOriginatorID;
  Result.aFile := id.FileNum;
  Result.Note := id.Note;
end;

//****************************************************
procedure TNotesView.Load;
begin
  LoadByID(GetViewID(Database,Name), aKey, nMaxDocs, true);
end;

//****************************************************
procedure TNotesView.LoadByID;
var
  hColl: HCOLLECTION;
begin
  // Get flags
  FNoteID := anID;
  FKey := aKey;
  FMaxDocs := nMaxDocs;
  Flags := LoadFlags (Database, ID);
  if (not IsFolder) and (self.className = TNotesFolder.className) then
    raise ELnInvalidFolder.Create ('Attempting to open view ' + Name + ' as a folder');

  // Load docs
  hColl := 0;
  try
    CheckError(NIFOpenCollection(
             Database.Handle,         { database handle }
             Database.Handle,         { database handle }
             ID,                      { view note ID}
             0,
             0,
             hColl,
             nil,
             nil,
             nil,
             nil));

    // Added by Matt Saint
    if not FreeCollectionOnClose
      then FHColl := hcoll
      else LoadColl(hColl, aKey, nMaxDocs);
  finally
    if (hColl <> 0) and (FreeCollectionOnClose) then CheckError(NIFCloseCollection (hColl));
  end;
end;

procedure TNotesView.LoadColl(hcoll: LHandle; aKey: string; nMaxDocs: integer);
var
  CollPosition: COLLECTIONPOSITION;
  hBuffer: LHandle;
  NotesFound: dword;
  SignalFlags: word;
  SkipNavigator: word;
  ReturnCount: longint;
  Error: word;
begin
  // Added by piton
  FillChar(CollPosition,sizeOf(CollPosition),0);
  if aKey = '' then begin
    SkipNavigator := NAVIGATE_NEXT;
    if nMaxDocs > 0
      then ReturnCount := nMaxDocs
      else ReturnCount := $FFFF;
  end
  else begin
    error := NIFFindByName(hColl, pChar(key), FIND_CASE_INSENSITIVE, @CollPosition, @ReturnCount );
    if (error and $3fff) = ERR_NOT_FOUND
      then ReturnCount := 0
      else CheckError(error);
    SkipNavigator := NAVIGATE_CURRENT;
    if nMaxDocs > 0 then ReturnCount := nMaxDocs;
  end;

  repeat
    hBuffer := 0;
    try
      CheckError(NIFReadEntries(
        hColl,            { handle to this collection }
        @CollPosition,    { where to start in collection }
        SkipNavigator,    { order to skip entries }
        1,                { number to skip }
        NAVIGATE_NEXT,    { order to use after skipping }
        ReturnCount,      { max return number }
        READ_MASK_NOTEID or READ_MASK_SUMMARY,      { info we want }
        @hBuffer,         { handle to info (return) }
        nil,              { length of buffer (return) }
        nil,              { entries skipped (return) }
        NotesFound,       { number of notes (return) }
        SignalFlags));    { share warning (return) }

      if (hBuffer<>0) then GetCollectionByHandle(hBuffer, NotesFound, True);
    finally
      if hBuffer <> 0 then OSMemFree(hBuffer);
    end;
  until (SignalFlags and SIGNAL_MORE_TO_DO) = 0;
end;

//****************************************************
class function TNotesView.LoadFlags;
var
  hNote: LHandle;
begin
  CheckError(NsfNoteOpen(Database.Handle, anID, OPEN_NOUPDATE, @hNote));
  try
    setLength (Result,31);
    NsfItemGetText(hNote, DESIGN_FLAGS, pchar(Result), 30);
    Result := strPas(pchar(Result));
  finally
    NsfNoteClose(hNote);
  end;
end;

//****************************************************
class function TNotesView.OpenView;
begin
  Result := OpenViewExt(notesDatabase, aName, '', 0);
end;

//****************************************************
class function TNotesView.OpenViewExt(notesDatabase: TNotesDatabase; aName,
  aKey: string; nMaxDocs: integer): TNotesView;
var
  Flags: string;
  ID: NOTEID;
begin
  Id := GetViewID(notesDatabase, aName);
  Flags := LoadFlags (notesDatabase, ID);
  if Pos(DESIGN_FLAG_FOLDER_VIEW,Flags) > 0
    then Result := TNotesFolder.CreateEmpty(notesDatabase,aName)
    else Result := TNotesView.CreateEmpty(notesDatabase,aName);
  Result.LoadByID(ID, aKey, nMaxDocs, true);
end;

//****************************************************
procedure TNotesView.Update;
var
  hColl: HCOLLECTION;
  ffree: boolean;
begin
  ffree := (FHColl = 0);
  if not ffree then hColl := FHColl
  else begin
    CheckError(NIFOpenCollection(Database.Handle,Database.Handle, ID, 0, 0, hColl,
      nil, nil, nil, nil));
    end;
  try
    CheckError(NIFUpdateCollection(hColl));
    Clear;
    LoadColl(hColl, Key, MaxDocs);
  finally
    if (hColl <> 0) and (ffree) then CheckError(NIFCloseCollection(hColl));
  end;
end;

//****************************************************
function TNotesView.GetAllDocumentByKey(akey: string; nmaxdocs:Integer): TNotesDocumentCollection;
var
  CollPosition: COLLECTIONPOSITION;
  hBuffer: LHandle;
  NotesFound: dword;
  SignalFlags: word;
  SkipNavigator: word;
  ReturnCount: longint;
  Error: word;
begin
  Result := nil;
  if FHcoll=0 then exit;

  Result := TNotesDocumentCollection.create(Database);
  Result.fUnreadDocs := false;
  FillChar(CollPosition,sizeOf(CollPosition),0);
  FKey:=akey;

  error := NIFFindByName(fhColl, pChar(key), FIND_CASE_INSENSITIVE, @CollPosition, @ReturnCount );
  if (error and $3fff) = ERR_NOT_FOUND then ReturnCount := 0 else CheckError(error);

  SkipNavigator := NAVIGATE_CURRENT;
  if nMaxDocs > 0 then ReturnCount := nMaxDocs;

  repeat
    hBuffer := 0;
    try
      CheckError(NIFReadEntries(
        fhColl,            { handle to this collection }
        @CollPosition,    { where to start in collection }
        SkipNavigator,    { order to skip entries }
        1,                { number to skip }
        NAVIGATE_NEXT,    { order to use after skipping }
        ReturnCount,      { max return number }
        READ_MASK_NOTEID or READ_MASK_SUMMARY,      { info we want }
        @hBuffer,         { handle to info (return) }
        nil,              { length of buffer (return) }
        nil,              { entries skipped (return) }
        NotesFound,       { number of notes (return) }
        SignalFlags));    { share warning (return) }

      if (hBuffer<>0) then Result.GetCollectionByHandle(hBuffer, NotesFound, True);
    finally
      if hBuffer <> 0 then OSMemFree(hBuffer);
    end;
  until (SignalFlags and SIGNAL_MORE_TO_DO) = 0;
end;

//ClassMarker_Method(TNotesView)


//****************************************************
// TNotesFolder
//****************************************************
constructor TNotesFolder.CreateNew;
var
  hFormatNote, hFormatDB: LHandle;
  dwDesign: DESIGN_TYPE;
  Name2: string;
begin
  CreateEmpty(notesDatabase, aName);
  Flags := DESIGN_FLAG_FOLDER_VIEW;
  if not fShared then appendStr (Flags,DESIGN_FLAG_PRIVATE_IN_DB);

  Name2 := Native2Lmbcs(aName);
  if FormatFolder = nil then begin
    hFormatNote := 0;
    hFormatDB := 0;
  end
  else begin
    hFormatNote := FormatFolder.ID;
    hFormatDB := FormatFolder.Database.Handle;
  end;

  if fShared then dwDesign := DESIGN_TYPE_SHARED else dwDesign := DESIGN_TYPE_PRIVATE_DATABASE;
  CheckError(FolderCreate(Database.Handle,0,hFormatNote,hFormatDB,pchar(Name2),
    length(Name2),dwDesign,0,FNoteID));
  Database.UpdateViews;
end;

//****************************************************
procedure TNotesFolder.AddDocument;
var
  Lst: TList;
begin
  Lst := TList.create;
  try
    Lst.add (Doc);
    AddDocuments(Lst);
  finally
    Lst.free;
  end;
end;

//****************************************************
procedure TNotesFolder.AddDocuments;
var
  hTable: LHandle;
  i: integer;
begin
  CheckError (IDCreateTable(sizeOf(NOTEID),@hTable));
  try
    for i := 0 to DocList.count-1 do
      CheckError(IDInsert(hTable,TNotesDocument(DocList[i]).DocID, nil));
    CheckError(FolderDocAdd(Database.Handle,0,ID,hTable,0));
  finally
    IDDestroyTable(hTable);
  end;
end;

//****************************************************
procedure TNotesFolder.RemoveDocuments;
begin
  CheckError(FolderDocRemoveAll(Database.Handle, 0, ID, 0));
end;

//****************************************************
function TNotesFolder.Copy;
var
  nid: NOTEID;
  Name2: string;
begin
  Name2 := Native2Lmbcs(NewName);
  CheckError(FolderCopy(Database.Handle,0,ID,pchar(Name2),length(Name2),0,nid));
  Database.UpdateViews;
  Result := TNotesFolder.Create(Database,NewName);
end;

//****************************************************
procedure TNotesFolder.Delete;
begin
  CheckError(FolderDelete(Database.Handle,0,ID,0));
  Database.UpdateViews;
end;

//****************************************************
procedure TNotesFolder.DeleteDocument;
var
  Lst: TList;
begin
  Lst := TList.create;
  try
    Lst.add (Doc);
    DeleteDocuments(Lst);
  finally
    Lst.free;
  end;
end;

//****************************************************
procedure TNotesFolder.DeleteDocuments;
var
  hTable: LHandle;
  i: integer;
begin
  CheckError (IDCreateTable(sizeOf(NOTEID),@hTable));
  try
    for i := 0 to DocList.count-1 do
      CheckError(IDInsert(hTable,TNotesDocument(DocList[i]).DocID, nil));
    CheckError(FolderDocRemove(Database.Handle,0,ID,hTable,0));
  finally
    IDDestroyTable(hTable);
  end;
end;

//****************************************************
procedure TNotesFolder.Move;
begin
  CheckError(FolderMove(Database.Handle,0,ID,0,ParentFolder.ID,0));
  Database.UpdateViews;
end;

//****************************************************
procedure TNotesFolder.SetName;
var
  Name2: string;
begin
  if (Value <> FName) then begin
    Name2 := Native2Lmbcs(Value);
    CheckError(FolderRename(Database.Handle,0,ID,pchar(Name2),length(Name2),0));
    FName := Value;
    Database.UpdateViews;
  end;
end;
//ClassMarker_Method(TNotesFolder)

//****************************************************
// TAclRoles
//****************************************************
type
  TAclRoles = class(TStringList)
  private
    ACL: TNotesACL;
    InInit: boolean;
  protected
    procedure Put(Index: Integer; const S: string); override;
    procedure Init;
  public
    function Add(const S: string): Integer; override;
    procedure Clear; override;
    procedure Delete(Index: Integer); override;
    procedure Insert(Index: Integer; const S: string); override;
  end;

//****************************************************
procedure TAclRoles.Init;
var
  i: word;
  buf, str: string;
  err: STATUS;
begin
  InInit := True;
  try
    Clear;
    Duplicates := dupIgnore;
    buf := '';
    SetLength(buf, ACL_PRIVSTRINGMAX+1);
    for i := 0 to ACL_PRIVCOUNT-1 do begin
      err := ACLGetPrivName(ACL.Handle,i,pchar(buf));
      if err <> 1060 then CheckError(err) else break;
      str := trim(strPas(pchar(buf)));
      if str <> '' then Add(str);
    end;
  finally
    InInit := False;
  end;
end;

//****************************************************
procedure TAclRoles.Put(Index: Integer; const S: string);
begin
  if not InInit then CheckError(ACLSetPrivName(ACL.Handle, Index, pchar(S)));
  inherited Put(Index,S);
end;

//****************************************************
function TAclRoles.Add(const S: string): Integer;
begin
  if not InInit then CheckError(ACLSetPrivName(ACL.Handle, Count+1, pchar(S)));
  Result := inherited Add(S);
end;

//****************************************************
procedure TAclRoles.Clear;
begin
  if InInit
    then inherited Clear
    else raise ELotusNotes.createErr(-1, 'Action is not supported');
end;

//****************************************************
procedure TAclRoles.Delete(Index: Integer);
begin
  raise ELotusNotes.createErr(-1, 'Action is not supported');
  //inherited Delete(Index);
end;

//****************************************************
procedure TAclRoles.Insert(Index: Integer; const S: string);
begin
  if not InInit then Add(S) else inherited Insert(Index,S);
end;

//****************************************************
// ACL maps
//****************************************************
const
  AccLevelMap: array [TNotesAclAccessLevel] of word = (
    ACL_LEVEL_NOACCESS,
    ACL_LEVEL_DEPOSITOR,
    ACL_LEVEL_READER,
    ACL_LEVEL_AUTHOR,
    ACL_LEVEL_EDITOR,
    ACL_LEVEL_DESIGNER,
    ACL_LEVEL_MANAGER
  );
  AclFlagMap: array [TNotesAclFlag] of word = (
    ACL_FLAG_AUTHOR_NOCREATE,
    ACL_FLAG_SERVER,
    ACL_FLAG_NODELETE,
    ACL_FLAG_CREATE_PRAGENT,
    ACL_FLAG_CREATE_PRFOLDER,
    ACL_FLAG_PERSON,
    ACL_FLAG_GROUP,
    ACL_FLAG_CREATE_FOLDER,
    ACL_FLAG_CREATE_LOTUSSCRIPT,
    ACL_FLAG_PUBLICREADER,
    ACL_FLAG_PUBLICWRITER,
    ACL_FLAG_ADMIN_READERAUTHOR,
    ACL_FLAG_ADMIN_SERVER
  );

//****************************************************
// TNotesACL
//****************************************************
constructor TNotesACL.Create;
begin
  inherited Create;
  FDatabase := aDatabase;
  FRoles := TAclRoles.create;
  TAclRoles(FRoles).ACL := self;
  FEntries := TList.create;

  CheckError(NSFDBReadACL(Database.Handle,@FHandle));
  TAclRoles(FRoles).Init;
  ReadEntries;
end;

//****************************************************
procedure TNotesACL.DeleteACLEntry;
var
  e: TNotesAclEntry;
begin
  e := Entry[aName];
  if e = nil then exit;
  CheckError(AclDeleteEntry(Handle, pchar(aName)));
  FEntries.remove(e);
  e.free;
end;

//****************************************************
destructor TNotesACL.Destroy;
var
  i: integer;
begin
  for i := 0 to FEntries.count-1 do TObject(FEntries[i]).free;
  FEntries.free;
  FRoles.free;
  if FHandle <> 0 then OsMemFree(FHandle);
  inherited Destroy;
end;

//****************************************************
function TNotesACL.GetEntriesCount;
begin
  Result := FEntries.count;
end;

//****************************************************
function TNotesACL.GetEntry;
var
  i: integer;
begin
  for i := 0 to EntriesCount-1 do begin
    Result := EntryByIndex[i];
    if compareText(Result.Name, aName) = 0 then exit;
  end;
  Result := nil;
end;

//****************************************************
function TNotesACL.GetEntryByIndex;
begin
  Result := TNotesAclEntry(FEntries[Index]);
end;

//****************************************************
function TNotesACL.GetUniformAccess;
var
  f: dword;
begin
  CheckError(AclGetFlags(Handle, f));
  Result := (ACL_UNIFORM_ACCESS and f) <> 0;
end;

//****************************************************
procedure TNotesACL.SetUniformAccess;
var
  f: dword;
begin
  if Value then f := ACL_UNIFORM_ACCESS else f := 0;
  CheckError(AclSetFlags(Handle, f));
end;

//****************************************************
function TNotesACL.CreateACLEntry;
begin
  if Entry[aName] <> nil then
    raise ELotusNotes.CreateErr(-1, 'Entry with name ' + aName + ' already exists');
  Result := TNotesAclEntry.CreateNew(self, aName, aclNoAccess, '');
  FEntries.Add(Result);
end;

//****************************************************
procedure EntriesProc (Param: pointer; Name: pchar; AccLevel: word;
  Privileges: PACL_PRIVILEGES; AccFlags: WORD); stdcall; far;
var
  e: TNotesACLEntry;
  AccessLevel: TNotesAclAccessLevel;
  AccessFlags: TNotesAclFlags;
  Flag: TNotesAclFlag;
begin
  for AccessLevel := System.Low(TNotesAclAccessLevel) to System.High(TNotesAclAccessLevel) do
    if AccLevelMap[AccessLevel] = AccLevel then break;
  AccessFlags := [];
  for Flag := System.Low(TNotesAclFlag) to System.High(TNotesAclFlag) do
    if (AccFlags and AclFlagMap[Flag]) <> 0 then include(AccessFlags, Flag);
  e := TNotesACLEntry.create(TNotesAcl(Param), strPas(Name), AccessLevel, Privileges^, AccessFlags);
  TNotesAcl(Param).FEntries.Add(e);
end;

//****************************************************
procedure TNotesACL.ReadEntries;
begin
  CheckError(AclEnumEntries(Handle,EntriesProc,self));
end;

//****************************************************
procedure TNotesACL.Save;
var
  i: integer;
begin
  for i := 0 to EntriesCount-1 do EntryByIndex[i].Update;
  CheckError(NsfDBStoreAcl(Database.Handle, FHandle, 0, 0));
end;

//****************************************************
function TNotesACL.GetMaxInternetAccess: TNotesAclAccessLevel;
begin
  CheckError(NSFGetMaxPasswordAccess(Database.Handle, PWord(@Result)));
end;

//****************************************************
procedure TNotesACL.SetMaxInternetAccess(const Value: TNotesAclAccessLevel);
begin
  CheckError(NSFSetMaxPasswordAccess(Database.Handle, Word(Value)));
end;
//ClassMarker_Method(TNotesACL)


//****************************************************
// TNotesACLEntry
//****************************************************
procedure TNotesACLEntry.AddRole;
var
  s: string;
begin
  s := Roles;
  if Pos (aRole+#13#10, s) = 0 then Roles := s + aRole + #13#10;
end;

//****************************************************
constructor TNotesACLEntry.Create;
begin
  inherited Create;
  FUpdateFlags := 0;
  FAcl := anAcl;
  FName := aName;
  FAccessLevel := AccLevel;
  FPrivileges := AclPrivs;
  FFlags := AccFlags;
end;

//****************************************************
constructor TNotesACLEntry.CreateNew;
begin
  inherited Create;
  FNew := True;
  FUpdateFlags := ACL_UPDATE_NAME or ACL_UPDATE_LEVEL or ACL_UPDATE_PRIVILEGES or ACL_UPDATE_FLAGS;
  FAcl := anAcl;
  Name := aName;
  AccessLevel := AccLevel;
  Roles := aRoles;
end;

//****************************************************
procedure TNotesACLEntry.SetAccessLevel;
begin
  if (FAccessLevel <> Value) then begin
    FAccessLevel := Value;
    FUpdateFlags := FUpdateFlags or ACL_UPDATE_LEVEL;
  end;
end;

//****************************************************
procedure TNotesACLEntry.SetFlags;
begin
  if (FFlags <> Value) then begin
    FFlags := Value;
    FUpdateFlags := FUpdateFlags or ACL_UPDATE_FLAGS;
  end;
end;

//****************************************************
procedure TNotesACLEntry.SetName;
begin
  if (FName <> Value) then begin
    if FOldName = '' then FOldName := FName;
    FName := Value;
    FUpdateFlags := FUpdateFlags or ACL_UPDATE_NAME;
  end;
end;

//****************************************************
procedure TNotesACLEntry.SetPrivileges;
begin
  FPrivileges := Value;
  FUpdateFlags := FUpdateFlags or ACL_UPDATE_PRIVILEGES;
end;

//****************************************************
procedure TNotesACLEntry.SetRoles;
var
  rl: TStringList;
  i, n: integer;
begin
  rl := TStringList.create;
  try
    rl.Text := Value;
    FillChar (FPrivileges, sizeOf(FPrivileges), 0);
    for i := 0 to rl.count-1 do begin
      n := ACL.Roles.IndexOf(rl[i]);
      if n >= 0 then AclSetPriv (FPrivileges, n+ACL_BITPRIVCOUNT);
    end;
    FUpdateFlags := FUpdateFlags or ACL_UPDATE_PRIVILEGES;
  finally
    rl.free;
  end;
end;

//****************************************************
procedure TNotesACLEntry.DeleteRole;
var
  s: string;
  n: integer;
begin
  s := Roles;
  n := Pos (aRole+#13#10, s);
  if n > 0 then begin
    delete (s, n, length(aRole + #13#10));
    Roles := s;
  end;
end;

//****************************************************
destructor TNotesACLEntry.Destroy;
begin
  Update;
  inherited Destroy;
end;

//****************************************************
function TNotesACLEntry.GetRoles;
var
  i: integer;
begin
  Result := '';
  for i := 0 to ACL_PRIVCOUNT-1 do if ACLIsPrivSet(FPrivileges, i) then begin
    AppendStr(Result, ACL.Roles[i-ACL_BITPRIVCOUNT] + #13#10);
  end;
end;

//****************************************************
procedure TNotesACLEntry.Update;
var
  flag: TNotesAclFlag;
  wFlags: word;
  pcName, pcOldName: pchar;
begin
  if (not FNew) and (FUpdateFlags = 0) then exit;
  if FOldName = '' then FOldName := FName;

  // By Joe Barbaretta
  // Default ACL entry is represented by NIL
  if (FName = '') or (compareText(FName,'-Default-') = 0) then pcName := nil else pcName := pchar(FName);
  if (FOldName = '') or (compareText(FOldName,'-Default-') = 0) then pcOldName := nil else pcOldName := pchar(FOldName);

  wFlags := 0;
  for flag := System.Low(TNotesAclFlag) to System.High(TNotesAclFlag) do
    if flag in Flags then wFlags := wFlags or AclFlagMap[flag];

  if FNew
    then CheckError(AclAddEntry(ACL.Handle, pcName, AccLevelMap[AccessLevel], FPrivileges, wFlags))
    else CheckError(AclUpdateEntry(Acl.Handle,pcOldName,FUpdateFlags,pcName,AccLevelMap[AccessLevel],FPrivileges,wFlags));

  FUpdateFlags := 0;
  FNew := False;
  FOldName := '';
end;

//****************************************************
{ TNotesServer }
//****************************************************
function DisplayTrav(Context: Pointer; Facility: PChar;
                     StatName: PChar; ValueType: Word;
                     Value: Pointer): STATUS; far; stdcall;
var
  StatRec: TStat;
  NameBuffer, ValueBuffer: String;
  begin
  SetLength(NameBuffer, 80);
  SetLength(ValueBuffer, 80);
  StatRec.Facility := Facility;
  StatRec.StatName := StatName;

  StatToText(Facility, StatName, ValueType, Value,
      PChar(NameBuffer), length(NameBuffer)-1,
      PChar(ValueBuffer), length(ValueBuffer)-1);

  StatRec.NameBuffer := Lmbcs2Native(NameBuffer);
  StatRec.ValueBuffer := Lmbcs2Native(ValueBuffer);

  TStatCollection(Context).Add( StatRec );

  Result := (NOERROR);
end;

//****************************************************
constructor TNotesServer.Create(ServerName: String);
begin
  inherited Create;
  StatCollection := TStatCollection.Create;
  FServerName := ServerName;
end;

//****************************************************
destructor TNotesServer.Destroy;
begin
  StatCollection.Free;
  inherited Destroy;
end;
//****************************************************
function TNotesServer.GetConsoleInfo(cmd: String): String;
var
  hRetInfo: HANDLE;
  pBuffer: PChar;
begin
  CheckError(NSFRemoteConsole (PChar(FServerName), PChar(cmd), @hRetInfo));
  pBuffer := OSLockObject(hRetInfo);

  Result := String(pBuffer);
  OSUnlockObject(hRetInfo);
  OSMemFree(hRetInfo);
end;

//****************************************************
function TNotesServer.GetHasMoreStatistics: boolean;
begin
  Result := StatCollection.HasMoreElements;
end;

//****************************************************
function TNotesServer.NextStat: TStat;
begin
  Result := StatCollection.NextElement;
end;

//****************************************************
function TNotesServer.QueryLocalStatistics(Facility, StatName: PChar): word;
begin
  StatTraverse(PChar(Facility),
               PChar(StatName),
               DisplayTrav,
               StatCollection);
  Result := StatCollection.Count;
end;

//****************************************************
procedure TNotesServer.parseToList( sList: TStrings );
var
  idx, apos: integer;
  tmp: String;
  tS: TStat;
begin
  // Build our list
  for idx := 0 to sList.Count -1 do begin
    tmp := sList.Strings[ idx ];
    aPos := pos('=', tmp);
    if ( aPos > 0) then begin
      tS.NameBuffer := Lmbcs2Native(trim(Copy(tmp, 1, aPos - 1)));
      tS.ValueBuffer := Lmbcs2Native(trim(Copy(tmp, aPos + 1, length(tmp))));
      StatCollection.Add( tS );
    end;
  end;
end;

//****************************************************
// use the NSFRemoteConsole to gain and extract whatever
// we can get that way...
function TNotesServer.QueryRemoteStatistics: word;
var
  sList: TStrings;
  aVar: String;
begin
  sList := TStringList.Create;
  try
    aVar := GetConsoleInfo('show stat');
    sList.Text := aVar;

    parseToList( sList );

    Result := StatCollection.Count;
  finally
    sList.Free;
  end;
end;

//****************************************************
{ TStatCollection }
{ Collection of stat records returned by TNotesServer }
//****************************************************

constructor TStatCollection.Create;
begin
  FList := TList.Create;   // Contains server statistical information
  FCount := 0;
end;

destructor TStatCollection.Destroy;
begin
  Clear;
  FList.Free;
  inherited Destroy;
end;

procedure TStatCollection.Add(aStatRec: TStat);
var
  StatRecord: PStatRecord;
begin
  new(StatRecord);
  StatRecord^.Facility := aStatRec.Facility;
  StatRecord^.NameBuffer := aStatRec.NameBuffer;
  StatRecord^.StatName := aStatRec.StatName;
  StatRecord^.ValueBuffer := aStatRec.ValueBuffer;

  FList.Add( StatRecord );   // And add the element
end;

procedure TStatCollection.Clear;
var
  idx: integer;
  StatRecord: PStatRecord;
begin
  // Clean up statistics list
  for idx := 0 to (FList.Count - 1) do begin
    StatRecord := FList.Items[idx];
    Dispose(StatRecord);
  end;
end;

function TStatCollection.GetHasMoreElements: boolean;
begin
  if (FCount < FList.Count)
    then Result := true
    else Result := false;
end;

function TStatCollection.GetMaxCount: word;
begin
  Result := FList.Count;
end;

function TStatCollection.NextElement: TStat;
var
  pStat: PStatRecord;
  aStat: TStat;
begin
  pStat := FList.Items[FCount];
  aStat.Facility := pStat.Facility;
  aStat.NameBuffer := pStat.NameBuffer;
  aStat.StatName := pStat.StatName;
  aStat.ValueBuffer := pStat.ValueBuffer;
  FCount := FCount + 1;
  Result := aStat;
end;

{$IFNDEF NO_INIT_SECTION}
initialization
{$IFDEF EXTENDED_INIT}
  InitNotesExt;
{$ELSE}
  InitNotes;
{$ENDIF}
{$IFNDEF NO_CTRL_BREAK}
  OldBreakProc := OsGetSignalHandler (OS_SIGNAL_CHECK_BREAK);
  OsSetSignalHandler (OS_SIGNAL_CHECK_BREAK, @NDBreakProc);
{$ENDIF}
finalization
{$IFNDEF NO_CTRL_BREAK}
  if InitDone then OsSetSignalHandler (OS_SIGNAL_CHECK_BREAK, @OldBreakProc);
{$ENDIF}
  if InitDone then NotesTerm;
{$ENDIF}
end.



