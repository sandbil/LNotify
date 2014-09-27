{==============================================================================|
| Project : Notes/Delphi class library                           | 3.9.2       |
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
|     25.08.01                                                                 |
|==============================================================================|
| Contributors and Bug Corrections:                                            |
|   Fujio Kurose                                                               |
|   Noah Silva                                                                 |
|   Tibor Egressi                                                              |
|   Andreas Pape                                                               |
|   Anatoly Ivkov                                                              |
|   Winalot                                                                    |
|   Olaf Hahnl                                                                 |
|     and others...                                                            |
|==============================================================================|
| History: see README.TXT                                                      |
|==============================================================================|
| Rich-text support classes and routines                                       |
|==============================================================================|}
unit Class_NotesRTF;

{$IFNDEF WIN32}
-- This unit is for Windows 32 environment
{$ENDIF}
{$RANGECHECKS OFF}
{$ALIGN OFF}

{$INCLUDE Util_LnVersion.inc}

interface
uses SysUtils, Classes, Windows, Class_LotusNotes, Util_LnAPIErr, Util_LnApi;

type
  TNotesRichTextItem = class;
  
  // Font and para parameters
  TRichTextJustification = (rjNone, rjLeft, rjCenter, rjRight, rjBlock);
  TRichTextFont = (rfRoman, rfSwiss, rfMonospace);
  TRichTextStyleOption = (
    rsPaginateBefore,   //* start new page with this par */
    rsKeepWithNext,     //* don't separate this and next par */
    rsKeepTogether,     //* don't split lines in paragraph */
    rsPropagate,        //* propagate even PAGINATE_BEFORE and KEEP_WITH_NEXT */
    rsHideReadOnly,     //* hide paragraph in R/O mode */
    rsHideEdit,         //* hide paragraph in R/W mode */
    rsHidePrint,        //* hide paragraph when printing */
    rsDisplayRM,        //* honor right margin when displaying to a window */
    rsHideCopy,         //* hide paragraph when copying/forwarding */
    rsBullet,           //* display paragraph with bullet */
    rsHideIF,           //*  use the hide when formula even if there is one. ??? - kol*/
    rsNumberList,       //* display paragraph with number */
    rsHidePreview,      //* hide paragraph when previewing*/
    rsHidePreviewPane,  //* hide paragraph when editing in the preview pane.    */
    rsHideNotes         //* hide paragraph from Notes clients */
  );
  TRichTextStyleOptions = set of TRichTextStyleOption;

  // Table parameters
  TRichTextTableOption = (rtAutoWidth, rtBorderEmboss, rtBorderExtrude);
  TRichTextTableOptions = set of TRichTextTableOption;
  TRichTextCellOption = (rcUseBkColor, rcInvisibleH, rcInvisibleV);
  TRichTextCellOptions = set of TRichTextCellOption;
  TRichTextCell = record            //use RichTextCell funct.to fill this quickly
    LeftMargin, RightMargin: word;
    Borders: array[1..4] of word;  //left,top,right,bottom
    Options: TRichTextCellOptions;
    BackColor: word;
    FractWidth: word;
    bRowSpan: boolean;
    bColSpan: boolean
  end;

  // Section parameters
  TRichTextSectionFlag = (
    rtsDisabledForNonEditors,
    rtsBorderInvisible,
    rtsHideExpanded,
    rtsAutoExpandRead,
    rtsAutoExpandPreview,
    rtsAutoExpandEdit,
    rtsAutoExpandPrint,
    rtsAutoCollapseRead,
    rtsAutoCollapsePreview,
    rtsAutoCollapseEdit,
    rtsAutoCollapsePrint,
    rtsExpanded,
    rtsDisabled
  );
  TRichTextSectionFlags = set of TRichTextSectionFlag;
  TRichTextSectionBorder = (
    rtsBorderShadow,
    rtsBorderNone,
    rtsBorderSingle,
    rtsBorderDouble,
    rtsBorderTriple,
    rtsBorderTwoline
  );

  // Enumeration routine prototype
  TNotesRTFReadProc = function (Item: TNotesRichTextItem;
                                 RecordPtr: pointer;
                                 RecordType: WORD;
                                 RecordLength: DWORD): STATUS of object;

  // OLE2 Attachment format
  TRichTextOle2AttachType = (
    rtaText,        //text with no formatting
    rtaMetafile,    //metafile
    rtaBitmap,      //bitmap
    rtaRTF,         //rich text
    rtaOwnerLink,   //OLE owner link (ext.reference)
    rtaObjectLink,  //OLE object link (ext.reference)
    rtaNative,      //OLE native format (ext.reference)
    rtaIcon         //only file icon
  );

  // Rich-text item class
  TNotesRichTextItem = class(TNotesItem)
  private
    FAttach: TStrings;
    FContext: pointer;
    FCurPtr: pointer;
    FLength: dword;
    FCurStyle: integer;
    FNewStyle: boolean;
    FHasPara: boolean;
    FParaJustification: TRichTextJustification;
    FParaIntProperties: array [0..5] of integer;
    FParaStyleOptions: TRichTextStyleOptions;
    FFontBoolProperties: array [0..8] of boolean;
    FFontColor: word;
    FFontFace: TRichTextFont;
    FFontFaceName: string;
    FFontSize: integer;
    FFontTableChanged: boolean;
    FPlainText: boolean;
    FPostedAttachments: TStringList;
    FStringsValue: TStringList;
    FReadProc: TNotesRTFReadProc;

    // ************** by Olaf **************
    FLinks: TList;            // the list holding the Link-definitions
    //**************************************

    procedure AddStyle;
    procedure AddStyle2(Flags: word);
    procedure AddMem (const Sz: dword);
    function CheckOdd (const sz: dword): dword;
    procedure LoadAttachmentInfo;
    function GetAttachment(Index: integer): string;
    function GetAttachmentCount: integer;
    function GetFontBoolProperty (Index: integer): boolean;
    function  GetParaIntProperty (Index: integer): integer;
    procedure SetFontBoolProperty (Index: integer; Value: boolean);
    procedure SetParaIntProperty (Index: integer; Value: integer);
    procedure SetParaJustification (Value: TRichTextJustification);
    procedure SetParaStyleOptions (Value: TRichTextStyleOptions);
    procedure AttachFiles;
    // ************** by Olaf **************
    procedure LoadLinkInfo;
    function GetLink(Index : integer) : LinkDef;
    function GetLinkCount : integer;
    //**************************************
  protected
    procedure CreateDefaults;
    function GetRichText: TStrings; override;
    procedure SetRichText (Value: TStrings); override;
    procedure IntAddFile(FileName, DllName, DllName2: string);
    procedure AddLink(const DatabaseID: DBID; const ViewID, DocID: UNID; Title: string; TitleLen: integer);
  public
    constructor Create(notesDocument: TNotesNote; aName: string); override;
    constructor CreateNew (notesDocument: TNotesNote; aName: string); override;
    constructor CreateNext (notesItem: TNotesItem); override;
    constructor CreateFromFile(aDocument: TNotesNote; aItemName: string; aFile: string);
    destructor Destroy; override;

    // Add content of given file to document. If context exists, saves it
    // Note that this function does file import producing new item with
    // the same name as this one
    procedure AddRtfFile(FileName: string);    //RTF
    procedure AddJpgFile(FileName: string);    //JPEG
    procedure AddBmpFile(FileName: string);    //BMP
    procedure AddFile(FileName: string);       //by extension

    // Context operations
    procedure CreateContext;  //creates or resets Rich-text context
    procedure CheckContext;   //checks the context existence and creates if neccessary
    procedure SaveContext;    //save context to field
    procedure AddPara;
    procedure AddText(Text: string);
    procedure AddTextPara(Text: string); //adds text with correct paragraphs instead of CRLFs
    {$IFDEF D4}
    procedure AddDocLink(Doc: TNotesDocument; Title: string = ' '; View: TNotesView = nil);   //View is optional
    procedure AddViewLink(View: TNotesView; Title: string = ' ');
    procedure AddDbLink(Database: TNotesDatabase; Title: string = ' ');
    {$ELSE}
    procedure AddDocLink(Doc: TNotesDocument; Title: string; View: TNotesView);
    procedure AddViewLink(View: TNotesView; Title: string);
    procedure AddDbLink(Database: TNotesDatabase; Title: string);
    {$ENDIF}
    procedure AddAnchorLink(Doc: TNotesDocument; Title, Anchor: string); //this one requires title name
    procedure AddUrl(aURL, aTitle: string);
    procedure AddPassThroughHtml(Html: string);
    // By A.Pape
    procedure AddFormulaButton(aCaption, aFormula: string);

    // Item enumeration and context filling routines
    procedure ReadItem(enumProc: TNotesRTFReadProc);
    procedure AddToContext(RecordPtr: pointer; RecordType: WORD; RecordLength: DWORD);
    procedure AddTextToContext(FontID: word; Text: string);

    // RTF sections
    // Use NOTES_COLOR_xxx for color values
    // Use ColorToNotes to convert from Delphi color
    procedure StartSection (aTitle: string; aColor: word;
                            aFlags: TRichTextSectionFlags;
                            aBorder: TRichTextSectionBorder);
    procedure EndSection;

    // Attachments support
    property AttachmentCount: integer read GetAttachmentCount;  //number of file attachments
    property Attachment[Index: integer]: string read GetAttachment;
    function FindAttachment(aName: string): integer;  //return index in Attachment or -1
    procedure Detach (Index: integer; FileName: string);
    {$IFDEF D4}
    procedure Attach(AName: string; fIcon: boolean = False); //attach a file by its name to the current context
    procedure AttachOleObject(AName: string; aType: TRichTextOle2AttachType; aHint: string = '');
      //attach a file as an OLE2 object
    {$ELSE}
    procedure Attach(AName: string; fIcon: boolean);         //attach a file by its name to the current context
    procedure AttachOleObject(AName: string; aType: TRichTextOle2AttachType; aHint: string);
    {$ENDIF}

    // Font properties
    property FontBold: boolean Index 0 read GetFontBoolProperty write SetFontBoolProperty default False;
    property FontColor: word read FFontColor write FFontColor default 0;  //Use NotesToColor to get Delphi color
    property FontItalic: boolean Index 1 read GetFontBoolProperty write SetFontBoolProperty default False;
    property FontUnderline: boolean Index 2 read GetFontBoolProperty write SetFontBoolProperty default False;
    property FontStrikeout: boolean Index 3 read GetFontBoolProperty write SetFontBoolProperty default False;
    property FontSuperScript: boolean Index 4 read GetFontBoolProperty write SetFontBoolProperty default False;
    property FontSubScript: boolean Index 5 read GetFontBoolProperty write SetFontBoolProperty default False;
    property FontShadow: boolean Index 6 read GetFontBoolProperty write SetFontBoolProperty default False;
    property FontEmboss: boolean Index 7 read GetFontBoolProperty write SetFontBoolProperty default False;
    property FontExtrude: boolean Index 8 read GetFontBoolProperty write SetFontBoolProperty default False;
    property FontFace: TRichTextFont read FFontFace write FFontFace default rfSwiss;
    property FontFaceName: string read FFontFaceName write FFontFaceName;
    property FontSize: integer read FFontSize write FFontSize default 10;

    // Paragraph properties
    property ParaJustification: TRichTextJustification read FParaJustification write SetParaJustification default rjLeft;
    property ParaAfterSpace: integer index 0 read GetParaIntProperty write SetParaIntProperty default DEFAULT_BELOW_PAR_SPACING;
    property ParaBeforeSpace: integer index 1 read GetParaIntProperty write SetParaIntProperty default DEFAULT_ABOVE_PAR_SPACING;
    property ParaFirstLeftMargin: integer index 2 read GetParaIntProperty write SetParaIntProperty default DEFAULT_FIRST_LEFT_MARGIN;
    property ParaLeftMargin: integer index 3 read GetParaIntProperty write SetParaIntProperty default DEFAULT_LEFT_MARGIN;
    property ParaRightMargin: integer index 4 read GetParaIntProperty write SetParaIntProperty default DEFAULT_RIGHT_MARGIN;
    property ParaLineSpacing: integer index 5 read GetParaIntProperty write SetParaIntProperty default DEFAULT_LINE_SPACING;
    property ParaStyleOptions: TRichTextStyleOptions read FParaStyleOptions write SetParaStyleOptions;
    procedure UpdateParaStyle;

    // RTF import/export
    property AsRichText;
    property PlainText: boolean read FPlainText write FPlainText default True;
    procedure ExportRtfFile(RtfFile: string);

    // Table creation
    procedure CreateTable(LeftMargin: word; Options: TRichTextTableOptions;
      HorizInterCellSpace: word; VertInterCellSpace: word);
    procedure AddCell(nRow, nCol: word; Cell: TRichTextCell);
    procedure EndTable;

    // Link support
    // ************** by Olaf **************
    property LinkCount : integer read GetLinkCount; //number of links
    property Link[Index: integer] : LinkDef read GetLink;
  end;

// RichText utils
{$IFNDEF D4}
function RichTextCell (
    LeftMargin, RightMargin: word;
    Options: TRichTextCellOptions;
    FractWidth: word;
    BorderLeft, BorderTop, BorderRight, BorderBottom: word;
    BackColor: word;
    bRowSpan: boolean;
    bColSpan: boolean
): TRichTextCell;
{$ELSE}
function RichTextCell (
    LeftMargin, RightMargin: word;
    Options: TRichTextCellOptions;
    FractWidth: word;
    BorderLeft: word = 1;
    BorderTop: word = 1;
    BorderRight: word = 1;
    BorderBottom: word = 1;
    BackColor: word = 0;
    bRowSpan: boolean = False;
    bColSpan: boolean = False
): TRichTextCell;
{$ENDIF}

// Version of HugeNsfItemAppend for RTF
procedure HugeNSFRTFItemAppend(hNote: NOTEHANDLE;
                              ItemFlags: Word;
                              Name: PChar;
                              NameLength: Word;
                              Value: Pointer;
                              ValueLength: DWord);


// Functions are provided for backward-compatibility only
// Use GetNotesExeDir, GetNotesDataDir instead
function GetActualNotesDir( var NotesDataDir: string ): string;
function AddSlashAtPathEnd( aFilePath: string; SlashAtEnd: boolean ): string;

// Filled during initialization
var
  NotesDir,
  NotesDataDir: string;

implementation
uses Registry, ShellAPI, Graphics, ActiveX, ComObj;

(******************************************************************************)
// By Andy
function AddSlashAtPathEnd( aFilePath: string; SlashAtEnd: boolean ): string;
begin
  Result := aFilePath;
  if SlashAtEnd then begin
    if (Result <> '') and (Result[length(Result)] <> '\') then appendStr(Result,'\');
  end
  else begin
    if (Result <> '') and (Result[length(Result)] = '\') then delete(Result,length(Result),1);
  end;
end;

(******************************************************************************)
// By Andy
function GetActualNotesDir( var NotesDataDir: string ): string;
var
  Key: string;
  Reg: TRegistry;
  KeyNames: TStringList;
begin
  Result      := '';
  NotesDataDir:= '';

  // Check whether we're in Notes directory
  Result := AddSlashAtPathEnd (ExtractFilePath(ParamStr(0)), True);
  if (FileExists( Result + 'nlnotes.exe' )) then begin
    NotesDataDir := GetNotesDataDir;
  end
  else begin
    Reg := TRegistry.create;
    KeyNames := TStringList.create;
    try
      Reg.RootKey := HKEY_LOCAL_MACHINE;
      Key := 'Software\Lotus\Notes';
      if Reg.OpenKey(Key,False) then begin
        Reg.GetKeyNames( KeyNames );
        if KeyNames.count > 0 then begin
          appendStr(Key, '\' + KeyNames[0]);
          Reg.CloseKey;
          if Reg.OpenKey(Key,False) then begin
            Result      := AddSlashAtPathEnd (Reg.ReadString( 'Path'     ), True);
            NotesDataDir:= AddSlashAtPathEnd (Reg.ReadString( 'DataPath' ), True);
          end;
        end;
      end;
    finally
      KeyNames.free;
      Reg.free;
    end;
  end;
end;

(******************************************************************************)
{function GetActualNotesDir (var NotesDataDir: string): string;
begin
  Result := GetNotesExeDir;
  NotesDataDir := GetNotesDataDir;
end;}

(******************************************************************************)
function GetFilterByName(Prefix, FilterName, Ext: string; var DllName, DllName2: string): boolean;
type
  tpart = array[1..5] of string;

procedure StrSplit(Str: string; var parts: tpart);
var
  i, n: integer;
begin
  for i := System.Low(parts) to System.High(parts) do parts[i] := '';
  for i := System.Low(parts) to System.High(parts) do begin
    n := Pos(',',str);
    if (n = 0) or (i = System.High(parts)) then begin
      parts[i] := str;
      break;
    end
    else begin
      parts[i] := trim(copy(str,1,n-1));
      delete(str,1,n);
    end;
  end;
end;

var
  i: integer;
  buf: string;
  parts: tpart;
begin
  Result := false;
  DllName := '';
  DllName2 := '';
  for i := 1 to 255 do begin
    buf := '';
    SetLength(buf, 256);
    OsGetEnvironmentString(pchar(Prefix + inttoStr(i)),pchar(buf), 255);
    buf := trim(strPas(pchar(buf)));
    //name,flag,dll name,dll name2,ext
    StrSplit(buf,parts);
    if FilterName <> '' then Result := compareText(parts[1], FilterName) = 0
    else if Ext <> '' then Result := Pos(upperCase(Ext),upperCase(parts[5])) > 0;
    if Result then begin
      DllName := parts[3];
      if (DllName <> '') and (DllName[1] = '_') then DllName[1] := 'N';
      DllName2 := parts[4];
      if (DllName2 <> '') and (DllName2[1] = '_') then DllName2[1] := 'N';

      // KOL - Workaround on W4W import issue (it's just don't work!)
      if (compareText(DllName, 'NIW4W') = 0) and (DllName2 <> '') then begin
        DllName := DllName2;
        DllName2 := '';
      end;
      break;
    end;
  end;
end;

(******************************************************************************)
function RichTextCell;
begin
  Result.LeftMargin   := LeftMargin;
  Result.RightMargin  := RightMargin;
  Result.Options      := Options;
  Result.Borders[1]   := BorderLeft;
  Result.Borders[2]   := BorderTop;
  Result.Borders[3]   := BorderRight;
  Result.Borders[4]   := BorderBottom;
  Result.BackColor    := BackColor;
  Result.FractWidth   := FractWidth;
  Result.bRowSpan     := bRowSpan;
  Result.bColSpan     := bColSpan;
end;

//***************************************************
function GetTextLen(var aText: string): dword;
begin
  Result := length(aText);
  if (Result mod 2) = 0 then begin
    appendStr(aText,'  ');
    inc(Result,2);
    aText[Result-1] := #0;
    aText[Result] := #0;
  end;
end;


//***************************************************
// NT ImportLib functions - by Andy
//***************************************************
const
  ntERRIMPORTLIB_NOERROR                = 0;
  ntERRIMPORTLIB_UNKNOW                 = -1;
  ntERRIMPORTLIB_DLLNOTFOUND            = -2;
  ntERRIMPORTLIB_GRAPHICNOTFOUND        = -3;
  ntERRIMPORTLIB_LOADFAILED             = -4;   { OSLoadLibrary failed. }
  ntERRIMPORTLIB_FUNCTIONNOTFOUND       = -5;
  ntERRIMPORTLIB_IMPORTERROR            = -6;
  ntERRIMPORTLIB_NULLHANDLE             = -7;
  ntERRIMPORTLIB_FILENOTFOUND           = -8;
  ntERRIMPORTLIB_OPENERROR              = -9;
  ntERRIMPORTLIB_PUTPARA                = -10;
  ntERRIMPORTLIB_PUTPABDEF              = -11;
  ntERRIMPORTLIB_PUTPABREF              = -12;
  ntERRIMPORTLIB_SEEKERROR              = -13;
  ntERRIMPORTLIB_READERROR              = -14;

  TAB_LEFT                              = 0;
  TAB_DEFAULT                           = TAB_LEFT;

type
  tpCDBuffer      = array[0..CD_BUFFER_LENGTH-1] of byte;

{ --------------------------------------------------------------------------- }
{ Puts a new paragraph a the beginning of the graphic file
{ --------------------------------------------------------------------------- }
function PutPara (  var   pbRTItem    : TpCDBuffer;
                          wLength     : WORD;
                    var   pwRTLength  : WORD ): boolean;
var
  pCDPara: CDPARAGRAPH;
begin
  { If not enough space in buffer for this paragraph, then exit.   }
  if (wLength < sizeof(CDPARAGRAPH))
  then  begin
          result := false;
          exit;
        end;

  { Fill in PARAGRAPH item structure  }
  pCDPara.Header.Length     := sizeof(CDPARAGRAPH);
  pCDPara.Header.Signature  := SIG_CD_PARAGRAPH;

  Move ( pCDPara, pbRTItem[pwRTLength], sizeof(CDPARAGRAPH) );

  { Adjust current record length, forcing to an even byte count.   }
  pwRTLength := pwRTLength + pCDPara.Header.Length;

  if ( pwRTLength mod 2) <> 0
  then inc ( pwRTLength);

  Result := true;
end;

{ --------------------------------------------------------------------------- }
{ Puts a new paragraph definition a the beginning of the graphic file
{ --------------------------------------------------------------------------- }
function PutPabDef (  var pbRTItem      : TpCDBuffer;
                          wPabDefNumber : WORD;
                          wLength       : WORD;
                      var pwRTLength    : WORD ): boolean;
var
  pcdPabDef: CDPABDEFINITION;  { style definition for this para }
begin
  { If not enough space in buffer for this paragraph, then exit.   }
  if (wLength < sizeof(CDPARAGRAPH))
  then  begin
          result := false;
          exit;
        end;

  { Fill in paragraph definition block.  We use all defaults.         }
  FillChar ( pcdPabDef, sizeof(CDPABDEFINITION), 0 );
  pcdPabDef.Header.Signature       := SIG_CD_PABDEFINITION;
  pcdPabDef.Header.Length          := sizeof(CDPABDEFINITION);
  pcdPabDef.PABID                  := wPabDefNumber;
  pcdPabDef.JustifyMode            := DEFAULT_JUSTIFICATION;
  pcdPabDef.LineSpacing            := DEFAULT_LINE_SPACING;
  pcdPabDef.ParagraphSpacingBefore := DEFAULT_ABOVE_PAR_SPACING;
  pcdPabDef.ParagraphSpacingAfter  := DEFAULT_BELOW_PAR_SPACING;
  pcdPabDef.LeftMargin             := DEFAULT_LEFT_MARGIN;
  pcdPabDef.RightMargin            := DEFAULT_RIGHT_MARGIN;
  pcdPabDef.FirstLineLeftMargin    := DEFAULT_FIRST_LEFT_MARGIN;
  pcdPabDef.Tabs                   := DEFAULT_TABS;
  pcdPabDef.Tab[0]                 := DEFAULT_TAB_INTERVAL;
  pcdPabDef.Flags                  := 0;
  pcdPabDef.TabTypes               := TAB_DEFAULT;

  Move ( pcdPabDef, pbRTItem[pwRTLength], sizeof( CDPABDEFINITION ) );

  { Adjust current record length, forcing to an even byte count.   }
  pwRTLength := pwRTLength + pcdPabDef.Header.Length;

  if ( pwRTLength mod 2) <> 0
  then inc ( pwRTLength);

  result := true;
end;

{ --------------------------------------------------------------------------- }
{ Puts a new paragraph reference a the beginning of the graphic file
{ --------------------------------------------------------------------------- }
function PutPabRef (  var pbRTItem      : TpCDBuffer;
                          wPabDefNumber : WORD;
                          wLength       : WORD;
                      var pwRTLength    : WORD ): boolean;
var
  pcdPabRef: CDPABREFERENCE;      { style definition for this para }
begin
  { If not enough space in buffer for this paragraph, then exit.   }
  if (wLength < sizeof(CDPARAGRAPH))
  then  begin
          result := false;
          exit;
        end;

  pcdPabRef.Header.Signature := SIG_CD_PABREFERENCE;
  pcdPabRef.Header.Length    := sizeof(CDPABREFERENCE);
  pcdPabRef.PABID            := wPabDefNumber;

  Move ( pcdPabRef, pbRTItem[pwRTLength], sizeof(CDPABREFERENCE) );

  { Adjust current record length, forcing to an even byte count.   }
  pwRTLength := pwRTLength + sizeof ( CDPABREFERENCE );

  if ( pwRTLength mod 2) <> 0
  then inc ( pwRTLength);

  result := true;
end;

{ --------------------------------------------------------------------------- }
{ Import the CD file into a notes document (notes item)
{ --------------------------------------------------------------------------- }
function ntImportLib_ImportCDFile ( hNote       : NOTEHANDLE ;
                                    pszCDFile   : string;
                                    pszItemName : string ): longint;
var
  bError          : BOOL;       { Returncode from PutPara, PutPabDef, PutPabRef }
  pCDBuffer       : TpCDBuffer; { Rich Text memory buffer                       }
  dwCDRecordLength: DWORD;      { Length of current CD record                   }
  lCombinedLength : DWORD;
  dwCDBufferLength: DWORD;      { Length of current CD buffer                   }
  wTemp           : WORD;       { Temporary length                              }
  bTemp           : BYTE;       { Temporary length                              }
  lLength         : DWORD;      { Length of current read buffer                 }
  wReadLength     : WORD;       { Length of current read buffer                 }
  longpos         : DWORD;      { Initialy seek past TYPE_COMPOSITE             }
                                {      at start of file                         }
  ltmpItemLength  : DWORD;
  wItemLength     : WORD;       { Index for buffer manipulation                 }
  bFlag           : BOOL;       { termination flag                              }
  CDFileFD        : integer;
  Position        : word;
  nRead           : DWORD;
begin
  // handle valid?
  if ( hNote = NULLHANDLE )
  then  begin
          Result := ntERRIMPORTLIB_NULLHANDLE;
          exit;
        end;

  // does the temp exists?
  if ( not ( FileExists ( pszCDFile ) ) )
  then  begin
          Result := ntERRIMPORTLIB_FILENOTFOUND;
          exit;
        end;

  dwCDBufferLength := CD_BUFFER_LENGTH;
  bFlag           := false;
  lLength         := 0;

  Fillchar ( pCDBuffer, sizeOf(pCDBuffer), 0 );

  // open the temp. file
  CDFileFD := FileOpen( pszCDFile, fmOpenRead );
  if CDFileFD < 0
  then  begin
          Result := ntERRIMPORTLIB_OPENERROR;
          exit;
        end;

 { Set start length to zero  }

  wItemLength := 0;

  {  Put a paragraph record in buffer.  }

  bError := PutPara(pCDBuffer, dwCDBufferLength, wItemLength );

  if (bError = FALSE)
  then  begin
          FileClose ( CDFileFD );
          Result := ntERRIMPORTLIB_PUTPARA;
          exit;
        end;

  { Setup a pabdef }

  bError := PutPabDef(pCDBuffer,1, dwCDBufferLength - wItemLength, wItemLength );

  { Leave if error returned...   }

  if (bError = FALSE)
  then  begin
          FileClose ( CDFileFD );
          Result := ntERRIMPORTLIB_PUTPABDEF;
          exit;
        end;

  { Now add a pabref }

  bError := PutPabRef(pCDBuffer, 1, dwCDBufferLength - wItemLength, wItemLength );

  { Leave if error returned...    }

  if (bError = FALSE)
  then  begin
          FileClose ( CDFileFD );
          Result := ntERRIMPORTLIB_PUTPABREF;
          exit;
        end;


  { Keep on writing items until entire cd file hase been appended   }

  longpos := 0;

  while (bFlag = FALSE) do
    begin
      { Seek file to end of previous CD record   }

      nRead := FileSeek ( CDFileFD, longpos, 0 );

      if ( nRead <> longPos )
      then  begin
              { Leave if error returned... }
              FileClose ( CDFileFD );
              Result := ntERRIMPORTLIB_SEEKERROR;
              exit;
            end;

      { Read the contents of the file into memory  }
      wReadLength := FileRead(  CDFileFD, pCDBuffer[wItemLength],
                                dwCDBufferLength - wItemLength );

      { check for error    }
      if (wReadLength = $ffff )
      then  begin
              { Leave if error returned...    }
              FileClose ( CDFileFD );
              Result := ntERRIMPORTLIB_READERROR;
              exit;
            end;

      { See whether the contents will fit in current item....  }

      if (wReadLength < CD_HIGH_WATER_MARK)
      then  begin
              { we can fit what is left in a single buffer and leave  }
              bFlag       := TRUE;
              wItemLength := wItemLength + wReadLength;
            end
      else  begin
               {
               * Parse the buffer one CD record at a time, adding up the lengths
               * of the CD records.  When the length approaches CD_HIGH_WATER_MARK,
               * append the buffer to the note.  Set the file pointer to the first
               * record not parsed, read from the temp file into the buffer again,
               * and repeat until end of temp file.
               *
               * All CD records begin with a signature word that indicates its
               * type and record length.  The low order byte is the type, and
               * the high order byte is the length.  If the indicated length is 0,
               * then the next DWORD (32 bits) contains the record length.  If the
               * indicated length is 0xff, the next WORD (16 bits) contains the
               * recordlength.   Else, then the high order BYTE, itself, contains
               * the record length.
               }

              dwCDRecordLength := 0 ;
              lLength          := 0 ;
              lCombinedLength  := wItemLength + dwCDRecordLength ;

              while (lCombinedLength < CD_HIGH_WATER_MARK) do
                begin
                  // store the current position
                  Position := wItemLength;
                  // inc the current position, because the first byte contains
                  // the record type
                  Inc(Position);

                  { find length of CD record. }
                  if ( pCDBuffer[Position] = 0 )  { record length is a DWORD  }
                  then  begin
                          Inc(Position);
                          Move ( pCDBuffer[Position], dwCDRecordLength, sizeOf(DWORD) );
                        end
                  else  begin
                          if ( pCDBuffer[Position] = $FF )   { record length is a WORD  }
                          then  begin
                                  Inc(Position);
                                  Move ( pCDBuffer[Position], wTemp, sizeOf(WORD) );
                                  dwCDRecordLength := wTemp;
                                end
                          else  begin   { record length is the BYTE }
                                  Move ( pCDBuffer[Position], bTemp, sizeOf(BYTE) );
                                  dwCDRecordLength := bTemp;
                                end;
                        end;

                  if (dwCDRecordLength mod 2) <> 0 then Inc(dwCDRecordLength);

                  lLength         := lLength + dwCDRecordLength;
                  ltmpItemLength  := wItemLength + dwCDRecordLength;

                  if (ltmpItemLength < CD_BUFFER_LENGTH)
                  then  wItemLength := wItemLength + dwCDRecordLength;

                  lCombinedLength := ltmpItemLength + dwCDRecordLength;
                end;
            end;

      if (wItemLength > 0)
      then  begin
              Result := NSFItemAppend(hNote,
                                 0,
                                 PChar(pszItemName),
                                 length(pszItemName),
                                 TYPE_COMPOSITE,
                                 @pCDBuffer,
                                 wItemLength);
              if Result <> 0 then begin
                FileClose ( CDFileFD );
                exit;
              end;
          end;

      longpos     := longpos + lLength;
      wItemLength := 0;

      Fillchar ( pCDBuffer, sizeOf(pCDBuffer), 0 );
    end;

    FileClose ( CDFileFD );

    Result := ntERRIMPORTLIB_NOERROR;
end;

//*************************************************************************
// appends a large richtext-item to a document, if needed the item is split
// up into several smaler parts, this routine preserves the unity of one
// cd_record
// by Olaf Hahnl
//*************************************************************************
procedure HugeNSFRTFItemAppend(hNote: NOTEHANDLE;
                              ItemFlags: Word;
                              Name: PChar;
                              NameLength: Word;
                              Value: Pointer;
                              ValueLength: DWord);
var
  curPtr,
  tmpPtr : pointer;
  position : dword;

  RecordType,
  FixedSize : Word;
  RecordLength : DWord;

begin
  // if length fits into one item append it
  if ValueLength <= CD_HIGH_WATER_MARK then
    CheckError (NSFItemAppend(hNote, ItemFlags, Name, NameLength,
                              TYPE_COMPOSITE, Value, ValueLength))
  // else split up the whole item into several parts
  else begin
    // start of buffer
    CurPtr := Value;
    tmpPtr := CurPtr;
    // do until length fits into one last item
    while ValueLength > CD_HIGH_WATER_MARK do begin
      position := 0;

      // add cd_records to buffer until critical mark reached
      while position < CD_HIGH_WATER_MARK do begin

        copymemory(@RecordType, CurPtr, sizeof(Word));
        // which type of record is it, calculate the size
        case (RecordType and $FF00) of
          LONGRECORDLENGTH: begin
                              RecordLength := LSIG(CurPtr^).Length;
              FixedSize := sizeof(LSIG);
                            end;
          WORDRECORDLENGTH: begin
                              RecordLength := WSIG(CurPtr^).Length;
            FixedSize := sizeof(WSIG);
                            end;
      else begin
            RecordLength := DWORD ((RecordType shr 8) and $00FF);
            RecordType := RecordType and $00FF; // Length not part of signature */
            FixedSize := sizeof(BSIG);
          end;
        end;

        // step the pointer forward to the next CD_Record
        if RecordLength <> 0 then begin
          CurPtr := Pointer(PChar(CurPtr) + RecordLength);
          inc(position, RecordLength);
        end else begin
          CurPtr := Pointer(PChar(CurPtr) + FixedSize);
          inc(position, FixedSize);
        end;

        // new CD_Record always starts at an even address
        if (DWORD (CurPtr) mod 2) <> 0 then begin
          CurPtr := Pointer(PChar(CurPtr)+1);
          inc(position);
        end;

      end;

      // decrease length to be procesed
      dec(ValueLength,position);
      // append this part of the whole item
      CheckError (NSFItemAppend(hNote, ItemFlags, Name, NameLength,
                  TYPE_COMPOSITE, tmpPtr, position));
      tmpPtr := CurPtr;
    end;
    // append the last part of the whole item
    CheckError (NSFItemAppend(hNote, ItemFlags, Name, NameLength,
                TYPE_COMPOSITE, tmpPtr, ValueLength));
  end;
end;

function getTempFileName: string;
var
  dir: string;
begin
  setLength(dir, MAX_PATH);
  setLength(Result, MAX_PATH);
  GetTempPath(MAX_PATH, pchar(dir));
  Windows.GetTempFileName(pchar(dir),'ndlib',0,pchar(Result));
  Result := strPas(pchar(Result));
end;

(******************************************************************************)
// returns the temporary path to a bitmap-file which is the icon to be shown
// in the Notes-Client - By Olaf
function getAppIcon(FileName : String) : String;
var
  Reg : TRegistry;

  // tries to get the "DefaultIcon"-Entry from the current path
  // opened in the registry (includes recursive searching)
  function FindDefaultIcon(Path : String) : String;
  var Sl : TStringList;
       i : word;
  begin
    Result := '';
    Sl := TStringList.Create;
    Reg.GetKeyNames(Sl);
    If SL.Count > 0 Then
      For i := 0 to Sl.Count-1 Do Begin
        If Reg.OpenKeyReadOnly(Sl.Strings[i]) Then Begin
          If Sl.Strings[i] = 'DefaultIcon' Then Begin
            Result := Reg.ReadString('');
            Break;
          End Else Begin
            Result := FindDefaultIcon(Reg.Currentpath);
            If Result <> '' Then Break;
          End;
          Reg.CloseKey;
          Reg.OpenKeyReadOnly(Path);
        End;
      End;
    Sl.Free;
    Reg.CloseKey;
    Reg.OpenKeyReadOnly(Path);
  End;

var
  Height,
  Width    : Integer;
  App    : String;
  Icon   : TIcon;
  Bmp    : TBitmap;
  Index  : Word;

Begin
  Icon := TIcon.Create;
  Reg := TRegistry.Create;
  Bmp := TBitmap.Create;
  Result := '';
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;
    If Reg.OpenKeyReadOnly(ExtractFileExt(Filename)) Then Begin
      app := FindDefaultIcon(Reg.CurrentPath);
      If app='' then begin
        app := Reg.ReadString('');
        Reg.CloseKey;
        If Reg.OpenKeyReadOnly(app) Then app := FindDefaultIcon(Reg.currentpath);
      end;
      If app <> '' then
        Icon.Handle := ExtractIcon(hInstance,PChar(copy(app,1,pos(',',App)-1)),
                                   strtoint(copy(app,pos(',',App)+1,5)))
      else Icon.Handle := ExtractAssociatedIcon(hInstance,PChar(FileName),Index);
    end
    else if FileExists(FileName) Then Begin
      Index := 0;
      Icon.Handle := ExtractAssociatedIcon(hInstance,PChar(FileName),Index);
    end
    else begin
      // No registry key and file doesn't exist - cannot get an icon!
      Bmp.free;
      Reg.free;
      Icon.free;
      exit;
    end;

    FileName := extractfilename(FileName);

    width := Bmp.Canvas.TextWidth(FileName);
    height := Bmp.Canvas.TextHeight(FileName);
    if width < Icon.Width then width := Icon.Width;

    Bmp.Width := width+4;
    Bmp.Height := Height+Icon.Height+4;
    Bmp.Canvas.Draw((width-Icon.Width) Div 2,2,Icon);
    Bmp.Canvas.TextOut(2,Icon.Height+2,FileName);
    // needed to avoid wrong colors in Notes
    Bmp.PixelFormat := pf8Bit;

    Result := getTempFileName;
    Bmp.SaveToFile(Result);
  finally
    Bmp.free;
    Reg.free;
    Icon.free;
  end;
end;

(********************* By Olaf *************************************************)
// get the parts that follow a link eg comment, server and anchor text
Procedure getVariablePart(RecordPtr : pchar;RecordLength : dword;
                          Var sComment, sHint, sAnchor : String);
Var  Comment,Hint,Anchor : PChar;
                Len,Len2 : Word;
Begin
  Len := RecordLength;
  If Len > 0 Then Begin
    Comment := stralloc(len+1);
    CopyMemory(Comment,RecordPtr,Len);
    sComment := String(Comment);
    Len2 := Length(sComment)+1;
    StrDispose(Comment);
    If Len2 < Len Then Begin
      Dec(Len,Len2);
      Hint := stralloc(len+1);
      CopyMemory(Hint,RecordPtr+Len2,Len);
      sHint := String(Hint);
      StrDispose(Hint);
      // special test, used for hotspotlinks I think
      Len2 := Length(sHint)+1;
      If (Len2=1) or ((Len2>=3) And (sHint[1]='C') And (sHint[2]='N'))
      or ((Len2>=2) And (sHint[1]='C')) Then Begin
        If Len2 < Len Then Begin
          Dec(Len,Len2);
          Anchor := stralloc(len+1);
          CopyMemory(Anchor,RecordPtr+length(sComment)+length(sHint)+2,Len);
          sAnchor := String(Anchor);
          StrDispose(Anchor);
        End;
      end
      else begin
        sHint := '';
        sAnchor := '';
      end;
    end;
  end
  else begin
    sComment := '';
    sAnchor := '';
    sHint := '';
  end;
end;

//***************************************************
// Internal class to support attachments
//***************************************************
type
  TAttachInfo = class
    blockitem: BLOCKID;     //BLOCKID of file attachment
    fileName: string;       //original file name
    alias: string;          //alias
    objectName: string;     //object name of OLE attachment
    progID: string;         //progID of OLE attachment
  end;


//***************************************************
// TNotesRichTextItem
//***************************************************
constructor TNotesRichTextItem.Create(notesDocument: TNotesNote; aName: string);
begin
  inherited Create (notesDocument, aName);
  if ItemType <> TYPE_COMPOSITE then
    raise ELotusNotes.create ('This item is not Rich-text: ' + aName);

  CreateDefaults;
  LoadAttachmentInfo;
  LoadLinkInfo;
end;

//***************************************************
constructor TNotesRichTextItem.CreateFromFile;
var
  name2: string;
begin
  name2 := Native2Lmbcs(aItemName);
  NsfItemDelete (aDocument.Handle, pchar(name2), length(name2));  //just in case...

  inherited CreateNew (aDocument, aItemName);

  CreateDefaults;
  AddFile(aFile);
end;

//***************************************************
constructor TNotesRichTextItem.CreateNew (notesDocument: TNotesNote; aName: string);
begin
  inherited CreateNew (notesDocument, aName);
  ItemType := TYPE_COMPOSITE;
  CreateDefaults;
end;

//***************************************************
constructor TNotesRichTextItem.CreateNext(notesItem: TNotesItem);
begin
  inherited CreateNext(notesItem);
  CreateDefaults;
  LoadAttachmentInfo;
  LoadLinkInfo;
end;

//**********************************************
procedure TNotesRichTextItem.CreateDefaults;
begin
  IsSummary := False;
  FAttach := TStringList.create;
  FPostedAttachments := TStringList.create;
  ParaAfterSpace := DEFAULT_ABOVE_PAR_SPACING;
  ParaBeforeSpace := DEFAULT_BELOW_PAR_SPACING;
  ParaFirstLeftMargin := DEFAULT_FIRST_LEFT_MARGIN;
  ParaLeftMargin := DEFAULT_LEFT_MARGIN;
  ParaRightMargin := DEFAULT_RIGHT_MARGIN;
  ParaLineSpacing := DEFAULT_LINE_SPACING;
  FontFace := rfSwiss;
  FontSize := 10;
  FPlainText := True;
  FParaStyleOptions := [];
  FLinks := TList.Create;
end;

//**********************************************
destructor TNotesRichTextItem.Destroy;
var
  i: integer;
begin
  if FContext <> nil then FreeMem (FContext);
  if FAttach <> nil then
    for i := 0 to FAttach.count-1 do FAttach.Objects[i].free;
  FAttach.free;
  FPostedAttachments.free;
  FStringsValue.free;
  if FLinks <> nil then begin
    for i := 0 to FLinks.count-1 do Dispose(PLinkDef(FLinks.Items[i]));
    FLinks.Free;
  end;
  inherited Destroy;
end;

//***************************************************
function FileEnumProc(RecordPtr: pchar; RecordType: word; RecordLength: dword; vContext: pointer): STATUS; stdcall;
type
  PCDHOTSPOTBEGIN = ^CDHOTSPOTBEGIN;
var
  PRec: PCDHOTSPOTBEGIN;
  pc: pchar;
  s1, s2: string;
begin
  Result := NOERROR;
  if RecordType = SIG_CD_HOTSPOTBEGIN then begin
    PRec := PCDHOTSPOTBEGIN(RecordPtr);
    if PRec^.aType = HOTSPOTREC_TYPE_FILE then begin
      pc := pchar (dword(PRec) + sizeOf(CDHOTSPOTBEGIN));
      s1 := strPas(pc);                                  //alias
      s2 := strPas(pchar (dword(pc) + length(s1)+1));    //file name
      TStringList(vContext).add(s1 + '=' + s2);
    end;
  end;
end;

//***************************************************
function OleEnumProc(RecordPtr: pchar; RecordType: word; RecordLength: dword; vContext: pointer): STATUS; stdcall;
type
  PCDOLEBEGIN = ^CDOLEBEGIN;
var
  PRec: PCDOLEBEGIN;
  pc: pchar;
  s1, s2: string;
begin
  Result := NOERROR;
  if RecordType = SIG_CD_OLEBEGIN then begin
    PRec := PCDOLEBEGIN(RecordPtr);
    pc := pchar(dword(prec) + sizeof(CDOLEBEGIN));
    s1 := NotesToString(pc, prec^.AttachNameLength);
    pc := pchar(dword(pc) + prec^.AttachNameLength);
    s2 := NotesToString(pc, prec^.ClassNameLength);
    TStringList(vContext).add(s1 + '=' + s2);
  end;
end;

//***************************************************
procedure TNotesRichTextItem.LoadAttachmentInfo;
type
  PFILEOBJECT = ^FILEOBJECT;
var
  bhItem, bhValue: BLOCKID;
  szValue: dword;
  err: STATUS;
  Obj: PFileObject;
  s1, s2: string;
  pc: pchar;
  Files, Objects: TStringList;
  atcItem: TAttachInfo;
  i: integer;
begin
  FAttach.clear;
  Files := TStringList.create;
  Objects := TStringList.create;
  try
    // Process all items sequentally in order to handle all attachments
    err := NOERROR;
    bhItem := ItemBid;
    bhValue := ValueBid;
    szValue := ValueLength;
    s1 := Native2Lmbcs(Name);
    while err = NOERROR do begin
      CheckError(EnumCompositeBuffer(bhValue, szValue, FileEnumProc, Files));
      CheckError(EnumCompositeBuffer(bhValue, szValue, OleEnumProc, Objects));
      err := NSFItemInfoNext(Document.Handle, bhItem, PChar(s1), Length(s1),
        @bhItem, nil, @bhValue, @szValue);
    end;

    // Loops through all $FILE items and find block ids
    Err := NSFItemInfo (Document.Handle, PChar(ITEM_NAME_ATTACHMENT), Length(ITEM_NAME_ATTACHMENT),
      @bhItem, nil, @bhValue, @szValue);
    while Err = NOERROR do begin
      Obj := PFileObject(dword(OsLockBlock(bhValue)) + sizeOf(WORD));
      try
        pc := pchar (longint (Obj) + sizeOf(FILEOBJECT));
        s1 := NotesToString(pc, Obj^.FileNameLength); //alias
        s2 := Files.Values[s1];   //file name
        if s2 <> '' then begin
          // OK, we found value handle
          atcItem := TAttachInfo.Create;
          atcItem.BlockItem := bhItem;
          atcItem.fileName := s2;
          atcItem.alias := s1;
          FAttach.AddObject(AnsiUpperCase(Lmbcs2Native(s2)), atcItem);
        end;
      finally
        OsUnlockBlock (bhValue);
      end;
      Err := NSFItemInfoNext(Document.Handle, bhItem, PChar(ITEM_NAME_ATTACHMENT), Length(ITEM_NAME_ATTACHMENT),
        @bhItem, nil, @bhValue, @szValue);
    end;

    // Process objects
    for i := 0 to Objects.count-1 do begin
      atcItem := TAttachInfo.Create;
      atcItem.objectName := Objects.Names[i];
      atcItem.fileName := atcItem.objectName + '.file';
      atcItem.progID := Objects.Values[atcItem.objectName];
      FAttach.AddObject(AnsiUpperCase(atcItem.fileName), atcItem);
    end;
  finally
    Files.free;
    Objects.free;
  end;
end;

//***************************************************
function TNotesRichTextItem.GetAttachmentCount;
begin
  Result := FAttach.count;
end;

//***************************************************
function TNotesRichTextItem.GetAttachment;
begin
  Result := FAttach[Index];
end;

//***************************************************
function TNotesRichTextItem.FindAttachment;
begin
  Result := FAttach.indexOf(upperCase(aName));
end;

//**********************************************
procedure TNotesRichTextItem.Attach;
var
  MemSz, NameSz: dword;
  sFileName, sAlias, tmpFileName: string;
  pBegin: CDHOTSPOTBEGIN;
  pEnd: CDHOTSPOTEND;
  cText: CDTEXT;
  pData: pchar;
begin
  if not FileExists (aName) then
    raise ELotusNotes.Create ('File ' + aName + ' does not exist');

  // Get name and alias
  sFileName := Native2Lmbcs(extractFileName(aName));
  sAlias := format ('file%.4d', [Document.FMaxAttachment]);
  inc(Document.FMaxAttachment);

  // Create hotspot begin record
  CheckContext;
  NameSz := length(sFileName) + length(sAlias) + 2;
  MemSz := ODSLength(_CDHOTSPOTBEGIN) + NameSz;
  AddMem (CheckOdd (MemSz));
  pBegin.Header.Signature := SIG_CD_HOTSPOTBEGIN;
  pBegin.Header.Length := MemSz;
  pBegin.aType := HOTSPOTREC_TYPE_FILE;
  pBegin.Flags := HOTSPOTREC_RUNFLAG_BEGIN {or HOTSPOTREC_RUNFLAG_NOBORDER};
  pBegin.DataLength := NameSz;
  ODSWriteMemory (@FCurPtr, _CDHOTSPOTBEGIN, @pBegin, 1);
  Move (sFileName[1], FCurPtr^, NameSz);

  // Add alias file name
  pData := pchar (FCurPtr);
  strCopy (pData, pchar(sAlias));
  pData := pchar (dword(pData) + length(sAlias) + 1);
  strCopy (pData, pchar(sFileName));
  FCurPtr := pointer (dword(FCurPtr) + CheckOdd(NameSz));

  If fIcon then begin
    // attach file with its icon as representation
    tmpFileName := getAppIcon(aName);

    // Add icon to indicate attachment
    if tmpFileName <> '' then begin
      AddBmpFile(tmpFileName);
      DeleteFile (PChar(tmpFileName));
    end;
  end;
  if (not fIcon) or (tmpFileName = '') then begin
    // Add text with attachment name
    sFileName := '<' + sFileName + '>';
    NameSz := CheckOdd(length(sFileName));
    Memsz := ODSLength( _CDTEXT ) + NameSz;
    AddMem (CheckOdd (Memsz));
    ctext.Header.Signature := SIG_CD_TEXT;
    ctext.Header.Length := MemSz;
    ctext.FontID := 0;
    FontIDSetFaceID (ctext.FontID, FONT_FACE_TYPEWRITER);
    FontIDSetSize (ctext.FontID, 10);
    FontIDSetStyle (ctext.FontID, CF_ISUNDERLINE);
    FontIDSetColor (ctext.FontID, 4);
    ODSWriteMemory (@FCurPtr, _CDTEXT, @ctext, 1 );
    Move (sFileName[1], FCurPtr^, NameSz);
    FCurPtr := pointer (longint (FCurPtr) + CheckOdd(NameSz));
  end;

  // Create hotspot end record
  MemSz := ODSLength(_CDHOTSPOTEND);
  AddMem (CheckOdd(MemSz));
  pEnd.Header.Signature := SIG_CD_HOTSPOTEND;
  pEnd.Header.Length := MemSz;
  ODSWriteMemory (@FCurPtr, _CDHOTSPOTEND, @pEnd, 1);

  // Mark required attachment
  FPostedAttachments.add(aName + '=' + sAlias);
end;

//**********************************************
procedure TNotesRichTextItem.AttachFiles;
var
  i: integer;
  sName, sAlias: string;
begin
  for i := 0 to FPostedAttachments.count-1 do begin
    sName := FPostedAttachments.Names[i];
    sAlias := FPostedAttachments.Values[sName];
    if not FileExists (sName) then
      raise ELotusNotes.Create ('File ' + sName + ' does not exist');
    CheckError (NsfNoteAttachFile(Document.Handle, ITEM_NAME_ATTACHMENT, length(ITEM_NAME_ATTACHMENT),
      pchar(Native2Lmbcs(sName)), pchar(sAlias), EFLAGS_INDOC or HOST_LOCAL or COMPRESS_HUFF));
  end;
  FPostedAttachments.clear;
end;

//**********************************************
procedure TNotesRichTextItem.Detach(Index: integer; FileName: string);
var
  info: TAttachInfo;
  pStg, pSubStg, pRootStg: IStorage;
  pObj: IUnknown;
  ppStorage: IPersistStorage;
  wfn, wsn, wtmpfn: WideString;
  clsid: TGUID;
  tmpFile: string;
begin
  if FileName = '' then FileName := FAttach[Index];
  info := TAttachInfo(FAttach.Objects[Index]);
  if FileName = '' then FileName := info.fileName;
  FileName := Native2Lmbcs(FileName);

  if info.objectName = '' then begin
    // File attachment
    CheckError(NSFNoteExtractFile(Document.Handle, info.blockitem, pchar(fileName), nil));
  end
  else begin
    // OLE Attachment
    tmpFile := getTempFileName;
    CheckError(NSFNoteExtractOLE2Object(Document.Handle, pchar(info.objectName), pchar(tmpFile), nil, true, 0));
    try
      wfn := WideString(fileName);
      wsn := WideString(info.progID);
      wtmpfn := WideString(tmpFile);
      if StgIsStorageFile(pwidechar(wtmpfn)) = S_OK then begin
        // Open the root storage and then object storage (named by object progID)
        OleCheck(StgOpenStorage(pwidechar(wtmpfn), nil, STGM_DIRECT or STGM_READ or STGM_SHARE_DENY_WRITE, nil, 0, pStg));
        OleCheck(pStg.OpenStorage(pwidechar(wsn), nil, STGM_READ or STGM_SHARE_EXCLUSIVE, nil, 0, pSubStg));

        // Load data into the object
        CLSIDFromProgID(pwidechar(wsn), clsid);
        OleCheck(CoCreateInstance(clsid, nil, CLSCTX_ALL, IUnknown, pObj));
        ppStorage := pObj as IPersistStorage;
        ppStorage.Load(pSubStg);

        // Save as root storage
        OleCheck(StgCreateDocfile(pwidechar(wfn), STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_CREATE,0,pRootStg));
        OleCheck(OleSave(ppStorage, pRootStg, false));
        OleCheck(ppStorage.SaveCompleted(nil));
      end;
    finally
      ppStorage := nil;
      pSubStg := nil;
      pStg := nil;
      DeleteFile(pchar(tmpFile));
    end;
  end;
end;

//**********************************************
procedure TNotesRichTextItem.AddStyle;
const
  OptionsMap: array [TRichTextStyleOption] of word = (
    PABFLAG_PAGINATE_BEFORE,
    PABFLAG_KEEP_WITH_NEXT,
    PABFLAG_KEEP_TOGETHER,
    PABFLAG_PROPAGATE,
    PABFLAG_HIDE_RO,
    PABFLAG_HIDE_RW,
    PABFLAG_HIDE_PR,
    PABFLAG_DISPLAY_RM,
    PABFLAG_HIDE_CO,
    PABFLAG_BULLET,
    PABFLAG_HIDE_IF,
    PABFLAG_NUMBEREDLIST,
    PABFLAG_HIDE_PV,
    PABFLAG_HIDE_PVE,
    PABFLAG_HIDE_NOTES
  );
var
  opt: TRichTextStyleOption;
  val: word;
begin
  val := 0;
  for opt := System.Low(TRichTextStyleOption) to System.High(TRichTextStyleOption) do
    if opt in ParaStyleOptions then val := val or OptionsMap[opt];
  AddStyle2(val);
end;

//**********************************************
procedure TNotesRichTextItem.AddStyle2;
const
  JustifyMap: array [TRichTextJustification] of word = (
    JUSTIFY_NONE, JUSTIFY_LEFT, JUSTIFY_CENTER, JUSTIFY_RIGHT, JUSTIFY_BLOCK
  );
var
  pabdef: CDPABDEFINITION;
  ref: CDPABREFERENCE;
begin
  AddMem (CheckOdd (ODSLength(_CDPABDEFINITION)));
  FillChar(pabdef, sizeOf(pabdef), 0);
  FillChar(ref, sizeOf(ref), 0);

  // Add style definition
  inc (FCurStyle);
  pabdef.Header.Signature := SIG_CD_PABDEFINITION;
  pabdef.Header.Length := ODSLength(_CDPABDEFINITION);
  pabdef.PABID := FCurStyle;
  pabdef.JustifyMode := JustifyMap[ParaJustification];
  pabdef.LineSpacing := ParaLineSpacing;
  pabdef.ParagraphSpacingBefore := ParaBeforeSpace;
  pabdef.ParagraphSpacingAfter := ParaAfterSpace;

  // Changed by AP - 1.11.99
  if ((ParaLeftMargin = ONEINCH) and
     ((Flags and PABFLAG_BULLET      = PABFLAG_BULLET) or
     ( Flags and PABFLAG_NUMBEREDLIST = PABFLAG_NUMBEREDLIST)))
      then pabdef.LeftMargin := 1800
      else pabdef.LeftMargin := ParaLeftMargin;
  pabdef.RightMargin := ParaRightMargin;
  if ((ParaFirstLeftMargin = ONEINCH ) and
       ((Flags and PABFLAG_BULLET       = PABFLAG_BULLET) or
       ( Flags and PABFLAG_NUMBEREDLIST = PABFLAG_NUMBEREDLIST)))
        then pabdef.FirstLineLeftMargin := 1800
        else pabdef.FirstLineLeftMargin := ParaFirstLeftMargin;

  pabdef.Tabs := DEFAULT_TABS;
  pabdef.Tab[0] := trunc(DEFAULT_TAB_INTERVAL);
  pabdef.Flags := Flags;

  { Call ODSWriteMemory to convert the CDPABDEFINITION structure to
    Notes canonical format and write the converted structure into
    the buffer at location buff_ptr. This advances buff_ptr to the
    next byte in the buffer after the canonical format strucure.
   }
  ODSWriteMemory (@FCurPtr, _CDPABDEFINITION, @pabdef, 1 );
  fNewStyle := False;

  // Add style tag
  AddMem (CheckOdd (ODSLength(_CDPABREFERENCE)));
  ref.Header.Signature := SIG_CD_PABREFERENCE;
  ref.Header.Length := byte (ODSLength(_CDPABREFERENCE));
  ref.PABID := FCurStyle;
  ODSWriteMemory(@FCurPtr, _CDPABREFERENCE, @ref, 1 );
end;

//**********************************************
procedure TNotesRichTextItem.AddPara;
var
  para: CDPARAGRAPH;
begin
  // Add paragraph mark
  fillChar(para, sizeOf(para), 0);
  AddMem (CheckOdd (ODSLength (_CDPARAGRAPH)));
  para.Header.Signature := SIG_CD_PARAGRAPH;
  para.Header.Length := ODSLength(_CDPARAGRAPH);
  ODSWriteMemory (@FCurPtr, _CDPARAGRAPH, @para, 1 );

  // Set style for next para
  if fNewStyle then AddStyle;
end;

//**********************************************
procedure TNotesRichTextItem.CreateContext;
begin
  // Create buffer enought for one paragraph definition
  if FContext <> nil then FreeMem (FContext);
  FContext := nil;
  FCurPtr := FContext;
  FLength := 0;
  FPostedAttachments.clear;
  fNewStyle := True;
  fHasPara := False;
  FFontTableChanged := False;
  AddStyle;
end;

//**********************************************
procedure TNotesRichTextItem.SaveContext;
var
  s: string;
begin
  // Add empty string to normalize text
  s := FontFaceName;
  FontFaceName := '';
  AddText ('');
  FontFaceName := s;

  // Attach files
  AttachFiles;

  // Save context
  s := Native2Lmbcs(Name);

  //Version 3.7
  //HugeNsfItemAppend(Document.Handle, 0, pchar(s), length(s),
  //  TYPE_COMPOSITE, FContext, FLength);

  HugeNsfRtfItemAppend(Document.Handle, 0, pchar(s), length(s),
    FContext, FLength);
  if FFontTableChanged then Document.SaveFontTable;

  // Clear
  if FContext <> nil then FreeMem (FContext);
  FContext := nil;
  FCurPtr := nil;
  FLength := 0;
end;

//**********************************************
function ReadItemProc (RecordPtr: pchar;
                        RecordType: WORD;
                        RecordLength: DWORD;
                        vContext: pointer): STATUS; stdcall;
begin
  if (vContext <> nil) and assigned(TNotesRichTextItem(vContext).FReadProc) then
    Result := TNotesRichTextItem(vContext).FReadProc(TNotesRichTextItem(vContext),RecordPtr,
      RecordType,RecordLength)
  else
    Result := NOERROR;
end;

//**********************************************
procedure TNotesRichTextItem.ReadItem(enumProc: TNotesRTFReadProc);
begin
  FReadProc := EnumProc;
  EnumCompositeBuffer(ValueBid,ValueLength,@ReadItemProc,self);
end;

//**********************************************
procedure TNotesRichTextItem.AddToContext(RecordPtr: pointer;
  RecordType: WORD; RecordLength: DWORD);
begin
  CheckContext;
  AddMem(CheckOdd(RecordLength));
  Move(RecordPtr^,FCurPtr^,RecordLength);
  FCurPtr := pointer(dword(FCurPtr) + CheckOdd(RecordLength));
end;

//**********************************************
procedure TNotesRichTextItem.AddTextToContext;
var
  ctext: CDTEXT;
  szt,sz: integer;
begin
  CheckContext;
  Document.LoadFontTable;

  Text := Native2Lmbcs(Text);
  szt := length(Text);
  sz := ODSLength(_CDTEXT ) + szt;
  AddMem (CheckOdd (sz));

  fillChar(ctext,sizeOf(ctext),0);
  ctext.Header.Signature := SIG_CD_TEXT;
  ctext.Header.Length := sz;
  ctext.FontID := FontID;

  ODSWriteMemory (@FCurPtr, _CDTEXT, @ctext, 1);

  // Add text
  if Text <> '' then begin
    Move (Text[1], FCurPtr^, szt);
    FCurPtr := pointer (dword(FCurPtr) + CheckOdd(szt));
  end;
end;

//**********************************************
procedure TNotesRichTextItem.AddText;
const
  StyleMap: array [0..8] of integer = (
    CF_ISBOLD, CF_ISITALIC, CF_ISUNDERLINE, CF_ISSTRIKEOUT,
    CF_ISSUPER, CF_ISSUB, CF_ISSHADOW, CF_ISEMBOSS, CF_ISEXTRUDE
  );
  FontMap: array [TRichTextFont] of integer = (
    FONT_FACE_ROMAN, FONT_FACE_SWISS, FONT_FACE_TYPEWRITER
  );
var
  ctext: CDTEXT;
  szt, sz: dword;
  fstyle: dword;
  i: integer;
begin
  CheckContext;
  if fNewStyle then AddStyle;

  // Add CDTEXT record
  Text := Native2Lmbcs(Text);
  szt := length(Text);
  sz := ODSLength(_CDTEXT ) + szt;
  AddMem (CheckOdd (sz));
  fillChar(ctext,sizeOf(ctext),0);
  ctext.Header.Signature := SIG_CD_TEXT;
  ctext.Header.Length := sz;
  ctext.FontID := 0;

  fstyle := 0;
  for i := System.Low(FFontBoolProperties) to System.High(FFontBoolProperties) do
    if FFontBoolProperties[i] then fStyle := fStyle or StyleMap[i];
  FontIDSetStyle (ctext.FontID, fStyle);
  FontIDSetSize (ctext.FontID, FontSize);
  FontIDSetColor (ctext.FontID, FontColor);

  if FontFaceName = '' then begin
     // Static font
    FontIDSetFaceId (ctext.FontID, FontMap[FontFace]);
  end
  else begin
    // Check for presence in Document Fonttable
    Document.LoadFontTable;
    fstyle := $ff;
    for i := 0 to Document.FontTable.count-1 do begin
      fstyle := dword (Document.FontTable.Objects[i]);
      if (compareStr (FontFaceName, Document.FontTable[i]) = 0) and
      ((fstyle and (not $FF)) = ctext.FontID) then begin
        // target font exist
        ctext.FontID := fstyle;
        fstyle := 0;
        break;
      end;
    end;
    if fstyle <> 0 then begin
      // Font not found
      inc(Document.FMaxFontID);
      if Document.FMaxFontID >= 255 then raise ELotusNotes.createErr(-1,'Too many fonts defined');
      FontIDSetFaceId (ctext.FontId,Document.FMaxFontID);
      Document.FontTable.AddObject (FontFaceName, pointer(ctext.FontID));
      FFontTableChanged := True;
    end;
  end;

  CheckContext;
  ODSWriteMemory (@FCurPtr, _CDTEXT, @ctext, 1);

  // Add text
  if Text <> '' then begin
    Move (Text[1], FCurPtr^, szt);
    FCurPtr := pointer (dword(FCurPtr) + CheckOdd(szt));
  end;
end;

//**********************************************
procedure TNotesRichTextItem.AddTextPara;
var
  n: integer;
begin
  n := Pos (#13#10,Text);
  while n > 0 do begin
    AddText(copy(Text,1,n-1));
    AddPara;
    Delete(Text,1,n+1);
    n := Pos (#13#10,Text);
  end;
  if Text <> '' then AddText(Text);
end;

//**********************************************
procedure TNotesRichTextItem.AddLink;
var
  CdLink: CDLINKEXPORT2;
  sz: dword;
begin
  CheckContext;
  Title := Native2Lmbcs( Title ) + #0 + #0 + #0;
  fillChar(CdLink,sizeOf(CdLink),0);
  sz := ODSLength(_CDLINKEXPORT2);
  if TitleLen = 0 then TitleLen := Length(Title) + 3;
  inc(sz, TitleLen);
  AddMem(CheckOdd(sz));
  FillChar(CdLink, sizeOf(CdLink), 0);
  CdLink.Header.Signature := SIG_CD_LINKEXPORT2;
  CdLink.Header.Length := sz;
  CdLink.NoteLink.aFile := DatabaseID;
  CdLink.NoteLink.View := ViewID;
  CdLink.NoteLink.Note := DocID;
  ODSWriteMemory (@FCurPtr, _CDLINKEXPORT2, @CdLink, 1);
  if Title <> '' then begin
    Move (Title[1], FCurPtr^, CheckOdd(TitleLen));
    FCurPtr := pointer (dword(FCurPtr) + CheckOdd(TitleLen));
  end;
end;

//**********************************************
procedure TNotesRichTextItem.AddDocLink;
begin
  if View <> nil
    then AddLink(Doc.Database.DatabaseID, View.UniversalID, Doc.UniversalID, Title, 0)
    else AddLink(Doc.Database.DatabaseID, BlankUNID, Doc.UniversalID, Title, 0);
end;

//**********************************************
procedure TNotesRichTextItem.AddViewLink;
begin
  AddLink (View.Database.DatabaseID, View.UniversalID, BlankUNID, Title, 0);
end;

//**********************************************
procedure TNotesRichTextItem.AddDbLink;
begin
  AddLink (Database.DatabaseID, BlankUNID, BlankUNID, Title, 0);
end;

//**********************************************
procedure TNotesRichTextItem.AddAnchorLink;
var
  buf: string;
  sz: integer;
begin
  buf := Title;
  sz := Length(Title) + Length(Anchor) + 2;
  setLength(buf, sz);
  strCopy (@(buf[Length(Title)+1]), PChar(Anchor));
  AddLink(Doc.Database.DatabaseID, BlankUNID, Doc.UniversalID, Title, sz);
end;

//**********************************************
procedure TNotesRichTextItem.AddURL;
type
  PCDHOTSPOTBEGIN = ^CDHOTSPOTBEGIN;
  PCDHOTSPOTEND = ^CDHOTSPOTEND;
var
  MemSz, NameSz: dword;
  pBegin: CDHOTSPOTBEGIN;
  pEnd: CDHOTSPOTEND;
  cText: CDTEXT;
begin
  // Create hotspot begin record
  if aURL = '' then exit;
  aTitle:= Native2Lmbcs(aTitle);
  CheckContext;

  NameSz := GetTextLen(aURL);
  MemSz := ODSLength(_CDHOTSPOTBEGIN) + NameSz;
  AddMem (CheckOdd (MemSz));
  fillChar(pBegin,sizeOf(pBegin),0);
  pBegin.Header.Signature := SIG_CD_HOTSPOTBEGIN;
  pBegin.Header.Length := CheckOdd(MemSz);
  pBegin.aType :=  HOTSPOTREC_TYPE_HOTLINK;
  pBegin.Flags := HOTSPOTREC_RUNFLAG_INOTES or HOTSPOTREC_RUNFLAG_BEGIN or
    HOTSPOTREC_RUNFLAG_END;
  pBegin.DataLength := NameSz;
  ODSWriteMemory (@FCurPtr, _CDHOTSPOTBEGIN, @pBegin, 1);
  Move (aURL[1], FCurPtr^, NameSz);
  FCurPtr := pointer (dword(FCurPtr) + CheckOdd(NameSz));

  // Add text
  if aTitle = '' then aTitle := aURL;
  NameSz := length(aTitle);
  Memsz := ODSLength( _CDTEXT ) + CheckOdd(NameSz);
  AddMem (CheckOdd (Memsz));
  fillChar(ctext,sizeOf(ctext),0);
  ctext.Header.Signature := SIG_CD_TEXT;
  ctext.Header.Length := MemSz;
  ctext.FontID := 0;
  ODSWriteMemory (@FCurPtr, _CDTEXT, @ctext, 1 );
  Move (aTitle[1], FCurPtr^, NameSz);
  FCurPtr := pointer (dword(FCurPtr) + CheckOdd(NameSz));

  // Create hotspot end record
  MemSz := ODSLength(_CDHOTSPOTEND);
  AddMem (CheckOdd(MemSz));
  fillChar(pEnd,sizeOf(pEnd),0);
  pEnd.Header.Signature := SIG_CD_HOTSPOTEND;
  pEnd.Header.Length := MemSz;
  ODSWriteMemory (@FCurPtr, _CDHOTSPOTEND, @pEnd, 1);
end;

//**********************************************
procedure TNotesRichTextItem.AddMem;
var
  diff: dword;
begin
  diff := dword(FCurPtr) - dword(FContext);
  inc (FLength, sz);
  ReallocMem (FContext, FLength);
  FCurPtr := pointer (dword(FContext) + diff);
  FillChar (FCurPtr^, sz, #0);
end;

//***************************************************
procedure TNotesRichTextItem.CheckContext;
begin
  if FContext = nil then CreateContext;
end;

//**********************************************
function TNotesRichTextItem.CheckOdd;
begin
  if (sz mod 2) = 0 then Result := sz else Result := sz+1;
end;

//**********************************************
function TNotesRichTextItem.GetParaIntProperty;
begin
  Result := FParaIntProperties[Index];
end;

//**********************************************
procedure TNotesRichTextItem.SetParaIntProperty;
begin
  if Value <> FParaIntProperties[Index] then begin
    FParaIntProperties[Index] := Value;
    fNewStyle := True;
  end;
end;

//**********************************************
procedure TNotesRichTextItem.SetParaJustification (Value: TRichTextJustification);
begin
  if (FParaJustification <> Value) then begin
    FParaJustification := Value;
    fNewStyle := True;
  end;
end;

//**********************************************
procedure TNotesRichTextItem.SetParaStyleOptions;
begin
  if (ParaStyleOptions <> Value) then begin
    FParaStyleOptions := Value;
    fNewStyle := True;
  end;
end;

//**********************************************
function TNotesRichTextItem.GetFontBoolProperty;
begin
  Result := FFontBoolProperties[Index];
end;

//**********************************************
procedure TNotesRichTextItem.SetFontBoolProperty;
begin
  if (FFontBoolProperties[Index] <> Value) then begin
    FFontBoolProperties[Index] := Value;
  end;
end;

//**********************************************
function TNotesRichTextItem.GetRichText;
var
  TmpRtfFile: string;
begin
  if PlainText then begin
    Result := inherited GetRichText;
    exit;
  end;

  // Save the text to temporary RTF file
  TmpRtfFile := getTempFileName;
  try
    ExportRtfFile(TmpRtfFile);

    // Load to internal buffer
    if FStringsValue = nil then FStringsValue := TStringList.create;
    FStringsValue.LoadFromFile (TmpRtfFile);
    Result := FStringsValue;
  finally
    DeleteFile (pchar(TmpRtfFile));
  end;
end;

//**********************************************
procedure TNotesRichTextItem.SetRichText;
var
  TmpRtfFile: string;
begin
  if PlainText then begin
    inherited SetRichText (Value);
    exit;
  end;

  // Save the text to temporary RTF file
  TmpRtfFile := getTempFileName;
  try
    Value.SaveToFile (TmpRtfFile);
    AddRtfFile(TmpRtfFile);
  finally
    DeleteFile (pchar(TmpRtfFile));
  end;
end;

//***************************************************
procedure TNotesRichTextItem.ExportRtfFile;
var
  TmpFile: string;
  hf: THandle;
  Ptr: pointer;
  EditExportData: TEDITEXPORTDATA;
  ProcAddress: FARPROC;
  err: STATUS;
  DllName, DllName2: string;
  hmod: HMODULE;
begin
  // Get temporary CD file name
  TmpFile := getTempFileName;

  Ptr := OsLockBlock(ValueBid);
  try
    // Write item content to temporary CD file
    hf := FileCreate(TmpFile);
    try
      FileWrite(hf, Ptr^, ValueLength);
    finally
      FileClose(hf);
    end;

    // Run export filter
    GetFilterByName('EDITEXP', 'MicrosoftWord RTF', '', DllName, DllName2);
    if DllName = '' then GetFilterByName('EDITEXP', 'Microsoft RTF', '', DllName, DllName2);
    if DllName = '' then DllName := 'NXRTF.DLL';
    ProcAddress := nil;
    hmod := 0;
    CheckError (OSLoadLibrary(pchar(DllName), 0, hmod, ProcAddress));
    if not assigned(ProcAddress) then raise ELotusNotes.createErr (-1, 'Wrong export DLL');

    try
      FillChar (EditExportData, sizeOf(EditExportData), 0);
      with EditExportData do begin
        StrCopy (InputFileName, pchar(TmpFile));
        HeaderBuffer.Desc.Font := Default_Font_ID;
        FooterBuffer.Desc.Font := Default_Font_ID;
        PrintSettings.Flags := PS_Initialized;
      end;
      err := TLnExportProc(ProcAddress)(
          @EditExportData,
          IXFLAG_FIRST or IXFLAG_LAST,        //* Both 1st and last import */
          0,                                  //* Use default hmodule      */
          pchar(DllName2),                    //* 2nd DLL, if needed.      */
          pchar(RtfFile));                    //* File to import.          */
      CheckError(err);
    finally
      OSFreeLibrary(hmod);
    end;
  finally
    OsUnlockBlock(ValueBid);
    DeleteFile(pchar(TmpFile));
  end;
end;

(*//**********************************************
//This routine seems not to work always correct! Don't know why!
function TNotesRichTextItem.CreateCDFile;
 //(FileName, FilterTyp, FilterName, DLLName: string): string;
var
  TmpPath: string;
  hmod: HMODULE;
  EditImportData: TEDITIMPORTDATA;  //* Import DLL data structure      */
  ProcAddress: FARPROC;
  err: STATUS;
  ModuleName: string;
begin
  Result := getTempFileName;
  try
    // Import file to temporary 'CD' format by standard Notes import filter
    // We assume that the input file is always RTF
    ModuleName := GetFilterByName (FilterTyp, FilterName);
    if( ModuleName = '' )
    then ModuleName := DLLName;
    ProcAddress := nil;
    hmod := 0;
    CheckError (OSLoadLibrary (pchar(ModuleName), 0, hmod, ProcAddress));
    if not assigned(ProcAddress) then raise ELotusNotes.createErr (-1, 'Wrong import DLL');

    try
      FillChar (EditImportData, sizeOf(EditImportData), #0);
      StrCopy (EditImportData.OutputFileName, pchar(Result));
      EditImportData.FontID := DEFAULT_FONT_ID;
      err := TLnImportProc(ProcAddress)(
          @EditImportData,
          IXFLAG_FIRST or IXFLAG_LAST,
          0,
          nil,
          pchar(FileName));
      CheckError(err);
    finally
      OSFreeLibrary(hmod);
    end;
  except
    Result:= '';
    raise;
    end;
end;
*)

(*//********************************************
procedure TNotesRichTextItem.LoadFile;
var
  TmpFile: string;
  Name2: string;
  hCompound: LHANDLE;
  //Style: COMPOUNDSTYLE;
  //dwStyle: dword;
begin
  try
    TmpFile:= CreateCDFile (FileName, FilterTyp, FilterName, DLLName);

    // Now create compound text buffer and import the file into it
    Name2 := Native2Lmbcs(Name);
    NsfItemDelete(Document.Handle, pchar(Name2), length(Name2));
    CheckError (CompoundTextCreate(Document.Handle, pchar(Name), @hCompound));
    try
      //CompoundTextInitStyle (@Style);
      //CheckError (CompoundTextDefineStyle(hCompound, 'Normal', @Style, @dwStyle));
      //CheckError(CompoundTextAddParagraph(hCompound, dwStyle, Default_Font_ID, '', 0, 0));
      CheckError(CompoundTextAssimilateFile(hCompound, pchar(TmpFile), 0));
      CheckError(CompoundTextClose(hCompound, nil, nil, nil, 0));
    except
      CompoundTextDiscard(hCompound);
      raise;
    end;
  finally
    DeleteFile (pchar(TmpFile));
  end;
end;
*)

//**********************************************
procedure TNotesRichTextItem.IntAddFile;
var
  data: TEDITIMPORTDATA;
  aStr: string;
  ContextExists: boolean;
  ProcAddress: FARPROC;
  hmod: LHandle;
begin
  // Load library
  if DllName = '' then raise ELotusNotes.createErr(-1,'Filter was not found for file ' + fileName);
  if extractFilePath(DllName) = '' then DllName := NotesDir + DllName;
  if (DllName2 <> '') and (extractFilePath(DllName2) = '') then DllName2 := NotesDir + DllName2;

  hmod := 0;
  CheckError(OSLoadLibrary(pchar(DLLName), 0, hmod, ProcAddress));
  if ProcAddress = nil then raise ELotusNotes.CreateErr(-1, 'Invalid import DLL');

  // Prepare temp.file and import structure
  FillChar(data, sizeof(data), #0);
  try
    data.FontID := DEFAULT_FONT_ID;
    StrPCopy(data.OutputFileName, getTempFileName);

    // Convert file to temporary CD
    CheckError(IXENTRYPROC(ProcAddress)(data, IXFLAG_FIRST and IXFLAG_LAST, 0,
                                        pchar(DllName2), pchar(FileName)));

    // Release context
    aStr:= Native2Lmbcs(Name);
    ContextExists:= ( FContext <> nil );
    if ContextExists then SaveContext;

    // Import temporary file
    ntImportLib_ImportCDFile( Document.Handle, data.OutputFileName, aStr );
  finally
    OSFreeLibrary (hMod);
    if data.OutputFileName[1] <> #0 then DeleteFile(data.OutputFileName);
  end;
end;

//**********************************************
procedure TNotesRichTextItem.AddRtfFile;
var
  fn, fn2: string;
begin
  GetFilterByName('EDITIMP', 'MicrosoftWord RTF', '', fn, fn2);
  if fn = '' then GetFilterByName('EDITIMP', 'Microsoft RTF', '', fn, fn2);
  IntAddFile(FileName, fn, fn2);
end;

//**********************************************
procedure TNotesRichTextItem.AddJpgFile( FileName : string );
var
  fn, fn2: string;
begin
  GetFilterByName('EDITIMP', '', '.jpg', fn, fn2);
  IntAddFile(FileName, fn, fn2);
end;

//**********************************************
procedure TNotesRichTextItem.AddBmpFile( FileName : string );
var
  fn, fn2: string;
begin
  GetFilterByName('EDITIMP', '', '.bmp', fn, fn2);
  IntAddFile(FileName, fn, fn2);
end;

//**********************************************
procedure TNotesRichTextItem.AddFile;
var
  fn, fn2: string;
begin
  GetFilterByName('EDITIMP', '', extractFileExt(FileName), fn, fn2);
  IntAddFile(FileName,fn, fn2);
end;

//**********************************************
procedure TNotesRichTextItem.CreateTable;
const
  OptionsMap: array [TRichTextTableOption] of word = (
    CDTABLE_AUTO_CELL_WIDTH,
    CDTABLE_3D_BORDER_EMBOSS,
    CDTABLE_3D_BORDER_EXTRUDE
  );
var
  rec: CDTABLEBEGIN;
  opt: TRichTextTableOption;
begin
  CheckContext;
  FillChar(rec,sizeOf(rec),0);
  rec.Header.Signature := SIG_CD_TABLEBEGIN;
  rec.Header.Length := ODSLength(_CDTABLEBEGIN);
  rec.LeftMargin := LeftMargin;
  rec.HorizInterCellSpace := HorizInterCellSpace;
  rec.VertInterCellSpace := VertInterCellSpace;

  // by Olaf Hahnl
  if (rtBorderEmboss in Options) or (rtBorderExtrude in Options) then begin
    // For these borders V4... vars must be set to 0
    rec.V4HorizInterCellSpace := 0;
    rec.V4VertInterCellSpace := 0;
  end
  else begin
    rec.V4HorizInterCellSpace := HorizInterCellSpace;
    rec.V4VertInterCellSpace := VertInterCellSpace;
  end;
  rec.Flags := CDTABLE_V4_BORDERS;
  for opt := System.Low(TRichTextTableOption) to System.High(TRichTextTableOption) do
    if opt in Options then rec.Flags := rec.Flags or OptionsMap[opt];
  AddMem(CheckOdd(rec.Header.Length));
  ODSWriteMemory (@FCurPtr, _CDTABLEBEGIN, @rec, 1);
end;

//**********************************************
procedure TNotesRichTextItem.AddCell;
const
  OptionsMap: array [TRichTextCellOption] of word  = (
    CDTABLECELL_USE_BKGCOLOR,
    CDTABLECELL_INVISIBLEH,
    CDTABLECELL_INVISIBLEV
  );
var
  rec: CDTABLECELL;
  opt: TRichTextCellOption;
begin
  FillChar(rec,sizeOf(rec),0);
  rec.Header.Signature := SIG_CD_TABLECELL;
  rec.Header.Length := ODSLength(_CDTABLECELL);
  rec.Row := nRow;
  rec.Column := nCol;
  rec.LeftMargin := Cell.LeftMargin;
  rec.RightMargin := Cell.RightMargin;
  rec.FractionalWidth := Cell.FractWidth;
  rec.RowSpan := ord(Cell.bRowSpan);
  rec.ColumnSpan := ord(Cell.bColSpan);
  rec.BackgroundColor := Cell.BackColor;

  rec.Border := rec.Border or ((Cell.Borders[1] shl CDTC_S_Left) and CDTC_M_Left);
  rec.Border := rec.Border or ((Cell.Borders[2] shl CDTC_S_Top) and CDTC_M_Top);
  rec.Border := rec.Border or ((Cell.Borders[3] shl CDTC_S_Right) and CDTC_M_Right);
  rec.Border := rec.Border or ((Cell.Borders[4] shl CDTC_S_Bottom) and CDTC_M_Bottom);

  rec.v42Border := rec.v42Border or ((Cell.Borders[1] shl CDTC_S_V42_Left) and CDTC_M_V42_Left);
  rec.v42Border := rec.v42Border or ((Cell.Borders[2] shl CDTC_S_V42_Top) and CDTC_M_V42_Top);
  rec.v42Border := rec.v42Border or ((Cell.Borders[3] shl CDTC_S_V42_Right) and CDTC_M_V42_Right);
  rec.v42Border := rec.v42Border or ((Cell.Borders[4] shl CDTC_S_V42_Bottom) and CDTC_M_V42_Bottom);

  rec.Flags := CDTABLECELL_USE_V42BORDERS;
  for opt := System.Low(TRichTextCellOption) to System.High(TRichTextCellOption) do
    if opt in Cell.Options then rec.Flags := rec.Flags or OptionsMap[opt];

  AddMem(CheckOdd(rec.Header.Length));
  ODSWriteMemory (@FCurPtr, _CDTABLECELL, @rec, 1);

  // Initialize a style
  ParaStyleOptions := ParaStyleOptions + [rsDisplayRM];
  ParaLeftMargin := Cell.LeftMargin;
  ParaRightMargin := Cell.RightMargin;
  AddPara;
  //AddStyle2(PABFLAG_DISPLAY_RM);
end;

//**********************************************
procedure TNotesRichTextItem.EndTable;
var
  rec: CDTABLEEND;
begin
  FillChar(rec,sizeOf(rec),0);
  rec.Header.Signature := SIG_CD_TABLEEND;
  rec.Header.Length := ODSLength(_CDTABLEEND);
  AddMem(CheckOdd(rec.Header.Length));
  ODSWriteMemory (@FCurPtr, _CDTABLEEND, @rec, 1);
  AddText('');
  AddPara;
end;

(******************************************************************************)
procedure TNotesRichTextItem.UpdateParaStyle;
begin
  if fNewStyle then AddStyle;
end;

(******************************************************************************)
procedure TNotesRichTextItem.StartSection;
const
  FlagsMap: array [TRichTextSectionFlag] of word = (
    BARREC_DISABLED_FOR_NON_EDITORS,
    BARREC_BORDER_INVISIBLE,
    BARREC_HIDE_EXPANDED,
    BARREC_AUTO_EXP_READ,
    BARREC_AUTO_EXP_PRE,
    BARREC_AUTO_EXP_EDIT,
    BARREC_AUTO_EXP_PRINT,
    BARREC_AUTO_COL_READ,
    BARREC_AUTO_COL_PRE,
    BARREC_AUTO_COL_EDIT,
    BARREC_AUTO_COL_PRINT,
    BARREC_EXPANDED,
    BARREC_IS_DISABLED
  );
  BorderMap: array [TRichTextSectionBorder] of word = (
    BARREC_BORDER_SHADOW,
    BARREC_BORDER_NONE,
    BARREC_BORDER_SINGLE,
    BARREC_BORDER_DOUBLE,
    BARREC_BORDER_TRIPLE,
    BARREC_BORDER_TWOLINE
  );

var
  pBegin: CDHOTSPOTBEGIN;
  pBar:   CDBAR;
  pText:  CDTEXT;
  NameSz: dword;
  flag:   TRichTextSectionFlag;
begin
  CheckContext;
  fillChar(pBegin,sizeOf(pBegin),0);
  fillChar(pBar,sizeOf(pBar),0);
  fillChar(pText, sizeOf(pText), 0);

  pBegin.Header.Signature := SIG_CD_V4HOTSPOTBEGIN;
  pBegin.Header.Length := ODSLength(_CDHOTSPOTBEGIN);
  pBegin.aType :=  HOTSPOTREC_TYPE_BUNDLE;
  pBegin.Flags := HOTSPOTREC_RUNFLAG_BEGIN or HOTSPOTREC_RUNFLAG_NOBORDER;
  pBegin.DataLength := 0;
  AddMem(CheckOdd(pBegin.Header.Length));
  ODSWriteMemory (@FCurPtr, _CDHOTSPOTBEGIN, @pBegin, 1);

  NameSz := GetTextLen(aTitle);
  pBar.Header.Signature := SIG_CD_BAR;
  pBar.Header.Length := ODSLength(_CDBAR) + CheckOdd(NameSz) + ODSLength(_WORD);
  pBar.Flags := BARREC_HAS_COLOR;
  for flag := System.Low(TRichTextSectionFlag) to System.High(TRichTextSectionFlag) do
    if flag in aFlags then pBar.Flags := pBar.Flags or FlagsMap[flag];
  SetBorderType (pBar.Flags, BorderMap[aBorder]);

  AddMem(CheckOdd(pBar.Header.Length));
  ODSWriteMemory (@FCurPtr, _CDBAR, @pBar, 1);
  ODSWriteMemory (@FCurPtr, _WORD, @aColor, 1);

  Move (aTitle[1], FCurPtr^, NameSz);
  FCurPtr := pointer (dword(FCurPtr) + CheckOdd(NameSz));

  pText.Header.Signature := SIG_CD_TEXT;
  pText.Header.Length    := ODSLength(_CDTEXT);
  AddMem(CheckOdd(pText.Header.Length));
  ODSWriteMemory (@FCurPtr, _CDTEXT, @pText, 1);

  fNewStyle := True;
  AddPara;
end;

(******************************************************************************)
procedure TNotesRichTextItem.EndSection;
var
  pEnd:   CDHOTSPOTEND;
  MemSz:  dword;
begin
  fillChar(pEnd,sizeOf(pEnd),0);
  MemSz := ODSLength(_CDHOTSPOTEND);
  AddMem (CheckOdd (MemSz));
  pEnd.Header.Signature := SIG_CD_V4HOTSPOTEND;
  pEnd.Header.Length := MemSz;
  ODSWriteMemory (@FCurPtr, _CDHOTSPOTEND, @pEnd, 1);
end;


const maxlinks = 5000;      // maximum count of Links Field ($Links)

//********************* Link By Olaf *****************
Type
  LinkDefI = packed record
    // Mirrors LinkDef from LnApi
    aFile : TIMEDATE; // File's replica ID
    View : UNID;      // View's Note Creation TIMEDATE
    Note : UNID;      // Note's Creation TIMEDATE
    Comment : String; // comment of doclink
    Hint : String;    // server
    Anchor : String;  // anchor text
    LinkType : TLinkType;
    //***
    Index : Integer;  // index in field ($Links)
  end;
  PLinkDef = ^LinkDefI;

// callback function to parse a rtf looking for links
function RTIEnumProcLink (RecordPtr: pchar; RecordType: word; RecordLength: dword; vContext: pointer): STATUS; stdcall;
type
  PCDLINK2 = ^CDLINK2;
  PCDLINKEXPORT2 = ^CDLINKEXPORT2;

var
  IRec : PCDLINK2;
  ERec : PCDLINKEXPORT2;
  Comment,
  Hint,
  Anchor : string;
  LInfo  : PLinkDef;

begin
  Result := NOERROR;
  if RecordType = SIG_CD_LINK2 then begin
    IRec := PCDLINK2(RecordPtr);

    getVariablePart(RecordPtr+SizeOf(CDLink2),RecordLength-SizeOf(CDLink2),Comment,Hint,Anchor);

    new(LInfo);
    LInfo^.Comment := Lmbcs2Native(Comment);
    LInfo^.Hint := Lmbcs2Native(Hint);
    LInfo^.Anchor := Lmbcs2Native(Anchor);
    If Anchor <> '' Then LInfo^.LinkType := rtlAnchorLink
    Else LInfo^.LinkType := rtlUnknown;
    LInfo^.Index := IRec^.LinkID;
    TList(vContext).Add(linfo);
  End Else If RecordType = SIG_CD_LINKEXPORT2 Then Begin
    ERec := PCDLINKEXPORT2(RecordPtr);

    getVariablePart(RecordPtr+SizeOf(CDLinkExport2), RecordLength-SizeOf(CDLinkExport2),
                    Comment, Hint, Anchor);

    new(LInfo);
    LInfo^.Comment := Lmbcs2Native(Comment);
    LInfo^.Hint := Lmbcs2Native(Hint);
    LInfo^.Anchor := Lmbcs2Native(Anchor);

    If Anchor <> '' Then LInfo^.LinkType := rtlAnchorLink
    Else LInfo^.LinkType := rtlUnknown;

    LInfo^.aFile := ERec^.NoteLink.aFile;
    LInfo^.View := ERec^.NoteLink.View;
    LInfo^.Note := ERec^.NoteLink.Note;
    LInfo^.Index := -1;
    TList(vContext).Add(linfo);
  end;
end;

(******************************************************************************)
// definitions needed to navigate in the system-field ($Links)
type
   NListe = record
        d : word;
    Liste : List;
   end;
   PNListe = ^NListe;

   NEntry = record
        L : NListe;
        N : array[0..maxlinks] of NoteLink;
   end;
   PNEntry = ^NEntry;

(******************************************************************************)
procedure TNotesRichTextItem.LoadLinkInfo;

function UNIDEmpty(u : UNID): Boolean;
begin
  Result:=((U.aFile.T1=0) and (U.aFile.T2=0) and (U.Note.T1=0) and (U.Note.T2=0));
end;

Var
  bhItem, bhValue: BLOCKID;
  szValue: dword;

  plinksvalue : ^byte;
        liste : PNListe;
            i : Word;
        Entry : PNEntry;
        linfo : PLinkDef;
       SEntry : NoteLink;

Begin
  // Parse original item
  EnumCompositeBuffer(ValueBid, ValueLength, RTIEnumProcLink, FLinks);

  // Parse other items with the same name
  bhItem := ItemBid;
  while NSFItemInfoNext(Document.Handle,bhItem,PChar(Name),length(Name),
                        @bhItem, nil, @bhValue, @szValue) = NoError do
     EnumCompositeBuffer (bhValue, szValue, RTIEnumProcLink, FLinks);

  if FLinks.Count > 0 then begin
    if NSFItemInfo(Document.Handle, PChar(ITEM_NAME_LINK), Length(ITEM_NAME_LINK),
                   @bhItem, nil, @bhValue, @szValue) = NoError then begin

      pLinksValue:=OSLockBlock(bhValue);
      liste:=PNListe(pLinksValue);
      entry:=PNEntry(pLinksValue);

      if Liste^.Liste.ListEntries = 0 then OSUnlockBlock(bhValue);

      for i := 0 to FLinks.Count-1 do begin

        LInfo := PLinkDef(FLinks.Items[i]);

        if LInfo.Index <> -1 then begin

          if LInfo.LinkType = rtlUnknown then begin
            if Not UNIDEmpty(entry^.N[LInfo.Index].Note) then LInfo.LinkType := rtlDocumentLink
                else if UNIDEmpty(entry^.N[LInfo.Index].View) then LInfo.LinkType := rtlDatabaseLink
                else LInfo.LinkType := rtlViewLink;
          end;

          SEntry := entry^.n[LInfo.Index];
          LInfo^.aFile := SEntry.aFile;
          LInfo^.View := SEntry.View;
          LInfo^.Note := SEntry.Note;
        end else begin
          if LInfo.LinkType = rtlUnknown Then Begin
            if Not UNIDEmpty(LInfo^.Note) then LInfo.LinkType := rtlDocumentLink
            else if UNIDEmpty(LInfo^.View) then LInfo.LinkType := rtlDatabaseLink
            else LInfo.LinkType := rtlViewLink;
          end;
        end;
      end;
      if Liste^.Liste.ListEntries <> 0 Then OSUnlockBlock(bhValue);
    end
    else begin
      for i := 0 to FLinks.Count-1 Do Begin
        LInfo := PLinkDef(FLinks.Items[i]);
        if LInfo.Index = -1 then begin

          if LInfo.LinkType = rtlUnknown then begin
            if Not UNIDEmpty(LInfo^.Note) then LInfo.LinkType := rtlDocumentLink
            else if UNIDEmpty(LInfo^.View) then LInfo.LinkType := rtlDatabaseLink
            else LInfo.LinkType := rtlViewLink;
          end;
        end;
      end;
    end;
  end;
end;

(******************************************************************************)
function TNotesRichTextItem.GetLink(Index : integer) : LinkDef;
var
  tmp : LinkDef;
begin
  if index < FLinks.Count then begin
    tmp.aFile := PLinkDef(FLinks[Index])^.aFile;
    tmp.View := PLinkDef(FLinks[Index])^.View;
    tmp.Note := PLinkDef(FLinks[Index])^.Note;
    tmp.Comment := PLinkDef(FLinks[Index])^.Comment;
    tmp.Hint := PLinkDef(FLinks[Index])^.Hint;
    tmp.Anchor := PLinkDef(FLinks[Index])^.Anchor;
    tmp.LinkType := PLinkDef(FLinks[Index])^.LinkType;
  end
  else begin
    tmp.aFile.T1 := 0;
    tmp.aFile.T2 := 0;
    tmp.View := BlankUNID;
    tmp.Note := BlankUNID;
    tmp.Comment := '';
    tmp.Anchor := '';
    tmp.LinkType := rtlUnknown;
  end;
  result := tmp;
end;

(******************************************************************************)
function TNotesRichTextItem.GetLinkCount : integer;
begin
  result := FLinks.Count;
end;

(******************************************************************************)
procedure TNotesRichTextItem.AddPassThroughHtml(Html: string);
var
  cdBegin: CDHTMLBEGIN;
  cdEnd: CDHTMLEND;
begin
  FillChar(cdBegin, sizeOf(cdBegin), #0);
  FillChar(cdEnd, sizeOf(cdEnd), #0);
  cdBegin.Header.Signature := SIG_CD_HTMLBEGIN;
  cdBegin.Header.Length := sizeOf(cdBegin);
  cdEnd.Header.Signature := SIG_CD_HTMLEND;
  cdEnd.Header.Length := sizeOf(cdEnd);

  CheckContext;
  AddToContext(@cdBegin, cdBegin.Header.Signature, cdBegin.Header.Length);
  AddTextToContext(0, html);
  AddToContext(@cdEnd, cdEnd.Header.Signature, cdEnd.Header.Length);
end;

(******************************************************************************)
procedure TNotesRichTextItem.AddFormulaButton(aCaption, aFormula: string);
var
  MemSz, NameSz: dword;
  wFormulaLen,wdc0,wdc1: WORD;
  pBegin: CDHOTSPOTBEGIN;
  pEnd: CDHOTSPOTEND;
  pFormula: pChar;
  hFormula: FORMULAHANDLE;
  cButton: CDBUTTON;
  NumOfLines: Word;

//----------------------------------------------------------------------------
  function GetNumOfButtonLines(var aBtnCap: String): Word;
  var
    n: integer;
    aText: String;
  begin
    aText:= aBtnCap;
    n := System.Pos (#13#10,aText);
    Result := 1;
    while n > 0 do begin
      System.Inc( Result, 1 );
      System.Delete(aText,1,n+1);
      n := System.Pos (#13#10,aText);
    end;//of while

    aBtnCap:= SysUtils.StringReplace (aBtnCap,#13#10, #10,[rfReplaceAll,rfIgnoreCase]);
  end;//of function

//----------------------------------------------------------------------------
begin
  CheckContext;

  try
    aFormula := Native2Lmbcs(aFormula);
    NumOfLines := GetNumOfButtonLines(aCaption);
    aCaption := Native2Lmbcs(aCaption);
    try
      if( aFormula <> '' )
      then CheckError(NSFFormulaCompile(
                         nil,                   { name of formula (none) }
                         WORD(0),               { length of name }
                         PChar(aFormula),       { the ASCII formula }
                         WORD(length(aFormula)),{ length of ASCII formula }
                         @hFormula,             { handle to compiled formula }
                         @wFormulaLen,          { compiled formula length }
                         @wdc1,                 { return code from compile (don't care) }
                         @wdc0, @wdc0, @wdc0, @wdc0)) { compile error info (don't care) }
      else begin
        hFormula   := 0;
        wFormulaLen:= 0;
      end;//of else

      // Create hotspot begin record
      MemSz := ODSLength(_CDHOTSPOTBEGIN) + wFormulaLen;
      AddMem (CheckOdd (MemSz));
      pBegin.Header.Signature := SIG_CD_HOTSPOTBEGIN;
      pBegin.Header.Length := MemSz;
      pBegin.aType := HOTSPOTREC_TYPE_BUTTON;
      pBegin.Flags := HOTSPOTREC_RUNFLAG_BEGIN {or HOTSPOTREC_RUNFLAG_NOBORDER};
      pBegin.DataLength := wFormulaLen;
      ODSWriteMemory (@FCurPtr, _CDHOTSPOTBEGIN, @pBegin, 1);

      if( hFormula <> 0 )
      then begin
         pFormula:= OsLockObject(THandle(hFormula));
         Move (pFormula^, FCurPtr^, wFormulaLen);
      end;//of if
      FCurPtr := pointer (dword(FCurPtr) + CheckOdd(wFormulaLen));
    finally
      if( hFormula <> 0 )
      then begin
         OsUnlockObject (hFormula);
         OSMemFree (hFormula);
       end;//of if
      end;//of finally

    // Create the button record
    NameSz := length(aCaption);
    Memsz := ODSLength( _CDBUTTON ) + NameSz;
    AddMem (CheckOdd (MemSz));

    cButton.Header.Signature := SIG_CD_BUTTON;
    cButton.Header.Length := MemSz;
    cButton.Flags := 0; {No flags set for button}
    cButton.Width := 2*ONEINCH;
    cButton.Height := 0;
    cButton.Lines := NumOfLines;
    cButton.FontID := 0;
    FontIDSetFaceID (cButton.FontID, FONT_FACE_SWISS);
    FontIDSetColor (cButton.FontID, NOTES_COLOR_BLACK);
    FontIDSetSize (cButton.FontID, 10);
    ODSWriteMemory (@FCurPtr, _CDBUTTON, @cButton, 1 );
    if( aCaption <> '' )
    then Move (aCaption[1], FCurPtr^, NameSz);
    FCurPtr := pointer (dword(FCurPtr) + CheckOdd(NameSz));

    // Create hotspot end record
    MemSz := ODSLength(_CDHOTSPOTEND);
    AddMem (CheckOdd(MemSz));
    pEnd.Header.Signature := SIG_CD_HOTSPOTEND;
    pEnd.Header.Length := MemSz;
    ODSWriteMemory (@FCurPtr, _CDHOTSPOTEND, @pEnd, 1);
  except
    raise;
  end;//of except
end;

(******************************************************************************)
procedure TNotesRichTextItem.AttachOleObject(AName: string; aType: TRichTextOle2AttachType; aHint: string);
const
  FmtMap: array [TRichTextOle2AttachType] of word = (
    DDEFORMAT_TEXT, DDEFORMAT_METAFILE, DDEFORMAT_BITMAP, DDEFORMAT_RTF, DDEFORMAT_OWNERLINK,
    DDEFORMAT_OBJECTLINK, DDEFORMAT_NATIVE, DDEFORMAT_ICON
  );
var
  clsid: TGUID;
  wsName, wsTempName, wsProgID: WideString;
  pwsName, pwsProgID: PWideChar;
  sTempName, sIconName, sProgID: string;
  pMalloc: IMalloc;
  pstgRoot, pstgObj: IStorage;
  pOleObj: IOleObject;
  pPersist: IPersistStorage;
  stat: STATSTG;
  sLmbcsName, sObject: string;
  ob: CDOLEBEGIN;
  oe: CDOLEEND;
  fp: pchar;
begin
  // To attach an OLE object from file we need to do the following
  // - find an application responsible for it;
  // - load the file and
  // - produce a "structured storage" file
  // - call NoteAttachOleObject passing this file as a parameter
  wsName := aName;
  pwsName := @wsName[1];
  SetLength(wsProgID, 255);
  pwsProgID := @wsProgID[1];
  wsTempName := '';
  sTempName := '';
  pstgRoot := nil;
  pstgObj := nil;
  pOleObj := nil;
  pPersist := nil;
  fillChar(stat, sizeOf(stat), 0);

  // Get class ID of file owner and its string reprs.
  OleCheck(GetClassFile(pwsName, clsid));
  OleCheck(ProgIDFromCLSID(clsid, pwsProgID));
  sProgID := pwsProgID;

  // Get a Malloc allocator
  OleCheck(CoGetMalloc(MEMCTX_TASK, pMalloc));

  // Make in-memory storage
  // We need 2 levels: root and object storage itself
  OleCheck(StgCreateDocfile(nil, STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_CREATE,0,pstgRoot));
  OleCheck(pstgRoot.CreateStorage(pwsProgID,STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_CREATE,0,0,pstgObj));

  // Load the file into the storage
  // Upon return we'll have live OleObject residing in there
  OleCheck(OleCreateFromFile(GUID_NULL, pwsName, IOleObject, OLERENDER_DRAW, nil, nil, pstgObj, pOleObj));

  // Now, save the object producing a temporary on-disk file
  OleCheck(pOleObj.QueryInterface(IPersistStorage, pPersist));
  OleCheck(OleSave(pPersist, pstgObj, True));
  OleCheck(pPersist.SaveCompleted(nil));

  // Get a name of temporary file
  OleCheck(pstgRoot.Stat(stat, STATFLAG_DEFAULT));
  try
    wsTempName := stat.pwcsName;
    sTempName := wsTempName
  finally
    if stat.pwcsName <> nil then pMalloc.Free(stat.pwcsName);
  end;

  // Generate an unique object name
  sObject := 'OLEOBJ' + inttoStr(Random(100)) + inttoStr(Random(100));

  try
    // Import the temporary file into a note
    // This would produce a next-in-sequence item
    SaveContext;
    sLmbcsName := Native2Lmbcs(Name);
    CheckError(NSFNoteAttachOLE2Object(Document.Handle,   //where to
                                      pchar(sTempName),   //storage file
                                      pchar(sObject),     //unique object ID
                                      true,               //create info?
                                      '',                 //description
                                      pchar(sLmbcsName),  //item name
                                      clsid,              //CLSID
                                      FmtMap[aType],      //data format
                                      false,              //scipted?
                                      false,              //ActiveX?
                                      0));

    // Add record showing the object
    CheckContext;
    fillChar(ob, sizeof(ob), 0);
    fillChar(oe, sizeof(oe), 0);

    // OLE begin
    ob.Version := NOTES_OLEVERSION2;
    ob.Flags := OLEREC_FLAG_OBJECT;
    ob.ClipFormat := FmtMap[aType];
    ob.AttachNameLength := length(sObject);
    ob.ClassNameLength := length(sProgID);
    ob.Header.Signature := SIG_CD_OLEBEGIN;
    ob.Header.Length := ODSLength(_CDOLEBEGIN) + CheckOdd(ob.AttachNameLength) + CheckOdd(ob.ClassNameLength);
    AddMem(ob.Header.Length);
    ODSWriteMemory (@FCurPtr, _CDOLEBEGIN, @ob, 1);
    fp := FCurPtr;
    Move (sObject[1], fp^, ob.AttachNameLength);
    fp := pointer (dword(fp) + CheckOdd(ob.AttachNameLength));
    Move (sProgID[1], fp^, ob.ClassNameLength);
    FCurPtr := pointer (dword(FCurPtr) + ob.Header.Length - ODSLength(_CDOLEBEGIN));

    // Hint
    if aHint <> '' then AddTextToContext(0, aHint)
    else begin
      // Attach file with its icon as representation
      sIconName := getAppIcon(aName);

      // Add icon to indicate attachment
      if sIconName <> '' then begin
        AddBmpFile(sIconName);
        DeleteFile(PChar(sIconName));
      end;
    end;

    // OLE end
    oe.Header.Length := ODSLength(_CDOLEEND);
    oe.Header.Signature := SIG_CD_OLEEND;
    AddMem(oe.Header.Length);
    ODSWriteMemory (@FCurPtr, _CDOLEEND, @oe, 1);

    // Save context
    SaveContext;
  finally
    pPersist := nil;
    pOleObj := nil;
    pStgObj := nil;
    pStgRoot := nil;
    if sTempName <> '' then DeleteFile(pchar(sTempName));
  end;
end;

{$IFNDEF NO_INIT_SECTION}

initialization
  NotesDir:= GetActualNotesDir(NotesDataDir);
  randomize;
{$IFNDEF NO_OLE_INITIALIZE}
  OleInitialize(nil);
{$ENDIF}
{$ENDIF}
end.


