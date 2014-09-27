{==============================================================================|
| Project : Notes/Delphi class library                           | 3.8         |
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
|     27.02.00, Sergey Kolchin                                                 |
|==============================================================================|
| Contributors and Bug Corrections:                                            |
|   Fujio Kurose                                                               |
|   Noah Silva                                                                 |
|   Tibor Egressi                                                              |
|   Andreas Pape                                                               |
|   Anatoly Ivkov                                                              |
|   Winalot                                                                    |
|     and others...                                                            |
|==============================================================================|
| History: see README.TXT                                                      |
|==============================================================================|
| This unit contains Open Database dialog and LnBrowse function                |
|  used to visually browse for a database on local or remote server            |
|                                                                              |
| In some configurations, Notes returns file names in OEM codepage rather than |
|   in ANSI. If this happened (dialog displays incorrect characters), define   |
|   a symbol USE_OEM_CODEPAGE (see below)
|==============================================================================|}
unit Form_LNBrowse;

interface

{.$DEFINE USE_OEM_CODEPAGE}

// Delphi version
{$IFDEF VER130}
  {$DEFINE D5}
  {$DEFINE D4}
{$ELSE}
  {$IFDEF VER120}
    {$DEFINE D4}
  {$ELSE}
    {$DEFINE D3}
  {$ENDIF}
{$ENDIF}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, Class_LotusNotes, ImgList
{$IFDEF D4}
  , ImgList;
{$ELSE}
  ;
{$ENDIF}


type
  // Browse form
  TLnBrowseDlg = class(TForm)
    Label1: TLabel;
    CbServer: TComboBox;
    TreeView: TTreeView;
    Label2: TLabel;
    Label3: TLabel;
    EFileName: TEdit;
    BtOpen: TButton;
    BtCancel: TButton;
    BtBrowse: TButton;
    OpenDialog: TOpenDialog;
    BtRefresh: TButton;
    ImageList1: TImageList;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure CbServerDropDown(Sender: TObject);
    procedure BtOpenClick(Sender: TObject);
    procedure CbServerChange(Sender: TObject);
    procedure TreeViewChange(Sender: TObject; Node: TTreeNode);
    procedure TreeViewExpanding(Sender: TObject; Node: TTreeNode;
      var AllowExpansion: Boolean);
    procedure BtBrowseClick(Sender: TObject);
    procedure BtRefreshClick(Sender: TObject);
    procedure TreeViewGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeViewGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeViewDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    Directory: TNotesDirectory;
    Port: string;
    function GetPath (Item: TTreeNode): string;
    procedure ListDir (Item: TTreeNode);
  public
  end;

var
  LnBrowseDlg: TLnBrowseDlg;
  SedFldr:string;
function LnBrowse (aPort: string; var aServer, aPath: string): boolean;

// Combines/divides LN path to be used in editors
// ! is a separator
const LN_EDIT_SEPARATOR = '!';
procedure ParseLnPath (const aPath: string; var Server, Path: string);
function CombineLnPath  (const Server, Path: string): string;

implementation

{$R *.DFM}
procedure ParseLnPath;
var
  n: integer;
begin
  n := Pos (LN_EDIT_SEPARATOR, aPath);
  if n = 0 then begin
    Server := '';
    Path := aPath;
  end
  else begin
    Server := Trim(copy (aPath, 1, n-1));
    Path := Trim(copy (aPath, n+1, length(aPath)-n));
  end;
end;

function CombineLnPath;
begin
  Result := Server;
  if Result <> ''
    then Result:= Result + LN_EDIT_SEPARATOR + Path //appendStr (Result, LN_EDIT_SEPARATOR + Path)
    else Result := Path;
end;

function ChgOemToAnsi (aStr : string): string;
begin
{$IFDEF USE_OEM_CODEPAGE}
  setLength(Result, length(aStr)+1);
  OemToChar(pchar(aStr), pchar(Result));
  Result := strPas(pchar(Result));
{$ELSE}
  Result := aStr;
{$ENDIF}
end;

function ChgNotesSep (aStr : string): string;
const
  NotesSeperator = '|';
var
  i: integer;
begin
  Result:= aStr;
  repeat
    i := Pos (#10, Result);
    if i > 0 then Result[i]:= NotesSeperator;
  until (I = 0);
end;

function LnBrowse;
begin
  SedFldr:=aPath;
  LnBrowseDlg := TLnBrowseDlg.create (nil);
  with LnBrowseDlg do try
    Port := aPort;
    CbServer.text := aServer;
    Result := False;
    showModal;
    if modalResult = mrOk then begin
      aServer := CbServer.text;
      if compareText (aServer,'Local') = 0 then aServer := '';
      aPath := EFileName.Text;
      Result := True;
    end;
  finally
    LnBrowseDlg.free;
  end;
end;

procedure TLnBrowseDlg.FormCreate(Sender: TObject);
begin
  Directory := TNotesDirectory.create;

end;

procedure TLnBrowseDlg.FormDestroy(Sender: TObject);
begin
  Directory.free;
end;

procedure TLnBrowseDlg.CbServerDropDown(Sender: TObject);
begin
  if CbServer.Items.Count = 0 then begin
    Screen.Cursor := crHourglass;
    try
      Directory.ListServers ('', CbServer.Items);
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

procedure TLnBrowseDlg.BtOpenClick(Sender: TObject);
begin
  try
    if (TreeView.Items.count = 0) and (EFileName.text = '') then begin
      // List root directory
      ModalResult := mrNone;
      ListDir (nil);
    end
    else begin
      // Exiting
      if TreeView.Items.Count <> 0 then
        if (EFileName.text = '') or ((TreeView.Selected <> nil) and (TreeView.Selected.ImageIndex <> 0)) then
          raise Exception.create ('Select a database to open');
      ModalResult := mrOk;
    end;
  except
    ModalResult := mrNone;
    raise;
  end;
end;

procedure TLnBrowseDlg.CbServerChange(Sender: TObject);
begin
  TreeView.Items.BeginUpdate;
  TreeView.Items.Clear;
  TreeView.Items.EndUpdate;
  EFileName.text := '';
end;

function TLnBrowseDlg.GetPath;
var
  s: string;
  n: integer;
begin
  Result := '';
  while Item <> nil do begin
    s := Item.Text;
    n := Pos ('[', s);
    if n <> 0 then begin
      delete (s, 1, n);
      n := Pos (']', s);
      if n <> 0 then delete (s, n, length(s)-n+1);
    end;
    Result := s + '\' + Result;
    Item := Item.Parent;
  end;
  if (Result <> '') and (Result[length(Result)] = '\') then delete(Result,length(Result),1);
end;

procedure TLnBrowseDlg.ListDir;
var
  Entry: TNotesDirEntry;
  NotDone: boolean;
  aPath: string;
  Node: TTreeNode;
begin
  Screen.Cursor := crHourglass;
  try
    if Item = nil then aPath := '' else aPath := GetPath (Item);
    NotDone := Directory.FindFirst (cbServer.Text, aPath, [nfoFiles, nfoTemplates, nfoSubDirs], Entry);
    while NotDone do begin
      if Entry.EntryType then begin
        Node := TreeView.Items.AddChild (Item, ChgOemToAnsi(Entry.FileName));
        Node.HasChildren := True;
        Node.ImageIndex := 1;
        Node.SelectedIndex := 1;
        Node.StateIndex := -1;
        Node.Data := nil;
      end
      else begin
        Node := TreeView.Items.AddChild (Item,
          ChgNotesSep(ChgOemToAnsi(Entry.FileInfo))
           + ' [' + ChgOemToAnsi(Entry.fileName) + ']');
        Node.HasChildren := False;
        Node.ImageIndex := 0;
      end;
      NotDone := Directory.FindNext (Entry);
    end;
    Directory.FindClose;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TLnBrowseDlg.TreeViewChange(Sender: TObject; Node: TTreeNode);
begin
  EFileName.text := GetPath (Node);
end;

procedure TLnBrowseDlg.TreeViewExpanding(Sender: TObject; Node: TTreeNode;
  var AllowExpansion: Boolean);
begin
  AllowExpansion := False;
  if Node.ImageIndex = 0 then exit;
  if Node.Data <> nil then AllowExpansion := Node.HasChildren
  else begin
    // Listing sub-directory
    ListDir (Node);
    AllowExpansion := Node.count > 0;
    Node.HasChildren := AllowExpansion;
    Node.Data := pointer(1);
  end;
end;

procedure TLnBrowseDlg.BtBrowseClick(Sender: TObject);
begin
  if OpenDialog.execute then begin
    CbServer.Text := '';
    TreeView.Items.Clear;
    EFileName.text := OpenDialog.FileName;
  end;
end;

procedure TLnBrowseDlg.BtRefreshClick(Sender: TObject);
begin
  TreeView.Items.Clear;
  ListDir(nil);
end;

procedure TLnBrowseDlg.TreeViewGetImageIndex(Sender: TObject;
  Node: TTreeNode);
begin
  if Node.ImageIndex <> 0 then
    if Node.Expanded then Node.ImageIndex := 2 else Node.ImageIndex := 1;
end;

procedure TLnBrowseDlg.TreeViewGetSelectedIndex(Sender: TObject;
  Node: TTreeNode);
begin
  if Node.ImageIndex <> 0 then
    if Node.Expanded then Node.SelectedIndex := 2 else Node.SelectedIndex := 1;
end;

procedure TLnBrowseDlg.TreeViewDblClick(Sender: TObject);
begin
  if (TreeView.Selected <> nil) and (TreeView.Selected.ImageIndex = 0) then BtOpenClick(Sender);
end;

function GetNodeByText (ATree : TTreeView; AValue:String; AVisible: Boolean): TTreeNode;
var
    Node: TTreeNode;
begin
  Result := nil;
  if ATree.Items.Count = 0 then Exit;
  Node := ATree.Items[0];
  while Node <> nil do
  begin
    if UpperCase(Node.Text) = UpperCase(AValue) then
    begin
      Result := Node;
      if AVisible then
        Result.MakeVisible;
      Break;
    end;
    Node := Node.GetNext;
  end;
end;

procedure TLnBrowseDlg.FormShow(Sender: TObject);
var
Node: TTreeNode;
begin
    ListDir (nil);
    if SedFldr <> '' then Node:=GetNodeByText(TreeView,SedFldr,true);
    if Node <> nil then
    begin
        TreeView.SetFocus;
        Node.Selected:=true;
        if Node.HasChildren then
        begin
            ListDir (Node);
            Node.Expand(true);
        end;
    end;
end;

end.
