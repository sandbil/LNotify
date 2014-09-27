
unit MainNot;

interface

uses
  Windows, Messages,ComOBJ, SysUtils,  Classes, Graphics, Controls, Forms,
  shellapi, AppEvnts, Buttons, StdCtrls, Menus, ExtCtrls, ImgList,dialogs,Contnrs,
  Class_LotusNotes, Util_LnApi, Util_LnApiErr,Form_LNBrowse, IniFiles, ComCtrls,
  NxScrollControl, NxCustomGridControl, NxCustomGrid, NxGrid ,login,StrUtils,
  NxColumnClasses, NxColumns, NxCollection, NxEdit,
  LinkLabel;

const
  WM_ICONTRAY  = WM_USER + 1;
  //WM_EXIST_MSG = WM_USER+2;

type
  TMainForm = class(TForm)
    PopupMenu1: TPopupMenu;
    CheckNewDoc: TMenuItem;
    startLN: TMenuItem;
    N1: TMenuItem;
    Exit1: TMenuItem;
    N2: TMenuItem;
    Config: TMenuItem;
    TimerIconAnimate: TTimer;
    ImageList1: TImageList;
    SaveBt: TButton;
    TimerCheckDoc: TTimer;
    TimerCloseHint: TTimer;
    CloseBt: TButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    NextGrid1: TNextGrid;
    NxCheckBoxColumn1: TNxCheckBoxColumn;
    NxImageColumn1: TNxImageColumn;
    NxButtonColumn1: TNxButtonColumn;
    AddBt: TNxButton;
    DeleteBt: TNxButton;
    Label8: TLabel;
    GroupBox2: TGroupBox;
    Memo1: TNxMemo;
    SrvName: TNxComboBox;
    FolderSed: TNxComboBox;
    Groups: TNxComboBox;
    NxLabel1: TNxLabel;
    MailFile: TNxButtonEdit;
    BtClear: TNxButton;
    NxLabel2: TNxLabel;
    NxLabel3: TNxLabel;
    EdCheckTime: TNxSpinEdit;
    EdShowTime: TNxSpinEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure CheckNewDocDrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
    procedure startLNDrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
    procedure Exit1DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
    procedure Exit1Click(Sender: TObject);
    procedure ConfigDrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
    procedure startLNClick(Sender: TObject);
    procedure ConfigClick(Sender: TObject);
    procedure TimerIconAnimateTimer(Sender: TObject);
    procedure TimerCheckDocTimer(Sender: TObject);
    procedure HintLabelClick(Sender: TObject);
    procedure HintFormClose(Sender: TObject; var Action: TCloseAction);
    procedure TimerCloseHintTimer(Sender: TObject);
    function OpenDb(srv:string; dbName: string; msg: boolean): TNotesDatabase;
    procedure IconAnimate(IndexIcon:integer);
    procedure SaveBtClick(Sender: TObject);
    procedure CheckNewDocClick(Sender: TObject);
    procedure CloseBtClick(Sender: TObject);
        function NewDocCheck(showwind: boolean; server: string; dbname: string; windprefix:string):boolean;
        procedure CheckNewDocFromAllDb(showEmptywind: boolean);
    procedure NxButtonColumn1ButtonClick(Sender: TObject);
    procedure AddBtClick(Sender: TObject);
    procedure DeleteBtClick(Sender: TObject);
    procedure BtClearClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    TrayIconData: TNotifyIconData;
    ListHintWindHandle: TList;
    ListDbForCheck:TList;
    procedure fShowHint(DocLink:TStrings);
  public
    procedure TrayMessage(var Msg: TMessage); message WM_ICONTRAY;
    procedure DrawBar(ACanvas: TCanvas);
  end;

var
  MainForm: TMainForm;
  iniServer,SavedUser: string;

implementation

{$R *.dfm}
Function ShellExecute(hWnd:HWND;lpOperation:Pchar;lpFile:Pchar;lpParameter:Pchar;
                      lpDirectory:Pchar;nShowCmd:Integer):Thandle; Stdcall;
External 'Shell32.Dll' name 'ShellExecuteA';

procedure TMainForm.DrawBar(ACanvas: TCanvas);
var
  lf : TLogFont;
  tf : TFont;
begin
  with ACanvas do begin
    Brush.Color := clGray;
    FillRect(Rect(0,0,20,192));
    Font.Name := 'Tahoma';
    Font.Size := 7;
    Font.Style := Font.Style - [fsBold];
    Font.Color := clWhite;
    tf := TFont.Create;
    try
      tf.Assign(Font);
      GetObject(tf.Handle, sizeof(lf), @lf);
      lf.lfEscapement := 900;
      lf.lfHeight := Font.Height - 2;
      tf.Handle := CreateFontIndirect(lf);
      Font.Assign(tf);
    finally
      tf.Free;
    end;
    TextOut(2, 85, 'Lotus Email');
  end;
end;

procedure TMainForm.TrayMessage(var Msg: TMessage);
var
  p : TPoint;
begin
  case Msg.lParam of
    WM_LBUTTONDOWN:
    begin
      CheckNewDocFromAllDb(true);
      IconAnimate(0);
    end;
    WM_RBUTTONDOWN:
    begin
       SetForegroundWindow(Handle);
       GetCursorPos(p);
       PopUpMenu1.Popup(p.x, p.y);
       PostMessage(Handle, WM_NULL, 0, 0);
    end;
  end;
end;

function i2str0(I: Longint): string;   { Convert any integer type to a string }
var
  S: string[11];
begin
  Str(I, S);
  s:='0'+s;
  Result:= RightStr(s,2);
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
   FileIni:TIniFile;
   UserName : string;
   UserID, Password: string;
   LNotesConnected: boolean;
   i:integer;
   login:TLoginDlg;
   mNamespaces, mNames, mItems, mResults:TStrings;
begin

  ListHintWindHandle:=TList.Create;

  if Not FileExists(extractfilepath(paramstr(0))+'\LNotify.ini') then
  begin
    FileIni:=TIniFile.Create(extractfilepath(paramstr(0))+'\LNotify.ini');
    FileIni.WriteString('Connect','Server','enter server name');
    FileIni.WriteString('Connect','SedFolder','enter default folder with DB');
    FileIni.WriteString('Connect','Group','enter user group ');
    FileIni.WriteInteger('Timer','CheckTime(min)',15);
    FileIni.WriteInteger('Timer','TimeShowHint(sec)',10);
    FileIni.Free;
  end ;
  FileIni:=TIniFile.Create(extractfilepath(paramstr(0))+'\LNotify'+'.ini');
  iniServer:=FileIni.ReadString('Connect','Server','default srv');
  SrvName.Text:= iniServer;
  FolderSed.Text:=FileIni.ReadString('Connect','SedFolder','default fld');
  Groups.Text:=FileIni.ReadString('Connect','Group','default Users');
  MainForm.TimerCheckDoc.Interval:=FileIni.ReadInteger('Timer','CheckTime(min)',15)*60000;
  EdCheckTime.Text:= FileIni.ReadString('Timer','CheckTime(min)','');
  MainForm.TimerCloseHint.Interval:=FileIni.ReadInteger('Timer','TimeShowHint(sec)',5)*1000;
  EdShowTime.Text:=FileIni.ReadString('Timer','TimeShowHint(sec)','');

  //******************* login
    LNotesConnected:=false;
    UserName:='';
    setLength (UserName, MAXENVVALUE + 1);

   login:=TLoginDLG.create(Self);

     while not LNotesConnected   do begin
        if not login.passSaved then  login.ShowModal;
        UserID:=TUserID(login.UserName.Items.Objects[login.UserName.ItemIndex]).IDFile;
        Password:= login.Password.text;
        Memo1.Lines:=Login.SaveData;
        if (login.ModalResult = mrOk)  or (login.passSaved) then
          begin
           try
             CheckError(SECKFMSwitchToIDFile(pchar(Native2Lmbcs(UserID)),
             pchar(Native2Lmbcs(Password)),pchar(UserName),MAXENVVALUE,0,Nil)) ;
             MainForm.Caption:= MainForm.Caption +': '+ login.UserName.Text;
             LNotesConnected:=true;
             TimerCheckDoc.Enabled:=true;
//***************************************
  PopUpMenu1.OwnerDraw:=True;
  with TrayIconData do
  begin
    cbSize := SizeOf(TrayIconData);
    Wnd := Handle;
    uID := 0;
    uFlags := NIF_MESSAGE + NIF_ICON + NIF_TIP;
    uCallbackMessage := WM_ICONTRAY;
    hIcon := Application.Icon.Handle;
    StrPCopy(szTip, Application.Title);
  end;

  Shell_NotifyIcon(NIM_ADD, @TrayIconData);
//****************************************

           except
              on E: Exception do  Application.MessageBox(PChar(E.Message),
              PChar(Application.Title),MB_OK or MB_ICONERROR);
           end;
              login.passSaved:=false;
              login.Password.text:='';
          end
        else exit;
       end;
       login.free;
//проверка наличи€ настройки почтовой базы
// обновление настройки почтовой базы
// делаетс€ каждый раз на случай если логин€тс€ разные люди
 // If FileIni.ReadString('Mail','MailFile','') = '' then
 // begin
      mNamespaces:=TStringList.Create();
      mNamespaces.Add('$Users');
      mNames:=TStringList.Create();
      mNames.Add(UserName);
      mItems:=TStringList.Create();
      mItems.Add('MailFile');
      mResults:=TStringList.Create();
      TNotesName.LookupNameList(iniServer,mNamespaces, mNames, mItems,
      [nloExhaustive], mResults);
      FileIni.WriteString('Mail','MailFile',Lmbcs2Native(mResults.Strings[0])+'.nsf');
      MailFile.Text:=Lmbcs2Native(mResults.Strings[0])+'.nsf';
      mNamespaces.Free;
      mNames.Free;
      mItems.Free;
      mResults.Free;
//  end;
//проверка наличи€ настройки базы
  If FileIni.ReadString('Database','NSF01','') <> '' then
  begin
      mResults:=TStringList.Create();
      FileIni.ReadSectionValues('Database',mResults);
      for i:=0 to mResults.Count-1 do
      begin
        NextGrid1.AddRow;
        NextGrid1.BeginUpdate;
        NextGrid1.Cell[2,NextGrid1.RowCount-1].AsString :=RightStr(mResults.Strings[i], length(mResults.Strings[i])- pos('=',mResults.Strings[i]));
        NextGrid1.EndUpdate;
      end;
      mResults.Free;
  end;

  FileIni.Free;
  CheckNewDocFromAllDb(false); //первична€ проверка баз данных на наличие новых писем сразу после запуска
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  Shell_NotifyIcon(NIM_DELETE, @TrayIconData);
  ListHintWindHandle.Free;
  ListDbForCheck.Free;
end;

procedure TMainForm.CheckNewDocDrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; Selected: Boolean);
begin
 if Selected then
   ACanvas.Brush.Color := clHighlight
 else
   ACanvas.Brush.Color := clMenu;
 ARect.Left := 25;
 ACanvas.FillRect(ARect);
 DrawText(ACanvas.Handle, PChar('Check email'), -1, ARect, DT_LEFT or DT_VCENTER or DT_SINGLELINE or DT_NOCLIP);
end;

procedure TMainForm.startLNDrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; Selected: Boolean);
begin
 if Selected then
   ACanvas.Brush.Color := clHighlight
 else
   ACanvas.Brush.Color := clMenu;
 ARect.Left := 25;
 ACanvas.FillRect(ARect);
 DrawText(ACanvas.Handle, PChar('Run Lotus Notes'), -1, ARect, DT_LEFT or DT_VCENTER or DT_SINGLELINE or DT_NOCLIP);
end;

procedure TMainForm.Exit1DrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; Selected: Boolean);
begin
 if Selected then
   ACanvas.Brush.Color := clHighlight
 else
   ACanvas.Brush.Color := clMenu;
 ARect.Left := 25;
 ACanvas.FillRect(ARect);
 DrawText(ACanvas.Handle, PChar('Exit'), -1, ARect, DT_LEFT or DT_VCENTER or DT_SINGLELINE or DT_NOCLIP);
 DrawBar(ACanvas);
end;

procedure TMainForm.ConfigDrawItem(Sender: TObject;
  ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
begin
 if Selected then
   ACanvas.Brush.Color := clHighlight
 else
   ACanvas.Brush.Color := clMenu;

 ARect.Left := 25;
 ACanvas.FillRect(ARect);
 DrawText(ACanvas.Handle, PChar('Setting'), -1, ARect, DT_LEFT or DT_VCENTER or DT_SINGLELINE or DT_NOCLIP);
end;

procedure TMainForm.Exit1Click(Sender: TObject);
begin
if MessageDlg('You want to quit from "LNotes Info" ?!',
  mtConfirmation, [mbYes, mbNo], 0) = mrYes then  Application.Terminate;
end;


procedure TMainForm.startLNClick(Sender: TObject);
var
commandline:string;
begin
  commandline:=   'Notes.exe'       ;
  ShellExecute (MainForm.Handle, nil, PChar(commandline), nil, nil, SW_RESTORE);

end;

procedure TMainForm.ConfigClick(Sender: TObject);
begin
  MainForm.Show;
end;


procedure TMainForm.IconAnimate(IndexIcon:integer);
var
  Icon: TIcon;
begin
  Icon:=TIcon.Create;
  try
    ImageList1.GetIcon(IndexIcon,Icon);
    TrayIconData.hIcon := Icon.Handle;
    Shell_NotifyIcon(NIM_Modify, @TrayIconData);
  finally
    Icon.Free;
  end;
end;


procedure TMainForm.TimerIconAnimateTimer(Sender: TObject);
{$J+}
const
  Index : Integer = 0;
{$J-}
begin
  Inc(Index);
  if Index = 2 then Index:=0;
  IconAnimate(Index);
end;



procedure TMainForm.HintLabelClick(Sender: TObject);
begin
    TForm(TLabel(Sender).Parent).close;
end;

procedure TMainForm.HintFormClose(Sender: TObject; var Action: TCloseAction);
var
i,ind,wHeight:integer;
begin
 Action:=caFree;
    ind:=ListHintWindHandle.IndexOf(Sender);
    wHeight:=TForm(Sender).Height ;
    ListHintWindHandle.Remove(TForm(Sender));
    for i:=ind to ListHintWindHandle.count-1  do
    begin
       TForm(ListHintWindHandle[i]).Top:= TForm(ListHintWindHandle[i]).Top+wHeight;
       TForm(ListHintWindHandle[i]).Repaint;
    end;

end;

procedure TMainForm.fShowHint(DocLink:TStrings);
var H:HWND;
    Rec:TRect;
    HintForm:TForm;
    Label1:TLabel;
    HintLabel:TLinkLabel;
    aw:hwnd;
    i,ofHeight:integer;
    CapText:string;
begin

  H := FindWindow('Shell_TrayWnd', nil);
  if h=0 then exit;

  GetWindowRect(h, Rec);

  HintForm:=TForm.Create(MainForm);
  ListHintWindHandle.Add(HintForm);
  ofHeight:=0;
  with HintForm do
  begin
    Width:=345;
    If DocLink.Count =2 then Height:=80 else Height:=200;
    BorderIcons :=[biSystemMenu];
    Color:=clSkyBlue;
    BorderStyle:=bsSingle; //bsNone;
    AlphaBlend:=true;
    AlphaBlendValue:=220;
    Left:=Screen.Width-Width;
    for i:=0 to ListHintWindHandle.Count-1 do
    begin
        ofHeight:=ofHeight+TForm(ListHintWindHandle[i]).Height;
    end;
    Top:=Rec.Top - ofHeight;
    FormStyle:=fsStayOnTop;
    Caption:=DocLink[0];
    OnClose :=HintFormClose;



    //—оздаЄм текст
      Label1:=TLabel.Create(nil);
      with Label1 do
      begin
          Parent:=HintForm;
          Align:=alTop;
          Alignment:=taCenter;
          Font.Style:=[fsBold];
          if DocLink.count < 8 then Caption:=DocLink[0]+chr(10)+DocLink[1]
          else Caption:=DocLink[0]+chr(10)+DocLink[1]+' (show first 8)';

      end;

    //DocLink
    for i:=2 to DocLink.Count-1 do
    begin
      HintLabel:=TLinkLabel.Create(nil);
      with HintLabel do
      begin
          Parent:=HintForm;
          Top:=i*16;
          left:=10 ;
          //WordWrap:=true;
          LinkType:=ltNotes;
          ShowHint:=true;
          HyperLink:=LeftStr(DocLink[i],pos('#|',DocLink[i]));
          CapText:=RightStr(DocLink[i],length(DocLink[i])-pos('#|',DocLink[i])-1);
          Hint:=CapText;
          if trim(CapText)=''then Caption:= 'no caption'
          else if length(CapText)>55 then Caption:= leftstr(CapText,55)+'...'
                                      else Caption:= CapText;
  //        OnClick := HintLabelClick;
      end;
      if i>8 then break;
    end;
    aw:=GetActiveWindow;
    ShowWindow(handle,SW_SHOWNOACTIVATE);
    SetActiveWindow(aw);
    Repaint;
   end;

end;

procedure TMainForm.TimerCheckDocTimer(Sender: TObject);
begin
  CheckNewDocFromAllDb(false);
end;

procedure TMainForm.TimerCloseHintTimer(Sender: TObject);
begin
if ListHintWindHandle.Count>0  then TForm(ListHintWindHandle[0]).Close
else  TimerCloseHint.Enabled:=False;
end;

function TMainForm.OpenDb(srv:string; dbName: string; msg: boolean): TNotesDatabase;
var
  Server : string;
begin
  Server := srv;
  Result := TNotesDatabase.create;
    try
      Result.open (Server, dbName);
      if msg then showMessage ('Opened successfully, title is ' + Result.Title);
    except
      Result.free;
      raise;
    end;
end;

function TMainForm.NewDocCheck(showwind: boolean; server:string; dbname:string; windprefix:string): boolean;
var
  Db: TNotesDatabase;
  UnreadDocCount,i:integer;
    Doc: TNotesDocument;
    mDocLink,mDocNeedMarkRead:TStrings;

begin
  try
    db := OpenDb(server, dbname, false);
    if db <> nil then try
      UnreadDocCount:=db.UnreadDocuments.Count;

      mDocLink:=TStringList.Create();
      mDocNeedMarkRead:=TStringList.Create();
      mDocLink.Add(windprefix + Db.Title);
      mDocLink.Add('New messages: ');
      if UnreadDocCount>0 then begin
          for i := 0 to db.UnreadDocuments.Count-1 do begin
              try
              Doc := db.UnreadDocuments.Document[i];
                if UpperCase(Doc.Form) <> UpperCase('setting') then
                   begin
                    mDocLink.Add(leftstr(server,pos('/',server)-1)+'/'+TimeDatetoStr(db.DatabaseID)+'/0/'+UNIDtoStr(Doc.UniversalID,false)+'?OpenDocument'+'#|'+Doc.Subject);
                      //showmessage('UNID- '+UNIDtoStr(Doc.UniversalID,false) +chr(10)+
                        //        'Form- '+Doc.Form    +chr(10)+
                          //      'Subject- '+Doc.Subject);

                   end
                else mDocNeedMarkRead.AddObject(inttostr(i),tobject(db.UnreadDocuments.DocumentID[i]));
              finally
                Doc.free;
              end;
           end;

           mDocLink[1]:=mDocLink[1]+inttostr(mDocLink.Count-2);

           if mDocNeedMarkRead.Count>0 then    //если в базе есть непрочитанные настройки, то отмечаем их как прочитанные
              begin
                for i := 0 to mDocNeedMarkRead.Count-1 do begin
                  try
                          db.MarkRead(NoteID(mDocNeedMarkRead.Objects[i]),true);
                  finally
                      //Doc.free;
                  end;
                end;
              end;
           fShowHint(mDocLink);
           TimerCloseHint.Enabled:=true;
           result:=true ;
      end
      else
        begin
          result:=false ;
          mDocLink[1]:=mDocLink[1]+'0';
          if showwind then begin
            fShowHint(mDocLink);
            TimerCloseHint.Enabled:=true;
          end;
        end;
    finally
      Db.free;
      mDocLink.Free;
      mDocNeedMarkRead.Free;
    end;
  except
    on E: ELotusNotes do MessageDlg(E.message, mtError, [mbOK], 0);
    else raise;
  end;
end;


procedure TMainForm.CheckNewDocFromAllDb(showEmptywind: boolean);
var
i:integer;
prIcnAnim, pr:boolean;
begin
prIcnAnim:=NewDocCheck(showEmptywind,iniServer,MailFile.Text,'Email: ');
for i:=0 to NextGrid1.RowCount-1 do
begin
    pr:=NewDocCheck(showEmptywind,iniServer,NextGrid1.Cell[2,i].asString, 'NSF: ');
    prIcnAnim:=(prIcnAnim or pr );
end;
if prIcnAnim then  TimerIconAnimate.Enabled:=true
else TimerIconAnimate.Enabled:=false;

end;

procedure TMainForm.CheckNewDocClick(Sender: TObject);
begin
    CheckNewDocFromAllDb(true);
end;

procedure TMainForm.CloseBtClick(Sender: TObject);
begin
MainForm.Hide;
end;

procedure TMainForm.NxButtonColumn1ButtonClick(Sender: TObject);
var
NameNSF:string;
begin
  with NextGrid1 do
  if LnBrowse ('', iniServer, NameNSF) then
    begin
         EndEditing;
         Cell[SelectedColumn, SelectedRow].AsString:= NameNSF;
    end;
end;

procedure TMainForm.AddBtClick(Sender: TObject);
var
NameNSF:string;
begin
  NameNSF:= FolderSed.Text;
  if LnBrowse ('', iniServer, NameNSF) then
    begin
         NextGrid1.AddRow();
         NextGrid1.Cell[2, NextGrid1.RowCount-1].AsString:= NameNSF;
    end;

end;

procedure TMainForm.DeleteBtClick(Sender: TObject);
begin
if (NextGrid1.SelectedCount>0) then
  NextGrid1.DeleteRow(NextGrid1.SelectedRow);
end;

procedure TMainForm.SaveBtClick(Sender: TObject);
var
FileIni:TIniFile;
i: integer;
function MakeItAString(I: Longint): string;   { Convert any integer type to a string }
var
  S: string[11];
begin
  Str(I, S);
  Result:= S;
end;
begin
  FileIni:=TIniFile.Create(extractfilepath(paramstr(0))+'\LNotify.ini');

  FileIni.WriteString('Connect','Server',SrvName.Text);
  FileIni.WriteString('Connect','SedFolder',FolderSed.Text);
  FileIni.WriteString('Connect','Group',Groups.Text);
  FileIni.WriteInteger('Timer','CheckTime(min)',EdCheckTime.AsInteger);
  MainForm.TimerCheckDoc.Interval:=EdCheckTime.AsInteger*60000;
  FileIni.WriteInteger('Timer','TimeShowHint(sec)',EdShowTime.AsInteger);
  MainForm.TimerCloseHint.Interval:=EdShowTime.AsInteger*1000;
  FileIni.WriteString('Mail','MailFile',MailFile.Text);

  FileIni.EraseSection('Database');
  for i:=0 to NextGrid1.RowCount-1 do
          FileIni.WriteString('Database','NSF' + i2str0(i + 1),NextGrid1.Cell[2, i].AsString);
  FileIni.Free;

end;

procedure TMainForm.BtClearClick(Sender: TObject);
var
FileIni: TIniFile;
begin
    Memo1.Clear;
    FileIni:=TIniFile.Create(extractfilepath(paramstr(0))+'\LNotify'+'.ini');
    FileIni.WriteString('Connect','User','');
    FileIni.Free;
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caNone;
  MainForm.Hide;
end;

end.
