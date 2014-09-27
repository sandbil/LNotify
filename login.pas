unit login;

{$P+}

interface
uses SysUtils, Windows, Messages, Classes, Graphics, Controls, Dialogs, StrUtils,
  Forms, StdCtrls, ExtCtrls, Buttons,inifiles,Class_LotusNotes, Util_LnApi, Util_LnApiErr,rc4;

type
  TLoginDlg = class(TForm)
    Panel: TPanel;
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Password: TEdit;
    OKBitBtn1: TBitBtn;
    CancelBitBtn2: TBitBtn;
    Panel2: TPanel;
    Image1: TImage;
    UserName: TComboBox;
    savePass: TCheckBox;
    OpenDialog1: TOpenDialog;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure OKBitBtn1Click(Sender: TObject);
    procedure CancelBitBtn2Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure UserNameChange(Sender: TObject);
  private
    { Private declarations }
        KeyData: TRC4Data;
  public
    SaveData:Tstrings;
    passSaved:boolean;
    { Public declarations }
  end;


type
   TUserID = class
   private
     fName: string;
     fIDFile: string;
   public
     property Name : string read fName;
     property IDFile : string read fIDFile;
     constructor Create(const name : string; const idFile : string) ;
   end;

var
  LoginDLG: TLoginDLG;
const
  Key: array[0..4] of ansichar = ('U', 'p', 'r', 'C', 'h');

implementation
//uses main;

 constructor TUserID.Create(const name : string; const idFile : string) ;
 begin
   fName := name;
   fIDFile := idFile;
 end;


{$R *.dfm}

procedure GetIDFiles( Path: string; SpisFile: TStrings);
var sr: TSearchRec;
begin
  if FindFirst(Path+'\*.id', faAnyFile, sr) = 0 then
  begin
    repeat
      SpisFile.Add(Path+'\'+sr.Name);
    until FindNext(sr)<>0 ;
    SysUtils.FindClose(sr);
  end;
end;


procedure TLoginDLG.FormCreate(Sender: TObject);
var
origId, sPath, SavedUser, buffer: string;
len,cntId: integer;
FileIni: TIniFile;
sData: array[0..1024] of ansichar;
IDFromData:TStrings;

  BBuffer: PAnsiChar;

begin
  RC4Init(KeyData,@Key,Sizeof(Key));

  FileIni:=TIniFile.Create(extractfilepath(paramstr(0))+'\LNotify'+'.ini');
  SavedUser:=FileIni.ReadString('Connect','User','');
  FileIni.Free;
  origId := '';
  SaveData:=TStringList.Create();
  SaveData.Add('');
  SaveData.Add('Сохраненных данных нет');
  try
  if  SavedUser<>'' then
  begin
      Buffer:='';
      len:=length(SavedUser) div 2;
      GetMem(BBuffer, len);
      HexToBin(PAnsiChar(AnsiString(SavedUser)), BBuffer , len);
      //HexToBin(pchar(SavedUser), pchar(Buffer) , len);

      move(BBuffer[0],sData,len);
      RC4Reset(KeyData);
      RC4Crypt(KeyData,@sData,@sData,len);
      buffer:= strPas(sData);

      origId:=LeftStr(buffer,pos(',',buffer)-1);
      if Fileexists(origID) then begin
          SaveData.Clear;
          SaveData.Add('User: ' + UserNameFromID(origId));
          SaveData.Add('IDFile: ' + origId);

          UserName.Items.AddObject(UserNameFromID(origId),TUserID.Create(UserNameFromID(origId), origId));

          Password.Text:=RightStr(buffer,len-pos(',',buffer));
          UserName.ItemIndex:=0;
          passSaved:=true;
       end
  end;
    origId := '';
    setLength (origID, MAXENVVALUE + 1);
    OSGetEnvironmentString ('KeyFilename', pchar(origID), MAXENVVALUE);
    origID := Lmbcs2Native(strPas(pchar(origId)));
    sPath:='';
    setLength (sPath, MAXENVVALUE + 1);
    OSGetEnvironmentString ('Directory=', pchar(sPath), MAXENVVALUE);
    OpenDialog1.InitialDir:=Lmbcs2Native(strPas(pchar(sPath)));
    if Not Fileexists(origID) then
      begin
        origID := Lmbcs2Native(strPas(pchar(sPath)))+'\'+ origID;
        if Fileexists(origID) then
              UserName.Items.AddObject(UserNameFromID(origID),TUserID.Create(UserNameFromID(origID), origID));
      end
    else UserName.Items.AddObject(UserNameFromID(origID),TUserID.Create(UserNameFromID(origID), origID));

    //*******из каталога Notes
    IDFromData:=TStringList.Create();
    GetIDFiles(strPas(pchar(sPath)),IDFromData);
    for cntId:=0 to IDFromData.Count - 1 do
        UserName.Items.AddObject(UserNameFromID(IDFromData.Strings[cntId]),TUserID.Create(UserNameFromID(IDFromData.Strings[cntId]), IDFromData.Strings[cntId]));
    finally
    IDFromData.Free;
    UserName.Items.AddObject('Выбрать другого',TUserID.Create('Выбрать другого', 'Выбрать другого'));
    UserName.ItemIndex:=0;
    end;

end;


procedure TLoginDLG.FormShow(Sender: TObject);
begin
 password.SetFocus;
end;

procedure TLoginDLG.FormClose(Sender: TObject; var Action: TCloseAction);
var
FileIni: TIniFile;
sTmp: string;
sData: array[0..1024] of char;
len: integer;
Buffer, text:string;

begin
if savePass.Checked then
  begin
    len:= length(TUserID(UserName.Items.Objects[UserName.ItemIndex]).IDFile)+ length(password.Text)+1;
    sTmp:= TUserID(UserName.Items.Objects[UserName.ItemIndex]).IDFile+','+password.Text;
    move(sTmp[1],sData,len);
    RC4Reset(KeyData);
    RC4Crypt(KeyData,@sData,@sData,len);

    //Buffer := strPas(sData);
    Buffer := sData    ;
    text:='';
    setLength (text, len*2);
    BinToHex(pchar(Buffer), pchar(text) , len);

    FileIni:=TIniFile.Create(extractfilepath(paramstr(0))+'\LNotify'+'.ini');
    FileIni.WriteString('Connect','User',text);
    FileIni.Free;
  end;

end;

procedure TLoginDlg.OKBitBtn1Click(Sender: TObject);
begin
//
end;

procedure TLoginDlg.CancelBitBtn2Click(Sender: TObject);
begin
Application.Terminate;
end;

procedure FreeObjects(const strings: TStrings) ;   //удаляет созданные нами obj'ы
var
  idx : integer;
begin
  for idx := 0 to Pred(strings.Count) do
  begin
    strings.Objects[idx].Free;
    strings.Objects[idx] := nil;
  end;
end;

procedure TLoginDlg.FormDestroy(Sender: TObject);
begin
FreeObjects(UserName.Items) ;
SaveData.Free;

end;

procedure TLoginDlg.UserNameChange(Sender: TObject);
var
origID:string;
begin
with UserName do
if text='Выбрать другого' then
begin
   If OpenDialog1.Execute then
   begin
       origID:=OpenDialog1.FileName;
       Items.AddObject(UserNameFromID(origID),TUserID.Create(UserNameFromID(origID), origID));
       ItemIndex:= UserName.Items.Count-1;
   end;
end;

    
end;

end.
