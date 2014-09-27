program LNotify;

uses
  Forms,
  MainNot in 'MainNot.pas' {MainForm},
  Util_LNApi in 'Ln\Util_LNApi.pas',
  Class_LotusNotes in 'Ln\Class_LotusNotes.pas',
  Util_LnApiErr in 'Ln\Util_LnApiErr.pas',
  Class_NotesRTF in 'Ln\Class_NotesRTF.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.ShowMainForm := False;
  Application.Title := 'Lotus Notes Информатор';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
