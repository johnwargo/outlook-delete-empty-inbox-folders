program DeleteEmptyFolders;

uses
  Vcl.Forms,
  main in 'main.pas' {frmMain},
  log in 'log.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
