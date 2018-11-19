program fastpaste;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  Winapi.Windows,
  Vcl.Dialogs;

var
  hMutex: THandle;

{$R *.res}
{$R *.dkl_const.res}

begin
  hMutex := 0;
  hMutex := CreateMutex(nil, False, '{1EAB05F1-D23D-421D-8CC3-1B285FD1B72A}');
  try
    if (hMutex = INVALID_HANDLE_VALUE) or (GetLastError=ERROR_ALREADY_EXISTS) then
    begin
      MessageDlg('Another copy of this program is already running.', mtError, [mbOK], 0);
      Halt;
    end;
    Application.Initialize;
    Application.Initialize;
    Application.MainFormOnTaskbar := True;
    Application.Title := 'NI FastPaste';
    Application.CreateForm(TForm1, Form1);
    Application.Run;
  finally
    ReleaseMutex(hMutex);
    CloseHandle(hMutex);
  end;
end.
