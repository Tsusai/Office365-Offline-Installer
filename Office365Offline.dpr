program Office365Offline;

uses
  Forms,
  SysUtils,
  Main in 'Main.pas' {Mainform},
  Help in 'Help.pas' {Form2},
  Wait in 'Wait.pas' {WaitForm},
  OfficeVerification in 'OfficeVerification.pas' {VerifyForm},
  OfficeLanguageInfo in 'OfficeLanguageInfo.pas';

{$R *.res}

begin
	Application.Initialize;
	Application.Title := '';
	Application.CreateForm(TMainform, Mainform);
  Application.CreateForm(THelpForm, HelpForm);
  Application.CreateForm(TWaitForm, WaitForm);
  Application.CreateForm(TVerifyForm, VerifyForm);
  Application.Run;
end.
