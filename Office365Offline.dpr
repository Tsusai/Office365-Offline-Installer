program Office365Offline;

uses
	Forms,
	SysUtils,
	Main in 'Main.pas' {Mainform},
	Help in 'Help.pas' {Form2},
	OfficeVerification in 'OfficeVerification.pas',
	Wait in 'Wait.pas' {WaitForm};

{$R *.res}

begin
	Application.Initialize;
	Application.Title := '';
	Application.CreateForm(TMainform, Mainform);
	Application.CreateForm(THelpForm, HelpForm);
	Application.CreateForm(TWaitForm, WaitForm);
	Application.Run;
end.
