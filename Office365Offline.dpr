program Office365Offline;

uses
  Forms,
  SysUtils,
  Main in 'Main.pas' {Form1},
  Help in 'Help.pas' {Form2};

{$R *.res}

begin
	Application.Initialize;
	Application.Title := 'Office 2013/365 Install Assistant';
	Application.CreateForm(TForm1, Form1);
  Application.CreateForm(THelpForm, HelpForm);
	Application.Run;
end.
