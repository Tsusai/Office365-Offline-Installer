program Office2013Setup;

uses
	Forms,
	SysUtils,
	Main in 'Main.pas' {Form1};

{$R *.res}

begin
	Application.Initialize;
	Application.Title := 'Office 2013/365 Install Assistant';
	Application.CreateForm(TForm1, Form1);
	if FindCmdLineSwitch('download',true) then Application.ShowMainForm := false;
	Application.Run;
end.
