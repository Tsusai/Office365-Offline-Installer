program Office2013Setup;

uses
  Forms,
  Main in 'Main.pas' {Form1};

{$R *.res}
{$R 'Win8UAC.res'}

begin
  Application.Initialize;
  Application.Title := 'Office 2013/365 Install Assistant';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
