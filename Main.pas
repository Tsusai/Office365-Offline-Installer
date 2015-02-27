unit Main;

interface

uses
	Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
	Dialogs, StdCtrls, ExtCtrls;

type
	TForm1 = class(TForm)
		Button1: TButton;
		RadioGroup1: TRadioGroup;
		procedure Button1Click(Sender: TObject);
		procedure FormCreate(Sender: TObject);
	private
		{ Private declarations }
	public
		{ Public declarations }
	end;

var
	Form1: TForm1;

implementation
uses
	ShellAPI;

{$R *.dfm}
var
	AppPath : string = '';

procedure ExecuteHidden(
	Command : string;
	Params : string = '';
	StartFolder : string = '');
begin
	ShellExecute(
		Form1.Handle,
		'open',
		PChar(Command),
		PChar(Params),
		PChar(StartFolder),
		SW_HIDE);
end;

procedure TForm1.Button1Click(Sender: TObject);
Var
	Version : String;
begin
	//set version string
	Case RadioGroup1.ItemIndex of
	0: Version := '"HomeStudent"';
	1: Version := '"HomeBuisness"';
	2: Version := '"O365HomePremium"';
	3: Version := '"O365ProPlusRetail"';
	4: Version := '"O365SmallBuisnessPrem"';
	end;

	SetCurrentDirectory(PChar(AppPath+'Setup\'));
	ExecuteHidden('setup.exe', '/configure ' + Version + '.xml');
	Application.Minimize;
	Sleep(15000);
	Close;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
	Left:=(Screen.Width-Width)  div 2;
	Top:=(Screen.Height-Height) div 2;
	AppPath := ExtractFilePath(ParamStr(0));
end;

end.
