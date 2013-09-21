unit Main;

interface

uses
	Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
	Dialogs, StdCtrls, ExtCtrls, VistaAltFixUnit;

type
	TForm1 = class(TForm)
		Button1: TButton;
		RadioGroup1: TRadioGroup;
    VistaAltFix1: TVistaAltFix;
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
	JclFileUtils,
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

type
	TOSInfo = class(TObject)
	public
		class function IsWOW64: Boolean;
	end;

class function TOSInfo.IsWOW64: Boolean;
type
	TIsWow64Process = function(
		Handle: THandle;
		var Res: BOOL
	): BOOL; stdcall;
var
	IsWow64Result: BOOL;
	IsWow64Process: TIsWow64Process;
begin
	IsWow64Process := GetProcAddress(
		GetModuleHandle('kernel32'), 'IsWow64Process'
	);
	if Assigned(IsWow64Process) then
	begin
		if not IsWow64Process(GetCurrentProcess, IsWow64Result) then
			raise Exception.Create('Bad process handle');
		Result := IsWow64Result;
	end
	else
		Result := False;
end;



procedure TForm1.Button1Click(Sender: TObject);
Var
	Version : String;
begin
	//set version string
	Case RadioGroup1.ItemIndex of
	0: Version := '"HS"';
	1: Version := '"HB"';
	2: Version := '"O365Home"';
	3: Version := '"O365Pro"';
	4: Version := '"O365SMB"';
	end;

	case TOSInfo.IsWOW64 of
		true : Version:= Version + '64.xml';
		false: Version:= Version + '32.xml';
	end;

	SetCurrentDirectory(PChar(AppPath+'Setup\'));
	ExecuteHidden('setup.exe', '/configure ' + Version);
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
