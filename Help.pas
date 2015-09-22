unit Help;

interface

uses
	Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
	Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
	THelpForm = class(TForm)
		Memo1: TMemo;
		procedure FormCreate(Sender: TObject);
	private
		{ Private declarations }
	public
		{ Public declarations }
	end;

var
	HelpForm: THelpForm;

implementation

{$R *.dfm}

procedure THelpForm.FormCreate(Sender: TObject);
begin
	Self.Position := poDesktopCenter;
end;

end.
