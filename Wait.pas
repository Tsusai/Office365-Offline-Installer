unit Wait;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TWaitForm = class(TForm)
    Label1: TLabel;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  WaitForm: TWaitForm;

implementation
uses
	Main;

{$R *.dfm}

procedure TWaitForm.FormShow(Sender: TObject);
begin
//	WaitForm.Left := ((MainForm.Width - WaitForm.Width) div 2) + MainForm.Left;
//	WaitForm.Top := ((MainForm.Height - WaitForm.Height) div 2) + MainForm.Top;
end;

end.
