unit Main;

interface

uses
	Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
	Dialogs, StdCtrls, ExtCtrls, Vcl.CheckLst;

type
	TForm1 = class(TForm)
		Button1: TButton;
		RadioGroup1: TRadioGroup;
		CheckListBox1: TCheckListBox;
		CustomCheck: TCheckBox;
    ComboBox1: TComboBox;
    Button2: TButton;
    ComboBox2: TComboBox;
		procedure Button1Click(Sender: TObject);
		procedure FormCreate(Sender: TObject);
		procedure CustomCheckClick(Sender: TObject);
		procedure RadioGroup1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
	private
		{ Private declarations }
	public
		procedure AdjustCustomInstall;
		procedure GenerateXML;
		{ Public declarations }
	end;

var
	Form1: TForm1;

implementation
uses
  Help,
	ShellAPI,
	XMLIntf,
	XmlDoc;

{$R *.dfm}
var
	AppPath : string = '';

procedure Execute(
	Command : string;
	Params : string = '';
	StartFolder : string = '';
	Hidden : boolean = false);
var
	i : integer;
begin
	if Hidden then i := SW_HIDE else i := SW_NORMAL;
	ShellExecute(
		Form1.Handle,
		'open',
		PChar(Command),
		PChar(Params),
		PChar(StartFolder),
		i);
end;

//Generates the configuration.xml
procedure TForm1.GenerateXML;
var
	XML : IXMLDocument;
	RootNode, CurNode, ProdNode : IXMLNode;

	ProdVersion : string;
	Exclude : string;
	idx : Integer;

	AFile : TextFile;
begin

	ProdVersion := 'O365SmallBusPremRetail'; //Need a default for Download Mode
	Exclude := '';

	case RadioGroup1.ItemIndex of
	0: ProdVersion := 'HomeStudentRetail';
	1: ProdVersion := 'HomeBusinessRetail';
	2: ProdVersion := 'O365HomePremRetail';
	3: ProdVersion := 'O365ProPlusRetail';
	4: ProdVersion := 'O365BusinessRetail';
	5: ProdVersion := 'O365SmallBusPremRetail';
	end;



	XML := NewXMLDocument;
	XML.Options := [doNodeAutoIndent];
	RootNode := XML.AddChild('Configuration');
		CurNode := RootNode.AddChild('Add');

    case ComboBox1.ItemIndex of
    0: CurNode.Attributes['SourcePath'] := '.\2013';
    1: CurNode.Attributes['SourcePath'] := '.\2016';
    end;
		CurNode.Attributes['OfficeClientEdition'] := '32';
			ProdNode := CurNode.AddChild('Product');
			//Product
			ProdNode.Attributes['ID'] := ProdVersion;
				CurNode := ProdNode.AddChild('Language');
				CurNode.Attributes['ID'] := 'en-US';

				//Exclude App Section for custom install
				if (CustomCheck.Checked) then
				begin
					for idx := 0 to CheckListBox1.Count-1 do
					begin
						if (CheckListBox1.ItemEnabled[idx]) then
						begin
							if not (CheckListbox1.Checked[idx]) then
							begin
								Exclude := CheckListBox1.Items.Strings[idx];
								if Exclude = 'OneDrive' then Exclude := 'Groove';
								CurNode := ProdNode.AddChild('ExcludeApp');
								CurNode.Attributes['ID'] := Exclude;
							end;
						end;
					end;
				end;
		//Extra common goodies
		CurNode := RootNode.AddChild('Display');
		CurNode.Attributes['AcceptEULA'] := 'TRUE';
		CurNode := RootNode.AddChild('Property');
		CurNode.Attributes['Name'] := 'AutoActivate';
		CurNode.Attributes['Value'] := '1';


	//We don't need the XML Header.  Just dump the xml
	AssignFile(AFile, 'install.xml');
	ReWrite(AFile);
	Writeln(AFile,XML.DocumentElement.XML);
	CloseFile(AFile);

end;

procedure TForm1.Button1Click(Sender: TObject);
var
  Exec : string;
  Switch : string;
  Hide : boolean;
begin
	SetCurrentDirectory(PChar(AppPath+'Setup\'));
	GenerateXML;
  Hide := false;

  case ComboBox1.ItemIndex of
  0: Exec := 'setup2013.exe';
  1: Exec := 'setup2016.exe';
  end;

  case ComboBox2.ItemIndex of
  0:
    begin
      Hide := false;
      Switch := '/download install.xml';
    end;
  1:
    begin
      Hide := true;
      Switch := '/configure install.xml';
    end;
  end;

  Execute(Exec,Switch,'',Hide);


	Application.Minimize;
	Sleep(15000);
	Close;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  HelpForm.ShowModal;
end;

procedure TForm1.AdjustCustomInstall;
var
	idx : integer;
	Cidx : integer;
	//Limit : set of 0..9;
	Prog : TStringlist;
begin
	//Disable All
	Prog := TStringList.Create;
	Prog.Clear;
	CheckListBox1.ItemIndex := -1;
	for idx := 0 to CheckListBox1.Count -1 do
	Begin
		CheckListBox1.ItemEnabled[idx] := false;
		CheckListBox1.Checked[idx] := false;
	end;

	case RadioGroup1.ItemIndex of
	0: Prog.CommaText := 'Excel,OneNote,PowerPoint,Word';
	1: Prog.CommaText := 'Excel,OneNote,Outlook,PowerPoint,Word';
	2: Prog.CommaText := 'Access,Excel,OneNote,Outlook,PowerPoint,Publisher,Word';
	3: Prog.CommaText := 'Access,Excel,InfoPath,Lync,OneNote,Outlook,PowerPoint,Publisher,Word,OneDrive';
	4: Prog.CommaText := 'Excel,OneNote,Outlook,PowerPoint,Publisher,Word,OneDrive';
	5: Prog.CommaText := 'Access,Excel,InfoPath,Lync,OneNote,Outlook,PowerPoint,Publisher,Word,OneDrive';
	end;

	//Yes, Number searching is nanoseconds faster, however with Microsoft coming up
	//with new SKU combinations, it's just tons easier for this little program
	//to search by names
	(*case RadioGroup1.ItemIndex of
	0: Limit := [0..3];
	1: Limit := [0..4];
	2: Limit := [0..6];
	3: Limit := [0..9];
	4: Limit := [0..5,7];
	5: Limit := [0..9];
	end;
	{for idx := 0 to CheckListBox1.Count -1 do
	begin
		if idx in Limit then
		begin
			CheckListBox1.Checked[idx] := true;
			CheckListBox1.ItemEnabled[idx] := true;
		end;
	end;
	*)

	for idx := 0 to Prog.Count -1 do
	begin
		Cidx := CheckListBox1.Items.IndexOf(Prog[idx]);
		if Cidx <> - 1 then
		begin
			CheckListBox1.Checked[Cidx] := true;
			CheckListBox1.ItemEnabled[Cidx] := true;
		end;
	end;

 Prog.Free;

end;

procedure TForm1.CustomCheckClick(Sender: TObject);
begin
	if (Self.CustomCheck.Checked = true) then
	begin
		CheckListbox1.Color := clWindow;
		CheckListBox1.Enabled := true;
	end else
	begin
		CheckListBox1.Color := clBtnFace;
		CheckListbox1.Enabled := False;
	end;
	AdjustCustomInstall;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
	Left:=(Screen.Width-Width)  div 2;
	Top:=(Screen.Height-Height) div 2;
	AppPath := ExtractFilePath(ParamStr(0));
	AdjustCustomInstall;
end;

procedure TForm1.RadioGroup1Click(Sender: TObject);
begin
	AdjustCustomInstall;
end;

end.
