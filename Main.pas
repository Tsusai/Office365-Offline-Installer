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
		procedure Button1Click(Sender: TObject);
		procedure FormCreate(Sender: TObject);
		procedure CustomCheckClick(Sender: TObject);
		procedure RadioGroup1Click(Sender: TObject);
	private
		{ Private declarations }
	public
		procedure AdjustCustomInstall;
		procedure GenerateXML(Install : boolean);
		procedure Download;
		{ Public declarations }
	end;

var
	Form1: TForm1;

implementation
uses
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
procedure TForm1.GenerateXML(Install : boolean);
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

	if Install then
	begin
		case RadioGroup1.ItemIndex of
		0: ProdVersion := 'HomeStudentRetail';
		1: ProdVersion := 'HomeBusinessRetail';
		2: ProdVersion := 'O365HomePremRetail';
		3: ProdVersion := 'O365ProPlusRetail';
		4: ProdVersion := 'O365BusinessRetail';
		5: ProdVersion := 'O365SmallBusPremRetail';
		end;
	end;



	XML := NewXMLDocument;
	XML.Options := [doNodeAutoIndent];
	RootNode := XML.AddChild('Configuration');
		CurNode := RootNode.AddChild('Add');
		CurNode.Attributes['OfficeClientEdition'] := '32';
			ProdNode := CurNode.AddChild('Product');
			//Product
			ProdNode.Attributes['ID'] := ProdVersion;
				CurNode := ProdNode.AddChild('Language');
				CurNode.Attributes['ID'] := 'en-US';

				//Exclude App Section for custom install
				if ((CustomCheck.Checked) and (Install)) then
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

procedure TForm1.Download;
begin
	SetCurrentDirectory(PChar(AppPath+'Setup\'));
	GenerateXML(true);
	Execute('setup.exe', '/download install.xml');
	Application.Minimize;
	Sleep(15000);
	Application.Terminate;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
	SetCurrentDirectory(PChar(AppPath+'Setup\'));
	GenerateXML(true);
	Execute('setup.exe', '/configure install.xml','',true);
	Application.Minimize;
	Sleep(15000);
	Close;
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
	if (FindCmdLineSwitch('download',true)) then
	begin
		ShowMessage('Download switch detected.  Downloading Latest Office 2013/365');
		Download;
	end;

end;

procedure TForm1.RadioGroup1Click(Sender: TObject);
begin
	AdjustCustomInstall;
end;

end.
