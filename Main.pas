unit Main;

interface

uses
	Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
	Dialogs, StdCtrls, ExtCtrls, Vcl.CheckLst;

type
	TMainform = class(TForm)
		RunBtn: TButton;
		VersionList: TRadioGroup;
		SoftwareList: TCheckListBox;
		CustomCheck: TCheckBox;
		VerBox: TComboBox;
		InstrBtn: TButton;
		TaskBox: TComboBox;
		procedure RunBtnClick(Sender: TObject);
		procedure FormCreate(Sender: TObject);
		procedure CustomCheckClick(Sender: TObject);
		procedure VersionListClick(Sender: TObject);
		procedure InstrBtnClick(Sender: TObject);
	private
		{ Private declarations }
	public
		procedure AdjustCustomInstall;
		procedure GenerateXML;
		{ Public declarations }
	end;

var
	Mainform: TMainform;

implementation
uses
	Help,
	ShellAPI,
	XMLIntf,
	XmlDoc;

{$R *.dfm}
var
	AppPath : string = '';
	CD : boolean;

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
		Mainform.Handle,
		'open',
		PChar(Command),
		PChar(Params),
		PChar(StartFolder),
		i);
end;

//Generates the configuration.xml
procedure TMainform.GenerateXML;
var
	XML : IXMLDocument;
	RootNode, CurNode, ProdNode : IXMLNode;

	ProdVersion : string;
	Exclude : string;
	idx : Integer;

	AFile : TextFile;
	AFilePath : string;
begin

	ProdVersion := 'O365SmallBusPremRetail'; //Need a default for Download Mode
	Exclude := '';
	AFilePath := '';

	case VersionList.ItemIndex of
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
		//forced indent so I can follow flow
		CurNode := RootNode.AddChild('Add');

		case VerBox.ItemIndex of
		0: CurNode.Attributes['SourcePath'] := AppPath+'Setup\2013\';
		1: CurNode.Attributes['SourcePath'] := AppPath+'Setup\2016\';
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
					for idx := 0 to SoftwareList.Count-1 do
					begin
						if (SoftwareList.ItemEnabled[idx]) then
						begin
							if not (SoftwareList.Checked[idx]) then
							begin
								Exclude := SoftwareList.Items.Strings[idx];
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
	AFilePath := 'msofficeinstall.xml';
	if CD then AFilePath := GetEnvironmentVariable('TEMP')+'\msofficeinstall.xml';

	AssignFile(AFile, AFilePath);
	ReWrite(AFile);
	Writeln(AFile,XML.DocumentElement.XML);
	CloseFile(AFile);

end;

procedure TMainform.RunBtnClick(Sender: TObject);
var
	Exec : string;
	Switch : string;
	Hide : boolean;
begin
	SetCurrentDirectory(PChar(AppPath+'Setup\'));
	GenerateXML;
	Hide := false;

	case VerBox.ItemIndex of
	0: Exec := 'setup2013.exe';
	1: Exec := 'setup2016.exe';
	end;

	case CD of
	true  : Switch := GetEnvironmentVariable('TEMP')+'\msofficeinstall.xml';
	false : Switch := 'msofficeinstall.xml';
	end;
	Switch := '"'+Switch+'"';

	case TaskBox.ItemIndex of
	0:
		begin
			Hide := false;
			Switch := '/download ' + Switch;
		end;
	1:
		begin
			Hide := true;
			Switch := '/configure ' + Switch;
		end;
	end;
	Execute(Exec,Switch,'',Hide);


	Application.Minimize;
	Sleep(15000);
	Close;
end;

procedure TMainform.InstrBtnClick(Sender: TObject);
begin
	HelpForm.ShowModal;
end;

procedure TMainform.AdjustCustomInstall;
var
	idx : integer;
	Cidx : integer;
	//Limit : set of 0..9;
	Prog : TStringlist;
begin
	//Disable All
	Prog := TStringList.Create;
	Prog.Clear;
	SoftwareList.ItemIndex := -1;
	for idx := 0 to SoftwareList.Count -1 do
	Begin
		SoftwareList.ItemEnabled[idx] := false;
		SoftwareList.Checked[idx] := false;
	end;

	case VersionList.ItemIndex of
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
		Cidx := SoftwareList.Items.IndexOf(Prog[idx]);
		if Cidx <> - 1 then
		begin
			SoftwareList.Checked[Cidx] := true;
			SoftwareList.ItemEnabled[Cidx] := true;
		end;
	end;

	Prog.Free;

end;

procedure TMainform.CustomCheckClick(Sender: TObject);
begin
	if (Self.CustomCheck.Checked = true) then
	begin
		SoftwareList.Color := clWindow;
		SoftwareList.Enabled := true;
	end else
	begin
		SoftwareList.Color := clBtnFace;
		SoftwareList.Enabled := False;
	end;
	AdjustCustomInstall;
end;

procedure TMainform.FormCreate(Sender: TObject);
begin
	CD := false;
	Self.Position := poDesktopCenter;
	AppPath := ExtractFilePath(ParamStr(0));
	if Windows.GetDriveType(PChar(ExtractFileDrive(AppPath))) = 5{DRIVE_CDRIVE} then CD := true;
	AdjustCustomInstall;
end;

//Updates custom install options
procedure TMainform.VersionListClick(Sender: TObject);
begin
	AdjustCustomInstall;
end;

end.
