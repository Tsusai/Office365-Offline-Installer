unit Main;

interface

uses
	Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
	Dialogs, StdCtrls, ExtCtrls, Vcl.CheckLst;

type
	TMainform = class(TForm)
		VersionList: TRadioGroup;
		SoftwareList: TCheckListBox;
		CustomCheck: TCheckBox;
		VerBox: TComboBox;
		InstrBtn: TButton;
		TaskBox: TComboBox;
		VerifyBtn: TButton;
		UpdateBtn: TButton;
		RunBtn: TButton;
		Label1: TLabel;
		procedure RunBtnClick(Sender: TObject);
		procedure FormCreate(Sender: TObject);
		procedure CustomCheckClick(Sender: TObject);
		procedure VersionListClick(Sender: TObject);
		procedure InstrBtnClick(Sender: TObject);
		procedure VerifyBtnClick(Sender: TObject);
		procedure FormClose(Sender: TObject; var Action: TCloseAction);
	private
		{ Private declarations }
	public
		procedure AdjustCustomInstall;
		procedure GenerateXML;
		{ Public declarations }
	end;

type
	TConfigINI = record
		XML_Lang : string;
	end;

var
	Mainform: TMainform;
	AppPath : string = '';
	tmpPath : string;
	ConfigINI : TConfigINI;

implementation
uses
	Help,
	Wait,
	OfficeVerification,
	ShellAPI,
	XMLIntf,
	IOUtils,
	IniFiles,
	XmlDoc;

{$R *.dfm}

(*procedure Execute(
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
end;*)

procedure ExecuteAndWait(const aCommando: string);
var
	tmpStartupInfo: TStartupInfo;
	tmpProcessInformation: TProcessInformation;
	tmpProgram: String;
begin
	tmpProgram := trim(aCommando);
	FillChar(tmpStartupInfo, SizeOf(tmpStartupInfo), 0);
	with tmpStartupInfo do
	begin
		cb := SizeOf(TStartupInfo);
		//wShowWindow := SW_HIDE;
		wShowWindow := SW_NORMAL;
	end;

  //Probably helpful if I don't hide the damn office.
	if CreateProcess(nil, pchar(tmpProgram), nil, nil, true, NORMAL_PRIORITY_CLASS{CREATE_NO_WINDOW},
		nil, nil, tmpStartupInfo, tmpProcessInformation) then
	begin
		// loop every 10 ms
		while WaitForSingleObject(tmpProcessInformation.hProcess, 10) > 0 do
		begin
			Application.ProcessMessages;
		end;
		CloseHandle(tmpProcessInformation.hProcess);
		CloseHandle(tmpProcessInformation.hThread);
	end
	else
	begin
		RaiseLastOSError;
	end;
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
	tmp : string;
begin

	ProdVersion := 'O365SmallBusPremRetail'; //Need a default for Download Mode
	Exclude := '';
	AFilePath := '';
	tmp := '';

	case VersionList.ItemIndex of
	0: ProdVersion := 'HomeStudentRetail';
	1: ProdVersion := 'HomeBusinessRetail';
	2: ProdVersion := 'O365HomePremRetail';
	3: ProdVersion := 'O365ProPlusRetail';
	4,5: ProdVersion := 'O365BusinessRetail';
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
				CurNode.Attributes['ID'] := ConfigINI.XML_Lang;

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
			if VersionList.ItemIndex = 5 then
			begin
				CurNode := XML.DocumentElement.ChildNodes['Add'];
				ProdNode := CurNode.AddChild('Product');
				case VerBox.ItemIndex of
				0: ProdNode.Attributes['ID'] := 'LyncRetail'; //2013
				1: ProdNode.Attributes['ID'] := 'SkypeforBusinessRetail'; //2016
				end;

				CurNode := ProdNode.AddChild('Language');
				CurNode.Attributes['ID'] := ConfigINI.XML_Lang;
			end;
		//Extra common goodies
		CurNode := RootNode.AddChild('Display');
		CurNode.Attributes['AcceptEULA'] := 'TRUE';
		CurNode := RootNode.AddChild('Property');
		CurNode.Attributes['Name'] := 'AutoActivate';
		CurNode.Attributes['Value'] := '1';
		CurNode := RootNode.AddChild('Logging');
		CurNode.Attributes['Type'] := 'Standard';
		DateTimeToString(tmp,'log-hhmmss',Now);
		CurNode.Attributes['Path'] := tmpPath;
		CurNode.Attributes['Template'] := 'O365Offline(*).txt';


	//We don't need the XML Header.  Just dump the xml
	AFilePath := tmpPath + 'msofficeinstall.xml';

	AssignFile(AFile, AFilePath);
	ReWrite(AFile);
	Writeln(AFile,XML.DocumentElement.XML);
	CloseFile(AFile);

end;

procedure TMainform.RunBtnClick(Sender: TObject);
var
	Exec : string;
	Switch : string;
begin
	if TaskBox.ItemIndex = 0 then exit;

	(*if (TaskBox.ItemIndex = 2) and
	(Verbox.ItemIndex = 0) and
	(VersionList.ItemIndex = 5) then
	begin
		ShowMessage('365 Business + Skype is untested for Office 2013.');
		Exit;
	end;*)


	SetCurrentDirectory(PChar(AppPath+'Setup\'));
	GenerateXML;

	case VerBox.ItemIndex of
	0: Exec := 'setup2013.exe';
	1: Exec := 'setup2016.exe';
	end;

//	Switch := '"' + tmpPath + 'msofficeinstall.xml"';

	{case TaskBox.ItemIndex of
	1:Switch := '/download ' + Switch;
	2:Switch := '/configure ' + Switch;
	end;}
	case TaskBox.ItemIndex of
	1:Switch := '/download';
	2:Switch := '/configure';
	end;

	WaitForm.Show;
	Application.ProcessMessages;
	Sleep(1000);

	MainForm.Enabled := false;
	ExecuteAndWait(
		Format('%s %s %smsofficeinstall.xml',[Exec,Switch,tmpPath])
	);
	//Sleep(10000);
	WaitForm.Close;
	MainForm.Enabled := true;
	MainForm.BringToFront;
	//Execute(Exec,Switch,'',Hide);
	//Application.Minimize;
	//Sleep(15000);
	//Close;
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
	4,5: Prog.CommaText := 'Excel,OneNote,Outlook,PowerPoint,Publisher,Word,OneDrive';
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

procedure TMainform.VerifyBtnClick(Sender: TObject);
begin
	VerifyOfficeData(2013);
	VerifyOfficeData(2016);
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

procedure TMainform.FormClose(Sender: TObject; var Action: TCloseAction);
var
	iniFile : TIniFile;
begin
	TDirectory.Delete(tmpPath,true);
	iniFile := TIniFile.Create(ChangeFileExt(Application.ExeName,'.ini'));
	try
		iniFile.WriteString('XML','LanguageCode',ConfigINI.XML_Lang);
	finally
		iniFile.Free;
	end;
end;

function GetAppVersionStr: string;
var
	Exe: string;
	Size, Handle: DWORD;
	Buffer: TBytes;
	FixedPtr: PVSFixedFileInfo;
begin
	Exe := ParamStr(0);
	Size := GetFileVersionInfoSize(PChar(Exe), Handle);
	if Size = 0 then
		RaiseLastOSError;
	SetLength(Buffer, Size);
	if not GetFileVersionInfo(PChar(Exe), Handle, Size, Buffer) then
		RaiseLastOSError;
	if not VerQueryValue(Buffer, '\', Pointer(FixedPtr), Size) then
		RaiseLastOSError;
	Result := Format(' %d.%d.%d.%d',
		[LongRec(FixedPtr.dwFileVersionMS).Hi,  //major
		 LongRec(FixedPtr.dwFileVersionMS).Lo,  //minor
		 LongRec(FixedPtr.dwFileVersionLS).Hi,  //release
		 LongRec(FixedPtr.dwFileVersionLS).Lo]) //build
end;

procedure TMainform.FormCreate(Sender: TObject);
var
	iniFile : TIniFile;
begin
	AppPath := ExtractFilePath(ParamStr(0));
	tmpPath := TPath.GetTempPath + 'O365OfflineTmp\';
	ForceDirectories(tmpPath);
	AdjustCustomInstall;
	Caption := Caption + GetAppVersionStr;
	Application.Title := Caption;
	iniFile := TIniFile.Create(ChangeFileExt(Application.ExeName,'.ini'));
	try
		ConfigINI.XML_Lang := iniFile.ReadString('XML','LanguageCode','en-US');
	finally
		iniFile.Free;
	end;
end;

//Updates custom install options
procedure TMainform.VersionListClick(Sender: TObject);
begin
	AdjustCustomInstall;
end;

end.
