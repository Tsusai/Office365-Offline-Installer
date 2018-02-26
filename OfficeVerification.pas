unit OfficeVerification;

interface

uses
	Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
	Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
	TVerifyForm = class(TForm)
		OutputBox: TMemo;
		CloseBtn: TButton;
		procedure CloseBtnClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
	private
		{ Private declarations }
	public
		{ Public declarations }
		procedure MsgNewLine(s : string);
		procedure MsgAppLine(s : string);
	end;

	procedure VerifyOfficeData(Year : integer);
	function ReadVersion(aCABfile : string) : string;

var
	VerifyForm: TVerifyForm;

implementation

{$R *.dfm}

uses
	Main,
	JvCabFile,
	XMLIntf,
	XmlDoc,
	uTPLb_Hash,
	uTPLb_CryptographicLibrary,
	uTPLb_StreamUtils;

var
	Passed : boolean;

procedure TVerifyForm.MsgNewLine(s : string);
begin
	OutputBox.Lines.Add(s);
end;

procedure TVerifyForm.CloseBtnClick(Sender: TObject);
begin
	Self.Close;
end;

procedure TVerifyForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := Passed;
end;

procedure TVerifyForm.MsgAppLine(s : string);
begin
	OutputBox.Lines.Strings[OutputBox.Lines.Count-1] :=
		OutputBox.Lines.Strings[OutputBox.Lines.Count-1] + s;
end;

function CheckFile(f : string) : boolean;
begin
	VerifyForm.MsgNewLine(
		Format('Checking for %s: File ',
		[  StringReplace(f,AppPath,'',[rfReplaceAll])  ]
		)
	);
	Result := FileExists(f);
	if Not Result then VerifyForm.MsgAppLine('Missing, Please Re-Download') else
		VerifyForm.MsgAppLine('Found');
end;

function CheckFolder(p : string) : boolean;
begin
	VerifyForm.MsgNewLine(
		Format('Checking for %s: Folder ',
		[  StringReplace(p,AppPath,'',[rfReplaceAll])  ]
		)
	);
	Result := DirectoryExists(p);
	if Not Result then VerifyForm.MsgAppLine('Missing, Please Re-Download') else
		VerifyForm.MsgAppLine('Found');
end;

function Hash_GrabHash(
	f : string
):string;
var
	AFile : TStringList;
begin
	Result := '';
	AFile := TStringList.Create;
	try
		AFile.Clear;
		//The Encoding is required, because one of the files lacks the unicode header
		AFile.LoadFromFile(f,TEncoding.Unicode);
		Result := AFile.Strings[0];
	finally
		AFile.Free;
	end;
end;

function Hash_VerifyHash(const Hash1, Hash2 : string) : boolean;
begin
	Result := (CompareText(Hash1,Hash2) = 0);
	case Result of
		true:  VerifyForm.MsgAppLine('Complete');
		false: VerifyForm.MsgAppLine('Failure');
	end;
end;

function Hash_SHA256(
	const f : string
) : string;
var
	i : integer;
	Crypto : TCryptographicLibrary;
	SHA256 : THash;
	bytes : TBytes;
	{P,} Sz: integer;
	aByte: byte;

begin
	Result := '';
	Crypto := TCryptographicLibrary.Create(nil);
	SHA256 := uTPLb_Hash.THash.Create(nil);
	try
		SHA256.CryptoLibrary := Crypto;
		SHA256.HashId := 'native.hash.SHA-256';
		SHA256.Begin_Hash;
		SHA256.HashFile(f);
		SHA256.HashOutputValue.Position := 0;
		SetLength(Bytes, 32);
		Sz := SHA256.HashOutputValue.Size;
		if Sz <> 32 then
		begin
			VerifyForm.MsgAppLine('Failure (Hash Length Issue)');
			Exit;
		end;
		SHA256.HashOutputValue.Position := 0;
		Result := '';
		for i := 0 to 31 do
		begin
			SHA256.HashOutputValue.Read(aByte,1);
			Result := Result + Format('%0.2x', [abyte]);
		end;
		SHA256.End_Hash;
		Result := LowerCase(Trim(Result));
	finally
		Crypto.Free;
		SHA256.Free;
	end;
end;

function ReadVersion(aCABfile : string) : string;
var
	CAB : TJvCABFile;
	XML: IXMLDocument;
	ANode : IXMLNode;
begin
	CAB := TJvCABFile.Create(nil);
	Result := '';
	try
		VerifyForm.MsgNewLine('Extracting VersionDescriptor.xml');
		CAB.FileName := ACABfile;
		CAB.ExtractFile('VersionDescriptor.xml',tmpPath);
		if CheckFile(tmpPath+'VersionDescriptor.xml') then
		begin
			XML := LoadXMLDocument(tmpPath+'VersionDescriptor.xml');
			ANode := XML.DocumentElement.ChildNodes['Available'];
			Result := ANode.Attributes['Build'];
			if Result <> '' then VerifyForm.MsgNewLine('Version Identified') else
			VerifyForm.MsgNewLine('Could not read Office version');
		end;
	finally
		CAB.Free;
	end;
end;

function S1_CheckFilesExist(
	const Year : integer;
	const DataPath : string;
	const ver : string
) : boolean;
begin
	Result := false;
	if not CheckFile(Format('%sv32_%s.cab',[DataPath,ver])) then Exit;
	if not CheckFolder(DataPath + ver) then Exit;
	if not CheckFile(Format('%s%s\i32%d.cab',[DataPath,ver,ConfigINI.LangID])) then Exit;
	if not CheckFile(Format('%s%s\i64%d.cab',[DataPath,ver,ConfigINI.LangID])) then Exit;
	if not CheckFile(DataPath + ver +'\s320.cab') then Exit;
	if not CheckFile(Format('%s%s\s32%d.cab',[DataPath,ver,ConfigINI.LangID])) then Exit;
	if not CheckFile(Format('%s%s\stream.x86.%s.dat',[DataPath,Ver,ConfigINI.XML_Lang])) then Exit;
	if not CheckFile(DataPath + ver +'\stream.x86.x-none.dat') then Exit;
	case Year of
	2016:
		begin
			if not CheckFile(DataPath + ver+'\i320.cab') then Exit;
			if not CheckFile(DataPath + ver+'\i640.cab') then Exit;
		end;
	end;
	//Holy Crap we made it.
	Result := true;
end;

function S2_ReadXML2016Hashes(
	//const aCABfile : string; Already Extracted
	var I320Hash : string;
	var I640Hash : string//;
) : boolean;
var
	XML: IXMLDocument;
	ANode : IXMLNode;
begin
	VerifyForm.MsgNewLine('Begin Reading Office 2016 i320 & i640 Hashes: ');
	Result := false;
	I320Hash := '';
	I640Hash := '';
//	CAB := TJvCABFile.Create(nil);
	try
		//CAB.FileName := ACABfile;
		//CAB.ExtractFile('VersionDescriptor.xml',tmpPath);
		XML := LoadXMLDocument(tmpPath+'VersionDescriptor.xml');
		ANode := XML.DocumentElement.ChildNodes['Available'];
		I320Hash := ANode.Attributes['I320Hash'];
		I640Hash := ANode.Attributes['I640Hash'];
		Result := (I320Hash <> '') AND (I640Hash <> '');
	finally
		//CAB.Free;
		//Fix this, this is so fucking lame.
		case Result of
		true:  VerifyForm.MsgAppLine('Complete');
		false: VerifyForm.MsgAppLine('Failure');
		end;
	end;
end;

function S3_Verify2016Hashes(
	const I320Hash : string;
	const I640Hash : string;
	const SourcePath : string
) : boolean;
var
	CAB : TJvCABFile;
begin
	Result := true; //presume true
	CAB := TJvCABFile.Create(nil);
	try
		CAB.FileName := SourcePath+'i320.cab';
		CAB.ExtractFile('i320.hash',tmpPath);
		if not CheckFile(tmpPath+'i320.hash') then
		begin
			result := false;
			exit;
		end;

		CAB.FileName := SourcePath + 'i640.cab';
		CAB.ExtractFile('i640.hash',tmpPath);
		if not CheckFile(tmpPath+'i640.hash') then
		begin
			VerifyForm.MsgNewLine('Could not extract i640.hash');
			result := false;
			exit;
		end;
	finally
		CAB.Free;
	end;

	if not result then exit;
	Result := false; //now we presume false;

	VerifyForm.MsgNewLine('Verifying i320 Hash: ');
	if not Hash_VerifyHash(i320Hash,Hash_GrabHash(tmpPath+'i320.Hash')) then exit;
	VerifyForm.MsgNewLine('Verifying i640 Hash: ');
	if not Hash_VerifyHash(i640Hash,Hash_GrabHash(tmpPath+'i640.Hash')) then exit;
	Result := True;
end;

function S4_VerifyStreamHash(
	const SourcePath : string;
	const CABFile : string;
	const StreamName : string
	) : boolean;
var
	CAB : TJvCabFile;
	sourceLocalHash : string;
	FileHash : string;
begin
	Result := false;
	CAB := TJvCabFile.Create(nil);
	try
		CAB.FileName := SourcePath + CABFile;
		CAB.ExtractFile(StreamName+'.hash',tmpPath);
		CheckFile(tmpPath+StreamName+'.hash');
		sourceLocalHash := Hash_GrabHash(tmpPath+StreamName+'.hash');
		VerifyForm.MsgNewLine('Hashing ' + StreamName+'.dat: ');
		FileHash := Hash_SHA256(SourcePath + StreamName+'.dat');
		if FileHash = '' then exit;
		VerifyForm.MsgAppLine('Hashed, Comparing: ');
		Result := Hash_VerifyHash(sourceLocalHash,FileHash);
	finally
		CAB.Free;
	end;
end;

//NOTE:
//CANNOT VERIFY 2013 AT THIS TIME. Leaving code multi-year for future changes
procedure VerifyOfficeData(Year : integer);
var
	ver : string;
	DataPath : string;
	I320Hash : string;
	I640Hash : string;
begin
	ver := '';
	DataPath := '';
	I320Hash := '';
	I640Hash := '';
	Passed := false;
	DataPath := Format('%sSetup\%d\',[AppPath,Year]);
	VerifyForm.Show;
	while not Passed do
	begin
		VerifyForm.MsgNewLine('Starting verification of Office '+ IntToStr(Year));

		//Check Files existance (Starting point of verification)
		if not CheckFile(Format('%sSetup\setup%d.exe',[AppPath,Year])) then exit;
		if not CheckFile(DataPath + 'Office\Data\v32.cab') then Exit;
		//Get our Version
		ver := ReadVersion(DataPath + 'Office\Data\v32.cab');

		if not S1_CheckFilesExist(Year, DataPath + 'Office\Data\',ver) then break;

		//That's it for 2013 atm


(*                      .. ......~~?$+ZZ+~,,.:==?~,...... ..
                     . ..~+$DDD8NDDDDNN88ND8N8DDDDZZ?~:+....
                   ....+Z8DNNNNMNNNNDN8O8ON88DNNNDDDDD8O$+,~.
                 ..:$8DDNNDNNNMNMNNNNNNNNNDNMNNNNNDDNDNDN8Z7,...
               ..,88DNDDNNNMMMMMMMNNNMNMMMMMNNNNNNNNNDDNDD88D+~.
            ...IODDDDDNMNMNMMNNMMNMMMMMMNMMNMNMNNNNNNNDDDNNNO88: ..
         ... ::NDDNDNNNNMNMMMMMMNMMMMMNMMMMMMMMNNNNNNDNNNDNND888??.
          ..?DDNDDDNNNNMNMMMMMNMMMMMMMNMMMMMMMMMMMMNNDNNNDDND8D8O=:..
         ..78DDDDNNNNNMNMMMMNNMMNMMMMMMMMMMMMNNMNNNNN8DDNNNNNDNDOD$+, .
         ..8DNDDDNNNNNNMMNNNMNMMMMMMMNMMMMMMMNMNNNNNNNDNNNNNNDDDD8O7..
       ..,Z8DDDDDNDNNNNNNNMNMMMMMMMNMMMMNNMNMNDNNNNNNNDDNDNNNNDDD8DO+,..
       .,7D888DDDNDDNNNNNNNNMNNNNNNNMMMMMMNNNNNNDNNDNDDNDNDDNNNDDD8DO?..
      ..=D8NDNDDDDDDDNNDDNNNNNNNNNMNMMNMNNNNNNNNNNNNNNDDNDNNNNNDDDD88O...
       ,O8DDDDND8D8NMN8DNNNNNNNNNMNMMMNNNNNNNNND8DDNNDNDNNDDDDNNNNND8OO..
    . .=8DDDNNNNOO8MNDDNNNMNNDNDDNNNMMNNNNNNDDDD8D8NNNDNNDDNNDDDNDNNDOZ7.
    ...DONNDDNNDDDNDNDDN8NNDNDDNDNNNMNNDDNDDDD8OOD8DDNN8NNDNNDNNNNNNN8ZZ.
    ..,O88NNDNNDDDDDDDDDN8DDDDNNNNMMMNDDDO8DDOZ8OOO8DNDDNNDNDNDNDNNNDDOZ~
    . 8O8NNNNND8DDDNDDDDDDNDDDDDDNNNNDD888D8OZZ$OODDDDO8NDNNDNN8NNNNND8O?..
    ..D8NNNNNNN8NNNDDDDD8DDD8D8NNNNDD8OZO88ZZ7$$8Z8D8$OO8OONDDNNDNDNN88O+
    ..Z8NDNNNNNNDDDDDND8NND8DDDDNN8OZZOO88OZ$$$O$Z88ZOZZZZZODDDNN8ND8OZ7:
    ..ONNNNNNNDNDNNNNNDNNND88D8DDDO$+I$8OO?7$8$$O88ZZZ$$$$$$8DDD8O88OO8O
     .88NDNNNDDDNNNNNNDDD88NDDDDD8Z$I$OZ$Z7IIOOZO8ZO$777777IO8NDDDO888OO.
     .IDDNDNNDDDNDD8DD$?7INNNN8N$87$O$$7I7OO$$Z$ZZ7777IIIIII$8DNDN8NDD$$
      .Z8DD88DDOOZ$Z7?7$8DDNNNZZZ$Z$7I7II??I777$77IIIIIIIIIIIODDN8D8DO$$.
       OD8DDDO8?+++?+++??+??+?III?++==+????7IIII????III?II???ODDO8D8DO$O.
       .8ODNDZ7??+?++++??I?+?I7?++==+=++7?7ZZO8Z77$$$$I??I???7DDDNDDZZO7.
      ..ZO7NDZ??????7ZO88DONDD8D$I++++?$DNDDDD8DNNDNDNNDZII??I$NNDD88ZO$
      . D$8DDZ???IZDDD888NNNNN8ZOI+++??$O8DNNND8OOZ$Z888O$I??I7NMMD8OOZZ.
     . 8D88NNZI??DD887???ZOZOZ777I++++?7Z7ZD$DOZDO8NNDZ77??MNN,.~ND8OZOI.
    . .O888DDOI?IDZI?+$I?D$IZZ7$II+=++?$ZOIOD7INMN8$8$?I7$7INM,~DDDZ$7?O~
    .  .+7Z8NN??I77++ZD=?8O8OOOOO$7I$DMNDOIDI~$DDO7?$Z$III??+???OD8Z7IZZ?
     ....IZONMDNMO?=+OO=$8MD8+7I7D7??I7$MZZ$7$777$$$?IIIIII??I??788OO7$ZI
      ..:?7Z8DN??????I77I=+I7$$$?7Z???77ONN7I?IIIIIIIIIIIIII?I???8O$$7?Z+
    .....7IZ?NO+???????I?????I??=ZZI+?I7$M8ZI+?+++???IIIIII?II??+$8Z8$7$.
       ..+?I$D=+????????+++==++++DII??IIIZM$I??+++++????I?M?I???++?78Z7+.
    .... ,+?+D~Z+??+++++++===+++=ZI????II7$8I?+?==?$N8I+IIIIII??++=I8O$I.
    .    .=+=8===+++++++==~=====Z?II??I7I7I$7II???++??????I?????++=O8Z?~
         ..+=$~+=Z?+++??7O8Z+=???III??I777I$7I????++++?????????+?+=O7+7..
         ..+=?==+?+?+?++=======++III???I77$$7I???++???????????????=II?...
         ...+??~++++++++=======++?II+++III7777I???+???????????????+7I=. .
         ....~+~=+++++++=====+++?I??+~=???7I787???????????????????+.. .
         ....=+~=?+?+++++++==+=+?I+?+++?I7M8$Z$I??????????????????+..
          .  .+:=+?++++++++++++++?77III7$Z$ZZ$77I?????????????????=.
              .~==?+++++++++++++??III7$ZOZ$$7$777I???????????????+=.
              ..+=??+++++++++++???III?7$$I$77777I7II??????I??????+~
              ..:++?+++++++++??I?IIIII$III?I7I7II7III??????I??I?+~,
              . .=+?++++?+++?II?I????+?+++??I77I77IIII???III????=~ .
              .  ~=???????+??I?I7I?????++?77$7$$$7Z$II???III???+=:
                 .~++????????I?7I$Z88OZZZ$7II777777IIII??II???+::,.
                  .~+?????++II???III?????I?IIIII777I777II?I??+=:=,
                  .:~+???????IIIIII??+????II777$777777IIIIII+~==~=..
                   ..~+??????II??III$7$Z$Z$$7777777III7II?I=~?+=+=,...
                   . .~+???????IIIII?IIIIIIII?IIII7II7IIIII?77????7MN..
                      .:+????I?IIII??++=++?+??+?IIIIII77$$7$$I7??D+=NO...
                   .    .+?IIII?II?++==+?=+=??IIIIIIII$ZZZZ$77?ID?+8D8$.
                     .  ..?77III????+++=?????IIIIIII7OZZZZ$$77I???7NDODO.
                        ..==?I$III?????????IIII777ZZOZZZZ$$7O7Z+?NNDNNDD$...
                         .~?+I77$7II?I?I?I777$$Z$ZOOOZZ$$$7O77+?DNNDNNDDDD....
                        ..?+??II77$ZZ$$ZZZZZZZZZZZOO$$7$$Z$8===DNDNNNDNNND8...
                       . :I~????II777$$$Z$ZZZZZZZ$$$777$8Z==+MN8NNNNDNNN8NDO....
                      ..Z,:D???IIIIII77$$Z$$$$Z$77III7ZZ===~NNNNNDNNNNDNNDD8D$..
                    .. 7N~::??IIIIIIII777777777I??IIII=~~~ODNNDDNNNN8NNNN8ODND88
                  .....NN~~~II?IIIIII?III7III?????IO::::~NNDNNNNNDNNNNDO88DNDNND
                 ,88D8DN8~~~=I?IIIII?III7I??++++?7~::::.DNDNNNDNNNNN8NNNNNNMNDDN

          DON'T YOU KNOW IT'S 2016?! I MEAN COME ON

*)
		if Year = 2016 then
		begin
			if not S2_ReadXML2016Hashes(
				//DataPath + 'Office\Data\v32.cab',
				I320Hash,I640Hash) then break;

			if not S3_Verify2016Hashes(
				I320Hash,I640Hash,
				DataPath + 'Office\Data\'+ver+'\'
			)then break;

			//Verify the damn streams because Microsoft doesn't bother and nothing pisses
			//me off like copying to USB/DVD for it to be fucking bad.
			{if not S4_VerifyStreamHash(
				DataPath + 'Office\Data\'+ver+'\',
				's321033.cab',
				'stream.x86.'+ ConfigINI.XML_Lang
			) then break;}
			if not S4_VerifyStreamHash(
				Format('%sOffice\Data\%s\',[DataPath,ver]),
				Format('s32%d.cab',[ConfigINI.LangID]),
        Format('stream.x86.%s',[ConfigINI.XML_Lang])
			) then break;
			if not S4_VerifyStreamHash(
				DataPath + 'Office\Data\'+ver+'\',
				's320.cab',
				'stream.x86.x-none'
			) then break;
		end;
		VerifyForm.MsgNewLine(Format('Verification of Office %d complete',[Year]));
		Passed := true;
	end;
end;

end.
