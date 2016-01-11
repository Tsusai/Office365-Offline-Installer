unit OfficeVerification;

interface
	procedure VerifyOfficeData(Year : integer);

implementation
uses
	Main,
	JvCabFile,
	Classes,
	SysUtils,
	XMLIntf,
	XmlDoc,
	uTPLb_Hash,
	uTPLb_CryptographicLibrary,
	uTPLb_StreamUtils,
	Dialogs;

function CheckFile(f : string) : boolean;
begin
	Result := FileExists(f);
	if Not Result then ShowMessage(f+' Missing, redownload');
end;

function CheckFolder(p : string) : boolean;
begin
	Result := DirectoryExists(p);
	if Not Result then ShowMessage('Folder ' + p + ' Missing, redownload');
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
			ShowMessage('Something went bad with with file hashing, no idea why' + f);
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

function S1_ReadXMLVersion(
	var I320Hash : string;
	var I640Hash : string;
	var Ver : string;
	const Year : integer;
	const VDXML : string
) : boolean;
var
	XML: IXMLDocument;
	ANode : IXMLNode;
begin
	Result := false;
	I320Hash := '';
	I640Hash := '';
	ver := '';
	try
		XML := LoadXMLDocument(VDXML);
		ANode := XML.DocumentElement.ChildNodes['Available'];
		ver := ANode.Attributes['Build'];
		case Year of
		2013: Result := (ver <> '');
		2016:
			begin
				I320Hash := ANode.Attributes['I320Hash'];
				I640Hash := ANode.Attributes['I640Hash'];
				Result := (ver <> '') AND (I320Hash <> '') AND (I640Hash <> '');
			end;
		end;
	finally
		if not Result then ShowMessage('Could not read '+VDXML);
	end;
end;

function S2_CheckFilesExist(
	const Year : integer;
	const DataPath : string;
	const ver : string
) : boolean;
begin
	Result := false;
	if not CheckFile(DataPath + 'Office\Data\v32_'+ver+'.cab') then Exit;
	if not CheckFolder(DataPath + 'Office\Data\'+ver) then Exit;
	if not CheckFile(DataPath + 'Office\Data\'+ver+'\i321033.cab') then Exit;
	if not CheckFile(DataPath + 'Office\Data\'+ver+'\i641033.cab') then Exit;
	if not CheckFile(DataPath + 'Office\Data\'+ver+'\s320.cab') then Exit;
	if not CheckFile(DataPath + 'Office\Data\'+ver+'\s321033.cab') then Exit;
	if not CheckFile(DataPath + 'Office\Data\'+ver+'\stream.x86.en-us.dat') then Exit;
	if not CheckFile(DataPath + 'Office\Data\'+ver+'\stream.x86.x-none.dat') then Exit;
	if Year = 2016 then
	begin
		if not CheckFile(DataPath + 'Office\Data\'+ver+'\i320.cab') then Exit;
		if not CheckFile(DataPath + 'Office\Data\'+ver+'\i640.cab') then Exit;
  end;

	//Holy Crap we made it.
	Result := true;
end;

function S3_Verify2016Hashes(
	const CABFILE : TJvCABFile;
	const I320Hash : string;
	const I640Hash : string;
	const DPath : string;
	const Ver : string
) : boolean;
var
	SourcePath : string;
begin
	Result := false;
	SourcePath := (DPath + 'Office\Data\'+ver+'\');
	try
		CABFile.FileName := SourcePath+'i320.cab';
		CABFile.ExtractFile('i320.hash',tmpPath);
	except
		ShowMessage('Could not open '+SourcePath+'i320.cab');
		CABFile.Free;
		Exit;
	end;
	try
		CABFile.FileName := SourcePath + 'i640.cab';
		CABFile.ExtractFile('i640.hash',tmpPath);
	except
		ShowMessage('Could not open '+ SourcePath+ 'i640.cab');
		CABFile.Free;
		Exit;
	end;

	if not Hash_VerifyHash(i320Hash,Hash_GrabHash(tmpPath+'i320.Hash')) then
	begin
		ShowMessage('Could not verify hash for '+tmpPath+'i320.Hash');
		exit;
	end;
	if not Hash_VerifyHash(i640Hash,Hash_GrabHash(tmpPath+'i640.Hash')) then
	begin
		ShowMessage('Could not verify hash for '+tmpPath+'i640.Hash');
		exit;
	end;
  Result := True;
end;

function S4_VerifyLStreamHash(
	const CABFILE : TJvCABFile;
	const DPath : string;
	const Ver : string
) : boolean;
var
	SourcePath : string;
	sourceLocalHash : string;
	FileHash : string;
begin
	Result := false;
	try
		SourcePath := (DPath + 'Office\Data\'+ver+'\');
		CABFile.FileName := SourcePath + 's321033.cab';
		CABFile.ExtractFile('stream.x86.en-us.hash',tmpPath);
		sourceLocalHash := Hash_GrabHash(tmpPath+'stream.x86.en-us.hash');
		FileHash := Hash_SHA256(SourcePath + 'stream.x86.en-us.dat');
		if FileHash = '' then exit;
		Result := Hash_VerifyHash(sourceLocalHash,FileHash);
  finally
		if not Result then ShowMessage('Failed to verify stream.x86.en-us.dat');
	end;

end;

function S5_VerifyUStreamHash(
	const CABFILE : TJvCABFile;
	const DPath : string;
	const Ver : string
) : boolean;
var
	SourcePath : string;
	sourceLocalHash : string;
	FileHash : string;
begin
	Result := false;
	try
		SourcePath := (DPath + 'Office\Data\'+ver+'\');
		CABFile.FileName := SourcePath + 's320.cab';
		CABFile.ExtractFile('stream.x86.x-none.hash',tmpPath);
		sourceLocalHash := Hash_GrabHash(tmpPath+'stream.x86.x-none.hash');
		FileHash := Hash_SHA256(SourcePath + 'stream.x86.x-none.dat');
		if FileHash = '' then exit;
		Result := Hash_VerifyHash(sourceLocalHash,FileHash);
	finally
		if not Result then ShowMessage('Failed to verify stream.x86.x-none.dat');
	end;
end;

//NOTE:
//CANNOT VERIFY 2013 AT THIS TIME. Leaving code multi-year for future changes
procedure VerifyOfficeData(Year : integer);
var
	CABFile : TJvCABFile;
	i : integer;
	ver : string;
	DataPath : string;
	I320Hash : string;
	I640Hash : string;
begin
	ver := '';
	DataPath := '';
	I320Hash := '';
	I640Hash := '';
	DataPath := Format('%sSetup\%d\',[AppPath,Year]);

	//Check if v32.cab exists (Starting point of verification)
	if not CheckFile(DataPath + 'Office\Data\v32.cab') then exit;

	CABFile := TJvCABFile.Create(nil);
	try
		CABFile.FileName := DataPath + 'Office\Data\v32.cab';
	except
		ShowMessage('Could not open '+ DataPath + 'Office\Data\v32.cab');
		CABFile.Free;
		Exit;
	end;
	//ShowMessage('File Count: ' + IntToStr(i));
	CABFile.ExtractFile('VersionDescriptor.xml',tmpPath);
	for I := 0 to 0 do //run once? lol
	begin

		//Get a Build #
		if not S1_ReadXMLVersion(
			I320Hash,I640Hash,Ver,
			Year,tmpPath+'VersionDescriptor.xml'
		) then break;

		//Verify files exists in that version's folder
		if not S2_CheckFilesExist(Year,DataPath,ver) then break;

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
			if not S3_Verify2016Hashes(
				CABFile,
				I320Hash,I640Hash,
				DataPath,
				ver
			) then break;
			//Verify the damn streams because Microsoft doesn't bother and nothing pisses
			//me off like copying to USB/DVD for it to be fucking bad.
			if not S4_VerifyLStreamHash(CABFile,DataPath,Ver) then break;
			if not S5_VerifyUStreamHash(CABFile,DataPath,Ver) then break;
		end;

		ShowMessage('Office ' + IntToStr(Year) + ' Verified');

	end;
	CABFile.Free;
end;



end.
