unit SpComponentInstaller;

{==============================================================================
Version 3.5.5

The contents of this package are licensed under a disjunctive tri-license
giving you the choice of one of the three following sets of free
software/open source licensing terms:
  - Mozilla Public License, version 1.1
  - GNU General Public License, version 2.0
  - GNU Lesser General Public License, version 2.1

Software distributed under the License is distributed on an "AS IS" basis,
WITHOUT WARRANTY OF ANY KIND, either expressed or implied. See the License for
the specific language governing rights and limitations under the License.

The initial developer of this code is Robert Lee.
http://www.silverpointdevelopment.com

Development notes:
  - All the Delphi IDE changes are marked with: [IDE-Change].
  - All the Delphi IDE update changes are marked with: [IDE-Change-Update].
  - dcc32.exe won't execute if -Q option is not used.
    But it works fine without -Q if ShellExecute is used:
      ShellExecute(Application.Handle, 'open', DCC, ExtractFileName(PackageFilename), ExtractFilePath(PackageFilename), SW_SHOWNORMAL);
    There must be something wrong with SpExecuteDosCommand.
  - It seems that deleting the previous bpl and dcp files are not necessary.
  - Undocumented dcc32.exe -JL compiler switch, used to make Hpp files
    required for C++Builder 2006 and above.
    http://groups.google.com/group/borland.public.cppbuilder.ide/browse_thread/thread/456bece4c5665459/0c4c61ecec179ca8

==============================================================================}

interface

{$WARN SYMBOL_PLATFORM OFF}
{$WARN UNIT_PLATFORM OFF}
{$BOOLEVAL OFF} // Unit depends on short-circuit boolean evaluation

{$R 'SpComponentInstallerRes.res'}

uses
  Windows, Messages, SysUtils, Classes, Forms, Contnrs, Generics.Collections;

resourcestring
  SLogStartUnzip   = '=========================================' + #13#10 +
                     'Unzipping - Cloning' + #13#10 +
                     '=========================================';
  SLogStartExecute = '=========================================' + #13#10 +
                     'Executing patches' + #13#10 +
                     '=========================================';
  SLogStartCompile = '=========================================' + #13#10 +
                     'Compiling and installing:' + #13#10 +
                     '%s' + #13#10 +
                     '=========================================';
  SLogEnd          = '=========================================' + #13#10 +
                     'Finished' + #13#10 +
                     '=========================================';

  SLogInvalidPath = 'Error: %s doesn''t exist.';
  SLogInvalidIDE = 'Error: %s is not installed.';
  SLogNotAZip = 'Error: %s is not a zip file.';
  SLogNotAGit = 'Error: %s is not a Git repository.';
  SLogNotInstallable = '%s is not installable. Proceeding without unzipping/cloning';
  SLogCorruptedZip = 'Error: %s is corrupted.';
  SLogGitCloneFailed = 'Error: in git clone %s';

  SLogErrorCopying = 'Error copying %s to %s';
  SLogErrorDeleting = 'Error deleting %s';
  SLogErrorExecuting = 'Error executing %s';
  SLogErrorCompiling = 'Error compiling %s';

  SLogCopying = 'Copying:' +  #13#10 + '     %s' + #13#10 + 'To:' + #13#10 + '     %s';
  SLogExecuting = 'Executing:' +  #13#10 + '     %s';
  SLogExtracting = 'Extracting:' +  #13#10 + '     %s' + #13#10 + 'To:' + #13#10 + '     %s';
  SLogGitCloning = 'Git cloning:' +  #13#10 + '     %s' + #13#10 + 'To:' + #13#10 + '     %s';
  SLogCompiling = 'Compiling Package: %s';
  SLogInstalling = 'Installing Package: %s';
  SLogFinished = 'All the component packages have been successfully installed.' + #13#10 + 'Elapsed time: %f secs.';

  SGitCloneCommand = 'GIT.EXE clone --verbose --progress %s %s';

type
  TSpIDEType = (     // [IDE-Change-Update]
    ideNone,         //
    ideDelphi7,      // D7
    ideDelphi2005,   // D9
    ideDelphi2006,   // D10
    ideDelphi2007,   // D11
    ideDelphi2009,   // D12
    ideDelphi2010,   // D14
    ideDelphiXE,     // D15
    ideDelphiXE2,    // D16
    ideDelphiXE3,    // D17
    ideDelphiXE4,    // D18
    ideDelphiXE5,    // D19
    ideDelphiXE6,    // D20
    ideDelphiXE7,    // D21
    ideDelphiXE8,    // D22
    ideDelphiSeattle,// D23
    ideDelphiBerlin, // D24
    ideDelphiTokyo,  // D25
    ideDelphiRio     // D26
  );

  TSpIDETypeRec = record
    IDEVersion: string;
    IDEName: string;
    IDERegistryPath: string;
    IDERADStudioVersion: string;
  end;

const
  // [IDE-Change-Update]
  IDETypes: array [TSpIDEType] of TSpIDETypeRec = (
    (IDEVersion: 'None'; IDEName: 'None'; IDERegistryPath: 'None'; IDERADStudioVersion: ''),
    (IDEVersion: 'D7'; IDEName: 'Delphi 7'; IDERegistryPath: 'SOFTWARE\Borland\Delphi\7.0'; IDERADStudioVersion: ''),
    (IDEVersion: 'D9'; IDEName: 'Delphi 2005'; IDERegistryPath: 'SOFTWARE\Borland\BDS\3.0'; IDERADStudioVersion: '3.0'),
    (IDEVersion: 'D10'; IDEName: 'Developer Studio 2006'; IDERegistryPath: 'SOFTWARE\Borland\BDS\4.0'; IDERADStudioVersion: '4.0'),
    (IDEVersion: 'D11'; IDEName: 'RAD Studio 2007'; IDERegistryPath: 'SOFTWARE\Borland\BDS\5.0'; IDERADStudioVersion: '5.0'),
    (IDEVersion: 'D12'; IDEName: 'RAD Studio 2009'; IDERegistryPath: 'SOFTWARE\CodeGear\BDS\6.0'; IDERADStudioVersion: '6.0'),
    (IDEVersion: 'D14'; IDEName: 'RAD Studio 2010'; IDERegistryPath: 'SOFTWARE\CodeGear\BDS\7.0'; IDERADStudioVersion: '7.0'),
    (IDEVersion: 'D15'; IDEName: 'RAD Studio XE';  IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\8.0'; IDERADStudioVersion: '8.0'),
    (IDEVersion: 'D16'; IDEName: 'RAD Studio XE2'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\9.0'; IDERADStudioVersion: '9.0'),
    (IDEVersion: 'D17'; IDEName: 'RAD Studio XE3'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\10.0'; IDERADStudioVersion: '10.0'),
    (IDEVersion: 'D18'; IDEName: 'RAD Studio XE4'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\11.0'; IDERADStudioVersion: '11.0'),
    (IDEVersion: 'D19'; IDEName: 'RAD Studio XE5'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\12.0'; IDERADStudioVersion: '12.0'),
    (IDEVersion: 'D20'; IDEName: 'RAD Studio XE6'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\14.0'; IDERADStudioVersion: '14.0'),
    (IDEVersion: 'D21'; IDEName: 'RAD Studio XE7'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\15.0'; IDERADStudioVersion: '15.0'),
    (IDEVersion: 'D22'; IDEName: 'RAD Studio XE8'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\16.0'; IDERADStudioVersion: '16.0'),
    (IDEVersion: 'D23'; IDEName: 'RAD Studio 10 Seattle'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\17.0'; IDERADStudioVersion: '17.0'),
    (IDEVersion: 'D24'; IDEName: 'RAD Studio 10.1 Berlin'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\18.0'; IDERADStudioVersion: '18.0'),
    (IDEVersion: 'D25'; IDEName: 'RAD Studio 10.2 Tokyo'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\19.0'; IDERADStudioVersion: '19.0'),
    (IDEVersion: 'D26'; IDEName: 'RAD Studio 10.3 Rio'; IDERegistryPath: 'SOFTWARE\Embarcadero\BDS\20.0'; IDERADStudioVersion: '20.0')
  );

type
  TSpIDEPersonality = (persDelphiWin32, persDelphiNET, persCPPBuilder);

  TSpDelphiIDE = class
  public
    // IDE
    class function Installed(IDE: TSpIDEType): Boolean;
    class function PersonalityInstalled(IDE: TSpIDEType; IDEPersonality: TSpIDEPersonality): Boolean;
    class function StringToIDEType(S: string): TSpIDEType;
    class procedure IDEPersonalityTypeToString(A: TSpIDEPersonality; out IDERegName: string);

    // Path
    class function GetIDEDir(IDE: TSpIDEType): string;
    class function GetDCC32Filename(IDE: TSpIDEType): string;
    class function GetBPLOutputDir(IDE: TSpIDEType): string;
    class function GetDCPOutputDir(IDE: TSpIDEType): string;

    // Macros
    class function ReadEnvironmentProj(IDE: TSpIDEType; NamesAndValues: TStringList): Boolean;
    class function ExpandMacros(S: string; IDE: TSpIDEType): string;

    // SearchPath
    class function GetSearchPath(IDE: TSpIDEType; CPPBuilderPath: Boolean): string;
    class procedure AddToSearchPath(SourcesL: TStrings; IDE: TSpIDEType);
  end;

  TSpDelphiDPKFile = class
  private
    FDPKFilename: string;
    FBPLFilename: string;
    FExists: Boolean;
    FOnlyRuntime: Boolean;
    FOnlyDesigntime: Boolean;
    FDescription: string;
    FLibSuffix: string;
    procedure CreateAndCopyEmptyResIfNeeded;
    function RegisterPackage(IDE: TSpIDEType; Log: TStrings): Boolean;
  public
    property DPKFilename: string read FDPKFilename;
    property BPLFilename: string read FBPLFilename;
    property Exists: Boolean read FExists;
    property OnlyRuntime: Boolean read FOnlyRuntime;
    property OnlyDesigntime : Boolean read FOnlyDesigntime;
    property Description: string read FDescription;
    property LibSuffix: string read FLibSuffix;
    constructor Create(const Filename: string); virtual;
    function CompilePackage(DCC: string; IDE: TSpIDEType; SourcesL, IncludesL, Log: TStrings; TempDir: string = ''): Boolean;
  end;

  TSpDelphiDPKFilesList = class(TObjectList<TSpDelphiDPKFile>)
  public
    procedure Sort; reintroduce;
  end;

  TSpActionType = (satNone, satCopy, satCopyRun, satRun);

  TSpInstallType = (sitNotInstallable, sitInstallable, sitSearchPathOnly);

  TSpExecuteEntry = class
  private
    FAction: TSpActionType;
    FOrigin: string;
    FDestination: string;
  public
    property Action: TSpActionType read FAction write FAction;
    property Origin: string read FOrigin write FOrigin;
    property Destination: string read FDestination write FDestination;
  end;

  TSpExecuteList = class(TObjectList<TSpExecuteEntry>)
  public
    procedure LoadFromIni(Filename, Section: string);
    function ExecuteAll(BaseFolder: string; Log: TStrings): Boolean;
  end;

  TSpComponentPackage = class
    Name: string;
    ZipFile: string;
    Git : string;
    Destination: string;
    SearchPath: string;
    Includes: string;
    Installable: TSpInstallType;
    GroupIndex: Integer;
    ExecuteList: TSpExecuteList;
    PackageList: array [TSpIDEType] of string; // Array of commatext for each TSpIDEType i.e. ("Packages\SpTBXLib_XE8.dpk", "Packages\SpTBXLibDsgn_XE8.dpk")
    constructor Create; virtual;
    destructor Destroy; override;
  end;

  TSpComponentPackageList = class(TObjectList<TSpComponentPackage>)
  private
    FDefaultInstallIDE: TSpIDEType;
    FDefaultInstallFolder: string;
    FMinimumIDE: TSpIDEType;
  public
    procedure LoadFromIni(Filename: string);
    function ExtractAllZips(Source, Destination: string; Log: TStrings): Boolean;
    function ExecuteAll(BaseFolder: string; Log: TStrings): Boolean;
    function CompileAll(BaseFolder: string; IDE: TSpIDEType; Log: TStrings): Boolean;
    property DefaultInstallIDE: TSpIDEType read FDefaultInstallIDE write FDefaultInstallIDE;
    property DefaultInstallFolder: string read FDefaultInstallFolder write FDefaultInstallFolder;
    property MinimumIDE: TSpIDEType read FMinimumIDE;
  end;

  TSpMultiInstaller = class
  protected
    FComponentPackages: TSpComponentPackageList;
    FInstalling: Boolean;
  public
    constructor Create(IniFilename: string); virtual;
    destructor Destroy; override;
    function Install(ZipPath, BaseFolder: string; IDE: TSpIDEType; Log: TStrings): Boolean;
    property ComponentPackages: TSpComponentPackageList read FComponentPackages;
  end;

{ Misc }
procedure SpOpenLink(URL: string);
function SpStringSearch(S, SubStr: string; Delimiter: Char = ';'): Boolean;
procedure SpWriteLog(Log: TStrings; ResourceS, Arg1: string; Arg2: string = '');

{ Files }
function SpGetParameter(const ParamName: string; out ParamValue: string): Boolean;
function SpExecuteDosCommand(CommandLine, WorkDir: string; out OutputString: string): Cardinal;
function SpFileOperation(Origin, Destination: string; Operation: Cardinal): Boolean;
function SpSelectDirectory(Root: string; out Directory: string): Boolean;

{ Zip }
function SpExtractZip(ZipFilename, DestinationPath: string): Boolean;

{ Git }
function SpGitClone(AGit, DestinationPath: string; Log: TStrings): Boolean;

{ Ini and Registry }
function SpReadRegValue(Key, Name: string; out Value: string): Boolean;
function SpReadRegKey(Key: string; NamesAndValues: TStringList): Boolean;
function SpWriteRegValue(Key, Name, Value: string): Boolean;
procedure SpIniLoadStringList(L: TStringList; IniFilename, Section: string; NamePrefix: string = '');
procedure SpIniSaveStringList(L: TStringList; IniFilename, Section: string; NamePrefix: string = '');
function SpParseEntryValue(S: string; ValueList: TStringList; MinimumCount: Integer = 0): Boolean;

{ IDE }
function SpActionTypeToString(A: TSpActionType): string;
function SpStringToActionType(S: string): TSpActionType;

implementation

uses
  ActiveX, ShellApi, ShlObj, IniFiles, Registry,
  System.Zip, // Abbrevia is not needed anymore
  Vcl.FileCtrl, System.IOUtils, StrUtils, Generics.Defaults,
  Xml.XMLIntf, Xml.XMLDoc, themes;

const
  rvCount = 'Count';
  rvPackageIniSectionPrefix = 'Package -';
  rvName = 'Name';
  rvZip = 'Zip';
  rvGit = 'Git';
  rvFolder = 'Folder';
  rvSearchPath = 'SearchPath';
  rvGroupIndex = 'GroupIndex';
  rvIncludes = 'Includes';
  rvInstallable = 'Installable';
  rvExecuteIniPrefix = 'Execute';
  rvBaseFolder = '$BaseFolder';
  rvOptionsIniSection = 'Options';
  rvDefaultInstallIDE = 'DefaultInstallIDE';
  rvDefaultInstallFolder = 'DefaultInstallFolder';
  rvMinimumIDE = 'MinimumIDEVersion';
  ActionTypes: array [TSpActionType] of string = ('none', 'copy', 'copyandrun', 'run');
  IDEPersonalityRegNameTypes: array [TSpIDEPersonality] of string = ('Delphi.Win32', 'Delphi.NET', 'BCB');

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Helpers }

procedure SpOpenLink(URL: string);
begin
  ShellExecute(Application.Handle, 'open', PChar(URL), '', '', SW_SHOWNORMAL);
end;

function SpStringSearch(S, SubStr: string; Delimiter: Char): Boolean;
var
  L: TStringList;
begin
  L :=  TStringList.Create;
  try
    L.StrictDelimiter := True;
    L.Delimiter := Delimiter;
    L.DelimitedText := S;
    Result := L.IndexOf(SubStr) > -1;
  finally
    L.free;
  end;
end;

procedure SpWriteLog(Log: TStrings; ResourceS, Arg1: string; Arg2: string = '');
begin
  if Assigned(Log) then begin
    Log.Add(Format(ResourceS, [Arg1, Arg2]));
    Log.Add('');
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Files }

function SpGetParameter(const ParamName: string; out ParamValue: string): Boolean;
{ Determines whether a string was passed as a command line parameter to the
  application, and returns the parameter value if there is one.

  For example:
  File.exe /param1 c:\windows\internet files /param2 /param3
  File.exe /param1 "c:\windows\internet files" /param2 /param3
  File.exe -param1 c:\windows\internet files -param2 -param3
  File.exe -param1 "c:\windows\internet files" -param2 -param3

  SpGetParameter('param1', ParamValue) returns ParamValue = c:\windows\internet files }
begin
  ParamValue := '';
  Result := FindCmdLineSwitch(ParamName, ParamValue);
end;

function SpExecuteDosCommand(CommandLine, WorkDir: string; out OutputString: string): Cardinal;
// Executes a DOS file, waits until it terminates and logs the output.
// CommandLine param can be a file name with params, for example: CMD.exe /c dir D:\mp3
// Do not use pipes and redirections in CommandLine (|, >, <)
// Ported by Stephane Wierzbicki from JclSysUtils.InternalExecute
const
  BufferSize = 255;
  NativeLineFeed       = Char(#10);
  NativeCarriageReturn = Char(#13);
  NativeCrLf           = string(#13#10);
var
  Buffer: array [0..BufferSize] of AnsiChar;
  TempOutput: string;
  PipeBytesRead: Cardinal;

  function MuteCRTerminatedLines(const RawOutput: string): string;
  const
    Delta = 1024;
  var
    BufPos, OutPos, LfPos, EndPos: Integer;
    C: Char;
  begin
    SetLength(Result, Length(RawOutput));
    OutPos := 1;
    LfPos := OutPos;
    EndPos := OutPos;
    for BufPos := 1 to Length(RawOutput) do
    begin
      if OutPos >= Length(Result)-2 then
        SetLength(Result, Length(Result) + Delta);
      C := RawOutput[BufPos];
      case C of
        NativeCarriageReturn:
          OutPos := LfPos;
        NativeLineFeed:
          begin
            OutPos := EndPos;
            Result[OutPos] := NativeCarriageReturn;
            Inc(OutPos);
            Result[OutPos] := C;
            Inc(OutPos);
            EndPos := OutPos;
            LfPos := OutPos;
          end;
      else
        Result[OutPos] := C;
        Inc(OutPos);
        EndPos := OutPos;
      end;
    end;
    SetLength(Result, OutPos - 1);
  end;

  function CharIsReturn(const C: Char): Boolean;
  begin
    Result := (C = NativeLineFeed) or (C = NativeCarriageReturn);
  end;

  procedure ProcessLine(LineEnd: Integer);
  begin
    if (TempOutput[LineEnd] <> NativeCarriageReturn) then
    begin
      while (LineEnd > 0) and CharIsReturn(TempOutput[LineEnd]) do
        Dec(LineEnd);
    end;
  end;

  procedure ProcessBuffer;
  begin
    Buffer[PipeBytesRead] := #0;
    TempOutput := TempOutput + string(Buffer);
  end;

// outsourced from Win32ExecAndRedirectOutput
var
  StartupInfo: TStartupInfo;
  ProcessInfo: TProcessInformation;
  SecurityAttr: TSecurityAttributes;
  PipeRead, PipeWrite: THandle;
  PWorkDirChar : PChar;
begin
  Result := $FFFFFFFF;
  SecurityAttr.nLength := SizeOf(SecurityAttr);
  SecurityAttr.lpSecurityDescriptor := nil;
  SecurityAttr.bInheritHandle := True;
  if not CreatePipe(PipeRead, PipeWrite, @SecurityAttr, 0) then
  begin
    Result := GetLastError;
    Exit;
  end;
  FillChar(StartupInfo, SizeOf(TStartupInfo), #0);
  StartupInfo.cb := SizeOf(TStartupInfo);
  StartupInfo.dwFlags := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
  StartupInfo.wShowWindow := SW_HIDE;
  StartupInfo.hStdInput := GetStdHandle(STD_INPUT_HANDLE);
  StartupInfo.hStdOutput := PipeWrite;
  StartupInfo.hStdError := PipeWrite;

  if WorkDir = '' then PWorkDirChar := nil
  else PWorkDirChar := PChar(WorkDir);

  if CreateProcess(nil, PChar(CommandLine), nil, nil, True, NORMAL_PRIORITY_CLASS,
    nil, PWorkDirChar, StartupInfo, ProcessInfo) then
  begin
    CloseHandle(PipeWrite);

    while ReadFile(PipeRead, Buffer, BufferSize, PipeBytesRead, nil) and (PipeBytesRead > 0) do
      ProcessBuffer;
    if (WaitForSingleObject(ProcessInfo.hProcess, INFINITE) = WAIT_OBJECT_0) and
      not GetExitCodeProcess(ProcessInfo.hProcess, Result) then
        Result := $FFFFFFFF;
    CloseHandle(ProcessInfo.hThread);
    CloseHandle(ProcessInfo.hProcess);
  end
  else
    CloseHandle(PipeWrite);
  CloseHandle(PipeRead);

  if TempOutput <> '' then
    OutputString := OutputString + MuteCRTerminatedLines(TempOutput);
end;

function SpFileOperation(Origin, Destination: string; Operation: Cardinal): Boolean;
var
  F: TShFileOpStruct;
begin
   Result := False;
   // Operation can be: FO_COPY, FO_MOVE, FO_DELETE, FO_RENAME
   if not (Operation in [FO_MOVE..FO_RENAME]) then Exit;

   Origin := Origin + #0#0;
   Destination := Destination + #0#0;

   FillChar(F, SizeOf(F), #0);
   F.Wnd := Application.Handle;
   F.wFunc := Operation;
   F.pFrom := PChar(Origin);
   F.pTo := PChar(Destination);
   F.fFlags := FOF_SILENT or FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR;
   Result := SHFileOperation(F) = 0;
end;

function SpSelectDirectory(Root: string; out Directory: string): Boolean;
// SelectDirectory with new UI
var
  DArray: TArray<string>;
begin
  Result := Vcl.FileCtrl.SelectDirectory(Root, DArray, []);
  if Result then
    Directory := DArray[0];
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Zip }

function SpExtractZip(ZipFilename, DestinationPath: string): Boolean;
begin
  if not TDirectory.Exists(DestinationPath) then
    CreateDir(DestinationPath);
  TZipFile.ExtractZipFile(ZipFilename, DestinationPath);
  Result := True;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Git }

function SpGitClone(AGit, DestinationPath: string; Log: TStrings): Boolean;
var
  CommandLine, DosOutput: string;
begin
  CommandLine := Format(SGitCloneCommand, [AGit, DestinationPath]);
  Result := SpExecuteDosCommand(CommandLine, '', DosOutput) = 0;
  if Assigned(Log) then
    Log.Text := Log.Text + DosOutput + #13#10;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Ini and Registry }

function SpDeleteRegValue(Key, Name: string): Boolean;
var
  R: TRegistry;
begin
  Result := False;
  R := TRegistry.Create;
  try
    R.RootKey := HKEY_CURRENT_USER;
    if R.OpenKey(Key, False) then
      if R.ValueExists(Name) then
        Result := R.DeleteValue(Name);
  finally
    R.Free;
  end;
end;

function SpReadRegValue(Key, Name: string; out Value: string): Boolean;
var
  R: TRegistry;
begin
  Result := False;
  Value := '';
  R := TRegistry.Create;
  try
    R.RootKey := HKEY_CURRENT_USER;
    if R.OpenKey(Key, False) then
      if R.ValueExists(Name) then begin
        Value := R.ReadString(Name);
        Result := True;
      end;
  finally
    R.Free;
  end;
end;

function SpReadRegKey(Key: string; NamesAndValues: TStringList): Boolean;
var
  R: TRegistry;
  Names: TStringList;
  I: Integer;
begin
  Result := False;
  if not Assigned(NamesAndValues) then Exit;
  NamesAndValues.Clear;

  R := TRegistry.Create;
  Names := TStringList.Create;
  try
    R.RootKey := HKEY_CURRENT_USER;
    if R.OpenKey(Key, False) then begin
      R.GetValueNames(Names);
      for I := 0 to Names.Count - 1 do
        if R.ValueExists(Names[I]) then
          NamesAndValues.Values[Names[I]] := R.ReadString(Names[I]);
      Result := True;
    end;
  finally
    R.Free;
    Names.Free;
  end;
end;

function SpWriteRegValue(Key, Name, Value: string): Boolean;
var
  R: TRegistry;
begin
  Result := False;
  R := TRegistry.Create;
  try
    R.RootKey := HKEY_CURRENT_USER;
    if R.OpenKey(Key, True) then begin
      R.WriteString(Name, Value);
      Result := True;
    end;
  finally
    R.Free;
  end;
end;

procedure SpIniLoadStringList(L: TStringList; IniFilename, Section: string; NamePrefix: string = '');
var
  F: TMemIniFile;
  I, C: integer;
begin
  if not Assigned(L) then Exit;
  F := TMemIniFile.Create(IniFilename);
  try
    L.Clear;
    C := F.ReadInteger(Section, NamePrefix + rvCount, -1);
    for I := 0 to C - 1 do
      L.Add(F.ReadString(Section, NamePrefix + inttostr(I), ''));
  finally
    F.Free;
  end;
end;

procedure SpIniSaveStringList(L: TStringList; IniFilename, Section: string; NamePrefix: string = '');
var
  F: TMemIniFile;
  I: integer;
begin
  if not Assigned(L) then Exit;
  F := TMemIniFile.Create(IniFilename);
  try
    F.EraseSection(Section);
    if L.Count > 0 then begin
      F.WriteInteger(Section, NamePrefix + rvCount, L.Count);
      for I := 0 to L.Count - 1 do
        F.WriteString(Section, NamePrefix + IntToStr(I), L[I]);
      F.UpdateFile;
    end;
  finally
    F.Free;
  end;
end;

function SpParseEntryValue(S: string; ValueList: TStringList; MinimumCount: Integer = 0): Boolean;
begin
  ValueList.Clear;
  ValueList.CommaText := S;
  if MinimumCount < 1 then
    Result := ValueList.Count >= MinimumCount
  else
    Result := True;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ IDE }

function SpActionTypeToString(A: TSpActionType): string;
begin
  Result := ActionTypes[A];
end;

function SpStringToActionType(S: string): TSpActionType;
var
  A: TSpActionType;
begin
  Result := satNone;
  S := LowerCase(S);
  for A := Low(ActionTypes) to High(ActionTypes) do
    if AnsiSameText(S, ActionTypes[A]) then begin
      Result := A;
      Exit;
    end;
end;

procedure SpIDESearchPathRegKey(IDE: TSpIDEType; out Key, Name: string; CPPBuilderPath: Boolean);
begin
  Key := '';
  Name := '';
  if IDE = ideNone then Exit;

  // [IDE-Change]
  Key := IDETypes[IDE].IDERegistryPath;
  if CPPBuilderPath and (IDE >= ideDelphi2006) then begin
    if IDE = ideDelphi2006 then begin
      // '\CppPaths\SearchPath' with no space in the middle for C++Builder 2006
      Key := Key + '\CppPaths';
      Name := 'SearchPath';
    end
    else begin
      Name := 'LibraryPath';
      if IDE >= ideDelphiXE2 then
        Key := Key + '\C++\Paths\Win32' // '\C++\Paths\Win32\LibraryPath' with no space in the middle for C++Builder XE2 and above
      else
        Key := Key + '\C++\Paths';   // '\C++\Paths\LibraryPath' with no space in the middle for C++Builder 2009 and above
    end;
  end
  else begin
    Name := 'Search Path';
    if IDE >= ideDelphiXE2 then
      Key := Key + '\Library\Win32'
    else
      Key := Key + '\Library';
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpDelphiIDE }

class function TSpDelphiIDE.Installed(IDE: TSpIDEType): Boolean;
begin
  if IDE = ideNone then
    Result := False
  else
    Result := TFile.Exists(GetDCC32Filename(IDE));
end;

class function TSpDelphiIDE.PersonalityInstalled(IDE: TSpIDEType;
  IDEPersonality: TSpIDEPersonality): Boolean;
var
  S, PersReg: string;
begin
  Result := False;
  if IDE = ideNone then Exit;

  if IDE >= ideDelphi2006 then begin
    IDEPersonalityTypeToString(IDEPersonality, PersReg);
    SpReadRegValue(IDETypes[IDE].IDERegistryPath + '\Personalities',  PersReg, S);
    if S <> ''then
      Result := True;
  end;
end;

class function TSpDelphiIDE.StringToIDEType(S: string): TSpIDEType;
var
  A: TSpIDEType;
begin
  Result := ideNone;
  for A := Low(IDETypes) to High(IDETypes) do
    if AnsiSameText(S, IDETypes[A].IDEVersion) then begin
      Result := A;
      Exit;
    end;
end;

class procedure TSpDelphiIDE.IDEPersonalityTypeToString(A: TSpIDEPersonality;
  out IDERegName: string);
begin
  IDERegName := IDEPersonalityRegNameTypes[A];
end;

class function TSpDelphiIDE.GetIDEDir(IDE: TSpIDEType): string;
begin
  SpReadRegValue(IDETypes[IDE].IDERegistryPath, 'RootDir', Result);
end;

class function TSpDelphiIDE.GetDCC32Filename(IDE: TSpIDEType): string;
begin
  SpReadRegValue(IDETypes[IDE].IDERegistryPath, 'App', Result);
  if Result <> '' then
    Result := TPath.Combine(TPath.GetDirectoryName(Result), 'dcc32.exe');
end;

class function TSpDelphiIDE.GetBPLOutputDir(IDE: TSpIDEType): string;
begin
  // BPL Output Dir
  if IDE >= ideDelphiXE2 then
    SpReadRegValue(IDETypes[IDE].IDERegistryPath + '\Library\Win32', 'Package DPL Output', Result)
  else
    SpReadRegValue(IDETypes[IDE].IDERegistryPath + '\Library', 'Package DPL Output', Result);
  Result := TSpDelphiIDE.ExpandMacros(Result, IDE);
end;

class function TSpDelphiIDE.GetDCPOutputDir(IDE: TSpIDEType): string;
begin
  // DCP Output Dir
  if IDE >= ideDelphiXE2 then
    SpReadRegValue(IDETypes[IDE].IDERegistryPath + '\Library\Win32', 'Package DCP Output', Result)
  else
    SpReadRegValue(IDETypes[IDE].IDERegistryPath + '\Library', 'Package DCP Output', Result);
  Result := TSpDelphiIDE.ExpandMacros(Result, IDE);
end;

class function TSpDelphiIDE.ReadEnvironmentProj(IDE: TSpIDEType; NamesAndValues: TStringList): Boolean;
// Reads environment.proj file.
// In newer versions of RAD Studio (2007 and up) the macros are stored in:
// C:\Users\x\AppData\Roaming\Embarcadero\BDS\19.0\environment.proj
// This file is used by MSBuild
// https://docs.microsoft.com/en-us/visualstudio/msbuild/how-to-use-environment-variables-in-a-build?view=vs-2017
var
  LStr: array[0 .. MAX_PATH] of Char;
  Filename: string;
  Doc: IXMLDocument;
  Node, Root: IXMLNode;
begin
  Result := False;
  if not Assigned(NamesAndValues) or (IDE < ideDelphi2007) then Exit;
  NamesAndValues.Clear;

  SetLastError(ERROR_SUCCESS);
  if SHGetFolderPath(0, CSIDL_APPDATA, 0, 0, @LStr) = S_OK then begin
    Filename := TPath.Combine(LStr, 'Embarcadero\BDS\' + IDETypes[IDE].IDERADStudioVersion + '\environment.proj');
    if TFile.Exists(Filename) then begin
      Doc := TXMLDocument.Create(Filename);
      Root := Doc.ChildNodes.FindNode('Project');
      if Root <> nil then begin
        Root := Root.ChildNodes.FindNode('PropertyGroup');
        if Root <> nil then begin
          // Add $(Delphi), $(BDS), $(BDSPROJECTSDIR), $(BDSCOMMONDIR),
          // $(BDSUSERDIR), $(BDSLIB) macros
          Node := Root.ChildNodes.FindNode('Delphi');
          if Node <> nil then NamesAndValues.AddPair('Delphi', Node.Text);
          Node := Root.ChildNodes.FindNode('BDS');
          if Node <> nil then NamesAndValues.AddPair('BDS', Node.Text);
          Node := Root.ChildNodes.FindNode('BDSPROJECTSDIR');
          if Node <> nil then NamesAndValues.AddPair('BDSPROJECTSDIR', Node.Text);
          Node := Root.ChildNodes.FindNode('BDSCOMMONDIR');
          if Node <> nil then NamesAndValues.AddPair('BDSCOMMONDIR', Node.Text);
          Node := Root.ChildNodes.FindNode('BDSUSERDIR');
          if Node <> nil then NamesAndValues.AddPair('BDSUSERDIR', Node.Text);
          Node := Root.ChildNodes.FindNode('BDSLIB');
          if Node <> nil then NamesAndValues.AddPair('BDSLIB', Node.Text);
          Result := True;
        end;
      end;
    end;
  end;
end;

class function TSpDelphiIDE.ExpandMacros(S: string; IDE: TSpIDEType): string;
// Replace $(Delphi), $(BDS), $(BDSPROJECTSDIR), $(BDSCOMMONDIR),
// $(BDSUSERDIR), $(BDSLIB) macros and IDE Environment Variables Overrides
// with full directory paths.

  function ReplaceWithDefault(OverrideL, DefaultL: TStringList; Macro: string): Boolean;
  var
    I: Integer;
  begin
    Result := False;
    I := OverrideL.IndexOfName(Macro);
    if I = -1 then begin
      I := DefaultL.IndexOfName(Macro);
      if I >= 0 then begin
        // No override found, use default value
        OverrideL.Values[Macro] := DefaultL.ValueFromIndex[I];
        Result := True;
      end;
    end;

  end;

const
  // English, German, French strings
  DirArrayBDS: array [0..2] of string = ('Borland Studio Projects', 'Borland Studio-Projekte', 'Projets Borland Studio');
  DirArrayRAD: array [0..2] of string = ('Projects', 'Projekte', 'Projets');
var
  R, MyDocs: string;
  I: Integer;
  DefaultL, OverrideL: TStringList;
begin
  // In newer versions of RAD Studio (2007 and up) the macros are stored in:
  // C:\Users\x\AppData\Roaming\Embarcadero\BDS\19.0\environment.proj
  //
  // When the IDE is opened it reads environment.proj and sets the macros as
  // system Environment Variables.
  // When the IDE is closed it deletes the macros from the system Environment
  // Variables.
  // So when the IDE is not running we can't use SysUtils.GetEnvironmentVariable.

  Result := S;
  if IDE = ideNone then Exit;

  DefaultL := TStringList.Create;
  OverrideL := TStringList.Create;
  try
    // Try to read Environment.proj file
    ReadEnvironmentProj(IDE, DefaultL);
    // Get the Environment Variables Overrides
    SpReadRegKey(IDETypes[IDE].IDERegistryPath + '\Environment Variables', OverrideL);

    // Override the default macros
    // $(Delphi)
    if not ReplaceWithDefault(OverrideL, DefaultL, 'Delphi') then
      OverrideL.Values['Delphi'] := GetIDEDir(IDE);

    // $(BDS)
    if IDE >= ideDelphi2005 then
      if not ReplaceWithDefault(OverrideL, DefaultL, 'BDS') then
        OverrideL.Values['BDS'] := GetIDEDir(IDE);

    // $(BDSCOMMONDIR)
    // It points to a different directory according to how you install Delphi:
    // If you choose All Users during installation:
    //   C:\Users\Public\Documents\Embarcadero\Studio\19.0
    // If you choose Just Me during installation:
    //   C:\Users\x\Documents\Embarcadero\Studio\19.0
    if IDE >= ideDelphi2007 then
      ReplaceWithDefault(OverrideL, DefaultL, 'BDSCOMMONDIR');

    // $(BDSPROJECTSDIR)
    // Example: C:\Users\x\Documents\Embarcadero\Studio\Projects
    if IDE >= ideDelphi2005 then begin
      // This macro can be overrided by adding a string value called
      // 'DefaultProjectsDirectory' containing a different directory to:
      // HKCU\Software\Borland\BDS\4.0\Globals
      SpReadRegValue(IDETypes[IDE].IDERegistryPath + '\Globals', 'DefaultProjectsDirectory', R);
      if not TDirectory.Exists(R) then
        if not ReplaceWithDefault(OverrideL, DefaultL, 'BDSPROJECTSDIR') then begin
          // Try to guess it
          // Since BDSPROJECTSDIR is not defined in the registry we have to find it
          // manually, looking for all the localized dir names in MyDocuments folder
          MyDocs := TPath.GetDocumentsPath;
          if IDE in [ideDelphi2005, ideDelphi2006] then begin
            // For older BDS check if it's My Documents\Borland Studio Projects
            for I := 0 to High(DirArrayBDS) do begin
              R := TPath.Combine(MyDocs, DirArrayBDS[I]);
              if TDirectory.Exists(R) then Break;
            end;
          end
          else begin
            // For newer versions check if it's C:\Users\x\Documents\Embarcadero\Studio\Projects
            // or C:\Users\x\Documents\RAD Studio\Projects
            for I := 0 to High(DirArrayRAD) do begin
              if IDE >= ideDelphiXE6 then
                R := TPath.Combine(MyDocs, 'Embarcadero\Studio\' + DirArrayRAD[I])
              else
                R := TPath.Combine(MyDocs, 'RAD Studio\' + DirArrayRAD[I]);
              if TDirectory.Exists(R) then Break;
            end;
          end;
          if TDirectory.Exists(R) then
            OverrideL.Values['BDSPROJECTSDIR'] := R;
        end;
    end;

    // $(BDSUSERDIR)
    // Example: C:\Users\x\Documents\Embarcadero\Studio\19.0
    if IDE >= ideDelphi2007 then
      if not ReplaceWithDefault(OverrideL, DefaultL, 'BDSUSERDIR') then begin
        // Try to guess it
        // Get BDSPROJECTSDIR and add IDE version
        R := OverrideL.Values['BDSPROJECTSDIR'];
        if TDirectory.Exists(R) then begin
          R := TPath.Combine(ExtractFilePath(R), IDETypes[IDE].IDERADStudioVersion);
          if TDirectory.Exists(R) then
            OverrideL.Values['BDSUSERDIR'] := R;
        end;
      end;

    // $(BDSLIB)
    // Example: C:\Program Files\Embarcadero\Studio\19.0\lib
    if not ReplaceWithDefault(OverrideL, DefaultL, 'BDSLIB') then
      OverrideL.Values['BDSLIB'] := TPath.Combine(GetIDEDir(IDE), 'lib');

    // $(PLATFORM)
    // Not sure were to find this macro
    // Since we're using DCC32 to compile assume Win32
    OverrideL.Values['PLATFORM'] := 'Win32';

    // Replace all
    for I := 0 to OverrideL.Count - 1 do
      Result := StringReplace(Result, '$(' + OverrideL.Names[I] + ')', ExcludeTrailingPathDelimiter(OverrideL.ValueFromIndex[I]), [rfReplaceAll, rfIgnoreCase]);
  finally
    DefaultL.Free;
    OverrideL.Free;
  end;
end;

class function TSpDelphiIDE.GetSearchPath(IDE: TSpIDEType;
  CPPBuilderPath: Boolean): string;
var
  Key, Name: string;
begin
  Result := '';
  if IDE <> ideNone then begin
    SpIDESearchPathRegKey(IDE, Key, Name, CPPBuilderPath);
    SpReadRegValue(Key, Name, Result);
  end;
end;

class procedure TSpDelphiIDE.AddToSearchPath(SourcesL: TStrings; IDE: TSpIDEType);
var
  I : Integer;
  S, Key, Name: string;
begin
  for I := 0 to SourcesL.Count - 1 do begin
    SourcesL[I] := ExpandMacros(ExcludeTrailingPathDelimiter(SourcesL[I]), IDE);

    // Add the directory to the Delphi Win32 search path registry entry
    S := GetSearchPath(IDE, False);
    if (S <> '') and (SourcesL[I] <> '') then
      if not SpStringSearch(S, SourcesL[I]) then begin
        if S[Length(S)] <> ';' then
          S := S + ';';
        S := S + SourcesL[I];
        SpIDESearchPathRegKey(IDE, Key, Name, False);
        SpWriteRegValue(Key, Name, S)
      end;

    // Add the directory to the C++Builder search path registry entry
    if IDE >= ideDelphi2006 then begin
      S := GetSearchPath(IDE, True);
      if (S <> '') and (SourcesL[I] <> '') then
        if not SpStringSearch(S, SourcesL[I]) then begin
          if S[Length(S)] <> ';' then
            S := S + ';';
          S := S + SourcesL[I];
          SpIDESearchPathRegKey(IDE, Key, Name, True);
          SpWriteRegValue(Key, Name, S)
        end;
    end;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpDelphiDPKFile }

constructor TSpDelphiDPKFile.Create(const Filename: string);
var
  L: TStringList;
  P, P2: Integer;
  Suffix: string;
begin
  FDPKFilename := Filename;
  FBPLFilename := '';
  FExists := False;
  FOnlyRuntime := False;
  FOnlyDesigntime := False;
  FDescription := '';
  FLibSuffix := '';

  Suffix := '';
  if TFile.Exists(Filename) then begin
    FDPKFilename := Filename;
    FBPLFilename := TPath.ChangeExtension(TPath.GetFileName(Filename), 'bpl');

    L := TStringList.Create;
    try
      L.LoadFromFile(Filename);
      P := Pos('{$RUNONLY}', L.Text);
      if P > 0 then
        FOnlyRuntime := True;
      P := Pos('{$DESIGNONLY}', L.Text);
      if P > 0 then
        FOnlyDesigntime := True;
      P := Pos('{$DESCRIPTION ''', L.Text); // {$DESCRIPTION 'Package Description'}
      if P > 0 then begin
        P := P + Length('{$DESCRIPTION ''');
        P2 := PosEx('''}', L.Text, P);
        if P2 > 0 then
          FDescription := Copy(L.Text, P, P2 - P);
      end;
      P := Pos('{$LIBSUFFIX ''', L.Text); // {$LIBSUFFIX '100'}  //  file100.bpl
      if P > 0 then begin
        P := P + Length('{$LIBSUFFIX ''');
        P2 := PosEx('''}', L.Text, P);
        if P2 > 0 then begin
          Suffix := Copy(L.Text, P, P2 - P);
          // Rename BPL filename to include the suffix
          FBPLFilename := TPath.GetFileNameWithoutExtension(FBPLFilename) + Suffix + TPath.GetExtension(FBPLFilename);
        end;
      end;

      FExists := True;
    finally
      L.Free;
    end;
  end;
end;

procedure TSpDelphiDPKFile.CreateAndCopyEmptyResIfNeeded;
var
  ResFile: string;
  FStream: TResourceStream;
begin
  // Create and copy a res file if needed
  if Exists then begin
    ResFile := TPath.ChangeExtension(FDPKFilename, 'res');
    if not TFile.Exists(ResFile) then begin
      FStream := TResourceStream.Create(HInstance, 'EMPTYRES', RT_RCDATA);
      try
        FStream.SaveToFile(ResFile);
      finally
        FStream.Free;
      end;
    end;
  end;
end;

function TSpDelphiDPKFile.CompilePackage(DCC: string; IDE: TSpIDEType; SourcesL,
  IncludesL, Log: TStrings; TempDir: string): Boolean;
// DCC = full path of dcc32.exe, e.g. 'C:\Program Files\Borland\Delphi7\Bin\dcc32.exe
// IDE = IDE version to compile with
// SourcesL = list of source folders of the component package to add to the Library Search Path
// IncludesL = list of includes folder paths, e.g. 'C:\MyCompos\AnotherCompo\Source'
// Log = Log strings
// TempDir = Temp dir where the package dcu will be copied, e.g. 'C:\Windows\Temp\MyCompos'
var
  CommandLine, WorkDir, DOSOutput, DCCConfig: string;
  L: TStringList;
  I: Integer;
  S, R: string; // Auxiliary strings
begin
  Result := False;
  if IDE = ideNone then Exit;
  if not Exists then begin
    if Assigned(Log) then
      SpWriteLog(Log, SLogInvalidPath, FDPKFilename);
    Exit;
  end
  else begin
    // [IDE Bug]: dcc32.exe won't execute if -Q option is not used
    // But it works fine without -Q if ShellExecute is used:
    // ShellExecute(Application.Handle, 'open', DCC, ExtractFileName(FDPKFilename), ExtractFilePath(FDPKFilename), SW_SHOWNORMAL);
    // There must be something wrong with SpExecuteDosCommand
    // Example: dcc32.exe -Q sptbxlib.dpk
    CommandLine := DCC + ' -Q  ' + TPath.GetFileName(FDPKFilename);
    WorkDir := TPath.GetDirectoryName(FDPKFilename);

    // Create and save DCC32.CFG file on the Package directory
    // Example of cfg file:
    // -U"$(BDSLIB)\$(Platform)\release";"C:\TB2K\Source";"C:\SpTBXLib\Source"
    // -R"C:\SpTBXLib\Source"
    // -LE"C:\Users\Public\Documents\Embarcadero\Studio\19.0\Bpl"
    // -LN"C:\Users\Public\Documents\Embarcadero\Studio\19.0\Dcp"
    // -NB"C:\Users\Public\Documents\Embarcadero\Studio\19.0\Dcp"
    // -NSSystem.Win;Data.Win;Datasnap.Win;Web.Win;Soap.Win;Xml.Win;Bde;Vcl;Vcl.Imaging;Vcl.Touch;Vcl.Samples;Vcl.Shell;System;Xml;Data;Datasnap;Web;Soap;Winapi
    // -N"C:\Users\x\AppData\Local\Temp\SpMultiInstall"
    // -JL
    L := TStringList.Create;
    try
      // Add the SourcesL directories to the registry
      TSpDelphiIDE.AddToSearchPath(SourcesL, IDE);

      // Expand SearchPath, replace $(Delphi) and $(BDS) with real directories
      // and enclose the paths with " " to transform it to a valid
      // comma delimited string for the -U switch.
      L.Text := TSpDelphiIDE.ExpandMacros(TSpDelphiIDE.GetSearchPath(IDE, False), IDE);
      L.Text := StringReplace(L.Text, ';', #13#10, [rfReplaceAll, rfIgnoreCase]);
      for I := 0 to L.Count - 1 do
        L[I] := '"' + L[I] + '"';
      S := StringReplace(L.Text, #13#10, ';', [rfReplaceAll, rfIgnoreCase]);
      if S[Length(S)] = ';' then
        Delete(S, Length(S), 1);

      // Save the DCC32.CFG file on the Package directory
      DCCConfig := TPath.Combine(WorkDir, 'DCC32.CFG');
      R := IDETypes[IDE].IDERegistryPath;
      L.Clear;
      // SearchPath
      L.Add('-U' + S);
      // Resource directories, add the source folder as the default *.dcr search folder
      S := '';
      for I := 0 to SourcesL.Count - 1 do
        S := S + ';' + SourcesL[I];
      if S <> '' then begin
        Delete(S, 1, 1);
        S := '"' + S + '"';
        L.Add('-R' + S);
      end;
      // BPL Output
      S := TSpDelphiIDE.GetBPLOutputDir(IDE);
      L.Add('-LE"' + S + '"');
      // DCP Output
      S := TSpDelphiIDE.GetDCPOutputDir(IDE);
      L.Add('-LN"' + S + '"');
      // BPI Output for the compiled packages, required for C++Builder 2006 and above,
      // same as DCP Output
      if IDE >= ideDelphi2006 then
        L.Add('-NB"' + S + '"');
      // Unit namespaces for Delphi XE2:
      if IDE >= ideDelphiXE2 then
        L.Add('-NSSystem.Win;Data.Win;Datasnap.Win;Web.Win;Soap.Win;Xml.Win;Bde;Vcl;Vcl.Imaging;Vcl.Touch;Vcl.Samples;Vcl.Shell;System;Xml;Data;Datasnap;Web;Soap;Winapi');
      // Includes, dcc32.exe accepts Includes as a semicolon separated string
      // enclosed by double quotes, e.g. "C:\dir1;C:\dir2;C:\dir3"
      S := '';
      for I := 0 to IncludesL.Count - 1 do
        S := S + ';' + IncludesL[I];
      if S <> '' then begin
        Delete(S, 1, 1);
        S := '"' + S + '"';
        L.Add('-I' + S);
      end;
      // DCU Output for the compiled packages
      if TempDir <> '' then
        L.Add('-N"' + TempDir + '"');
      // Add -JL compiler switch to make Hpp files required for C++Builder 2006 and above
      // This switch is undocumented:
      // http://groups.google.com/group/borland.public.cppbuilder.ide/browse_thread/thread/456bece4c5665459/0c4c61ecec179ca8
      if IDE >= ideDelphi2006 then
        if TSpDelphiIDE.PersonalityInstalled(IDE, persCPPBuilder) then
          L.Add('-JL');

      L.SaveToFile(DCCConfig);
    finally
      L.Free;
    end;

    // Create and copy an empty res file if needed.
    // Some component libraries like VirtualTreeView don't include .res files.
    CreateAndCopyEmptyResIfNeeded;

    // Compile
    SpWriteLog(Log, SLogCompiling, FDPKFilename);
    try
      Result := SpExecuteDosCommand(CommandLine, WorkDir, DOSOutput) = 0;
      if Assigned(Log) then
        Log.Text := Log.Text + DosOutput + #13#10;
      if Result then
        Result := RegisterPackage(IDE, Log);
    finally
      DeleteFile(DCCConfig);
    end;
  end;

  if not Result and Assigned(Log) then
    SpWriteLog(Log, SLogErrorCompiling, FDPKFilename, '');
end;

function TSpDelphiDPKFile.RegisterPackage(IDE: TSpIDEType;
  Log: TStrings): Boolean;
var
  BPL, RegKey: string;
begin
  Result := False;
  if IDE = ideNone then Exit;

  // BPL filename
  BPL := TPath.Combine(TSpDelphiIDE.GetBPLOutputDir(IDE), FBPLFilename);

  RegKey := IDETypes[IDE].IDERegistryPath + '\Known Packages';

  if FOnlyRuntime then begin
    SpDeleteRegValue(RegKey, BPL);
    Result := True
  end
  else
    if FOnlyDesigntime and TFile.Exists(BPL) then begin
      if SpWriteRegValue(RegKey, BPL, FDescription) then begin
        SpWriteLog(Log, SLogInstalling, FDPKFilename);
        Result := True;
      end;
    end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpDelphiDPKFilesList }

procedure TSpDelphiDPKFilesList.Sort;
begin
  inherited Sort(TComparer<TSpDelphiDPKFile>.Construct(
      function (const Left, Right: TSpDelphiDPKFile): Integer
      begin
        // Runtime packages should be sorted first
        if not Left.FOnlyDesigntime and Right.FOnlyDesigntime then
          Result := -1
        else
          if Left.FOnlyDesigntime and not Right.FOnlyDesigntime then
            Result := 1
          else
            Result := 0;
      end
  ));
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpComponentPackage }

constructor TSpComponentPackage.Create;
begin
  inherited;
  ExecuteList := TSpExecuteList.Create;
end;

destructor TSpComponentPackage.Destroy;
begin
  ExecuteList.Free;
  inherited;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpComponentPackageList }

procedure TSpComponentPackageList.LoadFromIni(Filename: string);
var
  F: TMemIniFile;
  LSections: TStringList;
  Entry: TSpComponentPackage;
  I, Aux: integer;
  S: string;
  A: TSpIDEType;
begin
  if not TFile.Exists(Filename) then
    Exit;

  LSections := TStringList.Create;
  F := TMemIniFile.Create(Filename);
  try
    // Read Options
    S := F.ReadString(rvOptionsIniSection, rvDefaultInstallIDE, '');
    FDefaultInstallIDE := TSpDelphiIDE.StringToIDEType(S);
    FDefaultInstallFolder := F.ReadString(rvOptionsIniSection, rvDefaultInstallFolder, '');
    S := F.ReadString(rvOptionsIniSection, rvMinimumIDE, '');
    FMinimumIDE := TSpDelphiIDE.StringToIDEType(S);
    if FMinimumIDE = ideNone then
      FMinimumIDE := ideDelphi7;

    // Read Component Packages
    F.ReadSections(LSections);
    for I := 0 to LSections.Count - 1 do begin
      S := LSections[I];
      if Length(S) > Length(rvPackageIniSectionPrefix) then
        if AnsiSameText(Copy(S, 1, Length(rvPackageIniSectionPrefix)), rvPackageIniSectionPrefix) then begin
          Entry := TSpComponentPackage.Create;
          Entry.Name := F.ReadString(S, rvName, '');
          Entry.ZipFile := F.ReadString(S, rvZip, '');
          Entry.Git := F.ReadString(S, rvGit, '');
          Entry.Destination := F.ReadString(S, rvFolder, '');
          Entry.SearchPath := F.ReadString(S, rvSearchPath, '');
          Entry.GroupIndex := F.ReadInteger(S, rvGroupIndex, 0);
          Aux := F.ReadInteger(S, rvInstallable, 1);
          if Aux < 0 then Aux := 0;
          if Aux > Ord(High(TSpInstallType)) then Aux := 1;
          Entry.Installable := TSpInstallType(Aux);
          Entry.Includes := F.ReadString(S, rvIncludes, '');

          // [IDE-Change]
          if Entry.Installable = sitInstallable then begin
            for A := FMinimumIDE to High(TSpIDEType) do
              Entry.PackageList[A] := F.ReadString(S, IDETypes[A].IDEVersion, '');
          end;

          Entry.ExecuteList.LoadFromIni(Filename, S);

          Add(Entry);
        end;
    end;
  finally
    LSections.Free;
    F.Free;
  end;
end;

function TSpComponentPackageList.ExtractAllZips(Source, Destination: string;
  Log: TStrings): Boolean;
var
  I: integer;
  Item: TSpComponentPackage;
begin
  Result := False;
  SpWriteLog(Log, SLogStartUnzip, '');

  // Check if the files exist
  if not TDirectory.Exists(Destination) then begin
    SpWriteLog(Log, SLogInvalidPath, Destination);
    Exit;
  end;
  for I := 0 to Count - 1 do begin
    Item := Items[I];
    // Expand ZipFile
    if Item.ZipFile <> '' then
      Item.ZipFile := TPath.Combine(Source, Item.ZipFile);
    // Expand Destination
    if Item.Destination <> '' then
      Item.Destination := TPath.Combine(Destination, Item.Destination);
    if TFile.Exists(Item.ZipFile) then begin
      if not AnsiSameText(TPath.GetExtension(Item.ZipFile), '.ZIP') then begin
        SpWriteLog(Log, SLogNotAZip, Item.ZipFile);
        Exit;
      end;
    end
    else
      if Item.Git <> '' then begin
        if not AnsiSameText(TPath.GetExtension(Item.Git), '.GIT') then begin
          SpWriteLog(Log, SLogNotAGit, Item.Git);
          Exit;
        end;
      end;
  end;

  // Unzip - Git Clone
  for I := 0 to Count - 1 do begin
    Item := Items[I];

    if Item.ZipFile <> '' then begin
      SpWriteLog(Log, SLogExtracting, Item.ZipFile, Item.Destination);
      if not SpExtractZip(Item.ZipFile, Item.Destination) then begin
        SpWriteLog(Log, SLogCorruptedZip, Item.ZipFile);
        Exit;
      end;
    end
    else
      if Item.Git <> '' then begin
        SpWriteLog(Log, SLogGitCloning, Item.Git, Item.Destination);
        if not SpGitClone(Item.Git, Item.Destination, Log) then begin
          SpWriteLog(Log, SLogGitCloneFailed, Item.Git);
          Exit;
        end;
      end
      else
        SpWriteLog(Log, SLogNotInstallable, Item.Name); // Not a Zip nor a Git, keep going
  end;
  Result := True;
end;

function TSpComponentPackageList.ExecuteAll(BaseFolder: string; Log: TStrings): Boolean;
var
  I: Integer;
begin
  Result := True;
  if Count > 0 then begin
    SpWriteLog(Log, SLogStartExecute, '');
    for I := 0 to Count - 1 do begin
      Result := Items[I].ExecuteList.ExecuteAll(BaseFolder, Log);
      if not Result then Exit;
      Application.ProcessMessages;
    end;
  end;
end;

function TSpComponentPackageList.CompileAll(BaseFolder: string; IDE: TSpIDEType; Log: TStrings): Boolean;
var
  DCC, TempDir: string;
  I, J: integer;
  Item: TSpComponentPackage;
  SourcesL, CompileL, IncludesL: TStringList;
  DPKList: TSpDelphiDPKFilesList;
begin
  Result := False;
  if IDE = ideNone then begin
    Result := True;
    Exit;
  end
  else
    if not TSpDelphiIDE.Installed(IDE) then begin
      SpWriteLog(Log, SLogInvalidIDE, IDETypes[IDE].IDEName);
      Exit;
    end;

  // Create TempDir
  TempDir := TPath.Combine(TPath.GetTempPath, 'SpMultiInstall');
  CreateDir(TempDir);

  DCC := TSpDelphiIDE.GetDCC32Filename(IDE);
  SourcesL := TStringList.Create;
  try
    for I := 0 to Count - 1 do begin
      Item := Items[I];
      SpWriteLog(Log, SLogStartCompile, Item.Name);

      // Expand Search Path
      if Item.SearchPath <> '' then begin
        SourcesL.CommaText := Item.SearchPath;
        // Add the destination search path
        for J := 0 to SourcesL.Count - 1 do
          SourcesL[J] := TPath.Combine(Item.Destination, SourcesL[J]);
      end
      else
        SourcesL.Add(Item.Destination);
      Item.SearchPath := SourcesL.CommaText;

      case Item.Installable of
        sitNotInstallable: ; // do nothing
        sitSearchPathOnly:
          // If the package is not installable add the SearchPath to the registry
          // This is useful when installing utility libraries that doesn't have
          // components to install, for example GraphicEx, GDI+, DirectX, etc
          TSpDelphiIDE.AddToSearchPath(SourcesL, IDE);
        sitInstallable:
          begin
            IncludesL := TStringList.Create;
            DPKList := TSpDelphiDPKFilesList.Create;
            try
              // Expand Packages
              CompileL := TStringList.Create;
              try
                CompileL.CommaText := Item.PackageList[IDE];
                for J := 0 to CompileL.Count - 1 do
                  DPKList.Add(TSpDelphiDPKFile.Create(TPath.Combine(Item.Destination, CompileL[J])));
              finally
                CompileL.Free;
              end;
              // Runtime packages must be compiled first
              DPKList.Sort;
              // Expand Includes
              IncludesL.CommaText := StringReplace(Item.Includes, rvBaseFolder, ExcludeTrailingPathDelimiter(BaseFolder), [rfReplaceAll, rfIgnoreCase]);
              // Compile and Install
              for J := 0 to DPKList.Count - 1 do
                if not DPKList[J].CompilePackage(DCC, IDE, SourcesL, IncludesL, Log, TempDir) then
                  Exit;
            finally
              IncludesL.Free;
              DPKList.Free;
            end;
          end;
      end;

      Application.ProcessMessages;
    end;
    Result := True;
  finally
    SourcesL.Free;
    SpFileOperation(TempDir, '', FO_DELETE);
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpExecuteList }

procedure TSpExecuteList.LoadFromIni(Filename, Section: string);
var
  L, V: TStringList;
  ExecuteEntry: TSpExecuteEntry;
  Action: TSpActionType;
  I: integer;
begin
  L := TStringList.Create;
  V := TStringList.Create;
  try
    Clear;
    SpIniLoadStringList(L, Filename, Section, rvExecuteIniPrefix);
    for I := 0 to L.Count - 1 do
      if SpParseEntryValue(L[I], V, 3) then begin
        Action := SpStringToActionType(V[0]);
        if Action <> satNone then begin
          ExecuteEntry := TSpExecuteEntry.Create;
          ExecuteEntry.Action := Action;
          ExecuteEntry.Origin := V[1];
          ExecuteEntry.Destination := V[2];
          Add(ExecuteEntry);
        end;
      end;
  finally
    L.Free;
    V.Free;
  end;
end;

function TSpExecuteList.ExecuteAll(BaseFolder: string; Log: TStrings): Boolean;
var
  I: Integer;
  Item: TSpExecuteEntry;

  function ExecuteRun: Boolean;
  var
    S, DosOutput: string;
  begin
    // Run it if it's a valid file
    Result := False;
    S := TPath.GetFileName(Item.Origin);
    if S <> '' then begin
      S := TPath.Combine(Item.Destination, S);
      SpWriteLog(Log, SLogExecuting, S);
      if SpExecuteDosCommand(S, Item.Destination, DosOutput) = 0 then begin
        Log.Text := Log.Text + DosOutput + #13#10;
        Result := True;
      end
      else
        SpWriteLog(Log, SLogErrorExecuting, S, '');
    end;
  end;

begin
  Result := False;

  // Check if the files exist
  for I := 0 to Count - 1 do begin
    Item := Items[I];
    Item.Origin := StringReplace(Item.Origin, rvBaseFolder, BaseFolder, [rfReplaceAll, rfIgnoreCase]);
    Item.Destination := StringReplace(Item.Destination, rvBaseFolder, BaseFolder, [rfReplaceAll, rfIgnoreCase]);
    if not TFile.Exists(Item.Origin) then begin
      SpWriteLog(Log, SLogInvalidPath, Item.Origin);
      Exit;
    end;
  end;

  // Execute
  for I := 0 to Count - 1 do begin
    Item := Items[I];
    case Item.Action of
      satRun:
        if not ExecuteRun then
          Exit;
      satCopy, satCopyRun:
        if SpFileOperation(Item.Origin, Item.Destination, FO_COPY) then begin
          SpWriteLog(Log, SLogCopying, Item.Origin, Item.Destination);
          if Item.Action = satCopyRun then
            if not ExecuteRun then
              Exit;
        end
        else begin
          SpWriteLog(Log, SLogErrorCopying, Item.Origin, Item.Destination);
          Exit;
        end;
    end;
  end;

  Result := True;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpMultiInstaller }

constructor TSpMultiInstaller.Create(IniFilename: string);
begin
  FComponentPackages := TSpComponentPackageList.Create;
  FComponentPackages.LoadFromIni(IniFilename);
end;

destructor TSpMultiInstaller.Destroy;
begin
  FComponentPackages.Free;
  inherited;
end;

function TSpMultiInstaller.Install(ZipPath, BaseFolder: string; IDE: TSpIDEType; Log: TStrings): Boolean;
var
  N, Secs: Single;
begin
  Result := False;
  FInstalling := True;
  try
    Log.Clear;
    N := GetTickCount;
    if ComponentPackages.ExtractAllZips(ZipPath, BaseFolder, Log) then
      if ComponentPackages.ExecuteAll(BaseFolder, Log) then
        if ComponentPackages.CompileAll(BaseFolder, IDE, Log) then begin
          Secs := (GetTickCount - N) / 1000;
          SpWriteLog(Log, SLogEnd, '');
          Log.Add(Format(SLogFinished, [Secs]));

          // [IDE-Change]
          if IDE >= ideDelphi2007 then begin
            // From the Delphi 2007 readme:
            // http://edn.embarcadero.com/article/36648
            // If you your component installer updates paths in Delphi's registry to include paths
            // to your components, you should add the following registry key:
            // HKCU\Software\Borland\BDS\5.0\Globals\ForceEnvOptionsUpdate with a string value "1"
            // After adding this registry key the next time the IDE is run, it will update
            // the EnvOptions.proj file on disk to include the paths that your installer added.
            // The EnvOptions.proj file is the mechanism by which the new MSBuild build engine
            // in the IDE is able to include paths that are listed on the Library - Win32 page in
            // the IDE's Tools > Options dialog.
            // The EnvOptions.proj file is in:
            // Vista/7/8: C:\Users\...\AppData\Roaming\Borland\BDS\5.0
            //            C:\Users\...\AppData\Roaming\Embarcadero\BDS\15.0
            SpWriteRegValue(IDETypes[IDE].IDERegistryPath + '\Globals', 'ForceEnvOptionsUpdate', '1');
          end;

          Result := True;
        end;
  finally
    FInstalling := False;
  end;
end;

end.
