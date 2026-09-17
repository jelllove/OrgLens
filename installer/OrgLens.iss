#ifndef AppVersion
  #error AppVersion is required.
#endif
#ifndef PackageDir
  #error PackageDir is required and must be an absolute package directory.
#endif
#ifndef OutputDir
  #error OutputDir is required and must be an absolute output directory.
#endif
#ifndef AssemblyIdentity
  #error AssemblyIdentity must be the packaged OrgLens.Outlook assembly FullName.
#endif
#ifndef AssemblyVersion
  #error AssemblyVersion is required.
#endif
#if defined(TestOutlookPath) && !defined(TestRunId)
  #error TestOutlookPath requires TestRunId. Release installers have no prerequisite override.
#endif
#if defined(TestProcessName) && !defined(TestRunId)
  #error TestProcessName requires TestRunId. Release installers always check OUTLOOK.EXE.
#endif
#ifndef TestProcessName
  #define ProcessName "OUTLOOK.EXE"
#else
  #define ProcessName TestProcessName
  #if Len(ProcessName) == 0 || Pos("\", ProcessName) > 0 || Pos("/", ProcessName) > 0
    #error TestProcessName must be an executable filename, not a path.
  #endif
#endif

#define ProductName "OrgLens"
#define ProductId "OrgLens"
#define RegistrationPrefix ""
#define InstallDirectory "{localappdata}\OrgLens\Addin"
#define RollbackDirectory "{localappdata}\OrgLens\InstallerRollback"
#define OutputName "OrgLens-" + AppVersion + "-Setup"
#ifdef TestRunId
  #if Len(TestRunId) != 36
    #error TestRunId must be a GUID without braces.
  #endif
  #if Copy(TestRunId, 9, 1) != "-" || Copy(TestRunId, 14, 1) != "-" || Copy(TestRunId, 19, 1) != "-" || Copy(TestRunId, 24, 1) != "-"
    #error TestRunId must be a GUID without braces.
  #endif
  #define TestCharacter 0
  #sub ValidateTestCharacter
    #if TestCharacter != 9 && TestCharacter != 14 && TestCharacter != 19 && TestCharacter != 24 && Pos(Copy(TestRunId, TestCharacter, 1), "0123456789abcdefABCDEF") == 0
      #error TestRunId contains an invalid character.
    #endif
  #endsub
  #for {TestCharacter = 1; TestCharacter <= 36; TestCharacter++} ValidateTestCharacter
  #define ProductName "OrgLens Setup Test " + TestRunId
  #define ProductId "OrgLens.SetupTest." + TestRunId
  #define RegistrationPrefix "Software\OrgLens\SetupTests\" + TestRunId + "\"
  #define InstallDirectory "{localappdata}\OrgLens\SetupTests\" + TestRunId + "\Addin"
  #define RollbackDirectory "{localappdata}\OrgLens\SetupTests\" + TestRunId + "\Rollback"
  #define OutputName "OrgLens-" + AppVersion + "-Setup-Test-" + TestRunId
#endif

[Setup]
AppId={#ProductId}
AppName={#ProductName}
AppVersion={#AppVersion}
AppVerName={#ProductName} {#AppVersion}
AppPublisher=OrgLens
DefaultDirName={#InstallDirectory}
DefaultGroupName={#ProductName}
DisableProgramGroupPage=yes
DisableDirPage=auto
PrivilegesRequired=lowest
MinVersion=6.1sp1
LicenseFile=..\LICENSE
OutputDir={#OutputDir}
OutputBaseFilename={#OutputName}
SetupIconFile=..\assets\orglens.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\OrgLens.Preview.exe
UninstallDisplayName={#ProductName}
UninstallLogMode=append
UsePreviousAppDir=yes
CloseApplications=no
RestartApplications=no
AlwaysRestart=no
RestartIfNeededByRun=no
SetupMutex={#ProductId}.Setup
VersionInfoVersion={#AssemblyVersion}
VersionInfoProductName={#ProductName}
VersionInfoProductVersion={#AppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
WelcomeLabel2=This will install [name/ver] for your Windows account.%n%nClose classic Outlook before continuing. Setup never closes or starts Outlook automatically.%n%nYour saved accounts, profile cache, and Outlook views will be preserved.
FinishedLabelNoIcons=OrgLens is installed.%n%nOpen classic Outlook when you are ready, then choose the OrgLens tab and Formatting rules. No formatting is applied automatically.
ConfirmUninstall=Are you sure you want to remove %1?%n%nClose classic Outlook first. Saved account profiles, cache, and Outlook view formatting will be kept.
UninstalledAll=%1 was successfully removed.%n%nSaved account profiles and cache were kept. Outlook views were not changed.

[Files]
Source: "{#PackageDir}\OrgLens.Outlook.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\OrgLens.Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\OrgLens.Desktop.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\Newtonsoft.Json.LICENSE.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\Lucide.LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\OrgLens.Preview.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\OrgLens.Preview.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\CHANGELOG.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PackageDir}\docs\images\orglens-settings.png"; DestDir: "{app}\docs\images"; Flags: ignoreversion
Source: "{#PackageDir}\LICENSE"; DestDir: "{app}"; Flags: ignoreversion; AfterInstall: InstallRegistration

[Icons]
Name: "{group}\OrgLens Preview"; Filename: "{app}\OrgLens.Preview.exe"; WorkingDir: "{app}"
Name: "{group}\Uninstall {#ProductName}"; Filename: "{uninstallexe}"

[Code]
const
  RegistrationPrefix = '{#RegistrationPrefix}';
  ComClass = 'OrgLens.Outlook.Connect';
  ProgId = 'OrgLens.Connect';
  ClassId = '{D7E2D48A-9466-4D83-856F-AC197BC23A98}';
  ManagedCategory = '{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}';
  CurrentAssembly = '{#AssemblyIdentity}';
  CurrentAssemblyVersion = '{#AssemblyVersion}';
  CheckedProcessName = '{#StringChange(ProcessName, "'", "''")}';
  KeyRead = $20019;
  KeyWrite = $20006;
  KeyWriteDac = $40000;
  KeyWow64_32Key = $0200;
  KeyWow64_64Key = $0100;
  ErrorFileNotFound = 2;
  ErrorNoMoreFiles = 18;
  SnapshotProcesses = 2;
  UrlEscapeAsUtf8 = $00040000;

type
  TProcessEntry32W = record
    Size: LongWord;
    Usage: LongWord;
    ProcessId: LongWord;
    DefaultHeapId: LongWord;
    ModuleId: LongWord;
    Threads: LongWord;
    ParentProcessId: LongWord;
    Priority: Longint;
    Flags: LongWord;
    ExeFile: array[0..259] of Char;
  end;
  TRegistrationBackup = record
    Root: Integer;
    Key: String;
    BackupKey: String;
    Existed: Boolean;
  end;
  TFileBackup = record
    Original: String;
    Saved: String;
  end;

var
  OutlookPath, OutlookArchitecture, AssemblyCodeBase, BackupBase, BackupDirectory: String;
  Backups: array of TRegistrationBackup;
  FileBackups: array of TFileBackup;
  RegistrationChanged, FilesMayHaveChanged, InstallationSucceeded: Boolean;

function CreateToolhelp32Snapshot(Flags, ProcessId: LongWord): THandle;
  external 'CreateToolhelp32Snapshot@kernel32.dll stdcall';
function Process32FirstW(Snapshot: THandle; var Entry: TProcessEntry32W): Boolean;
  external 'Process32FirstW@kernel32.dll stdcall';
function Process32NextW(Snapshot: THandle; var Entry: TProcessEntry32W): Boolean;
  external 'Process32NextW@kernel32.dll stdcall';
function CloseHandle(Handle: THandle): Boolean;
  external 'CloseHandle@kernel32.dll stdcall';
function GetCurrentProcessId: LongWord;
  external 'GetCurrentProcessId@kernel32.dll stdcall';
function UrlCreateFromPathW(Path, Url: String; var Length: LongWord; Flags: LongWord): Integer;
  external 'UrlCreateFromPathW@shlwapi.dll stdcall';
function UrlEscapeW(Url, Escaped: String; var Length: LongWord; Flags: LongWord): Integer;
  external 'UrlEscapeW@shlwapi.dll stdcall';
function ExpandEnvironmentStringsW(Source, Destination: String; Size: LongWord): LongWord;
  external 'ExpandEnvironmentStringsW@kernel32.dll stdcall';
function RegOpenKeyExW(Root: THandle; Subkey: String; Options, Access: LongWord; var Key: THandle): Integer;
  external 'RegOpenKeyExW@advapi32.dll stdcall';
function RegCreateKeyExW(Root: THandle; Subkey: String; Reserved: LongWord; ClassName: THandle;
  Options, Access: LongWord; Security: THandle; var Key: THandle; var Disposition: LongWord): Integer;
  external 'RegCreateKeyExW@advapi32.dll stdcall';
function RegCopyTreeW(Source: THandle; Subkey: THandle; Destination: THandle): Integer;
  external 'RegCopyTreeW@advapi32.dll stdcall';
function RegCloseKey(Key: THandle): Integer;
  external 'RegCloseKey@advapi32.dll stdcall';
function SetFileAttributesW(Name: String; Attributes: LongWord): Boolean;
  external 'SetFileAttributesW@kernel32.dll stdcall';

function ClsidKey: String;
begin
  Result := RegistrationPrefix + 'Software\Classes\CLSID\' + ClassId;
end;

function ProgIdKey: String;
begin
  Result := RegistrationPrefix + 'Software\Classes\' + ProgId;
end;

function AddinKey: String;
begin
  Result := RegistrationPrefix + 'Software\Microsoft\Office\Outlook\Addins\' + ProgId;
end;

function ViewFlag(Root: Integer): LongWord;
begin
  if Root = HKCU64 then Result := KeyWow64_64Key
  else Result := KeyWow64_32Key;
end;

procedure CheckRegistryStatus(Status: Integer; const Operation, Key: String);
begin
  if Status <> 0 then
    RaiseException(Operation + ' failed for HKCU\' + Key + ': ' +
      SysErrorMessage(Status) + ' (' + IntToStr(Status) + ').');
end;

function KeyExistsChecked(Root: Integer; const Key: String): Boolean;
var
  Handle: THandle;
  Status: Integer;
begin
  Status := RegOpenKeyExW(HKCU, Key, 0, KeyRead or ViewFlag(Root), Handle);
  if Status = ErrorFileNotFound then begin
    Result := False;
    exit;
  end;
  CheckRegistryStatus(Status, 'Reading registration', Key);
  RegCloseKey(Handle);
  Result := True;
end;

function FileUri(const Path: String): String;
var
  Buffer, Encoded: String;
  Length: LongWord;
  Status: Integer;
begin
  Length := 32768;
  SetLength(Buffer, Length);
  Status := UrlCreateFromPathW(Path, Buffer, Length, 0);
  if Status < 0 then
    RaiseException('Cannot convert the installation path to a file URI (' + IntToStr(Status) + ').');
  SetLength(Buffer, Length);
  Length := 32768;
  SetLength(Encoded, Length);
  { UrlCreateFromPath escapes %, # and spaces; escape Unicode without escaping % a second time. }
  Status := UrlEscapeW(Buffer, Encoded, Length, UrlEscapeAsUtf8);
  if Status < 0 then
    RaiseException('Cannot UTF-8 encode the installation file URI (' + IntToStr(Status) + ').');
  SetLength(Encoded, Length);
  Result := Encoded;
end;

function OutlookProcessError: String;
var
  Snapshot: THandle;
  Entry: TProcessEntry32W;
  Found: Boolean;
  ErrorCode, I: Integer;
  Name: String;
begin
  Result := '';
  Snapshot := CreateToolhelp32Snapshot(SnapshotProcesses, 0);
  if Snapshot = THandle(-1) then begin
    ErrorCode := DLLGetLastError;
    Result := 'Cannot check whether Outlook is running: process snapshot failed. ' +
      SysErrorMessage(ErrorCode) + ' (' + IntToStr(ErrorCode) + ').';
    exit;
  end;
  try
    { Inno Setup 6 runs a 32-bit process, including on 64-bit Windows. }
    Entry.Size := 556;
    Found := Process32FirstW(Snapshot, Entry);
    if not Found then begin
      ErrorCode := DLLGetLastError;
      Result := 'Cannot check whether Outlook is running: process enumeration failed. ' +
        SysErrorMessage(ErrorCode) + ' (' + IntToStr(ErrorCode) + ').';
      exit;
    end;
    while Found do begin
      Name := '';
      I := 0;
      while (I < 260) and (Entry.ExeFile[I] <> #0) do begin
        Name := Name + Entry.ExeFile[I];
        I := I + 1;
      end;
      if CompareText(Name, CheckedProcessName) = 0 then begin
        Log('Blocking process detected: ' + CheckedProcessName + ' (PID ' +
          IntToStr(Entry.ProcessId) + ').');
        Result := 'Close classic Outlook before installing, upgrading, or uninstalling OrgLens, ' +
          'then try again. Outlook will never be closed automatically.';
        exit;
      end;
      Found := Process32NextW(Snapshot, Entry);
      if not Found then begin
        ErrorCode := DLLGetLastError;
        if ErrorCode <> ErrorNoMoreFiles then
          Result := 'Cannot check whether Outlook is running: process enumeration was incomplete. ' +
            SysErrorMessage(ErrorCode) + ' (' + IntToStr(ErrorCode) + ').';
      end;
    end;
  finally
    CloseHandle(Snapshot);
  end;
end;

function RequireOutlookClosed(Silent: Boolean): Boolean;
var
  Error: String;
begin
  Result := False;
  repeat
    Error := OutlookProcessError;
    if Error = '' then begin
      Result := True;
      exit;
    end;
    Log(Error);
    if Silent then exit;
  until SuppressibleMsgBox(Error, mbError, MB_RETRYCANCEL, IDCANCEL) <> IDRETRY;
end;

procedure CheckFramework;
var
  Release: Cardinal;
  Key: String;
begin
  Key := 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full';
  if not RegQueryDWordValue(HKLM32, Key, 'Release', Release) then Release := 0;
  if Release < 528040 then
    RaiseException('Microsoft .NET Framework 4.8 or later is required. Install it from Microsoft, ' +
      'complete any requested Windows restart, and run Setup again.');
  if IsWin64 then begin
    if not RegQueryDWordValue(HKLM64, Key, 'Release', Release) then Release := 0;
    if Release < 528040 then
      RaiseException('The 64-bit Microsoft .NET Framework 4.8 or later is required. ' +
        'Install or repair .NET Framework, then run Setup again.');
  end;
end;

function ExpandPath(const Value: String): String;
var
  Buffer: String;
  Count: LongWord;
begin
  SetLength(Buffer, 32768);
  Count := ExpandEnvironmentStringsW(Trim(Value), Buffer, 32768);
  if (Count = 0) or (Count > 32768) then
    RaiseException('Cannot expand the Outlook executable path.');
  SetLength(Buffer, Count - 1);
  Result := Trim(Buffer);
  if (Length(Result) >= 2) and (Result[1] = '"') and (Result[Length(Result)] = '"') then
    Result := Copy(Result, 2, Length(Result) - 2);
end;

function AppPath(Root: Integer): String;
var
  Value: String;
begin
  Result := '';
  if RegQueryStringValue(Root, 'Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
    '', Value) then begin
    Value := ExpandPath(Value);
    if FileExists(Value) then Result := Value
    else Log('Ignoring a missing Outlook App Paths target: ' + Value);
  end;
end;

function OfficePath(const Base: String): String;
var
  Version: Integer;
  Candidate: String;
begin
  Result := '';
  for Version := 16 downto 14 do begin
    Candidate := Base + '\Microsoft Office\root\Office' + IntToStr(Version) + '\OUTLOOK.EXE';
    if FileExists(Candidate) then begin
      Result := Candidate;
      exit;
    end;
    Candidate := Base + '\Microsoft Office\Office' + IntToStr(Version) + '\OUTLOOK.EXE';
    if FileExists(Candidate) then begin
      Result := Candidate;
      exit;
    end;
  end;
end;

function DetectOutlook: String;
begin
#ifdef TestOutlookPath
  Result := '{#StringChange(TestOutlookPath, "'", "''")}';
  if not FileExists(Result) then Result := '';
#else
  Result := AppPath(HKCU32);
  if (Result = '') and IsWin64 then Result := AppPath(HKCU64);
  if Result = '' then Result := AppPath(HKLM32);
  if (Result = '') and IsWin64 then Result := AppPath(HKLM64);
  if (Result = '') and IsWin64 then Result := OfficePath(ExpandConstant('{pf64}'));
  if Result = '' then Result := OfficePath(ExpandConstant('{pf32}'));
#endif
  if Result = '' then
    RaiseException('Classic Microsoft Outlook was not found. Install classic Outlook and run Setup again. ' +
      'The new Outlook for Windows does not support this COM add-in.');
end;

function ReadOutlookArchitecture(const Path: String): String;
var
  Stream: TFileStream;
  Data: AnsiString;
  ReadPath, NativeSystemDirectory: String;
  Offset: Int64;
  Machine: Integer;
begin
  ReadPath := Path;
  NativeSystemDirectory := ExpandConstant('{win}\System32\');
  { Inspect the specified native file, not WOW64's redirected System32 counterpart. }
  if IsWin64 and
    (CompareText(Copy(Path, 1, Length(NativeSystemDirectory)), NativeSystemDirectory) = 0) then
    ReadPath := ExpandConstant('{sysnative}\') + Copy(Path, Length(NativeSystemDirectory) + 1, MaxInt);
  Stream := TFileStream.Create(ReadPath, fmOpenRead or fmShareDenyNone);
  try
    if Stream.Size < 64 then RaiseException('Outlook has an invalid executable header: ' + Path);
    SetLength(Data, 64);
    Stream.ReadBuffer(Data, 64);
    if Copy(Data, 1, 2) <> 'MZ' then RaiseException('Outlook is not a Windows executable: ' + Path);
    Offset := Ord(Data[61]) + Ord(Data[62]) * 256 + Ord(Data[63]) * 65536;
    Offset := Offset + Int64(Ord(Data[64])) * 16777216;
    if (Offset < 64) or (Offset > Stream.Size - 6) then
      RaiseException('Outlook has an invalid PE header offset: ' + Path);
    Stream.Seek(Offset, soFromBeginning);
    SetLength(Data, 6);
    Stream.ReadBuffer(Data, 6);
    if Copy(Data, 1, 4) <> 'PE' + #0 + #0 then
      RaiseException('Outlook has an invalid PE signature: ' + Path);
    Machine := Ord(Data[5]) + Ord(Data[6]) * 256;
    case Machine of
      $014C: Result := 'x86';
      $8664: begin
        if not IsWin64 then RaiseException('64-bit Outlook requires 64-bit Windows.');
        Result := 'x64';
      end;
    else
      RaiseException('Unsupported Outlook executable architecture (PE machine ' +
        IntToStr(Machine) + '). OrgLens supports x86 and x64 classic Outlook, not ARM64 Outlook.');
    end;
  finally
    Stream.Free;
  end;
end;

procedure ExpectString(Root: Integer; const Key, Name, Expected: String);
var
  Value: String;
begin
  if not RegQueryStringValue(Root, Key, Name, Value) or (CompareText(Value, Expected) <> 0) then
    RaiseException('Existing registration is not owned by this OrgLens installer: HKCU\' +
      Key + ' [' + Name + ']. No foreign registration will be overwritten or removed.');
end;

procedure CheckManagedServer(Root: Integer; const Key: String; Uninstalling: Boolean);
var
  Value: String;
begin
  ExpectString(Root, Key, 'Class', ComClass);
  if not RegQueryStringValue(Root, Key, 'Assembly', Value) or
    (Pos('OrgLens.Outlook, Version=', Value) <> 1) then
    RaiseException('Unexpected assembly ownership at HKCU\' + Key + '.');
  ExpectString(Root, Key, 'RuntimeVersion', 'v4.0.30319');
  if Uninstalling then
    ExpectString(Root, Key, 'CodeBase', AssemblyCodeBase);
end;

procedure CheckOwnership(Root: Integer; Uninstalling: Boolean);
var
  HasClass, HasProgId, HasAddin: Boolean;
  Names: TArrayOfString;
  I: Integer;
  Version, NewVersion: Int64;
  Server: String;
begin
  HasClass := KeyExistsChecked(Root, ClsidKey);
  HasProgId := KeyExistsChecked(Root, ProgIdKey);
  HasAddin := KeyExistsChecked(Root, AddinKey);
  if not (HasClass or HasProgId or HasAddin) then exit;
  { Outlook's HKCU Addins and ProgID keys can be shared across registry views. }
  if HasProgId then ExpectString(Root, ProgIdKey + '\CLSID', '', ClassId);
  if HasClass then begin
    ExpectString(Root, ClsidKey, '', ComClass);
    ExpectString(Root, ClsidKey + '\ProgId', '', ProgId);
    Server := ClsidKey + '\InprocServer32';
    ExpectString(Root, Server, '', 'mscoree.dll');
    CheckManagedServer(Root, Server, Uninstalling);
    if not RegGetSubkeyNames(Root, Server, Names) then
      RaiseException('Cannot enumerate existing OrgLens assembly registrations.');
    if not StrToVersion(CurrentAssemblyVersion, NewVersion) then
      RaiseException('This installer has an invalid assembly version.');
    for I := 0 to GetArrayLength(Names) - 1 do begin
      if not StrToVersion(Names[I], Version) then
        RaiseException('Unexpected COM registration subkey: ' + Server + '\' + Names[I]);
      CheckManagedServer(Root, Server + '\' + Names[I], Uninstalling);
      if not Uninstalling and (ComparePackedVersion(Version, NewVersion) > 0) then
        RaiseException('A newer OrgLens version is already registered. Install that version or later; ' +
          'Setup will not downgrade an existing installation.');
    end;
  end;
  if HasAddin then begin
    ExpectString(Root, AddinKey, 'FriendlyName', 'OrgLens');
    if not HasProgId then
      RaiseException('An Outlook OrgLens add-in registration has no matching ProgID. ' +
        'Setup will not replace or remove this unexpected registration.');
  end;
  if HasProgId and not HasClass then begin
    if not IsWin64 then
      RaiseException('The existing OrgLens ProgID has no matching COM class.');
    if Root = HKCU32 then HasClass := KeyExistsChecked(HKCU64, ClsidKey)
    else HasClass := KeyExistsChecked(HKCU32, ClsidKey);
    if not HasClass then
      RaiseException('The existing OrgLens ProgID has no matching COM class in either registry view.');
  end;
end;

procedure CheckAllOwnership(Uninstalling: Boolean);
begin
  CheckOwnership(HKCU32, Uninstalling);
  if IsWin64 then CheckOwnership(HKCU64, Uninstalling);
end;

procedure CopyRegistryTree(SourceRoot: Integer; const Source: String;
  DestinationRoot: Integer; const Destination: String);
var
  SourceHandle, DestinationHandle: THandle;
  Disposition: LongWord;
begin
  CheckRegistryStatus(RegOpenKeyExW(HKCU, Source, 0, KeyRead or ViewFlag(SourceRoot),
    SourceHandle), 'Opening registry backup source', Source);
  try
    { RegCopyTree also copies the security descriptor, which requires WRITE_DAC. }
    CheckRegistryStatus(RegCreateKeyExW(HKCU, Destination, 0, 0, 0,
      KeyRead or KeyWrite or KeyWriteDac or ViewFlag(DestinationRoot), 0, DestinationHandle, Disposition),
      'Opening registry backup destination', Destination);
    try
      CheckRegistryStatus(RegCopyTreeW(SourceHandle, 0, DestinationHandle),
        'Copying registration for rollback', Source);
    finally
      RegCloseKey(DestinationHandle);
    end;
  finally
    RegCloseKey(SourceHandle);
  end;
end;

procedure BackupKey(Root: Integer; const Key: String);
var
  Index: Integer;
begin
  Index := GetArrayLength(Backups);
  SetArrayLength(Backups, Index + 1);
  Backups[Index].Root := Root;
  Backups[Index].Key := Key;
  Backups[Index].BackupKey := BackupBase + '\' + IntToStr(Index);
  Backups[Index].Existed := KeyExistsChecked(Root, Key);
  if Backups[Index].Existed then
    CopyRegistryTree(Root, Key, HKCU32, Backups[Index].BackupKey);
end;

procedure BackupRegistration;
begin
  BackupBase := RegistrationPrefix + 'Software\OrgLens\InstallerRollback\' +
    GetDateTimeString('yyyymmddhhnnss', #0, #0) + '-' + IntToStr(GetCurrentProcessId);
  Log('Registration rollback location: HKCU\' + BackupBase);
  BackupKey(HKCU32, ClsidKey);
  BackupKey(HKCU32, ProgIdKey);
  BackupKey(HKCU32, AddinKey);
  if IsWin64 then begin
    BackupKey(HKCU64, ClsidKey);
    BackupKey(HKCU64, ProgIdKey);
    BackupKey(HKCU64, AddinKey);
  end;
end;

procedure DeleteOwnedTree(Root: Integer; const Key: String);
begin
  if KeyExistsChecked(Root, Key) and not RegDeleteKeyIncludingSubkeys(Root, Key) then
    RaiseException('Cannot remove owned OrgLens registration: HKCU\' + Key + '.');
end;

procedure RestoreRegistration;
var
  I: Integer;
begin
  Log('Restoring pre-install COM and Outlook registration.');
  for I := 0 to GetArrayLength(Backups) - 1 do begin
    DeleteOwnedTree(Backups[I].Root, Backups[I].Key);
    if Backups[I].Existed then
      CopyRegistryTree(HKCU32, Backups[I].BackupKey, Backups[I].Root, Backups[I].Key);
  end;
  RegistrationChanged := False;
end;

procedure BackupManagedFile(const RelativePath: String);
var
  Original, Saved: String;
  Index: Integer;
begin
  Original := ExpandConstant('{app}\') + RelativePath;
  if not FileExists(Original) then exit;
  if not ForceDirectories(BackupDirectory) then
    RaiseException('Cannot create the upgrade rollback directory: ' + BackupDirectory);
  Index := GetArrayLength(FileBackups);
  Saved := BackupDirectory + '\' + IntToStr(Index) + '.bak';
  if not CopyFile(Original, Saved, True) then
    RaiseException('Cannot back up the existing OrgLens file: ' + Original);
  SetArrayLength(FileBackups, Index + 1);
  FileBackups[Index].Original := Original;
  FileBackups[Index].Saved := Saved;
end;

procedure BackupManagedFiles;
begin
  BackupDirectory := ExpandConstant('{#RollbackDirectory}') + '\' +
    GetDateTimeString('yyyymmddhhnnss', #0, #0) + '-' + IntToStr(GetCurrentProcessId);
  Log('Managed-file rollback directory: ' + BackupDirectory);
  { Inno's standard rollback does not restore overwritten files on an upgrade. }
  BackupManagedFile('OrgLens.Outlook.dll');
  BackupManagedFile('OrgLens.Core.dll');
  BackupManagedFile('OrgLens.Desktop.dll');
  BackupManagedFile('Newtonsoft.Json.dll');
  BackupManagedFile('Newtonsoft.Json.LICENSE.md');
  BackupManagedFile('Lucide.LICENSE.txt');
  BackupManagedFile('OrgLens.Preview.exe');
  BackupManagedFile('OrgLens.Preview.exe.config');
  BackupManagedFile('README.md');
  BackupManagedFile('CHANGELOG.md');
  BackupManagedFile('docs\images\orglens-settings.png');
  BackupManagedFile('LICENSE');
end;

procedure RestoreManagedFiles;
var
  I: Integer;
  Error: String;
begin
  Error := '';
  for I := 0 to GetArrayLength(FileBackups) - 1 do begin
    if not ForceDirectories(ExtractFileDir(FileBackups[I].Original)) or
      not CopyFile(FileBackups[I].Saved, FileBackups[I].Original, False) then begin
      Log('Cannot restore the previous file: ' + FileBackups[I].Original);
      Error := 'Cannot restore the previous file: ' + FileBackups[I].Original;
    end;
  end;
  if Error <> '' then RaiseException(Error);
  FilesMayHaveChanged := False;
end;

procedure DeleteFileBackups;
var
  I: Integer;
begin
  for I := 0 to GetArrayLength(FileBackups) - 1 do begin
    SetFileAttributesW(FileBackups[I].Saved, $80);
    if FileExists(FileBackups[I].Saved) and not DeleteFile(FileBackups[I].Saved) then
      RaiseException('Cannot remove an upgrade backup: ' + FileBackups[I].Saved);
  end;
  if (BackupDirectory <> '') and DirExists(BackupDirectory) then
    RemoveDir(BackupDirectory);
end;

procedure WriteString(Root: Integer; const Key, Name, Value: String);
begin
  if not RegWriteStringValue(Root, Key, Name, Value) then
    RaiseException('Cannot write OrgLens registration: HKCU\' + Key + ' [' + Name + '].');
end;

procedure WriteDWord(Root: Integer; const Key, Name: String; Value: Cardinal);
begin
  if not RegWriteDWordValue(Root, Key, Name, Value) then
    RaiseException('Cannot write OrgLens registration: HKCU\' + Key + ' [' + Name + '].');
end;

procedure WriteManagedServer(Root: Integer; const Key: String);
begin
  WriteString(Root, Key, 'Class', ComClass);
  WriteString(Root, Key, 'Assembly', CurrentAssembly);
  WriteString(Root, Key, 'RuntimeVersion', 'v4.0.30319');
  WriteString(Root, Key, 'CodeBase', AssemblyCodeBase);
end;

procedure RegisterView(Root: Integer);
var
  Server: String;
  Names: TArrayOfString;
  I: Integer;
begin
  Server := ClsidKey + '\InprocServer32';
  if KeyExistsChecked(Root, Server) then begin
    if not RegGetSubkeyNames(Root, Server, Names) then
      RaiseException('Cannot enumerate prior OrgLens assembly versions.');
    for I := 0 to GetArrayLength(Names) - 1 do
      if Names[I] <> CurrentAssemblyVersion then
        DeleteOwnedTree(Root, Server + '\' + Names[I]);
  end;
  WriteString(Root, ProgIdKey, '', 'OrgLens classic Outlook add-in');
  WriteString(Root, ProgIdKey + '\CLSID', '', ClassId);
  WriteString(Root, ClsidKey, '', ComClass);
  WriteString(Root, ClsidKey + '\ProgId', '', ProgId);
  WriteString(Root, Server, '', 'mscoree.dll');
  WriteString(Root, Server, 'ThreadingModel', 'Both');
  WriteManagedServer(Root, Server);
  WriteManagedServer(Root, Server + '\' + CurrentAssemblyVersion);
  if not RegWriteStringValue(Root, ClsidKey + '\Implemented Categories\' + ManagedCategory, '', '') then
    RaiseException('Cannot register the managed COM category.');
  WriteString(Root, AddinKey, 'FriendlyName', 'OrgLens');
  WriteString(Root, AddinKey, 'Description', 'Highlight emails from your management chain.');
  if not RegValueExists(Root, AddinKey, 'LoadBehavior') then
    WriteDWord(Root, AddinKey, 'LoadBehavior', 3);
  WriteDWord(Root, AddinKey, 'CommandLineSafe', 0);
end;

procedure InstallRegistration;
var
  Error: String;
begin
  Error := OutlookProcessError;
  if Error <> '' then RaiseException(Error);
  CheckAllOwnership(False);
  RegistrationChanged := True;
  RegisterView(HKCU32);
  { AnyCPU .NET assemblies support either Outlook bitness; refresh both views on upgrades. }
  if IsWin64 then RegisterView(HKCU64);
  Log('OrgLens registered. Detected Outlook: ' + OutlookArchitecture + ', ' + OutlookPath);
end;

function InitializeSetup: Boolean;
begin
  Result := False;
  try
    CheckFramework;
    OutlookPath := DetectOutlook;
    OutlookArchitecture := ReadOutlookArchitecture(OutlookPath);
    Log('Detected classic Outlook ' + OutlookArchitecture + ': ' + OutlookPath);
    CheckAllOwnership(False);
    Result := RequireOutlookClosed(WizardSilent);
  except
    Log(GetExceptionMessage);
    SuppressibleMsgBox(GetExceptionMessage, mbError, MB_OK, IDOK);
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  NeedsRestart := False;
  Result := '';
  try
    CheckFramework;
    OutlookPath := DetectOutlook;
    OutlookArchitecture := ReadOutlookArchitecture(OutlookPath);
    AssemblyCodeBase := FileUri(ExpandConstant('{app}\OrgLens.Outlook.dll'));
    CheckAllOwnership(False);
    Result := OutlookProcessError;
  except
    Result := GetExceptionMessage;
  end;
  if Result <> '' then Log(Result);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Error: String;
begin
  if CurStep = ssInstall then begin
    Error := OutlookProcessError;
    if Error <> '' then RaiseException(Error);
    CheckAllOwnership(False);
    BackupRegistration;
    BackupManagedFiles;
    FilesMayHaveChanged := True;
  end;
  if CurStep = ssPostInstall then InstallationSucceeded := True;
end;

procedure DeinitializeSetup;
var
  RecoveryError: String;
begin
  RecoveryError := '';
  if not InstallationSucceeded then begin
    try
      if FilesMayHaveChanged then RestoreManagedFiles;
    except
      RecoveryError := GetExceptionMessage;
      Log('Managed-file recovery error: ' + RecoveryError);
    end;
    try
      if RegistrationChanged then RestoreRegistration;
    except
      RecoveryError := RecoveryError + #13#10 + GetExceptionMessage;
      Log('Registration recovery error: ' + GetExceptionMessage);
    end;
  end;
  try
    if RecoveryError <> '' then RaiseException(RecoveryError);
    if (BackupBase <> '') and KeyExistsChecked(HKCU32, BackupBase) then
      DeleteOwnedTree(HKCU32, BackupBase);
    DeleteFileBackups;
  except
    Log('Upgrade recovery/cleanup error: ' + GetExceptionMessage);
    SuppressibleMsgBox('OrgLens could not finish upgrade recovery or backup cleanup. ' +
      'Keep any remaining backups at HKCU\' + BackupBase + ' and ' + BackupDirectory +
      ' for recovery.' + #13#10 + GetExceptionMessage, mbError, MB_OK, IDOK);
  end;
end;

function InitializeUninstall: Boolean;
begin
  Result := False;
  try
    AssemblyCodeBase := FileUri(ExpandConstant('{app}\OrgLens.Outlook.dll'));
    CheckAllOwnership(True);
    Result := RequireOutlookClosed(UninstallSilent);
  except
    Log(GetExceptionMessage);
    SuppressibleMsgBox(GetExceptionMessage, mbError, MB_OK, IDOK);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Error: String;
begin
  if CurUninstallStep = usUninstall then begin
    Error := OutlookProcessError;
    if Error <> '' then RaiseException(Error);
    CheckAllOwnership(True);
    BackupRegistration;
    RegistrationChanged := True;
    try
      DeleteOwnedTree(HKCU32, AddinKey);
      DeleteOwnedTree(HKCU32, ProgIdKey);
      DeleteOwnedTree(HKCU32, ClsidKey);
      if IsWin64 then begin
        DeleteOwnedTree(HKCU64, AddinKey);
        DeleteOwnedTree(HKCU64, ProgIdKey);
        DeleteOwnedTree(HKCU64, ClsidKey);
      end;
      RegistrationChanged := False;
    except
      RestoreRegistration;
      RaiseException('OrgLens could not remove its registration. The previous registration was restored.');
    end;
  end;
end;

procedure DeinitializeUninstall;
begin
  try
    if RegistrationChanged then RestoreRegistration;
    if (BackupBase <> '') and KeyExistsChecked(HKCU32, BackupBase) then
      DeleteOwnedTree(HKCU32, BackupBase);
  except
    Log('Uninstall registration recovery error: ' + GetExceptionMessage);
    SuppressibleMsgBox('OrgLens could not finish registration recovery or backup cleanup. ' +
      'Keep any remaining backup at HKCU\' + BackupBase + '.' + #13#10 +
      GetExceptionMessage, mbError, MB_OK, IDOK);
  end;
end;
