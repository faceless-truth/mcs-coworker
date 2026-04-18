; MCS CoWorker — Inno Setup Script
; Produces: MCSCoWorker_Setup.exe
; Run via build_installer.bat or manually:
;   iscc installer.iss /DMyBuildDir="C:\path\to\installer_build"

#define MyAppName      "MCS CoWorker"
#define MyAppVersion   "2.0"
#define MyAppPublisher "MC & S Accountants"
#define MyAppURL       "https://mcands.com.au"
#define MyAppExeName   "MCSCoWorker.exe"
#define MyAppId        "{B4E1F3A2-7C5D-4E8B-9F2A-1D6E3C4B5A7F}"

; MyBuildDir is passed in from build_installer.bat via /DMyBuildDir=...
; Default for manual runs:
#ifndef MyBuildDir
  #define MyBuildDir "installer_build"
#endif

[Setup]
AppId={{#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; Require admin so we can write to Program Files
PrivilegesRequired=admin
OutputDir=installer_output
OutputBaseFilename=MCSCoWorker_Setup
SetupIconFile=assets\icon.ico
WizardImageFile=assets\installer_banner.bmp
WizardSmallImageFile=assets\installer_icon.bmp
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
DisableDirPage=yes
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
; Minimum Windows version: Windows 10
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";    Description: "Create a &desktop shortcut";      GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startupentry";   Description: "Start MCS CoWorker when &Windows starts"; GroupDescription: "Startup:"; Flags: unchecked

[Files]
; Launcher executable
Source: "{#MyBuildDir}\MCSCoWorker.exe";    DestDir: "{app}";              Flags: ignoreversion

; Bundled Python runtime
Source: "{#MyBuildDir}\python\*";           DestDir: "{app}\python";       Flags: ignoreversion recursesubdirs createallsubdirs

; App source (the git repo clone)
Source: "{#MyBuildDir}\app\*";              DestDir: "{app}\app";          Flags: ignoreversion recursesubdirs createallsubdirs

; Assets (icon, etc.)
Source: "assets\*";                         DestDir: "{app}\assets";       Flags: ignoreversion recursesubdirs createallsubdirs; Check: AssetsExist

[Dirs]
; Create writable data directory for coworker.db and config
Name: "{app}\data"; Permissions: users-full

[Icons]
; Start Menu
Name: "{group}\{#MyAppName}";              Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\Uninstall {#MyAppName}";    Filename: "{uninstallexe}"

; Desktop shortcut (optional task)
Name: "{autodesktop}\{#MyAppName}";        Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Registry]
; Auto-start on Windows login (optional task)
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
  ValueType: string; ValueName: "{#MyAppName}"; \
  ValueData: """{app}\{#MyAppExeName}"""; \
  Flags: uninsdeletevalue; Tasks: startupentry

[Run]
; Launch the app after install (optional)
Filename: "{app}\{#MyAppExeName}"; \
  Description: "Launch {#MyAppName} now"; \
  Flags: nowait postinstall skipifsilent

[UninstallRun]
; Nothing special needed — uninstaller removes all files

[Code]
function AssetsExist(): Boolean;
begin
  Result := FileExists(ExpandConstant('{src}\assets\icon.ico'));
end;

procedure InitializeWizard();
begin
  WizardForm.WelcomeLabel2.Caption :=
    'This will install MCS CoWorker on your computer.' + #13#10 + #13#10 +
    'MCS CoWorker is your AI-powered accounting assistant. ' +
    'It monitors your inbox, tracks ASIC returns, follows up debtors, ' +
    'and prepares your daily briefing automatically.' + #13#10 + #13#10 +
    'The app updates itself silently whenever your administrator ' +
    'pushes a fix — no reinstall required.';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  AppDir, AppGitDir: String;
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    AppDir    := ExpandConstant('{app}\app');
    AppGitDir := AppDir + '\.git';

    // If the app directory does not have a .git folder, clone the repo
    if not DirExists(AppGitDir) then
    begin
      // Remove any partial copy so git clone can create a fresh one
      if DirExists(AppDir) then
        DelTree(AppDir, True, True, True);

      Exec('cmd.exe',
        '/c git clone --depth 1 https://github.com/faceless-truth/mcs-coworker.git "' + AppDir + '"',
        '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

      if ResultCode <> 0 then
        MsgBox(
          'Could not clone the app from GitHub.' + #13#10 +
          'Please ensure this machine has internet access and try again.',
          mbError, MB_OK);
    end;
  end;
end;
