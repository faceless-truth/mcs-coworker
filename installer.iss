; MCS CoWorker - Inno Setup Script
; Produces: MCSCoWorker_Setup.exe
; Run via build_installer.bat

#define MyAppName      "MCS CoWorker"
#define MyAppVersion   "2.0"
#define MyAppPublisher "MC & S Accountants"
#define MyAppURL       "https://mcands.com.au"
#define MyAppId        "{B4E1F3A2-7C5D-4E8B-9F2A-1D6E3C4B5A7F}"

#ifndef MyBuildDir
  #define MyBuildDir "installer_build"
#endif
#ifndef MyRepoDir
  #define MyRepoDir "."
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
PrivilegesRequired=lowest
OutputDir={#MyRepoDir}\installer_output
OutputBaseFilename=MCSCoWorker_Setup
SetupIconFile={#MyRepoDir}\assets\icon.ico
WizardImageFile={#MyRepoDir}\assets\installer_banner.bmp
WizardSmallImageFile={#MyRepoDir}\assets\installer_icon.bmp
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
DisableDirPage=yes
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\assets\icon.ico
UninstallDisplayName={#MyAppName}
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";  Description: "Create a &desktop shortcut";              GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startupentry"; Description: "Start MCS CoWorker when &Windows starts"; GroupDescription: "Startup:";          Flags: unchecked

[Files]
Source: "{#MyBuildDir}\MCSCoWorker.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#MyBuildDir}\python\*"; DestDir: "{app}\python"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyBuildDir}\app\*"; DestDir: "{app}\app"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyRepoDir}\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
; WebView2 Evergreen bootstrapper — installed only if not already present.
Source: "{#MyBuildDir}\installer_extras\MicrosoftEdgeWebview2Setup.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Dirs]
Name: "{app}\data"; Permissions: users-full

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\MCSCoWorker.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\assets\icon.ico"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\MCSCoWorker.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\assets\icon.ico"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "{#MyAppName}"; ValueData: """{sys}\wscript.exe"" ""{app}\MCSCoWorker.vbs"""; Flags: uninsdeletevalue; Tasks: startupentry

[Run]
; Install the WebView2 Runtime first — pywebview's edgechromium backend
; depends on it. Skipped if the runtime is already installed.
Filename: "{tmp}\MicrosoftEdgeWebview2Setup.exe"; Parameters: "/silent /install"; \
    StatusMsg: "Installing Microsoft WebView2 Runtime..."; \
    Flags: runhidden waituntilterminated; \
    Check: not IsWebView2Installed
Filename: "{sys}\wscript.exe"; Parameters: """{app}\MCSCoWorker.vbs"""; WorkingDir: "{app}"; Description: "Launch {#MyAppName} now"; Flags: nowait postinstall skipifsilent

[Code]
// Detect whether the Edge WebView2 Runtime is already installed on the
// target machine. The Evergreen runtime registers under one of two known
// keys depending on per-machine vs per-user install. If either exists we
// skip the bootstrapper to avoid a redundant install.
function IsWebView2Installed: Boolean;
begin
  Result :=
    RegKeyExists(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BEB-235B8DE5E7E2}') or
    RegKeyExists(HKLM, 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BEB-235B8DE5E7E2}') or
    RegKeyExists(HKCU, 'Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BEB-235B8DE5E7E2}');
end;

procedure InitializeWizard();
begin
  WizardForm.WelcomeLabel2.Caption :=
    'This will install MCS CoWorker on your computer.' + #13#10 + #13#10 +
    'MCS CoWorker is your AI-powered accounting assistant. ' +
    'It monitors your inbox, tracks ASIC returns, follows up debtors, ' +
    'and prepares your daily briefing automatically.' + #13#10 + #13#10 +
    'The app updates itself silently whenever your administrator ' +
    'pushes a fix - no reinstall required.';
end;
