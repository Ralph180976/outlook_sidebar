; InboxBar - Inno Setup Installer Script
; Creates a proper Windows installer with Add/Remove Programs support

#define MyAppName "InboxBar"
#define MyAppVersion "1.3.18"
#define MyAppPublisher "Coveya"
#define MyAppURL "https://github.com/Ralph180976/outlook_sidebar"
#define MyAppExeName "InboxBar.exe"
#define MyAppDescription "Outlook Sidebar - Email & Calendar at a glance"

[Setup]
; Unique App ID (do NOT change between versions - this is how Windows tracks upgrades)
AppId={{05C6F8E8-1329-4786-9C64-001004126162}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Install location (per-user, no admin required)
DefaultDirName={localappdata}\{#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

; Output installer
OutputDir=installer_output
OutputBaseFilename=InboxBar_Setup_v{#MyAppVersion}
SetupIconFile=icons\inboxbar.ico

; Compression
Compression=lzma2/ultra64
SolidCompression=yes

; Visual
WizardStyle=modern


; Version info shown in Add/Remove Programs
VersionInfoVersion={#MyAppVersion}.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppDescription}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}.0

; Uninstaller settings
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}

; Allow upgrading without uninstall
UsePreviousAppDir=yes

; Close running instances before install/update
CloseApplications=force
CloseApplicationsFilter=*.exe

; Minimum Windows version (Windows 10)
MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startupicon"; Description: "Start InboxBar automatically when Windows starts"; GroupDescription: "Startup:"; Flags: unchecked

[Files]
; Main application files from PyInstaller output
Source: "dist\InboxBar\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Start Menu entry
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Comment: "{#MyAppDescription}"
; Desktop shortcut (optional)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; Comment: "{#MyAppDescription}"
; Startup shortcut (optional)
Name: "{userstartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: startupicon

[Run]
; Launch after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; Kill running instance before uninstall
Filename: "taskkill"; Parameters: "/F /IM {#MyAppExeName}"; Flags: runhidden; RunOnceId: "KillApp"

[UninstallDelete]
; Clean up config and cache files on uninstall
Type: filesandordirs; Name: "{localappdata}\OutlookSidebar"

[Code]
// Kill running instances before install/upgrade
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  // Try to kill any running instance
  Exec('taskkill', '/F /IM {#MyAppExeName}', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := True;
end;

// Remove old startup shortcut if user unchecks the option during update
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if not WizardIsTaskSelected('startupicon') then
    begin
      DeleteFile(ExpandConstant('{userstartup}\{#MyAppName}.lnk'));
    end;
  end;
end;
