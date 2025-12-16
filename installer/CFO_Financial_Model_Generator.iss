; Inno Setup Script for CFO Financial Model Generator
; Download Inno Setup from: https://jrsoftware.org/isdl.php

#define MyAppName "CFO Financial Model Generator"
#define MyAppVersion "1.3.0"
#define MyAppPublisher "Vectra Finance LLC"
#define MyAppURL "https://vectrafinance.com"
#define MyAppExeName "CFO_Financial_Model_Generator.exe"

[Setup]
; Application identity
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Install location
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

; Output settings
OutputDir=..\dist\installer
OutputBaseFilename=CFO_Financial_Model_Generator_Setup_{#MyAppVersion}
SetupIconFile=..\webapp\icon.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

; Privileges
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; License and info
LicenseFile=license.txt
InfoBeforeFile=readme.txt

; Uninstaller
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; Main executable
Source: "..\dist\CFO_Financial_Model_Generator.exe"; DestDir: "{app}"; Flags: ignoreversion

; Include the DNA_Template.xlsm if it's not bundled in the exe
; Source: "..\webapp\DNA_Template.xlsm"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// Check if a previous version is installed and offer to uninstall
function InitializeSetup(): Boolean;
var
  UninstallKey: String;
  UninstallString: String;
  ResultCode: Integer;
begin
  Result := True;

  UninstallKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1';

  if RegQueryStringValue(HKEY_CURRENT_USER, UninstallKey, 'UninstallString', UninstallString) or
     RegQueryStringValue(HKEY_LOCAL_MACHINE, UninstallKey, 'UninstallString', UninstallString) then
  begin
    if MsgBox('A previous version of {#MyAppName} is already installed. Would you like to uninstall it first?',
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      Exec(RemoveQuotes(UninstallString), '/SILENT', '', SW_SHOW, ewWaitUntilTerminated, ResultCode);
    end;
  end;
end;
