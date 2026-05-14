; BMES NG Rate Standalone — Inno Setup script
; Build with:  ISCC.exe installer.iss
; Produces:    dist\BmesNgRateStandalone_Setup.exe

#define MyAppName        "BMES NG Rate"
#define MyAppExeName     "BmesNgRateStandalone.exe"
#ifndef MyAppVersion
#define MyAppVersion     "1.0.0"
#endif
#define MyAppPublisher   "Personal"
#define PublishDir       "bin\Release\net8.0-windows\win-x64\publish"
#define DistDir          "dist"

[Setup]
AppId={{8A7C9D40-1F84-4C5E-B6E1-2F1E9B0D5A22}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableDirPage=no
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir={#DistDir}
OutputBaseFilename=BmesNgRateStandalone_Setup
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "korean";  MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; .exe + DLLs at the publish root
Source: "{#PublishDir}\*"; Excludes: "wwwroot\*,data\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; static assets (wwwroot)
Source: "{#PublishDir}\wwwroot\*"; DestDir: "{app}\wwwroot"; Flags: ignoreversion recursesubdirs createallsubdirs
; bundled DBs — credentials already cleared, RoutingTable / ReasonTable / ModelGroups preserved
; "onlyifdoesntexist" so user-edited DBs are kept across reinstalls/upgrades
Source: "{#PublishDir}\data\*"; DestDir: "{app}\data"; Flags: recursesubdirs createallsubdirs onlyifdoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; do NOT delete user data on uninstall
