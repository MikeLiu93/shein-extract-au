; Inno Setup script for SheinExtractAU.
; Compile with:  iscc installer.iss
; Produces: dist\SheinExtractAU-Setup-{version}.exe

#define MyAppName "SHEIN 上架工具 AU"
#define MyAppNameAscii "SheinExtractAU"
#define MyAppVersion "0.3.9"          ; Keep in sync with version.py
#define MyAppPublisher "MikeLiu93"
#define MyAppURL "https://github.com/MikeLiu93/shein-extract-au"
#define MyAppExeName "SheinExtractAU.exe"

[Setup]
; AppId is a unique GUID identifying this app — do NOT change between versions.
; Distinct from the 母项目 SheinExtract AppId so they coexist.
AppId={{4678792F-846D-485B-B525-84CD457338C5}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases

DefaultDirName={localappdata}\{#MyAppNameAscii}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

OutputDir=dist
OutputBaseFilename={#MyAppNameAscii}-Setup-{#MyAppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern

UsePreviousAppDir=yes
UsePreviousGroup=yes

ShowLanguageDialog=no

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "在桌面创建快捷方式"; GroupDescription: "附加任务:"; Flags: unchecked
Name: "startmenuicon"; Description: "在开始菜单创建快捷方式"; GroupDescription: "附加任务:"

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; SINGLE program icon — no separate "配置" icon. Menu handles that in-app.
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: startmenuicon
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"; Tasks: startmenuicon
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Preserve user data: %APPDATA%\shein-extract-au\, %USERPROFILE%\shein-cdp-profile-au\.
; Only files we drop in {app} get cleaned automatically.
