#define MyAppName "Tool Spy Idea"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "ChThanh"
#define MyAppURL "https://www.facebook.com/kenju152004"
#define MyAppExeName "ToolSpyIdea.exe"

[Setup]
AppId={{B8F3C2A1-5D7E-4F9A-A1B3-C8D2E6F7A9B0}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer_output
OutputBaseFilename=Setup_ToolSpyIdea_v{#MyAppVersion}
SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
WizardImageFile=wizard_image.bmp
WizardSmallImageFile=wizard_small.bmp

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"

[Files]
Source: "dist\ToolSpyIdea\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Tool Spy Idea"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\data"
Type: filesandordirs; Name: "{app}\static"
Type: filesandordirs; Name: "{app}\modules"
Type: filesandordirs; Name: "{app}\extension"
Type: filesandordirs; Name: "{app}\etsy_profile"
Type: filesandordirs; Name: "{app}\chrome_app_profile"
Type: filesandordirs; Name: "{app}\playwright"
Type: filesandordirs; Name: "{app}\_internal"
Type: filesandordirs; Name: "{app}"

[UninstallRun]
Filename: "taskkill"; Parameters: "/f /im ToolSpyIdea.exe"; Flags: runhidden
