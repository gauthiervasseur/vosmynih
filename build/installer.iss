; Inno Setup 6 script for восьминіг (called by build_windows.ps1)
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#ifndef SourceDir
  #define SourceDir "..\dist\vosmynih"
#endif

[Setup]
AppId={{6A1F3C2E-8B0D-4E7A-9C55-0C70C7095A11}
AppName=восьминіг
AppVersion={#AppVersion}
AppPublisher=восьминіг
DefaultDirName={autopf}\vosmynih
DefaultGroupName=восьминіг
UninstallDisplayIcon={app}\vosmynih.exe
OutputBaseFilename=vosmynih-{#AppVersion}-windows-x64-setup
SetupIconFile=..\app\assets\icon.ico
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\восьминіг"; Filename: "{app}\vosmynih.exe"
Name: "{autodesktop}\восьминіг"; Filename: "{app}\vosmynih.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\vosmynih.exe"; Description: "Launch восьминіг"; Flags: nowait postinstall skipifsilent
