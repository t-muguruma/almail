[Setup]
AppName=Petal
AppVersion=1.0.0
DefaultDirName={autopf}\Petal
DefaultGroupName=Petal
UninstallDisplayIcon={app}\Petal.exe
Compression=lzma
SolidCompression=yes
OutputDir=W:\myProjects\Petal
OutputBaseFilename=Petal_Setup
SetupIconFile=Petal_icon.ico
PrivilegesRequired=lowest

[Files]
Source: "Petal\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Petal"; Filename: "{app}\Petal.exe"
Name: "{autodesktop}\Petal"; Filename: "{app}\Petal.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Run]
Filename: "{app}\Petal.exe"; Description: "{cm:LaunchProgram,Petal}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\Petal"