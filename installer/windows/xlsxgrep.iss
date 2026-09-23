; Inno Setup script for xlsxgrep — installs the standalone CLI executable
; and adds its folder to the user PATH so `xlsxgrep` can be run from any terminal.

#ifndef AppVersion
#define AppVersion "0.0.0"
#endif

#define AppName "xlsxgrep"
#define AppPublisher "Ivan Cvitic"
#define AppURL "https://github.com/zazuum/xlsxgrep"
#define AppExeName "xlsxgrep.exe"

[Setup]
AppId={{4E7C4B7A-6C8A-4C3E-9F3E-1B2C3D4E5F60}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DisableProgramGroupPage=yes
DisableDirPage=no
OutputBaseFilename=xlsxgrep-setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "..\..\dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "addtopath"; Description: "Add xlsxgrep to PATH (recommended)"; GroupDescription: "Additional options:"

[Code]
procedure AddToPath();
var
  Path: string;
begin
  if RegQueryStringValue(HKEY_CURRENT_USER, 'Environment', 'Path', Path) then
  begin
    if Pos(ExpandConstant('{app}'), Path) = 0 then
    begin
      Path := Path + ';' + ExpandConstant('{app}');
      RegWriteStringValue(HKEY_CURRENT_USER, 'Environment', 'Path', Path);
    end;
  end
  else
  begin
    RegWriteStringValue(HKEY_CURRENT_USER, 'Environment', 'Path', ExpandConstant('{app}'));
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep = ssPostInstall) and IsTaskSelected('addtopath') then
  begin
    AddToPath();
  end;
end;
