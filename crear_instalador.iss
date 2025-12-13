; Script de Inno Setup para crear el instalador
; Requiere Inno Setup: https://jrsoftware.org/isinfo.php

[Setup]
AppName=Demo
AppVersion=1.0
AppPublisher=Tu Empresa
AppPublisherURL=
AppSupportURL=
AppUpdatesURL=
DefaultDirName={autopf}\Demo
DefaultGroupName=Demo
AllowNoIcons=yes
LicenseFile=
OutputDir=instalador
OutputBaseFilename=Demo_Instalador
SetupIconFile=
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
Source: "dist\Demo.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "plantilla.xlsx"; DestDir: "{app}"; Flags: ignoreversion
; Nota: La plantilla también está empaquetada en el EXE, pero la incluimos por separado por si acaso

[Icons]
Name: "{group}\Demo"; Filename: "{app}\Demo.exe"
Name: "{group}\{cm:UninstallProgram,Demo}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Demo"; Filename: "{app}\Demo.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\Demo"; Filename: "{app}\Demo.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\Demo.exe"; Description: "{cm:LaunchProgram,Demo}"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

