; Inno Setup Script for Email Automation Tool

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
AppId={{a12d6ee9-3451-4c8f-a59c-a3dab4434506}}
AppName=Email Automation Tool
AppVersion=1.6
AppPublisher=Yilin (Elsa) Wang
DefaultDirName={autopf}\Email-Automation-Tool
DefaultGroupName=Email-Automation-Tool
AllowNoIcons=yes
LicenseFile=license.txt
OutputDir=Installer
OutputBaseFilename=EmailAutomationTool_Setup_v1.6
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; This points to the .exe you created with PyInstaller.
Source: "dist\email_app.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Email Automation Tool"; Filename: "{app}\email_app.exe"
Name: "{autodesktop}\Email Automation Tool"; Filename: "{app}\email_app.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\email_app.exe"; Description: "{cm:LaunchProgram,Email Automation Tool}"; Flags: nowait postinstall skipifsilent