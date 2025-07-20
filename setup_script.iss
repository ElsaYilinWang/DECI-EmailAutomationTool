; setup_script.iss

[Setup]
AppName=Email Automation Tool
AppVersion=1.8
DefaultDirName={autopf}\DECI Email Automation Tool
DefaultGroupName=DECI Email Automation Tool
UninstallDisplayIcon={app}\EmailAutomationTool.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
OutputBaseFilename=EmailAutomationTool_Setup_v1.8
; Corrected icon filename
SetupIconFile=icon.ico
UninstallFilesDir={app}
; Ensures the installer saves to the "Installer" folder
OutputDir=Installer

[Files]
; This now points to the correct folder created by the .spec file
Source: "dist\EmailAutomationTool\EmailAutomationTool.exe"; DestDir: "{app}"
; Include all other necessary files from the dist folder
Source: "dist\EmailAutomationTool\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Email Automation Tool"; Filename: "{app}\EmailAutomationTool.exe"
Name: "{autodesktop}\Email Automation Tool"; Filename: "{app}\EmailAutomationTool.exe"

[Run]
Filename: "{app}\EmailAutomationTool.exe"; Description: "Launch application"; Flags: nowait postinstall skipifsilent