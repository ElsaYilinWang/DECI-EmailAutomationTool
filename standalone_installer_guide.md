The Plan: From Script to Installer
The goal is to create a professional setup.exe file that any user can run to install your application, without needing to install Python or any libraries themselves. We will do this in two stages:

Bundle with PyInstaller: We'll use a tool called PyInstaller to package your Python script, all its libraries, and even the Python interpreter into a single, standalone .exe file.

Create an Installer with Inno Setup: We'll use a tool called Inno Setup to take that .exe file and wrap it in a familiar installation wizard that handles license agreements, installation paths, and creating shortcuts.

Step 1: Bundle the App into a Single .exe with PyInstaller
Install PyInstaller
Open your terminal (with your .venv activated) and run:

python -m pip install pyinstaller

Prepare an Application Icon
For a professional look, you'll want an icon for your .exe file.

Find or create an image for your icon.

Use a free online converter (like icoconverter.com) to convert your image into a .ico file.

Save this file as icon.ico in your main project folder (C:\Users\ElsaWang\Documents\DECI-EmailAutomationTool).

Run the PyInstaller Command
In your terminal (make sure your .venv is active and you are in your project directory), run the following command. This tells PyInstaller to bundle everything into one file (--onefile), make it a windowed application so no console appears (--windowed), and use your icon (--icon).

pyinstaller --onefile --windowed --icon="icon.ico" email_app.py

Find Your Standalone .exe
PyInstaller will run for a minute or two. When it's finished, you will find a new folder named dist in your project directory. Inside dist, you will find email_app.exe. This is your complete application! You can copy this single file to another computer that doesn't have Python, and it will run.

Step 2: Create the Professional Installer with Inno Setup
Download and Install Inno Setup
This is a free and very popular installer creator for Windows. Download it from the official website and install it: jrsoftware.org

Prepare Required Files

License Agreement: Create a simple text file named license.txt in your project folder. Paste your terms and conditions into this file.

GUID: Go to a free online GUID generator like guidgenerator.com and generate one unique ID. You'll need this for the script below.

Create the Installer Script (.iss file)
Inno Setup uses a script to define how the installer works.

Open a plain text editor (like Notepad or VS Code).

Create a new file and save it as setup_script.iss in your main project folder.

Copy and paste the code below into that file. I have already configured it for your application. Remember to paste your unique GUID where indicated.

; Inno Setup Script for Email Automation Tool

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Paste your generated GUID here.
AppId={{YOUR-UNIQUE-APP-ID}}
AppName=Email Automation Tool
AppVersion=1.0
AppPublisher=DECI-LTD
DefaultDirName={autopf}\Email Automation Tool
DefaultGroupName=Email Automation Tool
AllowNoIcons=yes
LicenseFile=C:\Users\ElsaWang\Documents\DECI-EmailAutomationTool\license.txt
OutputDir=C:\Users\ElsaWang\Documents\DECI-EmailAutomationTool\Installer
OutputBaseFilename=EmailAutomationTool_Setup
SetupIconFile=C:\Users\ElsaWang\Documents\DECI-EmailAutomationTool\icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; This is the most important line. It points to the .exe you created with PyInstaller.
Source: "C:\Users\ElsaWang\Documents\DECI-EmailAutomationTool\dist\email_app.exe"; DestDir: "{app}"; Flags: ignoreversion
; You can add other files here, like a UserGuide.pdf if you have one.

[Icons]
Name: "{group}\Email Automation Tool"; Filename: "{app}\email_app.exe"
Name: "{autodesktop}\Email Automation Tool"; Filename: "{app}\email_app.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\email_app.exe"; Description: "{cm:LaunchProgram,Email Automation Tool}"; Flags: nowait postinstall skipifsilent

Compile the Installer

Open the Inno Setup Compiler application from your Start Menu.

Go to File > Open and select your setup_script.iss file.

Go to Build > Compile (or press F9).

Find Your Final Installer
When it's finished, go to the Installer folder (C:\Users\ElsaWang\Documents\DECI-EmailAutomationTool\Installer). You will find EmailAutomationTool_Setup.exe.

This is the final file you can distribute. It will provide the professional installation experience you described.