### **Guide: Creating a Standalone .exe Installer**

This guide provides the definitive, step-by-step method for creating a professional `setup.exe` file that correctly includes all necessary libraries and handles all known issues.

---

### **Section 1: Initial Setup (Do this once per project)**

#### **1. Install Dependencies**
Open your terminal, activate your virtual environment (`.\.venv\Scripts\activate`), and install all libraries required for both running and building the application:
```bash
python -m pip install pywin32 python-json-logger pyinstaller
```

#### **2. Prepare Application Assets**
Before building, you need two files in your main project folder:
* `icon.ico`: An icon file for your application.
* `license.txt`: A text file containing your End-User License Agreement (EULA).

#### **3. Create the PyInstaller Build Script (`.spec` file)**
Using a `.spec` file is the most reliable way to build your `.exe` and ensure all libraries are included.

1.  **Generate the file:** In your activated terminal, run this command once:
    ```bash
    pyi-makespec --onefile --windowed --icon="icon.ico" email_app.py
    ```
    This creates a new file named `email_app.spec`.

2.  **Edit the file:** Open `email_app.spec` and find the `hiddenimports=[]` line. Add `'pythonjsonlogger'` to the list to fix the `ModuleNotFoundError` bug.
    ```python
    hiddenimports=['pythonjsonlogger'],
    ```

#### **4. Create the Inno Setup Installer Script (`.iss` file)**
Create a new file named `setup_script.iss` in your project folder and paste the following code into it. This script is portable and uses the latest version information.

```pascal
; Inno Setup Script for Email Automation Tool

[Setup]
AppId={{a12d6ee9-3451-4c8f-a59c-a3dab4434506}}
AppName=Email Automation Tool
AppVersion=1.4
AppPublisher=Yilin (Elsa) Wang
DefaultDirName={autopf}\Email-Automation-Tool
DefaultGroupName=Email-Automation-Tool
AllowNoIcons=yes
LicenseFile=license.txt
OutputDir=Installer
OutputBaseFilename=EmailAutomationTool_Setup_v1.4
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\email_app.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Email Automation Tool"; Filename: "{app}\email_app.exe"
Name: "{autodesktop}\Email Automation Tool"; Filename: "{app}\email_app.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\email_app.exe"; Description: "{cm:LaunchProgram,Email Automation Tool}"; Flags: nowait postinstall skipifsilent
```

---

### **Section 2: The Build & Compile Workflow (For New Versions)**

Follow these steps every time you update your `email_app.py` script and want to release a new version.

#### **Step 1: Clean Your Project Folder**
Before starting a new build, it's crucial to delete old build files to prevent errors. Delete the following folders from your project directory if they exist:
* `build`
* `dist`

#### **Step 2: Re-bundle the `.exe`**
1.  Activate your virtual environment (`.\.venv\Scripts\activate`).
2.  Run PyInstaller using your configured `.spec` file. This is the only command you need to build the `.exe`.
    ```bash
    pyinstaller email_app.spec
    ```
3.  A new, updated `email_app.exe` will be created in the `dist` folder.

#### **Step 3: Re-compile the Installer**
1.  Open the **Inno Setup Compiler** application. **Run it as an administrator** to prevent "file access denied" errors.
2.  Go to **File > Open** and select your `setup_script.iss` file.
3.  If you're creating a new version, remember to update the `AppVersion` and `OutputBaseFilename` in the script.
4.  Go to **Build > Compile**.

Your new, final installer (e.g., `EmailAutomationTool_Setup_v1.4.exe`) will be created in the `Installer` folder, ready for release.
