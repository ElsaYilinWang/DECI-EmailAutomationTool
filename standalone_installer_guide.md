### **Guide: Creating a Standalone .exe Installer**

The goal is to create a professional `setup.exe` file that any user can run to install your application, without needing to install Python or any libraries themselves. We will do this in two stages:

1.  **Bundle with PyInstaller:** We'll use a tool called **PyInstaller** to package your Python script, all its libraries, and even the Python interpreter into a single, standalone `.exe` file.
2.  **Create an Installer with Inno Setup:** We'll use a tool called **Inno Setup** to take that `.exe` file and wrap it in a familiar installation wizard that handles license agreements, installation paths, and creating shortcuts.

### **Step 1: Bundle the App into a Single `.exe` with PyInstaller**

1.  **Install Dependencies**
    Open your terminal (with your `.venv` activated) and install all required libraries:
    ```bash
    python -m pip install pywin32 python-json-logger pyinstaller
    ```

2.  **Prepare an Application Icon**
    For a professional look, you'll want an icon for your `.exe` file.
    * Find or create an image for your icon.
    * Use a free online converter (like `icoconverter.com`) to convert your image into a `.ico` file.
    * Save this file as `icon.ico` in your main project folder.

3.  **Run the PyInstaller Command**
    In your terminal (make sure your `.venv` is active and you are in your project directory), run the following command:
    ```bash
    pyinstaller --onefile --windowed --icon="icon.ico" email_app.py
    ```

4.  **Find Your Standalone `.exe`**
    When PyInstaller is finished, you will find a new folder named `dist` in your project directory. Inside `dist`, you will find **`email_app.exe`**. This is your complete application.

### **Step 2: Create the Professional Installer with Inno Setup**

1.  **Download and Install Inno Setup**
    Download the free installer creator from its official website and install it: [jrsoftware.org](https://jrsoftware.org/isinfo.php)

2.  **Prepare Required Files**
    * **License Agreement:** Create a text file named `license.txt` in your project folder with your EULA.
    * **Application Icon:** Ensure your `icon.ico` file is in the project folder.

3.  **Create the Installer Script (`.iss` file)**
    Create a new file named **`setup_script.iss`** in your main project folder and paste the code below into it. All file paths are relative, making this script portable.

    ```pascal
    ; Inno Setup Script for Email Automation Tool
    
    [Setup]
    ; NOTE: The value of AppId uniquely identifies this application.
    AppId={{a12d6ee9-3451-4c8f-a59c-a3dab4434506}}
    AppName=Email Automation Tool
    AppVersion=1.1
    AppPublisher=Elsa Wang
    DefaultDirName={autopf}\Email-Automation-Tool
    DefaultGroupName=Email-Automation-Tool
    AllowNoIcons=yes
    LicenseFile=license.txt
    OutputDir=Installer
    OutputBaseFilename=EmailAutomationTool_Setup_v1.1
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
    ```

4.  **Compile the Installer**
    * Open the **Inno Setup Compiler** application (run as administrator if you encounter permission errors).
    * Go to **File > Open** and select your `setup_script.iss` file.
    * Go to **Build > Compile** (or press F9).

5.  **Find Your Final Installer**
    When it's finished, go to the `Installer` folder (which will be created inside your project directory). You will find your setup file (e.g., `EmailAutomationTool_Setup_v1.1.exe`).

### **Workflow for Updates**

If you make any changes to your `email_app.py` script, you **must** rebuild the installer:

1.  **Re-bundle the `.exe`:** Run the `pyinstaller` command again to create a new `email_app.exe`.
2.  **Re-compile the Installer:** Open your `setup_script.iss` in Inno Setup and re-compile it. It will automatically package the new `email_app.exe` into a new installer.
