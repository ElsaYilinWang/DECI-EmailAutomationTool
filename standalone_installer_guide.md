### **Guide: Creating a Standalone .exe Installer**

This guide provides the definitive method for creating a professional `setup.exe` file that correctly includes all necessary libraries.

---

### **Step 1: Create and Configure the Build Script (`.spec` file)**

Instead of a long command, we will use a `.spec` file to give PyInstaller explicit instructions. This is the most reliable way to ensure all libraries are included.

1.  **Generate the Default `.spec` File:**
    * Open your terminal (with your `.venv` activated) and navigate to your project folder.
    * Run the following command once. This creates a file named `email_app.spec` in your project folder.
    ```bash
    pyi-makespec --onefile --windowed --icon="icon.ico" email_app.py
    ```

2.  **Edit the `email_app.spec` File:**
    * Open the newly created `email_app.spec` file in a text editor (like VS Code or Notepad).
    * You will see a section that looks like this:
        ```python
        a = Analysis(
            ['email_app.py'],
            pathex=[],
            binaries=[],
            datas=[],
            hiddenimports=[],
            hookspath=[],
            ...
        )
        ```
    * Find the line `hiddenimports=[]` and **add `'pythonjsonlogger'`** to the list, like this:
        ```python
        hiddenimports=['pythonjsonlogger'],
        ```
    * Save and close the file.

### **Step 2: Build the `.exe` from the `.spec` File**

Now, we tell PyInstaller to build the application using our new configuration file.

1.  **Run PyInstaller with the `.spec` file:**
    * In the same activated terminal, run this command:
    ```bash
    pyinstaller email_app.spec
    ```
2.  **Find Your Standalone `.exe`:**
    * When the process is finished, a new `email_app.exe` will be in your `dist` folder. This version is now guaranteed to contain the `python-json-logger` library.

### **Step 3: Re-compile the Installer with Inno Setup**

This step is the same as before.

1.  **Open Inno Setup Compiler** (run as administrator if needed).
2.  Go to **File > Open** and select your `setup_script.iss` file.
3.  **Update the version number** in the script to `v1.3` (or your current version) to avoid confusion.
4.  Go to **Build > Compile**.

The new `EmailAutomationTool_Setup_v1.3.exe` in your `Installer` folder will now be correct and will not have the `ModuleNotFoundError` when run on a new machine.
