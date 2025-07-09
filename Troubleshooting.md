### **FAQ & Troubleshooting Guide**

This guide helps solve common issues when running the Python Email Automation Tool on a new computer, especially errors like `ModuleNotFoundError: No module named 'win32com'` or the app failing to detect your Outlook account.

---

#### **Primary Issue: The app fails to start or can't find Outlook.**

When you encounter errors, follow these steps in order. **Step 1 solves the problem over 90% of the time.**

---

### **Step 1: Check Your Virtual Environment (Most Common Fix)**

Modern Python projects use virtual environments to keep their required packages separate from the main system. If the environment isn't active, the script can't find the packages it needs.

**Symptom:** You open a new terminal, run `python email_app.py`, and immediately get `ModuleNotFoundError`.

**Solution:** You must create and activate the virtual environment for your project folder.

1.  **Open your terminal** and navigate to your project folder:
    ```bash
    cd C:\Users\ElsaWang\Documents\DECI-EmailAutomationTool
    ```

2.  **Create a new virtual environment.** (You only need to do this once per machine).
    ```bash
    python -m venv .venv
    ```

3.  **Activate the environment.** (You must do this every time you open a new terminal for this project).
    ```bash
    .\.venv\Scripts\activate
    ```
    Your terminal prompt should now start with `(.venv)`.

4.  **Install the required packages *inside* the active environment.**
    ```bash
    python -m pip install pywin32
    ```

5.  **Run the app.**
    ```bash
    python email_app.py
    ```

---

### **Step 2: Check for a "Bitness" Mismatch**

For the script to work, your Python installation and your Microsoft Office/Outlook installation **must both be 64-bit or both be 32-bit.**

**Symptom:** The virtual environment is active and `pywin32` is installed, but the app still reports "Outlook not running or account not found."

**Solution:**

1.  **Check Python Bitness:** In your terminal, type `python`. The first line will say `64 bit (AMD64)` or `32 bit`.
2.  **Check Outlook Bitness:** In Outlook, go to **File > Office Account > About Outlook**. The top line of the pop-up window will state if it is 32-bit or 64-bit.

If they do not match, the easiest solution is to uninstall and reinstall Python with the version that matches your Office installation.

---

### **Step 3: Run a "Quick Repair" on Microsoft Office**

Sometimes, the communication components that allow Python to talk to Outlook can be improperly registered in Windows.

**Symptom:** Bitness matches and the virtual environment is active, but the connection to Outlook still fails.

**Solution:**

1.  Close Outlook and all other Office apps.
2.  Go to **Windows Settings > Apps > Installed apps**.
3.  Find your **Microsoft 356** or **Microsoft Office** installation.
4.  Click **Modify**.
5.  Select **Quick Repair** and let the process run.

---

### **Step 4: Check the Terminal in VS Code**

The terminal inside VS Code often defaults to **PowerShell**, which has stricter security policies than the standard **Command Prompt (CMD)**. These policies can block the script from communicating with Outlook.

**Symptom:** The script runs perfectly in an external Command Prompt window, but fails to detect Outlook when run from the terminal inside VS Code.

**Solution:** Change the default terminal profile in VS Code to Command Prompt.

1.  In VS Code, press **`Ctrl + Shift + P`** to open the Command Palette.
2.  Type **`Terminal: Select Default Profile`** and select it from the list.
3.  Choose **Command Prompt** from the options.
4.  Close your current terminal panel (click the trash can icon) and open a new one. It will now be a Command Prompt, and your script will run correctly.

---

### **Step 5: Check if You're Using "New" Outlook**

The script can only automate the **Classic** desktop version of Outlook. It cannot work with the **"New" Outlook**, which is a web-based application.

**Symptom:** You have tried all the above steps, and it still fails.

**Solution:**

* Ensure the version of Outlook you are running is the traditional desktop client. If you are using the "New" Outlook (which often has a "New" toggle switch at the top), you must switch back to the classic version for this tool to function.

---

### **General Question: Why do we need a virtual environment (.venv)?**

Think of a virtual environment as a **separate, isolated toolbox for each of your Python projects.** It's a standard best practice for a few key reasons:

* **Dependency Isolation:** Imagine Project A needs an old version of a library, but Project B needs a new version. Without a `.venv`, you can only have one version installed on your computer, which would break one of your projects. A virtual environment gives each project its own private set of tools, so they never conflict.

* **Reproducibility:** A virtual environment makes it easy to create a `requirements.txt` file, which is a perfect list of all the packages your project needs. When you move the project to a new computer (like we did), you can create a new `.venv` and install everything from that list in one command, guaranteeing the project will have the exact setup it needs to run.

* **Keeps Your System Clean:** It prevents your main computer's Python installation from being cluttered with hundreds of packages from dozens of different projects. Each project's dependencies stay neatly inside its own folder.

In our case, the problem was that `pywin32` was installed in the main system "toolbox," but our script was looking for it inside its own empty project "toolbox" (`.venv`). Once we created the `.venv` and installed the package inside it, the script worked perfectly.
