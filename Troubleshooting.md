### **FAQ & Troubleshooting Guide**

This guide helps solve common issues when running the Python Email Automation Tool, from initial setup to sending emails.

---

### **Common Usage Error: "Template Not Found"**

This is the most common error during normal use. It means the application could not find an email in your Outlook Drafts folder with the subject you provided.

**Solution:** Check the following in your Outlook:

1.  **Exact Subject Match:** The subject in the app must be an *exact* match to the subject of the draft email, including capitalization and spacing. Copy and paste it to be sure.
2.  **Location:** The email template **must** be saved in the main **Drafts** folder for your default email account.
3.  **Is it still a draft?** Make sure you haven't accidentally sent or deleted the draft email.

---

### **Initial Setup Issue: App fails to start or can't find Outlook.**

When setting up the app on a new computer, follow these steps in order if you encounter errors.

---

#### **Step 1: Check Your Virtual Environment (Most Common Setup Fix)**

**Symptom:** You open a terminal, run `python email_app.py`, and immediately get a `ModuleNotFoundError`.

**Solution:** You must create and activate the virtual environment for your project folder.

1.  **Open your terminal** and navigate to your project folder (e.g., `cd D:\DECI-EmailAutomationTool`).
2.  **Create a new virtual environment** (only needs to be done once per machine):
    ```bash
    python -m venv .venv
    ```
3.  **Activate the environment** (must be done every time you open a new terminal for this project):
    ```bash
    .\.venv\Scripts\activate
    ```
    Your terminal prompt should now start with `(.venv)`.
4.  **Install required packages** inside the active environment:
    ```bash
    python -m pip install pywin32
    ```
5.  **Run the app:**
    ```bash
    python email_app.py
    ```

---

#### **Step 2: Check for a "Bitness" Mismatch**

**Symptom:** The virtual environment is active and packages are installed, but the app still reports "Outlook not running or account not found."

**Solution:** Your Python and Microsoft Office installations must **both be 64-bit or both be 32-bit.**

1.  **Check Python Bitness:** In your terminal, type `python`. The first line will say `64 bit (AMD64)` or `32 bit`.
2.  **Check Outlook Bitness:** In Outlook, go to **File > Office Account > About Outlook**. The top line will state if it is 32-bit or 64-bit.

If they don't match, the easiest solution is to reinstall Python with the version that matches your Office installation.

---

#### **Step 3: Run a "Quick Repair" on Microsoft Office**

**Symptom:** Bitness matches and the environment is active, but the connection to Outlook still fails.

**Solution:** This re-registers the communication components between apps.

1.  Close all Office apps.
2.  Go to **Windows Settings > Apps > Installed apps**.
3.  Find your **Microsoft 365** or **Microsoft Office** installation, click it, and choose **Modify**.
4.  Select **Quick Repair** and let the process run.

---

#### **Step 4: Check the Terminal in VS Code**

**Symptom:** The script runs perfectly in an external Command Prompt, but fails to detect Outlook when run from the terminal inside VS Code.

**Solution:** Change VS Code's default terminal to Command Prompt. The default terminal (PowerShell) has stricter security policies that can block the script from communicating with Outlook.

1.  In VS Code, press **`Ctrl + Shift + P`** to open the Command Palette.
2.  Type and select **`Terminal: Select Default Profile`**.
3.  Choose **Command Prompt**.
4.  Close and reopen the terminal panel in VS Code.

**Note:** This is a development environment issue only. The final `.exe` created by the installer runs directly on Windows and is **not affected** by terminal security policies.

---

#### **Step 5: Check if You're Using "New" Outlook**

**Symptom:** You have tried all the above steps, and it still fails.

**Solution:** The script can only automate the **Classic** desktop version of Outlook. It cannot work with the **"New" Outlook**. Ensure you are running the traditional desktop client.
