FAQ & Troubleshooting Guide
This guide helps solve common issues when running the Email Automation Tool, from initial setup to sending emails. It is organized into sections for end-users of the installed application and for developers working with the source code.

Section 1: For End-Users (After Installing the .exe)
Issue: "File contains a virus or potentially unwanted software"
Symptom: When you run the downloaded setup.exe, Windows Defender blocks it.

Cause: This is a false positive. Because the installer is from an individual developer and not "code-signed" (an expensive verification process), security software is extra cautious. The file is safe.

Solution: Add an exclusion for the installer file in Windows Security.

Open Windows Security > Virus & threat protection.

Click Manage settings > Add or remove exclusions.

Click + Add an exclusion, select File, and choose the downloaded setup file.

Issue: App crashes on launch with ModuleNotFoundError: pythonjsonlogger
Symptซom: You successfully install the app, but it crashes immediately upon launching.

Cause: This was a build error in versions before v1.3 where a required library was not included in the installer.

Solution: This has been permanently fixed. Ensure you have downloaded and installed the latest version (v1.3 or newer) from the GitHub Releases page.

Issue: App requires "Run as administrator" and then can't find Outlook
Symptom: The app shows a "permission denied" error when run normally. When you "Run as administrator," it opens but then fails to detect your Outlook account.

Cause: This occurred with older versions (before v1.2) that incorrectly tried to save data to a protected system folder. This created a privilege mismatch between the app and Outlook.

Solution: This has been permanently fixed. The latest version correctly saves data to the proper user folder (AppData) and does not require administrator rights to run.

Issue: "Template Not Found" Error
Symptom: The application shows an error saying it could not find the draft with the specified subject.

Solution:

Exact Subject Match: The subject in the app must be an exact match to the draft's subject, including capitalization and any hidden spaces. Copy and paste the subject from Outlook into the app to be certain.

Location: The template must be in the Drafts folder of your default Outlook account (the one shown at the bottom of the app).

Is it still a draft? Make sure you haven't accidentally sent or deleted the draft email.

Issue: "This method can't be used with an inline response mail item" Error
Symptom: The application shows this error when trying to send emails.

Cause: This happens if the draft email was last viewed in Outlook's main reading pane instead of a separate window, which puts it in a read-only state.

Solution: This bug has been fixed in version v1.3 and newer. Please ensure you have the latest version.

Section 2: For Developers (Working with the Source Code)
Issue: ModuleNotFoundError when running python email_app.py
Symptom: You open a new terminal, activate the environment, but still get ModuleNotFoundError: No module named 'win32com' or 'pythonjsonlogger'.

Cause: The required libraries are not installed inside the virtual environment.

Solution: You must activate the virtual environment and then install the packages.

Navigate to the project folder: cd D:\DECI-EmailAutomationTool

Activate the environment: .\.venv\Scripts\activate

Install dependencies: python -m pip install pywin32 python-json-logger

Run the app: python email_app.py

Issue: Script works in Command Prompt but not in the VS Code Terminal
Symptom: The script runs perfectly in an external Command Prompt window, but fails to detect Outlook when run from the terminal inside VS Code.

Cause: The default terminal in VS Code is often PowerShell, which has stricter security policies that can block the script from communicating with Outlook.

Solution: Change VS Code's default terminal to Command Prompt.

In VS Code, press Ctrl + Shift + P to open the Command Palette.

Type and select Terminal: Select Default Profile.

Choose Command Prompt.

Close and reopen the terminal panel in VS Code.

Issue: "Bitness" Mismatch
Symptom: The environment is active and packages are installed, but the app still reports "Outlook not running or account not found."

Solution: Your Python and Microsoft Office installations must both be 64-bit or both be 32-bit.

Check Python Bitness: In your terminal, type python. The first line will say 64 bit (AMD64) or 32 bit.

Check Outlook Bitness: In Outlook, go to File > Office Account > About Outlook. The top line will state if it is 32-bit or 64-bit.

Issue: Connection to Outlook still fails after all other steps
Symptom: You've checked everything else, but the connection to Outlook is unreliable.

Solution: Run a Quick Repair on your Microsoft Office installation. This re-registers the communication components between apps.

Close all Office apps.

Go to Windows Settings > Apps > Installed apps.

Find your Microsoft 365 or Microsoft Office installation, click it, and choose Modify.

Select Quick Repair and let the process run.