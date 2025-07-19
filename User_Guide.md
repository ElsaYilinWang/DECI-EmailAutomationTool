Email Automation Tool - User Manual & Troubleshooting Guide
This document provides a complete guide for using the Email Automation Tool. It is divided into two sections: a User Manual for getting started and a detailed Troubleshooting Guide for solving common problems.

Section 1: User Manual for Beginners
This section explains the standard workflow for sending a batch of emails.

Step 1: Prepare Your Email Template in Outlook
The application works by sending a perfect copy of an email you have saved in your Outlook Drafts folder.

Open Microsoft Outlook (Classic version).

Click "New Email".

Compose your email exactly as you want it to appear.

Formatting: Copy and paste tables from Excel, add images, and use different fonts. All formatting will be preserved.

Attachments: Attach any files you want to be included in every email.

Signature: Your default Outlook signature will be automatically included.

Write a clear and unique Subject line. You will need this exact subject for the app.

Instead of sending, go to File > Save. This will save the email in your Drafts folder.

Step 2: Use the Email Automation Tool
Launch the Application: Open the Email Automation Tool from your desktop or Start Menu shortcut.

Paste Recipient Emails:

Copy your list of recipient emails.

Paste the entire list into the large "To (Paste Email List)" box on the left. Each email should be on a new line.

Paste CC Emails (Optional):

The "Cc (Paste Email List)" box on the right will default to mro@deci-ltd.com.

You can add more emails below it on new lines, or delete it if no CC is needed.

Enter the Draft Subject:

In the "Subject of Draft Template in Outlook" field, type or paste the exact subject line from the draft you saved in Step 1. This must be a perfect match.

(Recommended) Review Your Lists:

Click the "Review 'To' List" and "Review 'Cc' List" buttons to see a clean, final list of recipients and make sure everything is correct.

Step 3: Send the Emails
Click "Send in Batch": A confirmation box will appear, showing you the final list of CC'd emails.

Confirm: Click "OK" to confirm. The application will now start sending a copy of your draft to every recipient in the "To" list.

A final "Success" message will appear when the process is complete.

Section 2: Troubleshooting Guide
This section helps solve common errors and setup issues.

Installation & Usage Errors
Issue: "File contains a virus or potentially unwanted software"
Symptom: When you run the downloaded setup.exe, Windows Defender blocks it.
Cause: This is a false positive. Because the installer is from an individual developer and not "code-signed" (an expensive process), security software is extra cautious.
Solution: Add an exclusion for the installer file in Windows Security.

Open Windows Security > Virus & threat protection.

Click Manage settings > Add or remove exclusions.

Click + Add an exclusion, select File, and choose the downloaded setup file.

Issue: "Template Not Found" Error
Symptom: The application shows an error saying it could not find the draft with the specified subject.
Solution:

Exact Subject Match: The subject in the app must be an exact match to the draft's subject, including capitalization and any hidden spaces. Copy and paste the subject from Outlook into the app to be certain.

Location & Multiple Accounts: The template must be in the Drafts folder of your default Outlook account. If you have multiple accounts (e.g., work and personal), the app will only check the Drafts folder of the account shown at the bottom of the app ("Sending from: ..."). Make sure your draft is saved in the correct account's folder.

Issue: App cannot detect Outlook account on a new device
Symptom: The "Sending from:" label shows "Outlook not running or no account found," even though Outlook is open.
Solution: This is almost always an environment issue on the new computer. Follow these steps in order:

Check for "New" Outlook: The app only works with Classic Outlook. If you see a "New Outlook" toggle in the top-right of Outlook, turn it Off.

Check for 64-bit Mismatch: The app is 64-bit, so your Office/Outlook must also be 64-bit. Check in Outlook via File > Office Account > About Outlook.

Run an Online Repair: This is a more thorough fix than "Quick Repair." Go to Windows Settings > Apps, find your Microsoft 365 installation, click Modify, and select Online Repair.

Create a New Outlook Profile: If the repair fails, your Outlook profile may be corrupt. Go to Control Panel > Mail (Microsoft Outlook) > Show Profiles... > Add... to create a new profile and set it as the default. This does not delete your emails.

Issue: ModuleNotFoundError or "Permission Denied" after installing
Cause: These were bugs in older versions of the installer (before v1.7).
Solution: Ensure you have downloaded and installed the latest version from the GitHub Releases page, which has permanently fixed these issues.

Section 3: For Developers (Working with the Source Code)
Issue: ModuleNotFoundError when running python email_app.py
Solution: Activate the project's virtual environment and install dependencies first.

Navigate to the project folder: cd D:\DECI-EmailAutomationTool

Activate the environment: .\.venv\Scripts\activate

Install dependencies: python -m pip install pywin32 python-json-logger

Run the app: python email_app.py

Issue: Script works in Command Prompt but not in the VS Code Terminal
Cause: The default terminal in VS Code is often PowerShell, which has stricter security policies.
Solution: Change VS Code's default terminal to Command Prompt.

In VS Code, press Ctrl + Shift + P.

Type and select Terminal: Select Default Profile.

Choose Command Prompt.