Software Requirements Specification: Email Automation Tool v2.9
1. Introduction
This document outlines the requirements for a desktop application designed to automate and streamline the process of sending batch emails using Microsoft Outlook. The application provides a user-friendly, modern interface for bulk-managing recipient lists and sending perfectly formatted emails by using an Outlook draft as a template. The intended users are employees who need to send standardized, richly formatted communications to multiple recipients efficiently and with fewer errors.

2. Overall Description
The application is a standalone Python-based tool with a GUI styled to mimic the modern Microsoft Outlook aesthetic. It allows users to paste lists of recipient and CC email addresses in bulk and save them automatically.

The core functionality revolves around a "draft template" workflow. The user first composes a perfectly formatted email in their Outlook Drafts folder. The application uses the subject line of this draft to find it and send a copy to every valid recipient on the list. The application includes features to reduce errors, such as a pre-send review of the recipient lists and a confirmation dialog to verify CC addresses. A default CC address is also included as a placeholder. All major operations and errors are recorded in a dual-format system log for future reference and database integration.

3. Functional Requirements
3.1 User Interface (UI)
Main Window: A single main window with a clean, modern layout.

Receiver's Email Addresses ("To"): A single, large multi-line text area for pasting a list of recipient emails.

CC Email Addresses ("Cc"): A single, large multi-line text area for pasting a list of CC emails. This field will be pre-populated with mro@deci-ltd.com as a default.

Outlook Draft Subject (Template): A dedicated input field for the user to enter the exact subject line of the email template.

Sender Information Display: A non-editable text label showing the default sender's email address detected from Outlook.

Buttons:

"Review 'To' List" / "Review 'Cc' List": Buttons that validate and display a clean, final recipient list in a pop-up window.

"Send in Batch": The primary button to initiate the email sending process.

"Cancel": A button to stop an in-progress sending operation.

"Clear": A button that, after confirmation, erases all content from the input fields.

3.2 Core Functionality
Bulk Email Input and Validation:

The application shall accept newline-separated lists of emails in both the "To" and "Cc" text areas.

The "Review" functions shall extract, validate, and de-duplicate emails from the input text.

Email Sending via Draft Template:

The application integrates with the user's locally installed Microsoft Outlook client.

Before sending, a confirmation dialog shall display the final parsed list of CC'd email addresses and require user confirmation.

Upon confirmation, the application will find the draft in Outlook by its subject and send a copy to each recipient. This copy includes all original formatting and attachments.

Data Persistence:

All user-entered text will be automatically saved to a local email_data.json file, stored in the user's AppData folder.

System Logging:

The application shall generate and maintain two log files in the user's AppData folder:

app_log.log: A human-readable, plain-text log for easy debugging.

app_log.jsonl: A structured JSON log where each line is a JSON object, designed for automated parsing and database ingestion.

The logs shall record key events, including application start/shutdown, data loading, batch send initiation, individual emails sent, and all errors.

Both log files will automatically rotate to prevent excessive size.

4. Non-Functional Requirements
Platform: Developed in Python and compatible with Windows.

Usability & Design: An intuitive, modern UI with error-prevention features.

Performance: The application must remain responsive during operations.

Reliability: The application must handle potential errors gracefully and provide clear user feedback.

5. Assumptions and Dependencies
The user must have Microsoft Outlook (Classic version) installed and configured.

An active internet connection is required for Outlook to send emails.

The user is responsible for creating and saving the email template (including any attachments) in their Outlook Drafts folder before use.