Software Requirements Specification: Email Automation Tool v2.6
1. Introduction
This document outlines the requirements for a desktop application designed to automate and streamline the process of sending batch emails using Microsoft Outlook. The application provides a user-friendly, modern interface for bulk-managing recipient lists and sending perfectly formatted emails by using an Outlook draft as a template. The intended users are employees who need to send standardized, richly formatted communications to multiple recipients efficiently and with fewer errors.

2. Overall Description
The application is a standalone Python-based tool with a GUI styled to mimic the modern Microsoft Outlook aesthetic. It allows users to paste lists of recipient and CC email addresses in bulk and save them automatically.

The core functionality revolves around a "draft template" workflow. The user first composes a perfectly formatted email in their Outlook Drafts folder. The application uses the subject line of this draft to find it and send a copy to every valid recipient on the list. The application includes features to reduce errors, such as a pre-send review of the recipient lists and a confirmation dialog to verify CC addresses. A default CC address is also included as a placeholder to prevent accidental omissions.

3. Functional Requirements
3.1 User Interface (UI)
Main Window: A single main window with a clean, modern layout.

Receiver's Email Addresses ("To"):

This section features a single, large multi-line text area for pasting a list of recipient emails, with each email on a new line.

CC Email Addresses ("Cc"):

This section features a single, large multi-line text area for pasting a list of CC emails.

This field will be pre-populated with mro@deci-ltd.com as a default if no other CC addresses are loaded from memory.

Outlook Draft Subject (Template):

A dedicated input field for the user to enter the exact subject line of the email template saved in their Outlook Drafts folder.

Sender Information Display:

A non-editable text label at the bottom of the window, showing the default sender's email address detected from Outlook.

Buttons:

"Review 'To' List" / "Review 'Cc' List": Buttons that will parse the text in the corresponding field, validate the email formats, remove duplicates, and display a clean, final recipient list in a pop-up window for user verification.

"Send in Batch": The primary button to initiate the email sending process.

"Cancel": A button to stop an in-progress sending operation.

"Clear": A button that, after confirmation, erases all content from the input fields.

3.2 Core Functionality
Bulk Email Input and Validation:

The application shall accept a newline-separated list of emails pasted into both the "To" and "Cc" text areas.

The "Review" functions shall extract, validate, and de-duplicate emails from the input text, displaying the final list and count to the user.

Email Sending via Draft Template:

The application integrates with the user's locally installed Microsoft Outlook client.

The sender's email address is the default account configured in Outlook.

Before sending, the application will display a confirmation dialog showing the user the final parsed list of CC'd email addresses and require confirmation to proceed.

Upon confirmation, the application will:

Search the user's Outlook Drafts folder for an email with a subject that exactly matches the text in the "Outlook Draft Subject" field.

If no matching draft is found, display an error message.

If a draft is found, copy its entire formatted content (HTML body) and subject.

For each valid recipient in the "To" list, create and send a new email using the copied subject and formatted body.

Data Persistence:

The text content from the "To" and "Cc" fields, as well as the 'Outlook Draft Subject', will be automatically saved to a local configuration file.

When the application is launched, it will automatically load this saved information.

4. Non-Functional Requirements
Platform: Developed in Python and compatible with Windows.

Usability & Design: An intuitive, modern UI similar to Microsoft Outlook. Error-prevention features like confirmation dialogs and default values shall be included.

Performance: The application must remain responsive while sending emails.

Reliability: The application must handle potential errors gracefully (e.g., Outlook not running, template not found, invalid email formats) and provide clear user feedback.

5. Assumptions and Dependencies
The user must have Microsoft Outlook (Classic version) installed and configured.

An active internet connection is required for Outlook to send emails.

The user is responsible for creating and saving the email template in their Outlook Drafts folder before use.