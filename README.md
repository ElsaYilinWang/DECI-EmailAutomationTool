Software Requirements Specification: Email Automation Tool v2.4
1. Introduction
This document outlines the requirements for a desktop application designed to automate and streamline the process of sending batch emails using Microsoft Outlook. The application provides a user-friendly, modern interface for managing recipient lists and sending perfectly formatted emails by using an Outlook draft as a template.

2. Overall Description
The application is a standalone Python-based tool with a graphical user interface (GUI) styled to mimic the modern Microsoft Outlook aesthetic. It allows users to input and automatically save lists of 'To' and 'CC' email addresses.

The core functionality revolves around a "draft template" workflow. The user first composes a perfectly formatted email (including tables, images, and signatures) and saves it in their Outlook Drafts folder. The application then uses the subject line of this draft to find it and send a copy to every recipient on the list. This method leverages the full power of Outlook's editor while providing a simple, powerful interface for batch sending. The application interfaces directly with the user's installed Microsoft Outlook client, using the default signed-in account as the sender.

3. Functional Requirements
3.1 User Interface (UI)
Main Window: The application features a single main window with a clean, modern layout inspired by the new Microsoft Outlook.

Receiver's Email Addresses ("To"):

This section contains 24 input fields within a scrollable pane to efficiently manage space.

CC Email Addresses ("Cc"):

This section contains 6 input fields.

Outlook Draft Subject (Template):

A dedicated input field is provided for the user to enter the exact subject line of the email template saved in their Outlook Drafts folder.

Sender Information Display:

A non-editable text label is displayed at the bottom of the window, showing the default sender's email address as detected from the user's Microsoft Outlook.

Buttons:

"Send in Batch": A primary-styled button that initiates the email sending process.

"Cancel": A secondary-styled button to immediately stop an in-progress sending operation.

"Clear": A secondary-styled button that, after a confirmation prompt, erases all content from the 'To', 'Cc', and 'Draft Template Subject' fields.

3.2 Core Functionality
Email Sending via Draft Template:

The application integrates with the user's locally installed Microsoft Outlook client.

The sender's email address is the default account configured in the user's Outlook.

Upon clicking "Send in Batch", the application will:

Search the user's Outlook Drafts folder for an email with a subject line that exactly matches the text in the "Outlook Draft Subject" field.

If no matching draft is found, display an error message to the user.

If a draft is found, copy its entire formatted content (HTML body) and subject.

For each recipient in the 'To' list, create and send a new email using the copied subject and formatted body.

Sender Email Detection:

On startup, the application will automatically detect and display the email address of the default account in the user's Outlook profile.

Data Persistence:

All email addresses ('To' and 'Cc') and the 'Outlook Draft Subject' will be automatically saved to a local configuration file (email_data.json).

When the application is launched, it will automatically load this saved information into the appropriate fields.

4. Non-Functional Requirements
Platform: The application is developed in Python and is compatible with the Windows operating system.

Usability & Design:

The UI is designed to be intuitive, modern, and professional, with a look and feel similar to the new Microsoft Outlook.

Performance: The application must remain responsive while sending emails in the background.

Reliability: The application must handle potential errors gracefully (e.g., Outlook not running, template not found) and provide clear feedback to the user via message boxes.

5. Assumptions and Dependencies
The user must have Microsoft Outlook installed and configured on their Windows machine.

The application requires an active internet connection for Outlook to send emails.

The user's Outlook application must have a default email account set up.

The user is responsible for creating and saving the email template in their Outlook Drafts folder before using the application to send.