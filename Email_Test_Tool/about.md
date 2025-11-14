This PowerShell script is a graphical user interface (GUI) application for sending emails using the MailKit library, which is a robust mail-handling library for .NET. The script provides a modern, clean interface for sending SMTP emails and includes features like field validation, progress reporting, and log display. Here's a breakdown of its key functionalities and what each part does:

Key Features:

1. GUI-Based Email Sending:

It uses Windows Presentation Foundation (WPF) for a modern interface, built using PowerShell.

It includes fields for configuring the SMTP server, sender/recipient details, subject, body, and user credentials.

2. Supports TLS/SSL:

The script allows users to send emails securely using TLS/SSL encryption.

There's a checkbox to enable SSL/TLS connections.

3. Email Validation & Progress Reporting:

The script checks for empty or missing required fields (such as email addresses and password).

It includes a progress bar that tracks the email sending process.

4. Status Logging:

Real-time logs are displayed at the bottom of the window.

Log entries are color-coded based on the status of the email operation (e.g., success, error, warning).

The logs are scrollable and show time-stamped entries.

5. Dynamic Log Viewer:

The log section dynamically updates as new messages are added.

Log entries are color-coded with icons to help users easily identify the status (e.g., success, failure, warnings).

6. Automatic MailKit Module Installation:

If the Send-MailKitMessage module is not already installed, the script automatically installs it for the user.

It also logs the installation process in the status log.

7. SMTP Provider Pre-Fill:

The script includes pre-filled SMTP servers for common providers (Office365, Gmail, Outlook), but it also allows users to input their own custom SMTP server if needed.

Major Sections and Components:

8. GUI Layout (XAML):

Header: Displays the app name with a mail icon.

Main Form: Includes form fields for SMTP server, port, sender/recipient email addresses, subject, message body, and credentials.

SSL Checkbox: Lets the user choose whether to use a secure connection.

Progress Bar: Shows the progress of the email sending process.

Status Log: Displays real-time logs (color-coded based on the message type: info, success, error, warning).

9. Helper Functions:

LogStatus: Adds a log entry to the status window, color-coding the message depending on its type (info, error, success, etc.).

Ensure-MailKitModule: Checks if the Send-MailKitMessage module is installed. If not, it installs it and imports it. Logs success or failure to the status log.

10. Event Handlers:

SendButton: The button to send the email. It collects input values from the form, validates them, and attempts to send the email using the Send-MailKitMessage cmdlet.

ClearLogButton: Clears the status log window.

ExitButton: Closes the application window.

11 What the Script Does:

When you press the "Send Email" button:

The script ensures the MailKit module is installed.

It checks if all required fields are filled (SMTP server, sender and recipient email addresses, username, password).

If all fields are valid, it prepares the email content (SMTP settings, recipient list, subject, body) and sends the email using the MailKit library.

It updates the progress bar as the email is being sent and provides feedback in the status log.

If successful, it shows a success message, and if there is an error, it logs the error and shows an error message.

When you press the "Clear Log" button:

Clears the status log.

When you press the "Exit" button:

Closes the GUI application.
