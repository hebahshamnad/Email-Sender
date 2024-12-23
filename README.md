# Email-Sender
Python Script for Sending Mass Scheduled Emails via Microsoft Office 365

This script demonstrates how to send emails using the `exchangelib` library with Microsoft Office 365. It connects to the Exchange server, authenticates the sender, and sends HTML formatted emails. You can read recipient information from an Excel file, allowing for efficient email campaign management.

## Key Features:
- Schedule mass emails for specific times or intervals.
- Read recipient data from an Excel file for personalized delivery.
- Customize email content using HTML templates.
- Easy integration with Microsoft Office 365/Outlook accounts.
- Handles authentication and connectivity to the email server.

**Note:** Configure the schedule and mailing list according to your needs before running the script.

## Limitations:
- **Security:** Keep the sender's email and password confidential. Use a secure server environment.
- **Daily Sending Limit:** Be aware of your email provider's daily sending limits to avoid disruptions.
- **Compatibility:** Designed for Microsoft Office 365/Outlook accounts; may not work with other providers.
- **Provider Restrictions:** Familiarize yourself with your provider's policies on bulk or automated emails.
- **File Integrity:** Ensure the Excel file ("email.xlsx") is correctly formatted with the required columns ("Name", "Email", "Message").


