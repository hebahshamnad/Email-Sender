# Email-Sender
Python Script for Sending Mass Scheduled Emails via Microsoft Office 365 


This Python script demonstrates how to send emails using the exchangelib library and Microsoft Office 365. The script utilizes the exchangelib library to establish a connection with the Exchange server, authenticates the sender's credentials, and sends an email in HTML format. With support for HTML templates, you can easily customize the email content and deliver rich, dynamic messages. By reading recipient information from an Excel file, you can efficiently manage and organize your email campaigns. Whether you need to send regular updates, notifications, or newsletters, this script automates the process, saving you time and effort. The script also includes error handling to handle potential exceptions during the email sending process. This code can serve as a starting point for integrating email functionality into Python applications that utilize Microsoft Office 365 as the email service provider.


Key Features:

* Schedule mass emails to be sent at specific times or intervals
* Read recipient information from an Excel file for personalized email delivery
* Customize email content using HTML templates
* Easy integration with Microsoft Office 365/Outlook accounts
* Handles authentication and connectivity to the email server

Note: It is important to configure the schedule and mailing list according to your requirements before running the script.


Limitations:

* Security Considerations: When using this script, it is essential to consider security aspects. Make sure to keep the sender's email and password confidential and ensure that you are using a secure and trusted server environment to run the script.

* Daily Email Sending Limit: Most email service providers, including Microsoft Office 365/Outlook, impose a limit on the number of emails that can be sent per day. Ensure that you are aware of the daily sending limit set by your email provider and configure the script accordingly to avoid any disruptions.

* Compatibility: This script is designed specifically for use with Microsoft Office 365/Outlook email accounts. It may not be compatible with other email service providers or protocols. Ensure that you have a valid Microsoft 365/Outlook email account and the necessary credentials before running the script.

* Provider Restrictions: Different email service providers may have specific policies, restrictions, and guidelines related to sending bulk or automated emails. Familiarize yourself with your provider's terms of service and any restrictions they impose.For example, some providers may require you to use specific APIs, SMTP servers, or obtain special permissions for sending bulk emails.

* File and Data Integrity: The script assumes that the Excel file containing recipient information ("email.xlsx") is present in the same directory as the script file. Ensure that the file is correctly formatted with the required columns ("Name", "Email", "Message") and that the data is accurate and up-to-date.


