# EMAIL SENDER 
# 13 JUNE 2023
# H.S

import pandas as pd
import schedule
import time
from exchangelib import DELEGATE, Account, Message, Mailbox, HTMLBody, Credentials
from stdiomask import getpass
import pytz


def send_email(sender_email, sender_password, recipient_email, recipient_name, subject, message):
    try:
        # Create the credentials object
        credentials = Credentials(username=sender_email, password=sender_password)

        # Connect to the Microsoft Exchange server
        account = Account(
            primary_smtp_address=sender_email,
            credentials=credentials,
            autodiscover=True,
            access_type=DELEGATE
        )

        # Create the email message
        email = Message(
            account=account,
            folder=account.sent,
            subject=subject,
            body=HTMLBody(message),
            to_recipients=[Mailbox(email_address=recipient_email, name=recipient_name)]
        )

        # Send the email
        email.send()

        print("Email sent successfully!")
    except Exception as e:
        print("Failed to send email. Error:", str(e))


def send_scheduled_emails(sender_email, sender_password):
    # Read the Excel file
    df = pd.read_excel('email_list.xlsx')

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        recipient_name = row['Name']
        recipient_email = row['Email']
        message_body = row['Message']
        subject = 'Bot Test Email'

        message = f"""
        <html>
        <head>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                }}

                h1 {{
                    color: #007bff;
                }}

                p {{
                    line-height: 1.5;
                }}
            </style>
        </head>
        <body>
            <h1>Hello, {recipient_name}!</h1>
            <h2>This is a test email sent using Python through Microsoft Office 365.</h2>
            <p>Feel free to use this message to send reminders, updates, newsletters etc of your choice </p>
            
              # If you feel that you would like to send a different       message to each individual in the excel file. Make use of the 'Message column, enter your message for each person and then use the below code exerpt
           # <p> {message_body} </p>
           
        </body>
        </html>
        """

        # Send the email
        send_email(sender_email, sender_password, recipient_email, recipient_name, subject, message)


# Email configuration: Please note that the sender email has to be a Microsoft 365/Outlook email
sender_email = input('Enter your email: ')

sender_password = getpass('Enter sender password: ', mask='*')

# Set the timezone to Eastern Standard Time (EST)
est_tz = pytz.timezone('US/Eastern')

# Schedules the email to send everyday at a specified time
# schedule.every().day.at("09:30").do(send_scheduled_emails, sender_email, sender_password).\
#     timezone(est_tz)

#  Uncomment to schedule the email to send every few minutes specified
schedule.every(60).seconds.do(send_scheduled_emails, sender_email, sender_password)


while True:
    schedule.run_pending()
    time.sleep(1)
