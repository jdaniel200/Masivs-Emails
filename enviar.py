import os
import base64
import time
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Authenticate and build Gmail service
def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

# Function to create an email with an attachment
def create_message_with_attachment(to, subject, body, attachment_path):
    message = MIMEMultipart()
    message['to'] = to
    message['subject'] = subject

    # Attach the email body
    message.attach(MIMEText(body, 'plain'))

    # Attach the file
    if attachment_path:
        filename = os.path.basename(attachment_path)
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {filename}')
        message.attach(part)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw_message}

# Function to send an email
def send_email(service, to, subject, body, attachment_path=None):
    try:
        message = create_message_with_attachment(to, subject, body, attachment_path)
        sent_message = (service.users().messages().send(userId="me", body=message).execute())
        print(f"Message sent to {to}: {sent_message['id']}")
    except HttpError as error:
        print(f"An error occurred: {error}")

# Main function to send bulk emails
def main():
    service = authenticate_gmail()

    # Read recipients and attachments from CSV file
    with open('correos.csv', mode='r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            recipient = row['correo']
            attachment_path = "archivos/" + row['archivo']
            subject = "Archivos a revisar"
            body = "Hola, por favor revisa estos archivos adjuntos."
            send_email(service, recipient, subject, body, attachment_path)
            time.sleep(1)

if __name__ == "__main__":
    main()