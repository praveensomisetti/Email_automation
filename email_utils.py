# email_utils.py
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT'))
SMTP_USERNAME = os.getenv('SMTP_USERNAME')
SMTP_PASSWORD = os.getenv('SMTP_PASSWORD')

def send_email(to_email, subject, body):
    msg = MIMEMultipart()
    msg['From'] = SMTP_USERNAME
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(msg)
        return f"Email sent to {to_email}"
    except Exception as e:
        return f"Failed to send email to {to_email}. Error: {e}"

def send_emails(df):
    results = []
    for _, row in df.iterrows():
        recipient_email = row['contact email']
        email_subject = row['prepared_subject']
        email_body = row['prepared_body']
        result = send_email(recipient_email, email_subject, email_body)
        results.append(result)
    return results
