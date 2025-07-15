from fastapi import FastAPI
import base64
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders

app = FastAPI()

# CONFIG
GMAIL_SERVICE_ACCOUNT_FILE = "/etc/secrets/gmail-api-key"  # <-- Your Gmail API JSON key
SENDER_EMAIL = "p2p.x@bahrainrfc.com"
RECIPIENT_EMAIL = "finance@bahrainrfc.com"

# Path to test attachment
TEST_ATTACHMENT_PATH = "C:\\Users\\Finance Manager\\OneDrive\\Desktop\\Estimate 2663.pdf"

# SCOPES
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

def get_gmail_service():
    creds = Credentials.from_service_account_file(
        GMAIL_SERVICE_ACCOUNT_FILE,
        scopes=GMAIL_SCOPES,
        subject=SENDER_EMAIL  # impersonate the Gmail sender
    )
    return build("gmail", "v1", credentials=creds)

def send_email_with_attachment():
    service = get_gmail_service()

    # Compose email with attachment
    message = MIMEMultipart()
    message["to"] = RECIPIENT_EMAIL
    message["from"] = SENDER_EMAIL
    message["subject"] = "PO Request Test - File Attachment Only"

    # Add attachment
    with open(TEST_ATTACHMENT_PATH, "rb") as f:
        mime = MIMEBase("application", "octet-stream")
        mime.set_payload(f.read())
        encoders.encode_base64(mime)
        mime.add_header("Content-Disposition", "attachment", filename=os.path.basename(TEST_ATTACHMENT_PATH))
        message.attach(mime)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    body = {"raw": raw}

    try:
        send_message = service.users().messages().send(userId="me", body=body).execute()
        print(f"✅ Message sent! ID: {send_message['id']}")
    except HttpError as error:
        print(f"❌ An error occurred: {error}")

# Run test once on startup
@app.on_event("startup")
def startup_event():
    send_email_with_attachment()
