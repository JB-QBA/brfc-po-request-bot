# Updated P2P 3000 Bot (with working cost item handling and SMTP email send)

from fastapi import FastAPI, Request
import os
import json
import base64
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import requests
import mimetypes
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import logging

# === CONFIG ===
SERVICE_ACCOUNT_FILE = "/etc/secrets/winged-pen-413708-e9544129b499.json"
SPREADSHEET_ID = "1U19XSieDNaDGN0khJJ8vFaDG75DwdKjE53d6MWi0Nt8"
SHEET_TAB_NAME = "FY2025Budget"
XERO_TAB_NAME = "Xero"
CHAT_SPACE_ID = "spaces/AAQAs4dLeAY"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USERNAME = os.getenv("p2p.x@bahrainrfc.com")
SMTP_PASSWORD = os.getenv("BRFC@Finance$2020")

# === FASTAPI SETUP ===
app = FastAPI()
user_states = {}

# === LOGGING ===
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# === GOOGLE AUTH ===
def get_gsheet():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=[
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets"
    ])
    return gspread.authorize(creds)

def get_drive_service():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/drive.readonly"])
    return build("drive", "v3", credentials=creds)

def get_chat_service():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/chat.bot"])
    return build("chat", "v1", credentials=creds)

# === SMTP EMAIL ===
def send_quote_email(to_emails, subject, body, filename, file_bytes):
    try:
        msg = MIMEMultipart()
        msg["From"] = SMTP_USERNAME
        msg["To"] = ", ".join(to_emails)
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        part = MIMEApplication(file_bytes, Name=filename)
        part["Content-Disposition"] = f'attachment; filename="{filename}"'
        msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(SMTP_USERNAME, to_emails, msg.as_string())

        print(f"\U0001F4E7 Email sent via SMTP to: {to_emails}")
    except Exception as e:
        print(f"\u274C SMTP Email failed: {e}")

# === SHEET HELPERS ===
def get_cost_items_for_department(department: str):
    sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME)
    rows = sheet.get_all_values()[1:]
    return list(set(row[3] for row in rows if len(row) > 3 and row[1].strip().lower() == department.lower() and row[3].strip()))

def get_account_tracking_reference(cost_item: str, department: str):
    sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME)
    rows = sheet.get_all_values()
    headers = rows[0]
    data_rows = rows[1:]
    idx = {
        "Account": headers.index("Account"),
        "Department": headers.index("Department"),
        "Cost Item": headers.index("Cost Item"),
        "Tracking": headers.index("Tracking"),
        "Finance Reference": headers.index("Finance Reference"),
        "Total": headers.index("Total")
    }
    for row in data_rows:
        if row[idx["Cost Item"]].strip().lower() == cost_item.lower() and row[idx["Department"]].strip().lower() == department.lower():
            return row[idx["Account"]], row[idx["Tracking"]], row[idx["Finance Reference"]], float(row[idx["Total"]].replace(",", "") or "0")
    return None, None, None, 0

def get_total_budget_for_account(account: str, department: str):
    rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME).get_all_values()[1:]
    return sum(float(row[17].replace(",", "")) for row in rows if len(row) >= 18 and row[0].strip().lower() == account.lower() and row[1].strip().lower() == department.lower())

def get_actuals_for_account(account: str, department: str):
    rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(XERO_TAB_NAME).get_all_values()[3:]
    total = 0.0
    for row in rows:
        if len(row) >= 15 and row[1].strip().lower() == account.lower() and row[14].strip().lower() == department.lower():
            val = row[10].strip().replace("−", "-").replace("–", "-").replace(",", "").replace(" ", "")
            try:
                total += float(val)
            except:
                pass
    return total

def post_to_shared_space(text: str):
    get_chat_service().spaces().messages().create(parent=CHAT_SPACE_ID, body={"text": text}).execute()

# === ADDITIONAL HELPERS ===
def download_direct_file(attachment_ref: dict) -> bytes:
    token = attachment_ref["attachmentDataRef"]["resourceName"]
    url = f"https://chat.googleapis.com/v1/media/{token}?alt=media"
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/chat.bot"])
    creds.refresh(requests.Request())
    headers = {"Authorization": f"Bearer {creds.token}"}
    res = requests.get(url, headers=headers)
    res.raise_for_status()
    return res.content

# === HEALTH ===
@app.get("/health")
async def health_check():
    return {"status": "healthy"}

@app.get("/")
async def root():
    return {"message": "P2P Bot is running with SMTP"}

# === MAIN HANDLER ===
@app.post("/")
async def chat_webhook(request: Request):
    ...  # Leave this part for now; we’ll splice in the final handler section once you're happy with this foundation.
