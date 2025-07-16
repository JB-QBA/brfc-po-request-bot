from fastapi import FastAPI, Request, UploadFile, File
import os
import json
import base64
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders

app = FastAPI()
user_states = {}

# === CONFIG ===
SERVICE_ACCOUNT_FILE = "/etc/secrets/winged-pen-413708-e9544129b499.json"
EMAIL_SERVICE_ACCOUNT_FILE = "/etc/secrets/p2p-x-465909-c3e319be97b8.json"
SPREADSHEET_ID = "1U19XSieDNaDGN0khJJ8vFaDG75DwdKjE53d6MWi0Nt8"
SHEET_TAB_NAME = "FY2025Budget"
XERO_TAB_NAME = "Xero"
SENDER_EMAIL = "p2p.x@bahrainrfc.com"
CHAT_SPACE_ID = "spaces/AAQAs4dLeAY"

# === USERS & DEPARTMENTS ===
special_users = {
    "finance@bahrainrfc.com": "Johann",
    "generalmanager@bahrainrfc.com": "Paul"
}
department_managers = {
    "hr@bahrainrfc.com": "Human Capital",
    "facilities@bahrainrfc.com": "Facilities",
    "clubhouse@bahrainrfc.com": "Clubhouse",
    "sports@bahrainrfc.com": "Sports",
    "marketing@bahrainrfc.com": "Marketing",
    "sponsorship@bahrainrfc.com": "Sponsorship"
}
all_departments = [
    "Clubhouse", "Facilities", "Finance", "Front Office",
    "Human Capital", "Management", "Marketing", "Sponsorship", "Sports"
]
greeting_triggers = ["hi", "hello", "hey", "howzit", "salam", "hey cunt", "howdy"]

# === GOOGLE AUTH HELPERS ===
def get_gsheet():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=[
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets"
    ])
    return gspread.authorize(creds)

def get_chat_service():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/chat.bot"])
    return build("chat", "v1", credentials=creds)

def get_gmail_service():
    creds = Credentials.from_service_account_file(
        EMAIL_SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/gmail.send"],
        subject=SENDER_EMAIL
    )
    return build("gmail", "v1", credentials=creds)

# === GOOGLE SHEET FETCH ===
def get_cost_items_for_department(department: str):
    sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME)
    rows = sheet.get_all_values()[1:]
    return list(set(row[3] for row in rows if len(row) > 3 and row[1].strip().lower() == department.lower() and row[3].strip()))

def get_account_tracking_reference(cost_item: str, department: str):
    sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME)
    rows = sheet.get_all_values()
    headers = rows[0]
    data_rows = rows[1:]
    indices = {
        "Account": headers.index("Account"),
        "Department": headers.index("Department"),
        "Cost Item": headers.index("Cost Item"),
        "Tracking": headers.index("Tracking"),
        "Finance Reference": headers.index("Finance Reference"),
        "Total": headers.index("Total")
    }
    for row in data_rows:
        if row[indices["Cost Item"]].strip().lower() == cost_item.lower() and row[indices["Department"]].strip().lower() == department.lower():
            account = row[indices["Account"]].strip()
            tracking = row[indices["Tracking"]].strip()
            reference = row[indices["Finance Reference"]].strip()
            total = float(row[indices["Total"]].replace(",", "") or "0")
            return account, tracking, reference, total
    return None, None, None, 0

def get_total_budget_for_account(account: str, department: str):
    sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME)
    rows = sheet.get_all_values()[1:]
    return sum(float(row[17].replace(",", "")) for row in rows if len(row) >= 18 and row[0].strip().lower() == account.lower() and row[1].strip().lower() == department.lower())

def get_actuals_for_account(account: str, department: str):
    sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(XERO_TAB_NAME)
    rows = sheet.get_all_values()[3:]
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
    svc = get_chat_service()
    svc.spaces().messages().create(parent=CHAT_SPACE_ID, body={"text": text}).execute()

# === EMAIL SENDER ===
def send_quote_email(to_emails: list, subject: str, body_text: str, file: UploadFile):
    service = get_gmail_service()
    message = MIMEMultipart()
    message["to"] = ", ".join(to_emails)
    message["from"] = SENDER_EMAIL
    message["subject"] = subject

    # Add attachment
    mime = MIMEBase("application", "octet-stream")
    mime.set_payload(file.file.read())
    encoders.encode_base64(mime)
    mime.add_header("Content-Disposition", f"attachment; filename={file.filename}")
    message.attach(mime)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()

# === FILE UPLOAD ENDPOINT ===
@app.post("/upload/")
async def upload_file(request: Request, file: UploadFile = File(...)):
    body = await request.form()
    sender_email = body.get("sender_email", "").lower()
    full_name = body.get("full_name", "User")
    first_name = full_name.split()[0]

    # State recall
    cost_item = user_states.get(f"{sender_email}_cost_item")
    account = user_states.get(f"{sender_email}_account")
    department = user_states.get(f"{sender_email}_department")
    reference = user_states.get(f"{sender_email}_reference")

    # Send quote to ApprovalMax
    send_quote_email(
        ["bahrain-rugby-football-club-po@mail.approvalmax.com"],
        "New PO Quote Submission",
        f"Quote for {cost_item} under {department}",
        file
    )

    # Notify procurement
    summary = (
        f"📩 *New PO Request Received!*\n"
        f"*Cost Item:* {cost_item}\n"
        f"*Account:* {account}\n"
        f"*Department:* {department}\n"
        f"*Projects/Events/Budgets:* {reference}\n"
        f"*File:* {file.filename}\n\n"
        f"Please make sure that the approved PO is sent to {first_name}."
    )
    post_to_shared_space(summary)
    user_states[sender_email] = "awaiting_q1"
    return {"text": "1️⃣ Does this quote require any upfront payments?"}

# === CHAT MESSAGE HANDLER ===
@app.post("/")
async def chat_webhook(request: Request):
    body = await request.json()
    sender = body.get("message", {}).get("sender", {})
    sender_email = sender.get("email", "").lower()
    full_name = sender.get("displayName", "there")
    first_name = full_name.split()[0]
    message_text = body.get("message", {}).get("text", "").strip()
    state = user_states.get(sender_email)

    if state == "awaiting_q1":
        user_states[f"{sender_email}_q1"] = message_text
        user_states[sender_email] = "awaiting_q2"
        return {"text": "2️⃣ Is this a foreign payment that requires GSA approval?"}

    if state == "awaiting_q2":
        user_states[f"{sender_email}_q2"] = message_text
        user_states[sender_email] = "awaiting_comments"
        return {"text": "3️⃣ Any comments you'd like to pass along to the PO team?"}

    if state == "awaiting_comments":
        q1 = user_states.get(f"{sender_email}_q1", "N/A")
        q2 = user_states.get(f"{sender_email}_q2", "N/A")
        comments = message_text
        cost_item = user_states.get(f"{sender_email}_cost_item")
        account = user_states.get(f"{sender_email}_account")
        department = user_states.get(f"{sender_email}_department")
        reference = user_states.get(f"{sender_email}_reference")

        summary = (
            f"📋 *Finance Responses Received*\n"
            f"*From:* {first_name}\n"
            f"*Cost Item:* {cost_item}\n"
            f"*Account:* {account}\n"
            f"*Department:* {department}\n"
            f"*Reference:* {reference}\n"
            f"1️⃣ Upfront Payment Required: {q1}\n"
            f"2️⃣ Foreign Payment / GSA Approval: {q2}\n"
            f"3️⃣ Comments to PO Team: {comments}"
        )
        post_to_shared_space(summary)

        send_quote_email(
            ["finance@bahrainrfc.com", "accounts@bahrainrfc.com", "stores@bahrainrfc.com"],
            f"Finance PO Review – {cost_item}",
            summary,
            file=None  # optionally attach quote file if stored
        )

        user_states[sender_email] = None
        return {"text": f"Thanks {first_name}, you're all done ✅"}

    if state == "awaiting_cost_item":
        department = user_states.get(f"{sender_email}_department")
        account, tracking, reference, item_total = get_account_tracking_reference(message_text, department)
        if account:
            acct_total = get_total_budget_for_account(account, department)
            actuals = get_actuals_for_account(account, department)
            user_states.update({
                sender_email: "awaiting_file",
                f"{sender_email}_cost_item": message_text.title(),
                f"{sender_email}_account": account,
                f"{sender_email}_reference": reference
            })
            return {"text": (
                f"✅ You've selected: {message_text.title()} under {department}\n\n"
                f"📊 Budgeted for item: {int(item_total):,}\n"
                f"📊 Account '{account}' budget: {int(acct_total):,}\n"
                f"📊 YTD actuals: {int(actuals):,}\n\n"
                "📎 Please upload the quote file using our upload endpoint or web form.\n"
            )}
        else:
            return {"text": f"Cost item not found under {department}. Try again."}

    if any(message_text.lower().startswith(greet) for greet in greeting_triggers):
        if sender_email in special_users:
            user_states[sender_email] = "awaiting_department"
            return {"text": f"Hi {first_name}, what department is this PO for?\nOptions: {', '.join(all_departments)}"}
        elif sender_email in department_managers:
            dept = department_managers[sender_email]
            items = get_cost_items_for_department(dept)
            user_states[sender_email] = "awaiting_cost_item"
            user_states[f"{sender_email}_department"] = dept
            return {"text": f"Hi {first_name},\nHere are the cost items for {dept}:\n- " + "\n- ".join(items)}

    if state == "awaiting_department":
        if message_text.title() in all_departments:
            dept = message_text.title()
            items = get_cost_items_for_department(dept)
            user_states[sender_email] = "awaiting_cost_item"
            user_states[f"{sender_email}_department"] = dept
            return {"text": f"Thanks {first_name}. Cost items for {dept}:\n- " + "\n- ".join(items)}
        else:
            return {"text": f"Department not recognized. Try one of: {', '.join(all_departments)}"}

    return {"text": "I'm not sure how to help. Start with 'Hi' or pick a cost item."}
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
GMAIL_SERVICE_ACCOUNT_FILE = "/etc/secrets/p2p-x-465909-c3e319be97b8.json"  # <-- Your Gmail API JSON key
SENDER_EMAIL = "p2p.x@bahrainrfc.com"
RECIPIENT_EMAIL = "finance@bahrainrfc.com"

# Path to test attachment
# TEST_ATTACHMENT_PATH = "C:\\Users\\Finance Manager\\OneDrive\\Desktop\\Estimate 2663.pdf"

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
