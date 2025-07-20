# P2P 3000 Bot – FINAL WORKING VERSION (ApprovalMax-compatible) - FIXED FILE FORMAT

from fastapi import FastAPI, Request
import os
import json
import base64
import gspread
import requests
import mimetypes
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
from google.auth.transport.requests import Request as GoogleAuthRequest
import logging

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()
user_states = {}

# === CONFIG ===
SERVICE_ACCOUNT_FILE = "/etc/secrets/acoustic-agent-465113-s7-df3d0e19a05e.json"
SPREADSHEET_ID = "1U19XSieDNaDGN0khJJ8vFaDG75DwdKjE53d6MWi0Nt8"
SHEET_TAB_NAME = "FY2025Budget"
XERO_TAB_NAME = "Xero"
SENDER_EMAIL = "p2p.x@bahrainrfc.com"
CHAT_SPACE_ID = "spaces/AAQAs4dLeAY"

# === USER GROUPS ===
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
    "sponsorship@bahrainrfc.com": "Sponsorship",
    "gym@bahrainrfc.com": "Sports",
    "juniorsport@bahrainrfc.com": "Sports"
}

all_departments = [
    "Clubhouse", "Facilities", "Finance", "Front Office",
    "Human Capital", "Management", "Marketing", "Sponsorship", "Sports"
]

greeting_triggers = ["hi", "hello", "hey", "howzit", "salam", "hey cunt", "howdy"]
reset_triggers = ["cancel", "restart", "start over"]

# === GOOGLE AUTH ===
def get_creds(scopes):
    return Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)

def get_gsheet():
    return gspread.authorize(get_creds([
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets"
    ]))

def get_chat_service():
    return build("chat", "v1", credentials=get_creds(["https://www.googleapis.com/auth/chat.bot"]))

def get_gmail_service():
    return build("gmail", "v1", credentials=get_creds([
        "https://www.googleapis.com/auth/gmail.send"
    ]).with_subject(SENDER_EMAIL))

def get_drive_service():
    return build("drive", "v3", credentials=get_creds(["https://www.googleapis.com/auth/drive.readonly"]))

# === SHEET HELPERS ===
def get_cost_items_for_department(department: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME).get_all_values()[1:]
        return list(set(row[3] for row in rows if len(row) > 3 and row[1].strip().lower() == department.lower() and row[3].strip()))
    except Exception as e:
        logger.error(f"Error getting cost items: {e}")
        return []

def get_account_tracking_reference(cost_item: str, department: str):
    try:
        sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME)
        rows = sheet.get_all_values()
        headers = rows[0]
        data_rows = rows[1:]

        account_idx = headers.index("Account") if "Account" in headers else 0
        dept_idx = headers.index("Department") if "Department" in headers else 1
        item_idx = headers.index("Cost Item") if "Cost Item" in headers else 3
        tracking_idx = headers.index("Tracking") if "Tracking" in headers else 4
        ref_idx = headers.index("Finance Reference") if "Finance Reference" in headers else 5
        total_idx = headers.index("Total") if "Total" in headers else 17

        for row in data_rows:
            if len(row) > max(account_idx, dept_idx, item_idx, tracking_idx, ref_idx, total_idx):
                if row[item_idx].strip().lower() == cost_item.lower() and row[dept_idx].strip().lower() == department.lower():
                    account = row[account_idx].strip()
                    tracking = row[tracking_idx].strip()
                    reference = row[ref_idx].strip()
                    total_str = row[total_idx].replace(",", "") if row[total_idx] else "0"
                    try:
                        total_budget = float(total_str)
                    except:
                        total_budget = 0
                    return account, tracking, reference, total_budget

        return None, None, None, 0
    except Exception as e:
        logger.error(f"Error getting account tracking reference: {e}")
        return None, None, None, 0

def get_total_budget_for_account(account: str, department: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME).get_all_values()[1:]
        return sum(float(row[17].replace(",", "")) for row in rows if len(row) >= 18 and row[0].strip().lower() == account.lower() and row[1].strip().lower() == department.lower())
    except Exception as e:
        logger.error(f"Error getting total budget: {e}")
        return 0

def get_actuals_for_account(account: str, department: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(XERO_TAB_NAME).get_all_values()[3:]
        total = 0.0
        for row in rows:
            if len(row) >= 15 and row[1].strip().lower() == account.lower() and row[14].strip().lower() == department.lower():
                try:
                    val = row[10].strip().replace("−", "-").replace("–", "-").replace(",", "").replace(" ", "")
                    total += float(val)
                except:
                    pass
        return total
    except Exception as e:
        logger.error(f"Error getting actuals: {e}")
        return 0

def post_to_shared_space(text: str):
    try:
        get_chat_service().spaces().messages().create(parent=CHAT_SPACE_ID, body={"text": text}).execute()
    except Exception as e:
        logger.error(f"Error posting to shared space: {e}")

# === EMAIL SENDER (ApprovalMax-compatible) - FIXED FILE FORMAT ===
def send_quote_email(to_emails, subject, body, filename, file_bytes, content_type=None):
    """
    Fixed version that preserves original file format using attachment metadata
    """
    try:
        service = get_gmail_service()

        message = MIMEMultipart()
        message["To"] = ", ".join(to_emails)
        message["From"] = SENDER_EMAIL
        message["Subject"] = subject

        # Add plain text body
        message.attach(MIMEText(body, "plain"))

        # Use content_type from attachment metadata if available, otherwise guess
        if content_type:
            mime_type = content_type
            logger.info(f"Using attachment content type: {mime_type} for file: {filename}")
        else:
            mime_type, _ = mimetypes.guess_type(filename)
            if mime_type is None:
                mime_type = "application/octet-stream"
            logger.info(f"Guessed MIME type: {mime_type} for file: {filename}")
        
        # Create attachment with proper MIME type
        if mime_type.startswith('text/'):
            # For text files, decode and re-encode properly
            try:
                text_content = file_bytes.decode('utf-8')
                attachment = MIMEText(text_content, _subtype=mime_type.split('/')[-1])
                attachment.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            except:
                # Fallback to binary if decoding fails
                attachment = MIMEApplication(file_bytes, Name=filename)
                attachment.add_header('Content-Type', mime_type)
                attachment.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        else:
            # For binary files, use MIMEApplication with specific subtype
            try:
                main_type, sub_type = mime_type.split('/', 1)
                attachment = MIMEApplication(file_bytes, _subtype=sub_type, Name=filename)
            except:
                # Fallback to generic application
                attachment = MIMEApplication(file_bytes, Name=filename)
            
            # Set proper headers
            attachment.add_header('Content-Type', f'{mime_type}; name="{filename}"')
            attachment.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        
        message.attach(attachment)
        logger.info(f"MIME attachment hash (SHA256): {hashlib.sha256(attachment.as_bytes()).hexdigest()}")

        # Convert to raw format for Gmail API
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        
        # Send the email
        result = service.users().messages().send(userId="me", body={"raw": raw}).execute()
        logger.info(f"📧 Email sent successfully with message ID: {result.get('id')}")
        print(f"📧 Email sent correctly - File: {filename} ({mime_type})")

    except Exception as e:
        logger.error(f"❌ Email failed: {e}")
        print(f"❌ Email failed: {e}")
        # Re-raise the exception so calling code can handle it
        raise

# === FILE HANDLER - Enhanced with better error handling ===
def download_direct_file(attachment_ref: dict) -> bytes:
    """
    Enhanced file download with better error handling
    """
    try:
        data_ref = attachment_ref["attachmentDataRef"]
        attachment_token = data_ref["resourceName"]
        url = f"https://chat.googleapis.com/v1/media/{attachment_token}?alt=media"

        creds = get_creds(["https://www.googleapis.com/auth/chat.bot"])
        creds.refresh(GoogleAuthRequest())
        headers = {"Authorization": f"Bearer {creds.token}"}

        logger.info(f"Downloading file from: {url}")
        res = requests.get(url, headers=headers, timeout=30)  # Added timeout
        res.raise_for_status()
        
        logger.info(f"Successfully downloaded {len(res.content)} bytes")
        return res.content
        
    except requests.exceptions.RequestException as e:
        logger.error(f"Request error downloading file: {e}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error downloading file: {e}")
        raise

def download_drive_file(file_id: str) -> bytes:
    """
    Enhanced Drive file download with better error handling
    """
    try:
        logger.info(f"Downloading Drive file ID: {file_id}")
        drive_service = get_drive_service()
        file_content = drive_service.files().get_media(fileId=file_id).execute()
        logger.info(f"Successfully downloaded Drive file: {len(file_content)} bytes")
        return file_content
    except Exception as e:
        logger.error(f"Error downloading Drive file {file_id}: {e}")
        raise

# === MAIN CHAT HANDLER ===
@app.post("/")
async def chat_webhook(request: Request):
    try:
        body = await request.json()
        logger.info(json.dumps(body, indent=2))

        message = body.get("message", {})
        sender = message.get("sender", {})
        attachments = message.get("attachment", [])
        message_text = message.get("text", "").strip()
        sender_email = sender.get("email", "").lower()
        first_name = sender.get("displayName", "there").split()[0]
        state = user_states.get(sender_email)

        if message_text.lower() in reset_triggers:
            user_states.pop(sender_email, None)
            for k in [k for k in user_states if k.startswith(f"{sender_email}_")]:
                user_states.pop(k)
            return {"text": f"✅ No problem {first_name}, your PO flow has been reset. Just say hi to begin again."}

        if attachments:
            try:
                att = attachments[0]
                filename = att.get("name", "quote.pdf")
                content_type = att.get("contentType", None)  # Get original content type
                file_bytes = None

                logger.info(f"Processing attachment: {filename}")
                logger.info(f"Original content type: {content_type}")
                logger.info(f"Full attachment data: {json.dumps(att, indent=2)}")

                if "driveDataRef" in att:
                    file_id = att["driveDataRef"]["driveFileId"]
                    file_bytes = download_drive_file(file_id)
                    # For Drive files, try to get content type from Drive API
                    if not content_type:
                        try:
                            drive_service = get_drive_service()
                            file_metadata = drive_service.files().get(fileId=file_id, fields="mimeType,name").execute()
                            content_type = file_metadata.get("mimeType")
                            logger.info(f"Drive file content type: {content_type}")
                        except Exception as e:
                            logger.warning(f"Could not get Drive file metadata: {e}")
                            
                elif "attachmentDataRef" in att:
                    file_bytes = download_direct_file(att)

                if not file_bytes:
                    raise ValueError("File could not be loaded - no data received")

                logger.info(f"File loaded successfully: {filename} ({len(file_bytes)} bytes)")

                # Log SHA256 hash of the original file for diagnostics
                import hashlib
                logger.info(f"Original file hash (SHA256): {hashlib.sha256(file_bytes).hexdigest()}")

                # Send email with properly formatted file using original content type
                send_quote_email(
                    ["bahrain-rugby-football-club-po@mail.approvalmax.com"],
                    "PO Quote Submission",
                    f"Quote uploaded by {first_name} ({sender_email})\nFilename: {filename}\nOriginal content type: {content_type}",
                    filename,
                    file_bytes,
                    content_type  # Pass the original content type
                )

                post_to_shared_space(f"📩 *Quote uploaded by {first_name}* — {filename} ({content_type})")
                user_states[sender_email] = "awaiting_q1"
                return {"text": f"✅ File received and forwarded: {filename}\n\n1️⃣ Does this quote require any upfront payments?"}

            except Exception as e:
                logger.error(f"Attachment error: {e}")
                return {"text": f"⚠️ Error handling attachment '{filename}': {str(e)}\n\nPlease try uploading the file again or contact support."}

        # TEXT FLOW
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
            
            # Clear user state
            for k in [k for k in user_states if k.startswith(f"{sender_email}_")]:
                user_states.pop(k)
            user_states[sender_email] = None
            
            return {"text": f"Thanks {first_name}, you're all done ✅"}

        if any(message_text.lower().startswith(g) for g in greeting_triggers):
            if sender_email in special_users:
                user_states[sender_email] = "awaiting_department"
                return {"text": f"Hi {first_name}, what department is this PO for?\nOptions: {', '.join(all_departments)}"}
            elif sender_email in department_managers:
                dept = department_managers[sender_email]
                items = get_cost_items_for_department(dept)
                if not items:
                    return {"text": f"Hi {first_name}, I couldn't find any cost items for {dept}. Please contact support."}
                user_states[sender_email] = "awaiting_cost_item"
                user_states[f"{sender_email}_department"] = dept
                return {"text": f"Hi {first_name},\nHere are the cost items for {dept}:\n- " + "\n- ".join(items)}
            else:
                return {"text": f"Hi {first_name}! I don't recognize your email address. Please contact an administrator to set up your access."}

        if state == "awaiting_department":
            if message_text.title() in all_departments:
                dept = message_text.title()
                items = get_cost_items_for_department(dept)
                if not items:
                    return {"text": f"No cost items found for {dept}. Please contact support."}
                user_states[sender_email] = "awaiting_cost_item"
                user_states[f"{sender_email}_department"] = dept
                return {"text": f"Thanks {first_name}. Cost items for {dept}:\n- " + "\n- ".join(items)}
            else:
                return {"text": f"Department not recognized. Try one of: {', '.join(all_departments)}"}

        if state == "awaiting_cost_item":
            dept = user_states.get(f"{sender_email}_department")
            account, tracking, reference, item_total = get_account_tracking_reference(message_text, dept)
            if account:
                acct_total = get_total_budget_for_account(account, dept)
                actuals = get_actuals_for_account(account, dept)
                user_states.update({
                    sender_email: "awaiting_file",
                    f"{sender_email}_cost_item": message_text.title(),
                    f"{sender_email}_account": account,
                    f"{sender_email}_reference": reference,
                    f"{sender_email}_department": dept
                })
                return {"text": (
                    f"✅ You've selected: {message_text.title()} under {dept}\n\n"
                    f"📊 Budgeted for item: {int(item_total):,}\n"
                    f"📊 Account '{account}' budget: {int(acct_total):,}\n"
                    f"📊 YTD actuals: {int(actuals):,}\n\n"
                    "📎 Please upload the quote file directly here in Chat."
                )}
            else:
                return {"text": f"Cost item not found under {dept}. Please try again or type the exact item name."}

        return {"text": "🤖 I'm not sure how to help. Say 'hi' to start or 'restart' to reset."}

    except Exception as e:
        logger.error(f"Webhook error: {e}")
        return {"text": f"❌ Unexpected error: {str(e)}. Please try again or contact support."}

@app.get("/health")
async def health_check():
    return {"status": "healthy"}

@app.get("/")
async def root():
    return {"message": "P2P 3000 is running"}