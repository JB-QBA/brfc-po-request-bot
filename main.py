# P2P 3000 Bot – ENHANCED VERSION with File Format Preservation and Financial Year Support

from fastapi import FastAPI, Request
import os
import json
import base64
import gspread
import requests
import mimetypes
import hashlib
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
from google.auth.transport.requests import Request as GoogleAuthRequest
import logging

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USERNAME = "p2p.x@bahrainrfc.com"
SMTP_PASSWORD = os.getenv("paiicvggoolfgwnh")  # Must be set in environment

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()
user_states = {}

# === CONFIG ===
SERVICE_ACCOUNT_FILE = "/etc/secrets/acoustic-agent-465113-s7-df3d0e19a05e.json"
SPREADSHEET_ID = "1U19XSieDNaDGN0khJJ8vFaDG75DwdKjE53d6MWi0Nt8"
SHEET_TAB_NAME = "CY_OPEX"
CAPEX_TAB_NAME = "CY_CAPEX"
NEXT_YEAR_OPEX_TAB_NAME = "NY_OPEX"
NEXT_YEAR_CAPEX_TAB_NAME = "NY_CAPEX"
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

# === CURRENT YEAR OPEX SHEET HELPERS ===
def get_cost_items_for_department(department: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME).get_all_values()[2:]
        return list(set(row[3] for row in rows if len(row) > 3 and row[1].strip().lower() == department.lower() and row[3].strip()))
    except Exception as e:
        logger.error(f"Error getting cost items: {e}")
        return []

def get_account_tracking_reference(cost_item: str, department: str):
    try:
        sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME)
        rows = sheet.get_all_values()
        headers = rows[0]
        data_rows = rows[2:]

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
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB_NAME).get_all_values()[2:]
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

# === CURRENT YEAR CAPEX SHEET HELPERS ===
def get_capital_items_for_department(department: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(CAPEX_TAB_NAME).get_all_values()[2:]
        return list(set(row[1] for row in rows if len(row) > 5 and row[5].strip().lower() == department.lower() and row[1].strip()))
    except Exception as e:
        logger.error(f"Error getting capital items: {e}")
        return []

def get_capex_account_tracking_reference(asset_item: str, department: str):
    try:
        sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(CAPEX_TAB_NAME)
        rows = sheet.get_all_values()
        headers = rows[1]
        data_rows = rows[2:]

        asset_idx = 1
        cost_idx = 2
        dept_idx = 5
        project_idx = 10
        account_idx = 23

        for row in data_rows:
            if len(row) > max(asset_idx, cost_idx, dept_idx, project_idx, account_idx):
                if row[asset_idx].strip().lower() == asset_item.lower() and row[dept_idx].strip().lower() == department.lower():
                    account = row[account_idx].strip()
                    project_ref = row[project_idx].strip()
                    cost_str = row[cost_idx].replace(",", "") if row[cost_idx] else "0"
                    try:
                        item_cost = float(cost_str)
                    except:
                        item_cost = 0
                    return account, project_ref, item_cost

        return None, None, 0
    except Exception as e:
        logger.error(f"Error getting CAPEX account tracking reference: {e}")
        return None, None, 0

def get_capex_total_budget_for_account(account: str, project_ref: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(CAPEX_TAB_NAME).get_all_values()[2:]
        total = 0
        for row in rows:
            if len(row) > 23:
                if row[23].strip().lower() == account.lower() and row[10].strip().lower() == project_ref.lower():
                    try:
                        cost_value = row[2].replace(",", "") if row[2] else "0"
                        total += float(cost_value)
                    except:
                        pass
        return total
    except Exception as e:
        logger.error(f"Error getting CAPEX total budget: {e}")
        return 0

def get_capex_actuals_for_account(account: str, project_ref: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(XERO_TAB_NAME).get_all_values()[3:]
        total = 0.0
        for row in rows:
            if len(row) >= 16:
                account_match = row[1].strip().lower() == account.lower()
                project_match = row[15].strip().lower() == project_ref.lower()
                
                if account_match and project_match:
                    try:
                        val = row[10].strip()
                        val = val.replace("\u00a0", "").replace("\u202f", "").replace("\u2009", "").replace("\u2008", "")
                        val = val.replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")
                        val = val.replace("−", "-").replace("–", "-").replace("—", "-").replace(",", "")
                        
                        if val and val != "-":
                            total += float(val)
                    except (ValueError, TypeError) as e:
                        logger.warning(f"Could not parse value '{row[10]}' for CAPEX actuals: {e}")
                        pass
        
        logger.info(f"CAPEX actuals total for account '{account}' and project '{project_ref}': {total}")
        return total
    except Exception as e:
        logger.error(f"Error getting CAPEX actuals: {e}")
        return 0

# === NEXT YEAR OPEX SHEET HELPERS ===
def get_ny_cost_items_for_department(department: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(NEXT_YEAR_OPEX_TAB_NAME).get_all_values()[2:]
        return list(set(row[3] for row in rows if len(row) > 3 and row[1].strip().lower() == department.lower() and row[3].strip()))
    except Exception as e:
        logger.error(f"Error getting NY cost items: {e}")
        return []

def get_ny_account_tracking_reference(cost_item: str, department: str):
    try:
        sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(NEXT_YEAR_OPEX_TAB_NAME)
        rows = sheet.get_all_values()
        headers = rows[0]
        data_rows = rows[2:]

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
        logger.error(f"Error getting NY account tracking reference: {e}")
        return None, None, None, 0

def get_ny_total_budget_for_account(account: str, department: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(NEXT_YEAR_OPEX_TAB_NAME).get_all_values()[2:]
        return sum(float(row[17].replace(",", "")) for row in rows if len(row) >= 18 and row[0].strip().lower() == account.lower() and row[1].strip().lower() == department.lower())
    except Exception as e:
        logger.error(f"Error getting NY total budget: {e}")
        return 0

# === NEXT YEAR CAPEX SHEET HELPERS ===
def get_ny_capital_items_for_department(department: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(NEXT_YEAR_CAPEX_TAB_NAME).get_all_values()[2:]
        return list(set(row[1] for row in rows if len(row) > 5 and row[5].strip().lower() == department.lower() and row[1].strip()))
    except Exception as e:
        logger.error(f"Error getting NY capital items: {e}")
        return []

def get_ny_capex_account_tracking_reference(asset_item: str, department: str):
    try:
        sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(NEXT_YEAR_CAPEX_TAB_NAME)
        rows = sheet.get_all_values()
        headers = rows[1]
        data_rows = rows[2:]

        asset_idx = 1
        cost_idx = 2
        dept_idx = 5
        project_idx = 10
        account_idx = 23

        for row in data_rows:
            if len(row) > max(asset_idx, cost_idx, dept_idx, project_idx, account_idx):
                if row[asset_idx].strip().lower() == asset_item.lower() and row[dept_idx].strip().lower() == department.lower():
                    account = row[account_idx].strip()
                    project_ref = row[project_idx].strip()
                    cost_str = row[cost_idx].replace(",", "") if row[cost_idx] else "0"
                    try:
                        item_cost = float(cost_str)
                    except:
                        item_cost = 0
                    return account, project_ref, item_cost

        return None, None, 0
    except Exception as e:
        logger.error(f"Error getting NY CAPEX account tracking reference: {e}")
        return None, None, 0

def get_ny_capex_total_budget_for_account(account: str, project_ref: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(NEXT_YEAR_CAPEX_TAB_NAME).get_all_values()[2:]
        total = 0
        for row in rows:
            if len(row) > 23:
                if row[23].strip().lower() == account.lower() and row[10].strip().lower() == project_ref.lower():
                    try:
                        cost_value = row[2].replace(",", "") if row[2] else "0"
                        total += float(cost_value)
                    except:
                        pass
        return total
    except Exception as e:
        logger.error(f"Error getting NY CAPEX total budget: {e}")
        return 0

def post_to_shared_space(text: str):
    try:
        get_chat_service().spaces().messages().create(parent=CHAT_SPACE_ID, body={"text": text}).execute()
    except Exception as e:
        logger.error(f"Error posting to shared space: {e}")

# === ENHANCED EMAIL SENDER (ApprovalMax-compatible) ===

def send_quote_email(to_emails, subject, body, filename, file_bytes, content_type=None, sender_name=None):
    """
    Enhanced email sender with better file format preservation for ApprovalMax compatibility
    """
    try:
        smtp_password = os.getenv("SMTP_PASSWORD")
        logger.info(f"SMTP_PASSWORD loaded dynamically: {'yes' if smtp_password else 'no'}, length: {len(smtp_password) if smtp_password else 0}")

        # Clean filename and add extension if needed - ENHANCED for ApprovalMax
        import re
        
        # First, detect the file type if no extension exists
        detected_extension = ""
        if not '.' in filename or filename.split('.')[-1] not in ['pdf', 'xlsx', 'xls', 'docx', 'doc', 'png', 'jpg', 'jpeg']:
            # Detect file type from content
            if file_bytes.startswith(b'%PDF'):
                detected_extension = ".pdf"
            elif file_bytes.startswith(b'PK'):  # ZIP-based formats (xlsx, docx)
                if content_type and 'spreadsheet' in content_type:
                    detected_extension = ".xlsx"
                elif content_type and 'wordprocessing' in content_type:
                    detected_extension = ".docx"
                else:
                    detected_extension = ".xlsx"  # Default to xlsx for PK files
            elif file_bytes.startswith(b'\x89PNG'):
                detected_extension = ".png"
            elif file_bytes.startswith(b'\xff\xd8\xff'):
                detected_extension = ".jpg"
        
        # Clean filename and add extension if needed - ENHANCED for meaningful names
        base_filename = re.sub(r'[^a-zA-Z0-9_.-]', '_', filename)
        
        # Generate a meaningful filename if the original is a Google Chat attachment ID
        if (len(base_filename) > 50 or 
            'spaces_' in base_filename or 
            'messages_' in base_filename or 
            'attachments_' in base_filename):
            # Create a meaningful filename with sender name and timestamp
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            # Clean the sender name for filename use
            clean_sender_name = re.sub(r'[^a-zA-Z0-9]', '', sender_name) if sender_name else "User"
            base_filename = f"quote_{clean_sender_name}_{timestamp}"
            
        if detected_extension and not base_filename.lower().endswith(detected_extension.lower()):
            safe_filename = base_filename + detected_extension
        else:
            safe_filename = base_filename
            
        logger.info(f"Original filename: {filename}")
        logger.info(f"Safe filename: {safe_filename}")
        logger.info(f"Detected extension: {detected_extension}")
        
        # Ensure we have the correct file extension and MIME type
        if not content_type:
            content_type, _ = mimetypes.guess_type(filename)
            if not content_type:
                # Default based on file extension
                ext = filename.lower().split('.')[-1] if '.' in filename else ''
                content_type_map = {
                    'pdf': 'application/pdf',
                    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'xls': 'application/vnd.ms-excel',
                    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'doc': 'application/msword',
                    'png': 'image/png',
                    'jpg': 'image/jpeg',
                    'jpeg': 'image/jpeg',
                    'txt': 'text/plain',
                    'csv': 'text/csv'
                }
                content_type = content_type_map.get(ext, 'application/octet-stream')
        
        logger.info(f"Processing file: {filename}")
        logger.info(f"Content-Type: {content_type}")
        logger.info(f"File size: {len(file_bytes)} bytes")
        logger.info(f"File hash: {hashlib.sha256(file_bytes).hexdigest()}")

        # Write file to disk for debugging
        file_path = f"/tmp/{safe_filename}"
        with open(file_path, "wb") as f:
            f.write(file_bytes)
        logger.info(f"File written to disk at {file_path}")

        # Create message
        msg = MIMEMultipart('mixed')  # Use 'mixed' for better compatibility
        msg["From"] = SMTP_USERNAME
        msg["To"] = ", ".join(to_emails)
        msg["Subject"] = subject
        
        # Add body
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # Enhanced attachment handling
        try:
            # For specific file types that ApprovalMax commonly handles
            if content_type in ['application/pdf', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
                attachment = MIMEApplication(file_bytes, _subtype=content_type.split('/')[-1])
                attachment.add_header('Content-Disposition', f'attachment; filename="{safe_filename}"')
                attachment.add_header('Content-Type', content_type)
                msg.attach(attachment)
                logger.info(f"Attached using MIMEApplication with Content-Type: {content_type}")
            else:
                # Use MIMEBase for more control over headers
                attachment = MIMEBase(*content_type.split('/'))
                attachment.set_payload(file_bytes)
                encoders.encode_base64(attachment)
                attachment.add_header('Content-Disposition', f'attachment; filename="{safe_filename}"')
                attachment.add_header('Content-Type', f'{content_type}; name="{safe_filename}"')
                attachment.add_header('Content-Transfer-Encoding', 'base64')
                msg.attach(attachment)
                logger.info(f"Attached using MIMEBase with Content-Type: {content_type}")
                
        except Exception as attachment_error:
            logger.warning(f"Primary attachment method failed: {attachment_error}")
            # Fallback: Generic binary attachment
            attachment = MIMEApplication(file_bytes, Name=safe_filename)
            attachment['Content-Disposition'] = f'attachment; filename="{safe_filename}"'
            msg.attach(attachment)
            logger.info("Used fallback MIMEApplication attachment method")

        # Send email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, smtp_password)
            server.sendmail(SMTP_USERNAME, to_emails, msg.as_string())

        logger.info(f"📧 SMTP Email sent successfully to {to_emails} with file: {filename}")
        print(f"📧 Email sent correctly via SMTP - File: {filename}")

    except Exception as e:
        logger.error(f"❌ SMTP Email failed: {e}")
        print(f"❌ SMTP Email failed: {e}")
        raise

def download_direct_file(attachment_ref: dict) -> tuple:
    """
    Enhanced file download with content type preservation
    """
    try:
        data_ref = attachment_ref["attachmentDataRef"]
        attachment_token = data_ref["resourceName"]
        url = f"https://chat.googleapis.com/v1/media/{attachment_token}?alt=media"

        creds = get_creds(["https://www.googleapis.com/auth/chat.bot"])
        creds.refresh(GoogleAuthRequest())
        headers = {"Authorization": f"Bearer {creds.token}"}

        logger.info(f"Downloading file from: {url}")
        res = requests.get(url, headers=headers, timeout=30)
        res.raise_for_status()
        
        # Try to get content type from response headers
        response_content_type = res.headers.get('content-type', '')
        logger.info(f"Response Content-Type: {response_content_type}")
        logger.info(f"Response headers: {dict(res.headers)}")
        
        logger.info(f"Successfully downloaded {len(res.content)} bytes")
        return res.content, response_content_type
        
    except requests.exceptions.RequestException as e:
        logger.error(f"Request error downloading file: {e}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error downloading file: {e}")
        raise

def download_drive_file(file_id: str) -> tuple:
    """
    Enhanced Drive file download with content type preservation
    """
    try:
        logger.info(f"Downloading Drive file ID: {file_id}")
        drive_service = get_drive_service()
        
        # Get file metadata first
        file_metadata = drive_service.files().get(fileId=file_id, fields="mimeType,name,size").execute()
        content_type = file_metadata.get("mimeType", "application/octet-stream")
        file_name = file_metadata.get("name", "unknown_file")
        file_size = file_metadata.get("size", "unknown")
        
        logger.info(f"Drive file metadata - Name: {file_name}, Type: {content_type}, Size: {file_size}")
        
        # Download file content
        file_content = drive_service.files().get_media(fileId=file_id).execute()
        logger.info(f"Successfully downloaded Drive file: {len(file_content)} bytes")
        
        return file_content, content_type
        
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
                original_content_type = att.get("contentType", None)
                file_bytes = None
                detected_content_type = None

                logger.info(f"Processing attachment: {filename}")
                logger.info(f"Original content type from Chat API: {original_content_type}")
                logger.info(f"Full attachment data: {json.dumps(att, indent=2)}")

                if "driveDataRef" in att:
                    file_id = att["driveDataRef"]["driveFileId"]
                    file_bytes, detected_content_type = download_drive_file(file_id)
                elif "attachmentDataRef" in att:
                    file_bytes, detected_content_type = download_direct_file(att)

                if not file_bytes:
                    raise ValueError("File could not be loaded - no data received")

                # Use the most reliable content type available
                final_content_type = original_content_type or detected_content_type
                
                logger.info(f"File loaded successfully: {filename} ({len(file_bytes)} bytes)")
                logger.info(f"Final content type: {final_content_type}")

                # Validate file integrity for common types
                if filename.lower().endswith('.pdf') and not file_bytes.startswith(b'%PDF'):
                    logger.warning("PDF file does not start with %PDF header - potential corruption")
                elif filename.lower().endswith('.xlsx') and not file_bytes.startswith(b'PK'):
                    logger.warning("XLSX file does not start with PK header - potential corruption")

                # Log SHA256 hash of the original file for diagnostics
                logger.info(f"Original file hash (SHA256): {hashlib.sha256(file_bytes).hexdigest()}")

                # Send email with enhanced format preservation
                send_quote_email(
                    ["bahrain-rugby-football-club-po@mail.approvalmax.com"],
                    "PO Quote Submission - Enhanced Format",
                    f"Quote uploaded by {first_name} ({sender_email})\n"
                    f"Original filename: {filename}\n"
                    f"Processed filename: (will be auto-detected with proper extension)\n"
                    f"Original content type: {original_content_type}\n