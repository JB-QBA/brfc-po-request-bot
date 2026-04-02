# P2P 3000 Bot – ENHANCED VERSION with Financial Year Selection + Finance Requests + Google Tasks

from fastapi import FastAPI, Request
import os
import json
import base64
import gspread
import requests
import mimetypes
import hashlib
from datetime import datetime
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
NY_OPEX_TAB_NAME = "NY_OPEX"
NY_CAPEX_TAB_NAME = "NY_CAPEX"
XERO_TAB_NAME = "Xero"
SENDER_EMAIL = "p2p.x@bahrainrfc.com"
CHAT_SPACE_ID = "spaces/AAQAs4dLeAY"

# === FINANCE REQUEST CONFIG ===
FINANCE_CHAT_SPACE_ID = "spaces/AAAAA0zEepc"
FINANCE_MANAGER_EMAIL = "finance@bahrainrfc.com"
FINANCE_TEAM_EMAIL = "accounts@bahrainrfc.com"

finance_categories = {
    "1": "Supplier / Payment Query",
    "2": "Customer / Revenue Query",
    "3": "Unbudgeted Items",
    "4": "Staff Matters",
    "5": "Reporting Query",
    "6": "Other"
}

# Maps department manager emails to their title and name for Unbudgeted Items emails
department_manager_titles = {
    "hr@bahrainrfc.com": ("Human Capital Manager", "Human Capital"),
    "facilities@bahrainrfc.com": ("Facilities Manager", "Facilities"),
    "clubhouse@bahrainrfc.com": ("Clubhouse Manager", "Clubhouse"),
    "sports@bahrainrfc.com": ("Sports Manager", "Sports"),
    "marketing@bahrainrfc.com": ("Marketing Manager", "Marketing"),
    "sponsorship@bahrainrfc.com": ("Sponsorship Manager", "Sponsorship"),
    "gym@bahrainrfc.com": ("Gym Manager", "Sports"),
    "juniorsport@bahrainrfc.com": ("Junior Sport Manager", "Sports")
}

# Finance reference counter (in-memory, resets on restart - can be upgraded to persistent storage)
finance_ref_counter = {"count": 0}

# === USER GROUPS ===
special_users = {
    "finance@bahrainrfc.com": "Johann",
    "generalmanager@bahrainrfc.com": "Paul",
    "admins@bahrainrfc.com": "Sarah"
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

# === UTILITY FUNCTIONS ===
def is_july_or_august():
    """Check if current month is July or August"""
    current_month = datetime.now().month
    return current_month in [7, 8]

def get_sheet_tab_names(financial_year: str, request_type: str):
    """Get the appropriate sheet tab names based on financial year and request type"""
    if financial_year == "next":
        return NY_OPEX_TAB_NAME if request_type == "OPEX" else NY_CAPEX_TAB_NAME
    else:  # current year
        return SHEET_TAB_NAME if request_type == "OPEX" else CAPEX_TAB_NAME

def generate_finance_ref():
    """Generate a unique finance request reference number"""
    finance_ref_counter["count"] += 1
    return f"FIN-{datetime.now().strftime('%Y')}-{finance_ref_counter['count']:04d}"

def get_manager_title(sender_email: str):
    """Get the manager title for a given email"""
    if sender_email in department_manager_titles:
        return department_manager_titles[sender_email][0]
    elif sender_email in special_users:
        return "Finance Manager" if "finance" in sender_email else "General Manager"
    return "Team Member"

def get_manager_department(sender_email: str):
    """Get the department for a given email"""
    if sender_email in department_managers:
        return department_managers[sender_email]
    if sender_email in department_manager_titles:
        return department_manager_titles[sender_email][1]
    return "General"

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

def get_tasks_service(user_email: str):
    """Get Google Tasks service with domain-wide delegation for a specific user"""
    try:
        creds = get_creds(["https://www.googleapis.com/auth/tasks"])
        delegated_creds = creds.with_subject(user_email)
        return build("tasks", "v1", credentials=delegated_creds)
    except Exception as e:
        logger.error(f"Error creating Tasks service for {user_email}: {e}")
        return None

# === GOOGLE TASKS HELPER ===
def create_google_task(assignee_email: str, title: str, notes: str, due_date: str = None):
    """
    Create a Google Task for the specified user.
    due_date should be in YYYY/MM/DD format.
    """
    try:
        service = get_tasks_service(assignee_email)
        if not service:
            logger.error(f"Could not create Tasks service for {assignee_email}")
            return False

        task_body = {
            "title": title,
            "notes": notes
        }

        # Convert YYYY/MM/DD to RFC 3339 format for Google Tasks API
        if due_date and due_date != "URGENT":
            try:
                parsed_date = datetime.strptime(due_date, "%Y/%m/%d")
                task_body["due"] = parsed_date.strftime("%Y-%m-%dT00:00:00.000Z")
            except ValueError:
                logger.warning(f"Could not parse due date: {due_date}")

        # Create task in the user's primary task list
        task_lists = service.tasklists().list(maxResults=1).execute()
        if task_lists.get("items"):
            tasklist_id = task_lists["items"][0]["id"]
            result = service.tasks().insert(tasklist=tasklist_id, body=task_body).execute()
            logger.info(f"✅ Google Task created for {assignee_email}: {title} (ID: {result.get('id')})")
            return True
        else:
            logger.error(f"No task lists found for {assignee_email}")
            return False

    except Exception as e:
        logger.error(f"⚠️ Failed to create Google Task for {assignee_email}: {e}")
        return False

# === FINANCE SPACE + EMAIL HELPERS ===
def post_to_finance_space(text: str):
    """Post a message to the Finance Team Chat space"""
    try:
        get_chat_service().spaces().messages().create(parent=FINANCE_CHAT_SPACE_ID, body={"text": text}).execute()
        logger.info(f"📨 Posted to Finance Chat space")
    except Exception as e:
        logger.error(f"Error posting to Finance space: {e}")

def send_finance_request_email(ref_number: str, category: str, details: str, addressee: str,
                                deadline: str, requester_name: str, requester_email: str, department: str,
                                extra_fields: dict = None):
    """Send structured finance request email to Jo for tracking"""
    try:
        subject = f"[FINANCE-REQUEST] {ref_number} — {category} — {department}"

        body = (
            f"FINANCE-REQUEST-DATA\n"
            f"========================\n"
            f"Reference: {ref_number}\n"
            f"Category: {category}\n"
            f"Requester: {requester_name} ({requester_email})\n"
            f"Department: {department}\n"
            f"Addressed To: {addressee}\n"
            f"Deadline: {deadline}\n"
            f"Details: {details}\n"
        )

        if extra_fields:
            for key, value in extra_fields.items():
                body += f"{key}: {value}\n"

        body += (
            f"Request Timestamp: {datetime.now().isoformat()}\n"
            f"========================\n"
        )

        msg = MIMEMultipart('mixed')
        msg["From"] = SMTP_USERNAME
        msg["To"] = SENDER_EMAIL
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))

        smtp_password = os.getenv("SMTP_PASSWORD")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, smtp_password)
            server.sendmail(SMTP_USERNAME, [SENDER_EMAIL], msg.as_string())

        logger.info(f"📧 Finance request email sent for {ref_number}")
    except Exception as e:
        logger.error(f"⚠️ Failed to send Finance request email: {e}")

def send_unbudgeted_email(ref_number: str, requester_name: str, requester_email: str,
                           manager_title: str, department: str, filename: str, file_bytes: bytes,
                           content_type: str = None):
    """Send unbudgeted item email with attachment directly to Finance Manager.
    Uses same file extension detection as send_quote_email for proper format preservation."""
    try:
        import re

        subject = f"[UNBUDGETED] {department} — {requester_name} — {ref_number}"

        body = (
            f"Hi Johann,\n\n"
            f"The {manager_title}, {requester_name}, has requested confirmation on where the "
            f"attached quote is to be allocated within the existing budget.\n\n"
            f"Reference: {ref_number}\n"
            f"Department: {department}\n"
            f"Submitted: {datetime.now().strftime('%Y/%m/%d %H:%M')}\n\n"
            f"Please review the attached quote and advise on budget allocation.\n\n"
            f"Regards,\n"
            f"P2P 3000"
        )

        # --- File extension detection (same logic as send_quote_email) ---
        detected_extension = ""
        if file_bytes:
            if not '.' in filename or filename.split('.')[-1] not in ['pdf', 'xlsx', 'xls', 'docx', 'doc', 'png', 'jpg', 'jpeg']:
                if file_bytes.startswith(b'%PDF'):
                    detected_extension = ".pdf"
                elif file_bytes.startswith(b'PK'):
                    if content_type and 'spreadsheet' in content_type:
                        detected_extension = ".xlsx"
                    elif content_type and 'wordprocessing' in content_type:
                        detected_extension = ".docx"
                    else:
                        detected_extension = ".xlsx"
                elif file_bytes.startswith(b'\x89PNG'):
                    detected_extension = ".png"
                elif file_bytes.startswith(b'\xff\xd8\xff'):
                    detected_extension = ".jpg"

            # Clean filename for safe email attachment
            base_filename = re.sub(r'[^a-zA-Z0-9_.-]', '_', filename)

            # Generate meaningful filename if original is a Google Chat attachment ID
            if (len(base_filename) > 50 or
                'spaces_' in base_filename or
                'messages_' in base_filename or
                'attachments_' in base_filename):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                clean_requester = re.sub(r'[^a-zA-Z0-9]', '', requester_name) if requester_name else "User"
                base_filename = f"unbudgeted_{clean_requester}_{timestamp}"

            if detected_extension and not base_filename.lower().endswith(detected_extension.lower()):
                safe_filename = base_filename + detected_extension
            else:
                safe_filename = base_filename

            # Determine proper content type
            if not content_type:
                content_type, _ = mimetypes.guess_type(safe_filename)
                if not content_type:
                    ext = safe_filename.lower().split('.')[-1] if '.' in safe_filename else ''
                    content_type_map = {
                        'pdf': 'application/pdf',
                        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        'xls': 'application/vnd.ms-excel',
                        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'doc': 'application/msword',
                        'png': 'image/png',
                        'jpg': 'image/jpeg',
                        'jpeg': 'image/jpeg',
                    }
                    content_type = content_type_map.get(ext, 'application/octet-stream')

            logger.info(f"Unbudgeted file - Original: {filename}, Safe: {safe_filename}, Type: {content_type}")
        else:
            safe_filename = filename

        msg = MIMEMultipart('mixed')
        msg["From"] = SMTP_USERNAME
        msg["To"] = FINANCE_MANAGER_EMAIL
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # Attach the file with proper format handling (same as send_quote_email)
        if file_bytes:
            try:
                if content_type in ['application/pdf', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
                    attachment = MIMEApplication(file_bytes, _subtype=content_type.split('/')[-1])
                    attachment.add_header('Content-Disposition', f'attachment; filename="{safe_filename}"')
                    attachment.add_header('Content-Type', content_type)
                    msg.attach(attachment)
                else:
                    attachment = MIMEBase(*content_type.split('/'))
                    attachment.set_payload(file_bytes)
                    encoders.encode_base64(attachment)
                    attachment.add_header('Content-Disposition', f'attachment; filename="{safe_filename}"')
                    attachment.add_header('Content-Type', f'{content_type}; name="{safe_filename}"')
                    attachment.add_header('Content-Transfer-Encoding', 'base64')
                    msg.attach(attachment)
            except Exception as attachment_error:
                logger.warning(f"Primary attachment method failed: {attachment_error}")
                attachment = MIMEApplication(file_bytes, Name=safe_filename)
                attachment['Content-Disposition'] = f'attachment; filename="{safe_filename}"'
                msg.attach(attachment)

        smtp_password = os.getenv("SMTP_PASSWORD")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, smtp_password)
            server.sendmail(SMTP_USERNAME, [FINANCE_MANAGER_EMAIL], msg.as_string())

        logger.info(f"📧 Unbudgeted item email sent to Finance Manager for {ref_number} (file: {safe_filename})")
    except Exception as e:
        logger.error(f"⚠️ Failed to send unbudgeted email: {e}")

def parse_deadline_input(text: str):
    """Parse deadline input and return YYYY/MM/DD format or URGENT"""
    text = text.strip().lower()
    if text in ["urgent", "immediate", "asap", "now"]:
        return datetime.now().strftime("%Y/%m/%d"), True

    # Try common date formats
    current_year = datetime.now().year
    formats_to_try = [
        "%d %b",        # 10 Mar
        "%d %B",        # 10 March
        "%d/%m/%Y",     # 10/03/2026
        "%d/%m",        # 10/03
        "%Y/%m/%d",     # 2026/03/10
        "%d-%m-%Y",     # 10-03-2026
        "%d %b %Y",     # 10 Mar 2026
        "%d %B %Y",     # 10 March 2026
    ]

    for fmt in formats_to_try:
        try:
            parsed = datetime.strptime(text, fmt)
            # If no year was in the format, use current year
            if "%Y" not in fmt:
                parsed = parsed.replace(year=current_year)
                # If the date has already passed this year, use next year
                if parsed < datetime.now():
                    parsed = parsed.replace(year=current_year + 1)
            return parsed.strftime("%Y/%m/%d"), False
        except ValueError:
            continue

    # If we can't parse, return as-is with a note
    return text, False

def clear_user_state(sender_email: str):
    """Clear all state for a user"""
    for k in [k for k in user_states if k.startswith(f"{sender_email}_")]:
        user_states.pop(k)
    user_states[sender_email] = None

# === ENHANCED CAPEX SHEET HELPERS ===
def get_capital_items_for_department(department: str, sheet_tab: str = CAPEX_TAB_NAME):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(sheet_tab).get_all_values()[2:]  # Data starts at row 3
        return sorted(set(row[1] for row in rows if len(row) > 5 and row[5].strip().lower() == department.lower() and row[1].strip()))  # Asset Item (col B), Department (col F) - sorted alphabetically
    except Exception as e:
        logger.error(f"Error getting capital items from {sheet_tab}: {e}")
        return []

def get_capex_account_tracking_reference(asset_item: str, department: str, sheet_tab: str = CAPEX_TAB_NAME):
    try:
        sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(sheet_tab)
        rows = sheet.get_all_values()
        headers = rows[1]  # Headers in row 2 (index 1) for CAPEX
        data_rows = rows[2:]  # Data starts from row 3 (index 2)

        # Column mappings for CAPEX: B=Asset Item, C=Cost, F=Department, K=Projects/Events/Budgets, X=Xero Account
        asset_idx = 1  # Column B (Asset Item)
        cost_idx = 2   # Column C (Cost)
        dept_idx = 5   # Column F (Department)
        project_idx = 10  # Column K (Projects/Events/Budgets reference)
        account_idx = 23  # Column X (Xero account name)

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
        logger.error(f"Error getting CAPEX account tracking reference from {sheet_tab}: {e}")
        return None, None, 0

def get_capex_total_budget_for_account(account: str, project_ref: str, sheet_tab: str = CAPEX_TAB_NAME):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(sheet_tab).get_all_values()[2:]  # Data starts at row 3
        total = 0
        for row in rows:
            if len(row) > 23:  # Ensure row has enough columns
                if row[23].strip().lower() == account.lower() and row[10].strip().lower() == project_ref.lower():  # Column X (account) and Column K (project ref)
                    try:
                        cost_value = row[2].replace(",", "") if row[2] else "0"  # Column C (cost)
                        total += float(cost_value)
                    except:
                        pass
        return total
    except Exception as e:
        logger.error(f"Error getting CAPEX total budget from {sheet_tab}: {e}")
        return 0

def get_capex_actuals_for_account(account: str, project_ref: str):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(XERO_TAB_NAME).get_all_values()[3:]  # Same Xero logic
        total = 0.0
        for row in rows:
            if len(row) >= 16:  # Ensure we have enough columns (up to column P)
                # Match account name (column B, index 1), department (column O, index 14), and project reference (column P, index 15)
                account_match = row[1].strip().lower() == account.lower()
                project_match = row[15].strip().lower() == project_ref.lower()  # Column P (Projects/Events/Budgets)
                
                if account_match and project_match:
                    try:
                        # Clean the value thoroughly for negative amounts and whitespace
                        val = row[10].strip()  # Amount column (assuming same as OPEX)
                        
                        # Remove various types of whitespace and special characters
                        val = val.replace("\u00a0", "")  # Non-breaking space
                        val = val.replace("\u202f", "")  # Narrow no-break space
                        val = val.replace("\u2009", "")  # Thin space
                        val = val.replace("\u2008", "")  # Punctuation space
                        val = val.replace(" ", "")       # Regular space
                        val = val.replace("\t", "")      # Tab
                        val = val.replace("\n", "")      # Newline
                        val = val.replace("\r", "")      # Carriage return
                        
                        # Handle different minus signs
                        val = val.replace("−", "-")      # Unicode minus
                        val = val.replace("–", "-")      # En dash
                        val = val.replace("—", "-")      # Em dash
                        val = val.replace(",", "")       # Comma thousands separator
                        
                        # Convert to float if not empty
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

# === ENHANCED OPEX SHEET HELPERS ===
def get_cost_items_for_department(department: str, sheet_tab: str = SHEET_TAB_NAME):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(sheet_tab).get_all_values()[2:]  # Changed from [1:] to [2:] - now starts at row 3
        return sorted(set(row[3] for row in rows if len(row) > 3 and row[1].strip().lower() == department.lower() and row[3].strip()))
    except Exception as e:
        logger.error(f"Error getting cost items from {sheet_tab}: {e}")
        return []

def get_account_tracking_reference(cost_item: str, department: str, sheet_tab: str = SHEET_TAB_NAME):
    try:
        sheet = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(sheet_tab)
        rows = sheet.get_all_values()
        headers = rows[0]  # Headers still in row 1 (index 0)
        data_rows = rows[2:]  # Data now starts from row 3 (index 2) - skipping row 2

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
        logger.error(f"Error getting account tracking reference from {sheet_tab}: {e}")
        return None, None, None, 0

def get_total_budget_for_account(account: str, department: str, sheet_tab: str = SHEET_TAB_NAME):
    try:
        rows = get_gsheet().open_by_key(SPREADSHEET_ID).worksheet(sheet_tab).get_all_values()[2:]  # Changed from [1:] to [2:] - now starts at row 3
        return sum(float(row[17].replace(",", "")) for row in rows if len(row) >= 18 and row[0].strip().lower() == account.lower() and row[1].strip().lower() == department.lower())
    except Exception as e:
        logger.error(f"Error getting total budget from {sheet_tab}: {e}")
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

# === ENHANCED EMAIL SENDER (ApprovalMax-compatible) ===

def send_quote_email(to_emails, subject, body, filename, file_bytes, content_type=None, sender_name=None):
    """
    Enhanced email sender with better file format preservation for ApprovalMax compatibility
    """
    try:
        smtp_password = os.getenv("SMTP_PASSWORD")
        logger.info(f"SMTP_PASSWORD loaded dynamically: {'yes' if smtp_password else 'no'}, length: {len(smtp_password) if smtp_password else 0}")

        # Clean filename while preserving extension - ENHANCED for ApprovalMax
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
            return {"text": f"✅ No problem {first_name}, your request has been reset. Just say hi to begin again."}

        # === ATTACHMENT HANDLING ===
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

                # === UNBUDGETED ITEMS FILE UPLOAD ===
                if state == "awaiting_unbudgeted_file":
                    ref_number = generate_finance_ref()
                    department = get_manager_department(sender_email)
                    manager_title = get_manager_title(sender_email)

                    # Send email with attachment to Finance Manager
                    send_unbudgeted_email(ref_number, first_name, sender_email,
                                          manager_title, department, filename, file_bytes, final_content_type)

                    # Post to Finance Chat space
                    summary = (
                        f"📋 *Finance Request Received*\n"
                        f"*Reference:* {ref_number}\n"
                        f"*Category:* Unbudgeted Items\n"
                        f"*From:* {first_name} ({department})\n"
                        f"*File:* {filename}\n"
                        f"*Sent to:* Finance Manager\n"
                        f"*Submitted:* {datetime.now().strftime('%Y/%m/%d %H:%M')}"
                    )
                    post_to_finance_space(summary)

                    # Send tracking email to Jo
                    send_finance_request_email(ref_number, "Unbudgeted Items", 
                                               f"Quote uploaded: {filename}", "Finance Manager",
                                               "Review required", first_name, sender_email, department)

                    clear_user_state(sender_email)
                    return {"text": (
                        f"✅ File received. Thank you, {first_name}.\n\n"
                        f"Your request has been forwarded to the Finance Manager for review and budget allocation guidance.\n\n"
                        f"📋 **Reference:** {ref_number}\n"
                        f"📂 **Category:** Unbudgeted Items\n"
                        f"📧 **Sent to:** Finance Manager\n\n"
                        f"You will be notified once a decision has been made.\n"
                        f"Say **hi** to start a new request."
                    )}

                # === PO FLOW FILE UPLOAD (existing) ===
                if state == "awaiting_file":
                    # Send email with enhanced format preservation
                    send_quote_email(
                        ["bahrain-rugby-football-club-po@mail.approvalmax.com"],
                        "PO Quote Submission - Enhanced Format",
                        f"Quote uploaded by {first_name} ({sender_email})\n"
                        f"Original filename: {filename}\n"
                        f"Processed filename: (will be auto-detected with proper extension)\n"
                        f"Original content type: {original_content_type}\n"
                        f"Detected content type: {detected_content_type}\n"
                        f"Final content type: {final_content_type}\n"
                        f"File size: {len(file_bytes)} bytes\n"
                        f"File hash: {hashlib.sha256(file_bytes).hexdigest()}",
                        filename,
                        file_bytes,
                        final_content_type,
                        sender_name=first_name
                    )

                    # Get financial year and request type for shared space notification
                    financial_year = user_states.get(f"{sender_email}_financial_year", "current")
                    request_type = user_states.get(f"{sender_email}_request_type", "OPEX")
                    fy_text = "🔥 **NEXT YEAR REQUEST** 🔥" if financial_year == "next" else ""
                    
                    post_to_shared_space(f"📩 *Quote uploaded by {first_name}* — {filename} ({final_content_type}) {fy_text}")
                    user_states[sender_email] = "awaiting_supplier_name"
                    return {"text": f"✅ File received and forwarded: {filename}\n\nPlease carefully complete the following questions, in order to ensure an accurate and timely completion of your Purchase Order request:\n\n1️⃣ Kindly provide the name of the supplier?"}

                # If no state matches for file upload
                return {"text": f"📎 File received, but I wasn't expecting an upload right now. Say **hi** to start a new request."}

            except Exception as e:
                logger.error(f"Attachment error: {e}")
                return {"text": f"⚠️ Error handling attachment '{filename}': {str(e)}\n\nPlease try uploading the file again or contact support."}

        # ============================================================
        # TEXT FLOW - PO Questions (supplier details + existing Qs)
        # ============================================================
        if state == "awaiting_supplier_name":
            user_states[f"{sender_email}_supplier_name"] = message_text
            user_states[sender_email] = "awaiting_quote_value"
            return {"text": "2️⃣ Kindly confirm the total value for the supplier quote?"}

        if state == "awaiting_quote_value":
            user_states[f"{sender_email}_quote_value"] = message_text
            user_states[sender_email] = "awaiting_quote_reference"
            return {"text": "3️⃣ Kindly provide the quote or reference number/details?"}

        if state == "awaiting_quote_reference":
            user_states[f"{sender_email}_quote_reference"] = message_text
            user_states[sender_email] = "awaiting_q1"
            return {"text": "4️⃣ Does this quote require any upfront payments?"}

        if state == "awaiting_q1":
            user_states[f"{sender_email}_q1"] = message_text
            user_states[sender_email] = "awaiting_q2"
            return {"text": "5️⃣ Is this a foreign payment that requires GSA approval?"}

        if state == "awaiting_q2":
            user_states[f"{sender_email}_q2"] = message_text
            user_states[sender_email] = "awaiting_comments"
            return {"text": "6️⃣ Any comments you'd like to pass along to the PO team?"}

        if state == "awaiting_comments":
            q1 = user_states.get(f"{sender_email}_q1", "N/A")
            q2 = user_states.get(f"{sender_email}_q2", "N/A")
            supplier_name = user_states.get(f"{sender_email}_supplier_name", "N/A")
            quote_value = user_states.get(f"{sender_email}_quote_value", "N/A")
            quote_reference = user_states.get(f"{sender_email}_quote_reference", "N/A")
            comments = message_text
            cost_item = user_states.get(f"{sender_email}_cost_item")
            account = user_states.get(f"{sender_email}_account")
            department = user_states.get(f"{sender_email}_department")
            reference = user_states.get(f"{sender_email}_reference")
            request_type = user_states.get(f"{sender_email}_request_type", "OPEX")
            financial_year = user_states.get(f"{sender_email}_financial_year", "current")

            # Highlight next year requests for procurement team
            fy_notification = "🔥 **NEXT YEAR REQUEST - Please note this is for the upcoming financial year** 🔥\n" if financial_year == "next" else ""

            summary = (
                f"📋 *Finance Responses Received*\n"
                f"{fy_notification}"
                f"*From:* {first_name}\n"
                f"*Type:* {request_type} ({financial_year.title()} FY)\n"
                f"*Cost Item:* {cost_item}\n"
                f"*Account:* {account}\n"
                f"*Department:* {department}\n"
                f"*Reference:* {reference}\n"
                f"1️⃣ Supplier Name: {supplier_name}\n"
                f"2️⃣ Quote Value: {quote_value}\n"
                f"3️⃣ Quote/Reference Number: {quote_reference}\n"
                f"4️⃣ Upfront Payment Required: {q1}\n"
                f"5️⃣ Foreign Payment / GSA Approval: {q2}\n"
                f"6️⃣ Comments to PO Team: {comments}"
            )

            post_to_shared_space(summary)

            # Send structured PO request data to Jo for tracking and reconciliation
            try:
                po_request_subject = f"[PO-REQUEST] {request_type} - {supplier_name} - {department} ({financial_year.title()} FY)"
                po_request_body = (
                    f"PO-REQUEST-DATA\n"
                    f"================\n"
                    f"Requester: {first_name} ({sender_email})\n"
                    f"Request Type: {request_type}\n"
                    f"Financial Year: {financial_year.title()}\n"
                    f"Department: {department}\n"
                    f"Cost Item: {cost_item}\n"
                    f"Account: {account}\n"
                    f"Reference: {reference}\n"
                    f"Supplier Name: {supplier_name}\n"
                    f"Quote Value: {quote_value}\n"
                    f"Quote/Reference Number: {quote_reference}\n"
                    f"Upfront Payment Required: {q1}\n"
                    f"Foreign Payment / GSA Approval: {q2}\n"
                    f"Comments: {comments}\n"
                    f"Request Timestamp: {datetime.now().isoformat()}\n"
                    f"================\n"
                )
                
                # Send to bot email for Jo P2P Intelligence pickup
                msg = MIMEMultipart('mixed')
                msg["From"] = SMTP_USERNAME
                msg["To"] = SENDER_EMAIL
                msg["Subject"] = po_request_subject
                msg.attach(MIMEText(po_request_body, "plain", "utf-8"))
                
                smtp_password = os.getenv("SMTP_PASSWORD")
                with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                    server.starttls()
                    server.login(SMTP_USERNAME, smtp_password)
                    server.sendmail(SMTP_USERNAME, [SENDER_EMAIL], msg.as_string())
                
                logger.info(f"📧 PO Request email sent to Jo ({SENDER_EMAIL}) for {supplier_name} - {department}")
            except Exception as e:
                logger.error(f"⚠️ Failed to send PO request email to Jo: {e}")
                # Non-blocking: don't fail the user flow if Jo email fails
            
            # Clear user state
            clear_user_state(sender_email)
            
            return {"text": f"Thanks {first_name}, you're all done ✅\n\n📞 **Pro tip:** Feel free to follow up with the procurement team to make sure everything was received okay!"}

        # ============================================================
        # FINANCE REQUEST FLOW
        # ============================================================

        # --- Standard finance query: details → addressee → deadline ---
        if state == "awaiting_finance_details":
            user_states[f"{sender_email}_finance_details"] = message_text
            user_states[sender_email] = "awaiting_finance_addressee"
            return {"text": (
                "2️⃣ Who should this be addressed to?\n\n"
                "Type **A** for the **Finance Team**\n"
                "Type **B** for the **Finance Manager**"
            )}

        if state == "awaiting_finance_addressee":
            choice = message_text.strip().upper()
            if choice in ["A", "FINANCE TEAM", "TEAM"]:
                user_states[f"{sender_email}_finance_addressee"] = "Finance Team"
            elif choice in ["B", "FINANCE MANAGER", "MANAGER"]:
                user_states[f"{sender_email}_finance_addressee"] = "Finance Manager"
            else:
                return {"text": "Please type **A** for Finance Team or **B** for Finance Manager."}
            user_states[sender_email] = "awaiting_finance_deadline"
            return {"text": (
                "3️⃣ What is the required deadline for this request?\n\n"
                "Type a date in **YYYY/MM/DD** format (e.g. **2026/03/10**), or type **urgent** if this requires immediate attention."
            )}

        if state == "awaiting_finance_deadline":
            deadline_str, is_urgent = parse_deadline_input(message_text)
            category = user_states.get(f"{sender_email}_finance_category", "Other")
            details = user_states.get(f"{sender_email}_finance_details", "N/A")
            addressee = user_states.get(f"{sender_email}_finance_addressee", "Finance Team")
            department = get_manager_department(sender_email)
            ref_number = generate_finance_ref()

            urgent_flag = "🚨 " if is_urgent else ""
            deadline_display = f"URGENT — {deadline_str} (Today)" if is_urgent else deadline_str

            # Post to Finance Chat space
            summary = (
                f"{urgent_flag}📋 *Finance Request Received*\n"
                f"*Reference:* {ref_number}\n"
                f"*Category:* {category}\n"
                f"*From:* {first_name} ({department})\n"
                f"*Addressed To:* {addressee}\n"
                f"*Deadline:* {deadline_display}\n"
                f"*Details:* {details}"
            )
            post_to_finance_space(summary)

            # Send tracking email to Jo
            send_finance_request_email(ref_number, category, details, addressee,
                                        deadline_display, first_name, sender_email, department)

            # Create Google Task
            task_assignee = FINANCE_TEAM_EMAIL if addressee == "Finance Team" else FINANCE_MANAGER_EMAIL
            task_title = f"[{ref_number}] {category} — {first_name} ({department})"
            task_notes = f"Category: {category}\nFrom: {first_name} ({sender_email})\nDepartment: {department}\nDeadline: {deadline_display}\n\nDetails:\n{details}"
            create_google_task(task_assignee, task_title, task_notes, deadline_str)

            clear_user_state(sender_email)
            addressee_response = "The Finance Manager will respond to your query as a priority." if addressee == "Finance Manager" else "The Finance team will respond to your query shortly."
            return {"text": (
                f"Thanks {first_name}, your request has been logged and sent to Finance ✅\n\n"
                f"📋 **Reference:** {ref_number}\n"
                f"📂 **Category:** {category}\n"
                f"📧 **Addressed to:** {addressee}\n"
                f"⏰ **Deadline:** {deadline_display}\n\n"
                f"{addressee_response}\n"
                f"Say **hi** to start a new request."
            )}

        # --- Finance category selection ---
        if state == "awaiting_finance_category":
            choice = message_text.strip()
            if choice not in finance_categories:
                return {"text": (
                    "Please select a valid option (1-6):\n\n"
                    "**1** — Supplier / Payment Query\n"
                    "**2** — Customer / Revenue Query\n"
                    "**3** — Unbudgeted Items\n"
                    "**4** — Staff Matters\n"
                    "**5** — Reporting Query\n"
                    "**6** — Other"
                )}

            category = finance_categories[choice]
            user_states[f"{sender_email}_finance_category"] = category

            if choice == "3":  # Unbudgeted Items
                user_states[sender_email] = "awaiting_unbudgeted_file"
                return {"text": (
                    f"**Unbudgeted Items** — Got it! 📋\n\n"
                    f"Kindly upload the related quote for the costs that fall outside of our current budget.\n\n"
                    f"📎 Please upload the file directly here in Chat."
                )}
            else:  # Standard 3-question flow (1, 2, 4, 5, 6)
                emoji_map = {"1": "💳", "2": "📈", "4": "👥", "5": "📊", "6": "💬"}
                user_states[sender_email] = "awaiting_finance_details"
                return {"text": f"**{category}** — Got it! {emoji_map.get(choice, '✅')}\n\n1️⃣ Please provide the details of your query:"}

        # --- Finance or PO selection ---
        if state == "awaiting_finance_or_po":
            if message_text.strip().lower() in ["finance", "fin"]:
                user_states[sender_email] = "awaiting_finance_category"
                return {"text": (
                    "How can the Finance team help you today? 📊\n\n"
                    "**Please select from the options below:**\n\n"
                    "Press **1** for a **Supplier / Payment Query**\n"
                    "Press **2** for a **Customer / Revenue Query**\n"
                    "Press **3** for **Unbudgeted Items**\n"
                    "Press **4** for **Staff Matters**\n"
                    "Press **5** for a **Reporting Query**\n"
                    "Press **6** for **Other**"
                )}
            elif message_text.strip().lower() in ["po", "purchase", "purchase order"]:
                # Route to existing PO flow
                user_states[sender_email] = "awaiting_opex_capex"
                return {"text": (
                    "Is this request for:\n\n"
                    "💼 **Operational Expenditures** (type 'OPEX')\n"
                    "🏗️ **Capital Expenditures** (type 'CAPEX')\n\n"
                    "Please type either OPEX or CAPEX to continue."
                )}
            else:
                return {"text": "Please type either **Finance** or **PO** to continue with the related request."}

        # ============================================================
        # GREETING AND FINANCIAL YEAR SELECTION
        # ============================================================
        if any(message_text.lower().startswith(g) for g in greeting_triggers):
            if sender_email not in special_users and sender_email not in department_managers:
                return {"text": f"Hi {first_name}! I don't recognize your email address. Please contact an administrator to set up your access."}
            
            # Check if it's July or August - if so, ask about financial year
            if is_july_or_august():
                user_states[sender_email] = "awaiting_financial_year"
                return {"text": f"Hi {first_name}! 👋\n\nWhich financial year is this request for?\n\n📅 **Current** financial year (type 'current')\n📅 **Next** financial year (type 'next')\n\nPlease type either 'current' or 'next' to continue."}
            else:
                # Outside July/August, skip financial year question and default to current
                user_states[f"{sender_email}_financial_year"] = "current"
                user_states[sender_email] = "awaiting_finance_or_po"
                return {"text": (
                    f"Hi {first_name}! 👋\n\n"
                    f"Is this a **Finance** or **PO** request?\n\n"
                    f"Please type either **Finance** or **PO** to continue with the related request."
                )}

        # Handle financial year selection (only in July/August)
        if state == "awaiting_financial_year":
            if message_text.lower() in ["current", "current year", "this year"]:
                user_states[f"{sender_email}_financial_year"] = "current"
                user_states[sender_email] = "awaiting_finance_or_po"
                return {"text": (
                    "✅ Current financial year selected.\n\n"
                    "Is this a **Finance** or **PO** request?\n\n"
                    "Please type either **Finance** or **PO** to continue with the related request."
                )}
            elif message_text.lower() in ["next", "next year", "upcoming"]:
                user_states[f"{sender_email}_financial_year"] = "next"
                user_states[sender_email] = "awaiting_finance_or_po"
                return {"text": (
                    "✅ Next financial year selected.\n\n"
                    "Is this a **Finance** or **PO** request?\n\n"
                    "Please type either **Finance** or **PO** to continue with the related request."
                )}
            else:
                return {"text": f"Please type either 'current' for this financial year or 'next' for the upcoming financial year."}

        # Handle OPEX/CAPEX selection
        if state == "awaiting_opex_capex":
            financial_year = user_states.get(f"{sender_email}_financial_year", "current")
            
            if message_text.upper() in ["OPEX", "OPERATIONAL"]:
                user_states[f"{sender_email}_request_type"] = "OPEX"
                if sender_email in special_users:
                    user_states[sender_email] = "awaiting_department"
                    fy_text = f" ({financial_year.title()} FY)" if financial_year == "next" else ""
                    return {"text": f"Great! OPEX request confirmed{fy_text}. 💼\n\nWhat department is this PO for?\nOptions: {', '.join(all_departments)}"}
                elif sender_email in department_managers:
                    dept = department_managers[sender_email]
                    sheet_tab = get_sheet_tab_names(financial_year, "OPEX")
                    items = get_cost_items_for_department(dept, sheet_tab)
                    if not items:
                        return {"text": f"I couldn't find any cost items for {dept} in the {financial_year} financial year. Please contact support."}
                    user_states[sender_email] = "awaiting_cost_item"
                    user_states[f"{sender_email}_department"] = dept
                    fy_text = f" ({financial_year.title()} FY)" if financial_year == "next" else ""
                    return {"text": f"Great! OPEX request for {dept} confirmed{fy_text}. 💼\n\nHere are the cost items:\n- " + "\n- ".join(items)}
                else:
                    return {"text": f"I don't recognize your email address. Please contact an administrator to set up your access."}
            elif message_text.upper() in ["CAPEX", "CAPITAL"]:
                user_states[f"{sender_email}_request_type"] = "CAPEX"
                if sender_email in special_users:
                    user_states[sender_email] = "awaiting_department"
                    fy_text = f" ({financial_year.title()} FY)" if financial_year == "next" else ""
                    return {"text": f"Great! CAPEX request confirmed{fy_text}. 🏗️\n\nWhat department is this PO for?\nOptions: {', '.join(all_departments)}"}
                elif sender_email in department_managers:
                    dept = department_managers[sender_email]
                    sheet_tab = get_sheet_tab_names(financial_year, "CAPEX")
                    items = get_capital_items_for_department(dept, sheet_tab)
                    if not items:
                        return {"text": f"I couldn't find any capital items for {dept} in the {financial_year} financial year. Please contact support."}
                    user_states[sender_email] = "awaiting_cost_item"
                    user_states[f"{sender_email}_department"] = dept
                    fy_text = f" ({financial_year.title()} FY)" if financial_year == "next" else ""
                    return {"text": f"Great! CAPEX request for {dept} confirmed{fy_text}. 🏗️\n\nHere are the capital items:\n- " + "\n- ".join(items)}
                else:
                    return {"text": f"I don't recognize your email address. Please contact an administrator to set up your access."}
            else:
                return {"text": f"Please type either 'OPEX' for Operational Expenditures or 'CAPEX' for Capital Expenditures."}
                
        # Handle department selection (for special users)
        if state == "awaiting_department":
            financial_year = user_states.get(f"{sender_email}_financial_year", "current")
            request_type = user_states.get(f"{sender_email}_request_type", "OPEX")
            if message_text.title() in all_departments:
                dept = message_text.title()
                sheet_tab = get_sheet_tab_names(financial_year, request_type)
                
                if request_type == "CAPEX":
                    items = get_capital_items_for_department(dept, sheet_tab)
                    if not items:
                        fy_text = f" in the {financial_year} financial year" if financial_year == "next" else ""
                        return {"text": f"No capital items found for {dept}{fy_text}. Please contact support."}
                    user_states[sender_email] = "awaiting_cost_item"
                    user_states[f"{sender_email}_department"] = dept
                    fy_text = f" ({financial_year.title()} FY)" if financial_year == "next" else ""
                    return {"text": f"Thanks {first_name}. Capital items for {dept}{fy_text}:\n- " + "\n- ".join(items)}
                else:  # OPEX
                    items = get_cost_items_for_department(dept, sheet_tab)
                    if not items:
                        fy_text = f" in the {financial_year} financial year" if financial_year == "next" else ""
                        return {"text": f"No cost items found for {dept}{fy_text}. Please contact support."}
                    user_states[sender_email] = "awaiting_cost_item"
                    user_states[f"{sender_email}_department"] = dept
                    fy_text = f" ({financial_year.title()} FY)" if financial_year == "next" else ""
                    return {"text": f"Thanks {first_name}. Cost items for {dept}{fy_text}:\n- " + "\n- ".join(items)}
            else:
                return {"text": f"Department not recognized. Try one of: {', '.join(all_departments)}"}

        # Handle cost/capital item selection
        if state == "awaiting_cost_item":
            dept = user_states.get(f"{sender_email}_department")
            financial_year = user_states.get(f"{sender_email}_financial_year", "current")
            request_type = user_states.get(f"{sender_email}_request_type", "OPEX")
            sheet_tab = get_sheet_tab_names(financial_year, request_type)
            
            if request_type == "CAPEX":
                account, project_ref, item_cost = get_capex_account_tracking_reference(message_text, dept, sheet_tab)
                if account:
                    acct_total = get_capex_total_budget_for_account(account, project_ref, sheet_tab)
                    
                    # Only show actuals for current year, N/A for next year
                    if financial_year == "current":
                        actuals = get_capex_actuals_for_account(account, project_ref)
                        actuals_text = f"📊 YTD actuals: {int(actuals):,}"
                    else:
                        actuals_text = f"📊 YTD actuals: N/A (Next FY)"
                    
                    user_states.update({
                        sender_email: "awaiting_file",
                        f"{sender_email}_cost_item": message_text.title(),
                        f"{sender_email}_account": account,
                        f"{sender_email}_reference": project_ref,
                        f"{sender_email}_department": dept
                    })
                    
                    fy_text = f" ({financial_year.title()} FY)" if financial_year == "next" else ""
                    return {"text": (
                        f"✅ You've selected: {message_text.title()} under {dept}{fy_text}\n\n"
                        f"🏗️ **CAPEX Summary:**\n"
                        f"📊 Cost of this item: {int(item_cost):,}\n"
                        f"📊 Total budget for '{project_ref}': {int(acct_total):,}\n"
                        f"{actuals_text}\n\n"
                        "📎 Please upload the quote file directly here in Chat."
                    )}
                else:
                    fy_text = f" for the {financial_year} financial year" if financial_year == "next" else ""
                    return {"text": f"Capital item not found under {dept}{fy_text}. Please try again or type the exact item name."}
            else:  # OPEX
                account, tracking, reference, item_total = get_account_tracking_reference(message_text, dept, sheet_tab)
                if account:
                    acct_total = get_total_budget_for_account(account, dept, sheet_tab)
                    
                    # Only show actuals for current year, N/A for next year
                    if financial_year == "current":
                        actuals = get_actuals_for_account(account, dept)
                        actuals_text = f"📊 YTD actuals: {int(actuals):,}"
                    else:
                        actuals_text = f"📊 YTD actuals: N/A (Next FY)"
                    
                    user_states.update({
                        sender_email: "awaiting_file",
                        f"{sender_email}_cost_item": message_text.title(),
                        f"{sender_email}_account": account,
                        f"{sender_email}_reference": reference,
                        f"{sender_email}_department": dept
                    })
                    
                    fy_text = f" ({financial_year.title()} FY)" if financial_year == "next" else ""
                    return {"text": (
                        f"✅ You've selected: {message_text.title()} under {dept}{fy_text}\n\n"
                        f"💼 **OPEX Summary:**\n"
                        f"📊 Budgeted for item: {int(item_total):,}\n"
                        f"📊 Account '{account}' budget: {int(acct_total):,}\n"
                        f"{actuals_text}\n\n"
                        "📎 Please upload the quote file directly here in Chat."
                    )}
                else:
                    fy_text = f" for the {financial_year} financial year" if financial_year == "next" else ""
                    return {"text": f"Cost item not found under {dept}{fy_text}. Please try again or type the exact item name."}

        return {"text": "🤖 I'm not sure how to help. Say 'hi' to start or 'restart' to reset."}

    except Exception as e:
        logger.error(f"Webhook error: {e}")
        return {"text": f"❌ Unexpected error: {str(e)}. Please try again or contact support."}

@app.get("/health")
async def health_check():
    return {"status": "healthy"}

@app.get("/")
async def root():
    return {"message": "P2P 3000 Enhanced with Finance Requests + Google Tasks is running"}