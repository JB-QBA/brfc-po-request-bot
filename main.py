import smtplib
import mimetypes
import hashlib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import logging

logger = logging.getLogger(__name__)

def send_quote_email_enhanced(to_emails, subject, body, filename, file_bytes, content_type=None):
    """
    Enhanced email sender with better file format preservation for ApprovalMax compatibility
    """
    try:
        smtp_password = os.getenv("SMTP_PASSWORD")
        logger.info(f"SMTP_PASSWORD loaded: {'yes' if smtp_password else 'no'}")

        # Clean filename while preserving extension
        import re
        safe_filename = re.sub(r'[^a-zA-Z0-9_.-]', '_', filename)
        
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

        # Create message
        msg = MIMEMultipart('mixed')  # Use 'mixed' for better compatibility
        msg["From"] = SMTP_USERNAME
        msg["To"] = ", ".join(to_emails)
        msg["Subject"] = subject
        
        # Add body
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # Method 1: Try MIMEApplication first (usually best for binary files)
        try:
            # For specific file types that ApprovalMax commonly handles
            if content_type in ['application/pdf', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
                attachment = MIMEApplication(file_bytes, _subtype=content_type.split('/')[-1])
                attachment.add_header('Content-Disposition', f'attachment; filename="{safe_filename}"')
                attachment.add_header('Content-Type', content_type)
                msg.attach(attachment)
                logger.info(f"Attached using MIMEApplication with Content-Type: {content_type}")
            else:
                # Method 2: Use MIMEBase for more control over headers
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
            text = msg.as_string()
            server.sendmail(SMTP_USERNAME, to_emails, text)

        logger.info(f"📧 Email sent successfully to {to_emails} with file: {safe_filename}")
        return True

    except Exception as e:
        logger.error(f"❌ Email sending failed: {e}")
        raise

def download_direct_file_enhanced(attachment_ref: dict) -> tuple[bytes, str]:
    """
    Enhanced file download that preserves content type information
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
        
    except Exception as e:
        logger.error(f"Error downloading file: {e}")
        raise

def download_drive_file_enhanced(file_id: str) -> tuple[bytes, str]:
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

# Updated attachment handling in your main webhook
def handle_attachment_enhanced(att, sender_email, first_name):
    """
    Enhanced attachment handling with better format preservation
    """
    try:
        filename = att.get("name", "quote.pdf")
        original_content_type = att.get("contentType", None)
        file_bytes = None
        detected_content_type = None

        logger.info(f"Processing attachment: {filename}")
        logger.info(f"Original content type from Chat API: {original_content_type}")
        logger.info(f"Full attachment data: {json.dumps(att, indent=2)}")

        if "driveDataRef" in att:
            file_id = att["driveDataRef"]["driveFileId"]
            file_bytes, detected_content_type = download_drive_file_enhanced(file_id)
        elif "attachmentDataRef" in att:
            file_bytes, detected_content_type = download_direct_file_enhanced(att)

        if not file_bytes:
            raise ValueError("File could not be loaded - no data received")

        # Use the most reliable content type available
        final_content_type = original_content_type or detected_content_type
        
        logger.info(f"File loaded successfully: {filename} ({len(file_bytes)} bytes)")
        logger.info(f"Final content type: {final_content_type}")
        logger.info(f"File hash: {hashlib.sha256(file_bytes).hexdigest()}")

        # Validate file integrity for common types
        if filename.lower().endswith('.pdf') and not file_bytes.startswith(b'%PDF'):
            logger.warning("PDF file does not start with %PDF header - potential corruption")
        elif filename.lower().endswith('.xlsx') and not file_bytes.startswith(b'PK'):
            logger.warning("XLSX file does not start with PK header - potential corruption")

        # Send email with enhanced format preservation
        send_quote_email_enhanced(
            ["p2p.x@bahrainrfc.com"],
            "PO Quote Submission - Enhanced Format",
            f"Quote uploaded by {first_name} ({sender_email})\n"
            f"Filename: {filename}\n"
            f"Original content type: {original_content_type}\n"
            f"Detected content type: {detected_content_type}\n"
            f"Final content type: {final_content_type}\n"
            f"File size: {len(file_bytes)} bytes\n"
            f"File hash: {hashlib.sha256(file_bytes).hexdigest()}",
            filename,
            file_bytes,
            final_content_type
        )

        return True, filename, final_content_type

    except Exception as e:
        logger.error(f"Enhanced attachment handling error: {e}")
        return False, filename, str(e)