import imaplib
import email
from email.header import decode_header
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta

def gmail_login():
    try:
        load_dotenv(".env")
        # Account credentials
        username = os.environ["mail_username"]
        # get App password from: https://support.google.com/accounts/answer/185833?hl=en
        password = os.environ["mail_password"]
    except Exception:
        print("You should provide your username and password first in .env")
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        imap_ssl_host = 'imap.gmail.com'
        mail = imaplib.IMAP4_SSL(imap_ssl_host)
        mail.login(username, password)
    except Exception:
        print(f"Login failed: {Exception}")
    return mail

def last_day_mails(mail, folder = "inbox", days = 1, n = None):
    mail.select(folder)
    date_since = (datetime.now() - timedelta(days=days)).strftime('%d-%b-%Y')
    status, messages = mail.search(None, f'(SINCE {date_since})')
    mail_ids = messages[0].split()
    if n is not None:
        mail_ids = mail_ids[-n:]
    return mail_ids
    

def export_attachment(mail, mail_id, extensions = [".pdf", ".doc", ".docx", ".ppt", ".pptx"], attachment_folder = "attachments"):
    if not os.path.exists(attachment_folder):
        os.makedirs(attachment_folder)

    status, msg_data = mail.fetch(mail_id, "(RFC822)") 
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            # Parse the bytes email into a message object
            msg = email.message_from_bytes(response_part[1])
            sender = msg.get("From")
            if sender:
                sender_email = email.utils.parseaddr(sender)[1]
                sender_username = sender_email.split('@')[0] + "_"
            else:
                sender_username = ""
            # Check if the email has attachments
            if msg.is_multipart():
                for part in msg.walk():
                    # Check if the part is an attachment
                    if part.get_content_disposition() == "attachment":
                        # Get the filename and decode it if needed
                        filename = part.get_filename()
                        if sender_username.lower().startswith("granatz"):
                            return False
                        if sender_username.lower().startswith("barsony"):
                            return False
                        if sender_username.lower().startswith("szonja.toth"):
                            return False
                        if "tig" in filename.lower() or "szerz" in filename.lower():
                            return False
                        if "megbizasi" in filename.lower() or "szerz" in filename.lower():
                            return False
                        if "invoice" in filename.lower() or "szamla" in filename.lower():
                            return False
                        if "eshop" in filename.lower() or "szamla" in filename.lower():
                            return False
                        if "meghívó" in filename.lower() or "meghivo" in filename.lower():
                            return False
                        if "invite" in filename.lower() or "invitation" in filename.lower():
                            return False

                        try:
                            filename = decode_header(filename)[0][0]
                            if isinstance(filename, bytes):
                                filename = filename.decode('ISO-8859-1') # hungarian filenames
                            if any(filename.lower().endswith(ext) for ext in extensions):
                                filename = f"{sender_username}{filename}"
                                filepath = os.path.join(attachment_folder, filename)
                            else:
                                pass
                            if not os.path.exists(filepath):
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"Downloaded: {filename}")
                                return True
                            else:
                                print(f"{filename} already exist. Delete to force redownload")
                        except Exception:
                            print(f"Failed to download {filename}")

def export_attachments_from_folder(folder = "inbox", days = 30, n = None, extensions = [".pdf", ".doc", ".docx", ".ppt", ".pptx"], attachment_folder = "attachments"):
    mail = gmail_login()
    mail_ids = last_day_mails(mail = mail, folder = folder, days = days, n = n)
    for mail_id in mail_ids:
        export_attachment(mail = mail, mail_id = mail_id, extensions = extensions, attachment_folder = attachment_folder)
    mail.logout()

def main(days = 30) -> None:
    export_attachments_from_folder(folder = "inbox", days = days)
    export_attachments_from_folder(folder = "w_attachment", days = days)
    export_attachments_from_folder(folder = "mnb", days = days)

if __name__ == "__main__":
    mail = gmail_login()
    mail_ids = last_day_mails(mail = mail, folder="mnb", days = 30)
    for mail_id in mail_ids:
        export_attachment(mail = mail, mail_id = mail_id)
