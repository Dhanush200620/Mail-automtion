# fetch_emails.py
# This script fetches emails from an inbox and returns them in a format
# that classify_leave.py can work with.
#
# It has two modes:
#   1. LOCAL mode — reads .txt files from the sample_emails/ folder (for testing)
#   2. IMAP mode — connects to a real email inbox via IMAP (Gmail, Outlook, etc.)
#
# I'm starting with LOCAL mode since I don't have a real shared mailbox yet.
# IMAP mode uses Python's built-in imaplib and email modules — no extra libraries needed.

import imaplib
import email
from email.header import decode_header
import os
import json
from datetime import datetime

# load config
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(script_dir, "config.json")

with open(config_path) as f:
    config = json.load(f)

SUBJECT_KEYWORDS = config.get("subject_keywords", ["leave", "sick", "absent", "holiday", "emergency"])


def fetch_local_emails(folder_path=None):
    # reads .txt email files from a local folder
    # this is for testing without connecting to a real inbox
    # returns a list of dicts with subject, sender, date, body
    
    if folder_path is None:
        folder_path = os.path.join(script_dir, config["sample_emails_folder"])
    
    if not os.path.exists(folder_path):
        print(f"ERROR: folder not found: {folder_path}")
        return []
    
    emails = []
    txt_files = sorted([f for f in os.listdir(folder_path) if f.endswith(".txt")])
    
    print(f"Found {len(txt_files)} email files in {folder_path}")
    
    for filename in txt_files:
        filepath = os.path.join(folder_path, filename)
        
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
        
        # parse the simple header format
        email_data = parse_text_email(content)
        email_data["source_file"] = filename
        email_data["source"] = "local"
        
        # check if subject has any leave-related keywords
        if has_leave_keywords(email_data["subject"]):
            emails.append(email_data)
            print(f"  [{filename}] Subject: {email_data['subject']} — MATCHED")
        else:
            print(f"  [{filename}] Subject: {email_data['subject']} — skipped (no keywords)")
    
    return emails


def parse_text_email(content):
    # parses a simple text email file (the format I'm using for sample emails)
    # headers at the top, blank line, then body
    
    lines = content.split("\n")
    subject = ""
    sender = ""
    date_str = ""
    body_start = 0
    
    for i, line in enumerate(lines):
        if line.strip() == "":
            body_start = i + 1
            break
        if line.lower().startswith("subject:"):
            subject = line.split(":", 1)[1].strip()
        elif line.lower().startswith("from:"):
            sender = line.split(":", 1)[1].strip()
        elif line.lower().startswith("date:"):
            date_str = line.split(":", 1)[1].strip()
    
    body = "\n".join(lines[body_start:]).strip()
    
    return {
        "subject": subject,
        "sender": sender,
        "date": date_str,
        "body": body,
    }


def fetch_imap_emails(mark_as_read=False):
    # connects to a real inbox via IMAP and fetches unread emails
    # that have leave-related keywords in the subject
    #
    # you need to set up these in config.json:
    #   imap_server: "imap.gmail.com" (or "outlook.office365.com")
    #   imap_port: 993
    #   imap_user: your email address
    #   imap_password: your app password (NOT your regular password!)
    #
    # for Gmail: go to Google Account > Security > App passwords
    # for Outlook: go to account.microsoft.com > Security > App passwords
    
    server = config.get("imap_server", "")
    port = config.get("imap_port", 993)
    user = config.get("imap_user", "")
    password = config.get("imap_password", "")
    
    if not server or not user or not password:
        print("ERROR: IMAP credentials not set in config.json")
        print("  Set imap_server, imap_user, and imap_password")
        return []
    
    emails = []
    
    try:
        # connect to IMAP server
        print(f"Connecting to {server}:{port}...")
        mail = imaplib.IMAP4_SSL(server, port)
        mail.login(user, password)
        print("Logged in OK")
        
        # select inbox
        mail.select("INBOX")
        
        # search for unread emails
        # we search for UNSEEN emails first, then filter by subject keywords ourselves
        # because IMAP search with OR on multiple subjects is clunky
        status, message_ids = mail.search(None, "UNSEEN")
        
        if status != "OK":
            print("ERROR: couldn't search inbox")
            mail.logout()
            return []
        
        ids = message_ids[0].split()
        print(f"Found {len(ids)} unread emails")
        
        for msg_id in ids:
            # fetch the email
            status, msg_data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue
            
            # parse the raw email
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # decode the subject
            subject = decode_subject(msg["Subject"])
            
            # check if this is a leave-related email
            if not has_leave_keywords(subject):
                continue
            
            # get the sender
            sender = msg["From"] or ""
            date_str = msg["Date"] or ""
            
            # get the body
            body = get_email_body(msg)
            
            email_data = {
                "subject": subject,
                "sender": sender,
                "date": date_str,
                "body": body,
                "source": "imap",
                "message_id": msg_id.decode(),
            }
            
            emails.append(email_data)
            print(f"  [IMAP] Subject: {subject} — MATCHED")
            
            # optionally mark as read so we don't process it again
            if mark_as_read:
                mail.store(msg_id, "+FLAGS", "\\Seen")
        
        mail.logout()
        print(f"Fetched {len(emails)} leave-related emails")
        
    except imaplib.IMAP4.error as e:
        print(f"IMAP error: {e}")
        print("Check your credentials in config.json")
    except Exception as e:
        print(f"ERROR connecting to email: {e}")
    
    return emails


def decode_subject(subject_header):
    # email subjects can be encoded in weird ways (like =?UTF-8?Q?...)
    # this decodes them into normal strings
    
    if subject_header is None:
        return ""
    
    decoded_parts = decode_header(subject_header)
    subject = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            subject += part.decode(encoding or "utf-8", errors="replace")
        else:
            subject += part
    
    return subject


def get_email_body(msg):
    # extracts the plain text body from an email message
    # emails can be multipart (with attachments, HTML version, etc.)
    # we just want the plain text
    
    body = ""
    
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                try:
                    body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                except:
                    body = str(part.get_payload())
                break
    else:
        try:
            body = msg.get_payload(decode=True).decode("utf-8", errors="replace")
        except:
            body = str(msg.get_payload())
    
    return body.strip()


def has_leave_keywords(subject):
    # checks if the subject contains any leave-related keywords
    subject_lower = subject.lower()
    for keyword in SUBJECT_KEYWORDS:
        if keyword.lower() in subject_lower:
            return True
    return False


def save_email_to_file(email_data, output_dir):
    # saves a fetched email to a .txt file so we can keep a record
    # and also re-process it later if needed
    
    # TODO: maybe use a better filename format? timestamp + sender?
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    sender_short = email_data["sender"].split("<")[0].strip().replace(" ", "_")[:20]
    filename = f"{timestamp}_{sender_short}.txt"
    filepath = os.path.join(output_dir, filename)
    
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(f"Subject: {email_data['subject']}\n")
        f.write(f"From: {email_data['sender']}\n")
        f.write(f"Date: {email_data['date']}\n")
        f.write(f"\n{email_data['body']}\n")
    
    return filepath


# ============================================================
# Test it
# ============================================================

if __name__ == "__main__":
    
    print("=" * 60)
    print("  fetch_emails.py — Test Run")
    print("=" * 60)
    
    # --- Test 1: LOCAL mode (sample emails) ---
    print("\n[Test 1] Fetching from local sample_emails/ folder...")
    print("-" * 40)
    local_emails = fetch_local_emails()
    print(f"\nGot {len(local_emails)} leave-related emails from local files")
    
    for e in local_emails:
        print(f"  From: {e['sender'][:30]}  |  Subject: {e['subject'][:40]}")
    
    # --- Test 2: IMAP mode (will fail if creds not set, that's OK) ---
    print("\n" + "-" * 40)
    print("[Test 2] Trying IMAP mode (will skip if creds not configured)...")
    
    if config.get("imap_user") and config.get("imap_password"):
        imap_emails = fetch_imap_emails()
        print(f"Got {len(imap_emails)} emails from IMAP")
    else:
        print("  IMAP credentials not set in config.json — skipping")
        print("  (This is fine for testing, set them when you have a real inbox)")
    
    print("\n" + "=" * 60)
    print("  Done!")
    print("=" * 60)
