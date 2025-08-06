import os
from imap_tools import MailBox, AND
import pandas as pd
from dotenv import load_dotenv

IMAP_HOST = "imap.gmail.com"  # Change for Outlook
EMAIL = os.getenv("EMAIL_USER")
PASSWORD = os.getenv("EMAIL_PASSWORD")

EMAIL_LIST_PATH = "uploads/test-emails.xlsx"
#TODO: Update this path to be dynamic to the name
DOWNLOAD_DIR = "uploads/raw_docs"


def fetch_and_save_attachments():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True) # Ensure the download directory exists

    # Load email-to-name mapping
    df = pd.read_excel(EMAIL_LIST_PATH) #read the Excel file
    email_to_name = {e.lower(): n.replace(" ", "_") for e, n in zip(df['Email'], df['Full Name'])}

    with MailBox(IMAP_HOST).login(EMAIL, PASSWORD) as mailbox:
        # Fetch unseen emails with attachments
        # Note: You can adjust the search criteria as needed
        for msg in mailbox.fetch(AND(seen=False, subject='RE: [TEST EMAIL] ADNOC Internship: Please Submit Your ID and Passport')):
            if msg.attachments:
                sender_email = msg.from_.lower()
                student_name = email_to_name.get(sender_email, "Unknown")

                for att in msg.attachments:
                    ext = os.path.splitext(att.filename)[1]
                    new_filename = f"{student_name}_{att.filename}"
                    file_path = os.path.join(DOWNLOAD_DIR, new_filename)

                    with open(file_path, "wb") as f:
                        f.write(att.payload)

                # Mark as seen so we don't reprocess
                mailbox.flag(msg.uid, '\\Seen', True)
