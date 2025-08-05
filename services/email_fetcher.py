import os
from imap_tools import MailBox, AND
import pandas as pd

IMAP_HOST = "imap.gmail.com"  # Change for Outlook
EMAIL = "your_email@example.com"
PASSWORD = "your_password"

EMAIL_LIST_PATH = "uploads/email_list.csv"
DOWNLOAD_DIR = "uploads/raw_docs"


def fetch_and_save_attachments():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # Load email-to-name mapping
    df = pd.read_csv(EMAIL_LIST_PATH)
    email_to_name = {e.lower(): n.replace(" ", "_") for e, n in zip(df['Email'], df['FullName'])}

    with MailBox(IMAP_HOST).login(EMAIL, PASSWORD) as mailbox:
        for msg in mailbox.fetch(AND(seen=False, has_attachment=True)):
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
