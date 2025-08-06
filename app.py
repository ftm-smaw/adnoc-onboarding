from flask import Flask, request, render_template, jsonify
import pandas as pd
import os
from dotenv import load_dotenv
import smtplib
from email.message import EmailMessage
from services.email_fetcher import fetch_and_save_attachments
from pync import Notifier

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Load environment variables
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Global temporary email list
temp_list_student_email = []

@app.route("/upload", methods=["POST"])
def upload():
    global temp_list_student_email
    temp_list_student_email = []

    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    if not file.filename.endswith('.xlsx') and not file.filename.lower().endswith('.xls'):
        return "Invalid file type. Please upload an .xlsx or .xls file", 400

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    try:
        df = pd.read_excel(filepath)

        # Look for any column with 'email' in its name (case-insensitive)
        email_cols = [col for col in df.columns if 'email' in col.lower()]
        if not email_cols:
            return "No column containing 'email' found.", 400

        email_col = email_cols[0]  # Use the first one found
        temp_list_student_email = df[email_col].dropna().unique().tolist()

        # Send emails to all
        for email in temp_list_student_email:
            send_email(email)

        return f"Emails sent to {len(temp_list_student_email)} students!",200
        #notif = "Emails sent to {len(temp_list_student_email)} students!"
        ## Notify user
        #Notifier.notify(notif, title="Sending emails successful")

    except Exception as e:
        print("Error:", e)
        return f"Failed: {e}", 500


def send_email(to_email):
    print(f"Sending email to {to_email}")

    msg = EmailMessage()
    msg['Subject'] = "[TEST EMAIL] ADNOC Internship: Please Submit Your ID and Passport"
    msg['From'] = EMAIL_USER
    msg['To'] = to_email

    msg.set_content(f"""
        Hi,
        
        This is a reminder to submit your Emirates ID and Passport (both PDFs) by replying to this email with attachments.
        
        Any replies not in this thread will not be considered.
        
        Best regards,  
        HR Department @ ADNOC
    """)

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_USER, EMAIL_PASSWORD)
            smtp.send_message(msg)
    except Exception as e:
        print(f"Failed to send to {to_email}: {e}")

@app.route('/fetch-replies', methods=['POST'])
def fetch_replies():
    try:
        result = fetch_and_save_attachments()
        return jsonify({"status": "success", "message": result})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

@app.route("/")
def main():
    return render_template("main.html")

if __name__ == "__main__":
    app.run(debug=True)