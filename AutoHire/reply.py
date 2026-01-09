from pymongo import MongoClient
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ssl
import os
import time
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
EMAIL = os.getenv("SENDER_EMAIL")
PASSWORD = os.getenv("SENDER_PASSWORD")

# MongoDB setup
client = MongoClient("mongodb://localhost:27017/")
db = client["resume_db"]
collection = db["jd_evaluation"]

# Email templates
status_templates = {
    "Approved": (
        "Dear {name},\n\n🎉 This email is just for testing purpose. "
        "Congratulations! You have been approved for the next round. "
        "We’ll be in touch shortly.\n\nBest,\nRecruitment Team"
    ),
    "Rejected": (
        "Dear {name},\n\nThis email is just for testing purpose. "
        "Thank you for your application. Unfortunately, you were not selected this time. "
        "We wish you the best in your future endeavors.\n\nSincerely,\nRecruitment Team"
    ),
    "Pending": (
        "Dear {name},\n\nThis email is just for testing purpose. "
        "Your application is currently under review. We appreciate your patience "
        "and will update you soon.\n\nBest,\nRecruitment Team"
    )
}

def send_status_emails():
    candidates = list(collection.find({"mail_sent": False}))
    if not candidates:
        print("✅ No pending emails.")
        return

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=20) as server:
            server.starttls(context=ssl.create_default_context())
            server.login(EMAIL, PASSWORD)

            for i, candidate in enumerate(candidates, 1):
                to_email = candidate.get("email")
                name = candidate.get("name", "Candidate")
                status = candidate.get("status")

                if status not in status_templates:
                    continue

                msg = MIMEMultipart()
                msg["From"] = EMAIL
                msg["To"] = to_email
                msg["Subject"] = f"Application Status: {status}"
                body = status_templates[status].format(name=name)
                msg.attach(MIMEText(body, "plain"))

                try:
                    server.send_message(msg)
                    collection.update_one({"_id": candidate["_id"]}, {"$set": {"mail_sent": True}})
                    print(f"[{i}/{len(candidates)}] 📧 Sent {status} mail to {to_email}")
                except Exception as e:
                    print(f"❌ Failed to send to {to_email}: {e}")
                time.sleep(1.5)  # Gmail safe delay
    except Exception as e:
        print(f"❗ SMTP setup failed: {e}")

if __name__ == "__main__":
    send_status_emails()
