import os
import imaplib
import email
from email.header import decode_header
from pymongo import MongoClient
from datetime import datetime
import zipfile
from jd import get_jd_text, calculate_score, extract_text_from_pdf, extract_text_from_docx
from duplicate import check_and_insert
from dotenv import load_dotenv
from exl import export_to_excel

load_dotenv()

EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER = "imap.gmail.com"
DOWNLOAD_FOLDER = "downloads"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# MongoDB setup
client = MongoClient("mongodb://localhost:27017/")
db = client["resume_db"]
collection_jd = db["jd_evaluation"]
all_resumes_collection = db["all_resumes"]

PROCESSED_LOG = "processed_resumes.csv"

def is_already_processed(msg_id):
    if not os.path.exists(PROCESSED_LOG):
        return False
    with open(PROCESSED_LOG, "r") as f:
        return msg_id in f.read().splitlines()

def log_processed(msg_id):
    with open(PROCESSED_LOG, "a") as f:
        f.write(msg_id + "\n")

def extract_text_from_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(file_path)
    elif ext == ".docx":
        return extract_text_from_docx(file_path)
    elif ext == ".txt":
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    return ""

def extract_resume_info(text, filename):
    import re
    lines = text.splitlines()
    name = next((line.strip() for line in lines if line.strip()), "Not Found")
    email_match = re.search(r'\b[\w\.-]+@[\w\.-]+\.\w+\b', text)
    phone_match = re.search(r'(\+91[\-\s]?)?\b\d{10}\b', text)

    return {
        "name": name,
        "email": email_match.group(0) if email_match else "Not Found",
        "phone": phone_match.group(0) if phone_match else "Not Found",
        "filename": filename
    }

def process_resume_file(file_path):
    text = extract_text_from_file(file_path)
    info = extract_resume_info(text, os.path.basename(file_path))

    # ✅ Store metadata in all_resumes
    basic_info = {
        "name": info["name"],
        "email": info["email"],
        "phone": info["phone"],
        "filename": info["filename"],
        "timestamp": datetime.utcnow()
    }
    all_resumes_collection.insert_one(info.copy())
    print(f"📁 Stored in all_resumes: {basic_info}")

    # ✅ Check for duplicate
    result = check_and_insert(info)
    if result == "Duplicate":
        print("⛔ Skipped JD matching for duplicate")
        return

    # ✅ Score using JD
    jd_text = get_jd_text()
    if not jd_text:
        print("⚠️ JD file not found. Skipping scoring.")
        percent_score = 0.0
    else:
        raw_score = calculate_score(text, jd_text)
        percent_score = round(raw_score * 100, 2)

    info["score"] = percent_score
    info["status"] = "Approved" if percent_score >= 30 else "Rejected"
    info["timestamp"] = datetime.utcnow()
    info["is_duplicate"] = False
    info["mail_sent"] = False

    collection_jd.insert_one(info)
    print(f"✅ Inserted in jd_evaluation: {info['email']} | Score: {info['score']}% | {info['status']}")

def fetch_unread_job_emails():
    print("📧 Connecting to Gmail...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    mail.select("inbox")

    _, messages = mail.search(None, 'UNSEEN')
    for mail_id in messages[0].split():
        _, data = mail.fetch(mail_id, "(RFC822)")
        for part in data:
            if isinstance(part, tuple):
                msg = email.message_from_bytes(part[1])
                msg_id = msg.get("Message-ID")
                if not msg_id or is_already_processed(msg_id):
                    continue

                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding or "utf-8")

                for part in msg.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    if "attachment" in str(part.get("Content-Disposition", "")):
                        filename = part.get_filename()
                        if not filename:
                            continue
                        filepath = os.path.join(DOWNLOAD_FOLDER, filename)
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))

                        if filepath.endswith(".zip"):
                            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                                zip_ref.extractall(DOWNLOAD_FOLDER)
                                for extracted_file in zip_ref.namelist():
                                    extracted_path = os.path.join(DOWNLOAD_FOLDER, extracted_file)
                                    if os.path.isfile(extracted_path):
                                        process_resume_file(extracted_path)
                        else:
                            process_resume_file(filepath)
                log_processed(msg_id)
    mail.logout()
    print("✅ Done processing emails.")

if __name__ == "__main__":
    fetch_unread_job_emails()
    export_to_excel()  # ✅ Export latest Excel
