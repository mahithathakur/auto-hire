from main import fetch_unread_job_emails
from exl import export_to_excel
from emailer import send_email_with_attachment
from reply import send_status_emails
import os
from dotenv import load_dotenv

load_dotenv()

def run_all():
    print("🚀 Starting full automation flow...")

    # Step 1: Process resumes from Gmail
    fetch_unread_job_emails()

    # Step 2: Generate Excel Report
    export_to_excel()

    # Step 3: Email the report to HR
    hr_email = os.getenv("HR_EMAIL")
    if hr_email:
        send_email_with_attachment(hr_email, "Candidate_Report.xlsx")
    else:
        print("⚠️ HR_EMAIL not found in .env file. Skipping HR email.")

    # Step 4: Notify candidates of their status
    send_status_emails()

    print("✅ All steps completed successfully!")

if __name__ == "__main__":
    run_all()





#python run_all.py        # 🚀 All-in-one automation