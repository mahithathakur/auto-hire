import os
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()

def send_email_with_attachment(to_email, file_path):
    msg = EmailMessage()
    msg['Subject'] = "Daily Resume Reports"
    msg['From'] = os.getenv("SENDER_EMAIL")
    msg['To'] = to_email
    msg.set_content(
    "Dear HR Team,\n\n"
    "Attached is the latest resume screening report generated today. "
    "It includes evaluated candidate profiles with JD matching scores and status.\n\n"
    "Please review the shortlisted candidates and let us know if further action is needed.\n\n"
    "Best regards,\n"
    "RecruitBot"
)

    with open(file_path, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=os.path.basename(file_path)
        )

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(os.getenv("SENDER_EMAIL"), os.getenv("SENDER_PASSWORD"))
        server.send_message(msg)
        print("📤 Report emailed to HR!")

if __name__ == "__main__":
    # Use the refreshed file generated at 11:55
    send_email_with_attachment(os.getenv("HR_EMAIL"), "Candidate_Report.xlsx")