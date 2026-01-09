
# 🧠 AutoHire: Intelligent Resume Processing & JD Matching System

AutoHire is an end-to-end automated resume screening and reporting system that fetches resumes from Gmail, detects duplicates, matches them against a job description (JD), generates an Excel report, and emails candidates and HR accordingly.

---

## 🚀 Features

- 📥 Fetch resumes from Gmail attachments
- 🧠 Extract resume content and calculate JD match using NLP (TF-IDF + cosine similarity)
- 🔁 Detect duplicate resumes based on email and phone
- ✅ Automatically assign status: Approved / Rejected / Pending
- 📊 Export detailed Excel report with charts, filters, and summaries
- 📧 Send custom status emails to candidates
- 📤 Email report to HR with one click
- 💾 MongoDB backend for storing evaluations

---

## 🛠 Technologies Used

- **Python 3.10+**
- **MongoDB**
- **pandas** & **xlsxwriter** for Excel reporting
- **pdfplumber** and **python-docx** for resume parsing
- **imaplib** for Gmail access
- **smtplib** for sending emails
- **dotenv** for credential management

---

## 📁 Project Structure

```
project_folder/
├── main.py                # Fetch resumes, score, and store
├── jd.py                  # JD extraction and scoring logic
├── duplicate.py           # Detect duplicate resumes
├── exl.py                 # Generate Excel reports
├── reply.py               # Send emails to candidates
├── emailer.py             # Email report to HR
├── run_all.py             # Automates the whole flow
├── .env                   # Stores credentials (not committed)
├── requirements.txt       # Python dependencies
├── processed_resumes.csv  # Tracks processed email IDs
└── downloads/
    └── jd.docx            # Job description document
```

---

## ⚙️ Setup Instructions

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Create `.env` File

```
EMAIL=your_gmail@gmail.com
PASSWORD=your_app_password
SENDER_EMAIL=your_gmail@gmail.com
SENDER_PASSWORD=your_app_password
HR_EMAIL=hr@example.com
```

> 🔐 Use a Gmail App Password if 2FA is enabled.

### 3. Place JD File

Save your JD as `downloads/jd.docx`.

### 4. Start MongoDB

```bash
mongod
```

Or use Docker:

```bash
docker run -d -p 27017:27017 --name mongodb mongo
```

---

## ▶️ Running the Project

### Option A: Manual Steps

```bash
python main.py       # Process resumes + score + store
python exl.py        # Export Excel report
python emailer.py    # Email report to HR
python reply.py      # Send candidate emails
```

### Option B: All-in-One

```bash
python run_all.py    # Runs everything above in one go
```

---

## 📦 Python Libraries Used

- pandas
- xlsxwriter
- pymongo
- pdfplumber
- python-docx
- imaplib2
- smtplib
- dotenv
- docx2pdf
- openpyxl
- pypandoc

---

## 📌 Outputs

- 📊 `Candidate_Report.xlsx`: With charts, summaries, and filtering
- 📧 Emails sent to candidates and HR
- 🧠 Candidate data stored in `resume_db.jd_evaluation`
- 🔁 Duplicates stored in `resume_db.duplicate_candidates`


