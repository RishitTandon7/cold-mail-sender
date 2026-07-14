import json
import os
import smtplib
import ssl
import time
from datetime import datetime, timezone
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests
from openpyxl import load_workbook

EXCEL_PATH = os.getenv("EXCEL_PATH", "kaamkibaatein_HR_Contact_Database.xlsx")
SHEET_NAME = "HR Contacts"
CONTACT_CATEGORIES = [
    value.strip()
    for value in os.getenv("CONTACT_CATEGORIES", "MNC / Product & Funded Companies").split(",")
    if value.strip()
]

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASS = os.getenv("SMTP_PASS", "")

YOUR_NAME = "Rishit Tandon"
YOUR_PHONE = "+91 7394865520"
YOUR_LINKEDIN = "linkedin.com/in/rishit-tandon-928661287"
YOUR_PORTFOLIO = "portfolio.rishit.site"
YOUR_GITHUB = "github.com/RishitTandon7"
RESUME_PATH = os.getenv("RESUME_PATH", "Rishit-Tandon_Resume.pdf")

GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
GEMINI_API_KEY = next(
    (
        value
        for value in [
            os.getenv("GOOGLE_API_KEY"),
            os.getenv("GEMINI_API_KEY"),
            *[os.getenv(f"GOOGLE_API_KEY_{index}") for index in range(1, 17)],
        ]
        if value
    ),
    "",
)

SENT_LOG_PATH = "sent_log.json"
STATS_PATH = "stats.json"

BATCH_SIZE = int(os.getenv("BATCH_SIZE", "7"))
RUNS_PER_DAY = int(os.getenv("RUNS_PER_DAY", "2"))
DAILY_SEND_LIMIT = BATCH_SIZE * RUNS_PER_DAY
SEND_DELAY_SECONDS = int(os.getenv("SEND_DELAY_SECONDS", "20"))
DRY_RUN = os.getenv("DRY_RUN", "true").strip().lower() not in ("false", "0", "no")

SUBJECT_TEMPLATE = "Internship / Full-Time Opportunity at {company}"


def now_utc_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def count_sent_today(sent_log: dict) -> int:
    today = datetime.now(timezone.utc).date().isoformat()
    return sum(
        1
        for item in sent_log.values()
        if isinstance(item, dict) and str(item.get("sent_at", "")).startswith(today)
    )


def load_contacts() -> list[dict]:
    workbook = load_workbook(EXCEL_PATH, data_only=True)
    sheet = workbook[SHEET_NAME]
    headers = [cell.value for cell in sheet[1]]
    column = {name: index for index, name in enumerate(headers)}

    contacts = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name = row[column["Name"]]
        email = row[column["Email"]]
        category = (row[column["Category"]] or "").strip() if row[column["Category"]] else ""
        if not name or not email:
            continue
        if CONTACT_CATEGORIES and category not in CONTACT_CATEGORIES:
            continue
        contacts.append(
            {
                "name": str(name).strip(),
                "title": str(row[column["Title"]] or "").strip(),
                "company": str(row[column["Company"]] or "").strip(),
                "category": category,
                "email": str(email).strip(),
            }
        )
    return contacts


def load_sent_log() -> dict:
    if not os.path.exists(SENT_LOG_PATH):
        return {}
    with open(SENT_LOG_PATH, encoding="utf-8") as file:
        data = json.load(file)
    if isinstance(data, list):
        return {email: {"email": email} for email in data}
    return data


def save_sent_log(sent_log: dict) -> None:
    with open(SENT_LOG_PATH, "w", encoding="utf-8") as file:
        json.dump(sent_log, file, indent=2, sort_keys=True)


def record_sent_contact(sent_log: dict, contact: dict) -> None:
    sent_log[contact["email"]] = {
        "name": contact["name"],
        "title": contact["title"],
        "company": contact["company"],
        "email": contact["email"],
        "sent_at": now_utc_iso(),
    }
    save_sent_log(sent_log)


def write_stats(total_contacts: int, sent_log: dict, run_result: str) -> None:
    stats = {
        "generated_at": now_utc_iso(),
        "total_contacts": total_contacts,
        "total_sent": len(sent_log),
        "remaining": max(total_contacts - len(sent_log), 0),
        "sent_today": count_sent_today(sent_log),
        "batch_size": BATCH_SIZE,
        "runs_per_day": RUNS_PER_DAY,
        "daily_send_limit": DAILY_SEND_LIMIT,
        "dry_run": DRY_RUN,
        "contact_categories": CONTACT_CATEGORIES,
        "last_run_result": run_result,
    }
    with open(STATS_PATH, "w", encoding="utf-8") as file:
        json.dump(stats, file, indent=2)


def get_personalized_opener(contact: dict) -> str:
    if not GEMINI_API_KEY:
        return f"I wanted to reach out regarding opportunities at {contact['company']}."

    prompt = (
        f"Write one short professional opening line for a cold email to {contact['name']}, "
        f"{contact['title']} at {contact['company']}. Return only the line."
    )
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.6, "maxOutputTokens": 60},
    }

    try:
        response = requests.post(url, json=payload, timeout=30)
        response.raise_for_status()
        data = response.json()
        candidates = data.get("candidates", [])
        if candidates:
            parts = candidates[0].get("content", {}).get("parts", [])
            text = "".join(part.get("text", "") for part in parts).strip().strip('"')
            if text:
                return text
    except Exception as exc:
        print(f"Gemini fallback for {contact['email']}: {exc}")

    return f"I wanted to reach out regarding opportunities at {contact['company']}."


def build_email_body(contact: dict, opener: str) -> str:
    first_name = contact["name"].split()[0] if contact["name"] else "there"
    phone_line = f"\nPhone: {YOUR_PHONE}" if YOUR_PHONE else ""
    return f"""Hi {first_name},

{opener}

I'm {YOUR_NAME}, a final-year Computer Science student at SRM Institute of Science and Technology, focused on AI/ML and full-stack development.

I'm reaching out to ask whether there are any internship or full-time opportunities at {contact['company']} where my background could be a fit.

Portfolio: {YOUR_PORTFOLIO}
GitHub: {YOUR_GITHUB}
LinkedIn: {YOUR_LINKEDIN}{phone_line}

I've attached my resume and would really appreciate the chance to connect.

Thanks,
{YOUR_NAME}
"""


def send_email(contact: dict, subject: str, body: str) -> None:
    message = MIMEMultipart()
    message["From"] = SMTP_USER
    message["To"] = contact["email"]
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    if RESUME_PATH and os.path.exists(RESUME_PATH):
        with open(RESUME_PATH, "rb") as file:
            attachment = MIMEApplication(file.read(), Name=os.path.basename(RESUME_PATH))
        attachment["Content-Disposition"] = f'attachment; filename="{os.path.basename(RESUME_PATH)}"'
        message.attach(attachment)

    context = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls(context=context)
        server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(SMTP_USER, contact["email"], message.as_string())


def main() -> None:
    contacts = load_contacts()
    sent_log = load_sent_log()
    pending = [contact for contact in contacts if contact["email"] not in sent_log]
    sent_today = count_sent_today(sent_log)

    print(f"Total contacts: {len(contacts)}")
    print(f"Already sent: {len(sent_log)}")
    print(f"Remaining: {len(pending)}")
    print(f"Sent today: {sent_today}/{DAILY_SEND_LIMIT}")

    if not pending:
        print("All contacts are already sent.")
        write_stats(len(contacts), sent_log, "nothing_pending")
        return

    if not DRY_RUN and sent_today >= DAILY_SEND_LIMIT:
        print("Daily limit reached. Skipping this run.")
        write_stats(len(contacts), sent_log, "daily_limit_reached")
        return

    batch = pending[:BATCH_SIZE]
    sent_this_run = 0

    for index, contact in enumerate(batch, start=1):
        opener = get_personalized_opener(contact)
        subject = SUBJECT_TEMPLATE.format(company=contact["company"])
        body = build_email_body(contact, opener)

        print(f"\n[{index}/{len(batch)}] {contact['name']} | {contact['company']} | {contact['email']}")

        if DRY_RUN:
            print(subject)
            print(body)
            continue

        try:
            send_email(contact, subject, body)
            record_sent_contact(sent_log, contact)
            sent_this_run += 1
            print(f"Sent to {contact['email']}")
        except Exception as exc:
            print(f"Failed for {contact['email']}: {exc}")

        if index < len(batch):
            time.sleep(SEND_DELAY_SECONDS)

    result = "dry_run" if DRY_RUN else f"sent_{sent_this_run}"
    write_stats(len(contacts), sent_log, result)
    print(f"\nRun finished. Sent this run: {sent_this_run}")


if __name__ == "__main__":
    main()
