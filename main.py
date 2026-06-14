import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re
import os
import json
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from gspread.exceptions import APIError
from gspread_formatting import CellFormat, TextFormat, Color, format_cell_range

# ────────────────────────────────────────────────
# Config (Secured via Environment Variables)
# ────────────────────────────────────────────────

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive'
]

# Paths and IDs loaded from environment variables
CREDS_FILE = os.getenv('GOOGLE_CREDS_FILE', 'GOOGLE_CREDENTIALS')
SHEET1_ID = os.getenv('GOOGLE_SHEET1_ID')
SHEET2_ID = os.getenv('GOOGLE_SHEET2_ID')
WORKSHEET2_NAME = os.getenv('GOOGLE_WORKSHEET2_NAME', 'Inspections')

LAST_RUN_FILE = 'last_inspection_timestamp.json'
RED = Color(red=1.0, green=0.0, blue=0.0)

# Email Configuration (from GitHub Secrets)
EMAIL_SENDER = os.getenv('EMAIL_SENDER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER')

# ────────────────────────────────────────────────
# Helpers
# ────────────────────────────────────────────────

def get_client():
    creds_json = os.getenv('GOOGLE_CREDENTIALS')
    if creds_json:
        info = json.loads(creds_json)
        return gspread.authorize(Credentials.from_service_account_info(info, scopes=SCOPES))
    else:
        return gspread.authorize(Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES))

def retry_api_call(func, max_attempts=6, base_delay=5):
    for attempt in range(1, max_attempts + 1):
        try:
            return func()
        except APIError as e:
            if '503' not in str(e):
                raise
            delay = base_delay * (2 ** (attempt - 1)) + random.uniform(0, 4)
            print(f"  [503] Retrying in {delay:.1f}s (attempt {attempt}/{max_attempts})")
            import time
            time.sleep(delay)
    raise Exception("Max retries exceeded for Google Sheets API call")

# ────────────────────────────────────────────────
# Dynamic Worksheet Name
# ────────────────────────────────────────────────

def get_current_worksheet_name():
    """Returns format like: 'June 26', 'July 26', etc. based on Month and Year"""
    now = datetime.now()
    month_name = now.strftime('%B')   # Full month name (e.g., June)
    year_short = now.strftime('%y')   # Two-digit year (e.g., 26)
    return f"{month_name} {year_short}"

# ────────────────────────────────────────────────
# Phone Normalization
# ────────────────────────────────────────────────

def normalize_phone(phone: str) -> str:
    if not phone:
        return ''
    digits = re.sub(r'\D', '', phone.strip())
    if not digits:
        return ''
    if digits.startswith('0') and len(digits) == 11:
        digits = '91' + digits[1:]
    elif len(digits) == 10:
        digits = '91' + digits
    return digits

# ────────────────────────────────────────────────
# Formatting Functions
# ────────────────────────────────────────────────

def mark_as_yes(sheet, row_num):
    try:
        sheet.update(f'F{row_num}', [['YES']], value_input_option='RAW')
        format_cell_range(sheet, f'F{row_num}:G{row_num}', CellFormat(
            textFormat=TextFormat(foregroundColor=None)
        ))
        print(f"   → Marked as YES at row {row_num}")
    except Exception as e:
        print(f"   Failed to mark YES at row {row_num}: {e}")

def mark_as_no(sheet, row_num):
    try:
        sheet.update(range_name=f'F{row_num}', values=[['NO']], value_input_option='RAW')
        format_cell_range(sheet, f'F{row_num}:G{row_num}', CellFormat(
            textFormat=TextFormat(foregroundColor=RED)
        ))
        print(f"   ❌ Marked as NO at row {row_num}")
    except Exception as e:
        print(f"   Failed to mark NO at row {row_num}: {e}")

# ────────────────────────────────────────────────
# Main Processing
# ────────────────────────────────────────────────

def process_new_inspections():
    if not SHEET1_ID or not SHEET2_ID:
        print("❌ Error: GOOGLE_SHEET1_ID or GOOGLE_SHEET2_ID environment variable is missing.")
        return None

    client = get_client()
    worksheet_name = get_current_worksheet_name() # Dynamically calculates "June 26"
 
    print(f"Processing worksheet: {worksheet_name}")

    try:
        sheet1 = client.open_by_key(SHEET1_ID).worksheet(worksheet_name)
        sheet2 = client.open_by_key(SHEET2_ID).worksheet(WORKSHEET2_NAME)
    except Exception as e:
        print(f"Failed to open worksheet {worksheet_name}: {e}")
        return None

    today = datetime.now().date()
    today_yyyy_mm_dd = today.strftime('%Y-%m-%d')
    today_dd_mm_yyyy = today.strftime('%d/%m/%Y')
    today_dd_mm_dash = today.strftime('%d-%m-%Y')

    # Read Main Sheet
    try:
        data1 = retry_api_call(lambda: sheet1.get_all_values())
        print(f"Read {len(data1)} rows from main sheet")
    except Exception as e:
        print(f"Failed to read main sheet: {e}")
        return None

    today_rows = {}  # norm_phone → (row_num, current_status)

    # 1. Look through the monthly sheet, but isolate ONLY today's rows
    for i, row in enumerate(data1[1:], start=2):
        if len(row) < 9:
            continue

        date_b = (row[1] if len(row) > 1 else '').strip()
        raw_phone = (row[8] if len(row) > 8 else '').strip()
        current_status = (row[5] if len(row) > 5 else '').strip().upper()

        is_today = False
        if date_b:
            if (today_yyyy_mm_dd in date_b or
                today_dd_mm_yyyy in date_b or
                today_dd_mm_dash in date_b):
                is_today = True

        if is_today:
            norm_phone = normalize_phone(raw_phone)
            if norm_phone:
                today_rows[norm_phone] = (i, current_status)
                print(f"   Found today's row {i} → Phone: {norm_phone} | Status: {current_status}")

    print(f"→ Found {len(today_rows)} officer(s) scheduled for today")

    # Read Inspections Sheet
    try:
        data2 = retry_api_call(lambda: sheet2.get_all_values())
    except Exception as e:
        print(f"Failed to read Inspections sheet: {e}")
        return None

    today_inspections = set()

    for row in data2[1:]:
        if len(row) < 4:
            continue
        date_str = (row[0] or '').strip()
        phone_raw = (row[3] or '').strip()

        if today_dd_mm_yyyy in date_str and phone_raw:
            norm_phone = normalize_phone(phone_raw)
            if norm_phone:
                today_inspections.add(norm_phone)

    print(f"→ Found {len(today_inspections)} inspection(s) completed today")

    # 2. Update exclusively today's rows (Safe from touching tomorrow's/future schedules)
    for norm_phone, (row_num, current_status) in today_rows.items():
        if norm_phone in today_inspections:
            if current_status != 'YES':
                mark_as_yes(sheet1, row_num)
        else:
            if current_status not in ['YES']:
                mark_as_no(sheet1, row_num)

    return datetime.now()

# ────────────────────────────────────────────────
# Monthly Report
# ────────────────────────────────────────────────

def send_monthly_report():
    if datetime.now().day != 1:
        return False

    print("Generating Monthly Report...")

    try:
        if not SHEET1_ID:
            print("❌ Error: GOOGLE_SHEET1_ID configuration is missing.")
            return False

        client = get_client()
        worksheet_name = get_current_worksheet_name()
        sheet1 = client.open_by_key(SHEET1_ID).worksheet(worksheet_name)
        data = sheet1.get_all_values()

        not_done = []
        for row in data[1:]:
            if len(row) > 5:
                status = row[5].strip().upper()
                if status == 'NO':
                    name = row[0] if len(row) > 0 else 'Unknown'
                    phone = row[8] if len(row) > 8 else 'N/A'
                    not_done.append(f"• {name} ({phone})")

        count = len(not_done)
        if count == 0:
            not_done.append("✅ All inspections were completed this month!")

        body = f"""Monthly Inspection Report - {datetime.now().strftime('%B %Y')}

Total Officers who did NOT complete inspection: {count}

Officers with NO inspection:
""" + "\n".join(not_done)

        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER
        msg['Subject'] = f"Monthly Inspection Report - {datetime.now().strftime('%B %Y')}"

        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()

        print("✅ Monthly report sent successfully!")
        return True

    except Exception as e:
        print(f"Failed to send monthly report: {e}")
        return False

# ────────────────────────────────────────────────
# Main Execution
# ────────────────────────────────────────────────

if __name__ == '__main__':
    print(f"🚀 Inspection Monitor Started - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
 
    process_new_inspections()
    send_monthly_report()
 
    print("✅ Workflow completed successfully.")
