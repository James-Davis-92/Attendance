import streamlit as st
import pdfplumber
import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment

import gspread
from google.oauth2.service_account import Credentials

# 📂 Configuration
data_folder = r"C:\Users\james\PycharmProjects\GBEservices\.venv\Attendance_Records"
os.makedirs(data_folder, exist_ok=True)

days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']

# ----------- GOOGLE SHEETS SETUP ------------

# Google Sheets credentials from Streamlit secrets
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scope
)

client = gspread.authorize(creds)

# Replace with your actual Google Sheet name
SHEET_NAME = "AttendanceAlwaysIncluded"

try:
    sheet = client.open(SHEET_NAME).sheet1
except gspread.SpreadsheetNotFound:
    st.error(f"Google Sheet '{SHEET_NAME}' not found. Please create it and share with service account.")
    st.stop()

def get_always_included_names():
    # Returns list of names from first column
    try:
        data = sheet.col_values(1)
        return [name for name in data if name.strip()]
    except Exception as e:
        st.error(f"Error reading from Google Sheet: {e}")
        return []

def add_name(name):
    names = get_always_included_names()
    if name not in names:
        sheet.append_row([name])

def remove_name(name):
    names = get_always_included_names()
    if name in names:
        cell = sheet.find(name)
        if cell:
            sheet.delete_row(cell.row)

# ----------- PDF Processing functions ------------

def group_words_to_rows(words, tolerance=3):
    rows, current_row, last_top = [], [], None
    for w in words:
        top = w['top']
        if last_top is None or abs(top - last_top) <= tolerance:
            current_row.append(w)
            last_top = top if last_top is None else (last_top + top) / 2
        else:
            rows.append(current_row)
            current_row, last_top = [w], top
    if current_row:
        rows.append(current_row)
    return rows

def extract_table_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        page = pdf.pages[0]
        words = sorted(page.extract_words(), key=lambda w: (w['top'], w['x0']))
        rows = group_words_to_rows(words)
        return [[w['text'] for w in sorted(row, key=lambda w: w['x0'])] for row in rows]

def process_pdf(pdf_file):
    table = extract_table_from_pdf(pdf_file)
    filtered = [row for row in table if len(row) == 12 and row[0] == 'IMSL']
    attendance = {}
    cutoff_time = datetime.strptime("06:01:00", "%H:%M:%S").time()
    for row in filtered:
        surname, first_name, time_str, day_str = row[3].rstrip(','), row[4], row[8], row[6]
        time_obj = datetime.strptime(time_str, "%H:%M:%S").time()
        flag = 'Y' if time_obj < cutoff_time else 'L'
        attendance[(surname, first_name)] = (day_str, flag)
    return attendance

def extract_date_from_filename(filename):
    name, _ = os.path.splitext(os.path.basename(filename))
    for sep in ['_', '.']:
        parts = name.split(sep)
        if len(parts) >= 3:
            try:
                return datetime(int(parts[2]), int(parts[1]), int(parts[0]))
            except:
                continue
    return None

def style_and_save(df, week_key, day_headers):
    filename = os.path.join(data_folder, f"attendance_{week_key}.xlsx")
    df.to_excel(filename, index=False)
    wb = load_workbook(filename)
    ws = wb.active

    fill_map = {
        'Y': PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid'),  # Green
        'L': PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid'),  # Yellow
        'A': PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid'),  # Red
    }
    center_align = Alignment(horizontal='center', vertical='center')

    for col_idx in range(3, 8):
        ws.column_dimensions[chr(64 + col_idx)].width = 20

    for row in range(2, ws.max_row + 1):
        for col_idx in range(3, 8):
            cell = ws.cell(row=row, column=col_idx)
            if cell.value in fill_map:
                cell.fill = fill_map[cell.value]
            cell.alignment = center_align

    wb.save(filename)
    return filename

# ----------- Streamlit app UI ------------

st.title("📋 Unauthorised Absence Tracker")

# Load always-included names from Google Sheets
always_included_names = get_always_included_names()

# Section to add/remove always-included names
st.subheader("Manage always-included names")

new_name = st.text_input("Enter full name to add (Surname, FirstName)")

if st.button("Add Name"):
    if new_name.strip():
        add_name(new_name.strip())
        st.success(f"Added {new_name.strip()} to always-included list.")
        always_included_names = get_always_included_names()
    else:
        st.warning("Please enter a name.")

names_to_remove = st.multiselect("Select names to remove", always_included_names)

if st.button("Remove Selected Names"):
    for name in names_to_remove:
        remove_name(name)
    st.success("Removed selected names.")
    always_included_names = get_always_included_names()

st.write("### Current always-included names")
st.write(always_included_names)

uploaded_files = st.file_uploader("Upload attendance PDF(s)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    st.info(f"Processing {len(uploaded_files)} file(s)...")

    weeks = defaultdict(list)
    for file in uploaded_files:
        date = extract_date_from_filename(file.name)
        if date:
            year, week_num, _ = date.isocalendar()
            week_key = f"{year}-W{week_num:02d}"
            weeks[week_key].append((file, date))
        else:
            st.warning(f"Could not extract date from filename: {file.name}")

    for week_key, files in weeks.items():
        st.subheader(f"📅 Week {week_key}")
        all_attendance = defaultdict(lambda: {day: 'A' for day in days})
        year, week_str = map(int, week_key.split('-W'))
        monday_date = datetime.strptime(f"{year} {week_str} 1", "%G %V %u")

        day_headers = []
        for i, day_abbr in enumerate(days):
            current_date = monday_date + timedelta(days=i)
            day_headers.append(f"{day_abbr} {current_date.strftime('%d/%m/%Y')}")

        for file, _ in files:
            attendance_for_day = process_pdf(file)
            for (surname, first_name), (day_str, flag) in attendance_for_day.items():
                if day_str in days:
                    all_attendance[(surname, first_name)][day_str] = flag

        # Add always-included names with default 'A' flags if not present
        for full_name in always_included_names:
            try:
                surname, first_name = map(str.strip, full_name.split(',', 1))
            except Exception:
                continue
            if (surname, first_name) not in all_attendance:
                all_attendance[(surname, first_name)] = {day: 'A' for day in days}

        rows = []
        for (surname, first_name), day_flags in all_attendance.items():
            row = [surname, first_name] + [day_flags[day] for day in days]
            rows.append(row)

        df = pd.DataFrame(rows, columns=['Surname', 'FirstName'] + days)
        df = df.sort_values(by=['Surname', 'FirstName']).reset_index(drop=True)
        rename_map = {day: header for day, header in zip(days, day_headers)}
        df.rename(columns=rename_map, inplace=True)

        st.dataframe(df)

        filename = style_and_save(df, week_key, day_headers)
        st.success(f"✅ Updated attendance for week {week_key}.")
        with open(filename, "rb") as f:
            st.download_button("⬇️ Download updated Excel", f, file_name=os.path.basename(filename))

    st.info("Done processing all files.")








