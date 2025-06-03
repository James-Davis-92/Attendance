import streamlit as st
import pdfplumber
import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict
import os
import json
import shutil
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment

# 📂 Configuration
data_folder = r"Attendance_Records"  # Relative path for deployment convenience
os.makedirs(data_folder, exist_ok=True)

always_included_default_file = os.path.join(data_folder, "always_included_default.json")
always_included_file = os.path.join(data_folder, "always_included.json")

# On startup, copy default if working JSON does not exist
if not os.path.exists(always_included_file):
    if os.path.exists(always_included_default_file):
        shutil.copy(always_included_default_file, always_included_file)
    else:
        with open(always_included_file, 'w') as f:
            json.dump([], f)

def load_always_included():
    with open(always_included_file, 'r') as f:
        return json.load(f)

def save_always_included(names_list):
    with open(always_included_file, 'w') as f:
        json.dump(names_list, f)

days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']

# 📊 Group words by row tolerance
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

# 📊 Extract table data from PDF
def extract_table_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        page = pdf.pages[0]
        words = sorted(page.extract_words(), key=lambda w: (w['top'], w['x0']))
        rows = group_words_to_rows(words)
        return [[w['text'] for w in sorted(row, key=lambda w: w['x0'])] for row in rows]

# 🎛️ Process PDF for attendance
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

# 📆 Extract date from filename
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

# 📈 Format and save Excel
def style_and_save(df, week_key, day_headers):
    filename = os.path.join(data_folder, f"attendance_{week_key}.xlsx")
    df.to_excel(filename, index=False)
    wb = load_workbook(filename)
    ws = wb.active

    fill_map = {
        'Y': PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid'),
        'L': PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid'),
        'A': PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid'),
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

# 📄 Streamlit Interface
st.title("📋 Unauthorised Absence Tracker")

# Load always included names
always_included_names = load_always_included()

st.sidebar.header("Add always-included names (Surname FirstName)")
new_name = st.sidebar.text_input("Enter a name to add:")
if st.sidebar.button("Add Name"):
    if new_name.strip() != "":
        if new_name not in always_included_names:
            always_included_names.append(new_name.strip())
            save_always_included(always_included_names)
            st.sidebar.success(f"Added {new_name}")
        else:
            st.sidebar.warning(f"{new_name} already in the list.")
    else:
        st.sidebar.error("Please enter a valid name.")

if st.sidebar.button("Clear all always-included names"):
    if st.sidebar.confirm("Are you sure you want to clear all always-included names?"):
        always_included_names = []
        save_always_included(always_included_names)
        st.sidebar.success("Cleared all names.")

st.sidebar.markdown("### Current always-included names:")
for name in always_included_names:
    st.sidebar.write(name)

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

        # Add always-included names with 'A' if missing
        for full_name in always_included_names:
            parts = full_name.split(maxsplit=1)
            if len(parts) == 2:
                surname, first_name = parts
                key = (surname, first_name)
                if key not in all_attendance:
                    all_attendance[key] = {day: 'A' for day in days}

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
else:
    st.info("Upload PDF files to begin processing.")







