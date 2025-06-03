import streamlit as st
import pdfplumber
import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict
from io import BytesIO
from openpyxl.styles import PatternFill, Alignment
import os

# Constants
days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
NAMES_STORAGE_FILE = "always_included_names.txt"

# --- Functions ---

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
    name, _ = filename.rsplit('.', 1)[0], None
    for sep in ['_', '.']:
        parts = name.split(sep)
        if len(parts) >= 3:
            try:
                return datetime(int(parts[2]), int(parts[1]), int(parts[0]))
            except:
                continue
    return None

def style_excel(df):
    with BytesIO() as output:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            wb = writer.book
            ws = wb.active

            fill_map = {
                'Y': PatternFill(start_color='FF00FF00', end_color='FF00FF00', fill_type='solid'),  # Green
                'L': PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid'),  # Yellow
                'A': PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid'),  # Red
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

        return output.getvalue()

def load_saved_names():
    if os.path.exists(NAMES_STORAGE_FILE):
        with open(NAMES_STORAGE_FILE, "r") as f:
            lines = f.read().strip().splitlines()
        saved = []
        for line in lines:
            if ',' in line:
                surname, first_name = map(str.strip, line.split(',', 1))
                if surname and first_name:
                    saved.append((surname, first_name))
        return saved
    return []

def save_names_to_file(names_list):
    with open(NAMES_STORAGE_FILE, "w") as f:
        for surname, first_name in names_list:
            f.write(f"{surname}, {first_name}\n")

# --- Streamlit UI ---

st.title("📋 Attendance Tracker")

# Load saved always-included names
always_include = load_saved_names()

# --- Labour List input and save (Standalone block) ---
st.subheader("👥 Labour List")

# Pre-fill text area with saved names
names_str = "\n".join([f"{s}, {f}" for s, f in always_include])

names_input = st.text_area(
    "Enter names (one per line) in the format: Surname, FirstName (Example: Smith, John)",
    value=names_str,
    height=150
)

if st.button("💾 Save names"):
    # Parse input
    new_names = []
    for line in names_input.splitlines():
        if ',' in line:
            surname, first_name = map(str.strip, line.split(',', 1))
            if surname and first_name:
                new_names.append((surname, first_name))
    save_names_to_file(new_names)
    st.success("Names saved successfully!")
    always_include = new_names  # update current list in app

# --- Weekly attendance Excel upload ---

uploaded_excel = st.file_uploader(
    "Upload existing weekly attendance Excel file (optional)",
    type=['xlsx']
)

uploaded_pdfs = st.file_uploader(
    "Upload attendance PDF(s) for the week",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_pdfs:
    weeks = defaultdict(list)
    for file in uploaded_pdfs:
        date = extract_date_from_filename(file.name)
        if date:
            year, week_num, _ = date.isocalendar()
            week_key = f"{year}-W{week_num:02d}"
            weeks[week_key].append((file, date))
        else:
            st.warning(f"Could not extract date from filename: {file.name}")

    for week_key, files in weeks.items():
        st.subheader(f"📅 Week {week_key}")

        if uploaded_excel:
            df_existing = pd.read_excel(uploaded_excel)
            st.write("Loaded existing attendance data:")
            st.dataframe(df_existing)
        else:
            all_attendance = defaultdict(lambda: {day: 'A' for day in days})

            for file, _ in files:
                attendance_for_day = process_pdf(file)
                for (surname, first_name), (day_str, flag) in attendance_for_day.items():
                    if day_str in days:
                        all_attendance[(surname, first_name)][day_str] = flag

            # Add always-included names with all 'A' if missing
            for name_tuple in always_include:
                if name_tuple not in all_attendance:
                    all_attendance[name_tuple] = {day: 'A' for day in days}

            rows = []
            for (surname, first_name), day_flags in all_attendance.items():
                row = [surname, first_name] + [day_flags[day] for day in days]
                rows.append(row)

            df_existing = pd.DataFrame(rows, columns=['Surname', 'FirstName'] + days)
            df_existing = df_existing.sort_values(by=['Surname', 'FirstName']).reset_index(drop=True)

        year, week_num = map(int, week_key.split('-W'))
        monday_date = datetime.strptime(f"{year} {week_num} 1", "%G %V %u")
        day_headers = []
        for i, day_abbr in enumerate(days):
            current_date = monday_date + timedelta(days=i)
            day_headers.append(f"{day_abbr} {current_date.strftime('%d/%m/%Y')}")
        rename_map = {day: header for day, header in zip(days, day_headers)}
        df_display = df_existing.rename(columns=rename_map)

        st.dataframe(df_display)

        excel_bytes = style_excel(df_existing)
        st.download_button(
            label=f"⬇️ Download updated Excel for week {week_key}",
            data=excel_bytes,
            file_name=f"attendance_{week_key}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload PDFs to process weekly attendance.")












