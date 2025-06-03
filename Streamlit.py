import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# Use your uploaded JSON key file contents as a dict or load from a secure place
# For Streamlit Cloud, best is to add the JSON content as a secret (see below)
json_key = st.secrets["gcp_service_account"]

# Authenticate with Google
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = Credentials.from_service_account_info(json_key, scopes=scope)
client = gspread.authorize(creds)

# Open your Google Sheet by name or URL
sheet = client.open("AlwaysIncludedNames").sheet1

def get_always_included_names():
    records = sheet.col_values(1)  # get all values from first column
    if records and records[0].lower() == "name":
        return records[1:]  # exclude header
    return records

def add_name(name):
    names = get_always_included_names()
    if name and name not in names:
        sheet.append_row([name])

def remove_name(name):
    # Find and delete row with the name
    cell = sheet.find(name)
    if cell:
        sheet.delete_rows(cell.row)

# Streamlit UI

st.title("Always Included Names")

# Load current list
names = get_always_included_names()

st.write("### Current always-included names:")
for n in names:
    st.write(f"- {n}")

new_name = st.text_input("Add a new name (Format: Surname FirstName)")

if st.button("Add name"):
    if new_name.strip():
        add_name(new_name.strip())
        st.experimental_rerun()

del_name = st.text_input("Remove a name (exact match)")

if st.button("Remove name"):
    if del_name.strip():
        remove_name(del_name.strip())
        st.experimental_rerun()








