import streamlit as st
import pandas as pd
import docx
import json
import os

st.set_page_config(page_title="B'Cebu Class Schedule Finder", page_icon="📚", layout="wide")

st.title("🎓 B'Cebu Class Schedule System")
st.caption("API NEXT EDU, Inc. — Word Document Extractor & Search App")

DB_FILE = "schedules_data.json"

# Helper function to load data permanently
def load_schedules():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

# Helper function to save data permanently
def save_schedules(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# Load database into session state
if "schedules_db" not in st.session_state:
    st.session_state["schedules_db"] = load_schedules()

# Helper function to parse Word (.docx) files accurately
def parse_docx(file):
    doc = docx.Document(file)
    
    full_text_list = []
    
    # Read text from paragraphs
    for p in doc.paragraphs:
        if p.text.strip():
            full_text_list.append(p.text.strip())
            
    # Read text from tables
    table_rows_data = []
    for table in doc.tables:
        for row in table.rows:
            row_cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells if cell.text.strip()]
            # Remove consecutive duplicate text from merged cells
            cleaned_row = []
            for cell in row_cells:
                if not cleaned_row or cleaned_row[-1] != cell:
                    cleaned_row.append(cell)
            if cleaned_row:
                table_rows_data.append(cleaned_row)
                full_text_list.append(" | ".join(cleaned_row))

    full_text = " ".join(full_text_list)

    # Extract metadata flexibly
    name = "Unknown Student"
    nickname = "N/A"
    course = "N/A"
    duration = "N/A"
    nationality = "N/A"

    for row in table_rows_data:
        row_str = " ".join(row)
        if "Name:" in row_str or "Name" in row:
            for i, cell in enumerate(row):
                if "Name:" in cell and i + 1 < len(row):
                    name = row[i+1].replace("|", "").strip()
                if "Nickname:" in cell and i + 1 < len(row):
                    nickname = row[i+1].replace("|", "").strip()
                if "Course:" in cell and i + 1 < len(row):
                    course = row[i+1].replace("|", "").strip()
                if "Duration:" in cell and i + 1 < len(row):
                    duration = row[i+1].replace("|", "").strip()
                if "Nationality:" in cell and i + 1 < len(row):
                    nationality = row[i+1].replace("|", "").strip()

    student_data = {
        "Name": name,
        "Nickname": nickname,
        "Course": course,
        "Duration": duration,
        "Nationality": nationality,
        "Schedule": []
    }

    # Extract schedule periods (rows containing time ranges like 08:00 - 08:45)
    for row in table_rows_data:
        if len(row) >= 3 and ":" in row[0] and "-" in row[0]:
            student_data["Schedule"].append({
                "Time Period": row[0],
                "Period": row[1] if len(row) > 1 else "",
                "Room": row[2] if len(row) > 2 else "",
                "Teacher": row[3] if len(row) > 3 else "",
                "Type": row[4] if len(row) > 4 else "",
                "Subject / Book": row[5] if len(row) > 5 else (row[-1] if len(row) > 3 else "")
            })

    return student_data

# App Tabs
tab_search, tab_upload = st.tabs(["🔍 Search Student Schedules", "📤 Admin Upload Word Files (.docx)"])

# TAB 1: SEARCH INTERFACE
with tab_search:
    st.subheader("Search Schedules")
    
    # Show active students counter
    total_students = len(st.session_state["schedules_db"])
    st.caption(f"📊 Currently storing **{total_students}** student schedule(s).")
    
    search_query = st.text_input("Enter Student Full Name or Nickname (e.g. Emilia or LIN):", "").strip()

    if search_query:
        matches = [
            s for s in st.session_state["schedules_db"]
            if search_query.lower() in s["Name"].lower() or search_query.lower() in s["Nickname"].lower()
        ]

        if matches:
            for student in matches:
                with st.expander(f"📌 {student['Name']} ({student['Nickname']}) — {student['Course']}", expanded=True):
                    c1, c2 = st.columns(2)
                    c1.write(f"**Nationality:** {student['Nationality']}")
                    c2.write(f"**Duration:** {student['Duration']}")
                    st.divider()
                    
                    st.markdown("#### ⏰ Class Schedule")
                    if student["Schedule"]:
                        st.table(pd.DataFrame(student["Schedule"]))
                    else:
                        st.warning("No individual class periods detected in this document.")
        else:
            st.warning(f"No student found matching '{search_query}'.")
    else:
        st.info("💡 Type a student name or nickname above to view their schedule.")

# TAB 2: UPLOAD INTERFACE
with tab_upload:
    st.subheader("Upload Word Documents (.docx)")
    st.caption("Select one or multiple student schedule Word files to import them into the search system.")
    
    uploaded_files = st.file_uploader("Drop .docx files here", type=["docx"], accept_multiple_files=True)

    if uploaded_files:
        if st.button("Extract & Save Schedules"):
            count = 0
            for file in uploaded_files:
                try:
                    parsed_info = parse_docx(file)
                    
                    # Remove old version of student schedule if updating
                    st.session_state["schedules_db"] = [
                        s for s in st.session_state["schedules_db"] 
                        if s["Name"].lower() != parsed_info["Name"].lower()
                    ]
                    
                    st.session_state["schedules_db"].append(parsed_info)
                    count += 1
                except Exception as e:
                    st.error(f"Error reading {file.name}: {e}")

            # Save to permanent disk file
            save_schedules(st.session_state["schedules_db"])
            st.success(f"Successfully loaded and saved {count} schedule(s)! Switch to the Search tab to test it.")
