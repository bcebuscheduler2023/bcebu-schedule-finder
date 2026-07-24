import streamlit as st
import pandas as pd
import docx
import json
import os
import re

st.set_page_config(page_title="B'Cebu Class Schedule Finder", page_icon="📚", layout="wide")

st.title("🎓 B'Cebu Class Schedule System")
st.caption("API NEXT EDU, Inc. — Word Document Extractor & Search App")

DB_FILE = "schedules_data.json"

def load_schedules():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_schedules(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if "schedules_db" not in st.session_state:
    st.session_state["schedules_db"] = load_schedules()

def parse_docx(file):
    doc = docx.Document(file)
    
    # 1. Flatten all unique non-empty cells in the document
    raw_cells = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip().replace('\n', ' ')
                if txt and (not raw_cells or raw_cells[-1] != txt):
                    raw_cells.append(txt)

    # 2. Extract metadata by finding label indices and fetching next value
    def get_value_after(label_list):
        for idx, text in enumerate(raw_cells):
            for label in label_list:
                if text.lower() == label.lower():
                    # Return next item if available and not another label
                    if idx + 1 < len(raw_cells):
                        val = raw_cells[idx + 1].replace("|", "").strip()
                        if val and not any(l in val.lower() for l in ["nickname:", "course:", "duration:", "nationality:"]):
                            return val
        return "N/A"

    student_name = get_value_after(["Name:", "Name"])
    student_nickname = get_value_after(["Nickname:", "Nickname"])
    student_course = get_value_after(["Course:", "Course"])
    student_duration = get_value_after(["Duration:", "Duration"])
    student_nationality = get_value_after(["Nationality:", "Nationality"])

    student_data = {
        "Name": student_name if student_name != "N/A" else "Unknown Student",
        "Nickname": student_nickname,
        "Course": student_course,
        "Duration": student_duration,
        "Nationality": student_nationality,
        "Schedule": []
    }

    # 3. Extract Schedule rows matching time ranges (e.g., 08:00 - 08:45)
    for table in doc.tables:
        for row in table.rows:
            row_cells = []
            for cell in row.cells:
                txt = cell.text.strip().replace('\n', ' ')
                if not row_cells or row_cells[-1] != txt:
                    row_cells.append(txt)
            
            # Check if row starts with a time format
            if len(row_cells) >= 3 and re.search(r"\d{2}:\d{2}\s*-\s*\d{2}:\d{2}", row_cells[0]):
                clean_row = [c for c in row_cells if c]
                student_data["Schedule"].append({
                    "Time Period": clean_row[0] if len(clean_row) > 0 else "-",
                    "Period": clean_row[1] if len(clean_row) > 1 else "-",
                    "Room": clean_row[2] if len(clean_row) > 2 else "-",
                    "Teacher": clean_row[3] if len(clean_row) > 3 else "-",
                    "Type": clean_row[4] if len(clean_row) > 4 else "-",
                    "Subject / Book": clean_row[5] if len(clean_row) > 5 else (clean_row[-1] if len(clean_row) > 3 else "-")
                })

    return student_data

# App Navigation
tab_search, tab_upload, tab_records = st.tabs(["🔍 Search Student Schedules", "📤 Admin Upload Word Files (.docx)", "📋 Stored Records"])

# TAB 1: SEARCH INTERFACE
with tab_search:
    st.subheader("Search Schedules")
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
            st.warning(f"No student found matching '{search_query}'. Check 'Stored Records' tab to confirm names.")
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
                    
                    # Remove duplicate entries if re-uploading
                    st.session_state["schedules_db"] = [
                        s for s in st.session_state["schedules_db"] 
                        if s["Name"].lower() != parsed_info["Name"].lower()
                    ]
                    
                    st.session_state["schedules_db"].append(parsed_info)
                    count += 1
                except Exception as e:
                    st.error(f"Error reading {file.name}: {e}")

            save_schedules(st.session_state["schedules_db"])
            st.success(f"Successfully loaded and saved {count} schedule(s)! Switch to the Search tab to test it.")

# TAB 3: RECORDS LIST
with tab_records:
    st.subheader("List of All Stored Students")
    if st.session_state["schedules_db"]:
        student_list = [{"Full Name": s["Name"], "Nickname": s["Nickname"], "Course": s["Course"]} for s in st.session_state["schedules_db"]]
        st.table(pd.DataFrame(student_list))
        
        if st.button("Clear All Stored Schedules"):
            st.session_state["schedules_db"] = []
            save_schedules([])
            st.rerun()
    else:
        st.info("No records saved in database.")
