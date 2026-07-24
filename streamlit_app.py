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
    
    # Collect all text content from paragraphs and tables
    all_cells_text = []
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                clean_text = cell.text.strip().replace('\n', ' ')
                if clean_text:
                    all_cells_text.append(clean_text)

    # Combine all text into a single flat searchable string
    full_text = " | ".join(all_cells_text)

    # Smart regex fallbacks for B'Cebu header format
    name_search = re.search(r"Name:\s*\|?\s*([^|]+)", full_text, re.IGNORECASE)
    nickname_search = re.search(r"Nickname:\s*\|?\s*([^|]+)", full_text, re.IGNORECASE)
    course_search = re.search(r"Course:\s*\|?\s*([^|]+)", full_text, re.IGNORECASE)
    duration_search = re.search(r"Duration:\s*\|?\s*([^|]+)", full_text, re.IGNORECASE)
    nationality_search = re.search(r"Nationality:\s*\|?\s*([^|]+)", full_text, re.IGNORECASE)

    name = name_search.group(1).strip() if name_search else "Unknown Student"
    nickname = nickname_search.group(1).strip() if nickname_search else "N/A"
    course = course_search.group(1).strip() if course_search else "N/A"
    duration = duration_search.group(1).strip() if duration_search else "N/A"
    nationality = nationality_search.group(1).strip() if nationality_search else "N/A"

    student_data = {
        "Name": name,
        "Nickname": nickname,
        "Course": course,
        "Duration": duration,
        "Nationality": nationality,
        "Schedule": []
    }

    # Extract Schedule Rows from Tables
    for table in doc.tables:
        for row in table.rows:
            row_cells = []
            for cell in row.cells:
                text = cell.text.strip().replace('\n', ' ')
                # Remove repeated adjacent strings caused by cell merges
                if not row_cells or row_cells[-1] != text:
                    row_cells.append(text)
            
            # Identify schedule period rows starting with time formatting (e.g. 08:00 - 08:45)
            if len(row_cells) >= 3 and re.search(r"\d{2}:\d{2}\s*-\s*\d{2}:\d{2}", row_cells[0]):
                # Filter out blank elements
                filtered_row = [c for c in row_cells if c]
                student_data["Schedule"].append({
                    "Time Period": filtered_row[0] if len(filtered_row) > 0 else "-",
                    "Period": filtered_row[1] if len(filtered_row) > 1 else "-",
                    "Room": filtered_row[2] if len(filtered_row) > 2 else "-",
                    "Teacher": filtered_row[3] if len(filtered_row) > 3 else "-",
                    "Type": filtered_row[4] if len(filtered_row) > 4 else "-",
                    "Subject / Remarks": filtered_row[5] if len(filtered_row) > 5 else (filtered_row[-1] if len(filtered_row) > 3 else "-")
                })

    return student_data

# App Navigation
tab_search, tab_upload, tab_debug = st.tabs(["🔍 Search Student Schedules", "📤 Admin Upload Word Files (.docx)", "📋 Stored Records"])

# TAB 1: SEARCH INTERFACE
with tab_search:
    st.subheader("Search Schedules")
    
    total_students = len(st.session_state["schedules_db"])
    st.caption(f"📊 Currently storing **{total_students}** student schedule(s).")
    
    search_query = st.text_input("Enter Student Full Name or Nickname (e.g. Emilia or MURAKAMI):", "").strip()

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
            st.warning(f"No student found matching '{search_query}'. Check the 'Stored Records' tab to view all uploaded names.")
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

# TAB 3: DEBUG LIST
with tab_debug:
    st.subheader("List of All Stored Students")
    if st.session_state["schedules_db"]:
        student_list = [{"Full Name": s["Name"], "Nickname": s["Nickname"], "Course": s["Course"]} for s in st.session_state["schedules_db"]]
        st.table(pd.DataFrame(student_list))
        
        if st.button("Clear All Stored Schedules"):
            st.session_state["schedules_db"] = []
            save_schedules([])
            st.experimental_rerun()
    else:
        st.info("No records saved in database.")
