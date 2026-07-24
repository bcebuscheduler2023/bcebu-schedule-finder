import streamlit as st
import pandas as pd
import docx
import re

st.set_page_config(page_title="B'Cebu Class Schedule Finder", page_icon="📚", layout="wide")

st.title("🎓 B'Cebu Class Schedule System")
st.caption("API NEXT EDU, Inc. — Word Document Extractor & Search App")

# Initialize persistent session storage for schedules
if "schedules_db" not in st.session_state:
    st.session_state["schedules_db"] = []

# Helper function to extract schedule data from Word tables
def parse_docx(file):
    doc = docx.Document(file)
    
    # Extract text from paragraphs
    full_text = " ".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    # Extract text from tables if paragraphs are formatted in tables
    for table in doc.tables:
        for row in table.rows:
            full_text += " " + " ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])

    # Extract Metadata using pattern matching
    name_match = re.search(r"Name:\s*\|?\s*([A-Za-z,\-\s]+)", full_text, re.IGNORECASE)
    nickname_match = re.search(r"Nickname:\s*\|?\s*([A-Za-z0-9\-\s]+)", full_text, re.IGNORECASE)
    course_match = re.search(r"Course:\s*\|?\s*([A-Za-z0-9\-\s]+)", full_text, re.IGNORECASE)
    duration_match = re.search(r"Duration:\s*\|?\s*([A-Za-z0-9\-\s]+)", full_text, re.IGNORECASE)
    nationality_match = re.search(r"Nationality:\s*\|?\s*([A-Za-z0-9\-\s]+)", full_text, re.IGNORECASE)

    student_data = {
        "Name": name_match.group(1).strip() if name_match else "Unknown Student",
        "Nickname": nickname_match.group(1).strip() if nickname_match else "N/A",
        "Course": course_match.group(1).strip() if course_match else "N/A",
        "Duration": duration_match.group(1).strip() if duration_match else "N/A",
        "Nationality": nationality_match.group(1).strip() if nationality_match else "N/A",
        "Schedule": []
    }

    # Extract Schedule Rows from Tables
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            # Clean duplicate adjacent text in cells caused by merged cells
            cleaned_cells = []
            for cell in cells:
                if not cleaned_cells or cleaned_cells[-1] != cell:
                    cleaned_cells.append(cell)

            # Match rows starting with time periods (e.g. 08:00 - 08:45)
            if cleaned_cells and re.match(r"\d{2}:\d{2}\s*-\s*\d{2}:\d{2}", cleaned_cells[0]):
                student_data["Schedule"].append({
                    "Time": cleaned_cells[0],
                    "Period": cleaned_cells[1] if len(cleaned_cells) > 1 else "",
                    "Room": cleaned_cells[2] if len(cleaned_cells) > 2 else "",
                    "Teacher": cleaned_cells[3] if len(cleaned_cells) > 3 else "",
                    "Type": cleaned_cells[4] if len(cleaned_cells) > 4 else "",
                    "Subject / Remarks": cleaned_cells[5] if len(cleaned_cells) > 5 else ""
                })

    return student_data

# App Navigation
tab_search, tab_upload = st.tabs(["🔍 Search Student Schedules", "📤 Admin Upload Word Files (.docx)"])

# TAB 1: SEARCH INTERFACE
with tab_search:
    st.subheader("Search Schedules")
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
            st.warning(f"No student found matching '{search_query}'. Please ensure the document was uploaded in the Admin tab.")
    else:
        st.info("💡 Enter a student name or nickname above to view their schedule.")

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
                    
                    # Remove previous entry for this student if updating
                    st.session_state["schedules_db"] = [
                        s for s in st.session_state["schedules_db"] if s["Name"].lower() != parsed_info["Name"].lower()
                    ]
                    
                    st.session_state["schedules_db"].append(parsed_info)
                    count += 1
                except Exception as e:
                    st.error(f"Error reading {file.name}: {e}")

            st.success(f"Successfully loaded {count} schedule(s)! Switch to the Search tab to test it.")
