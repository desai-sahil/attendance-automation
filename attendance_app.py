import streamlit as st
import pandas as pd
import openpyxl
import io
from copy import copy

# --- Helpers ---
def _norm_email(x) -> str:
    return str(x or "").strip().lower()

def _best_effort_name(row: pd.Series) -> dict:
    """
    Try to infer a student's name from common PollEverywhere export columns.
    Returns dict with optional keys: first, last, full
    """
    # Common patterns: First Name / Last Name
    first = str(row.get("First Name", "") or "").strip()
    last = str(row.get("Last Name", "") or "").strip()

    # Sometimes a single Name / Full Name column exists
    full = ""
    for k in ["Name", "Full Name", "Participant", "Student", "User", "Respondent"]:
        if k in row.index:
            candidate = str(row.get(k, "") or "").strip()
            if candidate:
                full = candidate
                break

    # If first/last missing but full exists, keep full only.
    out = {}
    if first:
        out["first"] = first
    if last:
        out["last"] = last
    if full:
        out["full"] = full
    return out

def _find_header_col(ws, header_name: str):
    """Return 1-based column index for a header in row 1, else None."""
    for cell in ws[1]:
        if cell.value == header_name:
            return cell.column
    return None

def _copy_row_style(ws, src_row: int, dst_row: int, max_col: int):
    """
    Copy style/formatting from src_row to dst_row (best-effort).
    This helps the newly appended student row match the sheet formatting.
    """
    for col in range(1, max_col + 1):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        if src.has_style:
            dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)

def _last_row_with_email(ws, email_col_idx: int, start_row: int = 3) -> int:
    """Find the last row (>= start_row) that contains a non-empty email value."""
    last = start_row - 1
    for r in range(start_row, ws.max_row + 1):
        v = ws.cell(row=r, column=email_col_idx).value
        if v is not None and str(v).strip() != "":
            last = r
    return last


# --- Core Logic ---
def process_attendance(master_file_obj, poll_file_obj, lecture_header, poll_search_string):
    """
    Reads files from memory, updates attendance, appends new students found in PollEverywhere,
    and returns a binary stream of the saved Excel file.
    """
    # 1) Read Poll Data
    try:
        if poll_file_obj.name.endswith(".csv"):
            df_poll = pd.read_csv(poll_file_obj)
        else:
            df_poll = pd.read_excel(poll_file_obj)
    except Exception as e:
        return None, f"Error reading Poll file: {str(e)}"

    # Validate poll email column presence (best effort)
    # (Original code hard-coded 'Email'. We'll keep that, but give a clearer error.)
    if "Email" not in df_poll.columns:
        return None, "Poll report must contain a column named 'Email'."

    # Identify relevant question columns
    target_poll_cols = [col for col in df_poll.columns if poll_search_string.lower() in str(col).lower()]
    if not target_poll_cols:
        return None, f"Could not find any columns in Poll report matching '{poll_search_string}'"

    # Build:
    # - poll_students: email -> {'first','last','full'} (for roster reconciliation)
    # - attendance_map: email -> 1 if answered_any among target_poll_cols
    poll_students = {}
    attendance_map = {}

    for _, row in df_poll.iterrows():
        email = _norm_email(row.get("Email", ""))
        if "@" not in email:
            continue

        poll_students[email] = _best_effort_name(row)

        answered_any = False
        for col in target_poll_cols:
            val = row[col]
            if pd.notna(val) and str(val).strip() != "":
                answered_any = True
                break

        if answered_any:
            attendance_map[email] = 1

    if not poll_students:
        return None, "No valid student emails were found in the Poll report."

    # 2) Load Master Workbook
    try:
        wb = openpyxl.load_workbook(master_file_obj)
        ws = wb.active
    except Exception as e:
        return None, f"Error reading Master Excel file: {str(e)}"

    # Locate required columns in master
    email_col_idx = _find_header_col(ws, "Email")
    target_col_idx = _find_header_col(ws, lecture_header)

    if not email_col_idx:
        return None, "Column 'Email' not found in Master Sheet Row 1."
    if not target_col_idx:
        return None, f"Column '{lecture_header}' not found in Master Sheet Row 1."

    # Optional name columns in master (best-effort)
    first_col_idx = _find_header_col(ws, "First Name")
    last_col_idx = _find_header_col(ws, "Last Name")
    name_col_idx = _find_header_col(ws, "Name") or _find_header_col(ws, "Full Name")

    # Build a set of emails already in master
    master_emails = set()
    for row_num in range(3, ws.max_row + 1):  # preserve your "row 2 is dates" assumption
        cell_email = ws.cell(row=row_num, column=email_col_idx).value
        if cell_email is None or str(cell_email).strip() == "":
            continue
        master_emails.add(_norm_email(cell_email))

    # 3) Append NEW students from poll report that aren't in master
    # Find where to append
    last_student_row = _last_row_with_email(ws, email_col_idx, start_row=3)
    append_row = last_student_row + 1

    # Copy formatting from the last student row (if it exists), else from row 3
    style_src_row = last_student_row if last_student_row >= 3 else 3
    max_col = ws.max_column

    added_count = 0
    for email, nm in poll_students.items():
        if email in master_emails:
            continue

        # Create new row with formatting
        if style_src_row >= 3:
            _copy_row_style(ws, style_src_row, append_row, max_col)

        # Fill cells
        ws.cell(row=append_row, column=email_col_idx).value = email

        # Prefer explicit First/Last columns if present
        if first_col_idx and "first" in nm:
            ws.cell(row=append_row, column=first_col_idx).value = nm["first"]
        if last_col_idx and "last" in nm:
            ws.cell(row=append_row, column=last_col_idx).value = nm["last"]

        # If there's a "Name"/"Full Name" column, use it
        if name_col_idx:
            if "full" in nm and nm["full"].strip():
                ws.cell(row=append_row, column=name_col_idx).value = nm["full"]
            else:
                # If we only have first/last, synthesize full name
                first = nm.get("first", "").strip()
                last = nm.get("last", "").strip()
                full_guess = (first + " " + last).strip()
                if full_guess:
                    ws.cell(row=append_row, column=name_col_idx).value = full_guess

        # Initialize attendance value for this lecture:
        # - 1 if they answered any relevant poll question
        # - else 0 (since they appear on poll roster but didn't answer this set)
        ws.cell(row=append_row, column=target_col_idx).value = 1 if email in attendance_map else 0

        master_emails.add(email)
        append_row += 1
        added_count += 1

    # 4) Update attendance for existing rows (Start at Row 3 to protect Date Row)
    updates_count = 0
    for row_num in range(3, ws.max_row + 1):
        cell_email = ws.cell(row=row_num, column=email_col_idx).value
        if not cell_email:
            continue

        master_email = _norm_email(cell_email)

        if master_email in attendance_map:
            ws.cell(row=row_num, column=target_col_idx).value = 1
            updates_count += 1
        else:
            # Only overwrite if empty (preserve manual edits)
            current_val = ws.cell(row=row_num, column=target_col_idx).value
            if current_val is None or str(current_val).strip() == "":
                ws.cell(row=row_num, column=target_col_idx).value = 0

    # 5) Save to Memory Buffer
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    return (
        output_buffer,
        f"Success. Updated {updates_count} students for '{lecture_header}'. "
        f"Added {added_count} new students from Poll report."
    )


# --- Streamlit Interface ---
st.set_page_config(page_title="BioNB 2220 Attendance Tool")

st.title("BioNB 2220 Attendance Tool")
st.markdown(
    """
This tool merges PollEverywhere reports into the Master Attendance Sheet.

Created by: Sahil Desai
"""
)

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Files")
    master_file = st.file_uploader("Upload Master Excel (xlsx)", type=["xlsx"])
    poll_file = st.file_uploader("Upload Poll Report (csv/xlsx)", type=["csv", "xlsx"])

with col2:
    st.subheader("2. Settings")
    lecture_header = st.text_input("Master Attendance sheet column name", placeholder="e.g., Lecture 2")
    poll_string = st.text_input("Poll search string in PollEV report", placeholder="e.g., Lecture 2")
    st.info(
        f"The tool will search for columns containing '{poll_string}' in the Poll Report and map them to '{lecture_header}' in the Master Sheet."
    )

st.divider()

if st.button("Process Attendance", type="primary"):
    if not master_file or not poll_file or not lecture_header or not poll_string:
        st.error("Please upload both files and fill in all text fields.")
    else:
        with st.spinner("Processing files..."):
            result_file, message = process_attendance(master_file, poll_file, lecture_header, poll_string)

            if result_file is None:
                st.error(f"{message}")
            else:
                st.success(f"{message}")

                original_name = master_file.name.replace(".xlsx", "")
                new_filename = f"{original_name}_UPDATED.xlsx"

                st.download_button(
                    label="Download Updated Master Sheet",
                    data=result_file,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
