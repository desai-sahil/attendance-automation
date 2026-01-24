import streamlit as st
import pandas as pd
import openpyxl
import io

# --- The Core Logic (Adapted for Web) ---
def process_attendance(master_file_obj, poll_file_obj, lecture_header, poll_search_string):
    """
    Reads files from memory, updates attendance, and returns a binary stream 
    of the saved Excel file.
    """
    # 1. Process Poll Data (Pandas)
    try:
        if poll_file_obj.name.endswith('.csv'):
            df_poll = pd.read_csv(poll_file_obj)
        else:
            df_poll = pd.read_excel(poll_file_obj)
    except Exception as e:
        return None, f"Error reading Poll file: {str(e)}"

    # Identify relevant question columns
    target_poll_cols = [col for col in df_poll.columns if poll_search_string.lower() in col.lower()]
    
    if not target_poll_cols:
        return None, f"Could not find any columns in Poll report matching '{poll_search_string}'"

    # Build Attendance Dictionary
    attendance_map = {}
    for _, row in df_poll.iterrows():
        raw_email = str(row.get('Email', '')).strip().lower()
        if '@' not in raw_email:
            continue
            
        answered_any = False
        for col in target_poll_cols:
            val = row[col]
            if pd.notna(val) and str(val).strip() != "":
                answered_any = True
                break
        
        if answered_any:
            attendance_map[raw_email] = 1

    # 2. Surgical Update (OpenPyXL)
    try:
        # Load the workbook from the uploaded file object
        wb = openpyxl.load_workbook(master_file_obj)
        ws = wb.active 
    except Exception as e:
        return None, f"Error reading Master Excel file: {str(e)}"

    # Locate Columns
    email_col_idx = None
    target_col_idx = None

    for cell in ws[1]:
        if cell.value == "Email":
            email_col_idx = cell.column
        elif cell.value == lecture_header:
            target_col_idx = cell.column

    if not email_col_idx:
        return None, "Column 'Email' not found in Master Sheet Row 1."
    if not target_col_idx:
        return None, f"Column '{lecture_header}' not found in Master Sheet Row 1."

    # Update Rows (Start at Row 3 to protect Date Row)
    updates_count = 0
    for row_num in range(3, ws.max_row + 1):
        cell_email = ws.cell(row=row_num, column=email_col_idx).value
        if not cell_email:
            continue
            
        master_email = str(cell_email).strip().lower()

        if master_email in attendance_map:
            ws.cell(row=row_num, column=target_col_idx).value = 1
            updates_count += 1
        else:
            # Only overwrite if empty (preserve manual edits)
            current_val = ws.cell(row=row_num, column=target_col_idx).value
            if current_val is None or current_val == "":
                ws.cell(row=row_num, column=target_col_idx).value = 0

    # 3. Save to Memory Buffer
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0) # Rewind buffer to start so it can be downloaded

    return output_buffer, f"Success. Updated {updates_count} students for '{lecture_header}'."

# --- The Streamlit Interface ---
st.set_page_config(page_title="BioNB 2220 Attendance Tool")

st.title("BioNB 2220 Attendance Tool")
st.markdown("""
This tool merges PollEverywhere reports into the Master Attendance Sheet.
""")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Files")
    master_file = st.file_uploader("Upload Master Excel (xlsx)", type=['xlsx'])
    poll_file = st.file_uploader("Upload Poll Report (csv/xlsx)", type=['csv', 'xlsx'])

with col2:
    st.subheader("2. Settings")
    lecture_header = st.text_input("Master Attendance sheet column name", placeholder="e.g., Lecture 2")
    poll_string = st.text_input("Poll search string in PollEV report", placeholder="e.g., Lecture 2")
    
    st.info(f"The tool will search for columns containing '{poll_string}' in the Poll Report and map them to '{lecture_header}' in the Master Sheet.")

st.divider()

if st.button("Process Attendance", type="primary"):
    if not master_file or not poll_file or not lecture_header or not poll_string:
        st.error("Please upload both files and fill in all text fields.")
    else:
        with st.spinner("Processing files..."):
            # Run the logic
            result_file, message = process_attendance(master_file, poll_file, lecture_header, poll_string)
            
            if result_file is None:
                st.error(f"{message}")
            else:
                st.success(f"{message}")
                
                # Create a smart filename for the download
                original_name = master_file.name.replace(".xlsx", "")
                new_filename = f"{original_name}_UPDATED.xlsx"
                
                st.download_button(
                    label="Download Updated Master Sheet",
                    data=result_file,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )