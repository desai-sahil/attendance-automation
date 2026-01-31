import io
from copy import copy
from datetime import datetime, date

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

import pandas as pd
import streamlit as st


# =========================================================
# App constants
# =========================================================
APP_NAME = "Big Red Roll Call"
CREATOR_LINE = "Created by: Sahil Desai (desai.sahil97@gmail.com)"


# =========================================================
# Helpers
# =========================================================
def _norm_email(x) -> str:
    return str(x or "").strip().lower()


def _is_blank(x) -> bool:
    return x is None or str(x).strip() == ""


def _cell_text(x) -> str:
    return str(x or "").strip()


def _cell_text_lower(x) -> str:
    return _cell_text(x).lower()


def _find_header_col_ci(ws, header_name: str):
    """Case-insensitive header lookup in row 1. Returns 1-based col index or None."""
    target = str(header_name).strip().lower()
    for cell in ws[1]:
        if cell.value is None:
            continue
        if str(cell.value).strip().lower() == target:
            return cell.column
    return None


def _copy_row_style(ws, src_row: int, dst_row: int, max_col: int):
    """Copy style/formatting from src_row to dst_row (best-effort)."""
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
    """Find last row (>= start_row) with a non-empty email value."""
    last = start_row - 1
    for r in range(start_row, ws.max_row + 1):
        v = ws.cell(row=r, column=email_col_idx).value
        if not _is_blank(v):
            last = r
    return last


def _pd_get_ci(row: pd.Series, wanted: str, default=""):
    """Case-insensitive get for pandas Series by column name."""
    wanted_l = wanted.strip().lower()
    for c in row.index:
        if str(c).strip().lower() == wanted_l:
            v = row.get(c, default)
            if pd.isna(v):
                return default
            return v
    return default


def _make_full_and_sortable(first: str, last: str):
    first = str(first or "").strip()
    last = str(last or "").strip()
    full = (first + " " + last).strip()
    if last and first:
        sortable = f"{last}, {first}"
    elif last:
        sortable = last
    else:
        sortable = first
    return full, sortable


def _date_like_equal(a, b: date) -> bool:
    """Compare ws row1 cell (could be datetime/date/string) to a date object b."""
    if a is None:
        return False

    if isinstance(a, datetime):
        return a.date() == b
    if isinstance(a, date):
        return a == b

    s = str(a).strip()
    for fmt in (
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m-%d-%Y",
        "%Y-%m-%d",
        "%d-%b",
        "%d-%b-%Y",
        "%b %d, %Y",
    ):
        try:
            dt = datetime.strptime(s, fmt)
            if fmt == "%d-%b":
                return (dt.month, dt.day) == (b.month, b.day)
            return dt.date() == b
        except Exception:
            pass
    return False


def _set_column_width(ws, col_idx: int, width: float):
    ws.column_dimensions[get_column_letter(col_idx)].width = width


def _last_used_col_by_headers(ws, header_rows=(1, 2), min_col=1) -> int:
    """
    Returns the last column index that has any non-empty cell in the given header rows.
    Avoids ws.max_column being inflated by formatting.
    """
    last = min_col
    for col in range(min_col, ws.max_column + 1):
        for r in header_rows:
            if not _is_blank(ws.cell(row=r, column=col).value):
                last = col
                break
    return last


def _ensure_lecture_column(ws, lecture_label: str, lecture_date: date) -> int:
    """
    Ensure there is a column whose row2 equals lecture_label (case-insensitive).
    If none exists, append immediately after the last *actually used* header column.

    Enforces formatting:
      - row1: date with "d-mmm" format
      - row2: text
      - rows 3+: integer "0" format (prevents ##### and date rendering)
    """
    lecture_label_clean = lecture_label.strip()
    lecture_label_l = lecture_label_clean.lower()

    # Find existing matches in row 2
    matches = []
    for col in range(1, ws.max_column + 1):
        v2 = ws.cell(row=2, column=col).value
        if _cell_text_lower(v2) == lecture_label_l:
            matches.append(col)

    # Choose best match if multiple
    if len(matches) == 1:
        col_idx = matches[0]
    elif len(matches) > 1:
        col_idx = matches[0]
        for c in matches:
            v1 = ws.cell(row=1, column=c).value
            if _date_like_equal(v1, lecture_date):
                col_idx = c
                break
    else:
        # Append after last REAL header column (not ws.max_column)
        last_used = _last_used_col_by_headers(ws, header_rows=(1, 2), min_col=1)
        col_idx = last_used + 1

    # Write headers
    c1 = ws.cell(row=1, column=col_idx)
    c1.value = lecture_date
    c1.number_format = "d-mmm"  # e.g. 23-Jan

    c2 = ws.cell(row=2, column=col_idx)
    c2.value = lecture_label_clean  # ensures capital L

    _set_column_width(ws, col_idx, 12)

    # Force attendance format for the column
    for r in range(3, ws.max_row + 1):
        cell = ws.cell(row=r, column=col_idx)
        cell.number_format = "0"
        cell.alignment = Alignment(horizontal="center", vertical="center")

    return col_idx


# =========================================================
# Core logic (GENERALIZED PRESENCE)
# =========================================================
def process_attendance(
    master_file_obj,
    poll_file_obj,
    lecture_number: int,
    lecture_date: date,
    backfill_names_if_blank: bool = True,
):
    """
    GENERALIZED ATTENDANCE LOGIC:
      Present (1) if the student appears in the PollEverywhere export (i.e., their email is listed),
      regardless of whether they answered any question.

    Master sheet expectations:
      Row 1 = date headers for lecture columns (and normal headers for student info cols)
      Row 2 = lecture labels ("Lecture 1", "Lecture 2", ...)
      Row 3+ = student rows
      Student info headers include: Full name, Sortable name, Email (SIS Id optional)
    """
    lecture_label = f"Lecture {int(lecture_number)}"  # capital L enforced

    # --- Read Poll file ---
    try:
        if poll_file_obj.name.lower().endswith(".csv"):
            df_poll = pd.read_csv(poll_file_obj)
        else:
            df_poll = pd.read_excel(poll_file_obj)
    except Exception as e:
        return None, f"Error reading Poll file: {e}"

    # Require Email column in Poll (case-insensitive)
    if not any(str(c).strip().lower() == "email" for c in df_poll.columns):
        return None, "Poll report must contain a column named 'Email'."

    # Build poll roster and presence map:
    poll_students = {}   # email -> {"first","last","full","sortable"}
    presence_map = {}    # email -> 1 (present if listed)

    for _, row in df_poll.iterrows():
        email = _norm_email(_pd_get_ci(row, "Email", ""))
        if "@" not in email:
            continue

        first = str(_pd_get_ci(row, "First name", "") or "").strip()
        last = str(_pd_get_ci(row, "Last name", "") or "").strip()
        full, sortable = _make_full_and_sortable(first, last)

        poll_students[email] = {"first": first, "last": last, "full": full, "sortable": sortable}
        presence_map[email] = 1  # <-- key change: listed = present

    if not poll_students:
        return None, "No valid student emails were found in the Poll report."

    # --- Load Master workbook ---
    try:
        wb = openpyxl.load_workbook(master_file_obj)
        ws = wb.active
    except Exception as e:
        return None, f"Error reading Master Excel file: {e}"

    # Roster columns (row 1 headers)
    email_col_idx = _find_header_col_ci(ws, "Email")
    full_name_col_idx = _find_header_col_ci(ws, "Full name")
    sortable_name_col_idx = _find_header_col_ci(ws, "Sortable name")

    if not email_col_idx:
        return None, "Column 'Email' not found in Master Sheet (row 1)."

    # Ensure / create lecture column based on row2 label + row1 date
    lecture_col_idx = _ensure_lecture_column(ws, lecture_label, lecture_date)

    # Map existing master emails -> row
    master_email_to_row = {}
    for r in range(3, ws.max_row + 1):
        v = ws.cell(row=r, column=email_col_idx).value
        if _is_blank(v):
            continue
        master_email_to_row[_norm_email(v)] = r

    # --- Append NEW students missing from master ---
    last_student_row = _last_row_with_email(ws, email_col_idx, start_row=3)
    append_row = last_student_row + 1
    style_src_row = last_student_row if last_student_row >= 3 else 3
    max_col = ws.max_column

    added_count = 0
    for email, nm in poll_students.items():
        if email in master_email_to_row:
            continue

        if ws.max_row >= style_src_row:
            _copy_row_style(ws, style_src_row, append_row, max_col)

        ws.cell(row=append_row, column=email_col_idx).value = email

        if full_name_col_idx:
            ws.cell(row=append_row, column=full_name_col_idx).value = nm["full"]
        if sortable_name_col_idx:
            ws.cell(row=append_row, column=sortable_name_col_idx).value = nm["sortable"]

        # Initialize attendance for new student
        att_cell = ws.cell(row=append_row, column=lecture_col_idx)
        att_cell.value = 1 if email in presence_map else 0
        att_cell.number_format = "0"
        att_cell.alignment = Alignment(horizontal="center", vertical="center")

        master_email_to_row[email] = append_row
        append_row += 1
        added_count += 1

    # --- Update attendance for everyone in master ---
    updated_present_count = 0
    wrote_zero_count = 0
    backfilled_name_cells = 0

    for email, r in master_email_to_row.items():
        if email in presence_map:
            cell = ws.cell(row=r, column=lecture_col_idx)
            cell.value = 1
            cell.number_format = "0"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            updated_present_count += 1
        else:
            # Only write 0 if blank (preserve manual edits)
            cell = ws.cell(row=r, column=lecture_col_idx)
            if _is_blank(cell.value):
                cell.value = 0
                cell.number_format = "0"
                cell.alignment = Alignment(horizontal="center", vertical="center")
                wrote_zero_count += 1

        # Optional name backfill (only when blank in master)
        if backfill_names_if_blank and (email in poll_students):
            nm = poll_students[email]
            if full_name_col_idx:
                cur_full = ws.cell(row=r, column=full_name_col_idx).value
                if _is_blank(cur_full) and not _is_blank(nm["full"]):
                    ws.cell(row=r, column=full_name_col_idx).value = nm["full"]
                    backfilled_name_cells += 1
            if sortable_name_col_idx:
                cur_sort = ws.cell(row=r, column=sortable_name_col_idx).value
                if _is_blank(cur_sort) and not _is_blank(nm["sortable"]):
                    ws.cell(row=r, column=sortable_name_col_idx).value = nm["sortable"]
                    backfilled_name_cells += 1

    # --- Save result ---
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    msg = (
        f"Success. Used column '{lecture_label}' dated {lecture_date.strftime('%b %d, %Y')}. "
        f"Present=1 set for {updated_present_count} students (listed in Poll report). "
        f"Absent=0 written for {wrote_zero_count} blank cells. "
        f"Added {added_count} new students. "
    )
    if backfill_names_if_blank:
        msg += f"Backfilled {backfilled_name_cells} name cells."

    return output_buffer, msg


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title=APP_NAME)
st.title(APP_NAME)

st.markdown(
    f"""
This tool merges PollEverywhere reports into a Master Attendance Sheet.

**{CREATOR_LINE}**
"""
)

with st.expander("How to use", expanded=False):
    st.markdown(
        """
**Step 1 — Download your files**
- Export the **PollEverywhere** report as CSV (or XLSX).
- Download your **Master Attendance Sheet** as XLSX.

**Step 2 — Upload**
- Upload the Master Excel file
- Upload the Poll report

**Step 3 — Enter lecture info**
- Choose the **Lecture number**
- Choose the **Lecture date**

**Step 4 — Process & download**
- Click **Process Attendance**
- Download the updated master sheet

**Attendance rule (generalized)**
- A student is marked **present (1)** if their **email appears anywhere in the PollEverywhere export**.
- Absences (**0**) are written **only if the attendance cell is blank**, to preserve manual edits.
- Students in Poll but not in Master are **appended** (Email + Full name + Sortable name).
"""
    )

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1) Upload files")
    master_file = st.file_uploader("Upload Master Excel (xlsx)", type=["xlsx"])
    poll_file = st.file_uploader("Upload Poll report (csv/xlsx)", type=["csv", "xlsx"])

with col2:
    st.subheader("2) Lecture settings")
    lecture_number = st.number_input("Lecture number", min_value=1, step=1, value=1)
    lecture_date = st.date_input("Lecture date", value=date.today())

    backfill_names = st.checkbox(
        "Backfill names for existing students (only if blank in master)",
        value=True,
    )

lecture_label_preview = f"Lecture {int(lecture_number)}"
st.info(
    "Generalized attendance logic: a student is marked present (1) if their email appears in the "
    "PollEverywhere export (listed), regardless of answers. New students in the Poll report are appended."
)
st.write(
    f"**This run will write into:** `{lecture_label_preview}` with date `{lecture_date.strftime('%b %d, %Y')}`"
)

st.divider()

if st.button("Process Attendance", type="primary"):
    if not master_file or not poll_file:
        st.error("Please upload both files.")
    else:
        with st.spinner("Processing files..."):
            result_file, message = process_attendance(
                master_file_obj=master_file,
                poll_file_obj=poll_file,
                lecture_number=int(lecture_number),
                lecture_date=lecture_date,
                backfill_names_if_blank=backfill_names,
            )

        if result_file is None:
            st.error(message)
        else:
            st.success(message)

            base = master_file.name.rsplit(".", 1)[0]
            new_filename = f"{base}_UPDATED.xlsx"

            st.download_button(
                label="Download Updated Master Sheet",
                data=result_file,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

st.caption(
    "Tip: Keep your Master sheet roster headers consistent (Full name, Sortable name, Email). "
    "The app creates a new lecture column automatically if it doesn't exist."
)
