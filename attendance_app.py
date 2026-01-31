import io
import os
from copy import copy
from datetime import datetime, date
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd
import requests
import streamlit as st


# =========================================================
# GitHub "Last updated" helper
# =========================================================
@st.cache_data(ttl=3600)  # cache for 1 hour (avoids rate-limit pain)
def get_github_last_updated_str() -> str | None:
    """
    Returns a formatted "Last updated" timestamp from the most recent commit
    that touched the configured file path in a GitHub repo.

    Configure via Streamlit Secrets (preferred) or environment variables:
      - GITHUB_OWNER
      - GITHUB_REPO
      - GITHUB_BRANCH
      - GITHUB_FILE_PATH
      - GITHUB_TOKEN (optional)
    """
    owner = st.secrets.get("GITHUB_OWNER", os.getenv("GITHUB_OWNER", "")).strip()
    repo = st.secrets.get("GITHUB_REPO", os.getenv("GITHUB_REPO", "")).strip()
    branch = st.secrets.get("GITHUB_BRANCH", os.getenv("GITHUB_BRANCH", "main")).strip()
    path = st.secrets.get("GITHUB_FILE_PATH", os.getenv("GITHUB_FILE_PATH", "")).strip()
    token = st.secrets.get("GITHUB_TOKEN", os.getenv("GITHUB_TOKEN", "")).strip()

    if not (owner and repo and path):
        return None

    url = f"https://api.github.com/repos/{owner}/{repo}/commits"
    params = {"path": path, "sha": branch, "per_page": 1}
    headers = {"Accept": "application/vnd.github+json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"

    try:
        r = requests.get(url, params=params, headers=headers, timeout=10)
        r.raise_for_status()
        data = r.json()
        if not data:
            return None

        iso_dt = data[0]["commit"]["committer"]["date"]  # e.g. 2026-01-28T15:04:05Z
        dt_utc = datetime.fromisoformat(iso_dt.replace("Z", "+00:00"))
        dt_local = dt_utc.astimezone(ZoneInfo("America/New_York"))
        return dt_local.strftime("%b %d, %Y %I:%M %p ET")
    except Exception:
        return None


# =========================================================
# Helpers
# =========================================================
def _norm_email(x) -> str:
    return str(x or "").strip().lower()


def _is_blank(x) -> bool:
    return x is None or str(x).strip() == ""


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


def _cell_text(x) -> str:
    return str(x or "").strip()


def _cell_text_lower(x) -> str:
    return _cell_text(x).lower()


def _date_like_equal(a, b: date) -> bool:
    """
    Compare ws row1 cell (could be datetime/date/string) to a date object b.
    """
    if a is None:
        return False

    if isinstance(a, datetime):
        return a.date() == b
    if isinstance(a, date):
        return a == b

    # If stored as string like "21-Jan" or "1/28/2026", try a few parses
    s = str(a).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%Y-%m-%d", "%d-%b", "%d-%b-%Y", "%b %d, %Y"):
        try:
            dt = datetime.strptime(s, fmt)
            # If format has no year (e.g. %d-%b), dt defaults to 1900; in that case compare month/day only
            if fmt == "%d-%b":
                return (dt.month, dt.day) == (b.month, b.day)
            return dt.date() == b
        except Exception:
            pass
    return False


def _ensure_lecture_column(ws, lecture_label: str, lecture_date: date) -> int:
    """
    Ensure there is a column whose row2 equals lecture_label (case-insensitive).
    If multiple matches, prefer one whose row1 matches lecture_date.
    If none, append new column at end and set row1=date, row2=lecture_label.

    Returns 1-based column index.
    """
    lecture_label_l = lecture_label.strip().lower()

    matches = []
    for col in range(1, ws.max_column + 1):
        v2 = ws.cell(row=2, column=col).value
        if _cell_text_lower(v2) == lecture_label_l:
            matches.append(col)

    if len(matches) == 1:
        col_idx = matches[0]
    elif len(matches) > 1:
        # Prefer matching date in row1
        preferred = None
        for col in matches:
            v1 = ws.cell(row=1, column=col).value
            if _date_like_equal(v1, lecture_date):
                preferred = col
                break
        col_idx = preferred if preferred is not None else matches[0]
    else:
        # Append new column
        col_idx = ws.max_column + 1

        # Try to copy style from previous column (optional)
        src_col = ws.max_column  # previous last col
        for r in range(1, ws.max_row + 1):
            src = ws.cell(row=r, column=src_col)
            dst = ws.cell(row=r, column=col_idx)
            if src.has_style:
                dst._style = copy(src._style)
            dst.number_format = src.number_format
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.protection = copy(src.protection

            )

    # Set header cells (row1=date, row2=Lecture X)
    c1 = ws.cell(row=1, column=col_idx)
    c1.value = lecture_date  # store as a true date
    c1.number_format = "d-mmm"  # display like 21-Jan (adjust if you prefer)

    c2 = ws.cell(row=2, column=col_idx)
    c2.value = lecture_label  # guarantees capital L

    return col_idx


# =========================================================
# Core logic
# =========================================================
def process_attendance(
    master_file_obj,
    poll_file_obj,
    lecture_number: int,
    lecture_date: date,
    poll_search_string: str,
    backfill_names_if_blank: bool = True,
):
    """
    Updates attendance for the specified lecture using the PollEverywhere export,
    and appends new students (email + names) found in Poll but missing from master.

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

    # Find relevant poll question columns by substring match
    target_poll_cols = [c for c in df_poll.columns if poll_search_string.lower() in str(c).lower()]
    if not target_poll_cols:
        return None, f"Could not find any Poll columns matching '{poll_search_string}'."

    # Build poll roster and attendance map
    poll_students = {}   # email -> {"first","last","full","sortable"}
    attendance_map = {}  # email -> 1 if answered any target column

    for _, row in df_poll.iterrows():
        email = _norm_email(_pd_get_ci(row, "Email", ""))
        if "@" not in email:
            continue

        first = str(_pd_get_ci(row, "First name", "") or "").strip()
        last = str(_pd_get_ci(row, "Last name", "") or "").strip()
        full, sortable = _make_full_and_sortable(first, last)

        poll_students[email] = {
            "first": first,
            "last": last,
            "full": full,
            "sortable": sortable,
        }

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

    # --- Load Master workbook ---
    try:
        wb = openpyxl.load_workbook(master_file_obj)
        ws = wb.active
    except Exception as e:
        return None, f"Error reading Master Excel file: {e}"

    # Master columns (case-insensitive) for roster fields
    email_col_idx = _find_header_col_ci(ws, "Email")
    full_name_col_idx = _find_header_col_ci(ws, "Full name")
    sortable_name_col_idx = _find_header_col_ci(ws, "Sortable name")

    if not email_col_idx:
        return None, "Column 'Email' not found in Master Sheet (row 1)."

    # Ensure / create lecture attendance column
    lecture_col_idx = _ensure_lecture_column(ws, lecture_label, lecture_date)

    # Map existing master emails to rows
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

        # Style new row like the last student row (keeps formatting consistent)
        if ws.max_row >= style_src_row:
            _copy_row_style(ws, style_src_row, append_row, max_col)

        # Fill Email
        ws.cell(row=append_row, column=email_col_idx).value = email

        # Fill Full name and Sortable name per your rule
        if full_name_col_idx:
            ws.cell(row=append_row, column=full_name_col_idx).value = nm["full"]
        if sortable_name_col_idx:
            ws.cell(row=append_row, column=sortable_name_col_idx).value = nm["sortable"]

        # Initialize lecture attendance for new student
        ws.cell(row=append_row, column=lecture_col_idx).value = 1 if email in attendance_map else 0

        master_email_to_row[email] = append_row
        append_row += 1
        added_count += 1

    # --- Update attendance for everyone in master ---
    updated_present_count = 0
    wrote_zero_count = 0
    backfilled_name_cells = 0

    for email, r in master_email_to_row.items():
        # Mark present
        if email in attendance_map:
            ws.cell(row=r, column=lecture_col_idx).value = 1
            updated_present_count += 1
        else:
            # Only write 0 if blank to preserve manual edits
            cur = ws.cell(row=r, column=lecture_col_idx).value
            if _is_blank(cur):
                ws.cell(row=r, column=lecture_col_idx).value = 0
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

    # --- Save result to buffer ---
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    msg = (
        f"Success. Used column '{lecture_label}' dated {lecture_date.strftime('%b %d, %Y')}. "
        f"Present=1 set for {updated_present_count} students. "
        f"Absent=0 written for {wrote_zero_count} blank cells. "
        f"Added {added_count} new students. "
    )
    if backfill_names_if_blank:
        msg += f"Backfilled {backfilled_name_cells} name cells."

    return output_buffer, msg


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="BioNB 2220 Attendance Tool")

st.title("BioNB 2220 Attendance Tool")

last_updated = get_github_last_updated_str()
last_updated_line = f"Last updated: {last_updated}" if last_updated else "Last updated: (not configured)"

st.markdown(
    f"""
This tool merges PollEverywhere reports into the Master Attendance Sheet.

Created by: Sahil Desai  
{last_updated_line}
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

    st.subheader("3) Poll matching")
    poll_string = st.text_input("Poll column search string", placeholder="e.g., Lecture 3")

    backfill_names = st.checkbox(
        "Backfill names for existing students (only if blank in master)",
        value=True,
    )

lecture_label_preview = f"Lecture {int(lecture_number)}"
st.info(
    "Attendance logic: a student is marked present (1) if they answered ANY poll question column "
    "that matches your search string. "
    "Lecture column is auto-created if missing: row 1 = date, row 2 = lecture label."
)
st.write(f"**This run will write into:** `{lecture_label_preview}` (capital L) with date `{lecture_date.strftime('%b %d, %Y')}`")

st.divider()

if st.button("Process Attendance", type="primary"):
    if not master_file or not poll_file or not poll_string:
        st.error("Please upload both files and fill in the Poll column search string.")
    else:
        with st.spinner("Processing files..."):
            result_file, message = process_attendance(
                master_file,
                poll_file,
                lecture_number=int(lecture_number),
                lecture_date=lecture_date,
                poll_search_string=poll_string,
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
    "GitHub Last Updated setup: add Secrets for GITHUB_OWNER, GITHUB_REPO, GITHUB_BRANCH, GITHUB_FILE_PATH "
    "(and optionally GITHUB_TOKEN)."
)
