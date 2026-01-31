# Big Red Roll Call

**Live App:** https://bigredrollcall.streamlit.app/

**Big Red Roll Call** is a Streamlit web app that automates attendance tracking by merging **PollEverywhere** exports into a **Master Excel attendance sheet**—without breaking formatting, headers, or manual edits.

> **Generalized attendance rule (current version):**  
> A student is marked **present (1)** if their **email appears in the PollEverywhere export**, regardless of whether they answered any poll question.

---

## What This Tool Does

- You upload:
  - **Master Attendance Sheet** (`.xlsx`)
  - **PollEverywhere Export** (`.csv` or `.xlsx`)
- You select:
  - **Lecture number**
  - **Lecture date**
- The app:
  - **Creates the correct lecture column** in the master sheet (if missing)
  - Writes **1** for present and **0** for absent (safely)
  - **Appends new students** found in PollEverywhere but missing from the master list
  - Preserves existing formatting and manual overrides

---

## Key Features

- **Safety-first edits (surgical merge):** Uses `openpyxl` to update only the intended attendance cells—preserving everything else.
- **Automatic lecture column creation:**
  - **Row 1:** lecture date (e.g., `23-Jan`)
  - **Row 2:** `Lecture X` (capital L enforced)
  - **Row 3+:** attendance values
- **Manual override safety:**
  - If an attendance cell already contains a value (e.g., a manual `1`, note, or override), the app **does not overwrite it**.
  - Absences (`0`) are written **only if the attendance cell is blank**.
- **New student handling:**
  - If a student appears in the PollEverywhere export but isn’t in the master sheet, they are appended with:
    - Email
    - Full name (`First Last`)
    - Sortable name (`Last, First`)
- **Avoids “random far-right columns”:**
  - The app appends new lecture columns based on the **last real header column**, not Excel’s “max column” (which can be inflated by formatting).
- **Zero-install / shareable:** Works via a browser when deployed on Streamlit Cloud.

---

## Master Sheet Assumptions (Important)

This tool expects a master sheet layout like:

- **Row 1:** column headers + lecture date cells for lecture columns
- **Row 2:** lecture labels (e.g., `Lecture 1`, `Lecture 2`, …)
- **Row 3+:** student rows

And the roster columns include (case-insensitive match):
- `Full name`
- `Sortable name`
- `Email`
- (optional) `SIS Id`

> The tool matches students by **Email** (normalized to lowercase and trimmed).

---

## PollEverywhere Export Assumptions

The PollEverywhere export must include a column named **Email** (case-insensitive).

If available, the app also uses:
- `First name`
- `Last name`

to populate `Full name` and `Sortable name` when appending new students.

---

## Attendance Logic (Generalized)

**Present = 1** if the student’s email appears anywhere in the PollEverywhere export.

- This ignores “answered vs not answered”
- It treats “listed in export” as attendance

> **Note:** Depending on how PollEverywhere exports are configured, some export types may only include students who responded. If you need “present even if they never responded,” make sure you export a report that includes all participants/attendees.

---

## The Algorithm (Safety First)

1. **Inputs**
   - Master `.xlsx`
   - PollEverywhere `.csv`/`.xlsx`
   - Lecture number + lecture date

2. **Poll Processing**
   - Normalize emails
   - Build a presence set: any valid email in the export ⇒ present

3. **Master Sheet Merge**
   - Find roster columns (`Email`, `Full name`, `Sortable name`)
   - Ensure a lecture column exists (create if missing):
     - Row 1 = date
     - Row 2 = lecture label (`Lecture X`)
   - For each student in the master:
     - If present ⇒ write `1`
     - Else if attendance cell blank ⇒ write `0`
     - Else ⇒ do nothing (preserve manual edits)

4. **Append New Students**
   - Add students in Poll export but missing from master
   - Initialize their attendance for that lecture (`1` if present else `0`)

5. **Output**
   - Save to memory and return a downloadable updated `.xlsx`
   - Original upload remains unchanged

---

## Running Locally

### Prerequisites
- Python 3.9+ recommended
- pip

### Install
```bash
git clone https://github.com/yourusername/attendance-automation.git
cd attendance-automation
pip install -r requirements.txt
