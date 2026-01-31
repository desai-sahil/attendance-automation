# Big Red Roll Call

**Live App:** https://bigredrollcall.streamlit.app/

**Big Red Roll Call** is a Streamlit web app that automates attendance tracking by merging **PollEverywhere** exports into a **Master Excel attendance sheet**—without breaking formatting, headers, or manual edits.

> **Attendance rule (generalized):**  
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
    - Full name (`First Last`) *(if available in the Poll export)*
    - Sortable name (`Last, First`) *(if available in the Poll export)*
- **Avoids “random far-right columns”:**
  - The app appends new lecture columns based on the **last real header column**, not Excel’s “max column” (which can be inflated by formatting).
- **Zero-install / shareable:** Works via a browser when deployed on Streamlit Cloud.

---

## Master Sheet Assumptions

This tool expects a master sheet layout like:

- **Row 1:** roster headers + lecture date cells for lecture columns
- **Row 2:** lecture labels (e.g., `Lecture 1`, `Lecture 2`, …)
- **Row 3+:** student rows

Roster headers (case-insensitive match):
- `Full name`
- `Sortable name`
- `Email`
- (optional) `SIS Id`

> The tool matches students by **Email** (normalized to lowercase and trimmed).

---

## PollEverywhere Export Requirements

Your PollEverywhere export must include a column named:

- **Email** (case-insensitive)

If the export includes:
- `First name`
- `Last name`

then the app can also populate `Full name` and `Sortable name` for newly added students.

> **Important:** Make sure your PollEverywhere export includes the students you consider “present.”  
> Some export types only include students who responded at least once.

---

## How Attendance Is Determined (Generalized)

- **Present = 1** if the student’s email appears anywhere in the PollEverywhere export.
- **Absent = 0** is written only when the attendance cell is blank (manual edits are preserved).

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
   - Locate roster columns (`Email`, `Full name`, `Sortable name`)
   - Ensure lecture column exists (create if missing):
     - Row 1 = date
     - Row 2 = lecture label (`Lecture X`)
   - Update attendance:
     - If present ⇒ write `1`
     - Else if attendance cell blank ⇒ write `0`
     - Else ⇒ do nothing (preserve manual edits)

4. **Append New Students**
   - Add students in Poll export but missing from master
   - Initialize their attendance for that lecture

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
