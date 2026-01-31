# Big Red Roll Call

**Live App:** https://bigredrollcall.streamlit.app/

**Big Red Roll Call** is a Streamlit web app that automates attendance tracking by merging **PollEverywhere** participation reports into a **Master Excel attendance sheet**—without breaking formatting, headers, or manual edits.

---

## What This Tool Does

- You upload:
  - **Master Attendance Sheet** (`.xlsx`)
  - **PollEverywhere Report** (`.csv` or `.xlsx`)
- You select:
  - **Lecture number**
  - **Lecture date**
  - **Poll column search string** (defaults to `Lecture X`)
- The app:
  - **Creates the correct lecture column** in the master sheet (if missing)
  - Writes **1** for present and **0** for absent (safely)
  - **Appends new students** found in PollEverywhere but missing from the master list

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
  - If a student appears in PollEverywhere but isn’t in the master sheet, they are appended with:
    - Email
    - Full name (`First Last`)
    - Sortable name (`Last, First`)
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

## How Attendance Is Determined

**Present = 1** if the student has a non-empty value in **any PollEverywhere column whose header contains the search string**.

Example:
- Poll search string = `"Lecture 3"`
- The tool looks for columns like:
  - `"Lecture 3 - Question 1"`
  - `"Lecture 3 - Poll"`
  - `"Lecture 3: Attendance Check"`

If any of those fields are non-empty → student marked present.

---

## The Algorithm (Safety First)

1. **Inputs**
   - Master `.xlsx`
   - PollEverywhere `.csv`/`.xlsx`
   - Lecture number + lecture date
   - Poll search string

2. **Poll Processing**
   - Normalize emails
   - Identify matching poll columns via substring search
   - Mark students present if they answered any matching poll column

3. **Master Sheet Merge**
   - Find roster columns (`Email`, `Full name`, `Sortable name`)
   - Ensure a lecture column exists (create if missing) with:
     - Row 1 = date
     - Row 2 = lecture label
   - For each student in the master:
     - If present → write `1`
     - Else if attendance cell blank → write `0`
     - Else → do nothing (preserve manual edits)

4. **Append New Students**
   - Add missing students from PollEverywhere to the bottom of the master sheet
   - Initialize their attendance for this lecture (`1` if present else `0`)

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
