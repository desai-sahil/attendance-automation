# BioNB 2220 Attendance Tool

**Live Application:** https://bionb2220-attendance.streamlit.app/

A Streamlit-based web application designed to automate the attendance tracking process for BioNB 2220 at Cornell University.

This tool merges attendance reports from PollEverywhere into a Master Excel Sheet without disrupting existing data, formatting, or dates.

## Key Features

* **Updates:** Uses openpyxl to modify specific cells in the Master Sheet, preserving all other formatting, formulas, and headers.
* **Date Protection:** Automatically detects and skips the "Date Row" (Row 2) to prevent data corruption.
* **Manual Override Safety:** If a student is absent in the Poll report but has a manually entered "1" (Present) or note in the Master Sheet, the tool will not overwrite it. It only fills empty cells.
* **Zero-Installation:** Can be deployed to the web so TAs and staff can use it without installing Python.

## How the Algorithm Works

The application follows a strict "Safety First" logic to ensure data integrity:

1. Input Normalization
   - The app accepts a Master Excel File (.xlsx) and a PollEverywhere Report (.csv or .xlsx).
   - It accepts two user-defined strings:
     - Master Attendance sheet column name: The specific lecture column to update (e.g., "Lecture 2").
     - Poll search string in PollEV report: A keyword to find relevant columns in the poll report (e.g., "Lecture 2").

2. Poll Data Processing (In-Memory)
   - The script scans the PollEverywhere report for any column headers containing the Poll Search String.
   - It standardizes all student emails (lowercase, trimmed whitespace) to ensure accurate matching.
   - Logic: If a student has a non-empty value in any of the identified question columns, they are marked as "Present" in a temporary dictionary.

3. Surgical Merge
   - The script loads the Master Sheet using openpyxl.
   - It identifies the column index for "Email" and the target "Master Attendance sheet column name".
   - It iterates through the Master Sheet starting at Row 3 (skipping headers and dates).
   - The Update Rule:
     - If Student is in Poll Data: Mark as 1.
     - If Student is NOT in Poll Data AND Cell is Empty: Mark as 0.
     - If Student is NOT in Poll Data AND Cell has content: Do Nothing (Preserves manual edits).

4. Output
   - The modified file is saved to a memory buffer and offered as a download. The original file uploaded by the user remains untouched.

## Installation & Usage (Local)

If you prefer to run this tool on your own computer, follow these steps.

### Prerequisites
* Python 3.8 or higher
* pip (Python package installer)

### Step 1: Clone the Repository
git clone https://github.com/yourusername/attendance-automation.git
cd attendance-automation

### Step 2: Install Dependencies
pip install -r requirements.txt

### Step 3: Run the App
streamlit run attendance_app.py

A browser tab will automatically open at http://localhost:8501.

## Deployment (Web Version)

To make this tool accessible to the department without them needing to install Python:

1. Upload this code to a GitHub repository.
2. Sign up for Streamlit Community Cloud (https://streamlit.io/cloud).
3. Connect your GitHub account and select this repository.
4. Click "Deploy".

You will receive a permanent URL (e.g., https://cornell-bionb-attendance.streamlit.app) that you can share with other TAs or professors.

## Project Structure

attendance-automation/
├── attendance_app.py      # The main application code
├── requirements.txt       # List of python libraries (pandas, openpyxl, streamlit)
└── README.md              # Documentation and logic explanation

## Requirements

* streamlit
* pandas
* openpyxl
