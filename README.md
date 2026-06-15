# Pole Survey Automation Tool

A web-based tool built during OJT that automates pole survey reporting — no more manual copy-pasting of images and data into Excel!

## The Problem
Field engineers export pole survey data as a CSV file. Each pole has 3 photos and several details that need to be manually copied into a formatted Excel report. This takes a lot of time and is very repetitive.

## The Solution
A web app where you just:
1. Fill in the project details
2. Upload the CSV file
3. Click Generate
4. Download the finished Excel report

The tool does everything automatically in the background.

---

## Features

### 📋 Report Generator
- Reads all pole data from the CSV
- Downloads all pole photos automatically (whole shot, attachments, tags)
- Opens the Excel template and fills in all text data
- Places and fits each photo into the correct cell
- Saves a finished Excel file ready for printing
- Live progress bar during generation
- Auto-fixes photo orientation (no more sideways images)
- Compressed images so the Excel file stays small
- Automatically names each sheet tab (e.g. FORTUNE 1, FORTUNE 2)
- Handles any number of poles — creates as many pages as needed
- Shareable on local office WiFi network

### 🧹 CSV Cleaner
A step-by-step wizard to clean and organize pole survey data before generating reports:

1. **Upload & Analyze** — detects issues automatically
2. **Pole Number Typo Fix** — detects values like NPT, NOTB, NONE and suggests corrections
3. **Spelling Fix** — detects similar barangay/municipality names using fuzzy matching
4. **Merge Barangays** — rename, merge, or reassign barangays to correct municipalities
5. **Merge Municipalities** — rename or merge municipality groups
6. **Duplicate Review** — flags duplicates by pole number, GPS, or photo URL. User decides what to keep — nothing is deleted automatically
7. **Pole Sequence Organizer** — reorder poles using #N item numbers from remarks field. Supports drag and drop, move up/down, renumber, and search
8. **Address Resolver** — reverse geocodes GPS coordinates using OpenStreetMap to generate formatted addresses like `ALONG FORTUNE ST, BESIDE MERCURY DRUG`
9. **Download** — export cleaned CSVs by barangay, municipality (ZIP with separate CSVs per barangay or single CSV), or all poles

**Other cleaner features:**
- Progress is saved automatically — leave the page and come back exactly where you left off
- Back button on every step
- Reset All clears everything completely
- Switch Original button on duplicates — choose which record to keep when the duplicate has better data

---

## Tech Stack
- **Python** — main programming language
- **Flask** — runs the web interface
- **win32com** — controls Excel to place images perfectly
- **Pillow** — handles image downloading and processing
- **thefuzz** — fuzzy string matching for spelling detection
- **geopy** — reverse geocoding via OpenStreetMap

---

## Requirements
- Windows (required for Excel generation via win32com)
- Python 3.11 recommended (3.14 may have compatibility issues with pywin32)

---

## Installation

pip install -r requirements.txt

Or manually:

pip install flask requests pillow thefuzz geopy pywin32

---

## How to Run

Place your `Template.xlsm` in the project folder then:

python app.py

Open the link shown in the terminal and share it with teammates on the same WiFi.

---

## File Structure

| File | Purpose |
|------|---------|
| `app.py` | Flask server, all routes |
| `excel_generator.py` | Excel report generation via win32com |
| `csv_cleaner.py` | All CSV cleaning and analysis logic |
| `address_resolver.py` | GPS reverse geocoding via OpenStreetMap |
| `templates/index.html` | Report Generator UI |
| `templates/cleaner.html` | CSV Cleaner UI (multi-step wizard) |
| `Template.xlsm` | Excel macro-enabled template |
| `requirements.txt` | Python dependencies |

---

## Developer
Nathan — OJT Project 2026
