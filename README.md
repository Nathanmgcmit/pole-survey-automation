# Pole Survey Automation System
A simple web-based tool I built during my OJT that automates the boring part of pole survey reporting — no more manual copy-pasting of images and data into Excel!

## The Problem
Every day, field engineers export pole survey data as a CSV file from the company website. Each pole has 3 photos and several details that need to be manually copied and pasted into a formatted Excel report. This takes a lot of time and is very repetitive.

## The Solution
I built a web app where you just:
1. Fill in the project details
2. Upload the CSV file
3. Click Generate
4. Download the finished Excel report

That's it! The tool does everything automatically in the background.

## What happens behind the scenes
- Reads all pole data from the CSV
- Downloads all pole photos from the website automatically
- Opens the Excel template
- Fills in all the text data (pole number, GPS coordinates, etc.)
- Places and fits each photo into the correct cell
- Saves a new Excel file ready for printing

## Tech Stack
- **Python** — main programming language
- **Flask** — runs the web interface
- **win32com** — controls Excel to place images perfectly
- **Pillow** — handles image downloading and processing

## How to Run

Install the required libraries:
```
pip install flask requests pillow pywin32
```

Place your `Template.xlsm` in the project folder then start the app:
```
python app.py
```

Open the link shown in the terminal and share it with your teammates on the same WiFi!

## Features
- ✅ Works on any browser — no installation needed for users
- ✅ Live progress bar so you know how far along it is
- ✅ Auto-fixes photo orientation (no more sideways images!)
- ✅ Compressed images so the Excel file stays small
- ✅ Automatically names each sheet tab (e.g. FORTUNE 1, FORTUNE 2)
- ✅ Handles any number of poles — creates as many pages as needed
- ✅ Shareable on local office WiFi network

## Developer
Nathan — OJT Project 2026
