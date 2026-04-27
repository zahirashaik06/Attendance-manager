# Attendance Manager (Python + Excel)

A beginner-friendly project that reads attendance data from an Excel file and generates a styled report.

## Files
- `attendance.xlsx` — sample input (10 students × 10 days, marked **P** / **A**)
- `attendance_manager.py` — the script
- `attendance_report.xlsx` — output (created when you run the script)

## Setup

```bash
pip install openpyxl
```

## Run

Put both files in the same folder, then:

```bash
python attendance_manager.py
```

## What it does
1. Reads each student's daily attendance (P = Present, A = Absent).
2. Calculates **Present count, Absent count, Attendance %**.
3. Flags students below **75%** as *Low Attendance* (highlighted orange).
4. Writes a styled **Summary** sheet with class stats (total students, average %, etc.).

## Customize
- Change `THRESHOLD = 75.0` in the script to adjust the low-attendance cutoff.
- Add more students/days to `attendance.xlsx` — the script auto-detects them.
- Mark cells as `P` or `A` (case-insensitive).

## Concepts you'll learn
- Reading & writing Excel with **openpyxl**
- Iterating rows, computing aggregates
- Conditional formatting via cell fills
- Using Excel formulas from Python (`=AVERAGE(...)`)
