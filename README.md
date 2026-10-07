# project-tracker

Project tracker for Suhail's digital marketing team at Dabur International (MENA and SSA, with support for India, Egypt, Turkey (Hobby) and Pakistan).

The live tracker is a Google Sheet. This repo keeps a copy of it and the script that builds it.

## Files

- `google-sheet/Dabur_Digital_Tracker.xlsx`: the workbook as uploaded to Google Drive. Upload it to Drive (it converts to a Google Sheet) to start a fresh copy.
- `google-sheet/build_sheet.py`: rebuilds the workbook from a JSON export of the old Dabur Digital Network hub.

## How the sheet is laid out

| Tab | Who edits | What it is |
|---|---|---|
| Start here | Suhail | How to use the tracker |
| Dashboard | nobody (formulas) | Counts per person and market, overdue/blocked list, due in 14 days, latest comments |
| All projects | nobody (formulas) | Every person's projects stacked into one list |
| Suhail, Nour, Sanya, Khushii, Team member 4 | that person + Suhail | One row per project |
| Comments | everyone | Updates, ideas, learnings, deck links, questions for Suhail |
| Lists | Suhail | Dropdown values: markets, brands, streams, agencies, people |

Edit rights are enforced with Google Sheets protected sheets (Data > Protect sheets and ranges), one per person's tab.

## Rebuilding

```
pip install openpyxl
python3 google-sheet/build_sheet.py <export-dir> Dabur_Digital_Tracker.xlsx
```

`<export-dir>` holds `people/`, `markets/`, `projects/` and `settings/main.json` exported from the hub.
