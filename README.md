# Work Tracker 2026

A local time-tracking app that logs hours directly into an Excel spreadsheet. Pick a category, run the timer, and your time gets written to the right cell for today automatically. No cloud, no accounts — just a Flask server running on your machine and an `.xlsx` file you can open anytime.

---

## Files

| File | Purpose |
|---|---|
| `app.py` | Flask backend — serves the UI and handles all timer/spreadsheet logic |
| `index.html` | Frontend UI — timer, category search, today's log |
| `Work Tracker 2026.xlsx` | The spreadsheet — one sheet, categories as rows, every day in 2026 as columns |
| `create_tracker.py` | Regenerates a fresh blank spreadsheet |
| `collapse_past_days.py` | Collapses all past-day columns in the spreadsheet |

---

## Setup

**Requirements:** Python 3.8+

Install dependencies:
```
pip install flask flask-cors openpyxl
```

Make sure all files are in the same folder, then run:
```
python app.py
```

The browser opens automatically at `http://localhost:5000`. Stop the server with `Ctrl+C`.

---

## How to use

1. Click the category input (or start typing) — existing categories appear as suggestions
2. Select one or type a new name — new categories are automatically added as a new row in the spreadsheet
3. Hit **Start**
4. Use **Pause / Resume** if you need to step away mid-session
5. Hit **Stop** — elapsed time gets added to today's cell and saved to the spreadsheet

**Today's log** at the bottom reads directly from the spreadsheet, sorted by time descending. It refreshes automatically every 30 seconds, or hit the `↻` button to force a refresh. Because it reads from the file, anything you manually type into the spreadsheet will show up here too.

---

## Spreadsheet structure

One sheet called `2026`. Layout:

- **Column A** — category names (row 2 onward)
- **Columns B → NJ** — one column per day, Jan 1 through Dec 31
- **Row 1** — header row with day abbreviation and date (e.g. `Thu\n3/26`)
- **Cell format** — `[h]:mm:ss` — the `[h]` bracket means hours can exceed 24 without rolling over to days

Past days are **grouped and collapsed** by default so the spreadsheet opens at today. Click the `+` button on the far left of the column headers to expand them. Run `collapse_past_days.py` periodically to re-collapse columns.

---

## Utilities

### Collapse past days
```
python collapse_past_days.py
```
Hides all columns before today using Excel's outline/grouping feature. Safe to re-run anytime — it also un-hides today and future columns if they were accidentally hidden.

### Regenerate spreadsheet
```
python create_tracker.py
```
Creates a fresh `Work Tracker Template.xlsx` with the default categories and past days collapsed. **This overwrites the existing file** — back it up first if you have data you want to keep.

---

## Default categories

- Classes
- Personal
- Coursework
- Employment
- UAV Lab
- Machine Learning
- Combat Clubs

Add new ones on the fly by just typing them into the search bar — no need to edit any files.

---

## Notes

- The server must be running for the UI to work — it's not a standalone webpage
- Time is stored as a decimal fraction of a day (Excel's native time format), so the cells are compatible with standard Excel `SUM` formulas
- If the server crashes mid-session, the in-progress time for that session is lost — it only writes to the spreadsheet when you hit Stop
- The timer state lives in memory, so restarting the server resets any active timer
