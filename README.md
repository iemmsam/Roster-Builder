# Duty Roster Builder

Generates a monthly office duty roster: a set number of slots a day (8 by default, changeable in the sidebar), Monday to Friday, built around the staff leave plan.

## Files

| File | What it does |
|---|---|
| `app.py` | The Streamlit screens: upload, staff list, roster editor, review, fairness tally, Excel download |
| `roster_engine.py` | The logic: reads the leave Excel, generates the roster, checks for conflicts, builds the Excel export |
| `requirements.txt` | Python packages Streamlit Cloud installs |

## Run on your computer

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Put it online (Streamlit Community Cloud, free)

1. Create a GitHub repository and upload the three files above.
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **Create app**, pick the repository, set the main file to `app.py`, and deploy.
4. Share the app link with the office lead. To keep it private, set the app to private in its settings and invite the lead by email.

## Leave file layouts it reads

- **Calendar grid:** one row per staff, names in the first column, dates or day numbers (1–31) across a header row. Any text in a cell (AL, MC, L) counts as leave.
- **List:** columns named Name, Start and End (or From/To), one row per leave block. Dates like `3/11/2026`, `3 Nov 2026` or real Excel dates all work.

## Current rules (placeholders)

1. Nobody is rostered on a leave day.
2. One slot per person per day.
3. Duties are spread evenly, and slot positions rotate.

Office-specific criteria go in `generate()` in `roster_engine.py`.

## Notes

- Changing the staff list, leave or month regenerates the roster and replaces manual edits. Download the Excel once the roster is final.
- Nothing is stored between sessions. The leave file is uploaded each time you build a roster.
