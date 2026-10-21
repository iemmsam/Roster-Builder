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

## Adding the leave plan and staff rules

Both can be **pasted from Excel**, so no file upload is needed (useful where uploads are blocked):

1. In Excel, select the table including its header row and press Ctrl+C.
2. In the app, open the **Leave plan** or **Staff & rules** tab, click in the paste box and press Ctrl+V.
3. Click **Use this leave plan** or **Use these rules**.

Uploading an .xlsx, .xls or .csv file is still available on the Leave plan tab.

### Leave plan layouts

- **Calendar grid:** one row per staff, names in the first column, dates or day numbers across a header row
  (1, 2, 3 … or 1-Nov, Mon 3, 03/11/2026). Any text in a cell (AL, MC, L) counts as leave.
- **List:** columns named Name, Start and End (or From/To), one row per leave block.

### Staff rules table

| Column | Meaning | Example |
|---|---|---|
| Name | Required | Aisyah |
| Exempt | Y = never rostered | Y |
| Not available on | Weekdays they can't do duty | Fri, or Mon, Wed |
| Max per week | Most duties in one week | 2 |
| Max per month | Most duties in the month | 8 |

Blank cells mean no limit. Pasting updates the people listed and adds anyone new.

## Rules the roster follows

1. Nobody is rostered on a leave day.
2. One slot per person per day.
3. Exempt staff, "not available" weekdays, and weekly and monthly maximums are respected.
4. Duties are spread evenly, and slot positions rotate.

More criteria go in `generate()` in `roster_engine.py`.

## Notes

- Changing the staff list, leave or month regenerates the roster and replaces manual edits. Download the Excel once the roster is final.
- Nothing is stored between sessions. Paste the leave plan and rules again when you reopen the app.
