# Duty Roster Builder

Generates a monthly office duty roster, Monday to Friday, built around the staff leave plan. Each duty post
(MSW 1, MSW 2, MSW 7, MSW 8 by default) has an AM and a PM slot, so 8 slots a day. Posts can be changed in the sidebar.

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

- **Attendance grid (the office leave plan):** names down the side, dates across the top (1-Dec, 2-Dec …),
  a weekday row under the dates is fine. **1** = in office; **0**, **off** or blank = away. A date column
  blank for everyone (e.g. 25-Dec) is a public holiday with no duty. The # column, the Total Working days
  column and the total/strength rows at the bottom are ignored, so they can be included in the selection.
  The month is read from the dates and the app switches to it.
- **Leave-mark grid:** same shape, but blank = present and a mark (AL, MC, L) = on leave.
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

1. Nobody is rostered on a leave day or a public holiday.
2. Nobody takes two slots in the same half-day (AM or PM).
3. One person may take an AM and a PM slot on the same day. The sidebar setting decides when:
   - **Only when short of staff** (default): everyone available gets one slot first; people double up only
     to fill slots that would otherwise be empty.
   - **Allowed on any day**: doubling up is allowed whenever that person is the fairest pick.
   - **Never**: one slot per person per day; short days keep unfilled slots.
4. Exempt staff, "not available" weekdays, and weekly and monthly maximums are respected
   (an AM + PM day counts as 2 duties).
5. **Same slot on the same weekday** (on by default, switch in the sidebar): each weekday keeps one line-up
   all month, so whoever does AM MSW 1 on the first Monday does it every Monday. When that person is away,
   someone else covers that day only. In the roster and the Excel download, a person in their own weekday
   slot is shaded **blue** and a stand-in is shaded **yellow**.
6. Duties are spread evenly, and slot positions rotate.

More criteria go in `generate()` in `roster_engine.py`.

## Notes

- Changing the staff list, leave or month regenerates the roster and replaces manual edits. Download the Excel once the roster is final.
- Nothing is stored between sessions. Paste the leave plan and rules again when you reopen the app.
