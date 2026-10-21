"""Scheduling logic for the duty roster: leave-file parsing, roster generation, checks and Excel export.

Kept separate from the Streamlit UI so office criteria can be added here without touching the app layout.
"""
from __future__ import annotations

import datetime as dt
import io
import random
import re

import pandas as pd

SLOTS = 8  # default; the app lets the user change it
MAX_SLOTS = 30


def slot_cols(n: int) -> list[str]:
    return [f"Slot {i + 1}" for i in range(n)]


def roster_slot_cols(roster: pd.DataFrame) -> list[str]:
    """Slot columns present in a roster, in order."""
    return [c for c in roster.columns if str(c).startswith("Slot ")]


SLOT_COLS = slot_cols(SLOTS)
MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
NAME_RE = re.compile(r"name|staff|officer|employee", re.I)
SKIP_RE = re.compile(r"^(day|days|date|dates|name|staff|total)$|^(total|daily|staff strength|attendance|sub-?total)\b",
                     re.I)

EXAMPLE_STAFF = ["Aisyah", "Benjamin", "Chloe", "Darren", "Farah", "Gopal",
                 "Hui Min", "Irfan", "Jasmine", "Kelvin", "Lakshmi", "Marcus"]


def example_leave(year: int, month: int) -> dict[str, set[dt.date]]:
    """Sample leave for the chosen month, used until a real file is uploaded."""
    days = workdays(year, month)
    pick = lambda idx: {days[i] for i in idx if i < len(days)}
    return {
        "Chloe": pick([5, 6, 7]),
        "Gopal": pick([10, 11, 12, 13, 14]),
        "Jasmine": pick([2]),
        "Marcus": pick([15, 16]),
        "Farah": pick([19]),
    }


# ---------- Leave file parsing ----------

def _is_blank(v) -> bool:
    if v is None:
        return True
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def _valid(y: int, m: int, d: int) -> dt.date | None:
    try:
        return dt.date(y, m, d)
    except ValueError:
        return None


def to_date(v, year: int, month: int) -> dt.date | None:
    """Turn an Excel cell into a date. Bare day numbers (1-31) are read as days of the chosen month."""
    if _is_blank(v):
        return None
    if isinstance(v, (pd.Timestamp, dt.datetime)):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if float(v).is_integer():
            n = int(v)
            if 1 <= n <= 31:
                return _valid(year, month, n)
            if 30000 < n < 80000:  # Excel serial date
                return dt.date(1899, 12, 30) + dt.timedelta(days=n)
        return None
    s = str(v).strip()
    s = re.sub(r"^(mon|tue|wed|thu|fri|sat|sun)[a-z]*[\s,.]*", "", s, flags=re.I)   # "Mon 3"
    s = re.sub(r"[\s,(]*(mon|tue|wed|thu|fri|sat|sun)[a-z]*\)?$", "", s, flags=re.I)  # "3 Mon", "3 (Mon)"
    if not s:
        return None
    if m := re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s):
        return _valid(int(m[1]), int(m[2]), int(m[3]))
    if m := re.match(r"^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})$", s):  # d/m/yyyy, Singapore style
        y = int(m[3]) + (2000 if int(m[3]) < 100 else 0)
        return _valid(y, int(m[2]), int(m[1]))
    if m := re.match(r"^(\d{1,2})[\s\-]*([A-Za-z]{3})[a-z]*[\s\-,]*(\d{2,4})?$", s):  # 3 Nov 2026
        if m[2].lower() not in MONTHS:
            return None
        y = int(m[3]) if m[3] else year
        y += 2000 if y < 100 else 0
        return _valid(y, MONTHS.index(m[2].lower()) + 1, int(m[1]))
    if m := re.match(r"^([A-Za-z]{3})[a-z]*[\s\-]*(\d{1,2})(?:[\s,\-]+(\d{2,4}))?$", s):  # Nov 3, 2026
        if m[1].lower() not in MONTHS:
            return None
        y = int(m[3]) if m[3] else year
        y += 2000 if y < 100 else 0
        return _valid(y, MONTHS.index(m[1].lower()) + 1, int(m[2]))
    if m := re.match(r"^(\d{1,2})[/.](\d{1,2})$", s):  # 3/11 (day/month, chosen year)
        return _valid(year, int(m[2]), int(m[1]))
    if re.fullmatch(r"\d{1,2}", s):
        return _valid(year, month, int(s))
    return None


def read_rows(file_name: str, data: bytes) -> list[list]:
    """Read the first sheet of an uploaded file as raw rows (no header guessing)."""
    buf = io.BytesIO(data)
    if file_name.lower().endswith(".csv"):
        df = pd.read_csv(buf, header=None, dtype=object)
    else:
        df = pd.read_excel(buf, header=None, sheet_name=0)
    df = df.astype(object).where(pd.notna(df), None)
    return df.values.tolist()


def rows_from_paste(text: str) -> list[list]:
    """Turn cells copied from Excel (tab-separated) or CSV text into raw rows."""
    import csv

    text = (text or "").strip("\r\n")
    if not text.strip():
        return []
    delim = "\t" if "\t" in text else ","
    rows = [[c.strip() or None for c in r] for r in csv.reader(io.StringIO(text), delimiter=delim)]
    rows = [r for r in rows if r] or [[None]]
    width = max(len(r) for r in rows)
    return [r + [None] * (width - len(r)) for r in rows]


def _cell(r: list, i: int):
    return r[i] if 0 <= i < len(r) else None


def _is_one(v) -> bool:
    if _is_blank(v):
        return False
    try:
        return float(str(v).strip()) == 1
    except ValueError:
        return False


def detect_month(rows: list[list], year: int) -> tuple[int, int] | None:
    """The month a calendar grid covers, read from headers like '1-Dec' or real dates. None for bare day numbers."""
    for head in rows[:12]:
        found = []
        for v in head:
            if _is_blank(v) or isinstance(v, (int, float)) or re.fullmatch(r"\s*\d{1,2}\s*", str(v)):
                continue
            if d := to_date(v, year, 1):
                found.append((d.year, d.month))
        if len(found) >= 5:
            return max(set(found), key=found.count)
    return None


def parse_leave(rows: list[list], year: int, month: int
                ) -> tuple[str, dict[str, set[dt.date]], set[dt.date]] | None:
    """Detect the layout and return (layout name, {staff: set of days away}, public holidays),
    or None if unrecognised.

    Layouts:
      - Attendance grid (the office leave plan): names down the side, dates across the top,
        1 = in office; 0, "off", or a blank cell = away. A date column that is blank for everyone
        is a public holiday.
      - Leave-mark grid: same shape, but cells are blank when present and marked (AL, MC, L) when on leave.
      - List: columns named Name, Start (or From/Date) and optionally End (or To/Until).
    """
    for hi, head in enumerate(rows[:12]):
        lower = ["" if _is_blank(v) else str(v).lower().strip() for v in head]

        name_idx = next((i for i, s in enumerate(lower) if NAME_RE.search(s)), -1)
        start_idx = next((i for i, s in enumerate(lower) if re.search(r"\b(start|from)\b", s)), -1)
        if start_idx < 0:
            start_idx = next((i for i, s in enumerate(lower)
                              if "date" in s and not re.search(r"end|to\b", s)), -1)
        end_idx = next((i for i, s in enumerate(lower) if re.search(r"\b(end|to|until)\b", s)), -1)

        if name_idx >= 0 and start_idx >= 0:
            leave: dict[str, set[dt.date]] = {}
            for r in rows[hi + 1:]:
                name = "" if _is_blank(_cell(r, name_idx)) else str(_cell(r, name_idx)).strip()
                if not name or SKIP_RE.match(name):
                    continue
                leave.setdefault(name, set())
                a = to_date(_cell(r, start_idx), year, month)
                if not a:
                    continue
                b = (to_date(_cell(r, end_idx), year, month) if end_idx >= 0 else None) or a
                d = a
                while d <= b and (d - a).days < 370:
                    leave[name].add(d)
                    d += dt.timedelta(days=1)
            if leave:
                return "List (Name / Start / End)", leave, set()

        date_cols = [(i, d) for i, v in enumerate(head) if (d := to_date(v, year, month))]
        if len(date_cols) >= 5:
            date_idx = {i for i, _ in date_cols}
            if name_idx < 0:
                name_idx = next((i for i, v in enumerate(head) if not _is_blank(v) and i not in date_idx), 0)
            people = []
            for r in rows[hi + 1:]:
                name = "" if _is_blank(_cell(r, name_idx)) else str(_cell(r, name_idx)).strip()
                if not name or SKIP_RE.search(name) or to_date(name, year, month):
                    continue
                people.append((name, [_cell(r, ci) for ci, _ in date_cols]))
            if not people:
                continue

            filled = [c for _, cells in people for c in cells if not _is_blank(c) and str(c).strip()]
            attendance = bool(filled) and sum(_is_one(c) for c in filled) / len(filled) >= 0.5
            holidays = set()
            if attendance and len(people) >= 2:
                for j, (_, d) in enumerate(date_cols):
                    if all(_is_blank(cells[j]) or not str(cells[j]).strip() for _, cells in people):
                        holidays.add(d)

            leave = {}
            for name, cells in people:
                away = leave.setdefault(name, set())
                for (_, d), cell in zip(date_cols, cells):
                    if d in holidays:
                        continue
                    blank = _is_blank(cell) or not str(cell).strip()
                    if attendance:
                        if not _is_one(cell):          # 0, off, AL, blank = not in office
                            away.add(d)
                    elif not blank and str(cell).strip() != "0":
                        away.add(d)
            layout = "Attendance grid (1 = in office)" if attendance else "Calendar grid (marks = leave)"
            return layout, leave, holidays
    return None


# ---------- Staff rules ----------

RULE_COLS = ["Name", "Exempt", "Not available on", "Max per week", "Max per month"]
DAY_ABBR = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]


def empty_rules(names: list[str]) -> pd.DataFrame:
    return pd.DataFrame({"Name": names, "Exempt": False, "Not available on": "",
                         "Max per week": None, "Max per month": None}, columns=RULE_COLS)


def parse_weekdays(text) -> set[int]:
    """'Mon, Fri' / 'Monday & Wednesday' / 'Tue/Thu' -> {0, 4} etc."""
    if _is_blank(text):
        return set()
    return {DAY_ABBR.index(w) for w in re.findall(r"[A-Za-z]{3}", str(text).lower()) if w in DAY_ABBR}


def _truthy(v) -> bool:
    if _is_blank(v):
        return False
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in ("y", "yes", "true", "1", "x", "exempt", "✓")


def _limit(v) -> int | None:
    if _is_blank(v) or str(v).strip() == "":
        return None
    try:
        n = int(float(str(v).strip()))
    except ValueError:
        return None
    return n if n >= 0 else None


def rules_from_rows(rows: list[list]) -> pd.DataFrame | None:
    """Read a pasted rules table. Needs a header row with a Name column; other columns are matched by keyword:
    exempt, not available / unavailable / days off, week, month."""
    for hi, head in enumerate(rows[:5]):
        lower = ["" if _is_blank(v) else str(v).lower().strip() for v in head]
        name_idx = next((i for i, s in enumerate(lower) if NAME_RE.search(s)), -1)
        if name_idx < 0:
            continue
        find = lambda pat: next((i for i, s in enumerate(lower) if i != name_idx and re.search(pat, s)), -1)
        ex_i = find(r"exempt|excluded?")
        na_i = find(r"not avail|unavail|days? off|cannot|can't")
        wk_i = find(r"week")
        mo_i = find(r"month")
        out = []
        for r in rows[hi + 1:]:
            name = "" if _is_blank(_cell(r, name_idx)) else str(_cell(r, name_idx)).strip()
            if not name or SKIP_RE.match(name):
                continue
            out.append({
                "Name": name,
                "Exempt": _truthy(_cell(r, ex_i)),
                "Not available on": "" if _is_blank(_cell(r, na_i)) else str(_cell(r, na_i)).strip(),
                "Max per week": _limit(_cell(r, wk_i)),
                "Max per month": _limit(_cell(r, mo_i)),
            })
        if out:
            return pd.DataFrame(out, columns=RULE_COLS)
    return None


def rules_dict(table: pd.DataFrame) -> dict[str, dict]:
    """Rules table -> {name: {exempt, off_days, max_week, max_month}}."""
    out = {}
    for _, r in table.iterrows():
        name = "" if _is_blank(r.get("Name")) else str(r["Name"]).strip()
        if not name:
            continue
        out[name] = {
            "exempt": _truthy(r.get("Exempt")),
            "off_days": parse_weekdays(r.get("Not available on")),
            "max_week": _limit(r.get("Max per week")),
            "max_month": _limit(r.get("Max per month")),
        }
    return out


def _week(d: dt.date) -> tuple[int, int]:
    return d.isocalendar()[:2]


# ---------- Roster generation ----------

def workdays(year: int, month: int, holidays: set[dt.date] | None = None) -> list[dt.date]:
    """Mon-Fri dates of the month, minus public holidays."""
    holidays = holidays or set()
    d, out = dt.date(year, month, 1), []
    while d.month == month:
        if d.weekday() < 5 and d not in holidays:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


def generate(staff: list[str], leave: dict[str, set[dt.date]], year: int, month: int, seed: int,
             slots: int = SLOTS, rules: dict[str, dict] | None = None,
             holidays: set[dt.date] | None = None) -> pd.DataFrame:
    """Fill the given number of duty slots per working day.

    Rules applied:
      1. Nobody is rostered on a leave day.
      2. One slot per person per day.
      3. Staff rules: exempt staff are never rostered; nobody is rostered on their "not available" weekdays;
         max duties per week and per month are respected.
      4. People with the fewest duties so far are picked first; slot positions rotate.
    Add further office criteria here (e.g. no consecutive days, seniority per slot).
    """
    rules = rules or {}
    week_count: dict[tuple, int] = {}

    def eligible(n: str, d: dt.date) -> bool:
        if d in leave.get(n, set()):
            return False
        r = rules.get(n)
        if not r:
            return True
        if r["exempt"] or d.weekday() in r["off_days"]:
            return False
        if r["max_month"] is not None and count[n] >= r["max_month"]:
            return False
        if r["max_week"] is not None and week_count.get((n, _week(d)), 0) >= r["max_week"]:
            return False
        return True

    rng = random.Random(seed)
    count = {n: 0 for n in staff}
    cols = slot_cols(slots)
    slot_count = {n: [0] * slots for n in staff}
    records = []
    for d in workdays(year, month, holidays):
        avail = [n for n in staff if eligible(n, d)]
        rng.shuffle(avail)
        avail.sort(key=lambda n: count[n])  # stable, so the shuffle breaks ties fairly
        chosen = avail[:slots]
        row = [None] * slots
        for s in range(slots):
            if not chosen:
                break
            chosen.sort(key=lambda n: slot_count[n][s])
            p = chosen.pop(0)
            row[s] = p
            count[p] += 1
            slot_count[p][s] += 1
            week_count[(p, _week(d))] = week_count.get((p, _week(d)), 0) + 1
        records.append({"Date": d, "Day": d.strftime("%a"), **dict(zip(cols, row))})
    return pd.DataFrame(records, columns=["Date", "Day", *cols])


# ---------- Checks ----------

def cell_status(roster: pd.DataFrame, leave: dict[str, set[dt.date]],
                rules: dict[str, dict] | None = None) -> pd.DataFrame:
    """Same shape as the slot columns: '', 'empty', 'leave', 'dup', 'exempt' or 'offday' for each cell."""
    rules = rules or {}
    cols = roster_slot_cols(roster)
    status = pd.DataFrame("", index=roster.index, columns=cols)
    for idx, row in roster.iterrows():
        names = [row[c] for c in cols]
        for c in cols:
            p = row[c]
            if _is_blank(p) or p == "":
                status.at[idx, c] = "empty"
            elif row["Date"] in leave.get(p, set()):
                status.at[idx, c] = "leave"
            elif names.count(p) > 1:
                status.at[idx, c] = "dup"
            elif rules.get(p, {}).get("exempt"):
                status.at[idx, c] = "exempt"
            elif pd.Timestamp(row["Date"]).weekday() in rules.get(p, {}).get("off_days", set()):
                status.at[idx, c] = "offday"
    return status


def issues_table(roster: pd.DataFrame, status: pd.DataFrame,
                 rules: dict[str, dict] | None = None) -> pd.DataFrame:
    labels = {"empty": "Unfilled slot", "leave": "Rostered while on leave", "dup": "Rostered twice that day",
              "exempt": "Staff is exempt from duty", "offday": "Not available on this weekday"}
    rows = []
    for idx, row in roster.iterrows():
        for c in roster_slot_cols(roster):
            kind = status.at[idx, c]
            if kind:
                rows.append({"Date": row["Date"], "Slot": c, "Staff": row[c] or "", "Issue": labels[kind]})
    rows += limit_breaches(roster, rules or {})
    return pd.DataFrame(rows, columns=["Date", "Slot", "Staff", "Issue"])


def limit_breaches(roster: pd.DataFrame, rules: dict[str, dict]) -> list[dict]:
    """Weekly and monthly maximums exceeded (usually after manual edits)."""
    cols = roster_slot_cols(roster)
    month_n: dict[str, int] = {}
    week_n: dict[tuple, list] = {}
    for _, row in roster.iterrows():
        d = pd.Timestamp(row["Date"]).date()
        for p in {row[c] for c in cols if not _is_blank(row[c]) and row[c] != ""}:
            month_n[p] = month_n.get(p, 0) + 1
            week_n.setdefault((p, _week(d)), []).append(d)
    out = []
    for p, n in month_n.items():
        cap = rules.get(p, {}).get("max_month")
        if cap is not None and n > cap:
            out.append({"Date": None, "Slot": "", "Staff": p, "Issue": f"{n} duties this month, limit is {cap}"})
    for (p, _), ds in week_n.items():
        cap = rules.get(p, {}).get("max_week")
        if cap is not None and len(ds) > cap:
            out.append({"Date": min(ds), "Slot": "", "Staff": p,
                        "Issue": f"{len(ds)} duties in the week of {min(ds):%d %b}, limit is {cap}"})
    return out


def tally(roster: pd.DataFrame, staff: list[str], leave: dict[str, set[dt.date]]) -> pd.DataFrame:
    days = list(roster["Date"])
    counts = {n: 0 for n in staff}
    for c in roster_slot_cols(roster):
        for p in roster[c]:
            if not _is_blank(p) and p != "":
                counts[p] = counts.get(p, 0) + 1
    rows = []
    for n, c in counts.items():
        lv = sum(1 for d in days if d in leave.get(n, set()))
        rows.append({"Staff": n, "Duties": c, "Leave days": lv, "Days available": len(days) - lv})
    return pd.DataFrame(rows).sort_values(["Duties", "Staff"], ascending=[False, True]).reset_index(drop=True)


# ---------- Excel export ----------

FILLS = {"leave": "FBE6E1", "dup": "FBEFD2", "empty": "EEF1F4", "exempt": "E8E1F5", "offday": "E8E1F5"}


def to_excel(roster: pd.DataFrame, status: pd.DataFrame, tally_df: pd.DataFrame,
             issues: pd.DataFrame, title: str) -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    cols = roster_slot_cols(roster)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        r = roster.copy()
        r["Date"] = pd.to_datetime(r["Date"])
        r.to_excel(xw, sheet_name="Roster", index=False, startrow=1)
        tally_df.to_excel(xw, sheet_name="Tally", index=False)
        issues.to_excel(xw, sheet_name="Issues", index=False)

        ws = xw.sheets["Roster"]
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=14)
        for cell in ws[2]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 6
        for i in range(len(cols)):
            ws.column_dimensions[get_column_letter(3 + i)].width = 14
        for ri, (_, srow) in enumerate(status.iterrows(), start=3):
            ws.cell(row=ri, column=1).number_format = "dd mmm yyyy"
            for ci, c in enumerate(cols, start=3):
                kind = srow[c]
                if kind:
                    ws.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor=FILLS[kind])
        ws.freeze_panes = "C3"

        for name in ("Tally", "Issues"):
            sh = xw.sheets[name]
            for cell in sh[1]:
                cell.font = Font(bold=True)
            for col in sh.columns:
                sh.column_dimensions[col[0].column_letter].width = 22
    return out.getvalue()
