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
SKIP_RE = re.compile(r"^(day|days|date|dates|name|staff|total)$", re.I)

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


def parse_leave(rows: list[list], year: int, month: int) -> tuple[str, dict[str, set[dt.date]]] | None:
    """Detect the layout and return (layout name, {staff: set of leave dates}), or None if unrecognised.

    Layouts:
      - Calendar grid: names in one column, dates or day numbers across a header row; any mark = leave.
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
                name = "" if _is_blank(r[name_idx]) else str(r[name_idx]).strip()
                if not name or SKIP_RE.match(name):
                    continue
                leave.setdefault(name, set())
                a = to_date(r[start_idx], year, month)
                if not a:
                    continue
                b = (to_date(r[end_idx], year, month) if end_idx >= 0 else None) or a
                d = a
                while d <= b and (d - a).days < 370:
                    leave[name].add(d)
                    d += dt.timedelta(days=1)
            if leave:
                return "List (Name / Start / End)", leave

        date_cols = [(i, d) for i, v in enumerate(head) if (d := to_date(v, year, month))]
        if len(date_cols) >= 5:
            date_idx = {i for i, _ in date_cols}
            if name_idx < 0:
                name_idx = next((i for i, v in enumerate(head) if not _is_blank(v) and i not in date_idx), 0)
            leave = {}
            for r in rows[hi + 1:]:
                name = "" if _is_blank(r[name_idx]) else str(r[name_idx]).strip()
                if not name or SKIP_RE.match(name):
                    continue
                marks = leave.setdefault(name, set())
                for ci, d in date_cols:
                    cell = r[ci] if ci < len(r) else None
                    if not _is_blank(cell) and str(cell).strip() not in ("", "0"):
                        marks.add(d)
            if leave:
                return "Calendar grid", leave
    return None


# ---------- Roster generation ----------

def workdays(year: int, month: int) -> list[dt.date]:
    d, out = dt.date(year, month, 1), []
    while d.month == month:
        if d.weekday() < 5:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


def generate(staff: list[str], leave: dict[str, set[dt.date]], year: int, month: int, seed: int,
             slots: int = SLOTS) -> pd.DataFrame:
    """Fill the given number of duty slots per working day.

    Current placeholder rules:
      1. Nobody is rostered on a leave day.
      2. One slot per person per day.
      3. People with the fewest duties so far are picked first; slot positions rotate.
    Add office criteria here (e.g. max duties per week, no consecutive days, seniority per slot).
    """
    rng = random.Random(seed)
    count = {n: 0 for n in staff}
    cols = slot_cols(slots)
    slot_count = {n: [0] * slots for n in staff}
    records = []
    for d in workdays(year, month):
        avail = [n for n in staff if d not in leave.get(n, set())]
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
        records.append({"Date": d, "Day": d.strftime("%a"), **dict(zip(cols, row))})
    return pd.DataFrame(records, columns=["Date", "Day", *cols])


# ---------- Checks ----------

def cell_status(roster: pd.DataFrame, leave: dict[str, set[dt.date]]) -> pd.DataFrame:
    """Same shape as the slot columns: '', 'empty', 'leave' or 'dup' for each cell."""
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
    return status


def issues_table(roster: pd.DataFrame, status: pd.DataFrame) -> pd.DataFrame:
    labels = {"empty": "Unfilled slot", "leave": "Rostered while on leave", "dup": "Rostered twice that day"}
    rows = []
    for idx, row in roster.iterrows():
        for c in roster_slot_cols(roster):
            kind = status.at[idx, c]
            if kind:
                rows.append({"Date": row["Date"], "Slot": c, "Staff": row[c] or "", "Issue": labels[kind]})
    return pd.DataFrame(rows, columns=["Date", "Slot", "Staff", "Issue"])


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

FILLS = {"leave": "FBE6E1", "dup": "FBEFD2", "empty": "EEF1F4"}


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
