"""Scheduling logic for the duty roster: leave-file parsing, roster generation, checks and Excel export.

Kept separate from the Streamlit UI so office criteria can be added here without touching the app layout.
"""
from __future__ import annotations

import datetime as dt
import io
import random
import re

import pandas as pd

SESSIONS = ("AM", "PM")
DEFAULT_POSTS = ["MSW 1", "MSW 2", "MSW 7", "MSW 8"]   # each post has an AM and a PM slot
SLOTS = len(SESSIONS) * len(DEFAULT_POSTS)

# Can one person take an AM and a PM slot on the same day? (Never two slots in the same half-day.)
DOUBLE_MODES = {
    "short": "Only when short of staff",
    "any": "Allowed on any day",
    "never": "Never",
}


def parse_posts(text) -> list[str]:
    """'MSW 1, MSW 2, MSW 7, MSW 8' -> ['MSW 1', 'MSW 2', 'MSW 7', 'MSW 8']."""
    posts = [p.strip() for p in re.split(r"[,;\n]", str(text or "")) if p.strip()]
    return list(dict.fromkeys(posts)) or list(DEFAULT_POSTS)


def slot_cols(posts: list[str] | None = None) -> list[str]:
    posts = posts or DEFAULT_POSTS
    return [f"{sess} {p}" for sess in SESSIONS for p in posts]


def roster_slot_cols(roster: pd.DataFrame) -> list[str]:
    """Slot columns present in a roster, in order."""
    return [c for c in roster.columns if c not in ("Date", "Day")]


def session_of(col: str) -> str:
    return str(col).split(" ", 1)[0]


SLOT_COLS = slot_cols()
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
             posts: list[str] | None = None, rules: dict[str, dict] | None = None,
             holidays: set[dt.date] | None = None, double: str = "short",
             weekly: bool = True) -> pd.DataFrame:
    """Fill the AM and PM slot of every post on each working day.

    Rules applied:
      1. Nobody is rostered on a leave day or a public holiday.
      2. Nobody takes two slots in the same half-day (AM or PM).
      3. A second slot on the same day (one AM + one PM) depends on `double`:
         "short" = only to fill slots left empty once everyone available has one slot,
         "any"   = allowed whenever that person is the fairest pick,
         "never" = one slot per person per day.
      4. Staff rules: exempt staff are never rostered; nobody is rostered on their "not available" weekdays;
         max duties per week and per month are respected (an AM + PM day counts as 2 duties).
      5. Weekly pattern (`weekly`): each weekday has a fixed line-up, so whoever has a slot on the first Monday
         has it every Monday. When they're away, someone else covers that day only.
      6. People with the fewest duties so far are picked first; slot positions rotate.
    The returned frame carries attrs["covers"] (slots filled by someone outside the pattern) and
    attrs["pattern"] ({weekday 0-4: [person per slot]}).
    """
    rules = rules or {}
    cols = slot_cols(posts)
    rng = random.Random(seed)
    count = {n: 0 for n in staff}
    slot_count = {n: [0] * len(cols) for n in staff}
    week_count: dict[tuple, int] = {}
    days = workdays(year, month, holidays)

    def can_work(n: str, d: dt.date) -> bool:
        """Leave, exemption and unavailable weekdays (no duty-count limits)."""
        if d in leave.get(n, set()):
            return False
        r = rules.get(n)
        return not (r and (r["exempt"] or d.weekday() in r["off_days"]))

    def eligible(n: str, d: dt.date) -> bool:
        if not can_work(n, d):
            return False
        r = rules.get(n)
        if r and r["max_month"] is not None and count[n] >= r["max_month"]:
            return False
        if r and r["max_week"] is not None and week_count.get((n, _week(d)), 0) >= r["max_week"]:
            return False
        return True

    def assign(row, s, p, d, today):
        row[s] = p
        today.setdefault(p, set()).add(session_of(cols[s]))
        count[p] += 1
        slot_count[p][s] += 1
        week_count[(p, _week(d))] = week_count.get((p, _week(d)), 0) + 1

    # ---- Weekly pattern: one line-up per weekday, balanced by expected duties ----
    pattern: dict[int, list[str | None]] = {}
    if weekly:
        expected = {n: 0 for n in staff}
        pat_slot = {n: [0] * len(cols) for n in staff}
        for wd in range(5):
            wd_days = [d for d in days if d.weekday() == wd]
            if not wd_days:
                continue
            avail = {n: sum(can_work(n, d) for d in wd_days) for n in staff}
            pool = [n for n in staff if avail[n] > 0]
            rng.shuffle(pool)
            # fewest expected duties first; among equals, the person who is in most often on this weekday
            pool.sort(key=lambda n: (expected[n], -avail[n]))
            chosen = pool[:len(cols)]
            line = [None] * len(cols)
            for s in range(len(cols)):
                if not chosen:
                    break
                chosen.sort(key=lambda n: pat_slot[n][s])
                p = chosen.pop(0)
                line[s] = p
                pat_slot[p][s] += 1
                expected[p] += avail[p]
            pattern[wd] = line

    covers = 0
    records = []
    for d in days:
        row = [None] * len(cols)
        today: dict[str, set[str]] = {}          # person -> sessions already worked today

        # 1. Pattern holders keep their slot when they can.
        for s, p in enumerate(pattern.get(d.weekday(), [])):
            if p and eligible(p, d) and session_of(cols[s]) not in today.get(p, set()):
                assign(row, s, p, d, today)
        filled_by_pattern = sum(v is not None for v in row)

        # 2. Cover the remaining slots, fairest first.
        avail = [n for n in staff if eligible(n, d)]
        rng.shuffle(avail)
        avail.sort(key=lambda n: count[n])
        free = [s for s in range(len(cols)) if row[s] is None]
        if double == "any":
            for s in free:
                sess = session_of(cols[s])
                cands = [n for n in avail if sess not in today.get(n, set()) and eligible(n, d)]
                if cands:
                    cands.sort(key=lambda n: (count[n], slot_count[n][s]))
                    assign(row, s, cands[0], d, today)
        else:
            # one slot each for people not yet on duty today
            pool = [n for n in avail if n not in today][:len(free)]
            for s in free:
                if not pool:
                    break
                pool.sort(key=lambda n: slot_count[n][s])
                assign(row, s, pool.pop(0), d, today)
            # still short: someone already on duty takes the other half-day
            if double == "short":
                for s in range(len(cols)):
                    if row[s] is not None:
                        continue
                    sess = session_of(cols[s])
                    cands = [n for n in today if sess not in today[n] and eligible(n, d)]
                    if cands:
                        cands.sort(key=lambda n: (count[n], slot_count[n][s]))
                        assign(row, s, cands[0], d, today)
        if weekly:
            covers += sum(v is not None for v in row) - filled_by_pattern
        records.append({"Date": d, "Day": d.strftime("%a"), **dict(zip(cols, row))})

    out = pd.DataFrame(records, columns=["Date", "Day", *cols])
    out.attrs["covers"] = covers
    out.attrs["pattern"] = pattern
    return out


# ---------- Checks ----------

def cell_status(roster: pd.DataFrame, leave: dict[str, set[dt.date]],
                rules: dict[str, dict] | None = None, double: str = "short",
                pattern: dict[int, list] | None = None) -> pd.DataFrame:
    """Same shape as the slot columns, one code per cell. Problems first:
    'empty', 'leave', 'dup' (twice in one half-day, or AM + PM when not allowed), 'exempt', 'offday'.
    Otherwise, with a weekly pattern: 'main' (the person's own slot for that weekday) or 'cover'
    (standing in for the slot's usual holder). Without one: 'double' (AM + PM that day) or ''."""
    rules = rules or {}
    pattern = pattern or {}
    cols = roster_slot_cols(roster)
    status = pd.DataFrame("", index=roster.index, columns=cols)
    for idx, row in roster.iterrows():
        by_sess = {sess: [row[c] for c in cols if session_of(c) == sess] for sess in SESSIONS}
        names = [row[c] for c in cols]
        for c in cols:
            p = row[c]
            if _is_blank(p) or p == "":
                status.at[idx, c] = "empty"
            elif pd.Timestamp(row["Date"]).date() in leave.get(p, set()):
                status.at[idx, c] = "leave"
            elif by_sess.get(session_of(c), []).count(p) > 1 or (double == "never" and names.count(p) > 1):
                status.at[idx, c] = "dup"
            elif rules.get(p, {}).get("exempt"):
                status.at[idx, c] = "exempt"
            elif pd.Timestamp(row["Date"]).weekday() in rules.get(p, {}).get("off_days", set()):
                status.at[idx, c] = "offday"
            elif pattern:
                line = pattern.get(pd.Timestamp(row["Date"]).weekday(), [])
                i = cols.index(c)
                status.at[idx, c] = "main" if i < len(line) and line[i] == p else "cover"
            elif names.count(p) > 1:
                status.at[idx, c] = "double"
    return status


def issues_table(roster: pd.DataFrame, status: pd.DataFrame,
                 rules: dict[str, dict] | None = None) -> pd.DataFrame:
    labels = {"empty": "Unfilled slot", "leave": "Rostered while on leave",
              "dup": "Two slots in the same half-day, or AM + PM when not allowed",
              "exempt": "Staff is exempt from duty", "offday": "Not available on this weekday"}
    rows = []
    for idx, row in roster.iterrows():
        for c in roster_slot_cols(roster):
            kind = status.at[idx, c]
            if kind in labels:
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
    cols = roster_slot_cols(roster)
    am = {n: 0 for n in counts}
    doubles = {n: 0 for n in counts}
    for _, row in roster.iterrows():
        names = [row[c] for c in cols if not _is_blank(row[c]) and row[c] != ""]
        for c in cols:
            p = row[c]
            if not _is_blank(p) and p != "" and session_of(c) == "AM":
                am[p] = am.get(p, 0) + 1
        for p in set(names):
            if names.count(p) > 1:
                doubles[p] = doubles.get(p, 0) + 1
    rows = []
    for n, c in counts.items():
        lv = sum(1 for d in days if pd.Timestamp(d).date() in leave.get(n, set()))
        rows.append({"Staff": n, "Duties": c, "AM": am.get(n, 0), "PM": c - am.get(n, 0),
                     "AM + PM days": doubles.get(n, 0), "Leave days": lv, "Days available": len(days) - lv})
    return pd.DataFrame(rows).sort_values(["Duties", "Staff"], ascending=[False, True]).reset_index(drop=True)


# ---------- Excel export ----------

FILLS = {"main": "DDEBF7", "cover": "FFF2A8",
         "leave": "F8CBC2", "dup": "F9C58D", "empty": "E7E9EC", "exempt": "D9CCF0", "offday": "D9CCF0",
         "double": "DDF3EC"}


def date_label(d) -> str:
    """'1-Dec Tue', matching the leave planning sheet's date headers."""
    d = pd.Timestamp(d)
    return f"{d.day}-{d:%b} {d:%a}"


def to_grid(roster: pd.DataFrame) -> pd.DataFrame:
    """Dates across the top, slots (AM MSW 1 … PM MSW 8) down the side."""
    cols = roster_slot_cols(roster)
    grid = roster.set_index(roster["Date"].map(date_label))[cols].T
    grid.index.name = "Slot"
    grid.columns.name = None
    return grid


def from_grid(grid: pd.DataFrame, roster: pd.DataFrame) -> pd.DataFrame:
    """Write an edited grid back into the roster's row-per-date shape."""
    out = roster.copy()
    labels = roster["Date"].map(date_label)
    for i, lab in enumerate(labels):
        if lab in grid.columns:
            for c in roster_slot_cols(roster):
                if c in grid.index:
                    v = grid.at[c, lab]
                    out.iat[i, out.columns.get_loc(c)] = None if _is_blank(v) or v == "" else v
    return out


def to_excel(roster: pd.DataFrame, status: pd.DataFrame, tally_df: pd.DataFrame,
             issues: pd.DataFrame, title: str) -> bytes:
    """Roster sheet laid out like the leave plan: dates across the top, a weekday row,
    then one row per slot (AM posts, then PM posts)."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    cols = roster_slot_cols(roster)
    dates = [pd.Timestamp(d).date() for d in roster["Date"]]

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        pd.DataFrame().to_excel(xw, sheet_name="Roster")
        tally_df.to_excel(xw, sheet_name="Tally", index=False)
        issues.to_excel(xw, sheet_name="Issues", index=False)

        ws = xw.sheets["Roster"]
        bold, center = Font(bold=True), Alignment(horizontal="center", vertical="center")
        head_fill = PatternFill("solid", fgColor="DCE6F1")
        thin = Side(style="thin", color="A6A6A6")
        thick = Side(style="medium", color="404040")
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=14)
        ws.cell(row=2, column=1, value="Session")
        ws.cell(row=2, column=2, value="Post")
        for j, d in enumerate(dates, start=3):
            c = ws.cell(row=2, column=j, value=d)
            c.number_format = "d-mmm"
            ws.cell(row=3, column=j, value=d.strftime("%a"))
        for rr in (2, 3):
            for j in range(1, len(dates) + 3):
                cell = ws.cell(row=rr, column=j)
                cell.font, cell.alignment, cell.fill = bold, center, head_fill
                cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)

        r0 = 4
        for i, col in enumerate(cols):
            rr = r0 + i
            sess, post = session_of(col), col.split(" ", 1)[1] if " " in col else col
            ws.cell(row=rr, column=1, value=sess).font = bold
            ws.cell(row=rr, column=2, value=post).font = bold
            last_of_session = i + 1 < len(cols) and session_of(cols[i + 1]) != sess
            for j in range(1, len(dates) + 3):
                cell = ws.cell(row=rr, column=j)
                cell.alignment = center
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thick if last_of_session else thin)
            for j, (_, row) in enumerate(roster.iterrows(), start=3):
                v = row[col]
                cell = ws.cell(row=rr, column=j, value=None if _is_blank(v) or v == "" else v)
                kind = status.iat[j - 3, status.columns.get_loc(col)] if col in status.columns else ""
                if kind:
                    cell.fill = PatternFill("solid", fgColor=FILLS[kind])
        # merge the session label down its block
        start = r0
        for i in range(1, len(cols) + 1):
            if i == len(cols) or session_of(cols[i]) != session_of(cols[start - r0]):
                if r0 + i - 1 > start:
                    ws.merge_cells(start_row=start, start_column=1, end_row=r0 + i - 1, end_column=1)
                start = r0 + i

        ws.column_dimensions["A"].width = 9
        ws.column_dimensions["B"].width = 9
        for j in range(3, len(dates) + 3):
            ws.column_dimensions[get_column_letter(j)].width = 14
        ws.freeze_panes = "C4"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        for name in ("Tally", "Issues"):
            sh = xw.sheets[name]
            for cell in sh[1]:
                cell.font = Font(bold=True)
            for col in sh.columns:
                sh.column_dimensions[col[0].column_letter].width = 22
    return out.getvalue()
