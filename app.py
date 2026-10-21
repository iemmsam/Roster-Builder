"""Duty Roster Builder — Streamlit app.

Run locally:  streamlit run app.py
"""
import calendar
import datetime as dt
import importlib

import pandas as pd
import streamlit as st

import roster_engine as eng

# Streamlit Cloud reruns the new app.py after a git push but keeps the old engine module in memory.
# Reloading keeps the two in step.
eng = importlib.reload(eng)

st.set_page_config(page_title="Duty Roster Builder", layout="wide")

EXAMPLE_RULES = eng.empty_rules(eng.EXAMPLE_STAFF)
EXAMPLE_RULES.loc[EXAMPLE_RULES.Name == "Kelvin", "Exempt"] = True
EXAMPLE_RULES.loc[EXAMPLE_RULES.Name == "Lakshmi", "Not available on"] = "Fri"
EXAMPLE_RULES.loc[EXAMPLE_RULES.Name == "Marcus", "Max per week"] = 2

# ---------- Session state ----------
ss = st.session_state
defaults = {
    "rules_table": EXAMPLE_RULES.copy(),
    "example_staff": True,
    "raw_rows": None,          # leave plan rows, from paste or upload
    "leave_source": None,      # "Pasted from Excel" or the file name
    "layout": None,
    "leave_msg": None,         # (kind, text) shown after loading
    "file_sig": None,
    "rules_msg": None,
    "seed": 1,
    "sig": None,
    "roster": None,
    "editor_ver": 0,
    "rules_ver": 0,
    "pending_month": None,     # (year, month) read from a pasted plan; applied before the sidebar draws
}
for k, v in defaults.items():
    ss.setdefault(k, v)

today = dt.date.today()
next_month = (today.replace(day=1) + dt.timedelta(days=32)).replace(day=1)


def plural(n, word):
    return f"{n} {word}{'' if n == 1 else 's'}"


def load_leave(rows, source, year, month):
    """Store leave rows if they can be read, switch to the month they cover, and add new names to the staff table."""
    ym = eng.detect_month(rows, year) if rows else None
    if ym:
        y, m = ym
        # headers like "1-Jan" carry no year: pick the year closest to the month being planned
        y = min((y - 1, y, y + 1), key=lambda yy: abs((yy * 12 + m) - (year * 12 + month)))
        if (y, m) != (year, month):
            ss.pending_month = (y, m)
        year, month = y, m
    res = eng.parse_leave(rows, year, month) if rows else None
    if not res:
        ss.leave_msg = ("error", f"Couldn't find staff names and leave dates in the {source.lower()}. "
                                 "Check that you copied the header row too. See the accepted layouts below.")
        return
    ss.raw_rows, ss.leave_source, ss.layout = rows, source, res[0]
    names = list(res[1].keys())
    table = ss.rules_table
    if ss.example_staff:
        table, added = eng.empty_rules(names), len(names)
        ss.example_staff = False
    else:
        new = [n for n in names if n not in set(table["Name"].astype(str))]
        table = pd.concat([table, eng.empty_rules(new)], ignore_index=True)
        added = len(new)
    ss.rules_table = table
    ss.rules_ver += 1
    msg = f"Read {plural(len(names), 'staff member')} for {calendar.month_name[month]} {year} as an {res[0].lower()}."
    if res[2]:
        msg += " Public holiday" + ("s" if len(res[2]) > 1 else "") + ": " + \
               ", ".join(d.strftime("%-d %b") for d in sorted(res[2])) + " (no duty)."
    if added:
        msg += f" {plural(added, 'name')} added to Staff & rules."
    ss.leave_msg = ("success", msg)


def clear_leave():
    for k in ("raw_rows", "leave_source", "layout", "leave_msg", "file_sig"):
        ss[k] = None


# ---------- Sidebar ----------
ss.setdefault("year_in", next_month.year)
ss.setdefault("month_in", next_month.month)
if ss.pending_month:
    ss.year_in, ss.month_in = ss.pending_month
    ss.pending_month = None

with st.sidebar:
    st.header("Month")
    c1, c2 = st.columns(2)
    year = c1.number_input("Year", min_value=2020, max_value=2100, step=1, key="year_in")
    month = c2.selectbox("Month", range(1, 13), key="month_in", format_func=lambda m: calendar.month_abbr[m])
    year, month = int(year), int(month)
    st.header("Duty slots")
    posts = eng.parse_posts(st.text_input(
        "Duty posts", value=", ".join(eng.DEFAULT_POSTS), key="posts_in",
        help="Separate posts with commas. Each post has an AM and a PM slot.",
    ))
    slots = len(eng.slot_cols(posts))
    st.caption(f"{plural(len(posts), 'post')} × AM and PM = {plural(slots, 'slot')} a day")
    double = st.selectbox(
        "Same person in an AM and a PM slot", list(eng.DOUBLE_MODES), format_func=eng.DOUBLE_MODES.get,
        key="double_in",
        help="Nobody is ever given two slots in the same half-day. This sets when someone may do both AM and PM "
             "on one day. 'Only when short of staff' gives everyone available one slot first, then doubles "
             "people up to fill what's left.",
    )
    per_session = len(posts)
    weekly = st.toggle(
        "Same slot on the same weekday", value=True, key="weekly_in",
        help="Each weekday keeps one line-up all month: whoever does AM MSW 1 on the first Monday does it every "
             "Monday. When that person is away, someone else covers that day only.",
    )
    st.divider()
    if st.button("Reshuffle roster", use_container_width=True,
                 help="Generate a different roster that follows the same rules."):
        ss.seed += 1
    st.caption("Steps: paste the leave plan, check Staff & rules, then review the roster and download it.")

title = f"{calendar.month_name[month]} {year}"
st.title("Duty Roster Builder")
st.caption(f"AM and PM · {', '.join(posts)} · {plural(slots, 'slot')} a day · Monday to Friday · {title}")
banner = st.container()
summary = st.container()

tab_leave, tab_rules, tab_roster, tab_review, tab_fair = st.tabs(
    ["Leave plan", "Staff & rules", "Roster", "Review", "Fairness"])

# ---------- Leave plan tab ----------
with tab_leave:
    method = st.radio("How do you want to add the leave plan?", ["Paste from Excel", "Upload a file"],
                      horizontal=True, key="leave_method")
    if method == "Paste from Excel":
        st.markdown(
            "1. In the leave planning Excel, click the **Name** heading cell and drag to the bottom-right "
            "of the last staff row, so the selection covers the date row (1-Dec, 2-Dec …), the day row "
            "and every staff row. Including the # column, the Total columns or the total rows is fine; "
            "they're ignored.\n"
            "2. Press **Ctrl+C**.\n"
            "3. Click in the box below, press **Ctrl+V**, then click **Use this leave plan**.\n\n"
            "The month is read from the dates, so the sidebar switches to it on its own."
        )
        with st.form("leave_paste", clear_on_submit=False):
            text = st.text_area("Leave plan cells", height=200, key="leave_text",
                                placeholder="Name\t1-Dec\t2-Dec\t3-Dec\t4-Dec …\n\tTue\tWed\tThu\tFri …\n"
                                            "Lim Boon Chee\t1\t1\t1\toff …\nChew Tee Kit\t1\t1\t0\t0 …")
            if st.form_submit_button("Use this leave plan", type="primary"):
                load_leave(eng.rows_from_paste(text), "Pasted from Excel", year, month)
                if ss.pending_month:
                    st.rerun()
    else:
        up = st.file_uploader("Leave planning Excel", type=["xlsx", "xls", "csv"], help="The first sheet is read.")
        if up is not None and (up.name, up.size) != ss.file_sig:
            ss.file_sig = (up.name, up.size)
            try:
                load_leave(eng.read_rows(up.name, up.getvalue()), up.name, year, month)
            except Exception:
                ss.leave_msg = ("error", f"Couldn't open {up.name}. Make sure it's an Excel or CSV file.")
            if ss.pending_month:
                st.rerun()

    if ss.leave_msg:
        getattr(st, ss.leave_msg[0])(ss.leave_msg[1])
    if ss.raw_rows is None:
        st.info("Example leave is in use until you add your leave plan.")

    with st.expander("Accepted layouts"):
        st.markdown(
            "**Attendance grid** (the office leave plan): names down the side, dates across the top. "
            "**1** = in office and available for duty. **0**, **off**, or a blank cell = away. "
            "A date column that's blank for everyone (e.g. 25-Dec) is a public holiday with no duty.\n\n"
            "**Leave-mark grid**: same shape, but cells are blank when present and marked (AL, MC, L) "
            "when on leave.\n\n"
            "**List**: columns named *Name*, *Start* and *End* (or From/To), one row per leave block.\n\n"
            "Day numbers on their own are read as days of the month chosen in the sidebar."
        )

# ---------- Leave for the chosen month ----------
if ss.raw_rows is not None:
    parsed = eng.parse_leave(ss.raw_rows, year, month)  # re-read: day-number headers depend on the month
    leave, holidays = (parsed[1], parsed[2]) if parsed else ({}, set())
    plan_month = eng.detect_month(ss.raw_rows, year)
    if plan_month and plan_month[1] != month:
        st.warning(f"The leave plan is for {calendar.month_name[plan_month[1]]}, but the roster is for "
                   f"{calendar.month_name[month]}. Change the month in the sidebar, or paste that month's plan.")
else:
    leave, holidays = eng.example_leave(year, month), set()

# ---------- Staff & rules tab ----------
with tab_rules:
    st.markdown("Each row is one staff member. Leave a cell blank if the rule doesn't apply to that person. "
                "You can type in the table, or paste a rules table from Excel.")
    with st.expander("Paste rules from Excel"):
        st.markdown(
            "Copy a table with a header row. The **Name** column is required. Other columns are matched by "
            "their header: *Exempt* (Y/N), *Not available on* (e.g. Fri, or Mon, Wed), *Max per week*, "
            "*Max per month*. Pasting updates the people listed and adds anyone new; others keep their rules."
        )
        with st.form("rules_paste"):
            rtext = st.text_area("Rules cells", height=160, key="rules_text",
                                 placeholder="Name\tExempt\tNot available on\tMax per week\tMax per month\n"
                                             "Aisyah\t\tFri\t2\t")
            if st.form_submit_button("Use these rules", type="primary"):
                table = eng.rules_from_rows(eng.rows_from_paste(rtext))
                if table is None:
                    ss.rules_msg = ("error", "Couldn't find a Name column. Copy the header row together with "
                                             "the rows below it.")
                else:
                    keep = ss.rules_table if not ss.example_staff else eng.empty_rules([])
                    keep = keep[~keep["Name"].astype(str).isin(set(table["Name"]))]
                    ss.rules_table = pd.concat([keep, table], ignore_index=True)
                    ss.example_staff = False
                    ss.rules_ver += 1
                    ss.rules_msg = ("success", f"Loaded rules for {plural(len(table), 'staff member')}. "
                                               f"The table now has {plural(len(ss.rules_table), 'staff member')}.")
    if ss.rules_msg:
        getattr(st, ss.rules_msg[0])(ss.rules_msg[1])
    if ss.example_staff:
        st.info("These are **example** staff. Paste your leave plan or rules to replace them.")

    rules_edit = st.data_editor(
        ss.rules_table, key=f"rules_{ss.rules_ver}", num_rows="dynamic", hide_index=True,
        use_container_width=True,
        column_config={
            "Name": st.column_config.TextColumn("Name", required=True),
            "Exempt": st.column_config.CheckboxColumn("Exempt", help="Never rostered", default=False),
            "Not available on": st.column_config.TextColumn(
                "Not available on", help="Weekdays this person can't do duty, e.g. Fri or Mon, Wed"),
            "Max per week": st.column_config.NumberColumn("Max per week", min_value=0, step=1, format="%d"),
            "Max per month": st.column_config.NumberColumn("Max per month", min_value=0, step=1, format="%d"),
        },
    )
    rules_edit = rules_edit[rules_edit["Name"].notna() & (rules_edit["Name"].astype(str).str.strip() != "")]
    rules_edit = rules_edit.assign(Name=rules_edit["Name"].astype(str).str.strip()).drop_duplicates("Name")
    staff = list(rules_edit["Name"])
    rules = eng.rules_dict(rules_edit)

    bad = [f"{n} ({r['Not available on']})" for _, r in rules_edit.iterrows()
           if (n := r["Name"]) and not eng._is_blank(r["Not available on"])
           and str(r["Not available on"]).strip() and not rules[n]["off_days"]]
    if bad:
        st.warning("Couldn't read the weekday for: " + ", ".join(bad) + ". Use Mon, Tue, Wed, Thu or Fri.")

# ---------- Generate when inputs change ----------
sig = (tuple(staff), year, month, tuple(posts), double, weekly, ss.seed, tuple(sorted(holidays)),
       tuple(sorted((n, tuple(sorted(ds))) for n, ds in leave.items())),
       tuple(sorted((n, r["exempt"], tuple(sorted(r["off_days"])), r["max_week"], r["max_month"])
                    for n, r in rules.items())))
if sig != ss.sig:
    ss.roster = eng.generate(staff, leave, year, month, ss.seed, posts, rules, holidays, double, weekly)
    ss.covers = ss.roster.attrs.get("covers", 0)
    ss.pattern = ss.roster.attrs.get("pattern", {})
    ss.sig = sig
    ss.editor_ver += 1

# ---------- Roster tab ----------
CSS = {"main": "background-color:#DDEBF7;color:#123B63",
       "cover": "background-color:#FFF2A8;color:#4A3B00",
       "leave": "background-color:#F8CBC2;color:#7A1F0F;font-weight:600",
       "dup": "background-color:#F9C58D;color:#5C2E00;font-weight:600",
       "exempt": "background-color:#D9CCF0;color:#3B2270;font-weight:600",
       "offday": "background-color:#D9CCF0;color:#3B2270;font-weight:600",
       "double": "background-color:#DDF3EC;color:#0F4A38",
       "empty": "background-color:#E7E9EC", "": ""}


def legend(weekly_on: bool) -> str:
    base = ("Blue: their own slot for that weekday · Yellow: covering for the usual person"
            if weekly_on else "Green: does both AM and PM that day")
    return (base + " · Red: on leave · Orange: two slots in the same half-day (or AM + PM when not allowed)"
            " · Purple: breaks a staff rule · Grey: unfilled")


with tab_roster:
    view_box = st.container()
    with st.expander("Make changes"):
        st.caption("Click any slot to swap someone in; the coloured roster above updates. Changing leave, staff, "
                   "rules, month or duty slots regenerates the roster and replaces manual edits.")
        grid = eng.to_grid(ss.roster)
        edited_grid = st.data_editor(
            grid, key=f"roster_{ss.editor_ver}", hide_index=False, use_container_width=True,
            height=38 * (len(grid) + 1) + 3,
            column_config={
                "Slot": st.column_config.TextColumn("Slot", disabled=True, pinned=True),
                **{c: st.column_config.SelectboxColumn(c, options=staff, required=False, width="medium")
                   for c in grid.columns},
            },
        )
    edited = eng.from_grid(edited_grid, ss.roster)

status = eng.cell_status(edited, leave, rules, double, ss.get("pattern") if weekly else None)
issues = eng.issues_table(edited, status, rules)
tally = eng.tally(edited, staff, leave)
n_empty = int((status == "empty").sum().sum())
n_conf = len(issues) - n_empty

view = eng.to_grid(edited)
status_grid = status.set_axis(view.columns, axis=0).T.reindex(index=view.index)
styled = view.fillna("").style.apply(lambda df: status_grid.apply(lambda col: col.map(CSS)), axis=None)

with view_box:
    st.dataframe(styled, use_container_width=True, height=38 * (len(view) + 1) + 3)
    st.caption(legend(weekly))

with banner:
    # Fewest people who can cover every slot: one per slot, or one per post when AM + PM doubling is allowed.
    need = slots if double == "never" else per_session
    on_duty = sum(1 for n in staff if not rules.get(n, {}).get("exempt"))
    if on_duty < need:
        st.warning(f"{plural(on_duty, 'staff member')} can do duty for {plural(slots, 'slot')} a day, so some "
                   "slots will stay unfilled. Add staff or remove a post.")
    else:
        short, doubled = [], []
        for d in edited["Date"]:
            d = pd.Timestamp(d).date()
            n = sum(1 for p in staff if d not in leave.get(p, set()) and not rules.get(p, {}).get("exempt")
                    and d.weekday() not in rules.get(p, {}).get("off_days", set()))
            if n < need:
                short.append(f"{d:%-d %b} ({n})")
            elif double == "short" and n < slots:
                doubled.append(f"{d:%-d %b} ({n})")
        if short:
            st.warning(f"Too few staff are available on {plural(len(short), 'day')} to fill every slot, so those "
                       "days have unfilled slots: " + ", ".join(short) + ".")
        if doubled:
            st.caption(f"Fewer than {slots} staff are in on {plural(len(doubled), 'day')}, so some people do both "
                       "AM and PM: " + ", ".join(doubled) + ".")
    if weekly and len(edited):
        total = len(edited) * slots
        st.caption(f"Weekly pattern: {total - ss.get('covers', 0)} of {total} slots follow each weekday's line-up. "
                   f"The other {ss.get('covers', 0)} are covers for people who are away that day.")
    if holidays:
        st.caption("No duty on public holidays: " + ", ".join(f"{d:%-d %b}" for d in sorted(holidays)) + ".")
    if ss.example_staff or ss.raw_rows is None:
        st.info("Some **example** data is in use. Add your leave plan and staff rules in the first two tabs.")

with summary:
    m1, m2, m3, m4, m5 = st.columns([1, 1, 1, 1, 1.4])
    m1.metric("Working days", len(edited))
    m2.metric("Slots", len(edited) * slots)
    m3.metric("Unfilled", n_empty)
    m4.metric("Conflicts", n_conf)
    m5.download_button(
        "Download Excel", type="primary", use_container_width=True,
        data=eng.to_excel(edited, status, tally, issues, f"Duty roster · {title}"),
        file_name=f"duty-roster-{year}-{month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------- Review tab ----------
with tab_review:
    st.dataframe(styled, use_container_width=True, height=38 * (len(view) + 1) + 3)
    st.caption(legend(weekly))
    if len(issues):
        st.dataframe(issues, hide_index=True, use_container_width=True,
                     column_config={"Date": st.column_config.DateColumn(format="DD MMM YYYY")})
    else:
        st.success("No issues. Every slot is filled and every rule is met.")

# ---------- Fairness tab ----------
with tab_fair:
    t = tally.merge(rules_edit[["Name", "Max per week", "Max per month"]].rename(columns={"Name": "Staff"}),
                    on="Staff", how="left")
    top = int(t["Duties"].max()) if len(t) else 1
    st.dataframe(t, hide_index=True, use_container_width=True, column_config={
        "Duties": st.column_config.ProgressColumn("Duties", min_value=0, max_value=max(top, 1), format="%d"),
        "Max per week": st.column_config.NumberColumn(format="%d"),
        "Max per month": st.column_config.NumberColumn(format="%d"),
    })
    pool = t[~t["Staff"].map(lambda n: rules.get(n, {}).get("exempt", False))]
    if len(pool):
        st.caption(f"Among staff who do duty, duties range from {int(pool['Duties'].min())} to "
                   f"{int(pool['Duties'].max())} this month.")

# ---------- Leave plan tab: what was read ----------
with tab_leave:
    if ss.raw_rows is not None:
        st.markdown(f"**{ss.leave_source}**, read as **{ss.layout}**. First rows:")
        st.dataframe(pd.DataFrame(ss.raw_rows[:8]).fillna("").astype(str), use_container_width=True)
        st.button("Clear leave plan", on_click=clear_leave)
    lv = pd.DataFrame([{"Staff": n, f"Leave days in {calendar.month_abbr[month]}":
                        sum(1 for d in ds if d.year == year and d.month == month)} for n, ds in leave.items()])
    if len(lv):
        lv = lv[lv.iloc[:, 1] > 0]
        st.markdown("**Leave this month**" + ("" if ss.raw_rows is not None else " (example)"))
        st.dataframe(lv, hide_index=True, use_container_width=True)
