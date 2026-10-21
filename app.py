"""Duty Roster Builder — Streamlit app.

Run locally:  streamlit run app.py
"""
import calendar
import datetime as dt

import pandas as pd
import streamlit as st

import roster_engine as eng

st.set_page_config(page_title="Duty Roster Builder", layout="wide")

# ---------- Session state ----------
ss = st.session_state
today = dt.date.today()
defaults = {
    "staff": list(eng.EXAMPLE_STAFF),
    "example": True,
    "raw_rows": None,
    "file_name": None,
    "file_sig": None,
    "file_error": None,
    "layout": None,
    "added": 0,
    "seed": 1,
    "sig": None,
    "roster": None,
    "editor_ver": 0,
    "staff_ver": 0,
}
for k, v in defaults.items():
    ss.setdefault(k, v)


def reset_to_example():
    for k in ("raw_rows", "file_name", "file_sig", "file_error", "layout"):
        ss[k] = None
    ss.staff, ss.example, ss.added = list(eng.EXAMPLE_STAFF), True, 0
    ss.staff_ver += 1


# ---------- Sidebar: inputs ----------
with st.sidebar:
    st.header("1 · Month")
    next_month = (today.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
    c1, c2 = st.columns(2)
    year = c1.number_input("Year", min_value=2020, max_value=2100, value=next_month.year, step=1)
    month = c2.selectbox("Month", range(1, 13), index=next_month.month - 1,
                         format_func=lambda m: calendar.month_abbr[m])
    year, month = int(year), int(month)
    slots = int(st.number_input(
        "Duty slots per day", min_value=1, max_value=eng.MAX_SLOTS, value=eng.SLOTS, step=1,
        help="How many people are on duty each working day. Changing this regenerates the roster.",
    ))

    st.header("2 · Leave plan")
    up = st.file_uploader("Upload the leave planning Excel", type=["xlsx", "xls", "csv"],
                          help="The first sheet is read.")
    if up is not None:
        sig = (up.name, up.size)
        if sig != ss.file_sig:
            ss.file_sig = sig
            try:
                rows = eng.read_rows(up.name, up.getvalue())
                res = eng.parse_leave(rows, year, month)
            except Exception:
                rows, res = None, None
                ss.file_error = f"Couldn't open {up.name}. Make sure it's an Excel or CSV file."
            if res:
                ss.raw_rows, ss.file_name, ss.layout, ss.file_error = rows, up.name, res[0], None
                names = list(res[1].keys())
                if ss.example:
                    ss.staff, ss.example, ss.added = names, False, len(names)
                else:
                    new = [n for n in names if n not in ss.staff]
                    ss.staff, ss.added = ss.staff + new, len(new)
                ss.staff_ver += 1
            elif rows is not None:
                ss.file_error = (f"Couldn't find staff names and leave dates in {up.name}. "
                                 "See the accepted layouts below.")

    if ss.file_error:
        st.error(ss.file_error)
    if ss.file_name:
        st.success(f"**{ss.file_name}** · {ss.layout}")
        if ss.added:
            st.caption(f"{ss.added} name{'s' if ss.added > 1 else ''} added to the staff list from this file.")
        st.button("Clear leave plan and use example data", on_click=reset_to_example)
    else:
        st.caption("Example leave is loaded until you upload a file.")

    with st.expander("Accepted layouts"):
        st.markdown(
            "**Calendar grid** — one row per staff, names in the first column, dates or day numbers "
            "across the top. Any text in a cell (AL, MC, L) counts as leave.\n\n"
            "**List** — columns named *Name*, *Start* and *End*, one row per leave block."
        )

    st.header("3 · Staff")
    staff_edit = st.data_editor(
        pd.DataFrame({"Name": ss.staff}), key=f"staff_{ss.staff_ver}",
        num_rows="dynamic", hide_index=True, use_container_width=True,
    )
    staff = list(dict.fromkeys(
        str(n).strip() for n in staff_edit["Name"] if n is not None and str(n).strip()
    ))
    st.caption("Add a row at the bottom, or select rows and delete them.")

    st.divider()
    if st.button("Reshuffle roster", use_container_width=True,
                 help="Generate a different roster that follows the same rules."):
        ss.seed += 1

# ---------- Leave for the chosen month ----------
if ss.raw_rows is not None:
    parsed = eng.parse_leave(ss.raw_rows, year, month)  # re-read: day-number headers depend on the month
    leave = parsed[1] if parsed else {}
else:
    leave = eng.example_leave(year, month)

# ---------- Generate when inputs change ----------
sig = (tuple(staff), year, month, slots, ss.seed,
       tuple(sorted((n, tuple(sorted(ds))) for n, ds in leave.items())))
if sig != ss.sig:
    ss.roster = eng.generate(staff, leave, year, month, ss.seed, slots)
    ss.sig = sig
    ss.editor_ver += 1

# ---------- Main page ----------
title = f"{calendar.month_name[month]} {year}"
st.title("Duty Roster Builder")
st.caption(f"{slots} duty slot{'s' if slots != 1 else ''} a day · Monday to Friday")
if len(staff) < slots:
    st.warning(f"There are {len(staff)} staff for {slots} slots a day, so some slots will stay unfilled. "
               "Add staff or reduce the slots per day.")
if ss.example:
    st.info("You're looking at **example** staff and leave. Upload the leave plan in the sidebar to use your own.")

summary = st.container()

tab_roster, tab_review, tab_fair, tab_leave = st.tabs(["Roster", "Review", "Fairness", "Leave file"])

with tab_roster:
    st.subheader(title)
    st.caption("Click any slot to swap someone in. Changing staff, leave or month regenerates the roster "
               "and replaces manual edits.")
    edited = st.data_editor(
        ss.roster, key=f"roster_{ss.editor_ver}", hide_index=True, use_container_width=True,
        disabled=["Date", "Day"], height=min(38 * (len(ss.roster) + 1) + 3, 900),
        column_config={
            "Date": st.column_config.DateColumn("Date", format="DD MMM", width="small"),
            "Day": st.column_config.TextColumn("Day", width="small"),
            **{c: st.column_config.SelectboxColumn(c, options=staff, required=False) for c in eng.slot_cols(slots)},
        },
    )

status = eng.cell_status(edited, leave)
issues = eng.issues_table(edited, status)
tally = eng.tally(edited, staff, leave)
n_empty = int((status == "empty").sum().sum())
n_conf = int(status.isin(["leave", "dup"]).sum().sum())

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

with tab_review:
    css_map = {"leave": "background-color:#FBE6E1;color:#8A2A1A;font-weight:600",
               "dup": "background-color:#FBEFD2;color:#6E4B00;font-weight:600",
               "empty": "background-color:#EEF1F4", "": ""}
    def _styles(df):
        out = pd.DataFrame("", index=df.index, columns=df.columns)
        cols = list(status.columns)
        out[cols] = status[cols].apply(lambda col: col.map(css_map))
        return out
    view = edited.copy()
    view["Date"] = pd.to_datetime(view["Date"]).dt.strftime("%d %b")
    st.dataframe(view.fillna("").style.apply(_styles, axis=None), hide_index=True, use_container_width=True,
                 height=min(38 * (len(view) + 1) + 3, 900))
    st.caption("Red: rostered while on leave · Amber: rostered twice that day · Grey: unfilled")
    if len(issues):
        st.dataframe(issues, hide_index=True, use_container_width=True,
                     column_config={"Date": st.column_config.DateColumn(format="DD MMM YYYY")})
    else:
        st.success("No issues. Every slot is filled and nobody is rostered on leave.")

with tab_fair:
    top = int(tally["Duties"].max()) if len(tally) else 1
    st.dataframe(tally, hide_index=True, use_container_width=True, column_config={
        "Duties": st.column_config.ProgressColumn("Duties", min_value=0, max_value=max(top, 1), format="%d"),
    })
    if len(tally):
        st.caption(f"Duties range from {int(tally['Duties'].min())} to {top} per person this month.")

with tab_leave:
    lv = pd.DataFrame(
        [{"Staff": n, "Leave days this month": sum(1 for d in ds if d.year == year and d.month == month)}
         for n, ds in leave.items()]
    )
    if ss.raw_rows is not None:
        st.markdown(f"**{ss.file_name}** read as **{ss.layout}**. First rows of the file:")
        raw = pd.DataFrame(ss.raw_rows[:10]).fillna("").astype(str)
        st.dataframe(raw, use_container_width=True)
    else:
        st.markdown("Example leave (no file uploaded):")
    if len(lv):
        st.dataframe(lv[lv["Leave days this month"] > 0], hide_index=True, use_container_width=True)
