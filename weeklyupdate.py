import streamlit as st
import pandas as pd
from collections import Counter
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Program Status (Exec Packet)", layout="wide")
st.title("Program Status (Exec Packet) — Streamlit")

st.markdown(
    "Upload the current Excel packet, review/update statuses, run quick checks, "
    "and export a fresh packet — no more cell gymnastics. 🎯"
)

# ---------- Helpers ----------
def load_excel(file):
    xls = pd.ExcelFile(file)
    data = {}
    for s in xls.sheet_names:
        try:
            df = xls.parse(s)
            data[s] = df
        except Exception as e:
            st.warning(f"Could not read sheet '{s}': {e}")
    return data

def to_excel_bytes(sheets_dict):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as xw:
        for name, df in sheets_dict.items():
            if isinstance(df, pd.DataFrame):
                df.to_excel(xw, sheet_name=name[:31], index=False)
    return out.getvalue()

def rag_emoji(val):
    if not isinstance(val, str):
        return "🟡"
    s = val.strip().lower()
    if s.startswith("g"): return "🟢"
    if s.startswith("y"): return "🟡"
    if s.startswith("r"): return "🔴"
    return "🟡"

def guess_status_columns(df: pd.DataFrame):
    return [c for c in df.columns if any(k in str(c).lower() for k in ["status", "rag"])]

def guess_key_columns(df: pd.DataFrame):
    keys = ["project", "initiative", "owner", "date", "risk", "issue", "status"]
    return [c for c in df.columns if any(k in str(c).lower() for k in keys)]

STATUS_BASE_OPTIONS = ["Green", "Yellow", "Red", "Amber", "On Track", "At Risk"]

def status_options_for(df: pd.DataFrame, column: str):
    """Merge default RAG values with any existing free-text entries."""
    if column not in df.columns:
        return STATUS_BASE_OPTIONS
    existing = (
        df[column]
        .dropna()
        .astype(str)
        .map(str.strip)
        .replace("", pd.NA)
        .dropna()
        .tolist()
    )
    ordered = []
    for value in STATUS_BASE_OPTIONS + existing:
        if value not in ordered:
            ordered.append(value)
    return ordered or STATUS_BASE_OPTIONS

def normalize_status_label(value) -> str:
    """Collapse various free-text status values into consistent buckets."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "Unknown"
    text = str(value).strip().lower()
    if not text:
        return "Unknown"
    if "hold" in text:
        return "On Hold"
    if "complete" in text or "done" in text or "closed" in text:
        return "Complete"
    if text.startswith("g") or "on track" in text or "green" in text:
        return "Green"
    if text.startswith("r") or "risk" in text or "blocked" in text or "critical" in text:
        return "Red"
    if text.startswith("y") or "amber" in text or "progress" in text or "watch" in text:
        return "Amber"
    return text.title()

def aggregate_status_counts(sheets: dict[str, pd.DataFrame]):
    overall = Counter()
    per_sheet = {}
    missing_status_columns = []
    for name, df in sheets.items():
        if not isinstance(df, pd.DataFrame) or df.empty:
            continue
        status_cols = guess_status_columns(df)
        if not status_cols:
            missing_status_columns.append(name)
            continue
        col = status_cols[0]
        normalized = df[col].map(normalize_status_label) if col in df.columns else pd.Series([], dtype=str)
        counts = Counter(normalized)
        # make sure totals include blank rows (Unknown)
        if len(df) and "Unknown" not in counts:
            counts["Unknown"] = len(df) - sum(counts.values())
        overall.update(counts)
        per_sheet[name] = counts
    return overall, per_sheet, missing_status_columns

def collect_upcoming_milestones(sheets: dict[str, pd.DataFrame], limit: int = 10) -> pd.DataFrame:
    records = []
    today = pd.Timestamp.now().normalize()
    for name, df in sheets.items():
        if not isinstance(df, pd.DataFrame) or df.empty:
            continue
        date_cols = [c for c in df.columns if "date" in str(c).lower() or "eta" in str(c).lower()]
        if not date_cols:
            continue
        label_candidates = [c for c in guess_key_columns(df) if c in df.columns] or list(df.columns)
        label_col = label_candidates[0] if label_candidates else None
        status_cols = guess_status_columns(df)
        status_col = status_cols[0] if status_cols else None
        for col in date_cols:
            parsed = pd.to_datetime(df[col], errors="coerce")
            for idx, date_val in parsed.dropna().items():
                item_label = (
                    str(df.at[idx, label_col]).strip()
                    if label_col and pd.notna(df.at[idx, label_col])
                    else f"Row {idx + 1}"
                )
                status_label = (
                    normalize_status_label(df.at[idx, status_col])
                    if status_col and status_col in df.columns
                    else ""
                )
                day_delta = (date_val.normalize() - today).days if not pd.isna(date_val) else None
                records.append(
                    {
                        "Sheet": name,
                        "Item": item_label or f"Row {idx + 1}",
                        "Date": date_val.date(),
                        "Status": status_label,
                        "Column": col,
                        "Days": day_delta,
                        "Bucket": "Overdue" if day_delta is not None and day_delta < 0 else "Upcoming",
                    }
                )
    if not records:
        return pd.DataFrame()
    events = pd.DataFrame(records).sort_values(["Bucket", "Date"])
    overdue = events[events["Bucket"] == "Overdue"].sort_values("Date")
    upcoming = events[events["Bucket"] == "Upcoming"].sort_values("Date")
    combined = pd.concat([overdue, upcoming], ignore_index=True)
    return combined.head(limit)

# ---------- Upload ----------
uploaded = st.file_uploader("Upload the Excel template", type=["xlsx"], accept_multiple_files=False)

if uploaded:
    sheets = load_excel(uploaded)
else:
    st.info("No file uploaded yet. You can start with a blank table: add a sheet name below.")
    default_name = st.text_input("New sheet name", value="Status")
    if st.button("Create blank sheet"):
        sheets = {default_name: pd.DataFrame(columns=["Project", "Owner", "Status", "Next Step", "ETA"])}
    else:
        sheets = {}

if not sheets:
    st.stop()

# ---------- Tabs (Scorecard + per sheet) ----------
tab_names = list(sheets.keys())
tabs = st.tabs(["Executive Scorecard", *tab_names])
scorecard_tab = tabs[0]
sheet_tabs = tabs[1:]

with scorecard_tab:
    st.subheader("Executive Scorecard")
    overall_counts, per_sheet_counts, missing_status = aggregate_status_counts(sheets)
    total_items = sum(overall_counts.values())
    if total_items:
        g_col, a_col, r_col, u_col = st.columns(4)
        g_col.metric("Green / On Track", overall_counts.get("Green", 0))
        a_col.metric("Amber / Watch", overall_counts.get("Amber", 0))
        r_col.metric("Red / At Risk", overall_counts.get("Red", 0))
        u_col.metric("Unknown / Blank", overall_counts.get("Unknown", 0))
        st.caption("Counts aggregate the first detected status column on each sheet.")
    else:
        st.info("Populate the status columns to see the executive roll-up.")

    if per_sheet_counts:
        rows = []
        for name, counts in per_sheet_counts.items():
            rows.append(
                {
                    "Sheet": name,
                    "Green": counts.get("Green", 0),
                    "Amber": counts.get("Amber", 0),
                    "Red": counts.get("Red", 0),
                    "On Hold": counts.get("On Hold", 0),
                    "Complete": counts.get("Complete", 0),
                    "Unknown": counts.get("Unknown", 0),
                    "Total": sum(counts.values()),
                }
            )
        summary_df = pd.DataFrame(rows).sort_values("Sheet")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

    milestones = collect_upcoming_milestones(sheets)
    if not milestones.empty:
        st.markdown("**Upcoming & Overdue Milestones**")
        st.dataframe(
            milestones,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Date": st.column_config.DateColumn("Date"),
                "Days": st.column_config.NumberColumn("Days From Today", format="%d"),
            },
        )
    else:
        st.info("Add date columns (e.g., 'Due Date', 'ETA') to surface upcoming milestones.")

    if missing_status:
        st.warning(
            "No status column detected on: " + ", ".join(missing_status) + ". "
            "Add a status field (e.g., 'Status' or 'RAG') so it rolls into the scorecard."
        )

for i, t in enumerate(sheet_tabs):
    name = tab_names[i]
    with t:
        st.subheader(f"Sheet: {name}")
        df = sheets[name]

        # Quick RAG preview if a likely status column exists
        likely = guess_status_columns(df)
        if likely:
            sc = likely[0]
            prev = df[[sc]].copy()
            prev["RAG"] = prev[sc].map(rag_emoji)
            with st.expander("RAG Preview", expanded=False):
                st.dataframe(prev, use_container_width=True)

        column_config = {}
        for col in likely:
            column_config[col] = st.column_config.SelectboxColumn(
                label=col,
                options=status_options_for(df, col),
                help="Pick a standard RAG value so the executive scorecard stays clean.",
            )

        edited = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            key=f"editor_{name}",
            column_config=column_config or None,
        )
        sheets[name] = edited

# ---------- Quality checks ----------
quality = st.sidebar
quality.header("Quality Checks")
overall_errors = False
overall_warnings = False

for idx, (sheet_name, df) in enumerate(sheets.items()):
    if not isinstance(df, pd.DataFrame):
        continue
    with quality.expander(sheet_name, expanded=(idx == 0 and len(sheets) <= 3)):
        if df.empty:
            st.info("Sheet is blank — skipping checks.")
            continue

        # Nulls on key columns
        keys = guess_key_columns(df)[:8]
        if keys:
            nulls = {k: int(df[k].isna().sum()) for k in keys if k in df.columns and df[k].isna().any()}
            if nulls:
                overall_errors = True
                st.error("Nulls detected in key fields: " + ", ".join([f"{k}: {v}" for k, v in nulls.items()]))
            else:
                st.success("Key fields look complete.")
        else:
            st.info("No obvious key columns detected (add Project, Owner, Status, etc.).")

        # Date sanity (if any column looks like a date)
        date_cols = [c for c in df.columns if "date" in str(c).lower() or "eta" in str(c).lower()]
        bad_dates = []
        for c in date_cols:
            try:
                parsed = pd.to_datetime(df[c], errors="coerce")
                n_bad = int(parsed.isna().sum()) if len(df) else 0
                if n_bad and df[c].notna().any():
                    bad_dates.append(f"{c} (~{n_bad} unparsable)")
            except Exception:
                bad_dates.append(f"{c} (parse error)")
        if bad_dates:
            overall_warnings = True
            st.warning("Date issues: " + ", ".join(bad_dates))
        elif date_cols:
            st.success("Date columns parse cleanly.")

if not overall_errors and not overall_warnings:
    quality.success("All checks passed across sheets.")

# ---------- Exports ----------
st.divider()
c1, c2, c3 = st.columns([1,1,2])
with c1:
    excel_bytes = to_excel_bytes(sheets)
    st.download_button(
        "⬇️ Download Excel Packet",
        data=excel_bytes,
        file_name=f"Program_Status_Export_{datetime.now():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with c2:
    # Export main sheet as CSV
    main_df = sheets[max(sheets, key=lambda s: sheets[s].shape[1])]
    st.download_button(
        "⬇️ Download CSV (Main Sheet)",
        data=main_df.to_csv(index=False).encode("utf-8"),
        file_name=f"Program_Status_{datetime.now():%Y-%m-%d}.csv",
        mime="text/csv",
        use_container_width=True,
    )

with c3:
    st.info("Tip: Keep consistent ‘Project’/‘Owner’ values week to week for clean deltas.")

# ---------- Optional: Snowflake sync (commented out) ----------
"""
# from snowflake.connector import connect
# def write_to_snowflake(df, table):
#     conn = connect(
#         account=st.secrets["snowflake"]["account"],
#         user=st.secrets["snowflake"]["user"],
#         password=st.secrets["snowflake"]["password"],
#         warehouse=st.secrets["snowflake"]["warehouse"],
#         database=st.secrets["snowflake"]["database"],
#         schema=st.secrets["snowflake"]["schema"],
#     )
#     cs = conn.cursor()
#     try:
#         # Create table if not exists (simple schema clone of current df)
#         cols = ", ".join([f'"{c}" VARCHAR' for c in df.columns])
#         cs.execute(f'CREATE TABLE IF NOT EXISTS "{table}" ({cols})')
#         success, nchunks, nrows, _ = cs.write_pandas(df, table_name=table, quote_identifiers=True)
#         st.success(f"Uploaded to Snowflake: {nrows} rows into {table}")
#     finally:
#         cs.close(); conn.close()
"""
