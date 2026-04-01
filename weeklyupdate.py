from __future__ import annotations

from datetime import datetime
from typing import Iterable, Mapping

import altair as alt
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Analyst Workflow Dashboard", layout="wide")
st.title("Analyst Workflow Dashboard")

WORKFLOW_COLUMNS = {
    "Request Type": ["type", "request type"],
    "Business": ["business"],
    "Division": ["division"],
    "Sector": ["sector"],
    "Vendor": ["vendor"],
    "Usage": ["usage"],
    "VA / Stocking Criteria": ["meets criteria", "va / stocking criteria", "stocking criteria"],
    "APL Indicator": ["compass apl", "apl indicator", "apl"],
    "Pantry Indicator": ["pantry", "pantry indicator"],
    "K12 Indicator": ["k12 apl", "k12 indicator", "k12"],
    "In CAT": ["in cat"],
    "On MOG": ["on mog"],
    "Cannot Add Not in Stock": ["cannot add not in stock", "if in stock: action"],
    "DIN": ["din"],
    "Conversion DIN": ["conversion din"],
    "Manufacturer": ["manufacturer"],
    "Brand": ["brand"],
    "Parent Category": ["parent category"],
    "Category": ["category"],
    "Subcategory": ["sub category", "subcategory"],
    "Reason for Request": ["reason for request"],
    "1x / Permanent": ["one-time or permanent", "1x / permanent indicator", "one time or permanent"],
    "Action": ["action"],
    "If In Stock Action": ["if in stock: action", "if in stock action"],
    "BuySmart Action": ["buysmart action", "buy smart action"],
    "Comments / Notes": ["comments / notes", "notes", "audit action"],
    "Assignment": ["assignment", "owner", "assigned to"],
    "Status": ["status"],
    "Exception Flag": ["exception flag", "exception"],
    "Rule Applied": ["rule applied"],
    "Override Reason": ["override reason"],
    "Case #": ["case#", "case #", "case"],
    "Date Created": ["date created"],
    "Unit Name": ["unit name"],
    "Item Description": ["description", "item description"],
}

ANONYMIZE_PREFIXES = {
    "Business": "Business",
    "Division": "Division",
    "Sector": "Sector",
    "Vendor": "Vendor",
    "Manufacturer": "Manufacturer",
    "Brand": "Brand",
    "Parent Category": "Parent Category",
    "Category": "Category",
    "Subcategory": "Subcategory",
    "Unit Name": "Unit",
}

REQUEST_BUCKET_ORDER = [
    "Mass Add",
    "PRF",
    "SORF",
    "SRF",
    "Already On MOG / Check Attribute",
    "Cannot Add Not in Stock",
    "Conversion DIN / Use Right",
    "1x request",
    "Permanent request",
    "Special exception / analyst review",
]

# ---------------------------------------------------------------------
# Data loading utilities
# ---------------------------------------------------------------------
def normalize_columns(df: pd.DataFrame, synonyms: Mapping[str, Iterable[str]]) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    lookup: dict[str, str] = {}
    for canonical, aliases in synonyms.items():
        lookup[canonical.lower()] = canonical
        for alias in aliases:
            lookup[str(alias).lower()] = canonical
    for col in df.columns:
        key = str(col).strip().lower()
        if key in lookup:
            rename_map[col] = lookup[key]
    df = df.rename(columns=rename_map)
    return df


def ensure_columns(df: pd.DataFrame, required: Iterable[str]) -> pd.DataFrame:
    for col in required:
        if col not in df.columns:
            df[col] = pd.NA
    ordered = list(required) + [c for c in df.columns if c not in required]
    return df[ordered]


def sanitize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def normalize_percent(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def normalize_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def anonymize_series(series: pd.Series, prefix: str) -> pd.Series:
    cleaned = series.fillna("").astype(str).str.strip()
    unique_values = [value for value in cleaned.unique().tolist() if value]
    replacement_map = {value: f"{prefix} {index:02d}" for index, value in enumerate(unique_values, start=1)}
    return cleaned.map(replacement_map).replace("", pd.NA)


def anonymize_numeric_identifier(series: pd.Series, prefix: str) -> pd.Series:
    cleaned = series.fillna("").astype(str).str.strip()
    unique_values = [value for value in cleaned.unique().tolist() if value]
    replacement_map = {value: f"{prefix}-{index:05d}" for index, value in enumerate(unique_values, start=1)}
    return cleaned.map(replacement_map).replace("", pd.NA)


def anonymize_free_text(series: pd.Series, prefix: str) -> pd.Series:
    cleaned = series.fillna("").astype(str).str.strip()
    masked = cleaned.where(cleaned.eq(""), f"{prefix} redacted")
    return masked.replace("", pd.NA)


def anonymize_workflow_data(workflow: pd.DataFrame) -> pd.DataFrame:
    masked = workflow.copy()

    for column, prefix in ANONYMIZE_PREFIXES.items():
        if column in masked.columns:
            masked[column] = anonymize_series(masked[column], prefix)

    if "Case #" in masked.columns:
        masked["Case #"] = anonymize_numeric_identifier(masked["Case #"], "CASE")
    if "DIN" in masked.columns:
        masked["DIN"] = anonymize_numeric_identifier(masked["DIN"], "DIN")
    if "Conversion DIN" in masked.columns:
        masked["Conversion DIN"] = anonymize_numeric_identifier(masked["Conversion DIN"], "CONV")
    if "Item Description" in masked.columns:
        masked["Item Description"] = anonymize_free_text(masked["Item Description"], "Item description")
    if "Reason for Request" in masked.columns:
        masked["Reason for Request"] = anonymize_free_text(masked["Reason for Request"], "Request reason")
    if "Comments / Notes" in masked.columns:
        masked["Comments / Notes"] = anonymize_free_text(masked["Comments / Notes"], "Analyst note")
    if "Override Reason" in masked.columns:
        masked["Override Reason"] = anonymize_free_text(masked["Override Reason"], "Override reason")

    return masked


def classify_request_bucket(row: pd.Series) -> str:
    request_type = str(row.get("Request Type", "") or "").strip().upper()
    on_mog = str(row.get("On MOG", "") or "").strip().upper()
    cannot_add = str(row.get("Cannot Add Not in Stock", "") or "").strip().lower()
    conversion_din = str(row.get("Conversion DIN", "") or "").strip()
    permanence = str(row.get("1x / Permanent", "") or "").strip().lower()
    exception_flag = str(row.get("Exception Flag", "") or "").strip().upper()
    status = str(row.get("Status", "") or "").strip().lower()
    buy_smart = str(row.get("BuySmart Action", "") or "").strip().lower()

    if "mass add" in request_type or request_type == "MASS ADD":
        return "Mass Add"
    if on_mog == "Y":
        return "Already On MOG / Check Attribute"
    if "cannot add" in cannot_add or "not in stock" in cannot_add:
        return "Cannot Add Not in Stock"
    if conversion_din:
        return "Conversion DIN / Use Right"
    if permanence == "one-time":
        return "1x request"
    if permanence == "permanent":
        return "Permanent request"
    if exception_flag == "Y" or "exception" in status or "review" in status or "exception" in buy_smart:
        return "Special exception / analyst review"
    if request_type == "PRF":
        return "PRF"
    if request_type == "SORF":
        return "SORF"
    if request_type == "SRF":
        return "SRF"
    return "Special exception / analyst review"


def load_workflow_sheet(uploaded_file) -> pd.DataFrame:
    uploaded_file.seek(0)
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = xls.sheet_names[0]
    workflow = sanitize_dataframe(xls.parse(sheet_name))
    workflow = normalize_columns(workflow, WORKFLOW_COLUMNS)
    workflow = ensure_columns(workflow, WORKFLOW_COLUMNS.keys())
    workflow = anonymize_workflow_data(workflow)
    uploaded_file.seek(0)
    return workflow


def build_workflow_seed(
    workflow_source: pd.DataFrame | None,
    reporting_date: datetime.date,
) -> pd.DataFrame:
    if workflow_source is not None and not workflow_source.empty:
        workflow = workflow_source.copy()
    else:
        workflow = pd.DataFrame(
            {
                "Request Type": ["PRF", "SORF", "PRF", "SRF"],
                "Business": ["Compass USA", "Compass USA", "Compass USA", "Compass USA"],
                "Division": ["HR", "Foodbuy", "HR", "Education"],
                "Sector": ["Chartwells Sector", "Canteen", "Morrison", "Chartwells K12"],
                "Vendor": ["Sysco Lincoln", "US Foods", "Performance Food Group", "Sysco Metro NY - Ritter"],
                "Usage": [2, 18, 6, 10],
                "VA / Stocking Criteria": ["Y", "Y", "N", "Y"],
                "APL Indicator": ["", "Y", "Y", ""],
                "Pantry Indicator": ["", "", "Y", ""],
                "K12 Indicator": ["", "", "", "Y"],
                "In CAT": ["Y", "N", "N", "Y"],
                "On MOG": ["", "Y", "", ""],
                "Cannot Add Not in Stock": ["", "Cannot Add. Not in Stock.", "", ""],
                "DIN": ["7145938", "883746", "462281", "73550"],
                "Conversion DIN": ["", "991188", "", ""],
                "Manufacturer": ["Sunset Foods Ltd", "Ventura Foods", "Kraft Heinz", "Gordon Foodservice"],
                "Brand": ["Boyles", "LouAna", "Heinz", "Mrs. Friday's"],
                "Parent Category": ["Protein", "Oils", "Condiments", "Frozen Seafood"],
                "Category": ["Beef", "Pantry", "Sauces", "Seafood"],
                "Subcategory": ["Breaded Beef", "Cooking Oil", "Ketchup", "Crab Cakes"],
                "Reason for Request": [
                    "Frequent catering request",
                    "Distributor request for stocking alignment",
                    "Unit requested local substitution",
                    "Cycle menu week one and week six",
                ],
                "1x / Permanent": ["Permanent", "Permanent", "One-Time", "One-Time"],
                "Action": ["OK", "Review", "Override", "OK"],
                "If In Stock Action": ["", "Use stocked item", "", "OK"],
                "BuySmart Action": ["Approved", "Pending Review", "Exception", "Denied"],
                "Comments / Notes": ["", "", "Awaiting category lead approval", ""],
                "Assignment": ["Analyst Queue", "Stocking Team", "Category Manager", "Analyst Queue"],
                "Status": ["New", "In Review", "Exception Review", "Closed"],
                "Exception Flag": ["N", "N", "Y", "N"],
                "Rule Applied": ["", "Stocking criteria met", "Manual override required", ""],
                "Override Reason": ["", "", "Local menu commitment", ""],
                "Case #": ["WO0000001", "WO0000002", "WO0000003", "WO0000004"],
                "Date Created": pd.to_datetime(
                    [reporting_date, reporting_date, reporting_date, reporting_date]
                ),
                "Unit Name": ["Mid-Plains CC", "Canteen East", "Morrison South", "Rutgers Newark"],
                "Item Description": [
                    "Breaded beef thin slice raw",
                    "Canola oil 35 lb",
                    "Tomato ketchup pouches",
                    "Crab cake patty",
                ],
            }
        )

    workflow["Date Created"] = normalize_date(workflow["Date Created"])
    workflow["Usage"] = normalize_percent(workflow["Usage"])
    workflow["Selected"] = False
    workflow["Last Saved"] = pd.Timestamp(reporting_date)
    workflow["Last Sync"] = pd.Timestamp(reporting_date)
    workflow["Queue Bucket"] = workflow["BuySmart Action"].fillna("Unassigned").replace("", "Unassigned")
    workflow["Request Bucket"] = workflow.apply(classify_request_bucket, axis=1)
    workflow["Needs Review"] = workflow["Status"].fillna("").astype(str).str.contains("review", case=False) | workflow[
        "Exception Flag"
    ].fillna("").astype(str).eq("Y")
    workflow["Analyst Notes"] = workflow["Comments / Notes"].fillna("")

    ordered_columns = [
        "Selected",
        "Case #",
        "Request Type",
        "Business",
        "Division",
        "Sector",
        "Vendor",
        "Unit Name",
        "Item Description",
        "Usage",
        "VA / Stocking Criteria",
        "APL Indicator",
        "Pantry Indicator",
        "K12 Indicator",
        "In CAT",
        "On MOG",
        "Cannot Add Not in Stock",
        "DIN",
        "Conversion DIN",
        "Manufacturer",
        "Brand",
        "Parent Category",
        "Category",
        "Subcategory",
        "Reason for Request",
        "1x / Permanent",
        "Action",
        "If In Stock Action",
        "BuySmart Action",
        "Analyst Notes",
        "Assignment",
        "Status",
        "Exception Flag",
        "Rule Applied",
        "Override Reason",
        "Date Created",
        "Queue Bucket",
        "Request Bucket",
        "Needs Review",
        "Last Sync",
        "Last Saved",
    ]
    return ensure_columns(workflow, ordered_columns)[ordered_columns]


def apply_workflow_quick_action(workflow_df: pd.DataFrame, action: str, reporting_date: datetime.date) -> tuple[pd.DataFrame, int]:
    updated = workflow_df.copy()
    selected_mask = updated["Selected"].fillna(False)
    affected = int(selected_mask.sum())

    if action == "Assign to analyst queue":
        updated.loc[selected_mask, "Assignment"] = "Analyst Queue"
        updated.loc[selected_mask, "Status"] = "In Review"
        updated.loc[selected_mask, "Queue Bucket"] = "Analyst Queue"
        updated.loc[selected_mask, "Rule Applied"] = "Assigned by queue rule"
    elif action == "Flag as exception":
        updated.loc[selected_mask, "Exception Flag"] = "Y"
        updated.loc[selected_mask, "Status"] = "Exception Review"
        updated.loc[selected_mask, "Needs Review"] = True
        updated.loc[selected_mask, "Rule Applied"] = "Exception routing"
    elif action == "Recommend stocked alternative":
        updated.loc[selected_mask, "If In Stock Action"] = "Use stocked item"
        updated.loc[selected_mask, "BuySmart Action"] = "Conversion Recommended"
        updated.loc[selected_mask, "Rule Applied"] = "Stocked alternative suggestion"
    elif action == "Approve selected requests":
        updated.loc[selected_mask, "BuySmart Action"] = "Approved"
        updated.loc[selected_mask, "Status"] = "Ready to Save"
        updated.loc[selected_mask, "Needs Review"] = False
        updated.loc[selected_mask, "Rule Applied"] = "Approval automation"

    updated["Request Bucket"] = updated.apply(classify_request_bucket, axis=1)
    updated.loc[selected_mask, "Last Sync"] = pd.Timestamp(reporting_date)
    return updated, affected
def render_workflow_dashboard(reporting_date: datetime.date, uploaded_file) -> None:
    st.subheader("Workflow Dashboard")
    st.caption(
        "Analyst-facing queue for recurring request maintenance: editable rows, bulk quick actions, draft saves, and workload rollups."
    )

    workflow_source = None
    source_key = "sample_workflow"
    if uploaded_file is not None:
        workflow_source = load_workflow_sheet(uploaded_file)
        source_key = uploaded_file.name

    if st.session_state.get("workflow_source_key") != source_key:
        st.session_state.workflow_drafts = build_workflow_seed(workflow_source, reporting_date)
        st.session_state.workflow_saved = st.session_state.workflow_drafts.copy()
        st.session_state.workflow_last_save = pd.Timestamp(reporting_date)
        st.session_state.workflow_action_count = 0
        st.session_state.workflow_source_key = source_key
    if "workflow_drafts" not in st.session_state:
        st.session_state.workflow_drafts = build_workflow_seed(workflow_source, reporting_date)
    if "workflow_saved" not in st.session_state:
        st.session_state.workflow_saved = st.session_state.workflow_drafts.copy()
    if "workflow_last_save" not in st.session_state:
        st.session_state.workflow_last_save = pd.Timestamp(reporting_date)
    if "workflow_action_count" not in st.session_state:
        st.session_state.workflow_action_count = 0

    workflow_df = st.session_state.workflow_drafts.copy()
    workflow_df["Request Bucket"] = workflow_df.apply(classify_request_bucket, axis=1)

    top_cols = st.columns([1.3, 1.1, 1.1, 1.1])
    selected_rows = int(workflow_df["Selected"].fillna(False).sum())
    review_rows = int(workflow_df["Needs Review"].fillna(False).sum())
    exception_rows = int(workflow_df["Exception Flag"].fillna("").astype(str).eq("Y").sum())
    top_cols[0].metric("Queue volume", len(workflow_df), delta=f"{review_rows} need review")
    top_cols[1].metric("Selected rows", selected_rows, delta=f"{exception_rows} exceptions")
    top_cols[2].metric("Automated updates", st.session_state.workflow_action_count)
    top_cols[3].metric("Last save", st.session_state.workflow_last_save.strftime("%b %d, %I:%M %p"))

    dashboard_left, dashboard_right = st.columns([1.2, 0.8])
    queue_rollup = (
        workflow_df["Request Bucket"]
        .value_counts()
        .reindex(REQUEST_BUCKET_ORDER, fill_value=0)
        .rename_axis("Request Bucket")
        .reset_index(name="Rows")
    )
    queue_chart = (
        alt.Chart(queue_rollup)
        .mark_bar()
        .encode(
            x=alt.X("Rows:Q", title="Rows"),
            y=alt.Y("Request Bucket:N", sort=REQUEST_BUCKET_ORDER),
            color=alt.Color("Request Bucket:N", legend=None, scale=alt.Scale(scheme="tableau20")),
            tooltip=["Request Bucket", "Rows"],
        )
    )
    dashboard_left.altair_chart(queue_chart, use_container_width=True)

    dashboard_right.markdown(
        """
        **Request buckets**
        - Mass Add
        - PRF / SORF / SRF
        - Already On MOG / Check Attribute
        - Cannot Add Not in Stock
        - Conversion DIN / Use Right
        - 1x request / Permanent request
        - Special exception / analyst review
        """
    )

    control_cols = st.columns([1.1, 1.1, 0.8, 0.8])
    action = control_cols[0].selectbox(
        "Quick action",
        options=[
            "Assign to analyst queue",
            "Flag as exception",
            "Recommend stocked alternative",
            "Approve selected requests",
        ],
    )
    control_cols[1].caption("Use the checkbox column to target rows before running an automation.")
    apply_clicked = control_cols[2].button("Run quick action", use_container_width=True)
    save_clicked = control_cols[3].button("Save changes", type="primary", use_container_width=True)

    edited_workflow = st.data_editor(
        workflow_df,
        use_container_width=True,
        hide_index=True,
        height=520,
        column_config={
            "Selected": st.column_config.CheckboxColumn("Select"),
            "Case #": st.column_config.TextColumn("Case #", disabled=True),
            "Request Type": st.column_config.SelectboxColumn("Request Type", options=["PRF", "SORF", "SRF"]),
            "Usage": st.column_config.NumberColumn("Usage", min_value=0, format="%d"),
            "VA / Stocking Criteria": st.column_config.SelectboxColumn("VA / Stocking Criteria", options=["", "Y", "N"]),
            "APL Indicator": st.column_config.SelectboxColumn("APL Indicator", options=["", "Y", "N"]),
            "Pantry Indicator": st.column_config.SelectboxColumn("Pantry Indicator", options=["", "Y", "N"]),
            "K12 Indicator": st.column_config.SelectboxColumn("K12 Indicator", options=["", "Y", "N"]),
            "In CAT": st.column_config.SelectboxColumn("In CAT", options=["", "Y", "N"]),
            "On MOG": st.column_config.SelectboxColumn("On MOG", options=["", "Y", "N"]),
            "1x / Permanent": st.column_config.SelectboxColumn("1x / Permanent", options=["One-Time", "Permanent", ""]),
            "Action": st.column_config.TextColumn("Action"),
            "If In Stock Action": st.column_config.TextColumn("If In Stock Action"),
            "BuySmart Action": st.column_config.TextColumn("BuySmart Action"),
            "Analyst Notes": st.column_config.TextColumn("Comments / Notes"),
            "Assignment": st.column_config.TextColumn("Assignment"),
            "Status": st.column_config.TextColumn("Status"),
            "Exception Flag": st.column_config.SelectboxColumn("Exception Flag", options=["N", "Y", ""]),
            "Needs Review": st.column_config.CheckboxColumn("Needs Review"),
            "Date Created": st.column_config.DateColumn("Date Created", disabled=True),
            "Last Sync": st.column_config.DateColumn("Last Sync", disabled=True),
            "Last Saved": st.column_config.DateColumn("Last Saved", disabled=True),
        },
        disabled=["Case #", "Date Created", "Last Sync", "Last Saved"],
    )

    edited_workflow["Request Bucket"] = edited_workflow.apply(classify_request_bucket, axis=1)
    st.session_state.workflow_drafts = edited_workflow.copy()

    if apply_clicked:
        updated_workflow, affected = apply_workflow_quick_action(st.session_state.workflow_drafts, action, reporting_date)
        st.session_state.workflow_drafts = updated_workflow
        st.session_state.workflow_action_count += affected
        st.success(f"{action} applied to {affected} row(s).")
        st.rerun()

    if save_clicked:
        saved_workflow = st.session_state.workflow_drafts.copy()
        saved_workflow["Last Saved"] = pd.Timestamp(datetime.now())
        st.session_state.workflow_drafts = saved_workflow
        st.session_state.workflow_saved = saved_workflow.copy()
        st.session_state.workflow_last_save = pd.Timestamp(datetime.now())
        st.success("Draft changes saved for this session.")
        st.rerun()

    saved_export = st.session_state.workflow_saved.copy()
    st.download_button(
        "Download workflow draft (CSV)",
        data=saved_export.to_csv(index=False).encode("utf-8"),
        file_name=f"workflow_wireframe_{reporting_date:%Y%m%d}.csv",
        mime="text/csv",
        use_container_width=True,
    )


# ---------------------------------------------------------------------
# App controls and layout
# ---------------------------------------------------------------------
uploaded_file = st.file_uploader("Upload analyst workflow workbook (Excel)", type=["xlsx"])

with st.sidebar:
    st.header("Guidance")
    st.caption("Upload the analyst request workbook to edit cases, run quick actions, and export a revised draft.")
    st.markdown(
        """
        **Primary workflow fields**
        - Request type, business, division, sector, vendor, usage
        - VA / stocking criteria, APL, pantry, K12, In CAT, On MOG
        - DIN and conversion DIN
        - Category, subcategory, BuySmart action, analyst notes
        - Assignment, status, exception flag, rule applied, override reason
        """
    )
    reporting_date = st.date_input("Reporting as of", datetime.now().date())

if uploaded_file is None:
    st.info("Using sample workflow data. Upload the analyst workbook to work with live request records.")

render_workflow_dashboard(reporting_date, uploaded_file)
