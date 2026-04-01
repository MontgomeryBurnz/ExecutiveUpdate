from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Iterable, Mapping

import altair as alt
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Program Executive Scorecard", layout="wide")
st.title("Program Executive Scorecard")


@dataclass
class ScorecardData:
    portfolio: pd.DataFrame
    milestones: pd.DataFrame
    risks: pd.DataFrame
    source_name: str


PORTFOLIO_COLUMNS = {
    "Initiative": ["project", "project name", "initiative name"],
    "Workstream": ["pillar", "tower", "portfolio"],
    "Owner": ["lead", "manager", "pm"],
    "Health": ["rag", "status", "overall status"],
    "Percent Complete": ["percent complete", "% complete", "progress", "progress (%)"],
    "Launch Date": ["start date", "kickoff"],
    "Target Date": ["end date", "go-live", "eta", "due date"],
    "Budget": ["planned spend", "allocated budget", "budget (usd)"],
    "Actual Spend": ["actuals", "actual spend", "spent"],
    "Status Summary": ["executive summary", "headline", "status notes"],
}

MILESTONE_COLUMNS = {
    "Initiative": ["project", "project name"],
    "Milestone": ["milestone name", "deliverable"],
    "Target Date": ["due date", "eta", "planned date"],
    "Status": ["status", "rag"],
    "Owner": ["lead", "manager", "pm"],
    "Notes": ["comments", "context"],
}

RISK_COLUMNS = {
    "Initiative": ["project", "project name"],
    "Risk": ["risk description", "risk statement"],
    "Impact": ["impact level", "impact rating"],
    "Probability": ["likelihood", "probability level"],
    "Status": ["status", "rag"],
    "Mitigation": ["mitigation plan", "response"],
    "Owner": ["risk owner", "owner"],
}

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

STATUS_ORDER = [
    "At Risk",
    "Watch",
    "On Hold",
    "Not Started",
    "On Track",
    "Complete",
    "Unknown",
]

STATUS_COLORS = {
    "At Risk": "#D64541",
    "Watch": "#E3B341",
    "On Hold": "#8E44AD",
    "Not Started": "#95A5A6",
    "On Track": "#2ECC71",
    "Complete": "#1ABC9C",
    "Unknown": "#BDC3C7",
}

RISK_LEVEL_MAP = {
    "low": 1,
    "medium": 2,
    "med": 2,
    "high": 3,
    "critical": 3,
}


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


def load_scorecard_from_excel(uploaded_file) -> ScorecardData:
    xls = pd.ExcelFile(uploaded_file)
    portfolio = pd.DataFrame()
    milestones = pd.DataFrame()
    risks = pd.DataFrame()

    if "Portfolio" in xls.sheet_names:
        portfolio = sanitize_dataframe(xls.parse("Portfolio"))
        portfolio = normalize_columns(portfolio, PORTFOLIO_COLUMNS)
        portfolio = ensure_columns(portfolio, PORTFOLIO_COLUMNS.keys())
    else:
        st.warning("No 'Portfolio' tab found. Using an empty portfolio sheet.")
        portfolio = pd.DataFrame(columns=list(PORTFOLIO_COLUMNS.keys()))

    if "Milestones" in xls.sheet_names:
        milestones = sanitize_dataframe(xls.parse("Milestones"))
        milestones = normalize_columns(milestones, MILESTONE_COLUMNS)
        milestones = ensure_columns(milestones, MILESTONE_COLUMNS.keys())
    else:
        milestones = pd.DataFrame(columns=list(MILESTONE_COLUMNS.keys()))

    if "Risks" in xls.sheet_names:
        risks = sanitize_dataframe(xls.parse("Risks"))
        risks = normalize_columns(risks, RISK_COLUMNS)
        risks = ensure_columns(risks, RISK_COLUMNS.keys())
    else:
        risks = pd.DataFrame(columns=list(RISK_COLUMNS.keys()))

    return ScorecardData(portfolio=portfolio, milestones=milestones, risks=risks, source_name=uploaded_file.name)


def load_sample_scorecard() -> ScorecardData:
    today = pd.Timestamp(datetime.now().date())
    portfolio = pd.DataFrame(
        {
            "Initiative": [
                "Digital Onboarding",
                "Customer 360 Rollout",
                "Data Warehouse Migration",
                "Field Mobile App",
                "Billing Modernization",
            ],
            "Workstream": ["Experience", "CRM", "Data", "Operations", "Finance"],
            "Owner": ["A. Lopez", "K. Chen", "R. Patel", "J. Gomez", "S. Woods"],
            "Health": ["Green", "Amber", "Red", "Green", "Amber"],
            "Percent Complete": [72, 58, 41, 86, 49],
            "Launch Date": [
                today - pd.Timedelta(days=120),
                today - pd.Timedelta(days=95),
                today - pd.Timedelta(days=60),
                today - pd.Timedelta(days=140),
                today - pd.Timedelta(days=80),
            ],
            "Target Date": [
                today + pd.Timedelta(days=40),
                today + pd.Timedelta(days=25),
                today + pd.Timedelta(days=70),
                today + pd.Timedelta(days=20),
                today + pd.Timedelta(days=55),
            ],
            "Budget": [550_000, 430_000, 620_000, 210_000, 390_000],
            "Actual Spend": [310_000, 260_000, 180_000, 150_000, 210_000],
            "Status Summary": [
                "MVP in pilot; adoption tracking 15% above target.",
                "Integration testing slipping; vendor patch due Friday.",
                "Blockers on data quality gating migration cut-over.",
                "Field feedback positive; higher hardware costs under review.",
                "Dependency on pricing APIs; replan under way with Finance.",
            ],
        }
    )

    milestones = pd.DataFrame(
        {
            "Initiative": [
                "Digital Onboarding",
                "Digital Onboarding",
                "Customer 360 Rollout",
                "Data Warehouse Migration",
                "Field Mobile App",
                "Billing Modernization",
            ],
            "Milestone": [
                "Self-service KYC live",
                "Paper process sunset",
                "CRM go-live (Phase 1)",
                "Production cut-over",
                "iOS field release",
                "Pricing rules migration",
            ],
            "Target Date": [
                today + pd.Timedelta(days=15),
                today + pd.Timedelta(days=45),
                today + pd.Timedelta(days=25),
                today + pd.Timedelta(days=70),
                today + pd.Timedelta(days=10),
                today - pd.Timedelta(days=5),
            ],
            "Status": ["Green", "Amber", "Amber", "Red", "Green", "Red"],
            "Owner": ["A. Lopez", "A. Lopez", "K. Chen", "R. Patel", "J. Gomez", "S. Woods"],
            "Notes": [
                "Beta conversion at 92%. Final analytics QA this week.",
                "Dependency on contact center training completion.",
                "Waiting on reference data clean-up from Data team.",
                "Load testing blocked by missing anonymized dataset.",
                "Final sprint in progress; hardware rollout scheduled.",
                "Legacy platform defects delaying migration window.",
            ],
        }
    )

    risks = pd.DataFrame(
        {
            "Initiative": [
                "Customer 360 Rollout",
                "Data Warehouse Migration",
                "Billing Modernization",
                "Field Mobile App",
            ],
            "Risk": [
                "Vendor CRM patch may slip, delaying UAT sign-off.",
                "Data quality issues could extend migration blackout.",
                "Pricing API throughput may not meet performance targets.",
                "Offline mode stability risk in low-connectivity regions.",
            ],
            "Impact": ["High", "High", "Medium", "Medium"],
            "Probability": ["Medium", "High", "Medium", "Low"],
            "Status": ["Mitigating", "Open", "Watching", "Mitigating"],
            "Mitigation": [
                "Escalated with vendor; tracking hotfix build.",
                "Profiling sprint and cleansing backlog with owners.",
                "Load test this sprint; scale-out plan agreed with Infra.",
                "Targeted field pilots with telemetry instrumentation.",
            ],
            "Owner": ["K. Chen", "R. Patel", "S. Woods", "J. Gomez"],
        }
    )

    return ScorecardData(
        portfolio=portfolio,
        milestones=milestones,
        risks=risks,
        source_name="Sample Program Scorecard",
    )


def create_template_workbook() -> bytes:
    sample = load_sample_scorecard()
    return to_excel_bytes(sample.portfolio, sample.milestones, sample.risks)


# ---------------------------------------------------------------------
# Transformation helpers
# ---------------------------------------------------------------------
def categorize_health(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "Unknown"
    text = str(value).strip().lower()
    if text in {"green", "on track", "good"}:
        return "On Track"
    if text in {"yellow", "amber", "watch", "caution"}:
        return "Watch"
    if text in {"red", "critical", "off track", "blocked"}:
        return "At Risk"
    if "hold" in text:
        return "On Hold"
    if "complete" in text or "done" in text or "closed" in text:
        return "Complete"
    if "not started" in text or text == "tbd":
        return "Not Started"
    return "Unknown"


def normalize_percent(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def normalize_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def normalize_risk_level(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "Unknown"
    text = str(value).strip().lower()
    if text.isdigit():
        score = int(text)
        if score <= 1:
            return "Low"
        if score == 2:
            return "Medium"
        return "High"
    for label in ["low", "medium", "high"]:
        if label in text:
            return label.capitalize()
    if "critical" in text or "severe" in text:
        return "High"
    return "Unknown"


def risk_score(level: str) -> int:
    return RISK_LEVEL_MAP.get(level.lower(), 0)


def format_currency(amount: float | int | None) -> str:
    if amount is None or pd.isna(amount):
        return "-"
    return f"${amount:,.0f}"


def to_excel_bytes(portfolio: pd.DataFrame, milestones: pd.DataFrame, risks: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        portfolio.to_excel(writer, sheet_name="Portfolio", index=False)
        milestones.to_excel(writer, sheet_name="Milestones", index=False)
        risks.to_excel(writer, sheet_name="Risks", index=False)
    output.seek(0)
    return output.getvalue()


def load_workflow_sheet(uploaded_file) -> pd.DataFrame:
    uploaded_file.seek(0)
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = xls.sheet_names[0]
    workflow = sanitize_dataframe(xls.parse(sheet_name))
    workflow = normalize_columns(workflow, WORKFLOW_COLUMNS)
    workflow = ensure_columns(workflow, WORKFLOW_COLUMNS.keys())
    uploaded_file.seek(0)
    return workflow


def build_workflow_seed(
    workflow_source: pd.DataFrame | None,
    portfolio: pd.DataFrame,
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

    updated.loc[selected_mask, "Last Sync"] = pd.Timestamp(reporting_date)
    return updated, affected


def render_scorecard_view(
    data: ScorecardData,
    reporting_date: datetime.date,
    selected_workstreams: list[str],
    selected_owners: list[str],
    selected_health: list[str],
    include_complete: bool,
    lookahead_days: int,
) -> None:
    portfolio = data.portfolio.copy()
    milestones = data.milestones.copy()
    risks = data.risks.copy()

    portfolio["Health Category"] = portfolio["Health"].map(categorize_health)
    portfolio["Percent Complete"] = normalize_percent(portfolio["Percent Complete"])
    portfolio["Launch Date"] = normalize_date(portfolio["Launch Date"])
    portfolio["Target Date"] = normalize_date(portfolio["Target Date"])

    milestones["Target Date"] = normalize_date(milestones["Target Date"])
    milestones["Status Category"] = milestones["Status"].map(categorize_health)

    risks["Impact Level"] = risks["Impact"].map(normalize_risk_level)
    risks["Probability Level"] = risks["Probability"].map(normalize_risk_level)
    risks["Severity Score"] = risks["Impact Level"].map(risk_score) * risks["Probability Level"].map(risk_score)

    filtered = portfolio.copy()
    if selected_workstreams:
        filtered = filtered[filtered["Workstream"].isin(selected_workstreams)]
    if selected_owners:
        filtered = filtered[filtered["Owner"].isin(selected_owners)]
    if selected_health:
        filtered = filtered[filtered["Health Category"].isin(selected_health)]
    if not include_complete:
        filtered = filtered[filtered["Health Category"] != "Complete"]

    milestones_filtered = milestones[milestones["Initiative"].isin(filtered["Initiative"])].copy()
    risks_filtered = risks[risks["Initiative"].isin(filtered["Initiative"])].copy()

    today = pd.Timestamp(reporting_date)
    milestones_filtered["Days From Today"] = (milestones_filtered["Target Date"] - today).dt.days
    upcoming_milestones = milestones_filtered[
        (milestones_filtered["Days From Today"] >= 0)
        & (milestones_filtered["Days From Today"] <= lookahead_days)
    ].sort_values("Target Date")
    overdue_milestones = milestones_filtered[milestones_filtered["Days From Today"] < 0].sort_values("Target Date")

    active_initiatives = len(filtered)
    at_risk_count = int((filtered["Health Category"] == "At Risk").sum())
    watch_count = int((filtered["Health Category"] == "Watch").sum())
    hold_count = int((filtered["Health Category"] == "On Hold").sum())

    avg_progress = filtered["Percent Complete"].dropna().mean()
    budget_total = pd.to_numeric(filtered["Budget"], errors="coerce").sum()
    actual_total = pd.to_numeric(filtered["Actual Spend"], errors="coerce").sum()
    budget_burn = actual_total / budget_total if budget_total else None

    metric_columns = st.columns(4)
    metric_columns[0].metric(
        "Active initiatives",
        active_initiatives,
        delta=f"{at_risk_count} at risk" if at_risk_count else "All healthy",
    )
    metric_columns[1].metric("Watch / Hold", f"{watch_count} watch | {hold_count} hold")
    metric_columns[2].metric(
        "Avg progress",
        f"{avg_progress:.0f}%" if not pd.isna(avg_progress) else "—",
        delta="vs. 100% target",
    )
    metric_columns[3].metric(
        "Budget burn",
        f"{budget_burn:.0%}" if budget_burn is not None else "—",
        delta=f"{format_currency(actual_total)} of {format_currency(budget_total)}",
    )

    st.caption(f"Reporting as of {reporting_date:%B %d, %Y} · Source: {data.source_name}")

    distribution_cols = st.columns(2)
    health_counts = (
        filtered["Health Category"]
        .value_counts()
        .reindex(STATUS_ORDER, fill_value=0)
        .rename_axis("Health Category")
        .reset_index(name="Count")
    )
    health_chart = (
        alt.Chart(health_counts)
        .mark_bar()
        .encode(
            y=alt.Y("Health Category:N", sort=STATUS_ORDER),
            x=alt.X("Count:Q", title="Initiatives"),
            color=alt.Color(
                "Health Category:N",
                scale=alt.Scale(domain=list(STATUS_COLORS.keys()), range=list(STATUS_COLORS.values())),
                legend=None,
            ),
            tooltip=["Health Category", "Count"],
        )
    )
    distribution_cols[0].altair_chart(health_chart, use_container_width=True)

    progress_by_workstream = (
        filtered.groupby("Workstream", dropna=False)["Percent Complete"]
        .mean()
        .reset_index()
        .sort_values("Percent Complete", ascending=True)
    )
    progress_chart = (
        alt.Chart(progress_by_workstream)
        .mark_bar()
        .encode(
            y=alt.Y("Workstream:N", title="Workstream"),
            x=alt.X("Percent Complete:Q", title="Avg % complete"),
            tooltip=["Workstream", alt.Tooltip("Percent Complete:Q", format=".0f")],
        )
    )
    distribution_cols[1].altair_chart(progress_chart, use_container_width=True)

    milestone_col1, milestone_col2 = st.columns(2)
    milestone_col1.subheader("Upcoming milestones")
    if not upcoming_milestones.empty:
        milestone_col1.dataframe(
            upcoming_milestones[
                ["Initiative", "Milestone", "Target Date", "Status Category", "Owner", "Days From Today", "Notes"]
            ],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Target Date": st.column_config.DateColumn("Target Date"),
                "Days From Today": st.column_config.NumberColumn("Days", format="%d"),
            },
        )
    else:
        milestone_col1.info("No milestones in the selected window.")

    milestone_col2.subheader("Overdue milestones")
    if not overdue_milestones.empty:
        milestone_col2.dataframe(
            overdue_milestones[
                ["Initiative", "Milestone", "Target Date", "Status Category", "Owner", "Days From Today", "Notes"]
            ],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Target Date": st.column_config.DateColumn("Target Date"),
                "Days From Today": st.column_config.NumberColumn("Days overdue", format="%d"),
            },
        )
    else:
        milestone_col2.success("No overdue milestones.")

    st.subheader("Risk posture")
    risk_columns = st.columns(2)
    if not risks_filtered.empty:
        heatmap_data = risks_filtered.groupby(["Impact Level", "Probability Level"]).size().reset_index(name="Count")
        heatmap_chart = (
            alt.Chart(heatmap_data)
            .mark_rect()
            .encode(
                x=alt.X("Probability Level:N", sort=["Low", "Medium", "High", "Unknown"]),
                y=alt.Y("Impact Level:N", sort=["Low", "Medium", "High", "Unknown"]),
                color=alt.Color("Count:Q", scale=alt.Scale(scheme="orangered")),
                tooltip=["Impact Level", "Probability Level", "Count"],
            )
        )
        risk_columns[0].altair_chart(heatmap_chart, use_container_width=True)

        top_risks = risks_filtered.sort_values(
            ["Severity Score", "Impact Level", "Probability Level"],
            ascending=[False, False, False],
        ).head(5)
        risk_columns[1].dataframe(
            top_risks[
                ["Initiative", "Risk", "Impact Level", "Probability Level", "Status", "Mitigation", "Owner"]
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        risk_columns[0].success("No risks logged for the filtered initiatives.")
        risk_columns[1].empty()

    st.subheader("Initiative detail")
    display_columns = [
        "Initiative",
        "Workstream",
        "Owner",
        "Health Category",
        "Percent Complete",
        "Launch Date",
        "Target Date",
        "Budget",
        "Actual Spend",
        "Status Summary",
    ]
    portfolio_display = filtered[display_columns].rename(columns={"Health Category": "Health", "Percent Complete": "% Complete"})
    st.dataframe(
        portfolio_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "% Complete": st.column_config.NumberColumn("% Complete", format="%.0f"),
            "Launch Date": st.column_config.DateColumn("Launch Date"),
            "Target Date": st.column_config.DateColumn("Target Date"),
            "Budget": st.column_config.NumberColumn("Budget", format="$%,.0f"),
            "Actual Spend": st.column_config.NumberColumn("Actual Spend", format="$%,.0f"),
        },
    )

    export_bytes = to_excel_bytes(filtered, milestones_filtered, risks_filtered)
    st.download_button(
        "Download filtered scorecard (Excel)",
        data=export_bytes,
        file_name=f"program_scorecard_{reporting_date:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def render_workflow_wireframe(data: ScorecardData, reporting_date: datetime.date, uploaded_file) -> None:
    st.subheader("Workflow and dashboard wireframe")
    st.caption(
        "Analyst-facing queue for recurring request maintenance: editable rows, bulk quick actions, draft saves, and workload rollups."
    )

    workflow_source = None
    source_key = data.source_name
    if uploaded_file is not None:
        workflow_source = load_workflow_sheet(uploaded_file)
        source_key = uploaded_file.name

    if st.session_state.get("workflow_source_key") != source_key:
        st.session_state.workflow_drafts = build_workflow_seed(workflow_source, data.portfolio, reporting_date)
        st.session_state.workflow_saved = st.session_state.workflow_drafts.copy()
        st.session_state.workflow_last_save = pd.Timestamp(reporting_date)
        st.session_state.workflow_action_count = 0
        st.session_state.workflow_source_key = source_key
    if "workflow_drafts" not in st.session_state:
        st.session_state.workflow_drafts = build_workflow_seed(workflow_source, data.portfolio, reporting_date)
    if "workflow_saved" not in st.session_state:
        st.session_state.workflow_saved = st.session_state.workflow_drafts.copy()
    if "workflow_last_save" not in st.session_state:
        st.session_state.workflow_last_save = pd.Timestamp(reporting_date)
    if "workflow_action_count" not in st.session_state:
        st.session_state.workflow_action_count = 0

    workflow_df = st.session_state.workflow_drafts.copy()

    top_cols = st.columns([1.3, 1.1, 1.1, 1.1])
    selected_rows = int(workflow_df["Selected"].fillna(False).sum())
    review_rows = int(workflow_df["Needs Review"].fillna(False).sum())
    exception_rows = int(workflow_df["Exception Flag"].fillna("").astype(str).eq("Y").sum())
    top_cols[0].metric("Queue volume", len(workflow_df), delta=f"{review_rows} need review")
    top_cols[1].metric("Selected rows", selected_rows, delta=f"{exception_rows} exceptions")
    top_cols[2].metric("Automated updates", st.session_state.workflow_action_count)
    top_cols[3].metric("Last save", st.session_state.workflow_last_save.strftime("%b %d, %I:%M %p"))

    dashboard_left, dashboard_right = st.columns([1.2, 0.8])
    queue_rollup = workflow_df.groupby(["BuySmart Action", "Status"]).size().reset_index(name="Rows")
    queue_chart = (
        alt.Chart(queue_rollup)
        .mark_bar()
        .encode(
            x=alt.X("Rows:Q", title="Rows"),
            y=alt.Y("BuySmart Action:N", sort="-x"),
            color=alt.Color("Status:N", scale=alt.Scale(range=["#b23a48", "#d88c2d", "#497174", "#6c757d"])),
            tooltip=["BuySmart Action", "Status", "Rows"],
        )
    )
    dashboard_left.altair_chart(queue_chart, use_container_width=True)

    dashboard_right.markdown(
        """
        **Analyst workflow**
        - Analysts work directly in request attributes, decision fields, and notes.
        - Quick actions apply repeatable rule-based updates to the selected cases.
        - Save captures the current draft state before export or downstream processing.
        - Dashboard metrics show review load, exceptions, and action distribution.
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
uploaded_file = st.file_uploader("Upload program scorecard (Excel)", type=["xlsx"])

with st.sidebar:
    st.header("Guidance")
    st.caption(
        "The workbook should include tabs for **Portfolio**, **Milestones**, and **Risks**. "
        "Use the template below for a quick start."
    )
    st.markdown(
        """
        **Workbook schema**
        - `Portfolio`: Initiative, Workstream, Owner, Health/RAG, Percent Complete, Launch Date, Target Date, Budget, Actual Spend, Status Summary
        - `Milestones`: Initiative, Milestone, Target Date, Status, Owner, Notes
        - `Risks`: Initiative, Risk, Impact, Probability, Status, Mitigation, Owner
        """
    )
    st.download_button(
        "Download template workbook",
        data=create_template_workbook(),
        file_name="program_scorecard_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

if uploaded_file:
    data = load_scorecard_from_excel(uploaded_file)
else:
    st.info("Using sample data. Upload your workbook to view live program results.")
    data = load_sample_scorecard()

# Sidebar filters
with st.sidebar:
    st.header("View")
    app_view = st.radio(
        "Experience",
        options=["Executive Scorecard", "Workflow Wireframe"],
        help="Switch between the current reporting experience and the conceptual operations UI.",
    )
    st.header("Filters")
    reporting_date = st.date_input("Reporting as of", datetime.now().date())
    available_workstreams = sorted([w for w in data.portfolio["Workstream"].dropna().unique()])
    selected_workstreams = st.multiselect(
        "Workstreams",
        options=available_workstreams,
        default=available_workstreams,
    )
    available_owners = sorted([o for o in data.portfolio["Owner"].dropna().unique()])
    selected_owners = st.multiselect(
        "Owners",
        options=available_owners,
        default=available_owners,
    )
    selected_health = st.multiselect(
        "Health categories",
        options=STATUS_ORDER,
        default=[s for s in STATUS_ORDER if s not in {"Complete"}],
    )
    include_complete = st.checkbox("Include completed initiatives", value=False)
    lookahead_days = st.slider("Milestone look-ahead (days)", min_value=14, max_value=120, value=45, step=1)

if app_view == "Workflow Wireframe":
    render_workflow_wireframe(data, reporting_date, uploaded_file)
else:
    render_scorecard_view(
        data,
        reporting_date,
        selected_workstreams,
        selected_owners,
        selected_health,
        include_complete,
        lookahead_days,
    )
