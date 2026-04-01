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


def build_workflow_seed(portfolio: pd.DataFrame, milestones: pd.DataFrame, reporting_date: datetime.date) -> pd.DataFrame:
    portfolio_base = portfolio.copy()
    portfolio_base["Percent Complete"] = normalize_percent(portfolio_base["Percent Complete"])
    portfolio_base["Target Date"] = normalize_date(portfolio_base["Target Date"])
    portfolio_base["Launch Date"] = normalize_date(portfolio_base["Launch Date"])
    portfolio_base["Health Category"] = portfolio_base["Health"].map(categorize_health)

    upcoming = (
        milestones.copy()
        .assign(Target_Date=normalize_date(milestones["Target Date"]))
        .sort_values("Target_Date")
        .groupby("Initiative", dropna=False)
        .first()
        .reset_index()
    )
    next_milestones = upcoming[["Initiative", "Milestone", "Target_Date"]].rename(
        columns={"Target_Date": "Next Milestone Date"}
    )

    workflow = portfolio_base.merge(next_milestones, on="Initiative", how="left")
    queue_map = {
        "At Risk": "Exceptions",
        "Watch": "Manager Review",
        "On Hold": "Dependency Hold",
        "Complete": "Archive",
    }
    priority_map = {"At Risk": "P1", "Watch": "P2", "On Hold": "P2", "On Track": "P3", "Not Started": "P3"}

    workflow = workflow.assign(
        Row_ID=[f"WF-{idx:03d}" for idx in range(1, len(workflow) + 1)],
        Queue=lambda df: df["Health Category"].map(queue_map).fillna("Active Workflow"),
        Priority=lambda df: df["Health Category"].map(priority_map).fillna("P3"),
        Workflow_Status=lambda df: df["Health Category"].replace(
            {
                "At Risk": "Needs Update",
                "Watch": "Pending Review",
                "On Hold": "Blocked",
                "Complete": "Completed",
                "On Track": "In Progress",
                "Not Started": "Queued",
            }
        ),
        Automation_Recommendation=lambda df: df["Health Category"].replace(
            {
                "At Risk": "Escalate timeline and request owner update",
                "Watch": "Refresh milestone status",
                "On Hold": "Check dependency owner",
                "Complete": "Archive record",
                "On Track": "No action needed",
                "Not Started": "Assign kickoff owner",
            }
        ),
    )

    workflow["Due Date"] = workflow["Target Date"].fillna(pd.Timestamp(reporting_date))
    workflow["Last Saved"] = pd.Timestamp(reporting_date)
    workflow["Last Sync"] = pd.Timestamp(reporting_date)
    workflow["Selected"] = False
    workflow["Needs Review"] = workflow["Health Category"].isin(["At Risk", "Watch"])

    return workflow[
        [
            "Selected",
            "Row_ID",
            "Initiative",
            "Workstream",
            "Owner",
            "Queue",
            "Workflow_Status",
            "Priority",
            "Health Category",
            "Percent Complete",
            "Due Date",
            "Next Milestone Date",
            "Needs Review",
            "Automation_Recommendation",
            "Status Summary",
            "Last Sync",
            "Last Saved",
        ]
    ].rename(
        columns={
            "Row_ID": "Row ID",
            "Health Category": "Health",
            "Percent Complete": "% Complete",
            "Status Summary": "Operator Notes",
        }
    )


def apply_workflow_quick_action(workflow_df: pd.DataFrame, action: str, reporting_date: datetime.date) -> tuple[pd.DataFrame, int]:
    updated = workflow_df.copy()
    selected_mask = updated["Selected"].fillna(False)
    affected = int(selected_mask.sum())

    if action == "Mark selected for review":
        updated.loc[selected_mask, "Workflow_Status"] = "Pending Review"
        updated.loc[selected_mask, "Needs Review"] = True
        updated.loc[selected_mask, "Automation_Recommendation"] = "Manager review queued"
    elif action == "Apply SLA +3 days":
        updated.loc[selected_mask, "Due Date"] = pd.to_datetime(updated.loc[selected_mask, "Due Date"], errors="coerce") + pd.Timedelta(days=3)
        updated.loc[selected_mask, "Automation_Recommendation"] = "Due date shifted by automation"
    elif action == "Normalize healthy items":
        healthy_mask = selected_mask & updated["Health"].isin(["On Track", "Complete"])
        updated.loc[healthy_mask, "Workflow_Status"] = "Ready to Save"
        updated.loc[healthy_mask, "Needs Review"] = False
        updated.loc[healthy_mask, "Automation_Recommendation"] = "Healthy item normalized"
        affected = int(healthy_mask.sum())
    elif action == "Auto-assign queue owners":
        queue_owner_map = {
            "Exceptions": "PMO Escalations",
            "Manager Review": "Program Manager",
            "Dependency Hold": "Dependency Lead",
            "Archive": "Operations Analyst",
            "Active Workflow": "Workflow Coordinator",
        }
        updated.loc[selected_mask, "Owner"] = updated.loc[selected_mask, "Queue"].map(queue_owner_map).fillna(updated.loc[selected_mask, "Owner"])
        updated.loc[selected_mask, "Automation_Recommendation"] = "Owner auto-assigned from queue rule"

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


def render_workflow_wireframe(data: ScorecardData, reporting_date: datetime.date) -> None:
    st.subheader("Workflow and dashboard wireframe")
    st.caption(
        "Concept view for an operations UI: editable queue, bulk quick actions, session-backed draft saves, and dashboard rollups."
    )

    if "workflow_drafts" not in st.session_state:
        st.session_state.workflow_drafts = build_workflow_seed(data.portfolio, data.milestones, reporting_date)
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
    ready_rows = int((workflow_df["Workflow_Status"] == "Ready to Save").sum())
    top_cols[0].metric("Queue volume", len(workflow_df), delta=f"{review_rows} need review")
    top_cols[1].metric("Selected rows", selected_rows, delta=f"{ready_rows} ready to save")
    top_cols[2].metric("Automated updates", st.session_state.workflow_action_count)
    top_cols[3].metric("Last save", st.session_state.workflow_last_save.strftime("%b %d, %I:%M %p"))

    dashboard_left, dashboard_right = st.columns([1.2, 0.8])
    queue_rollup = workflow_df.groupby(["Queue", "Priority"]).size().reset_index(name="Rows")
    queue_chart = (
        alt.Chart(queue_rollup)
        .mark_bar()
        .encode(
            x=alt.X("Rows:Q", title="Rows"),
            y=alt.Y("Queue:N", sort="-x"),
            color=alt.Color("Priority:N", scale=alt.Scale(range=["#b23a48", "#d88c2d", "#497174"])),
            tooltip=["Queue", "Priority", "Rows"],
        )
    )
    dashboard_left.altair_chart(queue_chart, use_container_width=True)

    dashboard_right.markdown(
        """
        **Quick action model**
        - Analysts make inline edits directly in the workflow grid.
        - Automation applies queue-level updates to the selected rows.
        - Save commits the current draft state for downstream processing.
        - Dashboard cards show what changed and what still needs attention.
        """
    )

    control_cols = st.columns([1.1, 1.1, 0.8, 0.8])
    action = control_cols[0].selectbox(
        "Quick action",
        options=[
            "Mark selected for review",
            "Apply SLA +3 days",
            "Normalize healthy items",
            "Auto-assign queue owners",
        ],
    )
    control_cols[1].caption("Use the checkbox column to target rows before running an automation.")
    apply_clicked = control_cols[2].button("Run quick action", use_container_width=True)
    save_clicked = control_cols[3].button("Save changes", type="primary", use_container_width=True)

    edited_workflow = st.data_editor(
        workflow_df,
        use_container_width=True,
        hide_index=True,
        height=460,
        column_config={
            "Selected": st.column_config.CheckboxColumn("Select"),
            "Row ID": st.column_config.TextColumn("Row ID", disabled=True),
            "Initiative": st.column_config.TextColumn("Initiative", disabled=True),
            "Workstream": st.column_config.TextColumn("Workstream", disabled=True),
            "% Complete": st.column_config.NumberColumn("% Complete", min_value=0, max_value=100, format="%d"),
            "Due Date": st.column_config.DateColumn("Due Date"),
            "Next Milestone Date": st.column_config.DateColumn("Next Milestone"),
            "Needs Review": st.column_config.CheckboxColumn("Needs Review"),
            "Last Sync": st.column_config.DateColumn("Last Sync", disabled=True),
            "Last Saved": st.column_config.DateColumn("Last Saved", disabled=True),
        },
        disabled=["Row ID", "Initiative", "Workstream", "Last Sync", "Last Saved"],
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
    render_workflow_wireframe(data, reporting_date)
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
