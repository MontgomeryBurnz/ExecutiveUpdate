from __future__ import annotations

from datetime import date, datetime, timedelta
import re
from textwrap import dedent
from urllib.parse import quote
from zoneinfo import ZoneInfo

import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Compliance Hub - Executive Tracker", layout="wide")

COLORS = {
    "navy": "#2a4561",
    "blue": "#2f6fed",
    "blue_soft": "#7ea6f8",
    "teal": "#4d7fb3",
    "panel": "#ffffff",
    "border": "#d9e3f0",
    "text": "#506377",
    "muted": "#7f91a6",
    "soft": "#f6f8fc",
    "green": "#2e8b6f",
    "yellow": "#d3a23c",
    "red": "#cf5c5c",
}

PORTFOLIO_NAME = "FY25 Strategic Transformation Portfolio"
APP_TIMEZONE = ZoneInfo("America/New_York")
PAGES = [
    "Impower Portfolio",
    "Program One-Pager",
    "Weekly Updates",
    "Settings",
    "Help & Support",
]

PORTFOLIO_SEEDS: dict[str, list[dict[str, object]]] = {
    "Core Delivery": [
        {
            "Program": "Data Integration",
            "Executive Sponsor": "R. Thompson",
            "Lead": "A. Miller",
            "Stage": "Discovery",
            "Status": "On Track",
            "Priority": "High",
            "Start": "2026-01-08",
            "End": "2026-02-21",
            "Progress": 78,
            "Milestone": "Scope baseline",
            "Milestone Date": "2026-02-14",
            "Delivery Health": "Green",
            "Tech Health": "Green",
            "Team Health": "Green",
            "Open Risks": 1,
            "Escalations": 0,
            "Summary": "Status unchanged from prior cycle. Discovery is landing cleanly with low escalation pressure.",
            "Upcoming Work": "Finalize scope, confirm integration dependencies, close architecture review comments.",
            "Risk Detail": "Minor data-mapping dependency remains open with a contingency path available.",
            "Decision Needed": "No executive decision required this cycle.",
            "Mitigation": "Weekly dependency check with architecture and vendor teams.",
            "Status Note": "Discovery remains stable and aligned to baseline.",
            "Next Step": "Finalize solution scope and confirm dependencies.",
        },
        {
            "Program": "Mobile App Delivery",
            "Executive Sponsor": "J. Carter",
            "Lead": "N. Patel",
            "Stage": "Phase 1",
            "Status": "Needs Attention",
            "Priority": "High",
            "Start": "2026-02-01",
            "End": "2026-04-18",
            "Progress": 52,
            "Milestone": "MVP release",
            "Milestone Date": "2026-04-30",
            "Delivery Health": "Yellow",
            "Tech Health": "Green",
            "Team Health": "Yellow",
            "Open Risks": 3,
            "Escalations": 1,
            "Summary": "Velocity is softer than plan but recoverable this month with focused delivery discipline.",
            "Upcoming Work": "Tighten sprint scope, complete UAT readiness plan, close design debt carryover.",
            "Risk Detail": "Integration dependency is increasing risk exposure and compressing test time.",
            "Decision Needed": "Confirm temporary UAT support allocation for the next two weeks.",
            "Mitigation": "Daily burn review and escalation path with product and QA leads.",
            "Status Note": "Needs attention due to throughput variance.",
            "Next Step": "Tighten sprint throughput and confirm UAT staffing.",
        },
        {
            "Program": "CRM Modernization",
            "Executive Sponsor": "M. Woods",
            "Lead": "J. Davis",
            "Stage": "Phase 2",
            "Status": "At Risk",
            "Priority": "Critical",
            "Start": "2026-03-01",
            "End": "2026-06-28",
            "Progress": 41,
            "Milestone": "Integration cutoff",
            "Milestone Date": "2026-05-17",
            "Delivery Health": "Red",
            "Tech Health": "Yellow",
            "Team Health": "Yellow",
            "Open Risks": 5,
            "Escalations": 2,
            "Summary": "Dependency delay has shifted confidence on the phase gate and requires active sponsorship.",
            "Upcoming Work": "Resolve vendor sequencing, confirm cutover path, re-baseline timeline assumptions.",
            "Risk Detail": "Vendor dependency has increased the probability of a milestone miss.",
            "Decision Needed": "Approve escalation with the vendor sponsor if cutoff confidence does not improve this week.",
            "Mitigation": "Parallel workstream for contingency design and twice-weekly leadership review.",
            "Status Note": "At risk because of external dependency and milestone pressure.",
            "Next Step": "Resolve vendor dependency and lock cutover path.",
        },
        {
            "Program": "Supply Chain Platform",
            "Executive Sponsor": "P. Stone",
            "Lead": "S. Nguyen",
            "Stage": "Phase 3",
            "Status": "On Track",
            "Priority": "Medium",
            "Start": "2026-05-01",
            "End": "2026-08-09",
            "Progress": 63,
            "Milestone": "Regional rollout",
            "Milestone Date": "2026-07-12",
            "Delivery Health": "Green",
            "Tech Health": "Green",
            "Team Health": "Green",
            "Open Risks": 2,
            "Escalations": 0,
            "Summary": "Delivery remains stable with good cross-functional support across the rollout region.",
            "Upcoming Work": "Confirm regional onboarding readiness and finalize launch communications.",
            "Risk Detail": "Moderate adoption risk remains but is contained through the rollout plan.",
            "Decision Needed": "No decision required; monitor rollout readiness cadence.",
            "Mitigation": "Regional readiness reviews and launch rehearsal checkpoints.",
            "Status Note": "Stable execution with manageable rollout risk.",
            "Next Step": "Confirm regional onboarding readiness.",
        },
    ],
    "Platform Modernization": [
        {
            "Program": "Cloud Foundation",
            "Executive Sponsor": "L. Hart",
            "Lead": "R. Flores",
            "Stage": "Discovery",
            "Status": "On Track",
            "Priority": "High",
            "Start": "2026-01-15",
            "End": "2026-03-15",
            "Progress": 70,
            "Milestone": "Architecture sign-off",
            "Milestone Date": "2026-03-05",
            "Delivery Health": "Green",
            "Tech Health": "Green",
            "Team Health": "Green",
            "Open Risks": 1,
            "Escalations": 0,
            "Summary": "Architecture decisions are converging on plan with no material leadership concerns.",
            "Upcoming Work": "Close security comments and finalize landing zone standards.",
            "Risk Detail": "Security review comments may slow sign-off if not resolved promptly.",
            "Decision Needed": "No current decision required.",
            "Mitigation": "Security architect embedded in the weekly review cadence.",
            "Status Note": "Stable and progressing toward architecture approval.",
            "Next Step": "Close remaining security review comments.",
        },
        {
            "Program": "Data Platform Refresh",
            "Executive Sponsor": "V. Shah",
            "Lead": "K. Young",
            "Stage": "Phase 2",
            "Status": "Needs Attention",
            "Priority": "Critical",
            "Start": "2026-02-22",
            "End": "2026-06-14",
            "Progress": 47,
            "Milestone": "Migration wave 1",
            "Milestone Date": "2026-05-28",
            "Delivery Health": "Yellow",
            "Tech Health": "Yellow",
            "Team Health": "Green",
            "Open Risks": 4,
            "Escalations": 1,
            "Summary": "Migration planning needs firmer control on sequencing to preserve milestone confidence.",
            "Upcoming Work": "Sequence migration windows and reduce backlog carryover.",
            "Risk Detail": "Wave sequencing remains the top portfolio risk for this program.",
            "Decision Needed": "Confirm tradeoff between migration scope and cutover stability.",
            "Mitigation": "Weekly integrated planning review and readiness criteria reset.",
            "Status Note": "Needs attention because of schedule and dependency complexity.",
            "Next Step": "Reduce backlog carryover and sequence migration windows.",
        },
    ],
    "Shared Services": [
        {
            "Program": "HR Service Refresh",
            "Executive Sponsor": "E. Cole",
            "Lead": "L. Brooks",
            "Stage": "Phase 1",
            "Status": "On Track",
            "Priority": "Medium",
            "Start": "2026-01-20",
            "End": "2026-04-11",
            "Progress": 66,
            "Milestone": "Pilot release",
            "Milestone Date": "2026-04-22",
            "Delivery Health": "Green",
            "Tech Health": "Green",
            "Team Health": "Yellow",
            "Open Risks": 2,
            "Escalations": 0,
            "Summary": "Program is healthy with minor capacity pressure but no material delivery drift.",
            "Upcoming Work": "Close pilot feedback and finalize rollout communications.",
            "Risk Detail": "Capacity strain could reduce responsiveness on late-stage changes.",
            "Decision Needed": "No executive decision required.",
            "Mitigation": "Capacity is being balanced through temporary PMO support.",
            "Status Note": "On track with manageable people pressure.",
            "Next Step": "Close pilot feedback and ready final rollout plan.",
        },
        {
            "Program": "Procurement Analytics",
            "Executive Sponsor": "C. Vega",
            "Lead": "T. Hall",
            "Stage": "Phase 3",
            "Status": "At Risk",
            "Priority": "Critical",
            "Start": "2026-03-10",
            "End": "2026-07-26",
            "Progress": 38,
            "Milestone": "Analytics adoption gate",
            "Milestone Date": "2026-06-18",
            "Delivery Health": "Red",
            "Tech Health": "Yellow",
            "Team Health": "Red",
            "Open Risks": 6,
            "Escalations": 2,
            "Summary": "Executive intervention may be required on capacity and scope to protect the adoption gate.",
            "Upcoming Work": "Stabilize staffing plan and unblock reporting requirements.",
            "Risk Detail": "Adoption and staffing risks are both trending negatively.",
            "Decision Needed": "Approve backfill support or descope lower-value analytics features.",
            "Mitigation": "Weekly adoption pulse and sponsor-led staffing review.",
            "Status Note": "At risk because of staffing and adoption pressures.",
            "Next Step": "Stabilize staffing plan and unblock reporting requirements.",
        },
    ],
}

PROGRAM_TO_PORTFOLIO = {
    row["Program"]: portfolio
    for portfolio, rows in PORTFOLIO_SEEDS.items()
    for row in rows
}
ALL_PROGRAMS = list(PROGRAM_TO_PORTFOLIO.keys())


def inject_styles() -> None:
    st.markdown(
        f"""
        <style>
            .stApp {{
                background:
                    radial-gradient(circle at top left, rgba(47,111,237,0.10), transparent 22%),
                    linear-gradient(180deg, #f8fafd 0%, #edf2f8 100%);
                color: {COLORS["text"]};
            }}
            .block-container {{
                max-width: 1520px;
                padding-top: 1rem;
                padding-bottom: 2rem;
            }}
            section[data-testid="stSidebar"] {{
                background: linear-gradient(180deg, #e9eff8 0%, #dce7f5 100%);
                border-right: 1px solid #c9d7ea;
            }}
            section[data-testid="stSidebar"] * {{
                color: {COLORS["text"]};
            }}
            section[data-testid="stSidebar"] [data-baseweb="select"] > div,
            section[data-testid="stSidebar"] [data-baseweb="input"] > div {{
                background: rgba(255,255,255,0.9);
                border-color: #c7d6e8;
            }}
            section[data-testid="stSidebar"] .stRadio > label,
            section[data-testid="stSidebar"] .stSelectbox > label,
            section[data-testid="stSidebar"] .stDateInput > label {{
                font-size: 0.78rem;
                letter-spacing: 0.12em;
                text-transform: uppercase;
                font-weight: 800;
                color: #6a7c93;
            }}
            section[data-testid="stSidebar"] [role="radiogroup"] label {{
                background: rgba(255,255,255,0.78);
                border: 1px solid #c7d6e8;
                border-radius: 14px;
                padding: 0.55rem 0.75rem;
                margin-bottom: 0.45rem;
            }}
            section[data-testid="stSidebar"] [role="radiogroup"] label[data-checked="true"] {{
                background: #ffffff;
                border-color: #97b4e4;
                box-shadow: 0 8px 20px rgba(47,111,237,0.10);
            }}
            .stTextArea label,
            .stTextInput label,
            .stDateInput label,
            .stSelectbox label,
            .stSlider label,
            .stRadio > label,
            .stMultiSelect label {{
                font-size: 0.78rem;
                letter-spacing: 0.12em;
                text-transform: uppercase;
                font-weight: 800;
                color: #6a7c93;
            }}
            .stTextArea textarea,
            .stTextInput input,
            .stDateInput input,
            .stNumberInput input {{
                background: linear-gradient(180deg, #ffffff 0%, #f5f9ff 100%);
                border: 1px solid #c9d8ea;
                border-radius: 14px;
                color: #23384f;
            }}
            .stTextArea textarea:focus,
            .stTextInput input:focus,
            .stDateInput input:focus,
            .stNumberInput input:focus {{
                border-color: #7ea6f8;
                box-shadow: 0 0 0 1px #7ea6f8;
            }}
            [data-baseweb="select"] > div,
            [data-baseweb="input"] > div,
            .stDateInput [data-baseweb="input"] > div,
            .stMultiSelect [data-baseweb="select"] > div {{
                background: linear-gradient(180deg, #ffffff 0%, #f5f9ff 100%);
                border-color: #c9d8ea;
                border-radius: 14px;
                color: #23384f;
            }}
            .stRadio [role="radiogroup"] {{
                gap: 0.45rem;
            }}
            .stRadio [role="radiogroup"] label {{
                background: rgba(255,255,255,0.9);
                border: 1px solid #c9d8ea;
                border-radius: 14px;
                padding: 0.55rem 0.8rem;
            }}
            .stRadio [role="radiogroup"] label[data-checked="true"] {{
                background: linear-gradient(180deg, #eef4ff 0%, #deebff 100%);
                border-color: #8eb0f1;
                box-shadow: 0 8px 20px rgba(47,111,237,0.10);
            }}
            .stSlider [data-baseweb="slider"] {{
                padding-top: 0.6rem;
            }}
            .stSlider [role="slider"] {{
                background: #376fe7;
                box-shadow: 0 0 0 4px rgba(47,111,237,0.16);
            }}
            .stSlider [data-baseweb="slider"] > div > div {{
                background: #d7e5fb;
            }}
            .hero {{
                background: linear-gradient(90deg, #f7faff 0%, #f7faff 74%, #dbe7fb 74%, #dbe7fb 100%);
                color: {COLORS["text"]};
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 1rem 1.15rem;
                border: 1px solid #d2ddec;
                margin-bottom: 1rem;
            }}
            .hero-title {{
                font-size: 1.95rem;
                font-weight: 800;
                letter-spacing: -0.03em;
                color: {COLORS["navy"]};
            }}
            .hero-badge {{
                min-width: 260px;
                text-align: center;
                font-size: 1.25rem;
                font-weight: 800;
                color: {COLORS["blue"]};
            }}
            .subtitle {{
                margin-bottom: 1rem;
                color: {COLORS["muted"]};
            }}
            .topbar {{
                display: grid;
                grid-template-columns: 240px minmax(280px, 1fr) auto auto;
                gap: 0.85rem;
                align-items: center;
                margin-bottom: 1rem;
            }}
            .brand-block {{
                display: flex;
                flex-direction: column;
                gap: 0.15rem;
            }}
            .brand-title {{
                font-size: 1.2rem;
                font-weight: 800;
                color: {COLORS["navy"]};
            }}
            .brand-copy {{
                font-size: 0.82rem;
                color: {COLORS["muted"]};
            }}
            .search-shell {{
                background: rgba(255,255,255,0.94);
                border: 1px solid {COLORS["border"]};
                border-radius: 12px;
                padding: 0.72rem 0.9rem;
                color: #94a3b8;
                font-size: 0.9rem;
            }}
            .toolbar-actions {{
                display: flex;
                gap: 0.5rem;
                align-items: center;
            }}
            .toolbar-icon {{
                width: 34px;
                height: 34px;
                border-radius: 10px;
                border: 1px solid {COLORS["border"]};
                background: rgba(255,255,255,0.95);
                display: flex;
                align-items: center;
                justify-content: center;
                color: {COLORS["muted"]};
                font-size: 0.9rem;
                font-weight: 700;
            }}
            .profile-shell {{
                display: flex;
                gap: 0.75rem;
                align-items: center;
                justify-content: flex-end;
            }}
            .profile-copy {{
                text-align: right;
                line-height: 1.2;
            }}
            .profile-name {{
                font-size: 0.92rem;
                font-weight: 800;
                color: {COLORS["navy"]};
            }}
            .profile-role {{
                font-size: 0.8rem;
                color: {COLORS["muted"]};
            }}
            .profile-avatar {{
                width: 36px;
                height: 36px;
                border-radius: 999px;
                background: linear-gradient(180deg, {COLORS["blue"]}, #2359c6);
                color: white;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 0.82rem;
                font-weight: 800;
            }}
            .dashboard-title-block {{
                margin-bottom: 1rem;
            }}
            .dashboard-title {{
                font-size: 2rem;
                font-weight: 800;
                color: {COLORS["navy"]};
                letter-spacing: -0.03em;
            }}
            .dashboard-meta {{
                display: flex;
                gap: 0.7rem;
                align-items: center;
                margin-top: 0.35rem;
                color: {COLORS["muted"]};
                font-size: 0.88rem;
            }}
            .update-pill {{
                padding: 0.28rem 0.55rem;
                border-radius: 999px;
                background: rgba(46,139,111,0.12);
                color: {COLORS["green"]};
                font-weight: 700;
            }}
            .metric-strip {{
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 0.9rem;
                margin-bottom: 1rem;
            }}
            .metric-card, .card, .editor-card {{
                background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
                border: 1px solid {COLORS["border"]};
                border-top: 4px solid {COLORS["blue"]};
                border-radius: 20px;
                padding: 1rem;
                box-shadow: 0 14px 34px rgba(22,50,79,0.05);
            }}
            .dashboard-shell {{
                background: rgba(255,255,255,0.82);
                border: 1px solid {COLORS["border"]};
                border-radius: 24px;
                padding: 1rem;
                box-shadow: 0 18px 42px rgba(22,50,79,0.05);
            }}
            .dash-kpi-grid {{
                display: grid;
                grid-template-columns: repeat(6, minmax(0, 1fr));
                gap: 0.95rem;
                margin-bottom: 1rem;
            }}
            .dashboard-page-grid {{
                display: grid;
                grid-template-columns: 220px minmax(0, 1fr);
                gap: 1rem;
                align-items: start;
            }}
            .left-rail {{
                background: linear-gradient(180deg, #f7f9fd 0%, #eef3fa 100%);
                border: 1px solid {COLORS["border"]};
                border-radius: 22px;
                padding: 0.9rem 0.7rem;
                box-shadow: 0 14px 34px rgba(22,50,79,0.04);
            }}
            .rail-item {{
                display: flex;
                align-items: center;
                gap: 0.65rem;
                padding: 0.7rem 0.75rem;
                border-radius: 14px;
                color: {COLORS["muted"]};
                font-size: 0.9rem;
                font-weight: 600;
                margin-bottom: 0.35rem;
            }}
            .rail-item.active {{
                background: linear-gradient(180deg, #3a73eb 0%, #275fd5 100%);
                color: white;
                box-shadow: 0 10px 24px rgba(47,111,237,0.22);
            }}
            .rail-icon {{
                width: 18px;
                text-align: center;
                opacity: 0.8;
            }}
            .rail-divider {{
                height: 1px;
                background: #dde7f2;
                margin: 0.7rem 0;
            }}
            .dash-kpi {{
                background: linear-gradient(180deg, #ffffff 0%, #f5f8fd 100%);
                border: 1px solid {COLORS["border"]};
                border-radius: 18px;
                padding: 1rem 1rem 0.95rem;
                min-height: 126px;
            }}
            .dash-kpi-name {{
                font-size: 0.74rem;
                text-transform: uppercase;
                letter-spacing: 0.14em;
                font-weight: 800;
                color: {COLORS["muted"]};
            }}
            .dash-kpi-value {{
                margin-top: 0.45rem;
                font-size: 2rem;
                font-weight: 800;
                color: {COLORS["navy"]};
            }}
            .dash-kpi-copy {{
                margin-top: 0.3rem;
                color: {COLORS["muted"]};
                font-size: 0.9rem;
                line-height: 1.35;
            }}
            .dashboard-main-grid {{
                display: grid;
                grid-template-columns: minmax(0, 1.7fr) minmax(320px, 0.96fr);
                gap: 1rem;
                align-items: start;
            }}
            .dashboard-stack {{
                display: grid;
                gap: 1rem;
            }}
            .dash-card {{
                background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
                border: 1px solid {COLORS["border"]};
                border-radius: 20px;
                padding: 1rem;
                box-shadow: 0 14px 34px rgba(22,50,79,0.05);
            }}
            .dash-card-title {{
                font-size: 0.8rem;
                text-transform: uppercase;
                letter-spacing: 0.14em;
                font-weight: 800;
                color: {COLORS["blue"]};
                margin-bottom: 0.28rem;
            }}
            .dash-card-heading {{
                font-size: 1.18rem;
                font-weight: 800;
                color: {COLORS["navy"]};
                margin-bottom: 0.2rem;
            }}
            .dash-card-copy {{
                color: {COLORS["muted"]};
                margin-bottom: 0.8rem;
                line-height: 1.4;
            }}
            .dash-table-wrap {{
                border: 1px solid #e4ebf4;
                border-radius: 16px;
                overflow: hidden;
                background: white;
            }}
            .dash-table {{
                width: 100%;
                border-collapse: collapse;
                font-size: 0.94rem;
            }}
            .dash-table th {{
                text-align: left;
                padding: 0.75rem 0.8rem;
                font-size: 0.72rem;
                text-transform: uppercase;
                letter-spacing: 0.12em;
                color: {COLORS["muted"]};
                background: #eff4fb;
                border-bottom: 1px solid #dfe7f1;
            }}
            .dash-table td {{
                padding: 0.82rem 0.8rem;
                border-bottom: 1px solid #e8eef6;
                vertical-align: top;
                color: {COLORS["text"]};
                line-height: 1.38;
            }}
            .dash-table tr:last-child td {{
                border-bottom: none;
            }}
            .dash-program {{
                font-weight: 700;
                color: {COLORS["navy"]};
            }}
            .dash-note {{
                color: {COLORS["muted"]};
            }}
            .dash-tag {{
                display: inline-block;
                border-radius: 999px;
                padding: 0.24rem 0.62rem;
                font-size: 0.76rem;
                font-weight: 800;
                white-space: nowrap;
            }}
            .dash-tag-green {{
                color: {COLORS["green"]};
                background: rgba(46,139,111,0.12);
            }}
            .dash-tag-yellow {{
                color: {COLORS["yellow"]};
                background: rgba(211,162,60,0.14);
            }}
            .dash-tag-red {{
                color: {COLORS["red"]};
                background: rgba(207,92,92,0.13);
            }}
            .dash-roadmap-note {{
                margin-top: 0.8rem;
                padding-top: 0.8rem;
                border-top: 1px solid #e7edf5;
                color: {COLORS["muted"]};
                font-size: 0.9rem;
                line-height: 1.42;
            }}
            .roadmap-grid {{
                display: grid;
                gap: 1.5rem;
                min-width: 0;
            }}
            .roadmap-quarter-row {{
                display: flex;
                gap: 1.8rem;
                flex-wrap: wrap;
                margin: 0.6rem 0 1.2rem;
                font-size: 1.1rem;
                color: #a3a7ad;
            }}
            .roadmap-quarter-row .active {{
                color: #294bb5;
                font-weight: 800;
            }}
            .roadmap-journey {{
                display: grid;
                gap: 0.4rem;
            }}
            .roadmap-journey-name {{
                font-size: 1.05rem;
                font-weight: 800;
                color: #3f4a58;
            }}
            .road-band-wrap {{
                position: relative;
                padding-bottom: 2rem;
                width: 100%;
                min-width: 0;
            }}
            .road-band {{
                display: flex;
                align-items: stretch;
                width: 100%;
                min-height: 3.15rem;
                border-radius: 999px;
                background: #eff1f4;
                overflow: hidden;
                min-width: 0;
            }}
            .road-band-segment {{
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 0.98rem;
                font-weight: 700;
                line-height: 1;
                min-width: 0;
                white-space: nowrap;
            }}
            .road-band-segment:first-child {{
                border-top-left-radius: 999px;
                border-bottom-left-radius: 999px;
            }}
            .road-band-segment:last-child {{
                border-top-right-radius: 999px;
                border-bottom-right-radius: 999px;
            }}
            .road-marker {{
                position: absolute;
                top: -0.35rem;
                width: 1rem;
                height: 4rem;
                transform: translateX(-50%);
                border-radius: 2px;
            }}
            .road-marker-ok {{ background: #294bb5; }}
            .road-marker-warn {{ background: #ff9800; }}
            .road-marker-alert {{ background: #f44336; }}
            .road-marker-label {{
                position: absolute;
                top: 3.7rem;
                transform: translateX(-50%);
                font-size: 0.82rem;
                font-weight: 800;
                white-space: nowrap;
            }}
            .road-marker-label.marker-ok {{ color: #294bb5; }}
            .road-marker-label.marker-warn {{ color: #d97800; }}
            .road-marker-label.marker-alert {{ color: #df2b33; }}
            .roadmap-row {{
                display: grid;
                grid-template-columns: 190px repeat(5, minmax(0, 1fr)) 70px;
                gap: 0.45rem;
                align-items: center;
            }}
            .roadmap-name {{
                font-size: 0.88rem;
                font-weight: 700;
                color: {COLORS["navy"]};
            }}
            .roadmap-pill {{
                border-radius: 999px;
                padding: 0.3rem 0.55rem;
                font-size: 0.74rem;
                font-weight: 700;
                text-align: center;
                background: #edf2f8;
                color: #98a5b5;
            }}
            .roadmap-pill.discovery {{ background: #d8e3fb; color: #496fae; }}
            .roadmap-pill.phase1 {{ background: #f7d7a8; color: #9c6918; }}
            .roadmap-pill.phase2 {{ background: #d7c6f7; color: #7454b5; }}
            .roadmap-pill.phase3 {{ background: #cfe6ff; color: #4c7cb8; }}
            .roadmap-pill.phase4 {{ background: #d6f1de; color: #3a8d62; }}
            .roadmap-progress {{
                font-size: 0.76rem;
                font-weight: 800;
                color: {COLORS["muted"]};
                text-align: right;
            }}
            .road-legend {{
                display: flex;
                flex-wrap: wrap;
                gap: 1rem;
                margin-top: 1.1rem;
                color: #737d8a;
                font-size: 0.92rem;
            }}
            .road-legend-item {{
                display: inline-flex;
                align-items: center;
                gap: 0.45rem;
            }}
            .road-legend-swatch {{
                width: 1.05rem;
                height: 1.05rem;
                border-radius: 0.4rem;
                display: inline-block;
            }}
            .milestone-list {{
                display: grid;
                gap: 0.75rem;
            }}
            .milestone-item, .risk-item {{
                background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
                border: 1px solid #e5ecf5;
                border-radius: 16px;
                padding: 0.85rem 0.9rem;
            }}
            .milestone-item {{
                display: grid;
                grid-template-columns: 48px 1fr auto;
                gap: 0.8rem;
                align-items: center;
            }}
            .milestone-datebox {{
                width: 48px;
                height: 48px;
                border-radius: 12px;
                background: #f1f5fb;
                border: 1px solid #dce5f1;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
            }}
            .milestone-day {{
                font-size: 1rem;
                font-weight: 800;
                color: {COLORS["navy"]};
                line-height: 1;
            }}
            .milestone-month {{
                font-size: 0.64rem;
                font-weight: 800;
                color: {COLORS["muted"]};
                text-transform: uppercase;
                letter-spacing: 0.08em;
                margin-top: 0.18rem;
            }}
            .milestone-title {{
                font-size: 0.93rem;
                font-weight: 700;
                color: {COLORS["navy"]};
            }}
            .milestone-sub {{
                font-size: 0.82rem;
                color: {COLORS["muted"]};
                margin-top: 0.15rem;
            }}
            .priority-badge {{
                border-radius: 999px;
                padding: 0.24rem 0.56rem;
                font-size: 0.74rem;
                font-weight: 800;
                white-space: nowrap;
            }}
            .priority-high {{
                background: rgba(46,139,111,0.12);
                color: {COLORS["green"]};
            }}
            .priority-medium {{
                background: rgba(211,162,60,0.14);
                color: {COLORS["yellow"]};
            }}
            .priority-low {{
                background: rgba(207,92,92,0.13);
                color: {COLORS["red"]};
            }}
            .risk-item {{
                display: grid;
                gap: 0.55rem;
                padding: 1rem 1.1rem;
            }}
            .risk-head {{
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                gap: 0.75rem;
            }}
            .risk-title {{
                font-size: 1rem;
                font-weight: 800;
                color: #212833;
            }}
            .risk-meta {{
                font-size: 0.84rem;
                color: #7f8795;
            }}
            .risk-trend {{
                font-size: 0.9rem;
                font-weight: 800;
            }}
            .risk-detail-line {{
                display: flex;
                gap: 0.8rem;
                flex-wrap: wrap;
                align-items: center;
            }}
            .risk-pill {{
                font-weight: 700;
            }}
            .risk-item.risk-worse {{
                background: linear-gradient(180deg, #fff7f7 0%, #fff1f1 100%);
                border-color: #ffbcbc;
            }}
            .risk-item.risk-better {{
                background: linear-gradient(180deg, #f7fff9 0%, #f2fcf5 100%);
                border-color: #b9ebc8;
            }}
            .risk-item.risk-stable {{
                background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
                border-color: #d9e3f0;
            }}
            .trend-worse {{ color: #ef3d39; }}
            .trend-stable {{ color: #8b95a3; }}
            .trend-better {{ color: #1cab55; }}
            .dash-bottom-space {{
                margin-top: 1rem;
            }}
            .metric-title, .eyebrow {{
                font-size: 0.76rem;
                text-transform: uppercase;
                letter-spacing: 0.14em;
                font-weight: 800;
                color: {COLORS["muted"]};
            }}
            .metric-number {{
                margin-top: 0.4rem;
                font-size: 2rem;
                font-weight: 800;
                color: {COLORS["navy"]};
            }}
            .metric-note, .copy {{
                margin-top: 0.2rem;
                color: {COLORS["muted"]};
            }}
            .section-bar {{
                background: linear-gradient(180deg, #6d96ea 0%, #4f7fe6 100%);
                color: white;
                font-weight: 800;
                text-align: center;
                border-radius: 14px;
                padding: 0.8rem 0.9rem;
                margin-bottom: 0.9rem;
            }}
            .sub-bar {{
                background: linear-gradient(180deg, #4f8ced 0%, #3476df 100%);
                color: white;
                font-weight: 800;
                text-align: center;
                border-radius: 12px;
                padding: 0.7rem 0.8rem;
                margin-bottom: 0.7rem;
            }}
            .heading {{
                font-size: 1.35rem;
                font-weight: 800;
                color: {COLORS["navy"]};
                margin: 0.12rem 0 0.2rem;
            }}
            .panel-header {{
                background: linear-gradient(180deg, #f8fbff 0%, #edf3fb 100%);
                border: 1px solid {COLORS["border"]};
                border-top: 4px solid {COLORS["blue"]};
                border-radius: 14px;
                padding: 0.85rem 1rem;
                margin-bottom: 0.9rem;
            }}
            .panel-header .heading {{
                margin: 0.1rem 0 0;
            }}
            div[data-testid="stVerticalBlockBorderWrapper"] {{
                border: 1px solid {COLORS["border"]} !important;
                border-top: 4px solid {COLORS["blue"]} !important;
                border-radius: 18px !important;
                background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
                box-shadow: 0 14px 34px rgba(22,50,79,0.05);
            }}
            .stDataFrame {{
                border-radius: 18px;
                overflow: hidden;
            }}
            .portfolio-table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 0.8rem;
                font-size: 0.95rem;
            }}
            .portfolio-table th {{
                text-align: left;
                font-size: 0.74rem;
                text-transform: uppercase;
                letter-spacing: 0.12em;
                color: {COLORS["muted"]};
                padding: 0.75rem 0.7rem;
                border-bottom: 1px solid {COLORS["border"]};
                background: rgba(47,111,237,0.04);
            }}
            .portfolio-table td {{
                padding: 0.85rem 0.7rem;
                border-bottom: 1px solid #e8eef6;
                vertical-align: top;
                color: {COLORS["text"]};
            }}
            .portfolio-table tr:last-child td {{
                border-bottom: none;
            }}
            .table-note {{
                font-size: 0.86rem;
                color: {COLORS["muted"]};
                line-height: 1.4;
            }}
            .mini-kpi-strip {{
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
                gap: 0.75rem;
                margin-bottom: 0.95rem;
            }}
            .mini-kpi {{
                background: linear-gradient(180deg, #ffffff 0%, #f5f9ff 100%);
                border: 1px solid {COLORS["border"]};
                border-radius: 16px;
                padding: 0.85rem 0.9rem;
            }}
            .mini-kpi-label {{
                font-size: 0.7rem;
                text-transform: uppercase;
                letter-spacing: 0.12em;
                font-weight: 800;
                color: {COLORS["muted"]};
            }}
            .mini-kpi-value {{
                margin-top: 0.3rem;
                font-size: 1.35rem;
                font-weight: 800;
                color: {COLORS["navy"]};
            }}
            .mini-list {{
                margin: 0;
                padding-left: 1.05rem;
            }}
            .mini-list li {{
                margin-bottom: 0.28rem;
            }}
            .status-pill {{
                display: inline-block;
                padding: 0.3rem 0.65rem;
                border-radius: 999px;
                font-size: 0.78rem;
                font-weight: 800;
                margin-top: 0.4rem;
            }}
            .status-green {{
                background: rgba(46,139,111,0.12);
                color: {COLORS["green"]};
            }}
            .status-yellow {{
                background: rgba(211,162,60,0.14);
                color: {COLORS["yellow"]};
            }}
            .status-red {{
                background: rgba(207,92,92,0.13);
                color: {COLORS["red"]};
            }}
            div[data-testid="stDataEditor"] {{
                border: 1px solid {COLORS["border"]};
                border-radius: 18px;
                overflow: hidden;
                background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%);
                box-shadow: 0 12px 28px rgba(22,50,79,0.04);
            }}
            div[data-testid="stDataEditor"] [data-testid="stDataFrameResizable"] {{
                background: transparent;
            }}
            div[data-testid="stDataEditor"] [role="grid"] {{
                background: transparent;
            }}
            div[data-testid="stDataEditor"] [role="columnheader"] {{
                background: linear-gradient(180deg, #edf4ff 0%, #e3edfb 100%);
                color: #55708f;
                font-size: 0.72rem;
                text-transform: uppercase;
                letter-spacing: 0.12em;
                font-weight: 800;
                border-bottom: 1px solid #d6e2f0;
            }}
            div[data-testid="stDataEditor"] [role="gridcell"] {{
                background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
                color: #23384f;
                border-bottom: 1px solid #e6edf6;
            }}
            div[data-testid="stDataEditor"] [role="gridcell"]:focus-within {{
                outline: 1px solid #8eb0f1;
                background: #f5f9ff;
            }}
            div[data-testid="stDataEditor"] input,
            div[data-testid="stDataEditor"] textarea,
            div[data-testid="stDataEditor"] [data-baseweb="select"] > div {{
                background: linear-gradient(180deg, #ffffff 0%, #f5f9ff 100%);
                color: #23384f;
            }}
            div[data-testid="stDataEditor"] button[kind="secondary"] {{
                background: linear-gradient(180deg, #eef4ff 0%, #deebff 100%);
                color: #2f6fed;
                border-color: #a7c1ef;
            }}
            .stButton > button {{
                border-radius: 999px;
                border: 1px solid {COLORS["blue"]};
                min-height: 2.75rem;
                font-weight: 700;
            }}
            .stButton > button:not([kind="primary"]) {{
                background: linear-gradient(180deg, #ffffff 0%, #f4f8ff 100%);
                color: {COLORS["navy"]};
            }}
            .stButton > button[kind="primary"] {{
                background: linear-gradient(90deg, {COLORS["blue"]}, #5b8ef1);
                color: white;
            }}
            .program-grid-link {{
                display: block;
                width: 100%;
                text-align: center;
                text-decoration: none;
                padding: 0.78rem 0.95rem;
                border-radius: 1.2rem;
                background: linear-gradient(180deg, #6f98ec 0%, #4f7fe6 100%);
                color: white !important;
                font-size: 0.95rem;
                font-weight: 800;
                line-height: 1.2;
                box-shadow: 0 10px 22px rgba(79,127,230,0.18);
                border: 1px solid rgba(79,127,230,0.55);
                transition: transform 120ms ease, box-shadow 120ms ease, filter 120ms ease;
            }}
            .program-grid-link:hover {{
                color: white !important;
                filter: brightness(1.02);
                box-shadow: 0 14px 26px rgba(79,127,230,0.22);
                transform: translateY(-1px);
            }}
            @media (max-width: 1100px) {{
                .metric-strip {{
                    grid-template-columns: repeat(2, minmax(0, 1fr));
                }}
                .dash-kpi-grid,
                .dashboard-main-grid,
                .dashboard-page-grid {{
                    grid-template-columns: 1fr;
                }}
                .topbar {{
                    grid-template-columns: 1fr;
                }}
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def load_portfolio_dataframe(portfolio: str) -> pd.DataFrame:
    df = pd.DataFrame(PORTFOLIO_SEEDS[portfolio]).copy()
    for col in ["Start", "End", "Milestone Date"]:
        df[col] = pd.to_datetime(df[col])
    return df


def ensure_state(portfolio: str) -> pd.DataFrame:
    key = f"portfolio_data_{portfolio}"
    if key not in st.session_state:
        st.session_state[key] = load_portfolio_dataframe(portfolio)
    return st.session_state[key].copy()


def save_state(portfolio: str, df: pd.DataFrame) -> None:
    st.session_state[f"portfolio_data_{portfolio}"] = df.copy()


def default_program_details(row: pd.Series, reporting_date: date) -> dict[str, object]:
    milestone_date = pd.to_datetime(row["Milestone Date"])
    week_ending = pd.Timestamp(reporting_date)
    milestones = pd.DataFrame(
        [
            {
                "Milestone Name": row["Milestone"],
                "Planned Date": milestone_date,
                "Forecast Date": milestone_date,
                "Status": row["Status"],
                "Comment": row["Status Note"],
            },
            {
                "Milestone Name": "Pilot Readiness Review",
                "Planned Date": milestone_date + timedelta(days=21),
                "Forecast Date": milestone_date + timedelta(days=21),
                "Status": "On Track",
                "Comment": "Readiness checkpoint scheduled with regional leads.",
            },
        ]
    )
    risks = pd.DataFrame(
        [
            {
                "Severity": "High" if row["Status"] == "At Risk" else "Medium",
                "Title": row["Milestone"],
                "Owner": row["Lead"],
                "Target Date": milestone_date,
                "Description": row["Risk Detail"],
                "Mitigation Plan": row["Mitigation"],
            }
        ]
    )
    decisions = pd.DataFrame(
        [
            {
                "Decision Topic": row["Decision Needed"] if "No " not in str(row["Decision Needed"]) else "No executive decision requested",
                "Required By": milestone_date - timedelta(days=7),
                "Impact if Unresolved": row["Status Note"],
                "Recommendation": row["Next Step"],
            }
        ]
    )
    return {
        "week_ending": week_ending,
        "current_phase": row["Stage"],
        "overall_status": row["Status"],
        "percent_complete": int(row["Progress"]),
        "trend": "Flat",
        "accomplishments": row["Summary"],
        "next_steps": row["Upcoming Work"],
        "dependencies": row["Risk Detail"],
        "executive_summary": row["Summary"],
        "milestones": milestones,
        "risks": risks,
        "decisions": decisions,
    }


def ensure_program_details(portfolio: str, program: str, reporting_date: date, row: pd.Series) -> dict[str, object]:
    key = f"weekly_detail_{portfolio}_{program}"
    if key not in st.session_state:
        st.session_state[key] = default_program_details(row, reporting_date)
    return {
        "week_ending": st.session_state[key]["week_ending"],
        "current_phase": st.session_state[key]["current_phase"],
        "overall_status": st.session_state[key]["overall_status"],
        "percent_complete": st.session_state[key]["percent_complete"],
        "trend": st.session_state[key]["trend"],
        "accomplishments": st.session_state[key]["accomplishments"],
        "next_steps": st.session_state[key]["next_steps"],
        "dependencies": st.session_state[key]["dependencies"],
        "executive_summary": st.session_state[key]["executive_summary"],
        "milestones": st.session_state[key]["milestones"].copy(),
        "risks": st.session_state[key]["risks"].copy(),
        "decisions": st.session_state[key]["decisions"].copy(),
    }


def save_program_details(portfolio: str, program: str, details: dict[str, object]) -> None:
    st.session_state[f"weekly_detail_{portfolio}_{program}"] = {
        "week_ending": details["week_ending"],
        "current_phase": details["current_phase"],
        "overall_status": details["overall_status"],
        "percent_complete": int(details["percent_complete"]),
        "trend": details["trend"],
        "accomplishments": details["accomplishments"],
        "next_steps": details["next_steps"],
        "dependencies": details["dependencies"],
        "executive_summary": details["executive_summary"],
        "milestones": details["milestones"].copy(),
        "risks": details["risks"].copy(),
        "decisions": details["decisions"].copy(),
    }


def clear_milestone_editor_state(portfolio: str, program: str) -> None:
    prefix = f"milestone_{portfolio}_{program}_"
    for key in list(st.session_state.keys()):
        if key.startswith(prefix):
            del st.session_state[key]


def render_milestone_editor(portfolio: str, program: str, milestones: pd.DataFrame) -> pd.DataFrame:
    prefix = f"milestone_{portfolio}_{program}"
    count_key = f"{prefix}_count"
    if count_key not in st.session_state:
        seed = milestones.reset_index(drop=True)
        st.session_state[count_key] = len(seed)
        for idx, row in seed.iterrows():
            st.session_state[f"{prefix}_name_{idx}"] = row["Milestone Name"]
            st.session_state[f"{prefix}_planned_{idx}"] = pd.to_datetime(row["Planned Date"]).date()
            st.session_state[f"{prefix}_forecast_{idx}"] = pd.to_datetime(row["Forecast Date"]).date()
            st.session_state[f"{prefix}_status_{idx}"] = row["Status"]
            st.session_state[f"{prefix}_comment_{idx}"] = row["Comment"]

    add_col, _ = st.columns([0.3, 1.7])
    if add_col.button("Add Milestone", key=f"{prefix}_add", use_container_width=True):
        idx = st.session_state[count_key]
        today = date.today()
        st.session_state[f"{prefix}_name_{idx}"] = ""
        st.session_state[f"{prefix}_planned_{idx}"] = today
        st.session_state[f"{prefix}_forecast_{idx}"] = today
        st.session_state[f"{prefix}_status_{idx}"] = "On Track"
        st.session_state[f"{prefix}_comment_{idx}"] = ""
        st.session_state[count_key] = idx + 1
        st.rerun()

    header_cols = st.columns([1.35, 0.9, 0.9, 0.8, 1.55, 0.35], gap="small")
    headers = ["Milestone Name", "Planned Date", "Forecast Date", "Status", "Comment", ""]
    for col, label in zip(header_cols, headers):
        col.markdown(f'<div class="metric-title" style="margin-bottom:0.35rem;">{label}</div>', unsafe_allow_html=True)

    rows = []
    current_count = st.session_state[count_key]
    for idx in range(current_count):
        cols = st.columns([1.35, 0.9, 0.9, 0.8, 1.55, 0.35], gap="small")
        name = cols[0].text_input("Milestone Name", key=f"{prefix}_name_{idx}", label_visibility="collapsed")
        planned = cols[1].date_input("Planned Date", key=f"{prefix}_planned_{idx}", label_visibility="collapsed")
        forecast = cols[2].date_input("Forecast Date", key=f"{prefix}_forecast_{idx}", label_visibility="collapsed")
        status = cols[3].selectbox("Status", ["On Track", "Needs Attention", "At Risk"], key=f"{prefix}_status_{idx}", label_visibility="collapsed")
        comment = cols[4].text_input("Comment", key=f"{prefix}_comment_{idx}", label_visibility="collapsed")
        remove = cols[5].button("X", key=f"{prefix}_remove_{idx}", use_container_width=True)
        if remove:
            for move_idx in range(idx, current_count - 1):
                for field in ["name", "planned", "forecast", "status", "comment"]:
                    st.session_state[f"{prefix}_{field}_{move_idx}"] = st.session_state.get(f"{prefix}_{field}_{move_idx + 1}")
            for field in ["name", "planned", "forecast", "status", "comment"]:
                st.session_state.pop(f"{prefix}_{field}_{current_count - 1}", None)
            st.session_state[count_key] = current_count - 1
            st.rerun()
        if str(name).strip():
            rows.append(
                {
                    "Milestone Name": name,
                    "Planned Date": pd.to_datetime(planned),
                    "Forecast Date": pd.to_datetime(forecast),
                    "Status": status,
                    "Comment": comment,
                }
            )

    return pd.DataFrame(rows, columns=["Milestone Name", "Planned Date", "Forecast Date", "Status", "Comment"])


def render_risk_editor(portfolio: str, program: str, risks: pd.DataFrame) -> pd.DataFrame:
    prefix = f"risk_{portfolio}_{program}"
    count_key = f"{prefix}_count"
    if count_key not in st.session_state:
        seed = risks.reset_index(drop=True)
        st.session_state[count_key] = len(seed)
        for idx, row in seed.iterrows():
            st.session_state[f"{prefix}_severity_{idx}"] = row["Severity"]
            st.session_state[f"{prefix}_title_{idx}"] = row["Title"]
            st.session_state[f"{prefix}_owner_{idx}"] = row["Owner"]
            st.session_state[f"{prefix}_target_{idx}"] = pd.to_datetime(row["Target Date"]).date()
            st.session_state[f"{prefix}_description_{idx}"] = row["Description"]
            st.session_state[f"{prefix}_mitigation_{idx}"] = row["Mitigation Plan"]

    add_col, _ = st.columns([0.24, 1.76])
    if add_col.button("Add Risk", key=f"{prefix}_add", use_container_width=True):
        idx = st.session_state[count_key]
        today = date.today()
        st.session_state[f"{prefix}_severity_{idx}"] = "Medium"
        st.session_state[f"{prefix}_title_{idx}"] = ""
        st.session_state[f"{prefix}_owner_{idx}"] = ""
        st.session_state[f"{prefix}_target_{idx}"] = today
        st.session_state[f"{prefix}_description_{idx}"] = ""
        st.session_state[f"{prefix}_mitigation_{idx}"] = ""
        st.session_state[count_key] = idx + 1
        st.rerun()

    header_cols = st.columns([0.6, 1.15, 0.9, 0.95, 1.7, 1.7, 0.28], gap="small")
    headers = ["Severity", "Risk Title", "Owner", "Target Date", "Description", "Mitigation", ""]
    for col, label in zip(header_cols, headers):
        col.markdown(f'<div class="metric-title" style="margin-bottom:0.35rem;">{label}</div>', unsafe_allow_html=True)

    rows = []
    current_count = st.session_state[count_key]
    for idx in range(current_count):
        cols = st.columns([0.6, 1.15, 0.9, 0.95, 1.7, 1.7, 0.28], gap="small")
        severity = cols[0].selectbox("Severity", ["High", "Medium", "Low", "DEP"], key=f"{prefix}_severity_{idx}", label_visibility="collapsed")
        title = cols[1].text_input("Risk Title", key=f"{prefix}_title_{idx}", label_visibility="collapsed")
        owner = cols[2].text_input("Owner", key=f"{prefix}_owner_{idx}", label_visibility="collapsed")
        target = cols[3].date_input("Target Date", key=f"{prefix}_target_{idx}", label_visibility="collapsed")
        description = cols[4].text_input("Description", key=f"{prefix}_description_{idx}", label_visibility="collapsed")
        mitigation = cols[5].text_input("Mitigation", key=f"{prefix}_mitigation_{idx}", label_visibility="collapsed")
        remove = cols[6].button("X", key=f"{prefix}_remove_{idx}", use_container_width=True)
        if remove:
            for move_idx in range(idx, current_count - 1):
                for field in ["severity", "title", "owner", "target", "description", "mitigation"]:
                    st.session_state[f"{prefix}_{field}_{move_idx}"] = st.session_state.get(f"{prefix}_{field}_{move_idx + 1}")
            for field in ["severity", "title", "owner", "target", "description", "mitigation"]:
                st.session_state.pop(f"{prefix}_{field}_{current_count - 1}", None)
            st.session_state[count_key] = current_count - 1
            st.rerun()
        if str(title).strip():
            rows.append(
                {
                    "Severity": severity,
                    "Title": title,
                    "Owner": owner,
                    "Target Date": pd.to_datetime(target),
                    "Description": description,
                    "Mitigation Plan": mitigation,
                }
            )

    return pd.DataFrame(rows, columns=["Severity", "Title", "Owner", "Target Date", "Description", "Mitigation Plan"])


def render_decision_editor(portfolio: str, program: str, decisions: pd.DataFrame) -> pd.DataFrame:
    prefix = f"decision_{portfolio}_{program}"
    count_key = f"{prefix}_count"
    if count_key not in st.session_state:
        seed = decisions.reset_index(drop=True)
        st.session_state[count_key] = len(seed)
        for idx, row in seed.iterrows():
            st.session_state[f"{prefix}_topic_{idx}"] = row["Decision Topic"]
            st.session_state[f"{prefix}_required_{idx}"] = pd.to_datetime(row["Required By"]).date()
            st.session_state[f"{prefix}_impact_{idx}"] = row["Impact if Unresolved"]
            st.session_state[f"{prefix}_recommendation_{idx}"] = row["Recommendation"]

    add_col, _ = st.columns([0.28, 1.72])
    if add_col.button("Add Request", key=f"{prefix}_add", use_container_width=True):
        idx = st.session_state[count_key]
        st.session_state[f"{prefix}_topic_{idx}"] = ""
        st.session_state[f"{prefix}_required_{idx}"] = date.today()
        st.session_state[f"{prefix}_impact_{idx}"] = ""
        st.session_state[f"{prefix}_recommendation_{idx}"] = ""
        st.session_state[count_key] = idx + 1
        st.rerun()

    header_cols = st.columns([1.45, 0.85, 1.55, 1.3, 0.28], gap="small")
    headers = ["Decision Topic", "Required By", "Impact if Unresolved", "Recommendation", ""]
    for col, label in zip(header_cols, headers):
        col.markdown(f'<div class="metric-title" style="margin-bottom:0.35rem;">{label}</div>', unsafe_allow_html=True)

    rows = []
    current_count = st.session_state[count_key]
    for idx in range(current_count):
        cols = st.columns([1.45, 0.85, 1.55, 1.3, 0.28], gap="small")
        topic = cols[0].text_input("Decision Topic", key=f"{prefix}_topic_{idx}", label_visibility="collapsed")
        required = cols[1].date_input("Required By", key=f"{prefix}_required_{idx}", label_visibility="collapsed")
        impact = cols[2].text_input("Impact if Unresolved", key=f"{prefix}_impact_{idx}", label_visibility="collapsed")
        recommendation = cols[3].text_input("Recommendation", key=f"{prefix}_recommendation_{idx}", label_visibility="collapsed")
        remove = cols[4].button("X", key=f"{prefix}_remove_{idx}", use_container_width=True)
        if remove:
            for move_idx in range(idx, current_count - 1):
                for field in ["topic", "required", "impact", "recommendation"]:
                    st.session_state[f"{prefix}_{field}_{move_idx}"] = st.session_state.get(f"{prefix}_{field}_{move_idx + 1}")
            for field in ["topic", "required", "impact", "recommendation"]:
                st.session_state.pop(f"{prefix}_{field}_{current_count - 1}", None)
            st.session_state[count_key] = current_count - 1
            st.rerun()
        if str(topic).strip():
            rows.append(
                {
                    "Decision Topic": topic,
                    "Required By": pd.to_datetime(required),
                    "Impact if Unresolved": impact,
                    "Recommendation": recommendation,
                }
            )

    return pd.DataFrame(rows, columns=["Decision Topic", "Required By", "Impact if Unresolved", "Recommendation"])


def status_class(status: str) -> str:
    return {
        "On Track": "status-green",
        "Needs Attention": "status-yellow",
        "At Risk": "status-red",
    }.get(status, "status-green")


def render_html(html: str) -> None:
    normalized = dedent(html).strip()
    normalized = re.sub(r"(?m)^[ \t]+(?=<)", "", normalized)
    st.markdown(normalized, unsafe_allow_html=True)


def metric_card(title: str, value: str | int, note: str) -> str:
    return f"""
    <div class="metric-card">
        <div class="metric-title">{title}</div>
        <div class="metric-number">{value}</div>
        <div class="metric-note">{note}</div>
    </div>
    """


def render_header(page: str, portfolio: str, reporting_date: date) -> None:
    if page in {"Impower Portfolio", "Program One-Pager"}:
        return
    badge_map = {
        "Impower Portfolio": "Impower Portfolio",
        "Program One-Pager": "Program One-Pager",
        "Weekly Updates": "Weekly Updates",
        "Settings": "Settings",
        "Help & Support": "Help & Support",
    }
    badge = badge_map[page]
    render_html(
        f"""
        <div class="hero">
            <div class="hero-title">Portfolio Command Center</div>
            <div class="hero-badge">{badge}</div>
        </div>
        <div class="subtitle">
            <strong>{PORTFOLIO_NAME}</strong> for <strong>{portfolio}</strong>. Reporting date:
            <strong>{reporting_date:%B %d, %Y}</strong>.
        </div>
        """,
    )


def render_program_picker(title: str, help_text: str, options: list[str], key: str) -> str:
    with st.container(border=True):
        render_html(
            f"""
            <div class="eyebrow" style="color:{COLORS["blue"]}; margin-bottom:0.45rem;">Program Selection</div>
            <div class="heading" style="display:flex; align-items:center; gap:0.5rem; margin-bottom:0.35rem;">
                {title}
                <span style="display:inline-flex; width:22px; height:22px; border-radius:999px; align-items:center; justify-content:center; background:#2f6fed; color:white; font-size:0.8rem; font-weight:800;" title="{help_text}">i</span>
            </div>
            <div class="copy" style="margin-bottom:0.7rem;">{help_text}</div>
            """
        )
        value = st.selectbox(
            title,
            options,
            key=key,
            label_visibility="collapsed",
            help=help_text,
        )
        return value


def render_metric_strip(df: pd.DataFrame) -> None:
    decisions_pending = int((~df["Decision Needed"].str.contains("No ", na=False)).sum())
    upcoming_milestones = int(df["Milestone Date"].notna().sum())
    st.markdown(
        f"""
        <div class="metric-strip">
            {metric_card("Total Programs", len(df), "Programs currently in the selected portfolio view.")}
            {metric_card("At Risk", int(df["Status"].eq("At Risk").sum()), "Programs with escalated delivery pressure.")}
            {metric_card("Decisions Pending", decisions_pending, "Leadership decisions currently called out by program leads.")}
            {metric_card("Upcoming Milestones", upcoming_milestones, "Active milestone checkpoints across the portfolio.")}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_app_navigation(current_page: str) -> str:
    nav_items = [
        "Impower Portfolio",
        "Program One-Pager",
        "Weekly Updates",
        "Settings",
        "Help & Support",
    ]

    with st.container(border=True):
        render_html(
            """
            <div class="eyebrow" style="color:#2f6fed; margin-bottom:0.6rem;">Navigation</div>
            """
        )
        for label in nav_items[:3]:
            if st.button(
                label,
                type="primary" if label == current_page else "secondary",
                use_container_width=True,
                key=f"nav_{label}",
            ):
                st.session_state["current_page"] = label
                st.rerun()
        render_html('<div class="rail-divider" style="margin:0.9rem 0;"></div>')
        for label in nav_items[3:]:
            if st.button(
                label,
                type="primary" if label == current_page else "secondary",
                use_container_width=True,
                key=f"nav_{label}",
            ):
                st.session_state["current_page"] = label
                st.rerun()

    return st.session_state["current_page"]


def sync_selected_program() -> None:
    st.session_state["selected_program"] = st.session_state["sidebar_selected_program"]


def navigate_to_program(program: str) -> None:
    st.session_state["pending_selected_program"] = program
    st.session_state["current_page"] = "Program One-Pager"
    st.rerun()


def timeline_chart(df: pd.DataFrame) -> alt.Chart:
    timeline = df[["Program", "Stage", "Start", "End", "Milestone", "Milestone Date"]].copy()
    return (
        alt.Chart(timeline)
        .mark_bar(cornerRadius=10, stroke="#17396c", strokeWidth=1.5, height=26)
        .encode(
            x=alt.X("Start:T", axis=alt.Axis(format="%b", labelAngle=0), title=None),
            x2="End:T",
            y=alt.Y("Program:N", sort=None, title=None),
            color=alt.Color(
                "Stage:N",
                legend=None,
                scale=alt.Scale(
                    domain=["Discovery", "Phase 1", "Phase 2", "Phase 3", "Phase 4"],
                    range=["#c7d7ee", "#bad0ea", "#a8c2e0", "#90b3d9", "#7aa4d2"],
                ),
            ),
            tooltip=[
                "Program",
                "Stage",
                alt.Tooltip("Start:T", title="Start"),
                alt.Tooltip("End:T", title="End"),
                "Milestone",
                alt.Tooltip("Milestone Date:T", title="Milestone Date"),
            ],
        )
        .properties(height=max(220, 54 * len(timeline)))
        .configure_view(strokeWidth=0)
        .configure_axis(gridColor="#d6dde8", labelColor="#223347", domain=False)
    )


def status_rollup_chart(df: pd.DataFrame) -> alt.Chart:
    chart_df = (
        df["Status"]
        .value_counts()
        .reindex(["On Track", "Needs Attention", "At Risk"], fill_value=0)
        .rename_axis("Status")
        .reset_index(name="Programs")
    )
    return (
        alt.Chart(chart_df)
        .mark_arc(innerRadius=70, outerRadius=110)
        .encode(
            theta=alt.Theta("Programs:Q"),
            color=alt.Color(
                "Status:N",
                legend=None,
                scale=alt.Scale(
                    domain=["On Track", "Needs Attention", "At Risk"],
                    range=[COLORS["green"], COLORS["yellow"], COLORS["red"]],
                ),
            ),
            tooltip=["Status", "Programs"],
        )
        .properties(height=300)
        .configure_view(strokeWidth=0)
    )


def milestone_table(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df[["Program", "Milestone", "Milestone Date", "Status"]]
        .sort_values("Milestone Date")
        .head(5)
        .reset_index(drop=True)
    )


def risk_table(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df[["Program", "Open Risks", "Escalations", "Risk Detail", "Mitigation"]]
        .sort_values(["Escalations", "Open Risks"], ascending=False)
        .head(5)
        .reset_index(drop=True)
    )


def decision_table(df: pd.DataFrame) -> pd.DataFrame:
    decisions = df[["Program", "Executive Sponsor", "Decision Needed", "Priority"]].copy()
    decisions = decisions[~decisions["Decision Needed"].str.contains("No ", na=False)]
    return decisions.reset_index(drop=True)


def cycle_label(reporting_date: date) -> str:
    start = reporting_date - timedelta(days=6)
    return f"{start:%d %b} - {reporting_date:%d %b %Y}"


def one_pager_status_label(status: str) -> str:
    return {
        "On Track": "GREEN",
        "Needs Attention": "AMBER",
        "At Risk": "RED",
    }.get(status, "GREEN")


def one_pager_status_class(status: str) -> str:
    return {
        "On Track": "op-green",
        "Needs Attention": "op-amber",
        "At Risk": "op-red",
    }.get(status, "op-green")


def split_bullets(text: str, limit: int = 4) -> list[str]:
    parts = [part.strip().capitalize() for part in str(text).split(",") if part.strip()]
    return parts[:limit] if parts else [str(text)]


def one_pager_accomplishments(row: pd.Series) -> list[str]:
    program = row["Program"]
    milestone = row["Milestone"]
    progress = int(row["Progress"])
    return [
        f"{program} advanced to {progress}% completion with core work aligned to the current plan.",
        f"{milestone} preparation remains active with milestone readiness reviews in motion.",
        f"Program leadership maintained delivery governance and sponsor visibility through the current cycle.",
        f"Cross-functional coordination remained focused on {str(row['Next Step']).lower()}.",
    ]


def one_pager_risks(row: pd.Series) -> list[dict[str, str]]:
    lead_initial = str(row["Lead"]).split()[-1]
    sponsor_initial = str(row["Executive Sponsor"]).split()[0]
    milestone_date = pd.to_datetime(row["Milestone Date"])
    return [
        {
            "severity": "HIGH" if row["Status"] == "At Risk" else ("MED" if row["Status"] == "Needs Attention" else "LOW"),
            "description": str(row["Risk Detail"]),
            "owner": str(row["Lead"]),
            "mitigation": str(row["Mitigation"]),
            "target": f"{milestone_date:%d %b %Y}",
        },
        {
            "severity": "MED",
            "description": f"Decision turnaround could affect readiness confidence for {row['Milestone'].lower()}.",
            "owner": sponsor_initial,
            "mitigation": "Weekly steering review and decision pre-read in advance of checkpoint.",
            "target": f"{(milestone_date - timedelta(days=10)):%d %b %Y}",
        },
        {
            "severity": "DEP",
            "description": f"Dependent on cross-team support to sustain execution against {row['Stage'].lower()} commitments.",
            "owner": lead_initial,
            "mitigation": "Integrated dependency review with PMO and workstream leads.",
            "target": f"{(milestone_date - timedelta(days=5)):%d %b %Y}",
        },
    ]


def one_pager_decisions(row: pd.Series) -> list[dict[str, str]]:
    milestone_date = pd.to_datetime(row["Milestone Date"])
    items = []
    if "No " not in str(row["Decision Needed"]):
        items.append(
            {
                "title": str(row["Decision Needed"]),
                "meta": f"Decision owner: {row['Executive Sponsor']} · Due: {(milestone_date - timedelta(days=7)):%d %b %Y}",
            }
        )
    items.append(
        {
            "title": f"Confirm resourcing support required to protect {str(row['Milestone']).lower()}.",
            "meta": f"Decision owner: PMO · Due: {(milestone_date - timedelta(days=12)):%d %b %Y}",
        }
    )
    items.append(
        {
            "title": f"Validate rollout and dependency assumptions for the current {str(row['Stage']).lower()} phase.",
            "meta": f"Decision owner: Steering Committee · Due: {(milestone_date - timedelta(days=5)):%d %b %Y}",
        }
    )
    return items[:3]


def one_pager_workstreams(row: pd.Series) -> list[dict[str, str | int]]:
    progress = int(row["Progress"])
    return [
        {"name": "Process Design", "pct": min(100, progress + 28), "note": "Core process definition and operating model alignment.", "owner": row["Executive Sponsor"]},
        {"name": "Technology Build", "pct": min(100, progress), "note": "Primary build motion and dependency resolution cadence.", "owner": row["Lead"]},
        {"name": "Data Migration", "pct": max(25, min(100, progress + 6)), "note": "Data readiness, mapping, and quality control checkpoints.", "owner": "Data Lead"},
        {"name": "Change Mgmt", "pct": max(20, progress - 12), "note": "Stakeholder readiness, comms, and adoption planning.", "owner": "Change Lead"},
        {"name": "Testing & Readiness", "pct": max(15, progress - 18), "note": "Environment readiness, scripts, and cutover preparation.", "owner": "Readiness Lead"},
    ]


def one_pager_milestones(row: pd.Series) -> list[dict[str, str]]:
    start = pd.to_datetime(row["Start"])
    end = pd.to_datetime(row["End"])
    milestone = pd.to_datetime(row["Milestone Date"])
    return [
        {"name": "Planning Complete", "date": start + timedelta(days=21), "note": "On time"},
        {"name": "Design Finalized", "date": start + timedelta(days=45), "note": "On time"},
        {"name": str(row["Milestone"]), "date": milestone, "note": "Current checkpoint"},
        {"name": "Pilot Readiness", "date": milestone + timedelta(days=42), "note": "On track"},
        {"name": "Cutover Readiness", "date": end - timedelta(days=20), "note": "On track"},
        {"name": "Go-Live", "date": end, "note": "At risk" if row["Status"] == "At Risk" else "On track"},
    ]


def render_program_one_pager(portfolio: str, program: str, df: pd.DataFrame, reporting_date: date) -> None:
    def esc(value: object) -> str:
        return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    row = df.loc[df["Program"] == program].iloc[0]
    status_label = one_pager_status_label(str(row["Status"]))
    status_cls = one_pager_status_class(str(row["Status"]))
    period = cycle_label(reporting_date)
    cycle_num = 24
    total_cycles = 36
    sponsor = str(row["Executive Sponsor"])
    lead = str(row["Lead"])
    pmo = "Sandra Kowalski"
    progress = int(row["Progress"])
    delta = "+6%" if row["Status"] != "At Risk" else "+2%"
    highlights = [
        f"Status {'unchanged from prior cycle' if row['Status'] != 'At Risk' else 'requires elevated attention this cycle'}",
        f"Percent complete {delta} (current {progress}%)",
        f"Risk exposure {'increased' if row['Open Risks'] >= 3 else 'remained stable'} due to {str(row['Risk Detail']).lower()}",
    ]
    accomplishments = one_pager_accomplishments(row)
    upcoming_work = split_bullets(str(row["Upcoming Work"]))
    risks = one_pager_risks(row)
    decisions = one_pager_decisions(row)
    workstreams = one_pager_workstreams(row)
    milestones = one_pager_milestones(row)
    confidence = "Medium" if row["Status"] != "On Track" else "High"
    confidence_delta = "was High" if confidence == "Medium" else "steady"
    risk_rows = []
    for item in risks:
        sev_class = {
            "HIGH": "sev-high",
            "MED": "sev-med",
            "LOW": "sev-low",
            "DEP": "sev-dep",
        }.get(item["severity"], "sev-med")
        risk_rows.append(
            f'<tr><td><span class="sev {sev_class}">{esc(item["severity"])}</span></td><td>{esc(item["description"])}</td><td>{esc(item["owner"])}</td><td>{esc(item["mitigation"])}</td><td>{esc(item["target"])}</td></tr>'
        )

    decision_cards = []
    decision_colors = ["decision-red", "decision-amber", "decision-blue"]
    for idx, item in enumerate(decisions, start=1):
        decision_cards.append(
            f'<div class="decision-card"><div class="decision-num {decision_colors[(idx - 1) % len(decision_colors)]}">{idx}</div><div><div class="decision-title">{esc(item["title"])}</div><div class="decision-meta">{esc(item["meta"])}</div></div></div>'
        )

    status_color = {"GREEN": "#22c55e", "AMBER": "#f59e0b", "RED": "#ef4444"}.get(status_label, "#22c55e")
    phase_label = str(row["Stage"])

    timeline_items = []
    current_index = min(2, len(milestones) - 1)
    for idx, item in enumerate(milestones):
        if idx < current_index:
            node_class = "done"
            icon = "✓"
        elif idx == current_index:
            node_class = "current"
            icon = "✦"
        else:
            node_class = "future"
            icon = "⚑"
        note = esc(item["note"])
        note_class = "note-red" if "risk" in note.lower() or "slip" in note.lower() else ("note-green" if "time" in note.lower() or "early" in note.lower() or "checkpoint" in note.lower() else "note-muted")
        timeline_items.append(
            f'<div class="timeline-step {node_class}"><div class="timeline-node">{icon}</div><div class="timeline-name">{esc(item["name"])}</div><div class="timeline-date">{item["date"]:%d %b %Y}</div><div class="timeline-note {note_class}">{note}</div></div>'
        )

    accomplishment_items = "".join(f"<li>{esc(item)}</li>" for item in accomplishments)
    upcoming_items = "".join(f"<li>{esc(item)}</li>" for item in upcoming_work)

    workstream_cards = []
    for item in workstreams:
        pct = int(item["pct"])
        tone = "ws-green" if pct >= 75 else ("ws-amber" if pct >= 45 else "ws-red")
        workstream_cards.append(
            f'<div class="ws-card"><div class="ws-head"><div class="ws-name">{esc(item["name"])}</div><div class="ws-dot {tone}"></div></div><div class="ws-bar"><span class="{tone}" style="width:{pct}%"></span></div><div class="ws-pct">{pct}%</div><div class="ws-note">{esc(item["note"])}</div><div class="ws-owner">Owner: {esc(item["owner"])}</div></div>'
        )

    summary_text = (
        f"{esc(row['Summary'])} The most significant item for leadership attention is {esc(row['Risk Detail']).lower()}, "
        f"which is being managed through {esc(row['Mitigation']).lower()}."
    )

    hero_bullets = [
        ("neutral", esc(highlights[0])),
        ("green", esc(highlights[1])),
        ("red", esc(highlights[2])),
    ]
    hero_bullet_html = "".join(
        f'<div class="hero-bullet {tone}"><span class="hero-dot"></span><span>{text}</span></div>'
        for tone, text in hero_bullets
    )

    html = f"""
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <style>
        html, body {{ width: 100%; min-height: 100%; overflow-x: hidden; }}
        * {{ box-sizing: border-box; }}
        body {{
            margin: 0;
            background: linear-gradient(180deg, #f4f7fb 0%, #eef3f9 100%);
            font-family: Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            color: #49576a;
        }}
        .page {{ width: 100%; max-width: 1560px; margin: 0 auto; padding: 10px 8px 28px; }}
        .hero {{
            background: #0f1837;
            color: white;
            border-radius: 22px;
            padding: 20px 24px 18px;
            box-shadow: 0 18px 42px rgba(8,16,40,0.18);
        }}
        .hero-top {{
            display: grid;
            grid-template-columns: minmax(0, 1.5fr) minmax(420px, 0.9fr);
            gap: 24px;
            align-items: start;
        }}
        .hero-title {{ font-size: 28px; font-weight: 800; letter-spacing: -0.03em; color: #ffffff; }}
        .hero-sub {{ margin-top: 10px; font-size: 16px; color: #9bb2e8; }}
        .hero-owners {{
            display: flex;
            flex-wrap: wrap;
            gap: 28px;
            margin-top: 22px;
            color: #a8b6d7;
            font-size: 14px;
        }}
        .hero-owners strong {{ color: white; margin-left: 6px; }}
        .hero-metrics {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 16px;
        }}
        .hero-metric-label {{ font-size: 12px; text-transform: uppercase; color: #9aa9cc; font-weight: 700; }}
        .hero-metric-value {{ margin-top: 8px; font-size: 18px; font-weight: 800; color: white; display:flex; align-items:center; gap:10px; }}
        .hero-status-dot {{ width: 18px; height: 18px; border-radius: 999px; background: {status_color}; display:inline-block; }}
        .hero-delta {{ font-size: 16px; color: #41d77a; font-weight: 800; }}
        .hero-divider {{ height: 1px; background: rgba(255,255,255,0.16); margin: 18px 0 14px; }}
        .hero-bottom {{ display: flex; flex-wrap: wrap; gap: 28px; }}
        .hero-bullet {{ display: flex; align-items: center; gap: 10px; font-size: 14px; }}
        .hero-dot {{ width: 10px; height: 10px; border-radius: 999px; background: #9ca3af; display:inline-block; }}
        .hero-bullet.green .hero-dot {{ background: #3ddb79; }}
        .hero-bullet.red .hero-dot {{ background: #ff6b6b; }}
        .hero-bullet span:last-child {{ color: #d9e2f8; }}
        .panel {{
            margin-top: 16px;
            background: linear-gradient(180deg, #ffffff 0%, #fafcff 100%);
            border: 1px solid #d8e1ed;
            border-radius: 22px;
            padding: 18px 22px;
            box-shadow: 0 12px 28px rgba(20,42,71,0.04);
        }}
        .section-title {{ display:flex; align-items:center; gap:10px; font-size: 16px; font-weight: 800; color: #1a2741; margin-bottom: 16px; }}
        .section-icon {{ color: #2f5bd2; font-size: 15px; }}
        .summary-copy {{ font-size: 15px; line-height: 1.45; color: #49576a; }}
        .timeline-wrap {{ position: relative; padding: 20px 6px 2px; }}
        .timeline-line {{
            position: absolute; left: 5%; right: 5%; top: 38px; height: 8px; border-radius: 999px;
            background: linear-gradient(90deg, #2f55c7 0%, #2f55c7 48%, #dfe4eb 48%, #dfe4eb 100%);
        }}
        .timeline-grid {{
            position: relative;
            display: grid;
            grid-template-columns: repeat(6, minmax(0, 1fr));
            gap: 10px;
        }}
        .timeline-step {{ text-align: center; }}
        .timeline-node {{
            width: 40px; height: 40px; border-radius: 999px; margin: 0 auto 10px;
            display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 18px;
            border: 3px solid #d8dce4; background: #eceff4; color: #8c93a1; position: relative; z-index: 2;
        }}
        .timeline-step.done .timeline-node {{ background: #2f55c7; border-color: #89a4ee; color: white; }}
        .timeline-step.current .timeline-node {{ background: #f6a609; border-color: #ffd58a; color: white; }}
        .timeline-name {{ font-size: 13px; font-weight: 800; color: #2a4cb8; line-height: 1.25; min-height: 34px; }}
        .timeline-step.future .timeline-name {{ color: #737b87; }}
        .timeline-step.current .timeline-name {{ color: #f29900; }}
        .timeline-date {{ margin-top: 6px; font-size: 12px; color: #7d8796; }}
        .timeline-note {{ margin-top: 6px; font-size: 12px; font-weight: 700; }}
        .note-green {{ color: #18a34a; }}
        .note-red {{ color: #ef4444; }}
        .note-muted {{ color: #97a0ad; }}
        .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-top: 16px; }}
        .list {{ margin: 0; padding-left: 1.2rem; color: #49576a; }}
        .list li {{ margin-bottom: 10px; line-height: 1.42; }}
        .risk-decision {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-top: 16px; }}
        .risk-table {{ width: 100%; border-collapse: collapse; }}
        .risk-table th {{ text-align: left; font-size: 11px; text-transform: uppercase; color: #9ba3af; padding: 10px 8px; border-bottom: 1px solid #e5e9f0; }}
        .risk-table td {{ padding: 10px 8px; border-bottom: 1px solid #eef2f7; vertical-align: top; font-size: 13px; color: #4f5d6f; }}
        .risk-table tr:last-child td {{ border-bottom: none; }}
        .sev {{ font-size: 12px; font-weight: 800; }}
        .sev-high {{ color: #dc2626; }}
        .sev-med {{ color: #b45309; }}
        .sev-low {{ color: #16a34a; }}
        .sev-dep {{ color: #2563eb; }}
        .decision-shell {{ background: #eef3ff; border: 1px solid #b8cbff; border-radius: 18px; padding: 16px; }}
        .decision-card {{
            display: grid; grid-template-columns: 34px 1fr; gap: 14px; align-items: start;
            background: white; border: 1px solid #d8e1f0; border-radius: 14px; padding: 14px 16px; margin-bottom: 14px;
        }}
        .decision-card:last-child {{ margin-bottom: 0; }}
        .decision-num {{
            width: 30px; height: 30px; border-radius: 999px; color: white; display: flex; align-items: center; justify-content: center;
            font-weight: 800; font-size: 13px;
        }}
        .decision-red {{ background: #ef4444; }}
        .decision-amber {{ background: #f59e0b; }}
        .decision-blue {{ background: #2f55c7; }}
        .decision-title {{ font-size: 14px; font-weight: 800; color: #1d2942; line-height: 1.35; }}
        .decision-meta {{ margin-top: 5px; font-size: 12px; color: #7f8997; }}
        .workstream-panel {{ margin-top: 16px; }}
        .workstream-grid {{ display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 16px; margin-top: 14px; }}
        .ws-card {{ border: 1px solid #dde5ef; border-radius: 16px; padding: 16px; background: white; }}
        .ws-head {{ display: flex; justify-content: space-between; align-items: center; gap: 12px; }}
        .ws-name {{ font-size: 14px; font-weight: 800; color: #1d2942; }}
        .ws-dot {{ width: 14px; height: 14px; border-radius: 999px; }}
        .ws-green {{ background: #22c55e; }}
        .ws-amber {{ background: #f59e0b; }}
        .ws-red {{ background: #ef4444; }}
        .ws-bar {{ margin-top: 14px; height: 10px; border-radius: 999px; background: #e7ebf0; overflow: hidden; }}
        .ws-bar span {{ display:block; height:100%; border-radius:999px; }}
        .ws-pct {{ margin-top: 8px; font-size: 13px; font-weight: 700; color: #7a8596; text-align: right; }}
        .ws-note {{ margin-top: 12px; font-size: 13px; line-height: 1.38; color: #5c6776; }}
        .ws-owner {{ margin-top: 12px; font-size: 12px; color: #98a1ae; }}
        @media (max-width: 1220px) {{
            .hero-top, .two-col, .risk-decision {{ grid-template-columns: 1fr; }}
            .hero-metrics {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
            .workstream-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
        }}
        @media (max-width: 820px) {{
            .page {{ padding: 8px 4px 24px; }}
            .hero-metrics, .timeline-grid, .workstream-grid {{ grid-template-columns: 1fr; }}
            .timeline-line {{ display:none; }}
        }}
    </style>
    </head>
    <body>
        <div class="page">
            <div class="hero">
                <div class="hero-top">
                    <div>
                        <div class="hero-title">{esc(program)}</div>
                        <div class="hero-sub">{esc(portfolio)} transformation program overview</div>
                        <div class="hero-owners">
                            <div>Executive Sponsor:<strong>{esc(sponsor)}</strong></div>
                            <div>Program Lead:<strong>{esc(lead)}</strong></div>
                            <div>PMO:<strong>{esc(pmo)}</strong></div>
                        </div>
                    </div>
                    <div class="hero-metrics">
                        <div>
                            <div class="hero-metric-label">Overall Status</div>
                            <div class="hero-metric-value"><span class="hero-status-dot"></span>{status_label}</div>
                        </div>
                        <div>
                            <div class="hero-metric-label">% Complete</div>
                            <div class="hero-metric-value">{progress}% <span class="hero-delta">{delta}</span></div>
                        </div>
                        <div>
                            <div class="hero-metric-label">Current Phase</div>
                            <div class="hero-metric-value">{esc(phase_label)}</div>
                        </div>
                    </div>
                </div>
                <div class="hero-divider"></div>
                <div class="hero-bottom">{hero_bullet_html}</div>
            </div>

            <div class="panel">
                <div class="section-title"><span class="section-icon">▣</span>Executive Summary</div>
                <div class="summary-copy">{summary_text}</div>
            </div>

            <div class="panel">
                <div class="section-title"><span class="section-icon">☰</span>Milestone Timeline</div>
                <div class="timeline-wrap">
                    <div class="timeline-line"></div>
                    <div class="timeline-grid">{''.join(timeline_items)}</div>
                </div>
            </div>

            <div class="two-col">
                <div class="panel">
                    <div class="section-title"><span class="section-icon">✳</span>Recent Accomplishments</div>
                    <ul class="list">{accomplishment_items}</ul>
                </div>
                <div class="panel">
                    <div class="section-title"><span class="section-icon">▸</span>Upcoming Work (Next 2 Weeks)</div>
                    <ul class="list">{upcoming_items}</ul>
                </div>
            </div>

            <div class="risk-decision">
                <div class="panel">
                    <div class="section-title"><span class="section-icon">⚠</span>Risks / Issues / Dependencies</div>
                    <table class="risk-table">
                        <thead><tr><th>Severity</th><th>Description</th><th>Owner</th><th>Mitigation</th><th>Target Date</th></tr></thead>
                        <tbody>{''.join(risk_rows)}</tbody>
                    </table>
                </div>
                <div class="panel decision-shell">
                    <div class="section-title"><span class="section-icon">✎</span>Leadership Decisions Needed</div>
                    {''.join(decision_cards)}
                </div>
            </div>

            <div class="panel workstream-panel">
                <div class="section-title"><span class="section-icon">▤</span>Workstream Status</div>
                <div class="workstream-grid">{''.join(workstream_cards)}</div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html, height=2450, scrolling=True)


def render_html_table(df: pd.DataFrame, columns: list[str], rename_map: dict[str, str] | None = None) -> None:
    rename_map = rename_map or {}
    headers = "".join(f"<th>{rename_map.get(col, col)}</th>" for col in columns)
    rows = []
    for _, row in df[columns].iterrows():
        cells = []
        for col in columns:
            value = row[col]
            if pd.isna(value):
                text = ""
            elif isinstance(value, pd.Timestamp):
                text = value.strftime("%b %d, %Y")
            else:
                text = str(value)
            cells.append(f"<td>{text}</td>")
        rows.append(f"<tr>{''.join(cells)}</tr>")
    table_html = f"""
    <table class="portfolio-table">
        <thead><tr>{headers}</tr></thead>
        <tbody>{''.join(rows)}</tbody>
    </table>
    """
    st.markdown(table_html, unsafe_allow_html=True)


def dashboard_status_tag(status: str) -> str:
    tag_class = {
        "On Track": "dash-tag dash-tag-green",
        "Needs Attention": "dash-tag dash-tag-yellow",
        "At Risk": "dash-tag dash-tag-red",
    }.get(status, "dash-tag dash-tag-green")
    return f'<span class="{tag_class}">{status}</span>'


def render_dashboard_program_grid(df: pd.DataFrame) -> None:
    with st.container(border=True):
        render_html(
            """
            <div class="dash-card-title">Portfolio Program Grid</div>
            <div class="dash-card-heading">Weekly Updates</div>
            <div class="dash-card-copy">Program, owner, phase, health, progress, next milestone, and current delivery note.</div>
            """
        )
        header_cols = st.columns([1.25, 0.85, 0.8, 0.45, 0.55, 0.45, 0.95, 1.4], gap="small")
        headers = ["Program", "Owner", "Phase", "RAG", "% Done", "Trend", "Next Milestone", "Status Note"]
        for col, label in zip(header_cols, headers):
            col.markdown(f'<div class="metric-title" style="margin-bottom:0.2rem;">{label}</div>', unsafe_allow_html=True)

        ordered = df.sort_values(["Milestone Date", "Program"]).reset_index(drop=True)
        for idx, (_, row) in enumerate(ordered.iterrows()):
            trend = "↑" if row["Status"] == "On Track" else ("!" if row["Status"] == "Needs Attention" else "↓")
            rag_html = {
                "On Track": '<span class="rag-dot rag-green"></span>',
                "Needs Attention": '<span class="rag-dot rag-yellow"></span>',
                "At Risk": '<span class="rag-dot rag-red"></span>',
            }.get(str(row["Status"]), '<span class="rag-dot rag-green"></span>')
            milestone_date = pd.to_datetime(row["Milestone Date"]).strftime("%b %d, %Y")
            program_url = f'?page={quote("Program One-Pager")}&program={quote(str(row["Program"]))}'
            row_cols = st.columns([1.25, 0.85, 0.8, 0.45, 0.55, 0.45, 0.95, 1.4], gap="small")
            row_cols[0].markdown(
                f'<a class="program-grid-link" href="{program_url}">{row["Program"]}</a>',
                unsafe_allow_html=True,
            )
            row_cols[1].markdown(f'<div class="copy" style="margin-top:0.45rem;">{row["Lead"]}</div>', unsafe_allow_html=True)
            row_cols[2].markdown(f'<div class="copy" style="margin-top:0.45rem;">{row["Stage"]}</div>', unsafe_allow_html=True)
            row_cols[3].markdown(f'<div style="margin-top:0.7rem;">{rag_html}</div>', unsafe_allow_html=True)
            row_cols[4].markdown(f'<div class="copy" style="margin-top:0.45rem;">{int(row["Progress"])}%</div>', unsafe_allow_html=True)
            row_cols[5].markdown(f'<div class="copy" style="margin-top:0.45rem; font-weight:800;">{trend}</div>', unsafe_allow_html=True)
            row_cols[6].markdown(
                f'<div class="copy" style="margin-top:0.2rem;"><strong style="color:{COLORS["navy"]};">{row["Milestone"]}</strong><br>{milestone_date}</div>',
                unsafe_allow_html=True,
            )
            row_cols[7].markdown(f'<div class="copy" style="margin-top:0.45rem;">{row["Status Note"]}</div>', unsafe_allow_html=True)
            if idx < len(ordered) - 1:
                st.markdown("<div style='height:1px;background:#e8eef6;margin:0.2rem 0 0.35rem;'></div>", unsafe_allow_html=True)


def render_dashboard_milestones(df: pd.DataFrame) -> None:
    items = []
    for _, row in milestone_table(df).iterrows():
        date_value = row["Milestone Date"]
        priority_class = "priority-high" if row["Status"] == "On Track" else ("priority-medium" if row["Status"] == "Needs Attention" else "priority-low")
        priority_label = "High" if row["Status"] == "On Track" else ("Medium" if row["Status"] == "Needs Attention" else "Low")
        items.append(
            f'<div class="milestone-item">'
            f'<div class="milestone-datebox"><div class="milestone-day">{date_value.strftime("%d")}</div><div class="milestone-month">{date_value.strftime("%b")}</div></div>'
            f'<div><div class="milestone-title">{row["Milestone"]}</div><div class="milestone-sub">{row["Program"]}</div></div>'
            f'<div class="priority-badge {priority_class}">{priority_label}</div>'
            f'</div>'
        )
    render_html('<div class="dash-card"><div class="section-bar">Upcoming Milestones</div><div class="milestone-list">' + "".join(items) + "</div></div>")


def render_dashboard_risks(df: pd.DataFrame) -> None:
    items = []
    ranked = (
        df[["Program", "Lead", "Open Risks", "Escalations", "Risk Detail", "Mitigation"]]
        .sort_values(["Escalations", "Open Risks"], ascending=False)
        .head(5)
        .reset_index(drop=True)
    )
    trend_classes = ["trend-worse", "trend-stable", "trend-better", "trend-stable", "trend-worse"]
    trend_labels = ["Worse", "Stable", "Better", "Stable", "Worse"]
    for index, (_, row) in enumerate(ranked.iterrows()):
        trend_class = trend_classes[index % len(trend_classes)]
        trend_label = trend_labels[index % len(trend_labels)]
        card_class = trend_class.replace("trend-", "risk-")
        severity = "High" if int(row["Escalations"]) > 1 or int(row["Open Risks"]) >= 5 else ("Medium" if int(row["Open Risks"]) >= 3 else "Low")
        items.append(
            f'<div class="risk-item {card_class}">'
            f'<div class="risk-head"><div class="risk-title">{row["Risk Detail"]}</div><div class="risk-trend {trend_class}">{"↑ " if trend_label == "Worse" else ("↓ " if trend_label == "Better" else "→ ")}{trend_label}</div></div>'
            f'<div class="risk-detail-line"><div class="risk-meta">{row["Program"]}</div><div class="risk-meta"><span class="risk-pill">Severity:</span> {severity}</div></div>'
            f'<div class="risk-meta"><span class="risk-pill">Mitigation:</span> {row["Lead"]}</div>'
            f'</div>'
        )
    render_html('<div class="dash-card"><div class="section-bar">Key Risks Across Portfolio</div><div class="dashboard-stack">' + "".join(items) + "</div></div>")


def render_dashboard_decisions(df: pd.DataFrame) -> None:
    decisions = decision_table(df)
    if decisions.empty:
        body = '<div class="dash-note">No executive decisions are currently outstanding.</div>'
    else:
        rows = []
        for _, row in decisions.iterrows():
            rows.append(
                f'<tr><td>{row["Decision Needed"]}</td><td><div class="dash-program">{row["Program"]}</div></td><td>{row["Executive Sponsor"]}</td><td>{row["Priority"]}</td></tr>'
            )
        body = (
            '<div class="dash-table-wrap"><table class="dash-table"><thead><tr>'
            '<th>Decision Statement</th><th>Owning Program</th><th>Owner</th><th>Priority</th>'
            '</tr></thead><tbody>' + "".join(rows) + '</tbody></table></div>'
        )
    render_html('<div class="dash-card dash-bottom-space"><div class="section-bar">Executive Action Items / Decisions Needed</div>' + body + '</div>')


def render_dashboard_roadmap(df: pd.DataFrame) -> None:
    def esc(value: object) -> str:
        return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    roadmap_rows = []
    for idx, (_, row) in enumerate(df.sort_values(["Stage", "Program"]).iterrows()):
        segments = roadmap_stage_segments(str(row["Stage"]), int(row["Progress"]))
        milestone_label = esc(row["Milestone"])
        marker_pos = max(12, min(88, int(row["Progress"])))
        marker_class = "road-marker-alert" if row["Status"] == "At Risk" else ("road-marker-warn" if row["Status"] == "Needs Attention" else "road-marker-ok")
        label_class = "marker-alert" if row["Status"] == "At Risk" else ("marker-warn" if row["Status"] == "Needs Attention" else "marker-ok")
        segment_html = "".join(
            f'<div class="road-band-segment" style="width:{seg["width"] * 100:.1f}%; background:{seg["bg"]}; color:{seg["fg"]};">{seg["label"]}</div>'
            for seg in segments
        )
        marker_html = ""
        if idx < 4:
            marker_html = (
                f'<div class="road-marker {marker_class}" style="left:{marker_pos}%"></div>'
                f'<div class="road-marker-label {label_class}" style="left:{marker_pos}%">{milestone_label}</div>'
            )
        roadmap_rows.append(
            f'<div class="roadmap-journey">'
            f'<div class="roadmap-journey-name">{esc(row["Program"])}</div>'
            f'<div class="road-band-wrap"><div class="road-band">{segment_html}</div>{marker_html}</div>'
            f'<div class="road-progress">{int(row["Progress"])}% complete</div>'
            f'</div>'
        )

    render_html(
        """
        <div class="dash-card">
            <div class="dash-card-title">Portfolio Roadmap - FY25</div>
            <div class="dash-card-heading">Program Timeline</div>
            <div class="roadmap-quarter-row"><span>Q1 FY25</span><span>Q2 FY25</span><span class="active">Q3 FY25</span><span>Q4 FY25</span><span>Q1 FY26</span></div>
        """
        + f'<div class="roadmap-grid">{"".join(roadmap_rows)}</div>'
        + """
            <div class="road-legend">
                <span class="road-legend-item"><span class="road-legend-swatch" style="background:#b8c6fb;"></span>Discover</span>
                <span class="road-legend-item"><span class="road-legend-swatch" style="background:#dcc7fb;"></span>Plan</span>
                <span class="road-legend-item"><span class="road-legend-swatch" style="background:#294bb5;"></span>Execute</span>
                <span class="road-legend-item"><span class="road-legend-swatch" style="background:#b7efc7;"></span>Stabilize</span>
                <span class="road-legend-item"><span class="road-legend-swatch" style="background:#9fe9c4;"></span>Realize Value</span>
            </div>
        </div>
        """
    )


def render_all_programs(df: pd.DataFrame) -> None:
    st.markdown(
        '<div class="panel-header"><div class="eyebrow">Portfolio Directory</div><div class="heading">All Programs</div><div class="copy">Browse the full program portfolio with current owner, phase, health, and next milestone context.</div></div>',
        unsafe_allow_html=True,
    )
    overview = df[["Program", "Executive Sponsor", "Lead", "Stage", "Status", "Progress", "Milestone", "Milestone Date", "Open Risks"]].copy()
    overview["Milestone Date"] = overview["Milestone Date"].dt.strftime("%b %d, %Y")
    overview.rename(
        columns={
            "Executive Sponsor": "Sponsor",
            "Lead": "Program Lead",
            "Stage": "Phase",
            "Status": "Status",
            "Progress": "% Complete",
            "Milestone": "Next Milestone",
            "Milestone Date": "Milestone Date",
            "Open Risks": "Open Risks",
        },
        inplace=True,
    )
    st.dataframe(overview, use_container_width=True, hide_index=True)


def render_roadmap_milestones(df: pd.DataFrame) -> None:
    left, right = st.columns([1.45, 1.0], gap="large")
    with left:
        st.markdown('<div class="card"><div class="eyebrow">Roadmap View</div><div class="heading">Program Timeline</div></div>', unsafe_allow_html=True)
        st.altair_chart(timeline_chart(df), use_container_width=True)
    with right:
        st.markdown('<div class="card"><div class="eyebrow">Milestones</div><div class="heading">Upcoming Checkpoints</div></div>', unsafe_allow_html=True)
        milestone_df = milestone_table(df).copy()
        milestone_df["Milestone Date"] = milestone_df["Milestone Date"].dt.strftime("%b %d, %Y")
        st.dataframe(milestone_df, use_container_width=True, hide_index=True)


def render_risks_issues(df: pd.DataFrame) -> None:
    st.markdown('<div class="panel-header"><div class="eyebrow">Risk Oversight</div><div class="heading">Risks & Issues</div><div class="copy">Cross-program risk register with escalation pressure and active mitigation plans.</div></div>', unsafe_allow_html=True)
    risk_df = risk_table(df).copy()
    risk_df.rename(columns={"Risk Detail": "Description"}, inplace=True)
    st.dataframe(risk_df, use_container_width=True, hide_index=True)


def render_action_items(df: pd.DataFrame) -> None:
    st.markdown('<div class="panel-header"><div class="eyebrow">Executive Queue</div><div class="heading">Action Items</div><div class="copy">Decision requests and near-term actions requiring leadership visibility.</div></div>', unsafe_allow_html=True)
    decisions = decision_table(df).copy()
    if decisions.empty:
        st.info("No active action items are currently outstanding.")
    else:
        st.dataframe(decisions, use_container_width=True, hide_index=True)


def render_trend_analytics(df: pd.DataFrame) -> None:
    chart_df = df[["Program", "Progress", "Open Risks", "Escalations", "Status"]].copy()
    progress_chart = (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopLeft=8, cornerRadiusTopRight=8)
        .encode(
            x=alt.X("Program:N", sort="-y", title=None),
            y=alt.Y("Progress:Q", title="% Complete"),
            color=alt.Color(
                "Status:N",
                scale=alt.Scale(
                    domain=["On Track", "Needs Attention", "At Risk"],
                    range=[COLORS["green"], COLORS["yellow"], COLORS["red"]],
                ),
                legend=None,
            ),
            tooltip=["Program", "Progress", "Status"],
        )
        .properties(height=320)
        .configure_view(strokeWidth=0)
    )
    risk_bubble = (
        alt.Chart(chart_df)
        .mark_circle(opacity=0.85)
        .encode(
            x=alt.X("Open Risks:Q", title="Open Risks"),
            y=alt.Y("Escalations:Q", title="Escalations"),
            size=alt.Size("Progress:Q", scale=alt.Scale(range=[300, 1600]), legend=None),
            color=alt.Color(
                "Status:N",
                scale=alt.Scale(
                    domain=["On Track", "Needs Attention", "At Risk"],
                    range=[COLORS["green"], COLORS["yellow"], COLORS["red"]],
                ),
                legend=None,
            ),
            tooltip=["Program", "Open Risks", "Escalations", "Progress", "Status"],
        )
        .properties(height=320)
        .configure_view(strokeWidth=0)
    )
    left, right = st.columns(2, gap="large")
    with left:
        st.markdown('<div class="card"><div class="eyebrow">Delivery Trend</div><div class="heading">Completion by Program</div></div>', unsafe_allow_html=True)
        st.altair_chart(progress_chart, use_container_width=True)
    with right:
        st.markdown('<div class="card"><div class="eyebrow">Risk Pressure</div><div class="heading">Risk vs. Escalation Mix</div></div>', unsafe_allow_html=True)
        st.altair_chart(risk_bubble, use_container_width=True)


def render_settings() -> None:
    st.markdown('<div class="panel-header"><div class="eyebrow">Configuration</div><div class="heading">Settings</div><div class="copy">Control portfolio display defaults and reporting preferences.</div></div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.selectbox("Default Portfolio", list(PORTFOLIO_SEEDS), key="settings_default_portfolio")
        st.toggle("Enable autosave for weekly updates", value=True, key="settings_autosave")
    with c2:
        st.selectbox("Default landing page", PAGES[:7], key="settings_default_page")
        st.toggle("Show executive preview by default", value=True, key="settings_preview")


def render_help_support() -> None:
    st.markdown('<div class="panel-header"><div class="eyebrow">Support</div><div class="heading">Help & Support</div><div class="copy">Quick guidance for program leads and portfolio leadership.</div></div>', unsafe_allow_html=True)
    render_html(
        """
        <div class="card">
            <div class="eyebrow">Getting Started</div>
            <ul class="mini-list">
                <li>Use <strong>Weekly Updates</strong> to enter the current week’s program status, milestones, risks, and decision requests.</li>
                <li>Use <strong>Program One-Pager</strong> to review how a single program is presented for leadership consumption.</li>
                <li>Use <strong>Impower Portfolio</strong> for portfolio-wide oversight of milestones, risks, and action items.</li>
            </ul>
        </div>
        """
    )


def roadmap_stage_segments(stage: str, progress: int) -> list[dict[str, object]]:
    labels = ["Discover", "Plan", "Execute", "Stabilize", "Value"]
    widths = [0.15, 0.18, 0.35, 0.20, 0.12]
    active_map = {
        "Discovery": 0,
        "Phase 1": 1,
        "Phase 2": 2,
        "Phase 3": 3,
        "Phase 4": 4,
    }
    colors = ["#b8c6fb", "#dcc7fb", "#294bb5", "#b7efc7", "#9fe9c4"]
    text_colors = ["#4454d8", "#7b33d6", "#ffffff", "#17804c", "#107f67"]
    active_index = active_map.get(stage, 0)
    reached = min(4, max(active_index, round(progress / 25)))
    segments = []
    for idx, (label, width, color, text_color) in enumerate(zip(labels, widths, colors, text_colors)):
        state = "future" if idx > reached else ("active" if idx == active_index else "complete")
        bg = color if state != "future" else "#e8edf6"
        fg = text_color if state != "future" else "#9aa5b2"
        segments.append({"label": label, "width": width, "bg": bg, "fg": fg})
    return segments


def render_dashboard(portfolio: str, df: pd.DataFrame, reporting_date: date, refreshed_at: datetime) -> None:
    total_programs = len(df)
    on_track = int(df["Status"].eq("On Track").sum())
    at_risk = int(df["Status"].eq("At Risk").sum())
    delayed = int(df["Status"].eq("Needs Attention").sum())
    avg_complete = int(round(df["Progress"].mean()))
    decisions_pending = int((~df["Decision Needed"].str.contains("No ", na=False)).sum())
    render_html(
        f"""
        <div class="topbar">
            <div class="brand-block">
                <div class="brand-title">Portfolio Command Center</div>
                <div class="brand-copy">Strategic Program Intelligence</div>
            </div>
            <div class="search-shell">Search programs, milestones, risks...</div>
            <div class="toolbar-actions">
                <div class="toolbar-icon">◐</div>
                <div class="toolbar-icon">◔</div>
                <div class="toolbar-icon">3</div>
                <div class="toolbar-icon">⚙</div>
            </div>
            <div class="profile-shell">
                <div class="profile-copy">
                    <div class="profile-name">Marcus Ellison</div>
                    <div class="profile-role">VP, Transformation Office</div>
                </div>
                <div class="profile-avatar">ME</div>
            </div>
        </div>
        <div class="dashboard-title-block">
            <div class="dashboard-title">FY25 Strategic Transformation Portfolio</div>
            <div class="dashboard-meta"><span>Week Ending {reporting_date:%b %d, %Y}</span><span class="update-pill">Refreshed {refreshed_at:%I:%M %p}</span></div>
        </div>
        """,
    )
    kpi_cols = st.columns(5, gap="large")
    kpi_data = [
        ("Total Programs", total_programs),
        ("At Risk", at_risk),
        ("Delayed", delayed),
        ("Avg % Complete", f"{avg_complete}%"),
        ("Decisions Pending", decisions_pending),
    ]
    for col, (label, value) in zip(kpi_cols, kpi_data):
        col.markdown(
            f'<div class="dash-kpi"><div class="dash-kpi-name">{label}</div><div class="dash-kpi-value">{value}</div></div>',
            unsafe_allow_html=True,
        )

    left, right = st.columns([1.62, 0.92], gap="large")
    with left:
        render_dashboard_program_grid(df)
        render_dashboard_roadmap(df)
    with right:
        render_dashboard_milestones(df)
        render_dashboard_risks(df)

    render_dashboard_decisions(df)


def render_program_update(portfolio: str, program: str, df: pd.DataFrame, reporting_date: date) -> None:
    row = df.loc[df["Program"] == program].iloc[0]
    details = ensure_program_details(portfolio, program, reporting_date, row)
    current_status = st.session_state.get(f"status_{portfolio}_{program}", details["overall_status"])
    current_phase_value = st.session_state.get(f"phase_{portfolio}_{program}", details["current_phase"])
    current_progress = int(st.session_state.get(f"pct_{portfolio}_{program}", details["percent_complete"]))
    current_trend = st.session_state.get(f"trend_{portfolio}_{program}", details["trend"])

    st.markdown(
        """
        <style>
            .weekly-hero {
                background: #0f1837;
                color: white;
                border-radius: 22px;
                padding: 20px 24px 18px;
                box-shadow: 0 18px 42px rgba(8,16,40,0.18);
                margin-bottom: 1rem;
            }
            .weekly-hero-top {
                display: grid;
                grid-template-columns: minmax(0, 1.4fr) minmax(320px, 0.95fr);
                gap: 24px;
                align-items: start;
            }
            .weekly-hero-title {
                font-size: 1.9rem;
                font-weight: 800;
                letter-spacing: -0.03em;
                color: #ffffff;
            }
            .weekly-hero-sub {
                margin-top: 0.45rem;
                font-size: 1rem;
                color: #9bb2e8;
            }
            .weekly-hero-owners {
                display: flex;
                flex-wrap: wrap;
                gap: 1.6rem;
                margin-top: 1rem;
                color: #a8b6d7;
                font-size: 0.92rem;
            }
            .weekly-hero-owners strong {
                color: white;
                margin-left: 0.35rem;
            }
            .weekly-hero-metrics {
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
                gap: 1rem;
            }
            .weekly-hero-label {
                font-size: 0.74rem;
                text-transform: uppercase;
                letter-spacing: 0.08em;
                color: #9aa9cc;
                font-weight: 700;
            }
            .weekly-hero-value {
                margin-top: 0.45rem;
                font-size: 1.12rem;
                font-weight: 800;
                color: white;
                display: flex;
                align-items: center;
                gap: 0.55rem;
            }
            .weekly-status-dot {
                width: 14px;
                height: 14px;
                border-radius: 999px;
                display: inline-block;
            }
            .weekly-status-dot.green { background: #22c55e; }
            .weekly-status-dot.amber { background: #f59e0b; }
            .weekly-status-dot.red { background: #ef4444; }
            .weekly-hero-divider {
                height: 1px;
                background: rgba(255,255,255,0.16);
                margin: 1rem 0 0.85rem;
            }
            .weekly-hero-note {
                color: #d9e2f8;
                font-size: 0.92rem;
                line-height: 1.45;
            }
            .weekly-section-title {
                display: flex;
                align-items: center;
                gap: 0.55rem;
                font-size: 1rem;
                font-weight: 800;
                color: #1a2741;
                margin-bottom: 0.85rem;
            }
            .weekly-section-icon {
                color: #2f5bd2;
                font-size: 0.95rem;
            }
            .weekly-section-copy {
                color: #67788d;
                margin: -0.15rem 0 0.9rem;
                line-height: 1.42;
            }
            div[data-testid="stVerticalBlockBorderWrapper"] {
                border: 1px solid #d8e1ed !important;
                border-top: 1px solid #d8e1ed !important;
                border-radius: 22px !important;
                background: linear-gradient(180deg, #ffffff 0%, #fafcff 100%);
                box-shadow: 0 12px 28px rgba(20,42,71,0.04);
            }
            .stTextArea textarea,
            .stTextInput input,
            .stDateInput input,
            .stNumberInput input,
            [data-baseweb="select"] > div,
            [data-baseweb="input"] > div,
            .stDateInput [data-baseweb="input"] > div,
            .stMultiSelect [data-baseweb="select"] > div {
                background: #ffffff !important;
                border: 1px solid #d8e1ed !important;
                color: #23384f !important;
                border-radius: 12px !important;
                box-shadow: none !important;
            }
            .stTextArea textarea:focus,
            .stTextInput input:focus,
            .stDateInput input:focus,
            .stNumberInput input:focus {
                border-color: #8aa7e6 !important;
                box-shadow: 0 0 0 1px #8aa7e6 !important;
            }
            .stRadio [role="radiogroup"] label {
                background: #ffffff;
                border: 1px solid #d8e1ed;
                border-radius: 12px;
                padding: 0.52rem 0.75rem;
            }
            .stRadio [role="radiogroup"] label[data-checked="true"] {
                background: #f5f8ff;
                border-color: #9bb3ea;
                box-shadow: none;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    status_dot_class = {
        "On Track": "green",
        "Needs Attention": "amber",
        "At Risk": "red",
    }.get(str(current_status), "green")
    status_label = one_pager_status_label(str(current_status))

    render_html(
        f"""
        <div class="weekly-hero">
            <div class="weekly-hero-top">
                <div>
                    <div class="weekly-hero-title">{program}</div>
                    <div class="weekly-hero-sub">Weekly Update Entry for {portfolio}</div>
                    <div class="weekly-hero-owners">
                        <div>Executive Sponsor:<strong>{row["Executive Sponsor"]}</strong></div>
                        <div>Program Lead:<strong>{row["Lead"]}</strong></div>
                        <div>Reporting Week:<strong>{pd.to_datetime(details["week_ending"]):%b %d, %Y}</strong></div>
                    </div>
                </div>
                <div class="weekly-hero-metrics">
                    <div>
                        <div class="weekly-hero-label">Overall Status</div>
                        <div class="weekly-hero-value"><span class="weekly-status-dot {status_dot_class}"></span>{status_label}</div>
                    </div>
                    <div>
                        <div class="weekly-hero-label">% Complete</div>
                        <div class="weekly-hero-value">{current_progress}%</div>
                    </div>
                    <div>
                        <div class="weekly-hero-label">Current Phase</div>
                        <div class="weekly-hero-value">{current_phase_value}</div>
                    </div>
                </div>
            </div>
            <div class="weekly-hero-divider"></div>
            <div class="weekly-hero-note">Submitted data will update the Impower Portfolio and Program One-Pager automatically. Use concise, executive-ready language and capture only the information leaders need to act on.</div>
        </div>
        """
    )

    with st.container(border=True):
        render_html('<div class="weekly-section-title"><span class="weekly-section-icon">▣</span><span>Core Status Inputs</span></div><div class="weekly-section-copy">Capture the current program snapshot for this reporting cycle.</div>')
        status_cols = st.columns([0.95, 0.95, 0.95, 1.15], gap="large")
        with status_cols[0]:
            week_ending = st.date_input("Week Ending Date", value=pd.to_datetime(details["week_ending"]).date(), key=f"we_{portfolio}_{program}")
            overall_status = st.selectbox("Overall Status", ["On Track", "Needs Attention", "At Risk"], index=["On Track", "Needs Attention", "At Risk"].index(details["overall_status"]), key=f"status_{portfolio}_{program}")
        with status_cols[1]:
            current_phase = st.selectbox("Current Phase", ["Discovery", "Phase 1", "Phase 2", "Phase 3", "Phase 4"], index=["Discovery", "Phase 1", "Phase 2", "Phase 3", "Phase 4"].index(details["current_phase"]), key=f"phase_{portfolio}_{program}")
            percent_complete = st.slider("% Complete", min_value=0, max_value=100, value=int(details["percent_complete"]), key=f"pct_{portfolio}_{program}")
        with status_cols[2]:
            trend = st.radio("Trend vs Prior Week", ["Up", "Flat", "Down"], index=["Up", "Flat", "Down"].index(details["trend"]), key=f"trend_{portfolio}_{program}")
        with status_cols[3]:
            st.markdown('<div class="metric-title" style="margin-bottom:0.5rem;">Status Guidance</div>', unsafe_allow_html=True)
            render_html(
                """
                <div class="copy">Use concise executive language. Capture the delta from the prior week, the most important blockers, and any decision requests that leadership needs to act on.</div>
                """
            )

    with st.container(border=True):
        render_html('<div class="weekly-section-title"><span class="weekly-section-icon">☰</span><span>Milestone Updates</span></div>')
        milestones = render_milestone_editor(portfolio, program, details["milestones"])

    with st.container(border=True):
        render_html('<div class="weekly-section-title"><span class="weekly-section-icon">✳</span><span>Key Accomplishments This Week</span></div>')
        accomplishments = st.text_area("Key Accomplishments This Week", value=str(details["accomplishments"]), height=130, key=f"acc_{portfolio}_{program}", label_visibility="collapsed")

    with st.container(border=True):
        render_html('<div class="weekly-section-title"><span class="weekly-section-icon">⚠</span><span>Dependencies & Blockers</span></div>')
        dependencies = st.text_area("Dependencies & Blockers", value=str(details["dependencies"]), height=130, key=f"deps_{portfolio}_{program}", label_visibility="collapsed")

    narrative_bottom = st.columns(2, gap="large")
    with narrative_bottom[0]:
        with st.container(border=True):
            render_html('<div class="weekly-section-title"><span class="weekly-section-icon">▸</span><span>Planned Next Steps</span></div>')
            next_steps = st.text_area("Planned Next Steps", value=str(details["next_steps"]), height=130, key=f"next_{portfolio}_{program}", label_visibility="collapsed")
    with narrative_bottom[1]:
        with st.container(border=True):
            render_html('<div class="weekly-section-title"><span class="weekly-section-icon">▣</span><span>Executive Summary Notes</span></div>')
            executive_summary = st.text_area("Executive Summary Notes", value=str(details["executive_summary"]), height=130, key=f"sum_{portfolio}_{program}", label_visibility="collapsed")

    with st.container(border=True):
        render_html('<div class="weekly-section-title"><span class="weekly-section-icon">⚠</span><span>Risks & Mitigations</span></div>')
        risks = render_risk_editor(portfolio, program, details["risks"])

    with st.container(border=True):
        render_html('<div class="weekly-section-title"><span class="weekly-section-icon">✎</span><span>Leadership Decision Requests</span></div>')
        decisions = render_decision_editor(portfolio, program, details["decisions"])
        st.caption("Decision requests will appear as action items on the executive-facing pages.")

    preview_status = one_pager_status_label(overall_status)
    preview_status_cls = status_class(overall_status)
    top_risks = risks.head(3)
    open_decisions = decisions[~decisions["Decision Topic"].str.contains("No executive decision", na=False)].head(3)
    risk_lines = "".join(
        f"<li><strong>{row_item['Severity']}:</strong> {row_item['Title']} - {row_item['Mitigation Plan']}</li>"
        for _, row_item in top_risks.iterrows()
    ) or "<li>No active risks captured.</li>"
    decision_lines = "".join(
        f"<li>{row_item['Decision Topic']} - by {pd.to_datetime(row_item['Required By']):%b %d}</li>"
        for _, row_item in open_decisions.iterrows()
    ) or "<li>No active executive requests.</li>"

    preview_cols = st.columns([1.15, 0.85], gap="large")
    with preview_cols[0]:
        with st.container(border=True):
            render_html('<div class="weekly-section-title"><span class="weekly-section-icon">▣</span><span>Executive Preview</span></div>')
            render_html(
                f"""
                <div class="status-pill {preview_status_cls}" style="margin-top:0;">{preview_status}</div>
                <div class="copy" style="margin-top:0.85rem;"><strong>{program}</strong> is at <strong>{percent_complete}% complete</strong> with a <strong>{trend}</strong> trend for the week ending {week_ending:%b %d, %Y}.</div>
                <div class="eyebrow" style="margin-top:1rem;">Executive Summary</div>
                <div class="copy">{executive_summary}</div>
                <div class="eyebrow" style="margin-top:1rem;">Decision Requests</div>
                <ul class="mini-list">{decision_lines}</ul>
                """
            )
    with preview_cols[1]:
        required_missing = sum(
            [
                1 if not executive_summary.strip() else 0,
                1 if milestones.empty else 0,
                1 if not accomplishments.strip() else 0,
            ]
        )
        with st.container(border=True):
            render_html(
                f"""
                <div class="weekly-section-title"><span class="weekly-section-icon">✓</span><span>Submission Readiness</span></div>
                <div class="heading" style="font-size:1.18rem;">{required_missing} required fields need attention</div>
                <div class="copy">Save draft to persist current inputs. Submit update once the executive summary, milestones, risks, and decisions reflect what leadership should read this cycle.</div>
                <div class="eyebrow" style="margin-top:1rem;">Top Risks</div>
                <ul class="mini-list">{risk_lines}</ul>
                """
            )

    for frame, cols in [
        (milestones, ["Planned Date", "Forecast Date"]),
        (risks, ["Target Date"]),
        (decisions, ["Required By"]),
    ]:
        for col in cols:
            frame[col] = pd.to_datetime(frame[col], errors="coerce")

    actions = st.columns([1.0, 1.0, 0.9, 0.9], gap="large")
    updated_df = df.copy()
    row_idx = updated_df.index[updated_df["Program"] == program][0]
    top_milestone = milestones.sort_values("Forecast Date").iloc[0] if not milestones.empty else None
    top_risk = risks.sort_values("Target Date").iloc[0] if not risks.empty else None
    top_decision = open_decisions.iloc[0] if not open_decisions.empty else None
    updated_df.at[row_idx, "Stage"] = current_phase
    updated_df.at[row_idx, "Status"] = overall_status
    updated_df.at[row_idx, "Progress"] = int(percent_complete)
    updated_df.at[row_idx, "Milestone"] = top_milestone["Milestone Name"] if top_milestone is not None else row["Milestone"]
    updated_df.at[row_idx, "Milestone Date"] = top_milestone["Forecast Date"] if top_milestone is not None else row["Milestone Date"]
    updated_df.at[row_idx, "Summary"] = executive_summary
    updated_df.at[row_idx, "Upcoming Work"] = next_steps
    updated_df.at[row_idx, "Risk Detail"] = top_risk["Description"] if top_risk is not None else dependencies
    updated_df.at[row_idx, "Mitigation"] = top_risk["Mitigation Plan"] if top_risk is not None else row["Mitigation"]
    updated_df.at[row_idx, "Decision Needed"] = top_decision["Decision Topic"] if top_decision is not None else "No executive decision required this cycle."
    updated_df.at[row_idx, "Status Note"] = dependencies
    updated_df.at[row_idx, "Next Step"] = next_steps
    updated_df.at[row_idx, "Open Risks"] = int(len(risks))
    updated_df.at[row_idx, "Escalations"] = int((risks["Severity"] == "High").sum()) if not risks.empty else 0

    details_payload = {
        "week_ending": pd.Timestamp(week_ending),
        "current_phase": current_phase,
        "overall_status": overall_status,
        "percent_complete": int(percent_complete),
        "trend": trend,
        "accomplishments": accomplishments,
        "next_steps": next_steps,
        "dependencies": dependencies,
        "executive_summary": executive_summary,
        "milestones": milestones,
        "risks": risks,
        "decisions": decisions,
    }

    if actions[0].button("Save Draft", type="primary", use_container_width=True):
        save_program_details(portfolio, program, details_payload)
        save_state(portfolio, updated_df)
        st.success("Weekly update draft saved.")
        st.rerun()
    if actions[1].button("Submit Update", use_container_width=True):
        save_program_details(portfolio, program, details_payload)
        save_state(portfolio, updated_df)
        st.success("Weekly update submitted to the reporting views.")
        st.rerun()
    if actions[2].button("Reset Program", use_container_width=True):
        clear_milestone_editor_state(portfolio, program)
        save_program_details(portfolio, program, default_program_details(row, reporting_date))
        st.success("Program entry reset to the starter template.")
        st.rerun()
    if actions[3].button("Reset Portfolio", use_container_width=True):
        save_state(portfolio, load_portfolio_dataframe(portfolio))
        st.success("Portfolio data reset to the starter template.")
        st.rerun()


def render_executive_dashboard(program: str, df: pd.DataFrame) -> None:
    row = df.loc[df["Program"] == program].iloc[0]
    left, right = st.columns([1.55, 1.0], gap="large")
    with left:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="eyebrow">Executive Summary Notes</div><div class="heading">Executive Preview</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="copy">{row["Summary"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="copy" style="margin-top:0.8rem;"><strong>Next milestone:</strong> {row["Milestone"]} on {row["Milestone Date"]:%B %d, %Y}</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        lower = st.columns([1.0, 1.0], gap="large")
        with lower[0]:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="eyebrow">Risks & Mitigations</div><div class="heading">Top Portfolio Risk</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="copy"><strong>Risk:</strong> {row["Risk Detail"]}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="copy" style="margin-top:0.6rem;"><strong>Mitigation:</strong> {row["Mitigation"]}</div>', unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        with lower[1]:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="eyebrow">Leadership Decision Requests</div><div class="heading">Decision Needed</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="copy">{row["Decision Needed"]}</div>', unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="section-bar">Program Health</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="card"><div class="metric-title">Delivery Health</div><div class="metric-number" style="font-size:1.35rem;">{row["Delivery Health"]}</div><div class="metric-note">Overall execution confidence.</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="card" style="margin-top:1rem;"><div class="metric-title">Tech Health</div><div class="metric-number" style="font-size:1.35rem;">{row["Tech Health"]}</div><div class="metric-note">Architecture and engineering readiness.</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="card" style="margin-top:1rem;"><div class="metric-title">Team Health</div><div class="metric-number" style="font-size:1.35rem;">{row["Team Health"]}</div><div class="metric-note">Capacity and team stability.</div></div>', unsafe_allow_html=True)

    bottom_left, bottom_right = st.columns([1.2, 1.0], gap="large")
    with bottom_left:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="eyebrow">Portfolio Snapshot</div><div class="heading">Executive Health Mix</div>', unsafe_allow_html=True)
        st.altair_chart(status_rollup_chart(df), use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    with bottom_right:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="eyebrow">Executive Preview</div><div class="heading">What Leaders Should Know</div>', unsafe_allow_html=True)
        st.markdown(
            f"""
            <ul class="mini-list">
                <li>{row["Program"]} is currently <strong>{row["Status"]}</strong>.</li>
                <li>{int(row["Open Risks"])} open risks and {int(row["Escalations"])} escalations are reported.</li>
                <li>Primary next step: {row["Next Step"]}</li>
            </ul>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)


inject_styles()

query_params = st.query_params
requested_page = query_params.get("page")
requested_program = query_params.get("program")
pending_program = st.session_state.pop("pending_selected_program", None)
if pending_program in ALL_PROGRAMS:
    st.session_state["selected_program"] = pending_program
    st.session_state["sidebar_selected_program"] = pending_program
if requested_program in ALL_PROGRAMS:
    st.session_state["selected_program"] = requested_program
    st.session_state["sidebar_selected_program"] = requested_program
if requested_page in PAGES:
    st.session_state["current_page"] = requested_page

with st.sidebar:
    if "current_page" not in st.session_state or st.session_state["current_page"] not in PAGES:
        st.session_state["current_page"] = PAGES[0]
    if "selected_program" not in st.session_state or st.session_state["selected_program"] not in ALL_PROGRAMS:
        st.session_state["selected_program"] = ALL_PROGRAMS[0]
    st.markdown("### Program Context")
    if "sidebar_selected_program" not in st.session_state or st.session_state["sidebar_selected_program"] not in ALL_PROGRAMS:
        st.session_state["sidebar_selected_program"] = st.session_state["selected_program"]
    st.selectbox(
        "Program",
        ALL_PROGRAMS,
        key="sidebar_selected_program",
        on_change=sync_selected_program,
    )
    selected_program = st.session_state["selected_program"]
    portfolio = PROGRAM_TO_PORTFOLIO[selected_program]
    reporting_date = st.date_input("Reporting Date", date.today())
    refreshed_at = datetime.now(APP_TIMEZONE)
    portfolio_df = ensure_state(portfolio)
    page = render_app_navigation(st.session_state["current_page"])
    st.markdown("### Context")
    st.caption(f"Selected portfolio: {portfolio}. Program leads update weekly status and leadership consumes program-relevant views.")

render_header(page, portfolio, reporting_date)

if page == "Impower Portfolio":
    render_dashboard(portfolio, portfolio_df, reporting_date, refreshed_at)
elif page == "Program One-Pager":
    render_program_one_pager(portfolio, selected_program, portfolio_df, reporting_date)
elif page == "Weekly Updates":
    render_program_update(portfolio, selected_program, portfolio_df, reporting_date)
elif page == "Settings":
    render_settings()
elif page == "Help & Support":
    render_help_support()
