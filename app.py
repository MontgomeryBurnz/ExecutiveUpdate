from __future__ import annotations

from datetime import date, timedelta
from textwrap import dedent
from urllib.parse import quote

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
PAGES = [
    "Impower Portfolio",
    "Program One-Pager",
    "Weekly Updates",
    "All Programs",
    "Roadmap & Milestones",
    "Risks & Issues",
    "Action Items",
    "Trend Analytics",
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
                gap: 0.75rem;
            }}
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
                gap: 0.45rem;
            }}
            .risk-head {{
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                gap: 0.75rem;
            }}
            .risk-title {{
                font-size: 0.94rem;
                font-weight: 700;
                color: {COLORS["navy"]};
            }}
            .risk-meta {{
                font-size: 0.8rem;
                color: {COLORS["muted"]};
            }}
            .risk-trend {{
                font-size: 0.74rem;
                font-weight: 800;
            }}
            .trend-worse {{ color: {COLORS["red"]}; }}
            .trend-stable {{ color: {COLORS["muted"]}; }}
            .trend-better {{ color: {COLORS["green"]}; }}
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


def status_class(status: str) -> str:
    return {
        "On Track": "status-green",
        "Needs Attention": "status-yellow",
        "At Risk": "status-red",
    }.get(status, "status-green")


def render_html(html: str) -> None:
    st.markdown(dedent(html).strip(), unsafe_allow_html=True)


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
        "All Programs": "All Programs",
        "Roadmap & Milestones": "Roadmap & Milestones",
        "Risks & Issues": "Risks & Issues",
        "Action Items": "Action Items",
        "Trend Analytics": "Trend Analytics",
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
        "All Programs",
        "Roadmap & Milestones",
        "Risks & Issues",
        "Action Items",
        "Trend Analytics",
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
            f"""
            <tr>
                <td><span class="sev {sev_class}">{esc(item["severity"])}</span></td>
                <td>{esc(item["description"])}</td>
                <td>{esc(item["owner"])}</td>
                <td>{esc(item["mitigation"])}</td>
                <td>{esc(item["target"])}</td>
            </tr>
            """
        )

    decision_cards = []
    for idx, item in enumerate(decisions, start=1):
        decision_cards.append(
            f"""
            <div class="op-decision">
                <div class="op-decision-num">{idx}</div>
                <div>
                    <div class="op-decision-title">{esc(item["title"])}</div>
                    <div class="op-decision-meta">{esc(item["meta"])}</div>
                </div>
            </div>
            """
        )

    workstream_cards = []
    for item in workstreams:
        pct = int(item["pct"])
        workstream_cards.append(
            f"""
            <div class="op-workstream">
                <div class="op-work-top">
                    <div class="op-work-name">{esc(item["name"])}</div>
                    <div class="op-work-pct">{pct}%</div>
                </div>
                <div class="op-progress-bar"><span style="width:{pct}%"></span></div>
                <div class="op-work-note">{esc(item["note"])}</div>
                <div class="op-work-owner">Owner: {esc(item["owner"])}</div>
            </div>
            """
        )

    timeline_nodes = []
    for item in milestones:
        note_cls = "on-track" if item["note"] in {"On time", "On track", "Current checkpoint"} else "at-risk"
        timeline_nodes.append(
            f"""
            <div class="op-timeline-item">
                <div class="op-timeline-title">{esc(item["name"])}</div>
                <div class="op-timeline-date">{item["date"]:%d %b %Y}</div>
                <div class="op-timeline-note {note_cls}">{esc(item["note"])}</div>
            </div>
            """
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
            background: linear-gradient(180deg, #f6f9fd 0%, #eef3f9 100%);
            font-family: Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            color: #4f6278;
        }}
        .page {{ width: 100%; max-width: 1520px; margin: 0 auto; padding: 18px 18px 28px; }}
        .layout {{ display: block; }}
        .rail {{
            display: none;
        }}
        .rail-item {{
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 11px 12px;
            border-radius: 14px;
            color: #94a3b8;
            font-size: 14px;
            font-weight: 600;
            margin-bottom: 6px;
        }}
        .rail-item.active {{ background: linear-gradient(180deg, #376fe7 0%, #255ed4 100%); color: white; }}
        .rail-divider {{ height: 1px; background: #dee7f2; margin: 10px 0; }}
        .content {{ display: grid; gap: 16px; min-width: 0; }}
        .topbar {{
            display: grid;
            grid-template-columns: minmax(280px, 0.95fr) minmax(360px, 1.35fr) auto auto;
            gap: 16px;
            align-items: center;
        }}
        .brand-title {{ font-size: 32px; font-weight: 800; color: #304a68; line-height: 1.05; letter-spacing: -0.03em; }}
        .brand-sub {{ font-size: 13px; color: #8b9db2; margin-top: 6px; }}
        .search {{
            background: rgba(255,255,255,0.92);
            border: 1px solid #d9e3f0;
            border-radius: 12px;
            padding: 14px 18px;
            color: #9aaabc;
            font-size: 14px;
            min-width: 0;
        }}
        .icon-row {{ display: flex; gap: 8px; }}
        .icon {{
            width: 34px; height: 34px; border: 1px solid #d9e3f0; border-radius: 10px; background: rgba(255,255,255,0.92);
            display: flex; align-items: center; justify-content: center; color: #91a2b7; font-size: 13px; font-weight: 700;
        }}
        .profile {{ display: flex; align-items: center; gap: 10px; justify-content: flex-end; }}
        .profile-copy {{ text-align: right; }}
        .profile-name {{ font-size: 13px; font-weight: 800; color: #304a68; }}
        .profile-role {{ font-size: 11px; color: #8b9db2; }}
        .avatar {{
            width: 36px; height: 36px; border-radius: 999px; background: linear-gradient(180deg, #376fe7, #255ed4);
            color: white; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 800;
        }}
        .context-bar {{
            display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap;
            background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%);
            border: 1px solid #d9e3f0; border-radius: 18px; padding: 14px 18px;
        }}
        .context-left {{ display: flex; gap: 10px; align-items: center; flex-wrap: wrap; color: #7f91a6; font-size: 13px; font-weight: 700; }}
        .context-pill {{ padding: 6px 10px; border-radius: 999px; background: #edf4ff; color: #376fe7; }}
        .context-actions {{ display: flex; gap: 10px; flex-wrap: wrap; }}
        .ghost-btn, .primary-btn {{
            border-radius: 12px; padding: 12px 16px; font-size: 14px; font-weight: 700; white-space: nowrap;
        }}
        .ghost-btn {{ background: rgba(255,255,255,0.92); border: 1px solid #d9e3f0; color: #52667b; }}
        .primary-btn {{ background: linear-gradient(180deg, #376fe7 0%, #255ed4 100%); color: white; }}
        .program-hero {{
            display: flex; justify-content: space-between; gap: 18px; align-items: flex-start;
            background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
            border: 1px solid #d9e3f0; border-radius: 22px; padding: 22px;
        }}
        .program-title {{ font-size: clamp(30px, 3vw, 40px); font-weight: 800; color: #304a68; letter-spacing: -0.03em; }}
        .program-sub {{ margin-top: 8px; font-size: 16px; color: #6f8298; }}
        .owner-grid {{
            display: grid; grid-template-columns: repeat(3, minmax(170px, 1fr)); gap: 12px; margin-top: 16px;
        }}
        .owner-card {{
            background: #f7fbff; border: 1px solid #e2eaf5; border-radius: 16px; padding: 12px 14px;
        }}
        .owner-label {{ font-size: 11px; text-transform: uppercase; letter-spacing: 0.12em; color: #8ea0b5; font-weight: 800; }}
        .owner-value {{ margin-top: 6px; font-size: 15px; font-weight: 700; color: #304a68; }}
        .status-stack {{
            min-width: 320px; display: grid; gap: 12px;
        }}
        .status-grid {{ display: grid; grid-template-columns: repeat(3, minmax(120px, 1fr)); gap: 12px; }}
        .status-card {{
            background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%);
            border: 1px solid #d9e3f0; border-radius: 18px; padding: 16px;
        }}
        .status-label {{ font-size: 11px; text-transform: uppercase; letter-spacing: 0.12em; color: #93a3b6; font-weight: 800; }}
        .status-value {{ margin-top: 8px; font-size: 28px; font-weight: 800; color: #304a68; }}
        .op-green .status-value {{ color: #2e8b6f; }}
        .op-amber .status-value {{ color: #d3962f; }}
        .op-red .status-value {{ color: #cf5c5c; }}
        .status-sub {{ margin-top: 6px; font-size: 12px; color: #8093a9; }}
        .highlights {{
            display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px;
        }}
        .highlight {{
            background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%); border: 1px solid #d9e3f0; border-radius: 16px; padding: 12px 14px;
            font-size: 14px; color: #5f7288;
        }}
        .main-grid {{ display: grid; grid-template-columns: minmax(0, 1.52fr) minmax(360px, 0.96fr); gap: 16px; }}
        .stack {{ display: grid; gap: 16px; min-width: 0; }}
        .card {{
            background: linear-gradient(180deg, #ffffff 0%, #f9fbff 100%); border: 1px solid #d9e3f0; border-radius: 22px; padding: 18px;
            box-shadow: 0 12px 32px rgba(22,50,79,0.04);
        }}
        .card-eyebrow {{ font-size: 11px; text-transform: uppercase; letter-spacing: 0.15em; color: #376fe7; font-weight: 800; margin-bottom: 8px; }}
        .card-title {{ font-size: 18px; font-weight: 800; color: #304a68; }}
        .card-copy {{ margin-top: 8px; color: #6f8298; font-size: 15px; line-height: 1.6; }}
        .bullet-list {{ margin: 14px 0 0; padding-left: 20px; color: #5f7288; }}
        .bullet-list li {{ margin-bottom: 10px; line-height: 1.5; }}
        .timeline-grid {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; margin-top: 14px; }}
        .op-timeline-item {{ border: 1px solid #e3eaf4; border-radius: 16px; padding: 14px; background: #fcfdff; }}
        .op-timeline-title {{ font-size: 15px; font-weight: 700; color: #304a68; }}
        .op-timeline-date {{ margin-top: 6px; font-size: 13px; color: #8193a9; }}
        .op-timeline-note {{ margin-top: 8px; font-size: 12px; font-weight: 800; }}
        .on-track {{ color: #2e8b6f; }}
        .at-risk {{ color: #cf5c5c; }}
        .table-scroll {{ width: 100%; overflow-x: auto; overflow-y: hidden; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 14px; table-layout: fixed; }}
        th {{
            text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.12em; color: #94a3b8;
            padding: 12px 10px; border-bottom: 1px solid #e6edf5;
        }}
        td {{ padding: 14px 10px; border-bottom: 1px solid #edf2f8; font-size: 14px; color: #516478; vertical-align: top; }}
        tr:last-child td {{ border-bottom: none; }}
        .sev {{ display: inline-block; border-radius: 999px; padding: 5px 9px; font-size: 11px; font-weight: 800; }}
        .sev-high {{ background: rgba(207,92,92,0.13); color: #cf5c5c; }}
        .sev-med {{ background: rgba(211,162,60,0.14); color: #d3a23c; }}
        .sev-low {{ background: rgba(46,139,111,0.12); color: #2e8b6f; }}
        .sev-dep {{ background: rgba(55,111,231,0.12); color: #376fe7; }}
        .op-decision {{
            display: grid; grid-template-columns: 34px 1fr; gap: 12px; align-items: start;
            padding: 14px 0; border-bottom: 1px solid #edf2f8;
        }}
        .op-decision:last-child {{ border-bottom: none; padding-bottom: 0; }}
        .op-decision-num {{
            width: 34px; height: 34px; border-radius: 999px; background: linear-gradient(180deg, #376fe7 0%, #255ed4 100%);
            color: white; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 14px;
        }}
        .op-decision-title {{ font-size: 15px; font-weight: 700; color: #304a68; line-height: 1.45; }}
        .op-decision-meta {{ margin-top: 5px; font-size: 12px; color: #8497ae; }}
        .workstream-grid {{ display: grid; gap: 12px; margin-top: 14px; }}
        .op-workstream {{ border: 1px solid #e1e9f3; border-radius: 16px; padding: 14px; background: #fcfdff; }}
        .op-work-top {{ display: flex; justify-content: space-between; gap: 12px; align-items: center; }}
        .op-work-name {{ font-size: 15px; font-weight: 700; color: #304a68; }}
        .op-work-pct {{ font-size: 15px; font-weight: 800; color: #376fe7; }}
        .op-progress-bar {{ height: 10px; border-radius: 999px; background: #edf2f7; margin-top: 10px; overflow: hidden; }}
        .op-progress-bar span {{ display: block; height: 100%; border-radius: 999px; background: linear-gradient(90deg, #5c8ef0 0%, #376fe7 100%); }}
        .op-work-note {{ margin-top: 10px; font-size: 13px; color: #6f8298; line-height: 1.45; }}
        .op-work-owner {{ margin-top: 8px; font-size: 12px; color: #8ea0b5; font-weight: 700; }}
        .change-grid {{ display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 12px; margin-top: 14px; }}
        .change-card {{ border: 1px solid #e1e9f3; border-radius: 16px; padding: 14px; background: #fcfdff; }}
        .change-label {{ font-size: 11px; text-transform: uppercase; letter-spacing: 0.12em; color: #8ea0b5; font-weight: 800; }}
        .change-value {{ margin-top: 8px; font-size: 16px; font-weight: 800; color: #304a68; }}
        .change-note {{ margin-top: 8px; font-size: 13px; color: #6f8298; line-height: 1.45; }}
        @media (max-width: 1320px) {{
            .topbar {{ grid-template-columns: minmax(240px, 0.95fr) minmax(260px, 1fr) auto; }}
            .profile {{ grid-column: 1 / -1; justify-content: flex-start; }}
            .main-grid {{ grid-template-columns: 1fr; }}
            .change-grid {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
        }}
        @media (max-width: 1080px) {{
            .topbar, .program-hero {{ grid-template-columns: 1fr; display: grid; }}
            .owner-grid, .status-grid, .highlights, .timeline-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
            .context-bar {{ align-items: flex-start; }}
        }}
        @media (max-width: 760px) {{
            .page {{ padding: 12px 10px 20px; }}
            .owner-grid, .status-grid, .highlights, .timeline-grid, .change-grid {{ grid-template-columns: 1fr; }}
            table, thead, tbody, th, td, tr {{ display: block; }}
            thead {{ display: none; }}
            tbody tr {{ padding: 10px 0; border-bottom: 1px solid #edf2f8; }}
            td {{ border-bottom: none; padding: 6px 0; }}
        }}
    </style>
    </head>
    <body>
        <div class="page">
            <div class="layout">
                <div class="rail">
                    <div class="rail-item">Impower Portfolio</div>
                    <div class="rail-item active">Program One-Pager</div>
                    <div class="rail-item">Weekly Update Entry</div>
                    <div class="rail-divider"></div>
                    <div class="rail-item">All Programs</div>
                    <div class="rail-item">Roadmap &amp; Milestones</div>
                    <div class="rail-item">Risks &amp; Issues</div>
                    <div class="rail-item">Action Items</div>
                    <div class="rail-item">Trend Analytics</div>
                    <div class="rail-item">Settings</div>
                    <div class="rail-item">Help &amp; Support</div>
                </div>
                <div class="content">
                    <div class="topbar">
                        <div>
                            <div class="brand-title">Portfolio Command Center</div>
                            <div class="brand-sub">Strategic Program Intelligence</div>
                        </div>
                        <div class="search">Search programs, milestones, risks...</div>
                        <div class="icon-row"><div class="icon">◐</div><div class="icon">◔</div><div class="icon">3</div><div class="icon">⚙</div></div>
                        <div class="profile">
                            <div class="profile-copy"><div class="profile-name">Marcus Ellison</div><div class="profile-role">VP, Transformation Office</div></div>
                            <div class="avatar">ME</div>
                        </div>
                    </div>
                    <div class="context-bar">
                        <div class="context-left">
                            <span class="context-pill">Reporting Period: {period}</span>
                            <span>|</span>
                            <span>Cycle {cycle_num} of {total_cycles}</span>
                        </div>
                        <div class="context-actions">
                            <div class="ghost-btn">Back to Impower Portfolio</div>
                            <div class="primary-btn">Edit Weekly Update</div>
                        </div>
                    </div>
                    <div class="program-hero">
                        <div>
                            <div class="program-title">{esc(program)}</div>
                            <div class="program-sub">{esc(portfolio)} transformation program overview</div>
                            <div class="owner-grid">
                                <div class="owner-card"><div class="owner-label">Executive Sponsor</div><div class="owner-value">{esc(sponsor)}</div></div>
                                <div class="owner-card"><div class="owner-label">Program Lead</div><div class="owner-value">{esc(lead)}</div></div>
                                <div class="owner-card"><div class="owner-label">PMO</div><div class="owner-value">{esc(pmo)}</div></div>
                            </div>
                        </div>
                        <div class="status-stack">
                            <div class="status-grid">
                                <div class="status-card {status_cls}"><div class="status-label">Overall Status</div><div class="status-value">{status_label}</div></div>
                                <div class="status-card"><div class="status-label">% Complete</div><div class="status-value">{progress}%</div><div class="status-sub">{delta}</div></div>
                                <div class="status-card"><div class="status-label">Current Phase</div><div class="status-value" style="font-size:22px;">{esc(row["Stage"])}</div></div>
                            </div>
                            <div class="highlights">
                                <div class="highlight">{esc(highlights[0])}</div>
                                <div class="highlight">{esc(highlights[1])}</div>
                                <div class="highlight">{esc(highlights[2])}</div>
                            </div>
                        </div>
                    </div>
                    <div class="main-grid">
                        <div class="stack">
                            <div class="card">
                                <div class="card-eyebrow">Executive Summary</div>
                                <div class="card-title">Program Narrative</div>
                                <div class="card-copy">{esc(row["Summary"])}</div>
                            </div>
                            <div class="card">
                                <div class="card-eyebrow">Recent Accomplishments</div>
                                <div class="card-title">What Landed This Cycle</div>
                                <ul class="bullet-list">
                                    <li>{esc(accomplishments[0])}</li>
                                    <li>{esc(accomplishments[1])}</li>
                                    <li>{esc(accomplishments[2])}</li>
                                    <li>{esc(accomplishments[3])}</li>
                                </ul>
                            </div>
                            <div class="card">
                                <div class="card-eyebrow">Upcoming Work (Next 2 Weeks)</div>
                                <div class="card-title">Near-Term Delivery Focus</div>
                                <ul class="bullet-list">
                                    {''.join(f'<li>{esc(item)}</li>' for item in upcoming_work)}
                                </ul>
                            </div>
                            <div class="card">
                                <div class="card-eyebrow">Workstream Status</div>
                                <div class="card-title">Execution by Workstream</div>
                                <div class="workstream-grid">
                                    {''.join(workstream_cards)}
                                </div>
                            </div>
                            <div class="card">
                                <div class="card-eyebrow">Change From Previous Reporting Cycle</div>
                                <div class="card-title">Cycle-over-Cycle View</div>
                                <div class="change-grid">
                                    <div class="change-card"><div class="change-label">Status Movement</div><div class="change-value">{status_label} → {status_label}</div><div class="change-note">{"No change this cycle" if row["Status"] != "At Risk" else "Escalated pressure remains active"}</div></div>
                                    <div class="change-card"><div class="change-label">Milestone Changes</div><div class="change-value">{esc(row["Milestone"])}</div><div class="change-note">{"Held on plan" if row["Status"] == "On Track" else "Monitoring schedule pressure"}</div></div>
                                    <div class="change-card"><div class="change-label">New Risks Opened</div><div class="change-value">{min(2, int(row["Open Risks"]))} new</div><div class="change-note">Total open risks: {int(row["Open Risks"])}</div></div>
                                    <div class="change-card"><div class="change-label">Closed Issues</div><div class="change-value">{max(1, 3 - int(row["Escalations"]))} closed</div><div class="change-note">Dependency follow-ups and execution blockers addressed.</div></div>
                                    <div class="change-card"><div class="change-label">Confidence Level</div><div class="change-value">{confidence}</div><div class="change-note">{confidence_delta}</div></div>
                                </div>
                            </div>
                        </div>
                        <div class="stack">
                            <div class="card">
                                <div class="card-eyebrow">Milestone Timeline</div>
                                <div class="card-title">Key Delivery Checkpoints</div>
                                <div class="timeline-grid">
                                    {''.join(timeline_nodes)}
                                </div>
                            </div>
                            <div class="card">
                                <div class="card-eyebrow">Risks / Issues / Dependencies</div>
                                <div class="card-title">Active Risk Register</div>
                                <div class="table-scroll">
                                    <table>
                                        <thead>
                                            <tr><th>Severity</th><th>Description</th><th>Owner</th><th>Mitigation</th><th>Target Date</th></tr>
                                        </thead>
                                        <tbody>
                                            {''.join(risk_rows)}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                            <div class="card">
                                <div class="card-eyebrow">Leadership Decisions Needed</div>
                                <div class="card-title">Executive Requests</div>
                                {''.join(decision_cards)}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html, height=2850, scrolling=True)


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
    rows = []
    for _, row in df.sort_values(["Milestone Date", "Program"]).iterrows():
        milestone_date = row["Milestone Date"].strftime("%b %d, %Y")
        rows.append(
            f'<tr>'
            f'<td><div class="dash-program">{row["Program"]}</div><div class="dash-note">{row["Lead"]}</div></td>'
            f'<td>{row["Milestone"]}<div class="dash-note">{milestone_date}</div></td>'
            f'<td>{dashboard_status_tag(row["Status"])}</td>'
            f'<td class="dash-note">{row["Status Note"]}</td>'
            f'</tr>'
        )
    html = (
        '<div class="dash-card">'
        '<div class="dash-card-title">Portfolio Program Grid</div>'
        '<div class="dash-card-heading">Weekly Updates</div>'
        '<div class="dash-card-copy">Program-level status, next milestone, and current delivery note aligned to the weekly portfolio rhythm.</div>'
        '<div class="dash-table-wrap"><table class="dash-table"><thead><tr>'
        '<th>Program</th><th>Next Milestone</th><th>Status</th><th>Status Note</th>'
        '</tr></thead><tbody>'
        + "".join(rows)
        + '</tbody></table></div></div>'
    )
    render_html(html)


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
    ranked = risk_table(df)
    trend_classes = ["trend-worse", "trend-stable", "trend-better", "trend-stable", "trend-worse"]
    trend_labels = ["Worse", "Stable", "Better", "Stable", "Worse"]
    for index, (_, row) in enumerate(ranked.iterrows()):
        trend_class = trend_classes[index % len(trend_classes)]
        trend_label = trend_labels[index % len(trend_labels)]
        severity = "High" if int(row["Escalations"]) > 1 or int(row["Open Risks"]) >= 5 else ("Medium" if int(row["Open Risks"]) >= 3 else "Low")
        items.append(
            f'<div class="risk-item">'
            f'<div class="risk-head"><div><div class="risk-title">{row["Program"]}</div><div class="risk-meta">Severity: {severity}</div></div><div class="risk-trend {trend_class}">{trend_label}</div></div>'
            f'<div class="risk-meta">Mitigation: {row["Mitigation"]}</div>'
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
    stage_order = ["Discovery", "Phase 1", "Phase 2", "Phase 3", "Phase 4"]
    class_map = {
        "Discovery": "discovery",
        "Phase 1": "phase1",
        "Phase 2": "phase2",
        "Phase 3": "phase3",
        "Phase 4": "phase4",
    }
    rows = []
    for _, row in df.sort_values(["Stage", "Program"]).iterrows():
        pills = []
        for stage in stage_order:
            active = row["Stage"] == stage
            pill_class = class_map[stage] if active else ""
            label = stage.replace("Phase ", "") if stage != "Discovery" else "Disc"
            pills.append(f'<div class="roadmap-pill {pill_class}">{label}</div>')
        rows.append(
            f'<div class="roadmap-row"><div class="roadmap-name">{row["Program"]}</div>{"".join(pills)}<div class="roadmap-progress">{int(row["Progress"])}%</div></div>'
        )
    render_html(
        '<div class="dash-card" style="margin-top:1rem;">'
        '<div class="dash-card-title">Portfolio Roadmap - FY25</div>'
        '<div class="dash-card-heading">Program Timeline</div>'
        '<div class="dash-card-copy">A phase-based roadmap aligned to the visual structure in the executive render.</div>'
        '<div class="roadmap-grid">' + "".join(rows) + "</div></div>"
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


def render_dashboard(portfolio: str, df: pd.DataFrame) -> None:
    def esc(value: object) -> str:
        return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    total_programs = len(df)
    on_track = int(df["Status"].eq("On Track").sum())
    at_risk = int(df["Status"].eq("At Risk").sum())
    delayed = int(df["Status"].eq("Needs Attention").sum())
    avg_complete = int(round(df["Progress"].mean()))
    decisions_pending = int((~df["Decision Needed"].str.contains("No ", na=False)).sum())

    grid_rows = []
    for _, row in df.sort_values(["Milestone Date", "Program"]).iterrows():
        trend = "↑" if row["Status"] == "On Track" else ("!" if row["Status"] == "Needs Attention" else "↓")
        rag_class = "rag-green" if row["Status"] == "On Track" else ("rag-yellow" if row["Status"] == "Needs Attention" else "rag-red")
        page_param = quote("Program One-Pager")
        program_param = quote(str(row["Program"]))
        grid_rows.append(
            f"""
            <tr>
                <td class="program-cell">
                    <div class="program-name">
                        <a
                            href="#"
                            onclick="window.top.location.href = window.top.location.origin + window.top.location.pathname + '?page={page_param}&program={program_param}'; return false;"
                        >{esc(row["Program"])}</a>
                    </div>
                </td>
                <td>{esc(row["Lead"])}</td>
                <td>{esc(row["Stage"])}</td>
                <td><span class="rag-dot {rag_class}"></span></td>
                <td>{int(row["Progress"])}%</td>
                <td>{trend}</td>
                <td>{esc(row["Milestone"])}</td>
                <td class="status-note">{esc(row["Status Note"])}</td>
            </tr>
            """
        )

    roadmap_rows = []
    for idx, (_, row) in enumerate(df.sort_values(["Stage", "Program"]).iterrows()):
        segments = roadmap_stage_segments(str(row["Stage"]), int(row["Progress"]))
        milestone_label = esc(row["Milestone"])
        marker_pos = max(12, min(88, int(row["Progress"])))
        marker_class = "road-marker-alert" if row["Status"] == "At Risk" else ("road-marker-warn" if row["Status"] == "Needs Attention" else "road-marker-ok")
        segment_html = "".join(
            f'<div class="road-band-segment" style="width:{seg["width"] * 100:.1f}%; background:{seg["bg"]}; color:{seg["fg"]};">{seg["label"]}</div>'
            for seg in segments
        )
        marker_html = ""
        if idx < 4:
            marker_html = (
                f'<div class="road-marker {marker_class}" style="left:{marker_pos}%"></div>'
                f'<div class="road-marker-label" style="left:{marker_pos}%">{milestone_label}</div>'
            )
        roadmap_rows.append(
            f"""
            <div class="roadmap-journey">
                <div class="roadmap-journey-name">{esc(row["Program"])}</div>
                <div class="road-band-wrap">
                    <div class="road-band">{segment_html}</div>
                    {marker_html}
                </div>
                <div class="roadmap-progress">{int(row["Progress"])}% complete</div>
            </div>
            """
        )

    milestone_cards = []
    for _, row in milestone_table(df).iterrows():
        dt = row["Milestone Date"]
        status_cls = "prio-high" if row["Status"] == "On Track" else ("prio-medium" if row["Status"] == "Needs Attention" else "prio-low")
        status_lbl = "High" if row["Status"] == "On Track" else ("Medium" if row["Status"] == "Needs Attention" else "Low")
        milestone_cards.append(
            f"""
            <div class="mile-card">
                <div class="datebox"><div class="day">{dt.strftime("%d")}</div><div class="mon">{dt.strftime("%b").upper()}</div></div>
                <div class="mile-copy"><div class="mile-title">{esc(row["Milestone"])}</div><div class="mile-sub">{esc(row["Program"])}</div></div>
                <div class="prio {status_cls}">{status_lbl}</div>
            </div>
            """
        )

    risk_cards = []
    trends = ["Worse", "Stable", "Better", "Stable", "Worse"]
    trend_classes = ["trend-red", "trend-muted", "trend-green", "trend-muted", "trend-red"]
    for idx, (_, row) in enumerate(risk_table(df).iterrows()):
        severity = "High" if int(row["Escalations"]) > 1 or int(row["Open Risks"]) >= 5 else ("Medium" if int(row["Open Risks"]) >= 3 else "Low")
        risk_cards.append(
            f"""
            <div class="risk-card">
                <div class="risk-top">
                    <div class="risk-title">{esc(row["Program"])}</div>
                    <div class="risk-trend {trend_classes[idx % len(trend_classes)]}">{trends[idx % len(trends)]}</div>
                </div>
                <div class="risk-meta">Severity: {severity}</div>
                <div class="risk-meta">Mitigation: {esc(row["Mitigation"])}</div>
            </div>
            """
        )

    decision_rows = []
    decisions = decision_table(df)
    for _, row in decisions.iterrows():
        impact = "High" if row["Priority"] == "Critical" else ("Medium" if row["Priority"] == "High" else "Low")
        decision_rows.append(
            f"""
            <tr>
                <td>{esc(row["Decision Needed"])}</td>
                <td>{esc(row["Program"])}</td>
                <td>{impact}</td>
                <td>{esc(row["Executive Sponsor"])}</td>
            </tr>
            """
        )

    html = f"""
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <style>
        html, body {{
            width: 100%;
            min-height: 100%;
            overflow-x: hidden;
        }}
        * {{
            box-sizing: border-box;
        }}
        body {{
            margin: 0;
            background: linear-gradient(180deg, #f6f9fd 0%, #eef3f9 100%);
            font-family: Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            color: #4f6278;
        }}
        .page {{
            width: 100%;
            max-width: 1520px;
            margin: 0 auto;
            padding: 18px 18px 28px;
        }}
        .layout {{
            display: block;
        }}
        .rail {{
            display: none;
        }}
        .rail-item {{
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 11px 12px;
            border-radius: 14px;
            color: #94a3b8;
            font-size: 14px;
            font-weight: 600;
            margin-bottom: 6px;
        }}
        .rail-item.active {{
            background: linear-gradient(180deg, #376fe7 0%, #255ed4 100%);
            color: white;
        }}
        .rail-divider {{
            height: 1px;
            background: #dee7f2;
            margin: 10px 0;
        }}
        .content {{
            display: grid;
            gap: 16px;
            min-width: 0;
        }}
        .topbar {{
            display: grid;
            grid-template-columns: minmax(280px, 0.9fr) minmax(360px, 1.35fr) auto auto;
            gap: 16px;
            align-items: center;
        }}
        .brand-title {{
            font-size: 32px;
            font-weight: 800;
            color: #304a68;
            line-height: 1.05;
            letter-spacing: -0.03em;
        }}
        .brand-sub {{
            font-size: 13px;
            color: #8b9db2;
            margin-top: 6px;
        }}
        .search {{
            background: rgba(255,255,255,0.92);
            border: 1px solid #d9e3f0;
            border-radius: 12px;
            padding: 14px 18px;
            color: #9aaabc;
            font-size: 14px;
            min-width: 0;
        }}
        .icon-row {{
            display: flex;
            gap: 8px;
        }}
        .icon {{
            width: 34px;
            height: 34px;
            border: 1px solid #d9e3f0;
            border-radius: 10px;
            background: rgba(255,255,255,0.92);
            display: flex;
            align-items: center;
            justify-content: center;
            color: #91a2b7;
            font-size: 13px;
            font-weight: 700;
        }}
        .profile {{
            display: flex;
            align-items: center;
            gap: 10px;
            justify-content: flex-end;
        }}
        .profile-copy {{
            text-align: right;
        }}
        .profile-name {{
            font-size: 13px;
            font-weight: 800;
            color: #304a68;
        }}
        .profile-role {{
            font-size: 11px;
            color: #8b9db2;
        }}
        .avatar {{
            width: 36px;
            height: 36px;
            border-radius: 999px;
            background: linear-gradient(180deg, #376fe7, #255ed4);
            color: white;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
            font-weight: 800;
        }}
        .hero {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 18px;
        }}
        .hero-title {{
            font-size: clamp(30px, 3vw, 38px);
            font-weight: 800;
            color: #304a68;
            letter-spacing: -0.02em;
        }}
        .hero-meta {{
            display: flex;
            gap: 10px;
            align-items: center;
            margin-top: 8px;
            color: #7f91a6;
            font-size: 13px;
        }}
        .updated {{
            background: rgba(46,139,111,0.12);
            color: #2e8b6f;
            border-radius: 999px;
            padding: 4px 8px;
            font-weight: 700;
        }}
        .actions {{
            display: flex;
            gap: 10px;
            align-items: center;
            flex-wrap: wrap;
            justify-content: flex-end;
        }}
        .ghost-btn, .primary-btn {{
            border-radius: 12px;
            padding: 12px 16px;
            font-size: 14px;
            font-weight: 700;
            white-space: nowrap;
        }}
        .ghost-btn {{
            background: rgba(255,255,255,0.92);
            border: 1px solid #d9e3f0;
            color: #52667b;
        }}
        .primary-btn {{
            background: linear-gradient(180deg, #376fe7 0%, #255ed4 100%);
            color: white;
        }}
        .kpis {{
            display: grid;
            grid-template-columns: repeat(5, minmax(140px, 1fr));
            gap: 12px;
        }}
        .kpi {{
            background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%);
            border: 1px solid #d9e3f0;
            border-radius: 18px;
            padding: 16px;
        }}
        .kpi-label {{
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.12em;
            color: #93a3b6;
            font-weight: 800;
        }}
        .kpi-value {{
            margin-top: 10px;
            font-size: 40px;
            line-height: 1;
            font-weight: 800;
            color: #304a68;
        }}
        .kpi-delta {{
            margin-top: 8px;
            font-size: 12px;
            color: #8ea0b5;
        }}
        .main-grid {{
            display: grid;
            grid-template-columns: minmax(0, 1.58fr) minmax(320px, 0.94fr);
            gap: 16px;
        }}
        .stack {{
            display: grid;
            gap: 16px;
            min-width: 0;
        }}
        .card {{
            background: linear-gradient(180deg, #ffffff 0%, #f9fbff 100%);
            border: 1px solid #d9e3f0;
            border-radius: 22px;
            padding: 18px;
            box-shadow: 0 12px 32px rgba(22,50,79,0.04);
        }}
        .card-eyebrow {{
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.15em;
            color: #376fe7;
            font-weight: 800;
            margin-bottom: 8px;
        }}
        .card-title {{
            font-size: 18px;
            font-weight: 800;
            color: #304a68;
        }}
        .card-copy {{
            margin-top: 6px;
            color: #8093a9;
            font-size: 14px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 14px;
            table-layout: fixed;
        }}
        th {{
            text-align: left;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.12em;
            color: #94a3b8;
            padding: 12px 10px;
            border-bottom: 1px solid #e6edf5;
        }}
        td {{
            padding: 14px 10px;
            border-bottom: 1px solid #edf2f8;
            font-size: 14px;
            color: #516478;
            vertical-align: top;
        }}
        tr:last-child td {{ border-bottom: none; }}
        .program-name {{
            color: #305da8;
            font-weight: 800;
        }}
        .program-name a {{
            color: #305da8;
            text-decoration: none;
            border-bottom: 1px solid transparent;
        }}
        .program-name a:hover {{
            color: #1e4fa6;
            border-bottom-color: #1e4fa6;
        }}
        .status-note {{
            color: #6f8298;
        }}
        .rag-dot {{
            display: inline-block;
            width: 10px;
            height: 10px;
            border-radius: 999px;
        }}
        .rag-green {{ background: #2e8b6f; }}
        .rag-yellow {{ background: #d3a23c; }}
        .rag-red {{ background: #cf5c5c; }}
        .section-pill {{
            background: linear-gradient(180deg, #6c95ea 0%, #4f80e6 100%);
            color: white;
            border-radius: 16px;
            text-align: center;
            padding: 16px 14px;
            font-size: 16px;
            font-weight: 800;
            margin-bottom: 14px;
        }}
        .roadmap-quarter-row {{
            display: flex;
            gap: 30px;
            flex-wrap: wrap;
            margin: 8px 0 18px;
            font-size: 22px;
            color: #a3a7ad;
        }}
        .roadmap-quarter-row .active {{
            color: #294bb5;
            font-weight: 800;
        }}
        .roadmap-grid {{
            display: grid;
            gap: 24px;
            margin-top: 12px;
        }}
        .roadmap-journey-name {{
            font-size: 18px;
            font-weight: 800;
            color: #3f4a58;
            margin-bottom: 8px;
        }}
        .road-band-wrap {{
            position: relative;
            padding-bottom: 32px;
        }}
        .road-band {{
            display: flex;
            align-items: stretch;
            width: 100%;
            height: 42px;
            border-radius: 999px;
            background: #f0f2f5;
            overflow: hidden;
        }}
        .road-band-segment {{
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            font-weight: 700;
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
            top: -8px;
            width: 18px;
            height: 56px;
            transform: translateX(-50%);
            border-radius: 2px;
        }}
        .road-marker-ok {{ background: #294bb5; }}
        .road-marker-warn {{ background: #ff9800; }}
        .road-marker-alert {{ background: #f44336; }}
        .road-marker-label {{
            position: absolute;
            top: 58px;
            transform: translateX(-50%);
            font-size: 12px;
            font-weight: 800;
            white-space: nowrap;
            color: #d97800;
        }}
        .road-progress {{
            margin-top: 2px;
            text-align: right;
            color: #8092a8;
            font-size: 12px;
            font-weight: 800;
        }}
        .mile-list, .risk-list {{
            display: grid;
            gap: 12px;
        }}
        .mile-card, .risk-card {{
            background: linear-gradient(180deg, #ffffff 0%, #f9fbff 100%);
            border: 1px solid #e1e9f3;
            border-radius: 18px;
            padding: 14px;
        }}
        .mile-card {{
            display: grid;
            grid-template-columns: 52px 1fr auto;
            gap: 12px;
            align-items: center;
        }}
        .datebox {{
            width: 52px;
            height: 52px;
            border-radius: 14px;
            border: 1px solid #dce5f1;
            background: #f2f6fb;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }}
        .day {{
            font-size: 22px;
            font-weight: 800;
            color: #304a68;
            line-height: 1;
        }}
        .mon {{
            margin-top: 4px;
            font-size: 10px;
            color: #8ea0b5;
            font-weight: 800;
            letter-spacing: 0.12em;
        }}
        .mile-title {{
            font-size: 15px;
            font-weight: 800;
            color: #304a68;
        }}
        .mile-sub {{
            margin-top: 4px;
            font-size: 12px;
            color: #8ea0b5;
        }}
        .prio {{
            border-radius: 999px;
            padding: 6px 10px;
            font-size: 12px;
            font-weight: 800;
        }}
        .prio-high {{ color: #2e8b6f; background: rgba(46,139,111,0.12); }}
        .prio-medium {{ color: #d3a23c; background: rgba(211,162,60,0.14); }}
        .prio-low {{ color: #cf5c5c; background: rgba(207,92,92,0.13); }}
        .risk-top {{
            display: flex;
            justify-content: space-between;
            gap: 12px;
            align-items: start;
        }}
        .risk-title {{
            font-size: 15px;
            font-weight: 800;
            color: #304a68;
        }}
        .risk-meta {{
            margin-top: 6px;
            font-size: 12px;
            color: #8697ac;
        }}
        .trend-red {{ color: #cf5c5c; font-size: 12px; font-weight: 800; }}
        .trend-green {{ color: #2e8b6f; font-size: 12px; font-weight: 800; }}
        .trend-muted {{ color: #8ea0b5; font-size: 12px; font-weight: 800; }}
        .decision-table th, .decision-table td {{
            font-size: 13px;
        }}
        .table-scroll {{
            width: 100%;
            overflow-x: auto;
            overflow-y: hidden;
        }}
        @media (max-width: 1280px) {{
            .topbar {{
                grid-template-columns: minmax(240px, 0.95fr) minmax(260px, 1fr) auto;
            }}
            .profile {{
                grid-column: 1 / -1;
            }}
            .kpis {{
                grid-template-columns: repeat(3, minmax(0, 1fr));
            }}
            .main-grid {{
                grid-template-columns: 1fr;
            }}
        }}
        @media (max-width: 1080px) {{
            .page {{
                padding: 16px 14px 24px;
            }}
            .topbar {{
                grid-template-columns: 1fr;
            }}
            .profile {{
                justify-content: flex-start;
            }}
            .hero {{
                flex-direction: column;
            }}
            .actions {{
                justify-content: flex-start;
            }}
            .kpis {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
        }}
        @media (max-width: 760px) {{
            .page {{
                padding: 12px 10px 20px;
            }}
            .brand-title {{
                font-size: 24px;
            }}
            .kpis {{
                grid-template-columns: 1fr;
            }}
            .mile-card {{
                grid-template-columns: 52px 1fr;
            }}
            .prio {{
                grid-column: 1 / -1;
                justify-self: start;
            }}
            .road-row {{
                grid-template-columns: 1fr;
            }}
            .road-progress {{
                text-align: left;
            }}
            table,
            thead,
            tbody,
            th,
            td,
            tr {{
                display: block;
            }}
            thead {{
                display: none;
            }}
            tbody tr {{
                padding: 10px 0;
                border-bottom: 1px solid #edf2f8;
            }}
            td {{
                border-bottom: none;
                padding: 6px 0;
            }}
        }}
    </style>
    </head>
    <body>
        <div class="page">
            <div class="layout">
                <div class="rail">
                    <div class="rail-item active">Impower Portfolio</div>
                    <div class="rail-item">Program One-Pager</div>
                    <div class="rail-item">Weekly Update Entry</div>
                    <div class="rail-divider"></div>
                    <div class="rail-item">All Programs</div>
                    <div class="rail-item">Roadmap &amp; Milestones</div>
                    <div class="rail-item">Risks &amp; Issues</div>
                    <div class="rail-item">Action Items</div>
                    <div class="rail-item">Trend Analytics</div>
                    <div class="rail-item">Settings</div>
                    <div class="rail-item">Help &amp; Support</div>
                </div>
                <div class="content">
                    <div class="topbar">
                        <div>
                            <div class="brand-title">Portfolio Command Center</div>
                            <div class="brand-sub">Strategic Program Intelligence</div>
                        </div>
                        <div class="search">Search programs, milestones, risks...</div>
                        <div class="icon-row"><div class="icon">◐</div><div class="icon">◔</div><div class="icon">3</div><div class="icon">⚙</div></div>
                        <div class="profile">
                            <div class="profile-copy"><div class="profile-name">Marcus Ellison</div><div class="profile-role">VP, Transformation Office</div></div>
                            <div class="avatar">ME</div>
                        </div>
                    </div>
                    <div class="hero">
                        <div>
                            <div class="hero-title">{PORTFOLIO_NAME}</div>
                            <div class="hero-meta"><span>Week Ending Sep 20, 2026</span><span class="updated">Updated 12 min ago</span></div>
                        </div>
                        <div class="actions"><div class="ghost-btn">Export PDF</div><div class="primary-btn">Share Report</div></div>
                    </div>
                    <div class="kpis">
                        <div class="kpi"><div class="kpi-label">Total Programs</div><div class="kpi-value">{total_programs}</div></div>
                        <div class="kpi"><div class="kpi-label">At Risk</div><div class="kpi-value">{at_risk}</div></div>
                        <div class="kpi"><div class="kpi-label">Delayed</div><div class="kpi-value">{delayed}</div></div>
                        <div class="kpi"><div class="kpi-label">Avg % Complete</div><div class="kpi-value">{avg_complete}%</div></div>
                        <div class="kpi"><div class="kpi-label">Decisions Pending</div><div class="kpi-value">{decisions_pending}</div></div>
                    </div>
                    <div class="main-grid">
                        <div class="stack">
                            <div class="card">
                                <div class="card-eyebrow">Portfolio Program Grid</div>
                                <div class="card-title">Weekly Updates</div>
                                <div class="card-copy">Program, owner, trend, progress, next milestone, and current status note.</div>
                                <div class="table-scroll">
                                <table>
                                    <thead>
                                        <tr><th>Program</th><th>Owner</th><th>Phase</th><th>RAG</th><th>% Done</th><th>Trend</th><th>Next Milestone</th><th>Status Note</th></tr>
                                    </thead>
                                    <tbody>
                                        {''.join(grid_rows)}
                                    </tbody>
                                </table>
                                </div>
                            </div>
                            <div class="card">
                                <div class="card-eyebrow">Portfolio Roadmap - FY25</div>
                                <div class="card-title">Program Timeline</div>
                                <div class="roadmap-quarter-row"><span>Q1 FY25</span><span>Q2 FY25</span><span class="active">Q3 FY25</span><span>Q4 FY25</span><span>Q1 FY26</span></div>
                                <div class="roadmap-grid">{''.join(roadmap_rows)}</div>
                            </div>
                        </div>
                        <div class="stack">
                            <div class="card">
                                <div class="section-pill">Upcoming Milestones</div>
                                <div class="mile-list">{''.join(milestone_cards)}</div>
                            </div>
                            <div class="card">
                                <div class="section-pill">Key Risks Across Portfolio</div>
                                <div class="risk-list">{''.join(risk_cards)}</div>
                            </div>
                        </div>
                    </div>
                    <div class="card">
                        <div class="section-pill">Executive Action Items / Decisions Needed</div>
                        <div class="table-scroll">
                        <table class="decision-table">
                            <thead>
                                <tr><th>Decision Statement</th><th>Owning Program</th><th>Impact if Delayed</th><th>Owner</th></tr>
                            </thead>
                            <tbody>
                                {''.join(decision_rows) if decision_rows else '<tr><td colspan="4">No executive decisions are currently outstanding.</td></tr>'}
                            </tbody>
                        </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html, height=2200, scrolling=True)


def render_program_update(portfolio: str, program: str, df: pd.DataFrame, reporting_date: date) -> None:
    row = df.loc[df["Program"] == program].iloc[0]
    details = ensure_program_details(portfolio, program, reporting_date, row)

    render_html(
        f"""
        <div class="panel-header">
            <div class="eyebrow">Weekly Update</div>
            <div class="heading">Draft <span style="color:{COLORS["muted"]}; font-size:0.95rem; font-weight:700;">· Auto-saved 2 min ago</span></div>
            <div class="copy"><strong>{program}</strong> for <strong>{portfolio}</strong>. Submitted data will update the portfolio dashboard and executive one-pager automatically.</div>
        </div>
        """
    )

    hero_cols = st.columns([1.15, 1.0, 0.85, 0.85], gap="large")
    hero_cols[0].markdown(f'<div class="metric-card"><div class="metric-title">Program Owner</div><div class="metric-number" style="font-size:1.18rem;">{row["Executive Sponsor"]}</div></div>', unsafe_allow_html=True)
    hero_cols[1].markdown(f'<div class="metric-card"><div class="metric-title">Reporting Cycle</div><div class="metric-number" style="font-size:1.18rem;">#24</div><div class="metric-note">{cycle_label(reporting_date)}</div></div>', unsafe_allow_html=True)
    hero_cols[2].markdown(f'<div class="metric-card"><div class="metric-title">Week Ending</div><div class="metric-number" style="font-size:1.18rem;">{pd.to_datetime(details["week_ending"]):%b %d, %Y}</div></div>', unsafe_allow_html=True)
    hero_cols[3].markdown('<div class="metric-card"><div class="metric-title">Submission Guidance</div><div class="metric-note" style="margin-top:0.55rem;">Use concise, report-ready language. This page is the source for the portfolio and executive views.</div></div>', unsafe_allow_html=True)

    main_left, main_right = st.columns([1.55, 0.95], gap="large")

    with main_left:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="eyebrow">Core Status Inputs</div><div class="heading">Weekly Program Snapshot</div>', unsafe_allow_html=True)
        status_cols = st.columns(2, gap="large")
        with status_cols[0]:
            week_ending = st.date_input("Week Ending Date", value=pd.to_datetime(details["week_ending"]).date(), key=f"we_{portfolio}_{program}")
            current_phase = st.selectbox("Current Phase", ["Discovery", "Phase 1", "Phase 2", "Phase 3", "Phase 4"], index=["Discovery", "Phase 1", "Phase 2", "Phase 3", "Phase 4"].index(details["current_phase"]), key=f"phase_{portfolio}_{program}")
            overall_status = st.selectbox("Overall Status", ["On Track", "Needs Attention", "At Risk"], index=["On Track", "Needs Attention", "At Risk"].index(details["overall_status"]), key=f"status_{portfolio}_{program}")
        with status_cols[1]:
            percent_complete = st.slider("% Complete", min_value=0, max_value=100, value=int(details["percent_complete"]), key=f"pct_{portfolio}_{program}")
            trend = st.radio("Trend vs Prior Week", ["Up", "Flat", "Down"], index=["Up", "Flat", "Down"].index(details["trend"]), horizontal=True, key=f"trend_{portfolio}_{program}")
            st.caption("Directional movement since the prior reporting cycle.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="card" style="margin-top:1rem;">', unsafe_allow_html=True)
        st.markdown('<div class="eyebrow">Milestone Updates</div><div class="heading">Milestone Tracker</div>', unsafe_allow_html=True)
        milestones = render_milestone_editor(portfolio, program, details["milestones"])
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="card" style="margin-top:1rem;">', unsafe_allow_html=True)
        st.markdown('<div class="eyebrow">Narrative Inputs</div><div class="heading">What Leadership Will Read</div>', unsafe_allow_html=True)
        accomplishments = st.text_area("Key Accomplishments This Week", value=str(details["accomplishments"]), height=110, key=f"acc_{portfolio}_{program}", help="Concise, report-ready language preferred. This appears in the executive summary.")
        next_steps = st.text_area("Planned Next Steps", value=str(details["next_steps"]), height=110, key=f"next_{portfolio}_{program}", help="Focus on actions planned for the next 1-2 weeks.")
        dependencies = st.text_area("Dependencies & Blockers", value=str(details["dependencies"]), height=110, key=f"deps_{portfolio}_{program}", help="Highlight cross-program and external dependencies.")
        executive_summary = st.text_area("Executive Summary Notes", value=str(details["executive_summary"]), height=130, key=f"sum_{portfolio}_{program}", help="2-3 sentences for the executive one-pager headline.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="card" style="margin-top:1rem;">', unsafe_allow_html=True)
        st.markdown('<div class="eyebrow">Risks & Mitigations</div><div class="heading">Risk Register</div>', unsafe_allow_html=True)
        risks = st.data_editor(
            details["risks"],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            height=260,
            key=f"risks_{portfolio}_{program}",
            column_config={
                "Severity": st.column_config.SelectboxColumn("Severity", options=["High", "Medium", "Low", "DEP"]),
                "Title": st.column_config.TextColumn("Title"),
                "Owner": st.column_config.TextColumn("Owner"),
                "Target Date": st.column_config.DateColumn("Target Date"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Mitigation Plan": st.column_config.TextColumn("Mitigation Plan", width="large"),
            },
        )
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="card" style="margin-top:1rem;">', unsafe_allow_html=True)
        st.markdown('<div class="eyebrow">Leadership Decision Requests</div><div class="heading">Actionable Executive Requests</div>', unsafe_allow_html=True)
        decisions = st.data_editor(
            details["decisions"],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            height=210,
            key=f"decisions_{portfolio}_{program}",
            column_config={
                "Decision Topic": st.column_config.TextColumn("Decision Topic", width="large"),
                "Required By": st.column_config.DateColumn("Required By"),
                "Impact if Unresolved": st.column_config.TextColumn("Impact if Unresolved", width="large"),
                "Recommendation": st.column_config.TextColumn("Recommendation", width="large"),
            },
        )
        st.caption("Decision requests appear as action items on the executive reporting dashboard.")
        st.markdown("</div>", unsafe_allow_html=True)

    with main_right:
        st.markdown('<div class="section-bar">Executive Preview</div>', unsafe_allow_html=True)
        preview_status = one_pager_status_label(overall_status)
        preview_status_cls = status_class(overall_status)
        top_risks = risks.head(2)
        open_decisions = decisions[~decisions["Decision Topic"].str.contains("No executive decision", na=False)].head(2)
        risk_lines = "".join(
            f"<li><strong>{row_item['Severity']}:</strong> {row_item['Title']}</li>"
            for _, row_item in top_risks.iterrows()
        ) or "<li>No active risks captured.</li>"
        decision_lines = "".join(
            f"<li>{row_item['Decision Topic']} - by {pd.to_datetime(row_item['Required By']):%b %d}</li>"
            for _, row_item in open_decisions.iterrows()
        ) or "<li>No active executive requests.</li>"
        render_html(
            f"""
            <div class="card">
                <div class="eyebrow">How Your Update Will Appear In Reporting</div>
                <div class="heading">{program}</div>
                <div class="status-pill {preview_status_cls}">{preview_status}</div>
                <div class="copy"><strong>{percent_complete}% Complete</strong> · {trend} trend · Week ending {week_ending:%b %d, %Y}</div>
                <div class="eyebrow" style="margin-top:1rem;">Executive Summary</div>
                <div class="copy">{executive_summary}</div>
                <div class="eyebrow" style="margin-top:1rem;">Key Risks</div>
                <ul class="mini-list">{risk_lines}</ul>
                <div class="eyebrow" style="margin-top:1rem;">Decisions Needed</div>
                <ul class="mini-list">{decision_lines}</ul>
            </div>
            """
        )
        required_missing = sum(
            [
                1 if not executive_summary.strip() else 0,
                1 if milestones.empty else 0,
            ]
        )
        render_html(
            f"""
            <div class="card" style="margin-top:1rem;">
                <div class="eyebrow">Submission Readiness</div>
                <div class="heading">{required_missing} required fields need attention</div>
                <div class="copy">Save draft to persist current inputs. Submit update when the preview reflects the message you want executives to read.</div>
            </div>
            """
        )

    for frame, cols in [
        (milestones, ["Planned Date", "Forecast Date"]),
        (risks, ["Target Date"]),
        (decisions, ["Required By"]),
    ]:
        for col in cols:
            frame[col] = pd.to_datetime(frame[col], errors="coerce")

    actions = st.columns([0.95, 0.95, 1.15, 1.55], gap="large")
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
    portfolio_df = ensure_state(portfolio)
    page = render_app_navigation(st.session_state["current_page"])
    st.markdown("### Context")
    st.caption(f"Selected portfolio: {portfolio}. Program leads update weekly status and leadership consumes program-relevant views.")

render_header(page, portfolio, reporting_date)

if page == "Impower Portfolio":
    render_dashboard(portfolio, portfolio_df)
elif page == "Program One-Pager":
    render_program_one_pager(portfolio, selected_program, portfolio_df, reporting_date)
elif page == "Weekly Updates":
    render_program_update(portfolio, selected_program, portfolio_df, reporting_date)
elif page == "All Programs":
    render_all_programs(portfolio_df)
elif page == "Roadmap & Milestones":
    render_roadmap_milestones(portfolio_df)
elif page == "Risks & Issues":
    render_risks_issues(portfolio_df)
elif page == "Action Items":
    render_action_items(portfolio_df)
elif page == "Trend Analytics":
    render_trend_analytics(portfolio_df)
elif page == "Settings":
    render_settings()
elif page == "Help & Support":
    render_help_support()
