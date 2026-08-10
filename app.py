import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date
import uuid

st.set_page_config(page_title="Snowflake POC Tracker", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── Header definitions ────────────────────────────────────────────────────────
REGISTRY_HEADERS = ["POC_ID", "Customer", "Engagement", "AE", "SE", "Status", "Created"]

OVERVIEW_HEADERS = [
    "POC_ID",
    "Snowflake AE", "Snowflake SE",
    "Labcorp Champion", "Champion Title", "Champion Email",
    "Executive Sponsor", "Sponsor Title",
    "Business Unit", "Primary Use Case", "Secondary Use Case",
    "Cloud Environment", "Current Data Platform", "Data Volume",
    "Compliance Requirements",
    "POC Start Date", "Target Completion Date", "Status",
    "POC Objective", "Technical Success Criteria", "Business Success Criteria",
    "POC Budget ($)", "Confirmed Spend ($)", "Potential ARR ($)",
    "Procurement Contact", "Budget Notes",
]

KPI_HEADERS    = ["POC_ID", "KPI", "Target", "Current Value", "Unit", "Status", "Notes"]
ACTION_HEADERS = ["POC_ID", "Action", "Owner", "Category", "Due Date", "Status", "Notes"]
UPDATE_HEADERS = ["POC_ID", "Date", "Update", "Posted By"]

USE_CASES = [
    "Real-World Evidence (RWE)",
    "Clinical Data Analytics",
    "Genomics & Next-Gen Sequencing",
    "Data Sharing with Pharma Partners",
    "Lab Results Platform Modernisation",
    "Clinical Trial Data Management",
    "Regulatory & Compliance Reporting",
    "Other",
]
BUSINESS_UNITS     = ["Diagnostics", "Drug Development (Biopharma Solutions)", "Genomics", "Technology Solutions", "Enterprise / Cross-BU"]
COMPLIANCE_OPTIONS = ["HIPAA", "21 CFR Part 11", "GxP", "SOC 2", "CLIA", "GDPR"]
STATUS_OPTIONS     = ["Planning", "Active", "At Risk", "Completed — Won", "Completed — Lost"]
KPI_STATUSES       = ["On Track", "At Risk", "Met", "Not Started"]
ACTION_OWNERS      = ["Snowflake", "Customer", "Joint"]
ACTION_CATS        = ["Technical", "Business", "Compliance", "Training", "Executive"]
ACTION_STATUSES    = ["Open", "In Progress", "Complete", "Blocked"]
CLOUD_OPTIONS      = ["AWS", "Azure", "Multi-Cloud"]

STATUS_EMOJI = {
    "Planning": "🔵", "Active": "🟢", "At Risk": "🟡",
    "Completed — Won": "✅", "Completed — Lost": "🔴",
}

# ── Google Sheets helpers ─────────────────────────────────────────────────────
@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    return gspread.authorize(creds)

def get_ss():
    return get_client().open_by_key(st.secrets["spreadsheet_id"])

def ensure_sheets():
    ss = get_ss()
    existing = {ws.title for ws in ss.worksheets()}
    for title, headers in [
        ("POC_Registry",  REGISTRY_HEADERS),
        ("Overview",      OVERVIEW_HEADERS),
        ("KPIs",          KPI_HEADERS),
        ("Action_Items",  ACTION_HEADERS),
        ("Updates",       UPDATE_HEADERS),
    ]:
        if title not in existing:
            ws = ss.add_worksheet(title=title, rows=1000, cols=max(len(headers), 10))
            ws.append_row(headers)
        else:
            ws = ss.worksheet(title)
            if not ws.row_values(1):
                ws.insert_row(headers, 1)

# ── Load functions ────────────────────────────────────────────────────────────
@st.cache_data(ttl=20)
def load_registry() -> pd.DataFrame:
    ws = get_ss().worksheet("POC_Registry")
    data = ws.get_all_records()
    if data:
        return pd.DataFrame(data)
    return pd.DataFrame(columns=REGISTRY_HEADERS)

@st.cache_data(ttl=20)
def load_overview(poc_id: str) -> dict:
    ws = get_ss().worksheet("Overview")
    rows = ws.get_all_records(expected_headers=OVERVIEW_HEADERS)
    for r in rows:
        if r.get("POC_ID") == poc_id:
            return r
    return {}

@st.cache_data(ttl=20)
def load_kpis(poc_id: str) -> pd.DataFrame:
    ws = get_ss().worksheet("KPIs")
    data = ws.get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=KPI_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        df = df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

@st.cache_data(ttl=20)
def load_actions(poc_id: str) -> pd.DataFrame:
    ws = get_ss().worksheet("Action_Items")
    data = ws.get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=ACTION_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        df = df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

@st.cache_data(ttl=20)
def load_updates(poc_id: str) -> pd.DataFrame:
    ws = get_ss().worksheet("Updates")
    data = ws.get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=UPDATE_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        df = df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

# ── Save / mutate functions ───────────────────────────────────────────────────
def save_overview(poc_id: str, values: dict):
    ws = get_ss().worksheet("Overview")
    all_rows = ws.get_all_values()
    row = [poc_id] + [values.get(h, "") for h in OVERVIEW_HEADERS[1:]]
    # find existing row for this poc_id
    for i, r in enumerate(all_rows[1:], start=2):
        if r and r[0] == poc_id:
            ws.delete_rows(i)
            ws.insert_row(row, i)
            st.cache_data.clear()
            return
    ws.append_row(row)
    st.cache_data.clear()

def append_to(sheet: str, poc_id: str, row: list):
    get_ss().worksheet(sheet).append_row([poc_id] + row, value_input_option="USER_ENTERED")
    st.cache_data.clear()

def update_sheet_row(sheet: str, poc_id: str, sheet_row: int, headers: list, values: dict):
    ws = get_ss().worksheet(sheet)
    full_row = [poc_id] + [values.get(h, "") for h in headers]
    # sheet_row is 1-indexed; row 1 = header
    for i, val in enumerate(full_row, start=1):
        ws.update_cell(sheet_row, i, val)
    st.cache_data.clear()

def delete_sheet_row(sheet: str, sheet_row: int):
    get_ss().worksheet(sheet).delete_rows(sheet_row)
    st.cache_data.clear()

def create_poc(poc_id: str, customer: str, engagement: str, ae: str, se: str, status: str):
    ws = get_ss().worksheet("POC_Registry")
    ws.append_row([poc_id, customer, engagement, ae, se, status, datetime.today().strftime("%Y-%m-%d")])
    st.cache_data.clear()

def get_all_sheet_row(sheet: str, poc_id: str) -> list[tuple[int, dict]]:
    """Return (sheet_row_number, record) pairs for a given poc_id."""
    ws = get_ss().worksheet(sheet)
    all_vals = ws.get_all_values()
    if not all_vals:
        return []
    headers = all_vals[0]
    result = []
    for i, row in enumerate(all_vals[1:], start=2):
        if row and row[0] == poc_id:
            record = {headers[j]: row[j] for j in range(len(headers)) if j < len(row)}
            result.append((i, record))
    return result

# ── Bootstrap ─────────────────────────────────────────────────────────────────
try:
    ensure_sheets()
except Exception as e:
    st.error(f"Could not connect to Google Sheets: {e}")
    st.stop()

# ── Premium CSS injection ─────────────────────────────────────────────────────
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
html, body, [class*="css"], .stApp { font-family: 'Inter', sans-serif !important; }

/* Hide Streamlit chrome */
#MainMenu, footer, header { visibility: hidden; height: 0; }
.block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: #0B1929 !important;
    border-right: 1px solid #1E3A5F !important;
}
section[data-testid="stSidebar"] .stRadio label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] small,
section[data-testid="stSidebar"] div { color: #CBD5E0 !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #FFFFFF !important; }
section[data-testid="stSidebar"] hr { border-color: #1E3A5F !important; }
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
    background: rgba(255,255,255,.04) !important;
    border: 1px solid #1E3A5F !important;
    border-radius: 8px !important;
    padding: 10px 14px !important;
    margin-bottom: 6px !important;
    transition: background .2s !important;
    font-size: .85rem !important;
}
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {
    background: rgba(41,181,232,.1) !important;
    border-color: #29B5E8 !important;
}
section[data-testid="stSidebar"] input[type="text"],
section[data-testid="stSidebar"] select,
section[data-testid="stSidebar"] textarea {
    background: rgba(255,255,255,.06) !important;
    border: 1px solid #1E3A5F !important;
    color: #FFFFFF !important;
    border-radius: 6px !important;
}
section[data-testid="stSidebar"] button[kind="primaryFormSubmit"],
section[data-testid="stSidebar"] button[kind="primary"] {
    background: #29B5E8 !important; color: #000 !important; border: none !important;
    border-radius: 6px !important; font-weight: 600 !important;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    border-bottom: 2px solid #E2E8F0;
}
.stTabs [data-baseweb="tab"] {
    font-size: .82rem !important; font-weight: 500 !important;
    letter-spacing: .04em !important;
    padding: 10px 20px !important;
    border-radius: 8px 8px 0 0 !important;
    color: #718096 !important;
}
.stTabs [aria-selected="true"] {
    background: #EBF8FF !important;
    color: #0B6E9B !important;
    border-bottom: 2px solid #29B5E8 !important;
}

/* ── Primary buttons ── */
.stButton > button[kind="primary"] {
    background: #29B5E8 !important; color: #000 !important;
    border: none !important; border-radius: 6px !important;
    font-weight: 600 !important; font-size: .82rem !important;
    letter-spacing: .04em !important;
    padding: 8px 24px !important;
}
.stButton > button {
    border-radius: 6px !important; font-size: .82rem !important;
    font-weight: 500 !important;
}

/* ── Inputs ── */
.stTextInput input, .stTextArea textarea, .stSelectbox select {
    border-radius: 6px !important;
    border: 1px solid #CBD5E0 !important;
    font-size: .88rem !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: #29B5E8 !important;
    box-shadow: 0 0 0 3px rgba(41,181,232,.15) !important;
}

/* ── Expanders ── */
.streamlit-expanderHeader {
    font-size: .85rem !important; font-weight: 600 !important;
    color: #2D3748 !important;
    background: #F7FAFC !important;
    border-radius: 8px !important;
    border: 1px solid #E2E8F0 !important;
}

/* ── Dividers ── */
hr { border-color: #E2E8F0 !important; }

/* ── Custom KPI card ── */
.kpi-grid { display: flex; flex-wrap: wrap; gap: 14px; margin-bottom: 24px; }
.kpi-card {
    flex: 1 1 200px; min-width: 180px;
    background: #fff;
    border: 1px solid #E2E8F0;
    border-top: 3px solid #CBD5E0;
    border-radius: 10px;
    padding: 18px 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.kpi-card.met   { border-top-color: #38A169; }
.kpi-card.risk  { border-top-color: #DD6B20; }
.kpi-card.track { border-top-color: #3182CE; }
.kpi-card.none  { border-top-color: #A0AEC0; }
.kpi-name  { font-size: .78rem; font-weight: 600; color: #4A5568; letter-spacing: .05em; text-transform: uppercase; margin-bottom: 8px; }
.kpi-value { font-size: 1.9rem; font-weight: 700; color: #1A202C; line-height: 1; margin-bottom: 4px; }
.kpi-target{ font-size: .76rem; color: #718096; margin-bottom: 10px; }
.kpi-badge { display: inline-block; font-size: .7rem; font-weight: 600; padding: 3px 10px; border-radius: 20px; letter-spacing: .06em; }
.badge-met   { background: #C6F6D5; color: #22543D; }
.badge-track { background: #BEE3F8; color: #1A365D; }
.badge-risk  { background: #FEEBC8; color: #7B341E; }
.badge-none  { background: #EDF2F7; color: #4A5568; }
.badge-blocked { background: #FED7D7; color: #742A2A; }
.badge-won   { background: #C6F6D5; color: #22543D; }
.badge-lost  { background: #FED7D7; color: #742A2A; }
.badge-active   { background: #BEE3F8; color: #1A365D; }
.badge-planning { background: #EDF2F7; color: #4A5568; }
.badge-atrisk   { background: #FEEBC8; color: #7B341E; }

/* ── Action item row ── */
.action-row {
    display: flex; align-items: flex-start; gap: 14px;
    padding: 14px 18px;
    border-radius: 8px;
    border: 1px solid #E2E8F0;
    background: #fff;
    margin-bottom: 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,.04);
}
.action-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; margin-top: 5px; }
.dot-open    { background: #CBD5E0; }
.dot-prog    { background: #3182CE; }
.dot-done    { background: #38A169; }
.dot-blocked { background: #E53E3E; }
.action-body { flex: 1; }
.action-title { font-size: .9rem; font-weight: 600; color: #1A202C; margin-bottom: 4px; }
.action-meta  { font-size: .75rem; color: #718096; }
.action-note  { font-size: .78rem; color: #4A5568; margin-top: 6px; font-style: italic; }
.action-right { text-align: right; flex-shrink: 0; }
.action-due   { font-size: .74rem; color: #718096; margin-bottom: 6px; }

/* ── Update feed ── */
.update-card {
    background: #fff;
    border: 1px solid #E2E8F0;
    border-left: 3px solid #29B5E8;
    border-radius: 0 8px 8px 0;
    padding: 16px 20px;
    margin-bottom: 12px;
    box-shadow: 0 1px 3px rgba(0,0,0,.04);
}
.update-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
.update-by   { font-size: .8rem; font-weight: 600; color: #2D3748; }
.update-date { font-size: .75rem; color: #A0AEC0; }
.update-text { font-size: .88rem; color: #4A5568; line-height: 1.7; }

/* ── Summary bar ── */
.summary-bar { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 24px; }
.summary-chip {
    background: #EBF8FF; border: 1px solid #BEE3F8;
    border-radius: 20px; padding: 6px 16px;
    font-size: .78rem; font-weight: 600; color: #0B6E9B;
}
.summary-chip.green { background:#F0FFF4; border-color:#9AE6B4; color:#22543D; }
.summary-chip.amber { background:#FFFAF0; border-color:#FBD38D; color:#7B341E; }
.summary-chip.gray  { background:#F7FAFC; border-color:#CBD5E0; color:#4A5568; }
</style>
""", unsafe_allow_html=True)


def status_badge(s: str) -> str:
    m = {
        "Met":             "badge-met",
        "On Track":        "badge-track",
        "At Risk":         "badge-risk",
        "Not Started":     "badge-none",
        "Complete":        "badge-met",
        "In Progress":     "badge-track",
        "Open":            "badge-none",
        "Blocked":         "badge-blocked",
        "Completed — Won": "badge-won",
        "Completed — Lost":"badge-lost",
        "Active":          "badge-active",
        "Planning":        "badge-planning",
    }
    cls = m.get(s, "badge-none")
    return f'<span class="kpi-badge {cls}">{s}</span>'

def kpi_card_class(s: str) -> str:
    return {"Met":"met","On Track":"track","At Risk":"risk"}.get(s,"none")

def action_dot_class(s: str) -> str:
    return {"Complete":"dot-done","In Progress":"dot-prog","Blocked":"dot-blocked"}.get(s,"dot-open")

def render_kpi_cards(df: pd.DataFrame):
    if df.empty:
        st.info("No KPIs defined yet.")
        return
    cols = st.columns(3)
    for i, (_, row) in enumerate(df.iterrows()):
        s    = str(row.get("Status",""))
        val  = str(row.get("Current Value","—"))
        unit = str(row.get("Unit",""))
        with cols[i % 3]:
            st.markdown(f"""
<div class="kpi-card {kpi_card_class(s)}">
  <div class="kpi-name">{row.get("KPI","")}</div>
  <div class="kpi-value">{val} <span style="font-size:1rem;font-weight:400;color:#718096">{unit}</span></div>
  <div class="kpi-target">Target: {row.get("Target","—")} {unit}</div>
  {status_badge(s)}
</div>""", unsafe_allow_html=True)

def render_actions(df: pd.DataFrame):
    if df.empty:
        st.info("No action items yet.")
        return
    for _, row in df.iterrows():
        s     = str(row.get("Status","Open"))
        dot   = action_dot_class(s)
        note  = str(row.get("Notes",""))
        note_html = f'<div class="action-note">💬 {note}</div>' if note else ""
        st.markdown(f"""
<div class="action-row">
  <div class="action-dot {dot}"></div>
  <div class="action-body">
    <div class="action-title">{row.get("Action","")}</div>
    <div class="action-meta">
      <b>{row.get("Owner","")}</b> &nbsp;·&nbsp; {row.get("Category","")}
    </div>
    {note_html}
  </div>
  <div class="action-right">
    <div class="action-due">📅 {row.get("Due Date","")}</div>
    {status_badge(s)}
  </div>
</div>""", unsafe_allow_html=True)

def render_updates(df: pd.DataFrame):
    if df.empty:
        st.info("No updates posted yet.")
        return
    sorted_df = df.sort_values("Date", ascending=False)
    for _, row in sorted_df.iterrows():
        st.markdown(f"""
<div class="update-card">
  <div class="update-header">
    <span class="update-by">🧑‍💼 {row.get("Posted By","")}</span>
    <span class="update-date">{row.get("Date","")}</span>
  </div>
  <div class="update-text">{row.get("Update","")}</div>
</div>""", unsafe_allow_html=True)

def render_summary_bar(ov: dict, kpis: pd.DataFrame, actions: pd.DataFrame):
    met   = (kpis["Status"] == "Met").sum()    if not kpis.empty else 0
    total_kpi = len(kpis)
    done  = (actions["Status"] == "Complete").sum()  if not actions.empty else 0
    open_ = (actions["Status"] == "Open").sum()      if not actions.empty else 0
    inprog= (actions["Status"] == "In Progress").sum()if not actions.empty else 0
    status = ov.get("Status","")
    arr    = ov.get("Potential ARR ($)","")
    arr_str = f"${int(arr):,}" if str(arr).isdigit() else (f"${arr}" if arr else "—")
    chips = [
        (f"{status_badge(status)}", ""),
        (f"💰 ARR: {arr_str}", "green"),
        (f"✅ KPIs Met: {met}/{total_kpi}", "green" if met==total_kpi else ""),
        (f"📋 Actions: {done} done · {inprog} in progress · {open_} open", "amber" if open_ > 0 else "green"),
    ]
    html = '<div class="summary-bar">'
    for text, cls in chips:
        html += f'<div class="summary-chip {cls}">{text}</div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


# ── Sidebar — POC selector ────────────────────────────────────────────────────
registry = load_registry()

with st.sidebar:
    st.markdown("## 🏢 POC Tracker")
    st.divider()

    if registry.empty:
        st.info("No POCs yet. Create one below.")
        poc_options = []
    else:
        poc_options = registry["POC_ID"].tolist()

    # POC selector
    if poc_options:
        labels = []
        for _, r in registry.iterrows():
            emoji      = STATUS_EMOJI.get(r.get("Status", ""), "⚪")
            engagement = r.get("Engagement", "")
            label      = f"{emoji}  {r['Customer']}  —  {engagement}" if engagement else f"{emoji}  {r['Customer']}"
            labels.append(label)

        selected_idx = st.radio(
            "Select Engagement",
            range(len(poc_options)),
            format_func=lambda i: labels[i],
            key="poc_radio",
        )
        active_poc_id = poc_options[selected_idx]
        active_row    = registry[registry["POC_ID"] == active_poc_id].iloc[0]
        st.caption(f"AE: {active_row.get('AE','')}  ·  SE: {active_row.get('SE','')}")
    else:
        active_poc_id = None
        active_row    = {}

    st.divider()

    # New POC
    with st.expander("＋  New POC", expanded=(active_poc_id is None)):
        with st.form("new_poc_form"):
            np_customer    = st.text_input("Customer Name *")
            np_engagement  = st.text_input("Engagement / BU", placeholder="e.g. Drug Development · RWE")
            np_ae          = st.text_input("Snowflake AE")
            np_se          = st.text_input("Snowflake SE")
            np_status      = st.selectbox("Status", STATUS_OPTIONS)
            if st.form_submit_button("Create POC", type="primary"):
                if np_customer.strip():
                    slug   = (np_customer.strip() + "-" + np_engagement.strip()).lower().replace(" ", "-").replace("·","").replace("/","")[:32]
                    create_poc(slug, np_customer.strip(), np_engagement.strip(), np_ae, np_se, np_status)
                    st.success(f"Created: {np_customer} — {np_engagement}")
                    st.rerun()
                else:
                    st.warning("Customer name required.")

# ── Guard ─────────────────────────────────────────────────────────────────────
if not active_poc_id:
    st.title("Snowflake POC Tracker")
    st.info("Create your first POC engagement using the sidebar.")
    st.stop()

# ── Load data for active POC ──────────────────────────────────────────────────
ov      = load_overview(active_poc_id)
kpis    = load_kpis(active_poc_id)
actions = load_actions(active_poc_id)
updates = load_updates(active_poc_id)

# ── Page header ───────────────────────────────────────────────────────────────
customer_name   = active_row.get("Customer", active_poc_id)
engagement_name = active_row.get("Engagement", "")
poc_status_emoji = STATUS_EMOJI.get(ov.get("Status", active_row.get("Status", "")), "⚪")

st.title(f"{customer_name} × Snowflake — POC Tracker")
if engagement_name:
    st.subheader(engagement_name, divider=False)

render_summary_bar(ov, kpis, actions)
st.divider()

tab0, tab1, tab2, tab3, tab4 = st.tabs(["📊 Sheet View", "🏢 Overview", "📈 KPIs & Budget", "✅ Action Plan", "📝 Updates"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 — SHEET VIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab0:
    st.subheader("Live Sheet Data")
    st.caption(f"Showing data for **{customer_name}** only. Refreshes every 20 s.")

    st.markdown("#### Overview")
    if ov:
        ov_df = pd.DataFrame([{k: v for k, v in ov.items() if k != "POC_ID"}])
        half = len(ov_df.columns) // 2
        cols = list(ov_df.columns)
        st.dataframe(ov_df[cols[:half]], use_container_width=True, hide_index=True)
        st.dataframe(ov_df[cols[half:]], use_container_width=True, hide_index=True)
    else:
        st.dataframe(pd.DataFrame(columns=OVERVIEW_HEADERS[1:]), use_container_width=True, hide_index=True)
        st.caption("No overview data yet.")

    st.divider()
    st.markdown("#### KPIs")
    st.dataframe(kpis if not kpis.empty else pd.DataFrame(columns=KPI_HEADERS[1:]),
                 use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("#### Mutual Action Plan")
    st.dataframe(actions if not actions.empty else pd.DataFrame(columns=ACTION_HEADERS[1:]),
                 use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("#### Status Updates")
    if not updates.empty:
        st.dataframe(updates.sort_values("Date", ascending=False),
                     use_container_width=True, hide_index=True)
    else:
        st.dataframe(pd.DataFrame(columns=UPDATE_HEADERS[1:]), use_container_width=True, hide_index=True)

    st.divider()
    if st.button("Refresh"):
        st.cache_data.clear()
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    if ov:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Engagement")
            for k, v in {
                "Business Unit":     ov.get("Business Unit"),
                "Primary Use Case":  ov.get("Primary Use Case"),
                "Secondary Use Case":ov.get("Secondary Use Case"),
                "Cloud Environment": ov.get("Cloud Environment"),
                "Current Platform":  ov.get("Current Data Platform"),
                "Data Volume":       ov.get("Data Volume"),
                "Compliance":        ov.get("Compliance Requirements"),
                "POC Start":         ov.get("POC Start Date"),
                "Target Completion": ov.get("Target Completion Date"),
                "Status":            ov.get("Status"),
            }.items():
                if v:
                    st.markdown(f"**{k}:** {v}")

        with col2:
            st.subheader("Contacts")
            for k, v in {
                "Snowflake AE":      ov.get("Snowflake AE"),
                "Snowflake SE":      ov.get("Snowflake SE"),
                "Champion":          ov.get("Labcorp Champion"),
                "Champion Title":    ov.get("Champion Title"),
                "Champion Email":    ov.get("Champion Email"),
                "Exec Sponsor":      ov.get("Executive Sponsor"),
                "Sponsor Title":     ov.get("Sponsor Title"),
            }.items():
                if v:
                    st.markdown(f"**{k}:** {v}")

        if ov.get("POC Objective"):
            st.divider()
            st.subheader("POC Objective")
            st.write(ov["POC Objective"])

        c1, c2 = st.columns(2)
        with c1:
            if ov.get("Technical Success Criteria"):
                st.subheader("Technical Success Criteria")
                st.write(ov["Technical Success Criteria"])
        with c2:
            if ov.get("Business Success Criteria"):
                st.subheader("Business Success Criteria")
                st.write(ov["Business Success Criteria"])

        st.divider()

    with st.expander("Edit Overview", expanded=not bool(ov)):
        with st.form("overview_form"):
            st.markdown("**Snowflake Team**")
            f1, f2 = st.columns(2)
            with f1:
                ae = st.text_input("Account Executive", value=ov.get("Snowflake AE", ""))
            with f2:
                se = st.text_input("Solutions Engineer", value=ov.get("Snowflake SE", ""))

            st.markdown("**Customer Contacts**")
            f3, f4 = st.columns(2)
            with f3:
                champion      = st.text_input("Champion Name",   value=ov.get("Labcorp Champion", ""))
                champ_title   = st.text_input("Champion Title",  value=ov.get("Champion Title", ""))
                champ_email   = st.text_input("Champion Email",  value=ov.get("Champion Email", ""))
            with f4:
                sponsor       = st.text_input("Executive Sponsor",       value=ov.get("Executive Sponsor", ""))
                sponsor_title = st.text_input("Executive Sponsor Title", value=ov.get("Sponsor Title", ""))

            st.markdown("**Engagement Details**")
            f5, f6 = st.columns(2)
            with f5:
                bu  = st.selectbox("Business Unit",     BUSINESS_UNITS,
                                   index=BUSINESS_UNITS.index(ov["Business Unit"]) if ov.get("Business Unit") in BUSINESS_UNITS else 0)
                uc  = st.selectbox("Primary Use Case",  USE_CASES,
                                   index=USE_CASES.index(ov["Primary Use Case"]) if ov.get("Primary Use Case") in USE_CASES else 0)
                uc2 = st.selectbox("Secondary Use Case",USE_CASES,
                                   index=USE_CASES.index(ov["Secondary Use Case"]) if ov.get("Secondary Use Case") in USE_CASES else 0)
            with f6:
                cloud    = st.selectbox("Cloud Environment", CLOUD_OPTIONS,
                                        index=CLOUD_OPTIONS.index(ov["Cloud Environment"]) if ov.get("Cloud Environment") in CLOUD_OPTIONS else 0)
                platform = st.text_input("Current Data Platform", value=ov.get("Current Data Platform", ""))
                volume   = st.text_input("Estimated Data Volume",  value=ov.get("Data Volume", ""))

            saved_compliance = [c.strip() for c in ov.get("Compliance Requirements", "").split(",") if c.strip()]
            compliance = st.multiselect("Compliance Requirements", COMPLIANCE_OPTIONS,
                                        default=[c for c in saved_compliance if c in COMPLIANCE_OPTIONS])

            st.markdown("**Timeline & Status**")
            f7, f8, f9 = st.columns(3)
            with f7:
                try:    start_val = datetime.strptime(ov.get("POC Start Date", ""), "%Y-%m-%d").date()
                except: start_val = date.today()
                start_date = st.date_input("POC Start Date", value=start_val)
            with f8:
                try:    end_val = datetime.strptime(ov.get("Target Completion Date", ""), "%Y-%m-%d").date()
                except: end_val = date.today()
                end_date = st.date_input("Target Completion Date", value=end_val)
            with f9:
                status = st.selectbox("POC Status", STATUS_OPTIONS,
                                      index=STATUS_OPTIONS.index(ov["Status"]) if ov.get("Status") in STATUS_OPTIONS else 0)

            objective     = st.text_area("POC Objective",              value=ov.get("POC Objective", ""), height=80)
            fa, fb = st.columns(2)
            with fa:
                tech_criteria = st.text_area("Technical Success Criteria", value=ov.get("Technical Success Criteria", ""), height=90)
            with fb:
                biz_criteria  = st.text_area("Business Success Criteria",  value=ov.get("Business Success Criteria", ""),  height=90)

            if st.form_submit_button("Save Overview", type="primary"):
                save_overview(active_poc_id, {
                    "Snowflake AE": ae, "Snowflake SE": se,
                    "Labcorp Champion": champion, "Champion Title": champ_title, "Champion Email": champ_email,
                    "Executive Sponsor": sponsor, "Sponsor Title": sponsor_title,
                    "Business Unit": bu, "Primary Use Case": uc, "Secondary Use Case": uc2,
                    "Cloud Environment": cloud, "Current Data Platform": platform, "Data Volume": volume,
                    "Compliance Requirements": ", ".join(compliance),
                    "POC Start Date": str(start_date), "Target Completion Date": str(end_date),
                    "Status": status, "POC Objective": objective,
                    "Technical Success Criteria": tech_criteria, "Business Success Criteria": biz_criteria,
                    "POC Budget ($)": ov.get("POC Budget ($)", ""),
                    "Confirmed Spend ($)": ov.get("Confirmed Spend ($)", ""),
                    "Potential ARR ($)": ov.get("Potential ARR ($)", ""),
                    "Procurement Contact": ov.get("Procurement Contact", ""),
                    "Budget Notes": ov.get("Budget Notes", ""),
                })
                st.success("Saved.")
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — KPIs & BUDGET
# ══════════════════════════════════════════════════════════════════════════════
with tab2:

    st.subheader("KPIs")
    DISPLAY_KPI = [h for h in KPI_HEADERS if h != "POC_ID"]

    # Styled KPI cards
    render_kpi_cards(kpis)

    if not kpis.empty:
        st.divider()
        st.markdown("**Update a KPI**")
        sheet_rows = get_all_sheet_row("KPIs", active_poc_id)
        for (sheet_row, rec), (df_idx, df_row) in zip(sheet_rows, kpis.iterrows()):
            with st.expander(f"{df_row.get('KPI', '')} — {df_row.get('Status', '')}"):
                with st.form(f"kpi_edit_{df_idx}"):
                    kc1, kc2 = st.columns(2)
                    with kc1:
                        kpi_name = st.text_input("KPI",    value=df_row.get("KPI", ""))
                        target   = st.text_input("Target", value=str(df_row.get("Target", "")))
                        unit     = st.text_input("Unit",   value=df_row.get("Unit", ""))
                    with kc2:
                        current  = st.text_input("Current Value", value=str(df_row.get("Current Value", "")))
                        ks_idx   = KPI_STATUSES.index(df_row["Status"]) if df_row.get("Status") in KPI_STATUSES else 0
                        k_status = st.selectbox("Status", KPI_STATUSES, index=ks_idx)
                        k_note   = st.text_input("Notes",  value=df_row.get("Notes", ""))
                    sc1, sc2 = st.columns([4, 1])
                    with sc1:
                        if st.form_submit_button("Update", type="primary"):
                            update_sheet_row("KPIs", active_poc_id, sheet_row, DISPLAY_KPI,
                                             {"KPI": kpi_name, "Target": target, "Current Value": current,
                                              "Unit": unit, "Status": k_status, "Notes": k_note})
                            st.success("Updated.")
                            st.rerun()
                    with sc2:
                        if st.form_submit_button("Delete"):
                            delete_sheet_row("KPIs", sheet_row)
                            st.rerun()

    st.divider()
    st.markdown("**Add KPI**")
    with st.form("new_kpi"):
        nk1, nk2, nk3 = st.columns(3)
        with nk1:
            nk_name   = st.text_input("KPI *")
            nk_target = st.text_input("Target")
            nk_unit   = st.text_input("Unit")
        with nk2:
            nk_current = st.text_input("Current Value")
            nk_status  = st.selectbox("Status", KPI_STATUSES)
        with nk3:
            nk_note = st.text_area("Notes", height=108)
        if st.form_submit_button("Add KPI", type="primary"):
            if nk_name.strip():
                append_to("KPIs", active_poc_id,
                          [nk_name.strip(), nk_target, nk_current, nk_unit, nk_status, nk_note])
                st.success("Added.")
                st.rerun()
            else:
                st.warning("KPI name required.")

    st.divider()

    # Budget — styled metric cards
    st.subheader("Budget & Funding")
    budget_val = ov.get("POC Budget ($)", "")
    spend_val  = ov.get("Confirmed Spend ($)", "")
    arr_val    = ov.get("Potential ARR ($)", "")
    if budget_val or spend_val or arr_val:
        b1, b2, b3 = st.columns(3)
        def fmt_money(v):
            try: return f"${int(v):,}"
            except: return f"${v}" if v else "—"
        with b1:
            st.metric("POC Budget", fmt_money(budget_val))
        with b2:
            st.metric("Confirmed Spend", fmt_money(spend_val))
        with b3:
            st.metric("Potential ARR", fmt_money(arr_val))
        if ov.get("Budget Notes"):
            st.caption(f"📌 {ov['Budget Notes']}")
    else:
        st.info("No budget details entered yet.")

    with st.expander("Edit Budget Details"):
        with st.form("budget_form"):
            bf1, bf2, bf3 = st.columns(3)
            with bf1:
                poc_budget = st.text_input("POC Budget ($)",      value=ov.get("POC Budget ($)", ""))
            with bf2:
                conf_spend = st.text_input("Confirmed Spend ($)", value=ov.get("Confirmed Spend ($)", ""))
            with bf3:
                arr        = st.text_input("Potential ARR ($)",   value=ov.get("Potential ARR ($)", ""))
            proc_contact = st.text_input("Procurement Contact", value=ov.get("Procurement Contact", ""))
            budget_notes = st.text_area("Budget Notes",         value=ov.get("Budget Notes", ""), height=80)
            if st.form_submit_button("Save Budget", type="primary"):
                updated = dict(ov)
                updated.update({
                    "POC Budget ($)": poc_budget, "Confirmed Spend ($)": conf_spend,
                    "Potential ARR ($)": arr, "Procurement Contact": proc_contact,
                    "Budget Notes": budget_notes,
                })
                save_overview(active_poc_id, updated)
                st.success("Saved.")
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — ACTION PLAN
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Mutual Action Plan")
    DISPLAY_ACT = [h for h in ACTION_HEADERS if h != "POC_ID"]

    # Styled action list
    render_actions(actions)

    if not actions.empty:
        st.divider()
        st.markdown("**Update an Action Item**")
        sheet_rows = get_all_sheet_row("Action_Items", active_poc_id)
        for (sheet_row, rec), (df_idx, df_row) in zip(sheet_rows, actions.iterrows()):
            label = f"{df_row.get('Action', '')} — {df_row.get('Owner', '')} — {df_row.get('Status', '')}"
            with st.expander(label):
                with st.form(f"action_edit_{df_idx}"):
                    action_text = st.text_input("Action", value=df_row.get("Action", ""))
                    ac1, ac2, ac3 = st.columns(3)
                    with ac1:
                        owner = st.selectbox("Owner", ACTION_OWNERS,
                                             index=ACTION_OWNERS.index(df_row["Owner"]) if df_row.get("Owner") in ACTION_OWNERS else 0)
                    with ac2:
                        cat   = st.selectbox("Category", ACTION_CATS,
                                             index=ACTION_CATS.index(df_row["Category"]) if df_row.get("Category") in ACTION_CATS else 0)
                    with ac3:
                        a_st  = st.selectbox("Status", ACTION_STATUSES,
                                             index=ACTION_STATUSES.index(df_row["Status"]) if df_row.get("Status") in ACTION_STATUSES else 0)
                    try:    due_val = datetime.strptime(df_row.get("Due Date", ""), "%Y-%m-%d").date()
                    except: due_val = date.today()
                    due  = st.date_input("Due Date", value=due_val, key=f"due_{df_idx}")
                    note = st.text_input("Notes", value=df_row.get("Notes", ""))
                    sc1, sc2 = st.columns([4, 1])
                    with sc1:
                        if st.form_submit_button("Update", type="primary"):
                            update_sheet_row("Action_Items", active_poc_id, sheet_row, DISPLAY_ACT,
                                             {"Action": action_text, "Owner": owner, "Category": cat,
                                              "Due Date": str(due), "Status": a_st, "Notes": note})
                            st.success("Updated.")
                            st.rerun()
                    with sc2:
                        if st.form_submit_button("Delete"):
                            delete_sheet_row("Action_Items", sheet_row)
                            st.rerun()

    st.divider()
    st.markdown("**Add Action Item**")
    with st.form("new_action"):
        new_text = st.text_input("Action *")
        na1, na2, na3, na4 = st.columns(4)
        with na1: new_owner  = st.selectbox("Owner",    ACTION_OWNERS)
        with na2: new_cat    = st.selectbox("Category", ACTION_CATS)
        with na3: new_due    = st.date_input("Due Date", value=date.today())
        with na4: new_status = st.selectbox("Status",   ACTION_STATUSES)
        new_note = st.text_input("Notes (optional)")
        if st.form_submit_button("Add", type="primary"):
            if new_text.strip():
                append_to("Action_Items", active_poc_id,
                          [new_text.strip(), new_owner, new_cat, str(new_due), new_status, new_note])
                st.success("Added.")
                st.rerun()
            else:
                st.warning("Action text required.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — UPDATES
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Status Updates")

    # Styled update feed
    render_updates(updates)

    st.divider()
    st.markdown("**Post a New Update**")
    with st.form("new_update"):
        uf1, uf2 = st.columns([3, 1])
        with uf1:
            update_text = st.text_area("Update *", height=100,
                                       placeholder="Key outcomes, decisions, blockers, or next steps…")
        with uf2:
            posted_by = st.text_input("Posted by *")
        if st.form_submit_button("Post Update", type="primary"):
            if update_text.strip() and posted_by.strip():
                append_to("Updates", active_poc_id,
                          [datetime.today().strftime("%Y-%m-%d %H:%M"), update_text.strip(), posted_by.strip()])
                st.success("Posted.")
                st.rerun()
            else:
                st.warning("Both fields are required.")
