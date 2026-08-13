import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date
import uuid
import time

def _retry(fn, retries=3, wait=8):
    for attempt in range(retries):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and attempt < retries - 1:
                time.sleep(wait * (attempt + 1))
            else:
                raise

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
    "Technical Champion", "Technical Champion Title",
    "Exec Business Sponsor", "Exec Business Sponsor Title",
    "Customer Participants",
    "Business Unit",
    "Cloud Environment", "Current Data Platform", "Data Volume",
    "Compliance Requirements",
    "POC Start Date", "Target Completion Date", "Status",
    "POC Objective", "Technical Success Criteria", "Business Success Criteria",
    "POC Budget ($)", "Confirmed Spend ($)", "Potential ARR ($)",
    "Procurement Contact", "Budget Notes",
]

KPI_HEADERS      = ["POC_ID", "KPI", "Target", "Current Value", "Unit", "Status", "Notes"]
ACTION_HEADERS   = ["POC_ID", "Action", "Assignee", "Category", "Due Date", "Status", "Notes"]
AUDIENCE_HEADERS = ["POC_ID", "Name", "Company", "Role", "Email"]
TIMELINE_HEADERS = ["POC_ID", "Milestone", "Due Date", "Owner", "Status", "Notes"]

BUSINESS_UNITS     = ["Diagnostics", "Drug Development (Biopharma Solutions)", "Genomics", "Technology Solutions", "Enterprise / Cross-BU"]
COMPLIANCE_OPTIONS = ["HIPAA", "21 CFR Part 11", "GxP", "SOC 2", "CLIA", "GDPR"]
STATUS_OPTIONS     = ["Planning", "Active", "At Risk", "Completed — Won", "Completed — Lost"]
KPI_STATUSES       = ["On Track", "At Risk", "Met", "Not Started"]
ACTION_CATS        = ["Technical", "Business", "Compliance", "Training", "Executive"]
ACTION_STATUSES    = ["Open", "In Progress", "Complete", "Blocked"]
TIMELINE_STATUSES  = ["Not Started", "In Progress", "At Risk", "Complete"]
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

@st.cache_resource
def get_ss():
    return get_client().open_by_key(st.secrets["spreadsheet_id"])

@st.cache_resource
def ensure_sheets():
    ss = get_ss()
    existing = {ws.title for ws in ss.worksheets()}
    for title, headers in [
        ("POC_Registry",  REGISTRY_HEADERS),
        ("Overview",      OVERVIEW_HEADERS),
        ("KPIs",          KPI_HEADERS),
        ("Action_Items",  ACTION_HEADERS),
        ("Audience",      AUDIENCE_HEADERS),
        ("Timeline",      TIMELINE_HEADERS),
    ]:
        if title not in existing:
            ws = ss.add_worksheet(title=title, rows=1000, cols=max(len(headers), 10))
            ws.append_row(headers)
        else:
            ws = ss.worksheet(title)
            if not ws.row_values(1):
                ws.insert_row(headers, 1)

# ── Load functions ────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_registry() -> pd.DataFrame:
    ws = get_ss().worksheet("POC_Registry")
    data = ws.get_all_records()
    if data:
        return pd.DataFrame(data)
    return pd.DataFrame(columns=REGISTRY_HEADERS)

@st.cache_data(ttl=300)
def load_overview(poc_id: str) -> dict:
    ws = get_ss().worksheet("Overview")
    rows = ws.get_all_records()
    for r in rows:
        if r.get("POC_ID") == poc_id:
            return r
    return {}

@st.cache_data(ttl=300)
def load_kpis(poc_id: str) -> pd.DataFrame:
    ws = get_ss().worksheet("KPIs")
    data = ws.get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=KPI_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        df = df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

@st.cache_data(ttl=300)
def load_actions(poc_id: str) -> pd.DataFrame:
    ws = get_ss().worksheet("Action_Items")
    data = ws.get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=ACTION_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        df = df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

@st.cache_data(ttl=300)
def load_audience(poc_id: str) -> pd.DataFrame:
    ws = get_ss().worksheet("Audience")
    data = ws.get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=AUDIENCE_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        df = df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

@st.cache_data(ttl=300)
def load_timeline(poc_id: str) -> pd.DataFrame:
    ws = get_ss().worksheet("Timeline")
    data = ws.get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=TIMELINE_HEADERS)
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

def save_all_rows(sheet: str, poc_id: str, display_headers: list, edited_df: pd.DataFrame):
    """Replace all rows for poc_id in a sheet with the edited dataframe."""
    ws = get_ss().worksheet(sheet)
    all_vals = ws.get_all_values()
    to_delete = [i + 2 for i, r in enumerate(all_vals[1:]) if r and r[0] == poc_id]
    for row_num in reversed(to_delete):
        ws.delete_rows(row_num)
    for _, row in edited_df.iterrows():
        ws.append_row(
            [poc_id] + [str(row.get(h, "")) for h in display_headers],
            value_input_option="USER_ENTERED",
        )
    st.cache_data.clear()

# ── Bootstrap ─────────────────────────────────────────────────────────────────
try:
    _retry(ensure_sheets)
except Exception as e:
    st.error(f"Could not connect to Google Sheets: {e}")
    st.stop()

# ── Premium CSS injection ─────────────────────────────────────────────────────
st.html("""
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
""")


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
        (f"✅ KPIs Met: {met}/{total_kpi}", "green" if met==total_kpi else ""),
        (f"📋 Actions: {done} done · {inprog} in progress · {open_} open", "amber" if open_ > 0 else "green"),
    ]
    html = '<div class="summary-bar">'
    for text, cls in chips:
        html += f'<div class="summary-chip {cls}">{text}</div>'
    html += '</div>'
    st.html(html)


# ── Sidebar — POC selector ────────────────────────────────────────────────────
registry = load_registry()

with st.sidebar:
    st.markdown("## 🏢 POC Tracker")
    st.divider()

    # ── Editor identity ───────────────────────────────────
    if "editor_name" not in st.session_state:
        st.session_state["editor_name"] = ""
    if "editor_role" not in st.session_state:
        st.session_state["editor_role"] = "Snowflake Team"

    st.markdown("**Editing as**")
    st.session_state["editor_name"] = st.text_input(
        "Your name", value=st.session_state["editor_name"],
        placeholder="e.g. Geet Sukhmani", label_visibility="collapsed"
    )
    st.session_state["editor_role"] = st.radio(
        "Role", ["Snowflake Team", "Customer"],
        index=0 if st.session_state["editor_role"] == "Snowflake Team" else 1,
        horizontal=True, label_visibility="collapsed"
    )
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
ov       = load_overview(active_poc_id)
kpis     = load_kpis(active_poc_id)
actions  = load_actions(active_poc_id)
audience = load_audience(active_poc_id)
timeline = load_timeline(active_poc_id)

# ── Page header ───────────────────────────────────────────────────────────────
customer_name   = active_row.get("Customer", active_poc_id)
engagement_name = active_row.get("Engagement", "")
poc_status_emoji = STATUS_EMOJI.get(ov.get("Status", active_row.get("Status", "")), "⚪")

st.title(f"{customer_name} × Snowflake — POC Tracker")
if engagement_name:
    st.subheader(engagement_name, divider=False)

render_summary_bar(ov, kpis, actions)
st.divider()

tab1, tab2, tab3, tab4, tab5 = st.tabs(["🏢 Overview", "📈 KPIs", "👥 Audience", "✅ Action Plan", "📅 Timeline"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    editor_name = st.session_state.get("editor_name", "")
    editor_role = st.session_state.get("editor_role", "Snowflake Team")

    with st.form("overview_form"):
        st.markdown("### Contacts")
        sf_col, cust_col = st.columns(2)

        with sf_col:
            st.markdown("**🔵 Snowflake Team**")
            ae = st.text_input("Account Executive", value=ov.get("Snowflake AE", ""))
            se = st.text_input("Solutions Engineer", value=ov.get("Snowflake SE", ""))

        with cust_col:
            st.markdown("**🟢 Customer Contacts**")
            champion      = st.text_input("Technical Champion",            value=ov.get("Technical Champion", ""))
            champ_title   = st.text_input("Technical Champion Title",      value=ov.get("Technical Champion Title", ""))
            sponsor       = st.text_input("Exec Business Sponsor",         value=ov.get("Exec Business Sponsor", ""))
            sponsor_title = st.text_input("Exec Business Sponsor Title",   value=ov.get("Exec Business Sponsor Title", ""))

        participants = st.text_area(
            "Customer Participants",
            value=ov.get("Customer Participants", ""),
            height=80,
            placeholder="e.g. Jane Smith (Data Engineer), John Doe (IT Lead)…",
        )

        st.divider()
        st.markdown("### Engagement Details")
        e1, e2 = st.columns(2)
        with e1:
            bu = st.selectbox("Business Unit", BUSINESS_UNITS,
                               index=BUSINESS_UNITS.index(ov["Business Unit"]) if ov.get("Business Unit") in BUSINESS_UNITS else 0)
        with e2:
            cloud    = st.selectbox("Cloud Environment", CLOUD_OPTIONS,
                                    index=CLOUD_OPTIONS.index(ov["Cloud Environment"]) if ov.get("Cloud Environment") in CLOUD_OPTIONS else 0)
            platform = st.text_input("Current Data Platform", value=ov.get("Current Data Platform", ""))
            volume   = st.text_input("Estimated Data Volume",  value=ov.get("Data Volume", ""))

        saved_compliance = [c.strip() for c in ov.get("Compliance Requirements", "").split(",") if c.strip()]
        compliance = st.multiselect("Compliance Requirements", COMPLIANCE_OPTIONS,
                                    default=[c for c in saved_compliance if c in COMPLIANCE_OPTIONS])

        st.divider()
        st.markdown("### Timeline & Status")
        t1, t2, t3 = st.columns(3)
        with t1:
            try:    start_val = datetime.strptime(ov.get("POC Start Date", ""), "%Y-%m-%d").date()
            except: start_val = date.today()
            start_date = st.date_input("POC Start Date", value=start_val)
        with t2:
            try:    end_val = datetime.strptime(ov.get("Target Completion Date", ""), "%Y-%m-%d").date()
            except: end_val = date.today()
            end_date = st.date_input("Target Completion Date", value=end_val)
        with t3:
            status = st.selectbox("POC Status", STATUS_OPTIONS,
                                  index=STATUS_OPTIONS.index(ov["Status"]) if ov.get("Status") in STATUS_OPTIONS else 0)

        st.divider()
        st.markdown("### Objectives & Success Criteria")
        objective = st.text_area("POC Objective", value=ov.get("POC Objective", ""), height=90)
        sc1, sc2 = st.columns(2)
        with sc1:
            tech_criteria = st.text_area("Technical Success Criteria", value=ov.get("Technical Success Criteria", ""), height=100)
        with sc2:
            biz_criteria  = st.text_area("Business Success Criteria",  value=ov.get("Business Success Criteria", ""),  height=100)

        st.divider()
        st.markdown("### Budget")
        b1, b2 = st.columns(2)
        with b1: poc_budget = st.text_input("POC Budget ($)",      value=ov.get("POC Budget ($)", ""))
        with b2: conf_spend = st.text_input("Confirmed Spend ($)", value=ov.get("Confirmed Spend ($)", ""))

        submitted = st.form_submit_button("💾  Save Overview", type="primary", use_container_width=True)
        if submitted:
            save_overview(active_poc_id, {
                "Snowflake AE": ae, "Snowflake SE": se,
                "Technical Champion": champion, "Technical Champion Title": champ_title,
                "Exec Business Sponsor": sponsor, "Exec Business Sponsor Title": sponsor_title,
                "Customer Participants": participants,
                "Business Unit": bu,
                "Cloud Environment": cloud, "Current Data Platform": platform, "Data Volume": volume,
                "Compliance Requirements": ", ".join(compliance),
                "POC Start Date": str(start_date), "Target Completion Date": str(end_date),
                "Status": status, "POC Objective": objective,
                "Technical Success Criteria": tech_criteria, "Business Success Criteria": biz_criteria,
                "POC Budget ($)": poc_budget, "Confirmed Spend ($)": conf_spend,
            })
            st.success(f"Saved by {editor_name or editor_role}.")
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — KPIs
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    DISPLAY_KPI = [h for h in KPI_HEADERS if h != "POC_ID"]

    st.subheader("KPIs")
    st.caption("Edit any cell directly, then click **Save KPIs** to write to the sheet.")

    kpi_edit_df = kpis.copy() if not kpis.empty else pd.DataFrame(columns=DISPLAY_KPI)
    # Ensure all columns exist
    for col in DISPLAY_KPI:
        if col not in kpi_edit_df.columns:
            kpi_edit_df[col] = ""
    kpi_edit_df = kpi_edit_df[DISPLAY_KPI]

    edited_kpis = st.data_editor(
        kpi_edit_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "KPI":           st.column_config.TextColumn("KPI", width="large"),
            "Target":        st.column_config.TextColumn("Target"),
            "Current Value": st.column_config.TextColumn("Current Value"),
            "Unit":          st.column_config.TextColumn("Unit"),
            "Status":        st.column_config.SelectboxColumn("Status", options=KPI_STATUSES),
            "Notes":         st.column_config.TextColumn("Notes", width="large"),
        },
        key="kpi_editor",
    )

    if st.button("💾  Save KPIs", type="primary", key="save_kpis"):
        clean = edited_kpis.dropna(how="all")
        clean = clean[clean["KPI"].astype(str).str.strip() != ""]
        save_all_rows("KPIs", active_poc_id, DISPLAY_KPI, clean)
        st.success(f"KPIs saved by {st.session_state.get('editor_name') or st.session_state.get('editor_role','')}")
        st.rerun()

    st.divider()

    # Budget summary (edit via Overview tab)
    st.subheader("Budget")

    def fmt_money(v):
        try: return f"${int(v):,}"
        except: return f"${v}" if v else "—"

    bm1, bm2 = st.columns(2)
    with bm1: st.metric("POC Budget",      fmt_money(ov.get("POC Budget ($)", "")))
    with bm2: st.metric("Confirmed Spend", fmt_money(ov.get("Confirmed Spend ($)", "")))
    st.caption("To edit, go to the Overview tab.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — AUDIENCE
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    DISPLAY_AUD = [h for h in AUDIENCE_HEADERS if h != "POC_ID"]

    st.subheader("Audience & Assignees")
    st.caption("Define the people involved in this POC. Assignees listed here appear in the Action Plan and Timeline.")

    aud_edit_df = audience.copy() if not audience.empty else pd.DataFrame(columns=DISPLAY_AUD)
    for col in DISPLAY_AUD:
        if col not in aud_edit_df.columns:
            aud_edit_df[col] = ""
    aud_edit_df = aud_edit_df[DISPLAY_AUD]

    edited_audience = st.data_editor(
        aud_edit_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Name":    st.column_config.TextColumn("Name", width="medium"),
            "Company": st.column_config.TextColumn("Company", width="medium"),
            "Role":    st.column_config.TextColumn("Role", width="medium"),
            "Email":   st.column_config.TextColumn("Email", width="large"),
        },
        key="audience_editor",
    )

    if st.button("💾  Save Audience", type="primary", key="save_audience"):
        clean = edited_audience.dropna(how="all")
        clean = clean[clean["Name"].astype(str).str.strip() != ""]
        save_all_rows("Audience", active_poc_id, DISPLAY_AUD, clean)
        st.success("Audience saved.")
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — ACTION PLAN
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    DISPLAY_ACT = [h for h in ACTION_HEADERS if h != "POC_ID"]
    assignee_names = audience["Name"].dropna().tolist() if not audience.empty else []

    st.subheader("Mutual Action Plan")
    st.caption("Edit any cell directly, then click **Save Actions**.")

    act_edit_df = actions.copy() if not actions.empty else pd.DataFrame(columns=DISPLAY_ACT)
    for col in DISPLAY_ACT:
        if col not in act_edit_df.columns:
            act_edit_df[col] = ""
    act_edit_df = act_edit_df[DISPLAY_ACT]

    assignee_col = (
        st.column_config.SelectboxColumn("Assignee", options=assignee_names)
        if assignee_names
        else st.column_config.TextColumn("Assignee")
    )

    edited_actions = st.data_editor(
        act_edit_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Action":   st.column_config.TextColumn("Action", width="large"),
            "Assignee": assignee_col,
            "Category": st.column_config.SelectboxColumn("Category", options=ACTION_CATS),
            "Due Date": st.column_config.TextColumn("Due Date", help="YYYY-MM-DD"),
            "Status":   st.column_config.SelectboxColumn("Status",   options=ACTION_STATUSES),
            "Notes":    st.column_config.TextColumn("Notes", width="large"),
        },
        key="action_editor",
    )

    if st.button("💾  Save Actions", type="primary", key="save_actions"):
        clean = edited_actions.dropna(how="all")
        clean = clean[clean["Action"].astype(str).str.strip() != ""]
        save_all_rows("Action_Items", active_poc_id, DISPLAY_ACT, clean)
        st.success(f"Actions saved by {st.session_state.get('editor_name') or st.session_state.get('editor_role','')}")
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — TIMELINE
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    DISPLAY_TL = [h for h in TIMELINE_HEADERS if h != "POC_ID"]

    st.subheader("POC Timeline")
    st.caption("Track milestones and key dates. Add rows for each phase or deliverable.")

    tl_edit_df = timeline.copy() if not timeline.empty else pd.DataFrame(columns=DISPLAY_TL)
    for col in DISPLAY_TL:
        if col not in tl_edit_df.columns:
            tl_edit_df[col] = ""
    tl_edit_df = tl_edit_df[DISPLAY_TL]

    owner_col = (
        st.column_config.SelectboxColumn("Owner", options=assignee_names)
        if assignee_names
        else st.column_config.TextColumn("Owner")
    )

    edited_timeline = st.data_editor(
        tl_edit_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Milestone": st.column_config.TextColumn("Milestone", width="large"),
            "Due Date":  st.column_config.TextColumn("Due Date", help="YYYY-MM-DD"),
            "Owner":     owner_col,
            "Status":    st.column_config.SelectboxColumn("Status", options=TIMELINE_STATUSES),
            "Notes":     st.column_config.TextColumn("Notes", width="large"),
        },
        key="timeline_editor",
    )

    if st.button("💾  Save Timeline", type="primary", key="save_timeline"):
        clean = edited_timeline.dropna(how="all")
        clean = clean[clean["Milestone"].astype(str).str.strip() != ""]
        save_all_rows("Timeline", active_poc_id, DISPLAY_TL, clean)
        st.success("Timeline saved.")
        st.rerun()
