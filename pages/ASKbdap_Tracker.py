import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date
import time

st.set_page_config(page_title="ASKbdap POC Tracker", layout="wide", initial_sidebar_state="collapsed")

# ── Single-POC config ─────────────────────────────────────────────────────────
POC_ID   = "askbdap"
CUSTOMER = "ASKbdap"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
REGISTRY_HEADERS = ["POC_ID", "Customer", "Engagement", "Status", "Created"]
OVERVIEW_HEADERS = [
    "POC_ID",
    "Technical Champion", "Technical Champion Title",
    "Exec Business Sponsor", "Exec Business Sponsor Title",
    "Customer Participants", "Business Unit",
    "Cloud Environment", "Current Data Platform", "Data Volume",
    "Compliance Requirements",
    "POC Start Date", "Target Completion Date", "Status",
    "POC Objective", "Technical Success Criteria", "Business Success Criteria",
    "POC Budget ($)", "Confirmed Spend ($)",
]
KPI_HEADERS    = ["POC_ID", "KPI", "Target", "Current Value", "Unit", "Status", "Notes"]
ACTION_HEADERS = ["POC_ID", "Action", "Assignee", "Category", "Due Date", "Status", "Notes"]

STATUS_OPTIONS     = ["Planning", "Active", "At Risk", "Completed — Won", "Completed — Lost"]
KPI_STATUSES       = ["On Track", "At Risk", "Met", "Not Started"]
ACTION_CATS        = ["Technical", "Business", "Compliance", "Training", "Executive"]
ACTION_STATUSES    = ["Open", "In Progress", "Complete", "Blocked"]
CLOUD_OPTIONS      = ["AWS", "Azure", "Multi-Cloud"]
BUSINESS_UNITS     = ["Diagnostics", "Drug Development (Biopharma Solutions)", "Genomics", "Technology Solutions", "Enterprise / Cross-BU"]
COMPLIANCE_OPTIONS = ["HIPAA", "21 CFR Part 11", "GxP", "SOC 2", "CLIA", "GDPR"]
STATUS_BADGE = {
    "Planning": "🔵 Planning", "Active": "🟢 Active", "At Risk": "🟡 At Risk",
    "Completed — Won": "✅ Won", "Completed — Lost": "🔴 Lost",
}

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, .stApp { font-family: 'Inter', sans-serif !important; }
#MainMenu, footer, .stDeployButton { display: none !important; }
.block-container { max-width: 1400px !important; padding: 1.5rem 2.5rem 3rem !important; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #E2E8F0; gap: 0; }
.stTabs [data-baseweb="tab"] {
    background: transparent; border-radius: 0; color: #64748B;
    font-weight: 500; font-size: .92rem; padding: 10px 24px;
    border-bottom: 2px solid transparent; margin-bottom: -2px;
}
.stTabs [aria-selected="true"] {
    color: #29B5E8 !important; border-bottom: 2px solid #29B5E8 !important;
    background: transparent !important;
}

/* Metric cards */
[data-testid="metric-container"] {
    background: #fff; border: 1px solid #E2E8F0; border-radius: 12px;
    padding: 16px 20px; box-shadow: 0 1px 4px rgba(0,0,0,.05);
}
[data-testid="stMetricLabel"] { font-size: .68rem !important; font-weight: 700 !important; text-transform: uppercase !important; letter-spacing: .07em !important; color: #94A3B8 !important; }
[data-testid="stMetricValue"] { font-size: 1.4rem !important; font-weight: 700 !important; color: #0F172A !important; }

/* Cards */
[data-testid="stVerticalBlockBorderWrapper"] {
    background: #fff !important; border: 1px solid #E2E8F0 !important;
    border-radius: 12px !important; box-shadow: 0 1px 4px rgba(0,0,0,.05) !important;
}
[data-testid="stVerticalBlockBorderWrapper"] > div > [data-testid="stVerticalBlock"] {
    padding: 20px 24px !important;
}

/* Inputs */
.stTextInput input, .stTextArea textarea {
    border-radius: 8px !important; border-color: #E2E8F0 !important;
    background: #FAFBFC !important; font-size: .84rem !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: #29B5E8 !important; box-shadow: 0 0 0 3px rgba(41,181,232,.12) !important;
    background: #fff !important;
}

/* Buttons */
.stButton > button[kind="primary"] {
    background: #29B5E8 !important; color: #0F172A !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important; font-size: .85rem !important;
}
.stButton > button { border-radius: 8px !important; font-size: .84rem !important; }

/* Progress */
.stProgress > div > div { background: #29B5E8 !important; height: 6px !important; border-radius: 4px !important; }
.stProgress > div { background: #E2E8F0 !important; height: 6px !important; border-radius: 4px !important; }
</style>
""", unsafe_allow_html=True)


# ── Retry ─────────────────────────────────────────────────────────────────────
def _retry(fn, retries=3, wait=8):
    for attempt in range(retries):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and attempt < retries - 1:
                time.sleep(wait * (attempt + 1))
            else:
                raise


# ── Sheets ────────────────────────────────────────────────────────────────────
@st.cache_resource
def _get_client():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource
def _get_ss():
    return _get_client().open_by_key(st.secrets["spreadsheet_id"])

@st.cache_resource
def _ensure_sheets():
    ss = _get_ss()
    existing = {ws.title for ws in ss.worksheets()}
    for title, headers in [
        ("POC_Registry", REGISTRY_HEADERS),
        ("Overview",     OVERVIEW_HEADERS),
        ("KPIs",         KPI_HEADERS),
        ("Action_Items", ACTION_HEADERS),
    ]:
        if title not in existing:
            ws = ss.add_worksheet(title=title, rows=1000, cols=max(len(headers), 10))
            ws.append_row(headers)

def _ensure_poc_exists():
    """Create the ASKbdap row in POC_Registry if not already there."""
    ws = _get_ss().worksheet("POC_Registry")
    ids = [r.get("POC_ID") for r in ws.get_all_records()]
    if POC_ID not in ids:
        _retry(lambda: ws.append_row(
            [POC_ID, CUSTOMER, "POC", "Planning", datetime.today().strftime("%Y-%m-%d")]
        ))
        st.cache_data.clear()

@st.cache_data(ttl=300)
def load_overview():
    for r in _get_ss().worksheet("Overview").get_all_records():
        if r.get("POC_ID") == POC_ID:
            return r
    return {}

@st.cache_data(ttl=300)
def load_kpis():
    data = _get_ss().worksheet("KPIs").get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=KPI_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        filtered = df[df["POC_ID"] == POC_ID].drop(columns=["POC_ID"])
        return filtered
    return pd.DataFrame(columns=[h for h in KPI_HEADERS if h != "POC_ID"])

@st.cache_data(ttl=300)
def load_actions():
    data = _get_ss().worksheet("Action_Items").get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=ACTION_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        filtered = df[df["POC_ID"] == POC_ID].drop(columns=["POC_ID"])
        return filtered
    return pd.DataFrame(columns=[h for h in ACTION_HEADERS if h != "POC_ID"])

def save_overview(values):
    ws = _get_ss().worksheet("Overview")
    rows = ws.get_all_values()
    row = [POC_ID] + [values.get(h, "") for h in OVERVIEW_HEADERS[1:]]
    for i, r in enumerate(rows[1:], start=2):
        if r and r[0] == POC_ID:
            _retry(lambda i=i, row=row: ws.update(f"A{i}", [row]))
            st.cache_data.clear()
            return
    _retry(lambda: ws.append_row(row))
    st.cache_data.clear()

def save_rows(sheet, headers, df):
    ws = _get_ss().worksheet(sheet)
    vals = ws.get_all_values()
    for n in reversed([i + 2 for i, r in enumerate(vals[1:]) if r and r[0] == POC_ID]):
        _retry(lambda n=n: ws.delete_rows(n))
    new_rows = [[POC_ID] + [str(row.get(h, "")) for h in headers] for _, row in df.iterrows()]
    if new_rows:
        _retry(lambda: ws.append_rows(new_rows, value_input_option="USER_ENTERED"))
    st.cache_data.clear()


def fmt_money(v):
    try:
        return f"₹{int(str(v).replace(',','').replace('₹','').replace('$','')):,}"
    except Exception:
        return str(v) if v else "—"

def parse_date(v):
    try:
        return datetime.strptime(str(v), "%Y-%m-%d").date()
    except Exception:
        return date.today()

def prep(df, cols):
    result = df.copy() if not df.empty else pd.DataFrame(columns=cols)
    for c in cols:
        if c not in result.columns:
            result[c] = ""
    return result[cols].astype(str).replace("nan", "").replace("None", "")


# ── Bootstrap ─────────────────────────────────────────────────────────────────
try:
    _retry(_ensure_sheets)
    _ensure_poc_exists()
except Exception as e:
    st.error(f"Could not connect to Google Sheets: {e}")
    st.stop()

# ── Show pending save notification (set before rerun, shown once after) ───────
if "notify" in st.session_state:
    msg = st.session_state.pop("notify")
    st.toast(msg, icon="✅")

# ── Load data ─────────────────────────────────────────────────────────────────
ov      = load_overview()
kpis    = load_kpis()
actions = load_actions()

poc_status = ov.get("Status", "Planning")
DISPLAY_KPI = [h for h in KPI_HEADERS if h != "POC_ID"]
DISPLAY_ACT = [h for h in ACTION_HEADERS if h != "POC_ID"]

kpis_met   = int((kpis["Status"] == "Met").sum())              if not kpis.empty    else 0
kpis_total = len(kpis)
acts_open  = int((actions["Status"].isin(["Open","In Progress"])).sum()) if not actions.empty else 0

today = date.today()
overdue = 0
if not actions.empty and "Due Date" in actions.columns:
    for _, row in actions.iterrows():
        try:
            if datetime.strptime(str(row["Due Date"]), "%Y-%m-%d").date() < today and str(row.get("Status","")) != "Complete":
                overdue += 1
        except Exception:
            pass

# ── Header ────────────────────────────────────────────────────────────────────
hdr_l, hdr_r = st.columns([5, 1])
with hdr_l:
    st.markdown(
        '<div style="padding:2px 0 10px;">'
        '<div style="font-size:1.6rem;font-weight:700;color:#0F172A;line-height:1.2;">ASKbdap × Snowflake</div>'
        '<div style="font-size:.84rem;color:#64748B;margin-top:3px;">POC Tracker</div>'
        '</div>',
        unsafe_allow_html=True,
    )
with hdr_r:
    st.markdown("<div style='margin-top:20px;'></div>", unsafe_allow_html=True)
    if st.button("↺ Refresh", help="Reload from Google Sheets"):
        with st.spinner("Refreshing…"):
            st.cache_data.clear()
        st.rerun()

# ── Metrics ───────────────────────────────────────────────────────────────────
m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.metric("Status", STATUS_BADGE.get(poc_status, poc_status or "—"))
with m2:
    st.metric("KPIs Met", f"{kpis_met} / {kpis_total}")
with m3:
    pct = round(kpis_met / kpis_total * 100) if kpis_total else 0
    st.metric("KPI Progress", f"{pct}%")
with m4:
    st.metric("Actions Open", str(acts_open),
              delta=f"{overdue} overdue" if overdue else None,
              delta_color="inverse" if overdue else "normal")
with m5:
    budget = fmt_money(ov.get("POC Budget ($)", ""))
    spent  = fmt_money(ov.get("Confirmed Spend ($)", ""))
    st.metric("Budget", budget,
              delta=f"Spent {spent}" if ov.get("Confirmed Spend ($)") else None,
              delta_color="off")

if kpis_total > 0:
    st.progress(kpis_met / kpis_total, text=f"{kpis_met} of {kpis_total} KPIs met")

st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_kpi, tab_actions, tab_details = st.tabs(["📊  KPIs", "✅  Action Plan", "📋  POC Details"])


# ─── KPIs ─────────────────────────────────────────────────────────────────────
with tab_kpi:
    kpi_df = prep(kpis, DISPLAY_KPI)

    with st.container(border=True):
        st.markdown("**Key Performance Indicators**")
        st.caption("Edit any cell directly · Use the ＋ row at the bottom to add new KPIs · Click **Save KPIs** when done.")

        edited_kpis = st.data_editor(
            kpi_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "KPI":           st.column_config.TextColumn("KPI Description", width="large"),
                "Target":        st.column_config.TextColumn("Target",          width="small"),
                "Current Value": st.column_config.TextColumn("Current",         width="small"),
                "Unit":          st.column_config.TextColumn("Unit",            width="small"),
                "Status":        st.column_config.SelectboxColumn("Status",     options=KPI_STATUSES, width="medium"),
                "Notes":         st.column_config.TextColumn("Notes",           width="large"),
            },
            key="kpi_editor",
        )

        if st.button("💾  Save KPIs", type="primary", key="save_kpis"):
            clean = edited_kpis.dropna(how="all")
            clean = clean[clean["KPI"].astype(str).str.strip() != ""]
            with st.spinner("Saving KPIs…"):
                save_rows("KPIs", DISPLAY_KPI, clean)
            st.session_state["notify"] = "KPIs saved successfully"
            st.rerun()


# ─── Action Plan ──────────────────────────────────────────────────────────────
with tab_actions:
    if overdue:
        st.warning(f"⚠️  {overdue} action item{'s are' if overdue > 1 else ' is'} past due date.")
    if not actions.empty:
        blocked = int((actions["Status"] == "Blocked").sum())
        if blocked:
            st.error(f"🚫  {blocked} action item{'s are' if blocked > 1 else ' is'} blocked.")

    act_df = prep(actions, DISPLAY_ACT)

    with st.container(border=True):
        st.markdown("**Action Plan**")
        st.caption("Edit any cell directly · Use the ＋ row at the bottom to add new actions · Click **Save Actions** when done.")

        edited_actions = st.data_editor(
            act_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Action":   st.column_config.TextColumn("Action Item",                        width="large"),
                "Assignee": st.column_config.TextColumn("Assignee",                          width="medium"),
                "Category": st.column_config.SelectboxColumn("Category", options=ACTION_CATS, width="medium"),
                "Due Date": st.column_config.TextColumn("Due Date (YYYY-MM-DD)",             width="medium"),
                "Status":   st.column_config.SelectboxColumn("Status",   options=ACTION_STATUSES, width="medium"),
                "Notes":    st.column_config.TextColumn("Notes",                             width="large"),
            },
            key="action_editor",
        )

        if st.button("💾  Save Actions", type="primary", key="save_actions"):
            clean = edited_actions.dropna(how="all")
            clean = clean[clean["Action"].astype(str).str.strip() != ""]
            with st.spinner("Saving actions…"):
                save_rows("Action_Items", DISPLAY_ACT, clean)
            st.session_state["notify"] = "Action plan saved successfully"
            st.rerun()


# ─── POC Details ─────────────────────────────────────────────────────────────
with tab_details:
    st.markdown("**POC Details** — all fields are editable")

    with st.form("details_form"):
        col_a, col_b = st.columns(2)

        with col_a:
            with st.container(border=True):
                st.markdown("**Contacts**")
                champion      = st.text_input("Technical Champion",          value=ov.get("Technical Champion", ""))
                champ_title   = st.text_input("Technical Champion Title",    value=ov.get("Technical Champion Title", ""))
                sponsor       = st.text_input("Exec Business Sponsor",       value=ov.get("Exec Business Sponsor", ""))
                sponsor_title = st.text_input("Exec Business Sponsor Title", value=ov.get("Exec Business Sponsor Title", ""))
                participants  = st.text_area(
                    "Customer Participants", value=ov.get("Customer Participants", ""),
                    height=80, placeholder="Jane Smith (Data Engineer), John Doe (IT Lead)…",
                )

        with col_b:
            with st.container(border=True):
                st.markdown("**Engagement**")
                bu = st.selectbox(
                    "Business Unit", BUSINESS_UNITS,
                    index=BUSINESS_UNITS.index(ov["Business Unit"]) if ov.get("Business Unit") in BUSINESS_UNITS else 0,
                )
                cloud = st.selectbox(
                    "Cloud Environment", CLOUD_OPTIONS,
                    index=CLOUD_OPTIONS.index(ov["Cloud Environment"]) if ov.get("Cloud Environment") in CLOUD_OPTIONS else 0,
                )
                platform   = st.text_input("Current Data Platform", value=ov.get("Current Data Platform", ""))
                volume     = st.text_input("Estimated Data Volume",  value=ov.get("Data Volume", ""))
                saved_c    = [x.strip() for x in ov.get("Compliance Requirements", "").split(",") if x.strip()]
                compliance = st.multiselect("Compliance", COMPLIANCE_OPTIONS,
                                            default=[x for x in saved_c if x in COMPLIANCE_OPTIONS])

        with st.container(border=True):
            st.markdown("**Timeline & Status**")
            t1, t2, t3 = st.columns(3)
            with t1:
                start_date = st.date_input("Start Date",         value=parse_date(ov.get("POC Start Date", "")))
            with t2:
                end_date   = st.date_input("Target Completion",  value=parse_date(ov.get("Target Completion Date", "")))
            with t3:
                status = st.selectbox(
                    "POC Status", STATUS_OPTIONS,
                    index=STATUS_OPTIONS.index(ov["Status"]) if ov.get("Status") in STATUS_OPTIONS else 0,
                )

        with st.container(border=True):
            st.markdown("**Objectives & Success Criteria**")
            objective = st.text_area("POC Objective",              value=ov.get("POC Objective", ""),              height=80)
            o1, o2 = st.columns(2)
            with o1:
                tc = st.text_area("Technical Success Criteria",    value=ov.get("Technical Success Criteria", ""), height=90)
            with o2:
                bc = st.text_area("Business Success Criteria",     value=ov.get("Business Success Criteria", ""),  height=90)

        with st.container(border=True):
            st.markdown("**Budget**")
            b1, b2 = st.columns(2)
            with b1:
                poc_budget = st.text_input("POC Budget ($)",      value=ov.get("POC Budget ($)", ""),      placeholder="e.g. 50000")
            with b2:
                conf_spend = st.text_input("Confirmed Spend ($)", value=ov.get("Confirmed Spend ($)", ""), placeholder="0")

        if st.form_submit_button("💾  Save Details", type="primary", use_container_width=True):
            with st.spinner("Saving POC details…"):
                save_overview({
                    "Technical Champion":          champion,
                    "Technical Champion Title":    champ_title,
                    "Exec Business Sponsor":       sponsor,
                    "Exec Business Sponsor Title": sponsor_title,
                    "Customer Participants":       participants,
                    "Business Unit":               bu,
                    "Cloud Environment":           cloud,
                    "Current Data Platform":       platform,
                    "Data Volume":                 volume,
                    "Compliance Requirements":     ", ".join(compliance),
                    "POC Start Date":              str(start_date),
                    "Target Completion Date":      str(end_date),
                    "Status":                      status,
                    "POC Objective":               objective,
                    "Technical Success Criteria":  tc,
                    "Business Success Criteria":   bc,
                    "POC Budget ($)":              poc_budget,
                    "Confirmed Spend ($)":         conf_spend,
                })
            st.session_state["notify"] = "POC details saved successfully"
            st.rerun()
