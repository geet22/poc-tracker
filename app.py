import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date
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

st.set_page_config(page_title="POC Tracker", layout="wide", initial_sidebar_state="expanded")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

REGISTRY_HEADERS = ["POC_ID", "Customer", "Engagement", "Status", "Created"]
OVERVIEW_HEADERS = [
    "POC_ID",
    "Technical Champion", "Technical Champion Title",
    "Exec Business Sponsor", "Exec Business Sponsor Title",
    "Customer Participants",
    "Business Unit",
    "Cloud Environment", "Current Data Platform", "Data Volume",
    "Compliance Requirements",
    "POC Start Date", "Target Completion Date", "Status",
    "POC Objective", "Technical Success Criteria", "Business Success Criteria",
    "POC Budget ($)", "Confirmed Spend ($)",
]
KPI_HEADERS    = ["POC_ID", "KPI", "Target", "Current Value", "Unit", "Status", "Notes"]
ACTION_HEADERS = ["POC_ID", "Action", "Assignee", "Category", "Due Date", "Status", "Notes"]

BUSINESS_UNITS     = ["Diagnostics", "Drug Development (Biopharma Solutions)", "Genomics", "Technology Solutions", "Enterprise / Cross-BU"]
COMPLIANCE_OPTIONS = ["HIPAA", "21 CFR Part 11", "GxP", "SOC 2", "CLIA", "GDPR"]
STATUS_OPTIONS     = ["Planning", "Active", "At Risk", "Completed — Won", "Completed — Lost"]
KPI_STATUSES       = ["On Track", "At Risk", "Met", "Not Started"]
ACTION_CATS        = ["Technical", "Business", "Compliance", "Training", "Executive"]
ACTION_STATUSES    = ["Open", "In Progress", "Complete", "Blocked"]
CLOUD_OPTIONS      = ["AWS", "Azure", "Multi-Cloud"]
STATUS_EMOJI       = {"Planning":"🔵","Active":"🟢","At Risk":"🟡","Completed — Won":"✅","Completed — Lost":"🔴"}

# ── Sheets ────────────────────────────────────────────────────────────────────
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
        ("POC_Registry", REGISTRY_HEADERS),
        ("Overview",     OVERVIEW_HEADERS),
        ("KPIs",         KPI_HEADERS),
        ("Action_Items", ACTION_HEADERS),
    ]:
        if title not in existing:
            ws = ss.add_worksheet(title=title, rows=1000, cols=max(len(headers), 10))
            ws.append_row(headers)
        else:
            ws = ss.worksheet(title)
            if not ws.row_values(1):
                ws.insert_row(headers, 1)

@st.cache_data(ttl=300)
def load_registry():
    data = get_ss().worksheet("POC_Registry").get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=REGISTRY_HEADERS)

@st.cache_data(ttl=300)
def load_overview(poc_id):
    for r in get_ss().worksheet("Overview").get_all_records():
        if r.get("POC_ID") == poc_id:
            return r
    return {}

@st.cache_data(ttl=300)
def load_kpis(poc_id):
    data = get_ss().worksheet("KPIs").get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=KPI_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        return df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

@st.cache_data(ttl=300)
def load_actions(poc_id):
    data = get_ss().worksheet("Action_Items").get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=ACTION_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        return df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

def save_overview(poc_id, values):
    ws = get_ss().worksheet("Overview")
    rows = ws.get_all_values()
    row = [poc_id] + [values.get(h, "") for h in OVERVIEW_HEADERS[1:]]
    for i, r in enumerate(rows[1:], start=2):
        if r and r[0] == poc_id:
            ws.delete_rows(i); ws.insert_row(row, i)
            st.cache_data.clear(); return
    ws.append_row(row); st.cache_data.clear()

def create_poc(poc_id, customer, engagement, status):
    get_ss().worksheet("POC_Registry").append_row(
        [poc_id, customer, engagement, status, datetime.today().strftime("%Y-%m-%d")]
    )
    st.cache_data.clear()

def save_all_rows(sheet, poc_id, headers, df):
    ws = get_ss().worksheet(sheet)
    vals = ws.get_all_values()
    for n in reversed([i+2 for i, r in enumerate(vals[1:]) if r and r[0] == poc_id]):
        ws.delete_rows(n)
    for _, row in df.iterrows():
        ws.append_row([poc_id]+[str(row.get(h,"")) for h in headers], value_input_option="USER_ENTERED")
    st.cache_data.clear()

# ── Bootstrap ─────────────────────────────────────────────────────────────────
try:
    _retry(ensure_sheets)
except Exception as e:
    st.error(f"Could not connect to Google Sheets: {e}"); st.stop()

# ── Styles ────────────────────────────────────────────────────────────────────
st.html("""
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*, html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
#MainMenu, footer, header { visibility:hidden; height:0; }

/* Page background */
[data-testid="stAppViewContainer"] > .main { background:#EBEDF0 !important; }
.block-container { padding:0.6rem 1.4rem 2rem !important; max-width:1300px; }

/* Vertical rhythm */
[data-testid="stVerticalBlock"] { gap:0.45rem !important; }

/* Bordered containers become white cards */
[data-testid="stVerticalBlockBorderWrapper"] {
    background:#fff !important;
    border:1px solid #DDE1E7 !important;
    border-radius:10px !important;
    box-shadow:0 1px 3px rgba(0,0,0,.06) !important;
    padding:0 !important;
}
[data-testid="stVerticalBlockBorderWrapper"] > div > [data-testid="stVerticalBlock"] {
    padding:16px 20px !important; gap:0.6rem !important;
}

/* Sidebar */
section[data-testid="stSidebar"] { background:#0F172A !important; border-right:1px solid #1E293B !important; width:230px !important; }
section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] small, section[data-testid="stSidebar"] div,
section[data-testid="stSidebar"] label { color:#94A3B8 !important; }
section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3, section[data-testid="stSidebar"] strong { color:#F1F5F9 !important; }
section[data-testid="stSidebar"] hr { border-color:#1E293B !important; margin:8px 0 !important; }
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
    background:rgba(255,255,255,.03) !important; border:1px solid #1E293B !important;
    border-radius:6px !important; padding:7px 12px !important; margin-bottom:3px !important;
    font-size:.8rem !important; color:#CBD5E1 !important;
}
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {
    background:rgba(41,181,232,.08) !important; border-color:#29B5E8 !important;
}
section[data-testid="stSidebar"] input[type="text"] {
    background:rgba(255,255,255,.05) !important; border:1px solid #1E293B !important;
    color:#F1F5F9 !important; border-radius:6px !important; font-size:.82rem !important;
}
section[data-testid="stSidebar"] button[kind="primaryFormSubmit"],
section[data-testid="stSidebar"] button[kind="primary"] {
    background:#29B5E8 !important; color:#0F172A !important; border:none !important;
    border-radius:6px !important; font-weight:600 !important; font-size:.78rem !important;
}

/* Buttons */
.stButton > button[kind="primary"] {
    background:#29B5E8 !important; color:#0F172A !important; border:none !important;
    border-radius:6px !important; font-weight:600 !important; font-size:.75rem !important;
    padding:5px 14px !important; height:30px !important; line-height:1 !important;
}
.stButton > button {
    border-radius:6px !important; font-size:.75rem !important; font-weight:500 !important;
    height:30px !important; padding:5px 12px !important;
}

/* Inputs */
.stTextInput input, .stTextArea textarea, .stSelectbox > div > div {
    border-radius:6px !important; border:1px solid #DDE1E7 !important;
    font-size:.83rem !important; background:#FAFBFC !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color:#29B5E8 !important; box-shadow:0 0 0 3px rgba(41,181,232,.1) !important;
    background:#fff !important;
}

/* Expander */
[data-testid="stExpander"] summary {
    font-size:.8rem !important; font-weight:600 !important; color:#374151 !important;
    background:#F9FAFB !important; border-radius:7px !important; padding:10px 14px !important;
    border:1px solid #DDE1E7 !important;
}
[data-testid="stExpander"] { border:none !important; background:transparent !important; }

/* Dividers */
hr { border-color:#EEF0F3 !important; margin:4px 0 !important; }

/* Stat row */
.stat-row { display:flex; gap:8px; margin:4px 0 8px; flex-wrap:wrap; }
.stat-card {
    flex:1 1 110px; background:#fff; border:1px solid #DDE1E7;
    border-radius:8px; padding:11px 15px; min-width:105px;
    box-shadow:0 1px 2px rgba(0,0,0,.04);
}
.stat-label { font-size:.6rem; font-weight:700; letter-spacing:.07em; text-transform:uppercase; color:#9CA3AF; margin-bottom:4px; }
.stat-value { font-size:1.35rem; font-weight:700; color:#111827; line-height:1.1; }
.stat-value-sm { font-size:.95rem; font-weight:700; color:#111827; }
.stat-sub { font-size:.65rem; color:#9CA3AF; margin-top:2px; }

/* Pills */
.pill { display:inline-block; font-size:.68rem; font-weight:600; padding:2px 9px; border-radius:20px; letter-spacing:.02em; }
.p-green { background:#D1FAE5; color:#065F46; }
.p-blue  { background:#DBEAFE; color:#1E40AF; }
.p-amber { background:#FEF3C7; color:#92400E; }
.p-red   { background:#FEE2E2; color:#991B1B; }
.p-gray  { background:#F3F4F6; color:#374151; }

/* Card heading row */
.ch { display:flex; justify-content:space-between; align-items:center; margin-bottom:2px; }
.ch-title { font-size:.78rem; font-weight:700; color:#111827; letter-spacing:-.01em; }
.ch-sub   { font-size:.7rem; color:#9CA3AF; }
</style>
""")

# ── Helpers ───────────────────────────────────────────────────────────────────
_PILL = {
    "Planning":"p-gray","Active":"p-blue","At Risk":"p-amber",
    "Completed — Won":"p-green","Completed — Lost":"p-red",
    "Met":"p-green","On Track":"p-blue","Not Started":"p-gray",
    "Complete":"p-green","In Progress":"p-blue","Open":"p-gray","Blocked":"p-red",
}
def pill(t): return f'<span class="pill {_PILL.get(t,"p-gray")}">{t}</span>'

def fmt_money(v):
    try: return f"${int(str(v).replace(',','').replace('$','')):,}"
    except: return f"${v}" if v else "—"

def prep_df(df, cols):
    """Ensure all cols exist, cast to str, append blank rows for editing."""
    for c in cols:
        if c not in df.columns: df[c] = ""
    df = df[cols].copy()
    for c in cols:
        df[c] = df[c].astype(str).replace("nan", "")
    blank = pd.DataFrame([{c: "" for c in cols}] * 3)
    return pd.concat([df, blank], ignore_index=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
registry = load_registry()

with st.sidebar:
    st.markdown("**POC Tracker**")
    st.divider()

    for k, d in [("editor_name",""),("editor_role","Snowflake Team")]:
        if k not in st.session_state: st.session_state[k] = d

    st.markdown("<small>Editing as</small>", unsafe_allow_html=True)
    st.session_state["editor_name"] = st.text_input(
        "name", value=st.session_state["editor_name"],
        placeholder="Your name", label_visibility="collapsed"
    )
    st.session_state["editor_role"] = st.radio(
        "role", ["Snowflake Team","Customer"],
        index=0 if st.session_state["editor_role"]=="Snowflake Team" else 1,
        horizontal=True, label_visibility="collapsed"
    )
    st.divider()

    poc_options = registry["POC_ID"].tolist() if not registry.empty else []
    if poc_options:
        labels = []
        for _, r in registry.iterrows():
            e = STATUS_EMOJI.get(r.get("Status",""),"⚪")
            eng = r.get("Engagement","")
            labels.append(f"{e} {r['Customer']}" + (f" · {eng}" if eng else ""))
        idx = st.radio("Engagements", range(len(poc_options)),
                       format_func=lambda i: labels[i], key="poc_radio",
                       label_visibility="collapsed")
        active_poc_id = poc_options[idx]
        active_row    = registry[registry["POC_ID"]==active_poc_id].iloc[0]
    else:
        st.info("No POCs yet.")
        active_poc_id = None
        active_row    = {}

    st.divider()
    with st.expander("＋ New POC", expanded=not poc_options):
        with st.form("new_poc"):
            c = st.text_input("Customer *")
            e = st.text_input("Engagement / BU")
            s = st.selectbox("Status", STATUS_OPTIONS)
            if st.form_submit_button("Create", type="primary"):
                if c.strip():
                    slug = (c.strip()+"-"+e.strip()).lower().replace(" ","-").replace("·","").replace("/","")[:32]
                    create_poc(slug, c.strip(), e.strip(), s)
                    st.success(f"Created"); st.rerun()
                else:
                    st.warning("Customer name required.")

# ── Guard ─────────────────────────────────────────────────────────────────────
if not active_poc_id:
    st.markdown("## POC Tracker")
    st.info("Create your first engagement in the sidebar.")
    st.stop()

# ── Data ──────────────────────────────────────────────────────────────────────
ov      = load_overview(active_poc_id)
kpis    = load_kpis(active_poc_id)
actions = load_actions(active_poc_id)

customer_name   = active_row.get("Customer", active_poc_id)
engagement_name = active_row.get("Engagement","")
poc_status      = ov.get("Status", active_row.get("Status",""))

kpis_met   = int((kpis["Status"]=="Met").sum())     if not kpis.empty    else 0
kpis_total = len(kpis)
acts_open  = int((actions["Status"]=="Open").sum()) if not actions.empty else 0
acts_total = len(actions)

# ── Header ────────────────────────────────────────────────────────────────────
h1, h2 = st.columns([3, 1])
with h1:
    st.markdown(f"### {customer_name} × Snowflake")
    if engagement_name:
        st.caption(engagement_name)
with h2:
    if st.button("↺ Refresh", key="refresh"):
        st.cache_data.clear(); st.rerun()

st.html(f"""
<div class="stat-row">
  <div class="stat-card">
    <div class="stat-label">Status</div>
    <div style="margin-top:4px">{pill(poc_status) if poc_status else "—"}</div>
  </div>
  <div class="stat-card">
    <div class="stat-label">KPIs Met</div>
    <div class="stat-value">{kpis_met}<span style="font-size:.9rem;color:#D1D5DB;font-weight:500"> /{kpis_total}</span></div>
  </div>
  <div class="stat-card">
    <div class="stat-label">Actions Open</div>
    <div class="stat-value">{acts_open}<span style="font-size:.9rem;color:#D1D5DB;font-weight:500"> /{acts_total}</span></div>
  </div>
  <div class="stat-card">
    <div class="stat-label">Budget</div>
    <div class="stat-value-sm">{fmt_money(ov.get("POC Budget ($)",""))}</div>
    <div class="stat-sub">Spent {fmt_money(ov.get("Confirmed Spend ($)",""))}</div>
  </div>
</div>
""")

# ── KPIs ──────────────────────────────────────────────────────────────────────
DISPLAY_KPI = [h for h in KPI_HEADERS if h != "POC_ID"]
kpi_df = prep_df(kpis.copy() if not kpis.empty else pd.DataFrame(columns=DISPLAY_KPI), DISPLAY_KPI)

with st.container(border=True):
    st.html('<div class="ch"><span class="ch-title">Key Performance Indicators</span><span class="ch-sub">Click any cell to edit · Save when done</span></div>')
    edited_kpis = st.data_editor(
        kpi_df, num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "KPI":           st.column_config.TextColumn("KPI",     width="large"),
            "Target":        st.column_config.TextColumn("Target",  width="small"),
            "Current Value": st.column_config.TextColumn("Current", width="small"),
            "Unit":          st.column_config.TextColumn("Unit",    width="small"),
            "Status":        st.column_config.SelectboxColumn("Status", options=KPI_STATUSES, width="medium"),
            "Notes":         st.column_config.TextColumn("Notes",   width="large"),
        },
        key="kpi_editor",
    )
    if st.button("Save KPIs", type="primary", key="save_kpis"):
        clean = edited_kpis.dropna(how="all")
        clean = clean[clean["KPI"].astype(str).str.strip() != ""]
        save_all_rows("KPIs", active_poc_id, DISPLAY_KPI, clean)
        st.success("Saved."); st.rerun()

# ── Action Plan ───────────────────────────────────────────────────────────────
DISPLAY_ACT = [h for h in ACTION_HEADERS if h != "POC_ID"]
act_df = prep_df(actions.copy() if not actions.empty else pd.DataFrame(columns=DISPLAY_ACT), DISPLAY_ACT)

with st.container(border=True):
    st.html('<div class="ch"><span class="ch-title">Action Plan</span><span class="ch-sub">Click any cell to edit · Save when done</span></div>')
    edited_actions = st.data_editor(
        act_df, num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "Action":   st.column_config.TextColumn("Action",   width="large"),
            "Assignee": st.column_config.TextColumn("Assignee", width="medium"),
            "Category": st.column_config.SelectboxColumn("Category", options=ACTION_CATS, width="medium"),
            "Due Date": st.column_config.TextColumn("Due Date (YYYY-MM-DD)", width="medium"),
            "Status":   st.column_config.SelectboxColumn("Status", options=ACTION_STATUSES, width="medium"),
            "Notes":    st.column_config.TextColumn("Notes",    width="large"),
        },
        key="action_editor",
    )
    if st.button("Save Actions", type="primary", key="save_actions"):
        clean = edited_actions.dropna(how="all")
        clean = clean[clean["Action"].astype(str).str.strip() != ""]
        save_all_rows("Action_Items", active_poc_id, DISPLAY_ACT, clean)
        st.success("Saved."); st.rerun()

# ── POC Details ───────────────────────────────────────────────────────────────
with st.expander("POC Details", expanded=False):
    editor_name = st.session_state.get("editor_name","")
    editor_role = st.session_state.get("editor_role","Snowflake Team")

    with st.form("overview_form"):
        st.markdown("**Contacts**")
        c1, c2 = st.columns(2)
        with c1:
            champion      = st.text_input("Technical Champion",          value=ov.get("Technical Champion",""))
            champ_title   = st.text_input("Technical Champion Title",    value=ov.get("Technical Champion Title",""))
        with c2:
            sponsor       = st.text_input("Exec Business Sponsor",       value=ov.get("Exec Business Sponsor",""))
            sponsor_title = st.text_input("Exec Business Sponsor Title", value=ov.get("Exec Business Sponsor Title",""))
        participants = st.text_area("Customer Participants", value=ov.get("Customer Participants",""),
                                    height=65, placeholder="Jane Smith (Data Engineer), John Doe (IT Lead)…")

        st.divider()
        st.markdown("**Engagement**")
        e1, e2 = st.columns(2)
        with e1:
            bu = st.selectbox("Business Unit", BUSINESS_UNITS,
                              index=BUSINESS_UNITS.index(ov["Business Unit"]) if ov.get("Business Unit") in BUSINESS_UNITS else 0)
        with e2:
            cloud    = st.selectbox("Cloud Environment", CLOUD_OPTIONS,
                                    index=CLOUD_OPTIONS.index(ov["Cloud Environment"]) if ov.get("Cloud Environment") in CLOUD_OPTIONS else 0)
            platform = st.text_input("Current Data Platform", value=ov.get("Current Data Platform",""))
            volume   = st.text_input("Estimated Data Volume",  value=ov.get("Data Volume",""))
        saved_c = [x.strip() for x in ov.get("Compliance Requirements","").split(",") if x.strip()]
        compliance = st.multiselect("Compliance", COMPLIANCE_OPTIONS,
                                    default=[x for x in saved_c if x in COMPLIANCE_OPTIONS])

        st.divider()
        st.markdown("**Timeline & Status**")
        t1, t2, t3 = st.columns(3)
        with t1:
            try:    sv = datetime.strptime(ov.get("POC Start Date",""), "%Y-%m-%d").date()
            except: sv = date.today()
            start_date = st.date_input("Start Date", value=sv)
        with t2:
            try:    ev = datetime.strptime(ov.get("Target Completion Date",""), "%Y-%m-%d").date()
            except: ev = date.today()
            end_date = st.date_input("Target Completion", value=ev)
        with t3:
            status = st.selectbox("POC Status", STATUS_OPTIONS,
                                  index=STATUS_OPTIONS.index(ov["Status"]) if ov.get("Status") in STATUS_OPTIONS else 0)

        st.divider()
        st.markdown("**Objectives**")
        objective = st.text_area("POC Objective", value=ov.get("POC Objective",""), height=70)
        o1, o2 = st.columns(2)
        with o1: tc = st.text_area("Technical Success Criteria", value=ov.get("Technical Success Criteria",""), height=80)
        with o2: bc = st.text_area("Business Success Criteria",  value=ov.get("Business Success Criteria",""),  height=80)

        st.divider()
        st.markdown("**Budget**")
        b1, b2 = st.columns(2)
        with b1: poc_budget = st.text_input("POC Budget ($)",      value=ov.get("POC Budget ($)",""))
        with b2: conf_spend = st.text_input("Confirmed Spend ($)", value=ov.get("Confirmed Spend ($)",""))

        if st.form_submit_button("Save Details", type="primary", use_container_width=True):
            save_overview(active_poc_id, {
                "Technical Champion": champion, "Technical Champion Title": champ_title,
                "Exec Business Sponsor": sponsor, "Exec Business Sponsor Title": sponsor_title,
                "Customer Participants": participants, "Business Unit": bu,
                "Cloud Environment": cloud, "Current Data Platform": platform, "Data Volume": volume,
                "Compliance Requirements": ", ".join(compliance),
                "POC Start Date": str(start_date), "Target Completion Date": str(end_date),
                "Status": status, "POC Objective": objective,
                "Technical Success Criteria": tc, "Business Success Criteria": bc,
                "POC Budget ($)": poc_budget, "Confirmed Spend ($)": conf_spend,
            })
            st.success(f"Saved by {editor_name or editor_role}."); st.rerun()
