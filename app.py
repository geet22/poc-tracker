import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date
import time
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

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

# ── Headers ───────────────────────────────────────────────────────────────────
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

STATUS_EMOJI = {
    "Planning": "🔵", "Active": "🟢", "At Risk": "🟡",
    "Completed — Won": "✅", "Completed — Lost": "🔴",
}

# ── Google Sheets ─────────────────────────────────────────────────────────────
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

# ── Load ──────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_registry() -> pd.DataFrame:
    data = get_ss().worksheet("POC_Registry").get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=REGISTRY_HEADERS)

@st.cache_data(ttl=300)
def load_overview(poc_id: str) -> dict:
    for r in get_ss().worksheet("Overview").get_all_records():
        if r.get("POC_ID") == poc_id:
            return r
    return {}

@st.cache_data(ttl=300)
def load_kpis(poc_id: str) -> pd.DataFrame:
    data = get_ss().worksheet("KPIs").get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=KPI_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        return df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

@st.cache_data(ttl=300)
def load_actions(poc_id: str) -> pd.DataFrame:
    data = get_ss().worksheet("Action_Items").get_all_records()
    df = pd.DataFrame(data) if data else pd.DataFrame(columns=ACTION_HEADERS)
    if not df.empty and "POC_ID" in df.columns:
        return df[df["POC_ID"] == poc_id].drop(columns=["POC_ID"])
    return df

# ── Save ──────────────────────────────────────────────────────────────────────
def save_overview(poc_id: str, values: dict):
    ws = get_ss().worksheet("Overview")
    all_rows = ws.get_all_values()
    row = [poc_id] + [values.get(h, "") for h in OVERVIEW_HEADERS[1:]]
    for i, r in enumerate(all_rows[1:], start=2):
        if r and r[0] == poc_id:
            ws.delete_rows(i)
            ws.insert_row(row, i)
            st.cache_data.clear()
            return
    ws.append_row(row)
    st.cache_data.clear()

def create_poc(poc_id, customer, engagement, status):
    get_ss().worksheet("POC_Registry").append_row(
        [poc_id, customer, engagement, status, datetime.today().strftime("%Y-%m-%d")]
    )
    st.cache_data.clear()

def save_all_rows(sheet: str, poc_id: str, headers: list, df: pd.DataFrame):
    ws = get_ss().worksheet(sheet)
    all_vals = ws.get_all_values()
    for row_num in reversed([i + 2 for i, r in enumerate(all_vals[1:]) if r and r[0] == poc_id]):
        ws.delete_rows(row_num)
    for _, row in df.iterrows():
        ws.append_row([poc_id] + [str(row.get(h, "")) for h in headers], value_input_option="USER_ENTERED")
    st.cache_data.clear()

# ── Bootstrap ─────────────────────────────────────────────────────────────────
try:
    _retry(ensure_sheets)
except Exception as e:
    st.error(f"Could not connect to Google Sheets: {e}")
    st.stop()

# ── Styles ────────────────────────────────────────────────────────────────────
st.html("""
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
html, body, [class*="css"], .stApp { font-family: 'Inter', sans-serif !important; }
#MainMenu, footer, header { visibility: hidden; height: 0; }
.block-container { padding: 0.75rem 2rem 1.5rem !important; }

/* Tighten vertical gaps between components */
div[data-testid="stVerticalBlock"] { gap: 0.35rem !important; }
div[data-testid="stVerticalBlockBorderWrapper"] { gap: 0.35rem !important; }

/* Sidebar */
section[data-testid="stSidebar"] { background:#0B1929 !important; border-right:1px solid #1E3A5F !important; }
section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] small, section[data-testid="stSidebar"] div,
section[data-testid="stSidebar"] .stRadio label { color:#CBD5E0 !important; }
section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color:#fff !important; }
section[data-testid="stSidebar"] hr { border-color:#1E3A5F !important; }
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
    background:rgba(255,255,255,.04) !important; border:1px solid #1E3A5F !important;
    border-radius:6px !important; padding:8px 12px !important; margin-bottom:4px !important; font-size:.82rem !important;
}
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {
    background:rgba(41,181,232,.1) !important; border-color:#29B5E8 !important;
}
section[data-testid="stSidebar"] input[type="text"], section[data-testid="stSidebar"] select,
section[data-testid="stSidebar"] textarea {
    background:rgba(255,255,255,.06) !important; border:1px solid #1E3A5F !important;
    color:#fff !important; border-radius:6px !important;
}
section[data-testid="stSidebar"] button[kind="primaryFormSubmit"],
section[data-testid="stSidebar"] button[kind="primary"] {
    background:#29B5E8 !important; color:#000 !important; border:none !important;
    border-radius:6px !important; font-weight:600 !important;
}

/* Buttons */
.stButton > button[kind="primary"] {
    background:#29B5E8 !important; color:#000 !important; border:none !important;
    border-radius:6px !important; font-weight:600 !important; font-size:.78rem !important; padding:6px 16px !important;
}
.stButton > button { border-radius:6px !important; font-size:.78rem !important; font-weight:500 !important; }

/* Inputs */
.stTextInput input, .stTextArea textarea {
    border-radius:6px !important; border:1px solid #E2E8F0 !important; font-size:.85rem !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color:#29B5E8 !important; box-shadow:0 0 0 3px rgba(41,181,232,.12) !important;
}

/* Expander */
.streamlit-expanderHeader {
    font-size:.82rem !important; font-weight:600 !important;
    background:#F8FAFC !important; border-radius:6px !important; border:1px solid #E2E8F0 !important;
    padding:8px 14px !important;
}
hr { border-color:#F1F5F9 !important; margin:0.4rem 0 !important; }

/* Compact stat cards */
.stat-row { display:flex; gap:10px; margin:0.5rem 0 0.75rem 0; flex-wrap:wrap; }
.stat-card {
    flex:1 1 120px; background:#fff; border:1px solid #E2E8F0;
    border-radius:8px; padding:12px 16px; min-width:120px;
    box-shadow:0 1px 3px rgba(0,0,0,.04);
}
.stat-label { font-size:.63rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; color:#94A3B8; margin-bottom:5px; }
.stat-value { font-size:1.4rem; font-weight:700; color:#1A202C; line-height:1.1; }
.stat-value-sm { font-size:1rem; font-weight:700; color:#1A202C; }
.stat-sub { font-size:.68rem; color:#A0AEC0; margin-top:3px; }

/* Section label */
.section-label {
    font-size:.63rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase;
    color:#94A3B8; padding-bottom:.4rem; border-bottom:2px solid #F1F5F9; margin-bottom:0;
}

/* Status pills */
.pill { display:inline-block; font-size:.7rem; font-weight:600; padding:2px 10px; border-radius:20px; }
.p-green { background:#DCFCE7; color:#166534; }
.p-blue  { background:#DBEAFE; color:#1E40AF; }
.p-amber { background:#FEF3C7; color:#92400E; }
.p-red   { background:#FEE2E2; color:#991B1B; }
.p-gray  { background:#F1F5F9; color:#475569; }
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

def ag_table(df: pd.DataFrame, col_defs: dict, key: str) -> pd.DataFrame:
    """Render a compact editable AG Grid table and return the edited DataFrame."""
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        editable=True, resizable=True, sortable=False,
        filter=False, suppressMenu=True, minWidth=60,
    )
    for field, opts in col_defs.items():
        gb.configure_column(field, **opts)
    gb.configure_grid_options(
        rowHeight=28, headerHeight=32,
        suppressMovableColumns=True, suppressContextMenu=True,
        domLayout="autoHeight",
    )
    resp = AgGrid(
        df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MANUAL,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        theme="alpine",
        fit_columns_on_grid_load=True,
        key=key,
    )
    return pd.DataFrame(resp["data"])

# ── Sidebar ───────────────────────────────────────────────────────────────────
registry = load_registry()

with st.sidebar:
    st.markdown("## POC Tracker")
    st.divider()

    for k, d in [("editor_name", ""), ("editor_role", "Snowflake Team")]:
        if k not in st.session_state:
            st.session_state[k] = d

    st.markdown("**Editing as**")
    st.session_state["editor_name"] = st.text_input(
        "name", value=st.session_state["editor_name"],
        placeholder="Your name", label_visibility="collapsed"
    )
    st.session_state["editor_role"] = st.radio(
        "role", ["Snowflake Team", "Customer"],
        index=0 if st.session_state["editor_role"] == "Snowflake Team" else 1,
        horizontal=True, label_visibility="collapsed"
    )
    st.divider()

    poc_options = registry["POC_ID"].tolist() if not registry.empty else []

    if poc_options:
        labels = []
        for _, r in registry.iterrows():
            e   = STATUS_EMOJI.get(r.get("Status", ""), "⚪")
            eng = r.get("Engagement", "")
            labels.append(f"{e}  {r['Customer']}" + (f"  —  {eng}" if eng else ""))

        idx = st.radio("Select Engagement", range(len(poc_options)),
                       format_func=lambda i: labels[i], key="poc_radio")
        active_poc_id = poc_options[idx]
        active_row    = registry[registry["POC_ID"] == active_poc_id].iloc[0]
    else:
        st.info("No POCs yet.")
        active_poc_id = None
        active_row    = {}

    st.divider()
    with st.expander("＋  New POC", expanded=not poc_options):
        with st.form("new_poc_form"):
            np_cust = st.text_input("Customer Name *")
            np_eng  = st.text_input("Engagement / BU")
            np_st   = st.selectbox("Status", STATUS_OPTIONS)
            if st.form_submit_button("Create", type="primary"):
                if np_cust.strip():
                    slug = (np_cust.strip() + "-" + np_eng.strip()).lower().replace(" ","-").replace("·","").replace("/","")[:32]
                    create_poc(slug, np_cust.strip(), np_eng.strip(), np_st)
                    st.success(f"Created: {np_cust}")
                    st.rerun()
                else:
                    st.warning("Customer name required.")

# ── Guard ─────────────────────────────────────────────────────────────────────
if not active_poc_id:
    st.title("Snowflake POC Tracker")
    st.info("Create your first POC engagement using the sidebar.")
    st.stop()

# ── Data ──────────────────────────────────────────────────────────────────────
ov      = load_overview(active_poc_id)
kpis    = load_kpis(active_poc_id)
actions = load_actions(active_poc_id)

# ── Header ────────────────────────────────────────────────────────────────────
customer_name   = active_row.get("Customer", active_poc_id)
engagement_name = active_row.get("Engagement", "")
poc_status      = ov.get("Status", active_row.get("Status", ""))

st.markdown(f"## {customer_name} × Snowflake")
if engagement_name:
    st.caption(engagement_name)

# Stat cards
kpis_met   = int((kpis["Status"] == "Met").sum())     if not kpis.empty    else 0
kpis_total = len(kpis)
acts_open  = int((actions["Status"] == "Open").sum()) if not actions.empty else 0
acts_total = len(actions)

st.html(f"""
<div class="stat-row">
  <div class="stat-card">
    <div class="stat-label">Status</div>
    <div style="margin-top:6px">{pill(poc_status) if poc_status else "—"}</div>
  </div>
  <div class="stat-card">
    <div class="stat-label">KPIs Met</div>
    <div class="stat-value">{kpis_met}<span style="font-size:1rem;color:#CBD5E0;font-weight:400"> / {kpis_total}</span></div>
  </div>
  <div class="stat-card">
    <div class="stat-label">Actions Open</div>
    <div class="stat-value">{acts_open}<span style="font-size:1rem;color:#CBD5E0;font-weight:400"> / {acts_total}</span></div>
  </div>
  <div class="stat-card">
    <div class="stat-label">POC Budget</div>
    <div class="stat-value-sm">{fmt_money(ov.get("POC Budget ($)",""))}</div>
    <div class="stat-sub">Spent: {fmt_money(ov.get("Confirmed Spend ($)",""))}</div>
  </div>
</div>
""")

# ── KPIs ──────────────────────────────────────────────────────────────────────
st.html('<div class="section-label">Key Performance Indicators</div>')

DISPLAY_KPI = [h for h in KPI_HEADERS if h != "POC_ID"]
kpi_df = kpis.copy() if not kpis.empty else pd.DataFrame(columns=DISPLAY_KPI)
for col in DISPLAY_KPI:
    if col not in kpi_df.columns:
        kpi_df[col] = ""
kpi_df = kpi_df[DISPLAY_KPI]
for col in DISPLAY_KPI:
    kpi_df[col] = kpi_df[col].astype(str).replace("nan", "")

edited_kpis = ag_table(kpi_df, {
    "KPI":           {"flex": 3},
    "Target":        {"width": 90},
    "Current Value": {"width": 110},
    "Unit":          {"width": 70},
    "Status":        {"width": 130, "cellEditor": "agSelectCellEditor",
                      "cellEditorParams": {"values": KPI_STATUSES}, "cellEditorPopup": True},
    "Notes":         {"flex": 2},
}, key="kpi_grid")

if st.button("Save KPIs", type="primary", key="save_kpis"):
    clean = edited_kpis.dropna(how="all")
    clean = clean[clean["KPI"].astype(str).str.strip() != ""]
    save_all_rows("KPIs", active_poc_id, DISPLAY_KPI, clean)
    st.success("Saved.")
    st.rerun()

st.divider()

# ── Action Plan ───────────────────────────────────────────────────────────────
st.html('<div class="section-label">Action Plan</div>')

DISPLAY_ACT = [h for h in ACTION_HEADERS if h != "POC_ID"]
act_df = actions.copy() if not actions.empty else pd.DataFrame(columns=DISPLAY_ACT)
for col in DISPLAY_ACT:
    if col not in act_df.columns:
        act_df[col] = ""
act_df = act_df[DISPLAY_ACT]
for col in DISPLAY_ACT:
    act_df[col] = act_df[col].astype(str).replace("nan", "")

edited_actions = ag_table(act_df, {
    "Action":   {"flex": 3},
    "Assignee": {"width": 130},
    "Category": {"width": 130, "cellEditor": "agSelectCellEditor",
                 "cellEditorParams": {"values": ACTION_CATS}, "cellEditorPopup": True},
    "Due Date": {"width": 110},
    "Status":   {"width": 120, "cellEditor": "agSelectCellEditor",
                 "cellEditorParams": {"values": ACTION_STATUSES}, "cellEditorPopup": True},
    "Notes":    {"flex": 2},
}, key="action_grid")

if st.button("Save Actions", type="primary", key="save_actions"):
    clean = edited_actions.dropna(how="all")
    clean = clean[clean["Action"].astype(str).str.strip() != ""]
    save_all_rows("Action_Items", active_poc_id, DISPLAY_ACT, clean)
    st.success("Saved.")
    st.rerun()

st.divider()

# ── POC Details (collapsible) ─────────────────────────────────────────────────
with st.expander("POC Details", expanded=False):
    editor_name = st.session_state.get("editor_name", "")
    editor_role = st.session_state.get("editor_role", "Snowflake Team")

    with st.form("overview_form"):
        st.markdown("**Contacts**")
        champion      = st.text_input("Technical Champion",          value=ov.get("Technical Champion", ""))
        champ_title   = st.text_input("Technical Champion Title",    value=ov.get("Technical Champion Title", ""))
        sponsor       = st.text_input("Exec Business Sponsor",       value=ov.get("Exec Business Sponsor", ""))
        sponsor_title = st.text_input("Exec Business Sponsor Title", value=ov.get("Exec Business Sponsor Title", ""))

        participants = st.text_area(
            "Customer Participants", value=ov.get("Customer Participants", ""),
            height=70, placeholder="e.g. Jane Smith (Data Engineer), John Doe (IT Lead)…"
        )

        st.divider()
        st.markdown("**Engagement**")
        e1, e2 = st.columns(2)
        with e1:
            bu = st.selectbox("Business Unit", BUSINESS_UNITS,
                              index=BUSINESS_UNITS.index(ov["Business Unit"]) if ov.get("Business Unit") in BUSINESS_UNITS else 0)
        with e2:
            cloud    = st.selectbox("Cloud Environment", CLOUD_OPTIONS,
                                    index=CLOUD_OPTIONS.index(ov["Cloud Environment"]) if ov.get("Cloud Environment") in CLOUD_OPTIONS else 0)
            platform = st.text_input("Current Data Platform", value=ov.get("Current Data Platform", ""))
            volume   = st.text_input("Estimated Data Volume",  value=ov.get("Data Volume", ""))

        saved_compliance = [c.strip() for c in ov.get("Compliance Requirements","").split(",") if c.strip()]
        compliance = st.multiselect("Compliance Requirements", COMPLIANCE_OPTIONS,
                                    default=[c for c in saved_compliance if c in COMPLIANCE_OPTIONS])

        st.divider()
        st.markdown("**Timeline & Status**")
        t1, t2, t3 = st.columns(3)
        with t1:
            try:    start_val = datetime.strptime(ov.get("POC Start Date",""), "%Y-%m-%d").date()
            except: start_val = date.today()
            start_date = st.date_input("POC Start Date", value=start_val)
        with t2:
            try:    end_val = datetime.strptime(ov.get("Target Completion Date",""), "%Y-%m-%d").date()
            except: end_val = date.today()
            end_date = st.date_input("Target Completion Date", value=end_val)
        with t3:
            status = st.selectbox("POC Status", STATUS_OPTIONS,
                                  index=STATUS_OPTIONS.index(ov["Status"]) if ov.get("Status") in STATUS_OPTIONS else 0)

        st.divider()
        st.markdown("**Objectives**")
        objective = st.text_area("POC Objective", value=ov.get("POC Objective",""), height=80)
        sc1, sc2 = st.columns(2)
        with sc1:
            tech_criteria = st.text_area("Technical Success Criteria", value=ov.get("Technical Success Criteria",""), height=90)
        with sc2:
            biz_criteria  = st.text_area("Business Success Criteria",  value=ov.get("Business Success Criteria",""),  height=90)

        st.divider()
        st.markdown("**Budget**")
        b1, b2 = st.columns(2)
        with b1: poc_budget = st.text_input("POC Budget ($)",      value=ov.get("POC Budget ($)",""))
        with b2: conf_spend = st.text_input("Confirmed Spend ($)", value=ov.get("Confirmed Spend ($)",""))

        if st.form_submit_button("Save Details", type="primary", use_container_width=True):
            save_overview(active_poc_id, {
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
