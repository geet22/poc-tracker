import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date

st.set_page_config(page_title="Labcorp × Snowflake", layout="wide", page_icon="❄️")

# ── Styling ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .header-banner {
        background: linear-gradient(135deg, #29B5E8 0%, #0C3D6E 100%);
        padding: 24px 32px;
        border-radius: 12px;
        margin-bottom: 24px;
        color: white;
    }
    .header-banner h1 { color: white; margin: 0; font-size: 1.8rem; }
    .header-banner p  { color: rgba(255,255,255,0.8); margin: 4px 0 0 0; font-size: 0.95rem; }
    .status-badge {
        display: inline-block;
        padding: 4px 14px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
    }
    .metric-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
    }
    .section-label {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #64748b;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

OVERVIEW_HEADERS = [
    "Snowflake AE", "Snowflake SE",
    "Labcorp Champion", "Champion Title", "Champion Email",
    "Executive Sponsor", "Sponsor Title",
    "Business Unit", "Primary Use Case", "Secondary Use Case",
    "Cloud Environment", "Current Data Platform", "Data Volume",
    "Compliance Requirements",
    "POC Start Date", "Target Completion Date", "Status",
    "POC Objective", "Technical Success Criteria", "Business Success Criteria",
]

ACTION_HEADERS = ["Action", "Owner", "Category", "Due Date", "Status", "Notes"]
UPDATE_HEADERS = ["Date", "Update", "Posted By"]

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

BUSINESS_UNITS = [
    "Diagnostics",
    "Drug Development (Biopharma Solutions)",
    "Genomics",
    "Technology Solutions",
    "Enterprise / Cross-BU",
]

COMPLIANCE_OPTIONS = ["HIPAA", "21 CFR Part 11", "GxP", "SOC 2", "CLIA", "GDPR"]

STATUS_OPTIONS  = ["Planning", "Active", "At Risk", "Completed — Won", "Completed — Lost"]
ACTION_OWNERS   = ["Snowflake", "Labcorp", "Joint"]
ACTION_CATS     = ["Technical", "Business", "Compliance", "Training", "Executive"]
ACTION_STATUSES = ["Open", "In Progress", "Complete", "Blocked"]
CLOUD_OPTIONS   = ["AWS", "Azure", "Multi-Cloud"]

STATUS_STYLE = {
    "Planning":          ("🔵", "#dbeafe", "#1e40af"),
    "Active":            ("🟢", "#dcfce7", "#166534"),
    "At Risk":           ("🟡", "#fef9c3", "#854d0e"),
    "Completed — Won":   ("✅", "#dcfce7", "#166534"),
    "Completed — Lost":  ("🔴", "#fee2e2", "#991b1b"),
}


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
        ("Overview", OVERVIEW_HEADERS),
        ("Action_Items", ACTION_HEADERS),
        ("Updates", UPDATE_HEADERS),
    ]:
        if title not in existing:
            ws = ss.add_worksheet(title=title, rows=500, cols=len(headers))
            ws.append_row(headers)
        else:
            ws = ss.worksheet(title)
            if not ws.row_values(1):
                ws.insert_row(headers, 1)


@st.cache_data(ttl=15)
def load_overview():
    ws = get_ss().worksheet("Overview")
    rows = ws.get_all_records()
    return rows[0] if rows else {}


@st.cache_data(ttl=15)
def load_actions():
    ws = get_ss().worksheet("Action_Items")
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=ACTION_HEADERS)


@st.cache_data(ttl=15)
def load_updates():
    ws = get_ss().worksheet("Updates")
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=UPDATE_HEADERS)


def save_overview(values: dict):
    ws = get_ss().worksheet("Overview")
    rows = ws.get_all_values()
    row = [values.get(h, "") for h in OVERVIEW_HEADERS]
    if len(rows) < 2:
        ws.append_row(row)
    else:
        ws.delete_rows(2)
        ws.insert_row(row, 2)
    st.cache_data.clear()


def append_action(row: list):
    get_ss().worksheet("Action_Items").append_row(row, value_input_option="USER_ENTERED")
    st.cache_data.clear()


def update_action(sheet_row: int, values: dict):
    ws = get_ss().worksheet("Action_Items")
    for i, h in enumerate(ACTION_HEADERS, 1):
        ws.update_cell(sheet_row, i, values.get(h, ""))
    st.cache_data.clear()


def delete_action(sheet_row: int):
    get_ss().worksheet("Action_Items").delete_rows(sheet_row)
    st.cache_data.clear()


def append_update(row: list):
    get_ss().worksheet("Updates").append_row(row, value_input_option="USER_ENTERED")
    st.cache_data.clear()


# ── Boot ───────────────────────────────────────────────────────────────────

try:
    ensure_sheets()
except Exception as e:
    st.error(f"Could not connect to Google Sheets: {e}")
    st.stop()

# ── Header ─────────────────────────────────────────────────────────────────

ov = load_overview()
status = ov.get("Status", "Planning")
emoji, bg, fg = STATUS_STYLE.get(status, ("🔵", "#dbeafe", "#1e40af"))

st.markdown(f"""
<div class="header-banner">
    <h1>Labcorp × Snowflake &nbsp; POC Engagement Tracker</h1>
    <p>
        {ov.get('Business Unit', '') or 'Business unit not set'} &nbsp;·&nbsp;
        {ov.get('Primary Use Case', '') or 'Use case not set'} &nbsp;·&nbsp;
        <span style="background:{bg};color:{fg};padding:2px 10px;border-radius:12px;font-weight:600;font-size:0.8rem;">
            {emoji} {status}
        </span>
    </p>
</div>
""", unsafe_allow_html=True)

# ── Navigation ─────────────────────────────────────────────────────────────

tab1, tab2, tab3 = st.tabs(["📋  Overview", "✅  Success Plan", "📝  Updates"])

# ══════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════

with tab1:

    # ── Key contacts summary ───────────────────────────────────────────────
    if ov:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown('<p class="section-label">Snowflake AE</p>', unsafe_allow_html=True)
            st.markdown(f"**{ov.get('Snowflake AE') or '—'}**")
        with c2:
            st.markdown('<p class="section-label">Snowflake SE</p>', unsafe_allow_html=True)
            st.markdown(f"**{ov.get('Snowflake SE') or '—'}**")
        with c3:
            st.markdown('<p class="section-label">Labcorp Champion</p>', unsafe_allow_html=True)
            st.markdown(f"**{ov.get('Labcorp Champion') or '—'}**")
            st.caption(ov.get("Champion Title") or "")
        with c4:
            st.markdown('<p class="section-label">Executive Sponsor</p>', unsafe_allow_html=True)
            st.markdown(f"**{ov.get('Executive Sponsor') or '—'}**")
            st.caption(ov.get("Sponsor Title") or "")

        st.markdown("---")

        # ── Timeline ──────────────────────────────────────────────────────
        d1, d2 = st.columns(2)
        with d1:
            st.markdown('<p class="section-label">POC Start</p>', unsafe_allow_html=True)
            st.markdown(f"**{ov.get('POC Start Date') or '—'}**")
        with d2:
            st.markdown('<p class="section-label">Target Completion</p>', unsafe_allow_html=True)
            st.markdown(f"**{ov.get('Target Completion Date') or '—'}**")

        st.markdown("---")

        # ── Objective + criteria ──────────────────────────────────────────
        if ov.get("POC Objective"):
            st.markdown('<p class="section-label">POC Objective</p>', unsafe_allow_html=True)
            st.markdown(ov["POC Objective"])
        col_a, col_b = st.columns(2)
        with col_a:
            if ov.get("Technical Success Criteria"):
                st.markdown('<p class="section-label">Technical Success Criteria</p>', unsafe_allow_html=True)
                st.markdown(ov["Technical Success Criteria"])
        with col_b:
            if ov.get("Business Success Criteria"):
                st.markdown('<p class="section-label">Business Success Criteria</p>', unsafe_allow_html=True)
                st.markdown(ov["Business Success Criteria"])

        st.markdown("---")

    with st.expander("✏️  Edit overview details", expanded=not bool(ov)):
        with st.form("overview_form"):

            st.markdown("**Snowflake Team**")
            sf1, sf2 = st.columns(2)
            with sf1:
                ae = st.text_input("Account Executive", value=ov.get("Snowflake AE", ""))
            with sf2:
                se = st.text_input("Solutions Engineer", value=ov.get("Snowflake SE", ""))

            st.markdown("**Labcorp Contacts**")
            lc1, lc2 = st.columns(2)
            with lc1:
                champion      = st.text_input("Champion Name", value=ov.get("Labcorp Champion", ""))
                champion_title = st.text_input("Champion Title", value=ov.get("Champion Title", ""))
                champion_email = st.text_input("Champion Email", value=ov.get("Champion Email", ""))
            with lc2:
                sponsor       = st.text_input("Executive Sponsor Name", value=ov.get("Executive Sponsor", ""))
                sponsor_title = st.text_input("Executive Sponsor Title", value=ov.get("Sponsor Title", ""))

            st.markdown("**Engagement Details**")
            ed1, ed2 = st.columns(2)
            with ed1:
                bu_idx  = BUSINESS_UNITS.index(ov["Business Unit"]) if ov.get("Business Unit") in BUSINESS_UNITS else 0
                bu      = st.selectbox("Business Unit", BUSINESS_UNITS, index=bu_idx)
                uc_idx  = USE_CASES.index(ov["Primary Use Case"]) if ov.get("Primary Use Case") in USE_CASES else 0
                uc      = st.selectbox("Primary Use Case", USE_CASES, index=uc_idx)
                uc2_idx = USE_CASES.index(ov["Secondary Use Case"]) if ov.get("Secondary Use Case") in USE_CASES else 0
                uc2     = st.selectbox("Secondary Use Case (optional)", USE_CASES, index=uc2_idx)
            with ed2:
                cloud_idx = CLOUD_OPTIONS.index(ov["Cloud Environment"]) if ov.get("Cloud Environment") in CLOUD_OPTIONS else 0
                cloud     = st.selectbox("Cloud Environment", CLOUD_OPTIONS, index=cloud_idx)
                platform  = st.text_input("Current Data Platform", value=ov.get("Current Data Platform", ""),
                                          placeholder="e.g. Teradata, Oracle, Azure Synapse")
                volume    = st.text_input("Estimated Data Volume", value=ov.get("Data Volume", ""),
                                          placeholder="e.g. 50 TB, 200M rows/day")

            saved_compliance = [c.strip() for c in ov.get("Compliance Requirements", "").split(",") if c.strip()]
            compliance = st.multiselect("Compliance Requirements", COMPLIANCE_OPTIONS,
                                        default=[c for c in saved_compliance if c in COMPLIANCE_OPTIONS])

            st.markdown("**Timeline & Status**")
            ts1, ts2, ts3 = st.columns(3)
            with ts1:
                try:
                    start_val = datetime.strptime(ov.get("POC Start Date", ""), "%Y-%m-%d").date()
                except ValueError:
                    start_val = date.today()
                start_date = st.date_input("POC Start Date", value=start_val)
            with ts2:
                try:
                    end_val = datetime.strptime(ov.get("Target Completion Date", ""), "%Y-%m-%d").date()
                except ValueError:
                    end_val = date.today()
                end_date = st.date_input("Target Completion Date", value=end_val)
            with ts3:
                st_idx  = STATUS_OPTIONS.index(ov["Status"]) if ov.get("Status") in STATUS_OPTIONS else 0
                s_status = st.selectbox("POC Status", STATUS_OPTIONS, index=st_idx)

            st.markdown("**Objectives & Success Criteria**")
            objective  = st.text_area("POC Objective", value=ov.get("POC Objective", ""), height=80,
                                      placeholder="What are we trying to prove with this POC?")
            cr1, cr2 = st.columns(2)
            with cr1:
                tech_criteria = st.text_area("Technical Success Criteria", value=ov.get("Technical Success Criteria", ""),
                                             height=90, placeholder="e.g. Query performance < 2s on 50TB dataset, Snowpark pipeline processes 10M records/hr")
            with cr2:
                biz_criteria  = st.text_area("Business Success Criteria", value=ov.get("Business Success Criteria", ""),
                                             height=90, placeholder="e.g. 3 validated use cases, RWE report delivered to pharma partner")

            if st.form_submit_button("Save Overview", type="primary"):
                save_overview({
                    "Snowflake AE": ae, "Snowflake SE": se,
                    "Labcorp Champion": champion, "Champion Title": champion_title, "Champion Email": champion_email,
                    "Executive Sponsor": sponsor, "Sponsor Title": sponsor_title,
                    "Business Unit": bu, "Primary Use Case": uc, "Secondary Use Case": uc2,
                    "Cloud Environment": cloud, "Current Data Platform": platform, "Data Volume": volume,
                    "Compliance Requirements": ", ".join(compliance),
                    "POC Start Date": str(start_date), "Target Completion Date": str(end_date),
                    "Status": s_status,
                    "POC Objective": objective,
                    "Technical Success Criteria": tech_criteria,
                    "Business Success Criteria": biz_criteria,
                })
                st.success("Overview saved.")
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════
# TAB 2 — SUCCESS PLAN
# ══════════════════════════════════════════════════════════════════════════

with tab2:
    actions = load_actions()

    # ── Summary counts ─────────────────────────────────────────────────────
    if not actions.empty:
        mc1, mc2, mc3, mc4 = st.columns(4)
        mc1.metric("Total Actions", len(actions))
        mc2.metric("Open", len(actions[actions["Status"] == "Open"]))
        mc3.metric("In Progress", len(actions[actions["Status"] == "In Progress"]))
        mc4.metric("Blocked", len(actions[actions["Status"] == "Blocked"]))
        st.markdown("---")

    # ── Existing action items ──────────────────────────────────────────────
    if not actions.empty:
        cat_filter = st.multiselect("Filter by category", ACTION_CATS, default=ACTION_CATS)
        filtered   = actions[actions["Category"].isin(cat_filter)] if "Category" in actions.columns else actions

        for idx, row in filtered.iterrows():
            sheet_row = idx + 2
            s = row.get("Status", "Open")
            icon = {"Open": "⬜", "In Progress": "🔄", "Complete": "✅", "Blocked": "🚫"}.get(s, "⬜")
            label = f"{icon} **{row.get('Action', '')}** — {row.get('Owner', '')} · {row.get('Category', '')} · Due {row.get('Due Date', '')}"
            with st.expander(label):
                with st.form(f"action_edit_{idx}"):
                    action_text = st.text_input("Action", value=row.get("Action", ""))
                    ec1, ec2, ec3 = st.columns(3)
                    with ec1:
                        ow_idx = ACTION_OWNERS.index(row["Owner"]) if row.get("Owner") in ACTION_OWNERS else 0
                        owner  = st.selectbox("Owner", ACTION_OWNERS, index=ow_idx)
                    with ec2:
                        ca_idx = ACTION_CATS.index(row["Category"]) if row.get("Category") in ACTION_CATS else 0
                        cat    = st.selectbox("Category", ACTION_CATS, index=ca_idx)
                    with ec3:
                        as_idx  = ACTION_STATUSES.index(row["Status"]) if row.get("Status") in ACTION_STATUSES else 0
                        a_status = st.selectbox("Status", ACTION_STATUSES, index=as_idx)
                    try:
                        due_val = datetime.strptime(row.get("Due Date", ""), "%Y-%m-%d").date()
                    except (ValueError, TypeError):
                        due_val = date.today()
                    due  = st.date_input("Due Date", value=due_val, key=f"due_{idx}")
                    note = st.text_input("Notes", value=row.get("Notes", ""))
                    sc1, sc2 = st.columns([3, 1])
                    with sc1:
                        if st.form_submit_button("Update", type="primary"):
                            update_action(sheet_row, {
                                "Action": action_text, "Owner": owner, "Category": cat,
                                "Due Date": str(due), "Status": a_status, "Notes": note,
                            })
                            st.success("Updated.")
                            st.rerun()
                    with sc2:
                        if st.form_submit_button("🗑 Delete"):
                            delete_action(sheet_row)
                            st.rerun()
    else:
        st.info("No action items yet. Add the first one below.")

    st.markdown("---")
    st.subheader("Add Action Item")
    with st.form("new_action"):
        new_text = st.text_input("Action *", placeholder="e.g. Provision Snowflake trial environment with HIPAA controls")
        na1, na2, na3, na4 = st.columns(4)
        with na1:
            new_owner  = st.selectbox("Owner", ACTION_OWNERS)
        with na2:
            new_cat    = st.selectbox("Category", ACTION_CATS)
        with na3:
            new_due    = st.date_input("Due Date", value=date.today())
        with na4:
            new_status = st.selectbox("Status", ACTION_STATUSES)
        new_note = st.text_input("Notes (optional)")
        if st.form_submit_button("Add Action Item", type="primary"):
            if new_text.strip():
                append_action([new_text.strip(), new_owner, new_cat, str(new_due), new_status, new_note])
                st.success("Added.")
                st.rerun()
            else:
                st.warning("Action text is required.")


# ══════════════════════════════════════════════════════════════════════════
# TAB 3 — UPDATES
# ══════════════════════════════════════════════════════════════════════════

with tab3:
    st.subheader("Post a Status Update")
    with st.form("new_update"):
        uf1, uf2 = st.columns([3, 1])
        with uf2:
            posted_by = st.text_input("Your name")
        with uf1:
            update_text = st.text_area("Update", height=90,
                                       placeholder="Key outcomes, decisions, blockers, or next steps from this week…")
        if st.form_submit_button("Post Update", type="primary"):
            if update_text.strip() and posted_by.strip():
                append_update([datetime.today().strftime("%Y-%m-%d %H:%M"), update_text.strip(), posted_by.strip()])
                st.success("Posted.")
                st.rerun()
            else:
                st.warning("Both name and update are required.")

    st.markdown("---")

    updates = load_updates()
    if not updates.empty:
        for _, u in updates.sort_values("Date", ascending=False).iterrows():
            st.markdown(f"**{u.get('Date', '')}** &nbsp;·&nbsp; *{u.get('Posted By', '')}*")
            st.markdown(u.get("Update", ""))
            st.markdown("---")
    else:
        st.caption("No updates posted yet.")
