import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date

st.set_page_config(page_title="Labcorp × Snowflake POC", layout="wide")

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
    "POC Budget ($)", "Confirmed Spend ($)", "Potential ARR ($)",
    "Procurement Contact", "Budget Notes",
]

KPI_HEADERS    = ["KPI", "Target", "Current Value", "Unit", "Status", "Notes"]
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
BUSINESS_UNITS    = ["Diagnostics", "Drug Development (Biopharma Solutions)", "Genomics", "Technology Solutions", "Enterprise / Cross-BU"]
COMPLIANCE_OPTIONS = ["HIPAA", "21 CFR Part 11", "GxP", "SOC 2", "CLIA", "GDPR"]
STATUS_OPTIONS    = ["Planning", "Active", "At Risk", "Completed — Won", "Completed — Lost"]
KPI_STATUSES      = ["On Track", "At Risk", "Met", "Not Started"]
ACTION_OWNERS     = ["Snowflake", "Labcorp", "Joint"]
ACTION_CATS       = ["Technical", "Business", "Compliance", "Training", "Executive"]
ACTION_STATUSES   = ["Open", "In Progress", "Complete", "Blocked"]
CLOUD_OPTIONS     = ["AWS", "Azure", "Multi-Cloud"]


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
        ("KPIs", KPI_HEADERS),
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
def load_kpis():
    ws = get_ss().worksheet("KPIs")
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=KPI_HEADERS)

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
    row = [values.get(h, "") for h in OVERVIEW_HEADERS]
    all_rows = ws.get_all_values()
    if len(all_rows) < 2:
        ws.append_row(row)
    else:
        ws.delete_rows(2)
        ws.insert_row(row, 2)
    st.cache_data.clear()

def append_row_to(sheet: str, row: list):
    get_ss().worksheet(sheet).append_row(row, value_input_option="USER_ENTERED")
    st.cache_data.clear()

def update_sheet_row(sheet: str, sheet_row: int, headers: list, values: dict):
    ws = get_ss().worksheet(sheet)
    for i, h in enumerate(headers, 1):
        ws.update_cell(sheet_row, i, values.get(h, ""))
    st.cache_data.clear()

def delete_sheet_row(sheet: str, sheet_row: int):
    get_ss().worksheet(sheet).delete_rows(sheet_row)
    st.cache_data.clear()


try:
    ensure_sheets()
except Exception as e:
    st.error(f"Could not connect to Google Sheets: {e}")
    st.stop()

ov = load_overview()

st.title("Labcorp × Snowflake — POC Tracker")
st.caption(f"{ov.get('Business Unit', 'Business unit not set')}  ·  {ov.get('Primary Use Case', 'Use case not set')}  ·  Status: **{ov.get('Status', 'Not set')}**")
st.divider()

tab0, tab1, tab2, tab3, tab4 = st.tabs(["Sheet View", "Overview", "KPIs & Budget", "Action Plan", "Updates"])


# ══════════════════════════════════════════════════════════════════════════
# TAB 0 — SHEET VIEW (live read-only view of all 4 sheets)
# ══════════════════════════════════════════════════════════════════════════

with tab0:
    st.subheader("Live Sheet Data")
    st.caption("Read-only mirror of your Google Sheet. Refreshes every 15 seconds.")

    # Overview
    st.markdown("#### Overview")
    if ov:
        ov_df = pd.DataFrame([ov])
        # split into two halves for readability
        cols = list(ov_df.columns)
        half = len(cols) // 2
        st.dataframe(ov_df[cols[:half]], use_container_width=True, hide_index=True)
        st.dataframe(ov_df[cols[half:]], use_container_width=True, hide_index=True)
    else:
        st.dataframe(pd.DataFrame(columns=OVERVIEW_HEADERS), use_container_width=True, hide_index=True)
        st.caption("No overview data yet — fill it in via the Overview tab.")

    st.divider()

    # KPIs
    st.markdown("#### KPIs")
    kpis_sv = load_kpis()
    st.dataframe(kpis_sv if not kpis_sv.empty else pd.DataFrame(columns=KPI_HEADERS),
                 use_container_width=True, hide_index=True)

    st.divider()

    # Action Items
    st.markdown("#### Mutual Action Plan")
    actions_sv = load_actions()
    st.dataframe(actions_sv if not actions_sv.empty else pd.DataFrame(columns=ACTION_HEADERS),
                 use_container_width=True, hide_index=True)

    st.divider()

    # Updates
    st.markdown("#### Status Updates")
    updates_sv = load_updates()
    if not updates_sv.empty:
        st.dataframe(updates_sv.sort_values("Date", ascending=False),
                     use_container_width=True, hide_index=True)
    else:
        st.dataframe(pd.DataFrame(columns=UPDATE_HEADERS), use_container_width=True, hide_index=True)

    st.divider()
    if st.button("Refresh"):
        st.cache_data.clear()
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════

with tab1:

    # Read-only snapshot of what is in the sheet
    if ov:
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Engagement")
            rows = {
                "Business Unit":        ov.get("Business Unit"),
                "Primary Use Case":     ov.get("Primary Use Case"),
                "Secondary Use Case":   ov.get("Secondary Use Case"),
                "Cloud Environment":    ov.get("Cloud Environment"),
                "Current Platform":     ov.get("Current Data Platform"),
                "Data Volume":          ov.get("Data Volume"),
                "Compliance":           ov.get("Compliance Requirements"),
                "POC Start":            ov.get("POC Start Date"),
                "Target Completion":    ov.get("Target Completion Date"),
                "Status":               ov.get("Status"),
            }
            for k, v in rows.items():
                if v:
                    st.markdown(f"**{k}:** {v}")

        with col2:
            st.subheader("Contacts")
            contacts = {
                "Snowflake AE":        ov.get("Snowflake AE"),
                "Snowflake SE":        ov.get("Snowflake SE"),
                "Labcorp Champion":    ov.get("Labcorp Champion"),
                "Champion Title":      ov.get("Champion Title"),
                "Champion Email":      ov.get("Champion Email"),
                "Executive Sponsor":   ov.get("Executive Sponsor"),
                "Sponsor Title":       ov.get("Sponsor Title"),
            }
            for k, v in contacts.items():
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

            st.markdown("**Labcorp Contacts**")
            f3, f4 = st.columns(2)
            with f3:
                champion       = st.text_input("Champion Name",   value=ov.get("Labcorp Champion", ""))
                champ_title    = st.text_input("Champion Title",  value=ov.get("Champion Title", ""))
                champ_email    = st.text_input("Champion Email",  value=ov.get("Champion Email", ""))
            with f4:
                sponsor        = st.text_input("Executive Sponsor Name",  value=ov.get("Executive Sponsor", ""))
                sponsor_title  = st.text_input("Executive Sponsor Title", value=ov.get("Sponsor Title", ""))

            st.markdown("**Engagement Details**")
            f5, f6 = st.columns(2)
            with f5:
                bu   = st.selectbox("Business Unit",     BUSINESS_UNITS,
                                    index=BUSINESS_UNITS.index(ov["Business Unit"]) if ov.get("Business Unit") in BUSINESS_UNITS else 0)
                uc   = st.selectbox("Primary Use Case",  USE_CASES,
                                    index=USE_CASES.index(ov["Primary Use Case"]) if ov.get("Primary Use Case") in USE_CASES else 0)
                uc2  = st.selectbox("Secondary Use Case", USE_CASES,
                                    index=USE_CASES.index(ov["Secondary Use Case"]) if ov.get("Secondary Use Case") in USE_CASES else 0)
            with f6:
                cloud    = st.selectbox("Cloud Environment", CLOUD_OPTIONS,
                                        index=CLOUD_OPTIONS.index(ov["Cloud Environment"]) if ov.get("Cloud Environment") in CLOUD_OPTIONS else 0)
                platform = st.text_input("Current Data Platform", value=ov.get("Current Data Platform", ""),
                                         placeholder="e.g. Teradata, Oracle, Azure Synapse")
                volume   = st.text_input("Estimated Data Volume", value=ov.get("Data Volume", ""),
                                         placeholder="e.g. 50 TB, 200M rows/day")

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

            objective     = st.text_area("POC Objective", value=ov.get("POC Objective", ""), height=80)
            fa, fb = st.columns(2)
            with fa:
                tech_criteria = st.text_area("Technical Success Criteria", value=ov.get("Technical Success Criteria", ""), height=90)
            with fb:
                biz_criteria  = st.text_area("Business Success Criteria",  value=ov.get("Business Success Criteria", ""), height=90)

            if st.form_submit_button("Save Overview", type="primary"):
                save_overview({
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


# ══════════════════════════════════════════════════════════════════════════
# TAB 2 — KPIs & BUDGET
# ══════════════════════════════════════════════════════════════════════════

with tab2:

    # ── KPIs ──────────────────────────────────────────────────────────────
    st.subheader("KPIs")
    kpis = load_kpis()

    if not kpis.empty:
        # Show live KPI table from the sheet
        st.dataframe(kpis, use_container_width=True, hide_index=True)
        st.divider()

        st.markdown("**Update a KPI**")
        for idx, row in kpis.iterrows():
            sheet_row = idx + 2
            with st.expander(f"{row.get('KPI', '')} — Target: {row.get('Target', '')} {row.get('Unit', '')} — {row.get('Status', '')}"):
                with st.form(f"kpi_edit_{idx}"):
                    kc1, kc2 = st.columns(2)
                    with kc1:
                        kpi_name = st.text_input("KPI",    value=row.get("KPI", ""))
                        target   = st.text_input("Target", value=str(row.get("Target", "")))
                        unit     = st.text_input("Unit",   value=row.get("Unit", ""), placeholder="e.g. seconds, TB, %")
                    with kc2:
                        current  = st.text_input("Current Value", value=str(row.get("Current Value", "")))
                        ks_idx   = KPI_STATUSES.index(row["Status"]) if row.get("Status") in KPI_STATUSES else 0
                        k_status = st.selectbox("Status", KPI_STATUSES, index=ks_idx)
                        k_note   = st.text_input("Notes", value=row.get("Notes", ""))
                    sc1, sc2 = st.columns([4, 1])
                    with sc1:
                        if st.form_submit_button("Update", type="primary"):
                            update_sheet_row("KPIs", sheet_row, KPI_HEADERS,
                                             {"KPI": kpi_name, "Target": target, "Current Value": current,
                                              "Unit": unit, "Status": k_status, "Notes": k_note})
                            st.success("Updated.")
                            st.rerun()
                    with sc2:
                        if st.form_submit_button("Delete"):
                            delete_sheet_row("KPIs", sheet_row)
                            st.rerun()
    else:
        st.info("No KPIs defined yet.")

    st.divider()
    st.markdown("**Add KPI**")
    with st.form("new_kpi"):
        nk1, nk2, nk3 = st.columns(3)
        with nk1:
            nk_name   = st.text_input("KPI *", placeholder="e.g. Query Response Time")
            nk_target = st.text_input("Target", placeholder="e.g. < 2")
            nk_unit   = st.text_input("Unit",   placeholder="e.g. seconds")
        with nk2:
            nk_current = st.text_input("Current Value", placeholder="Baseline or leave blank")
            nk_status  = st.selectbox("Status", KPI_STATUSES)
        with nk3:
            nk_note = st.text_area("Notes", height=108)
        if st.form_submit_button("Add KPI", type="primary"):
            if nk_name.strip():
                append_row_to("KPIs", [nk_name.strip(), nk_target, nk_current, nk_unit, nk_status, nk_note])
                st.success("Added.")
                st.rerun()
            else:
                st.warning("KPI name required.")

    st.divider()

    # ── Budget / Funding ───────────────────────────────────────────────────
    st.subheader("Budget & Funding")

    if ov:
        budget_rows = {
            "POC Budget ($)":      ov.get("POC Budget ($)"),
            "Confirmed Spend ($)": ov.get("Confirmed Spend ($)"),
            "Potential ARR ($)":   ov.get("Potential ARR ($)"),
            "Procurement Contact": ov.get("Procurement Contact"),
            "Budget Notes":        ov.get("Budget Notes"),
        }
        if any(budget_rows.values()):
            df_budget = pd.DataFrame(
                [(k, v) for k, v in budget_rows.items() if v],
                columns=["Field", "Value"]
            )
            st.dataframe(df_budget, use_container_width=True, hide_index=True)
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
                save_overview(updated)
                st.success("Saved.")
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════
# TAB 3 — ACTION PLAN
# ══════════════════════════════════════════════════════════════════════════

with tab3:
    st.subheader("Mutual Action Plan")
    actions = load_actions()

    if not actions.empty:
        # Live table from the sheet
        st.dataframe(actions, use_container_width=True, hide_index=True)
        st.divider()

        st.markdown("**Update an Action Item**")
        for idx, row in actions.iterrows():
            sheet_row = idx + 2
            label = f"{row.get('Action', '')} — {row.get('Owner', '')} — {row.get('Status', '')}"
            with st.expander(label):
                with st.form(f"action_edit_{idx}"):
                    action_text = st.text_input("Action", value=row.get("Action", ""))
                    ac1, ac2, ac3 = st.columns(3)
                    with ac1:
                        owner = st.selectbox("Owner", ACTION_OWNERS,
                                             index=ACTION_OWNERS.index(row["Owner"]) if row.get("Owner") in ACTION_OWNERS else 0)
                    with ac2:
                        cat   = st.selectbox("Category", ACTION_CATS,
                                             index=ACTION_CATS.index(row["Category"]) if row.get("Category") in ACTION_CATS else 0)
                    with ac3:
                        a_st  = st.selectbox("Status", ACTION_STATUSES,
                                             index=ACTION_STATUSES.index(row["Status"]) if row.get("Status") in ACTION_STATUSES else 0)
                    try:    due_val = datetime.strptime(row.get("Due Date", ""), "%Y-%m-%d").date()
                    except: due_val = date.today()
                    due  = st.date_input("Due Date", value=due_val, key=f"due_{idx}")
                    note = st.text_input("Notes", value=row.get("Notes", ""))
                    sc1, sc2 = st.columns([4, 1])
                    with sc1:
                        if st.form_submit_button("Update", type="primary"):
                            update_sheet_row("Action_Items", sheet_row, ACTION_HEADERS,
                                             {"Action": action_text, "Owner": owner, "Category": cat,
                                              "Due Date": str(due), "Status": a_st, "Notes": note})
                            st.success("Updated.")
                            st.rerun()
                    with sc2:
                        if st.form_submit_button("Delete"):
                            delete_sheet_row("Action_Items", sheet_row)
                            st.rerun()
    else:
        st.info("No action items yet.")

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
                append_row_to("Action_Items", [new_text.strip(), new_owner, new_cat, str(new_due), new_status, new_note])
                st.success("Added.")
                st.rerun()
            else:
                st.warning("Action text required.")


# ══════════════════════════════════════════════════════════════════════════
# TAB 4 — UPDATES
# ══════════════════════════════════════════════════════════════════════════

with tab4:
    st.subheader("Status Updates")

    updates = load_updates()
    if not updates.empty:
        st.dataframe(updates.sort_values("Date", ascending=False), use_container_width=True, hide_index=True)
        st.divider()

    with st.form("new_update"):
        uf1, uf2 = st.columns([3, 1])
        with uf1:
            update_text = st.text_area("Update *", height=100,
                                       placeholder="Key outcomes, decisions, blockers, or next steps…")
        with uf2:
            posted_by = st.text_input("Posted by *")
        if st.form_submit_button("Post Update", type="primary"):
            if update_text.strip() and posted_by.strip():
                append_row_to("Updates", [datetime.today().strftime("%Y-%m-%d %H:%M"), update_text.strip(), posted_by.strip()])
                st.success("Posted.")
                st.rerun()
            else:
                st.warning("Both fields are required.")
