import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date

st.set_page_config(page_title="POC & Success Plan Tracker", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

POC_HEADERS = [
    "Account Name", "Industry", "Cloud Platform",
    "Snowflake POC Owner", "Customer Champion", "Champion Email",
    "Start Date", "Target End Date", "Status",
    "Use Case Summary", "Success Criteria",
]

ACTION_HEADERS = ["Account Name", "Action", "Owner", "Due Date", "Status"]

UPDATE_HEADERS = ["Account Name", "Date", "Update", "Posted By"]

STATUS_OPTIONS = ["Planning", "Active", "At Risk", "Won", "Lost"]
ACTION_OWNER_OPTIONS = ["Snowflake", "Customer", "Joint"]
ACTION_STATUS_OPTIONS = ["Open", "In Progress", "Done", "Blocked"]
CLOUD_OPTIONS = ["AWS", "Azure", "GCP", "Multi-Cloud"]
INDUSTRY_OPTIONS = [
    "Financial Services", "Healthcare", "Retail", "Technology",
    "Media & Entertainment", "Manufacturing", "Public Sector", "Other",
]

STATUS_COLORS = {
    "Planning": "🔵",
    "Active": "🟢",
    "At Risk": "🟡",
    "Won": "✅",
    "Lost": "🔴",
}


@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)


def get_spreadsheet():
    return get_client().open_by_key(st.secrets["spreadsheet_id"])


def ensure_sheets():
    """Create missing tabs and add headers if needed."""
    ss = get_spreadsheet()
    existing = {ws.title for ws in ss.worksheets()}
    for title, headers in [
        ("POCs", POC_HEADERS),
        ("Action_Items", ACTION_HEADERS),
        ("Updates", UPDATE_HEADERS),
    ]:
        if title not in existing:
            ws = ss.add_worksheet(title=title, rows=1000, cols=len(headers))
            ws.append_row(headers)
        else:
            ws = ss.worksheet(title)
            if ws.row_count == 0 or ws.row_values(1) != headers:
                ws.insert_row(headers, 1)


@st.cache_data(ttl=20)
def load_pocs():
    ws = get_spreadsheet().worksheet("POCs")
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=POC_HEADERS)


@st.cache_data(ttl=20)
def load_actions():
    ws = get_spreadsheet().worksheet("Action_Items")
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=ACTION_HEADERS)


@st.cache_data(ttl=20)
def load_updates():
    ws = get_spreadsheet().worksheet("Updates")
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=UPDATE_HEADERS)


def append_row(sheet_name: str, row: list):
    get_spreadsheet().worksheet(sheet_name).append_row(row, value_input_option="USER_ENTERED")
    st.cache_data.clear()


def update_poc_row(account_name: str, updated: dict):
    ws = get_spreadsheet().worksheet("POCs")
    records = ws.get_all_records()
    for i, rec in enumerate(records, start=2):
        if rec.get("Account Name") == account_name:
            for col_idx, header in enumerate(POC_HEADERS, start=1):
                ws.update_cell(i, col_idx, updated.get(header, rec.get(header, "")))
            st.cache_data.clear()
            return
    st.error("POC record not found.")


def update_action_row(ws_row_idx: int, updated: dict):
    ws = get_spreadsheet().worksheet("Action_Items")
    for col_idx, header in enumerate(ACTION_HEADERS, start=1):
        ws.update_cell(ws_row_idx, col_idx, updated.get(header, ""))
    st.cache_data.clear()


# ── Sidebar ────────────────────────────────────────────────────────────────

st.sidebar.title("POC Tracker")
page = st.sidebar.radio("Navigate", ["Dashboard", "POC Detail", "New POC"])
st.sidebar.markdown("---")
st.sidebar.caption("Data syncs to Google Sheets every 20 seconds.")

# ── Initialise sheets on first run ─────────────────────────────────────────

try:
    ensure_sheets()
except Exception as e:
    st.error(f"Could not connect to Google Sheets: {e}")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════

if page == "Dashboard":
    st.title("POC Dashboard")

    pocs = load_pocs()

    if pocs.empty:
        st.info("No POCs yet. Use 'New POC' to add one.")
    else:
        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total POCs", len(pocs))
        col2.metric("Active", len(pocs[pocs["Status"] == "Active"]))
        col3.metric("At Risk", len(pocs[pocs["Status"] == "At Risk"]))
        col4.metric("Won", len(pocs[pocs["Status"] == "Won"]))

        st.markdown("---")

        # Filter
        filter_status = st.multiselect(
            "Filter by status", STATUS_OPTIONS, default=STATUS_OPTIONS
        )
        filtered = pocs[pocs["Status"].isin(filter_status)] if "Status" in pocs.columns else pocs

        # Display table
        display_cols = [
            "Account Name", "Industry", "Snowflake POC Owner",
            "Status", "Start Date", "Target End Date", "Use Case Summary",
        ]
        available = [c for c in display_cols if c in filtered.columns]

        def fmt_status(s):
            return f"{STATUS_COLORS.get(s, '⚪')} {s}"

        view = filtered[available].copy()
        if "Status" in view.columns:
            view["Status"] = view["Status"].apply(fmt_status)

        st.dataframe(view, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════
# PAGE: POC DETAIL
# ══════════════════════════════════════════════════════════════════════════

elif page == "POC Detail":
    st.title("POC Detail")

    pocs = load_pocs()

    if pocs.empty or "Account Name" not in pocs.columns:
        st.info("No POCs yet. Create one first.")
        st.stop()

    account = st.selectbox("Select account", pocs["Account Name"].tolist())
    rec = pocs[pocs["Account Name"] == account].iloc[0].to_dict()

    tab1, tab2, tab3 = st.tabs(["Overview", "Action Plan", "Updates"])

    # ── Tab 1: Overview ───────────────────────────────────────────────────
    with tab1:
        st.subheader("POC Details")
        with st.form("edit_poc"):
            c1, c2 = st.columns(2)
            with c1:
                industry = st.selectbox(
                    "Industry",
                    INDUSTRY_OPTIONS,
                    index=INDUSTRY_OPTIONS.index(rec.get("Industry", "Other"))
                    if rec.get("Industry") in INDUSTRY_OPTIONS else 0,
                )
                cloud = st.selectbox(
                    "Cloud Platform",
                    CLOUD_OPTIONS,
                    index=CLOUD_OPTIONS.index(rec.get("Cloud Platform", "AWS"))
                    if rec.get("Cloud Platform") in CLOUD_OPTIONS else 0,
                )
                se_owner = st.text_input("Snowflake POC Owner", value=rec.get("Snowflake POC Owner", ""))
                champion = st.text_input("Customer Champion", value=rec.get("Customer Champion", ""))
                champion_email = st.text_input("Champion Email", value=rec.get("Champion Email", ""))
            with c2:
                start_raw = rec.get("Start Date", "")
                end_raw = rec.get("Target End Date", "")
                try:
                    start_val = datetime.strptime(start_raw, "%Y-%m-%d").date() if start_raw else date.today()
                except ValueError:
                    start_val = date.today()
                try:
                    end_val = datetime.strptime(end_raw, "%Y-%m-%d").date() if end_raw else date.today()
                except ValueError:
                    end_val = date.today()

                start_date = st.date_input("Start Date", value=start_val)
                end_date = st.date_input("Target End Date", value=end_val)
                status = st.selectbox(
                    "Status",
                    STATUS_OPTIONS,
                    index=STATUS_OPTIONS.index(rec.get("Status", "Planning"))
                    if rec.get("Status") in STATUS_OPTIONS else 0,
                )

            use_case = st.text_area("Use Case Summary", value=rec.get("Use Case Summary", ""), height=80)
            success_criteria = st.text_area(
                "Success Criteria", value=rec.get("Success Criteria", ""), height=100,
                help="What does a successful POC look like? E.g. query latency < 2s, 3 use cases validated."
            )

            if st.form_submit_button("Save Changes", type="primary"):
                updated = {
                    "Account Name": account,
                    "Industry": industry,
                    "Cloud Platform": cloud,
                    "Snowflake POC Owner": se_owner,
                    "Customer Champion": champion,
                    "Champion Email": champion_email,
                    "Start Date": str(start_date),
                    "Target End Date": str(end_date),
                    "Status": status,
                    "Use Case Summary": use_case,
                    "Success Criteria": success_criteria,
                }
                update_poc_row(account, updated)
                st.success("Saved.")
                st.rerun()

    # ── Tab 2: Action Plan ────────────────────────────────────────────────
    with tab2:
        st.subheader("Mutual Action Plan")
        actions = load_actions()
        acct_actions = (
            actions[actions["Account Name"] == account]
            if not actions.empty and "Account Name" in actions.columns
            else pd.DataFrame(columns=ACTION_HEADERS)
        )

        if not acct_actions.empty:
            # Show editable rows
            for idx, row in acct_actions.iterrows():
                ws_row = idx + 2  # sheet row (1-indexed + header)
                with st.expander(f"{row.get('Action', 'Action')} — {row.get('Owner', '')} — {row.get('Status', '')}"):
                    with st.form(f"action_{idx}"):
                        a_action = st.text_input("Action", value=row.get("Action", ""))
                        a_col1, a_col2, a_col3 = st.columns(3)
                        with a_col1:
                            a_owner = st.selectbox(
                                "Owner",
                                ACTION_OWNER_OPTIONS,
                                index=ACTION_OWNER_OPTIONS.index(row.get("Owner", "Snowflake"))
                                if row.get("Owner") in ACTION_OWNER_OPTIONS else 0,
                            )
                        with a_col2:
                            try:
                                due_val = datetime.strptime(row.get("Due Date", ""), "%Y-%m-%d").date()
                            except (ValueError, TypeError):
                                due_val = date.today()
                            a_due = st.date_input("Due Date", value=due_val, key=f"due_{idx}")
                        with a_col3:
                            a_status = st.selectbox(
                                "Status",
                                ACTION_STATUS_OPTIONS,
                                index=ACTION_STATUS_OPTIONS.index(row.get("Status", "Open"))
                                if row.get("Status") in ACTION_STATUS_OPTIONS else 0,
                                key=f"astatus_{idx}",
                            )
                        if st.form_submit_button("Update"):
                            update_action_row(ws_row, {
                                "Account Name": account,
                                "Action": a_action,
                                "Owner": a_owner,
                                "Due Date": str(a_due),
                                "Status": a_status,
                            })
                            st.success("Updated.")
                            st.rerun()
        else:
            st.caption("No action items yet.")

        st.markdown("---")
        st.subheader("Add Action Item")
        with st.form("new_action"):
            new_action = st.text_input("Action")
            nc1, nc2, nc3 = st.columns(3)
            with nc1:
                new_owner = st.selectbox("Owner", ACTION_OWNER_OPTIONS)
            with nc2:
                new_due = st.date_input("Due Date", value=date.today())
            with nc3:
                new_astatus = st.selectbox("Status", ACTION_STATUS_OPTIONS)
            if st.form_submit_button("Add", type="primary"):
                if new_action.strip():
                    append_row("Action_Items", [account, new_action.strip(), new_owner, str(new_due), new_astatus])
                    st.success("Added.")
                    st.rerun()
                else:
                    st.warning("Action text is required.")

    # ── Tab 3: Updates ────────────────────────────────────────────────────
    with tab3:
        st.subheader("Status Updates")
        updates = load_updates()
        acct_updates = (
            updates[updates["Account Name"] == account].sort_values("Date", ascending=False)
            if not updates.empty and "Account Name" in updates.columns
            else pd.DataFrame(columns=UPDATE_HEADERS)
        )

        with st.form("new_update"):
            posted_by = st.text_input("Your name")
            update_text = st.text_area("Update", height=100, placeholder="What happened? Any blockers, wins, next steps?")
            if st.form_submit_button("Post Update", type="primary"):
                if update_text.strip() and posted_by.strip():
                    append_row("Updates", [
                        account,
                        datetime.today().strftime("%Y-%m-%d %H:%M"),
                        update_text.strip(),
                        posted_by.strip(),
                    ])
                    st.success("Posted.")
                    st.rerun()
                else:
                    st.warning("Name and update text are required.")

        st.markdown("---")
        if not acct_updates.empty:
            for _, u in acct_updates.iterrows():
                st.markdown(f"**{u.get('Date', '')}** — *{u.get('Posted By', '')}*")
                st.markdown(u.get("Update", ""))
                st.markdown("---")
        else:
            st.caption("No updates yet.")


# ══════════════════════════════════════════════════════════════════════════
# PAGE: NEW POC
# ══════════════════════════════════════════════════════════════════════════

elif page == "New POC":
    st.title("New POC")

    with st.form("new_poc"):
        c1, c2 = st.columns(2)
        with c1:
            account_name = st.text_input("Account Name *")
            industry = st.selectbox("Industry", INDUSTRY_OPTIONS)
            cloud = st.selectbox("Cloud Platform", CLOUD_OPTIONS)
            se_owner = st.text_input("Snowflake POC Owner *")
        with c2:
            champion = st.text_input("Customer Champion")
            champion_email = st.text_input("Champion Email")
            start_date = st.date_input("Start Date", value=date.today())
            end_date = st.date_input("Target End Date", value=date.today())
            status = st.selectbox("Status", STATUS_OPTIONS)

        use_case = st.text_area("Use Case Summary", height=80)
        success_criteria = st.text_area(
            "Success Criteria", height=100,
            placeholder="What does a successful POC look like?"
        )

        submitted = st.form_submit_button("Create POC", type="primary")
        if submitted:
            if not account_name.strip() or not se_owner.strip():
                st.warning("Account Name and POC Owner are required.")
            else:
                pocs = load_pocs()
                if not pocs.empty and account_name.strip() in pocs["Account Name"].values:
                    st.error(f"'{account_name}' already exists. Edit it from POC Detail.")
                else:
                    append_row("POCs", [
                        account_name.strip(), industry, cloud,
                        se_owner.strip(), champion, champion_email,
                        str(start_date), str(end_date), status,
                        use_case, success_criteria,
                    ])
                    st.success(f"POC created for {account_name}. Go to 'POC Detail' to manage it.")
