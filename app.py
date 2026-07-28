"""
DCT 2026 Data Collection Tool
-----------------------------
A Streamlit web app for entering the PEPFAR/DHIS2-style ART indicators
(TX_ML, TX_CURR, TX_NEW, TX_PVLS, TX_RTT, DSDM VLC/VLS) contained in
DCTs_2026.xlsx, and saving each submission to a Google Sheet.

Setup (see README.md for full walkthrough):
    1. Create a Google Cloud service account with the Sheets + Drive APIs
       enabled, and download its JSON key.
    2. Create a Google Sheet and share it (Editor access) with the service
       account's email address.
    3. Put the service account JSON and the sheet's key/name into
       .streamlit/secrets.toml (locally) or the app's "Secrets" panel
       (on Streamlit Community Cloud) -- see README.md for the exact format.

Run with:
    streamlit run app.py
"""

import hashlib
import secrets as pysecrets
import uuid
from datetime import datetime, date

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SUBMISSIONS_HEADER = ["submission_id", "facility", "org_unit", "period",
                       "entered_by", "submitted_at"]
ENTRIES_HEADER = ["submission_id", "sheet", "table_name", "row_label",
                   "col_label", "value"]
FACILITIES_HEADER = ["facility_name", "org_unit"]
USERS_HEADER = ["username", "salt", "password_hash", "full_name", "role"]
DEFAULT_ADMIN_PASSWORD = "Admin@123"

# ---------------------------------------------------------------------------
# Reference lists (mirrors the disaggregations in DCTs_2026.xlsx)
# ---------------------------------------------------------------------------

AGE_BANDS_15 = [
    "< 1 year", "1-4 years", "5-9 years", "10-14 years", "15-19 years",
    "20-24 years", "25-29 years", "30-34 years", "35-39 years", "40-44 years",
    "45-49 years", "50-54 years", "55-59 years", "60-64 years", "65+ years",
]

AGE_BANDS_10 = [
    "< 1 year", "1-4 years", "5-9 years", "10-14 years", "15-19 years",
    "20-24 years", "25-29 years", "30-39 years", "40-49 years", "50+ years",
]

TX_ML_OUTCOMES = [
    "Died",
    "Transfer Out (Silent)",
    "IIT after <3 months on treatment",
    "IIT after 3-5 months on treatment",
    "IIT after >5 months on treatment",
    "Refused (Stopped) Treatment",
]

TX_CURR_DISPENSING = [
    "<3 months of ARVs dispensed, <15yrs",
    "<3 months of ARVs dispensed, 15+yrs",
    "3-5 months of ARVs dispensed, <15yrs",
    "3-5 months of ARVs dispensed, 15+yrs",
    "6+ months of ARVs dispensed, <15yrs",
    "6+ months of ARVs dispensed, 15+yrs",
]

TX_CURR_DTG = ["<15yrs", "15+yrs"]

TX_RTT_DURATION = [
    "Treatment interruption <3 months before returning",
    "Treatment interruption 3-5 months before returning",
    "Treatment interruption 6+ months before returning",
]

PVLS_CATEGORIES = ["Eligible", "Tested", "Suppressed Viral Load"]

DSDM_MODELS = [
    "Fast Track Drug Refill (FTDR)",
    "Facility Based Individual Management (FBIM)",
    "Community Drug Distribution Point (CDDP)",
    "Community Client Led ART Distribution (CCLAD)",
    "Facility Based Groups (FBGs)",
    "Community Retail Pharmacy Drug Distribution Point (CRPDDP)",
]

SHEETS = ["TX_ML", "TX_CURR", "TX_NEW", "TX_PVLS", "TX_RTT", "DSDM_VLC-VLS"]

# ---------------------------------------------------------------------------
# Google Sheets backend
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def get_client():
    """Authorize a gspread client from Streamlit secrets."""
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    client = get_client()
    if "sheet_key" in st.secrets:
        return client.open_by_key(st.secrets["sheet_key"])
    return client.open(st.secrets["sheet_name"])


def init_db():
    """Ensure the 'submissions', 'entries', 'facilities', and 'users' worksheets exist."""
    sh = get_spreadsheet()
    existing = {ws.title for ws in sh.worksheets()}
    if "submissions" not in existing:
        ws = sh.add_worksheet(title="submissions", rows=1000, cols=len(SUBMISSIONS_HEADER))
        ws.append_row(SUBMISSIONS_HEADER)
    if "entries" not in existing:
        ws = sh.add_worksheet(title="entries", rows=5000, cols=len(ENTRIES_HEADER))
        ws.append_row(ENTRIES_HEADER)
    if "facilities" not in existing:
        ws = sh.add_worksheet(title="facilities", rows=200, cols=len(FACILITIES_HEADER))
        ws.append_row(FACILITIES_HEADER)
        ws.append_row(["Example Health Centre IV", "EX001"])
    if "users" not in existing:
        ws = sh.add_worksheet(title="users", rows=200, cols=len(USERS_HEADER))
        ws.append_row(USERS_HEADER)
        salt, pw_hash = hash_password(DEFAULT_ADMIN_PASSWORD)
        ws.append_row(["admin", salt, pw_hash, "Administrator", "admin"])


def hash_password(password: str, salt: str = None):
    if salt is None:
        salt = pysecrets.token_hex(8)
    digest = hashlib.sha256((salt + password).encode("utf-8")).hexdigest()
    return salt, digest


@st.cache_data(ttl=60, show_spinner=False)
def load_users():
    sh = get_spreadsheet()
    records = sh.worksheet("users").get_all_records()
    df = pd.DataFrame(records, columns=USERS_HEADER)
    return df


def authenticate(username: str, password: str):
    users = load_users()
    match = users[users["username"] == username]
    if match.empty:
        return None
    row = match.iloc[0]
    _, digest = hash_password(password, row["salt"])
    if digest == row["password_hash"]:
        return {"username": row["username"], "full_name": row["full_name"], "role": row["role"]}
    return None


def add_user(username: str, password: str, full_name: str, role: str):
    sh = get_spreadsheet()
    salt, pw_hash = hash_password(password)
    sh.worksheet("users").append_row([username, salt, pw_hash, full_name, role])
    load_users.clear()


@st.cache_data(ttl=300, show_spinner=False)
def load_facilities():
    """Read the facility list from the 'facilities' tab. Cached 5 min so the
    Sheet can be edited without a code change; use the Refresh button to
    pick up changes sooner."""
    sh = get_spreadsheet()
    records = sh.worksheet("facilities").get_all_records()
    df = pd.DataFrame(records, columns=FACILITIES_HEADER)
    df = df[df["facility_name"].astype(str).str.strip() != ""]
    return df


def save_submission(meta: dict, tables: dict):
    """tables: {(sheet, table_name): DataFrame(index=row_label, columns=col_label)}"""
    sh = get_spreadsheet()
    sub_id = str(uuid.uuid4())

    sh.worksheet("submissions").append_row([
        sub_id,
        meta["facility"],
        meta["org_unit"],
        meta["period"],
        meta["entered_by"],
        datetime.now().isoformat(timespec="seconds"),
    ])

    rows = []
    for (sheet, table_name), df in tables.items():
        melted = df.reset_index().melt(id_vars=df.index.name or "index",
                                        var_name="col_label", value_name="value")
        melted.columns = ["row_label", "col_label", "value"]
        for _, r in melted.iterrows():
            rows.append([
                sub_id, sheet, table_name, r["row_label"], r["col_label"],
                float(r["value"]) if pd.notna(r["value"]) else 0.0,
            ])
    if rows:
        sh.worksheet("entries").append_rows(rows, value_input_option="RAW")
    return sub_id


def load_submissions():
    sh = get_spreadsheet()
    records = sh.worksheet("submissions").get_all_records()
    df = pd.DataFrame(records, columns=SUBMISSIONS_HEADER)
    if not df.empty:
        df = df.sort_values("submitted_at", ascending=False)
    return df


def load_entries(submission_id):
    sh = get_spreadsheet()
    records = sh.worksheet("entries").get_all_records()
    df = pd.DataFrame(records, columns=ENTRIES_HEADER)
    if not df.empty:
        df = df[df["submission_id"] == submission_id]
    return df


# ---------------------------------------------------------------------------
# Reusable entry grid
# ---------------------------------------------------------------------------

def entry_grid(key, rows, columns, index_name="Disaggregation"):
    df = pd.DataFrame(0, index=rows, columns=columns)
    df.index.name = index_name
    edited = st.data_editor(
        df,
        key=key,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            c: st.column_config.NumberColumn(c, min_value=0, step=1, format="%d")
            for c in columns
        },
    )
    return edited


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

st.set_page_config(page_title="DCT 2026 Data Collection", layout="wide")
init_db()

if "auth" not in st.session_state:
    st.session_state.auth = None

if st.session_state.auth is None:
    st.title("DCT 2026 \u2014 Sign in")
    with st.form("login_form"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in", type="primary")
    if submitted:
        user = authenticate(u.strip(), p)
        if user:
            st.session_state.auth = user
            st.rerun()
        else:
            st.error("Incorrect username or password.")
    st.caption(
        "First time here? The default admin account is username **admin**, "
        f"password **{DEFAULT_ADMIN_PASSWORD}** — log in with that, then go to "
        "'Manage users' to add real accounts and change this password."
    )
    st.stop()

st.title("DCT 2026 \u2014 ART Program Data Collection Tool")
st.caption(
    "Enter facility-level ART indicators (TX_ML, TX_CURR, TX_NEW, TX_PVLS, "
    "TX_RTT, DSDM VLC/VLS) and submit. Every submission is saved locally so "
    "you can review or export it later."
)

with st.sidebar:
    st.success(f"Signed in as **{st.session_state.auth['full_name'] or st.session_state.auth['username']}**")
    if st.button("Log out"):
        st.session_state.auth = None
        st.rerun()
    st.divider()

    st.header("Reporting details")

    facilities_df = load_facilities()
    if facilities_df.empty:
        st.warning(
            "No facilities found in the 'facilities' tab of the Google Sheet. "
            "Add rows there (facility_name, org_unit) then hit Refresh."
        )
        facility = st.text_input("Facility name")
        org_unit = st.text_input("Org unit / facility code")
    else:
        facility = st.selectbox(
            "Facility name", options=facilities_df["facility_name"].tolist()
        )
        matched_unit = facilities_df.loc[
            facilities_df["facility_name"] == facility, "org_unit"
        ]
        org_unit = matched_unit.iloc[0] if not matched_unit.empty else ""
        st.caption(f"Org unit: {org_unit or '(none set)'}")

    if st.button("\U0001F504 Refresh facility list"):
        load_facilities.clear()
        st.rerun()

    period = st.text_input("Reporting period (e.g. 2026-Q2 or July 2026)",
                            value=date.today().strftime("%Y-%m"))
    entered_by = st.text_input("Entered by")
    st.divider()
    nav_options = ["Data entry", "Submission history"]
    if st.session_state.auth["role"] == "admin":
        nav_options.append("Manage users")
    nav = st.radio("View", nav_options)

if nav == "Manage users":
    st.subheader("Manage users")
    st.caption("Add accounts for staff who'll be entering data. Only admins can see this page.")
    with st.form("add_user_form", clear_on_submit=True):
        col_a, col_b = st.columns(2)
        new_username = col_a.text_input("Username")
        new_password = col_b.text_input("Temporary password", type="password")
        new_full_name = col_a.text_input("Full name")
        new_role = col_b.selectbox("Role", ["user", "admin"])
        add_submitted = st.form_submit_button("Add user", type="primary")
    if add_submitted:
        existing_users = load_users()
        if not new_username or not new_password:
            st.error("Username and password are required.")
        elif new_username in existing_users["username"].values:
            st.error("That username already exists.")
        else:
            add_user(new_username, new_password, new_full_name, new_role)
            st.success(f"User '{new_username}' added.")
            st.rerun()

    st.markdown("**Existing users**")
    users_display = load_users()[["username", "full_name", "role"]]
    st.dataframe(users_display, use_container_width=True, hide_index=True)
    st.caption(
        "To remove or reset a user, edit the 'users' tab directly in the Google "
        "Sheet (delete their row to remove access)."
    )
    st.stop()

if nav == "Submission history":
    st.subheader("Past submissions")
    subs = load_submissions()
    if subs.empty:
        st.info("No submissions yet.")
    else:
        st.dataframe(subs, use_container_width=True, hide_index=True)
        chosen = st.selectbox(
            "View details for submission",
            subs["submission_id"],
            format_func=lambda sid: f"{sid[:8]} \u2014 "
            f"{subs.loc[subs.submission_id == sid, 'facility'].iloc[0]} "
            f"({subs.loc[subs.submission_id == sid, 'period'].iloc[0]})",
        )
        if chosen:
            detail = load_entries(chosen)
            st.dataframe(detail, use_container_width=True, hide_index=True)
            st.download_button(
                "Download this submission as CSV",
                detail.to_csv(index=False).encode("utf-8"),
                file_name=f"submission_{chosen[:8]}.csv",
                mime="text/csv",
            )
    st.stop()

# ---- Data entry mode ----
tabs = st.tabs(SHEETS)
tables = {}

with tabs[0]:
    st.subheader("TX_ML \u2014 Interruption in Treatment / Outcomes")
    st.caption(
        "Enter Female/Male counts by age band for each outcome. The combined "
        "total across all outcomes, by age and sex, is calculated automatically below."
    )

    outcome_tables = {}
    for outcome in TX_ML_OUTCOMES:
        st.markdown(f"**{outcome}**")
        df = entry_grid(
            f"txml_{outcome}", AGE_BANDS_15, ["Female", "Male"], "Age band"
        )
        outcome_tables[outcome] = df
        tables[("TX_ML", outcome)] = df

    # Auto-computed: combined total across all outcome sub-categories, by age & sex
    combined = sum(outcome_tables.values())
    combined.index.name = "Age band"
    tables[("TX_ML", "Combined total (all outcomes) by age and sex")] = combined

    st.markdown("**Combined total across all outcomes \u2014 by age and sex** _(auto-calculated)_")
    combined_display = combined.copy()
    combined_display["Total"] = combined_display["Female"] + combined_display["Male"]
    st.dataframe(combined_display, use_container_width=True)

    totals_row = combined[["Female", "Male"]].sum()
    st.markdown(
        f"**Grand total \u2014 Female: {int(totals_row['Female'])} | "
        f"Male: {int(totals_row['Male'])} | "
        f"Overall: {int(totals_row.sum())}**"
    )

with tabs[1]:
    st.subheader("TX_CURR \u2014 Currently on ART")
    st.markdown("**By age and sex**")
    tables[("TX_CURR", "By age and sex")] = entry_grid(
        "txcurr_age", AGE_BANDS_15, ["Female", "Male"], "Age band"
    )
    st.markdown("**ARV dispensing quantity (Total)**")
    tables[("TX_CURR", "ARV dispensing quantity")] = entry_grid(
        "txcurr_dispense", TX_CURR_DISPENSING, ["Total"], "Dispensing category"
    )
    st.markdown("**On DTG-based regimen (Total)**")
    tables[("TX_CURR", "DTG regimen")] = entry_grid(
        "txcurr_dtg", TX_CURR_DTG, ["Total"], "Age/weight band"
    )

with tabs[2]:
    st.subheader("TX_NEW \u2014 Newly Enrolled on ART")
    tables[("TX_NEW", "By age and sex")] = entry_grid(
        "txnew_age", AGE_BANDS_15, ["Female", "Male"], "Age band"
    )

with tabs[3]:
    st.subheader("TX_PVLS \u2014 Viral Load Coverage & Suppression")
    for cat in PVLS_CATEGORIES:
        st.markdown(f"**{cat}**")
        tables[("TX_PVLS", cat)] = entry_grid(
            f"pvls_{cat.replace(' ', '_')}", AGE_BANDS_10, ["Female", "Male"], "Age band"
        )

with tabs[4]:
    st.subheader("TX_RTT \u2014 Returned to Treatment")
    st.markdown("**By age and sex**")
    tables[("TX_RTT", "By age and sex")] = entry_grid(
        "txrtt_age", AGE_BANDS_15, ["Female", "Male"], "Age band"
    )
    st.markdown("**Duration of treatment interruption before returning (Total)**")
    tables[("TX_RTT", "Duration before returning")] = entry_grid(
        "txrtt_duration", TX_RTT_DURATION, ["Total"], "Duration"
    )

with tabs[5]:
    st.subheader("DSDM VLC/VLS \u2014 Differentiated Service Delivery Models")
    st.markdown("**Active on DSD, by age and model**")
    tables[("DSDM_VLC-VLS", "Active on DSD")] = entry_grid(
        "dsdm_active", AGE_BANDS_10, DSDM_MODELS, "Age band"
    )
    for cat in PVLS_CATEGORIES:
        st.markdown(f"**{cat}, by age and model**")
        tables[("DSDM_VLC-VLS", cat)] = entry_grid(
            f"dsdm_{cat.replace(' ', '_')}", AGE_BANDS_10, DSDM_MODELS, "Age band"
        )

st.divider()
col1, col2 = st.columns([1, 3])
with col1:
    submit = st.button("Submit report", type="primary", use_container_width=True)

if submit:
    if not facility or not period:
        st.error("Please enter at least a facility name and reporting period in the sidebar.")
    else:
        meta = {
            "facility": facility,
            "org_unit": org_unit,
            "period": period,
            "entered_by": entered_by,
        }
        sub_id = save_submission(meta, tables)
        st.success(f"Report submitted and saved (ID: {sub_id[:8]}).")
