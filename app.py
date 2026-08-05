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
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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

IIT_OUTCOMES = [
    "IIT after <3 months on treatment",
    "IIT after 3-5 months on treatment",
    "IIT after >5 months on treatment",
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

PREP_CATEGORIES = [
    "Screened for PREP",
    "Screened and eligible for PREP",
    "Screened, eligible and initiated on PREP",
    "Currently on PREP (PREP_CT)",
]

SHEETS = ["TX_ML", "TX_CURR", "TX_NEW", "TX_PVLS", "TX_RTT", "DSDM_VLC-VLS", "PREP_BF_PREG"]

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


@st.cache_data(ttl=60, show_spinner=False)
def load_all_data():
    """All entries joined with their submission's facility/period/etc, for the dashboard."""
    sh = get_spreadsheet()
    subs = pd.DataFrame(sh.worksheet("submissions").get_all_records(), columns=SUBMISSIONS_HEADER)
    entries = pd.DataFrame(sh.worksheet("entries").get_all_records(), columns=ENTRIES_HEADER)
    if subs.empty or entries.empty:
        return pd.DataFrame(columns=list(ENTRIES_HEADER) +
                             ["facility", "org_unit", "period", "entered_by", "submitted_at"])
    merged = entries.merge(subs, on="submission_id", how="left")
    merged["value"] = pd.to_numeric(merged["value"], errors="coerce").fillna(0)
    return merged


def load_submission_tables(submission_id):
    """Reshape one submission's flat entry rows back into {(sheet, table_name): DataFrame}
    for pre-filling the entry grids when editing."""
    entries = load_entries(submission_id)
    result = {}
    if entries.empty:
        return result
    for (sheet, table_name), grp in entries.groupby(["sheet", "table_name"]):
        pivot = grp.pivot_table(index="row_label", columns="col_label", values="value",
                                 aggfunc="sum", fill_value=0)
        result[(sheet, table_name)] = pivot
    return result


def update_submission(submission_id, meta, tables):
    """Overwrite an existing submission in place: same submission_id, new values."""
    sh = get_spreadsheet()

    subs_ws = sh.worksheet("submissions")
    subs_df = pd.DataFrame(subs_ws.get_all_records(), columns=SUBMISSIONS_HEADER)
    subs_df = subs_df[subs_df["submission_id"] != submission_id]
    new_sub_row = {
        "submission_id": submission_id,
        "facility": meta["facility"],
        "org_unit": meta["org_unit"],
        "period": meta["period"],
        "entered_by": meta["entered_by"],
        "submitted_at": datetime.now().isoformat(timespec="seconds") + " (edited)",
    }
    subs_df = pd.concat([subs_df, pd.DataFrame([new_sub_row])], ignore_index=True)
    subs_ws.clear()
    subs_ws.append_row(SUBMISSIONS_HEADER)
    if not subs_df.empty:
        subs_ws.append_rows(subs_df[SUBMISSIONS_HEADER].values.tolist(), value_input_option="RAW")

    entries_ws = sh.worksheet("entries")
    entries_df = pd.DataFrame(entries_ws.get_all_records(), columns=ENTRIES_HEADER)
    entries_df = entries_df[entries_df["submission_id"] != submission_id]

    new_rows = []
    for (sheet, table_name), df in tables.items():
        melted = df.reset_index().melt(id_vars=df.index.name or "index",
                                        var_name="col_label", value_name="value")
        melted.columns = ["row_label", "col_label", "value"]
        for _, r in melted.iterrows():
            new_rows.append({
                "submission_id": submission_id, "sheet": sheet, "table_name": table_name,
                "row_label": r["row_label"], "col_label": r["col_label"],
                "value": float(r["value"]) if pd.notna(r["value"]) else 0.0,
            })
    entries_df = pd.concat([entries_df, pd.DataFrame(new_rows)], ignore_index=True)
    entries_ws.clear()
    entries_ws.append_row(ENTRIES_HEADER)
    if not entries_df.empty:
        entries_ws.append_rows(entries_df[ENTRIES_HEADER].values.tolist(), value_input_option="RAW")

    load_all_data.clear()
    return submission_id


# ---------------------------------------------------------------------------
# Reusable entry grid
# ---------------------------------------------------------------------------


def entry_grid(key, rows, columns, index_name="Disaggregation", initial_data=None):
    if initial_data is not None and not initial_data.empty:
        df = initial_data.reindex(index=rows, columns=columns).fillna(0).astype(int)
    else:
        df = pd.DataFrame(0, index=rows, columns=columns)
    df.index.name = index_name
    edited = st.data_editor(
        df,
        key=key,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            c: st.column_config.NumberColumn(c, min_value=0, step=1, format="%d",
                                              alignment="center")
            for c in columns
        },
    )
    return edited


def tracked_entry_grid(tables_dict, sheet, table_name, base_key, rows, columns,
                        index_name="Disaggregation"):
    """Wraps entry_grid: in edit mode, gives each widget a submission-specific key
    and pre-fills it from the loaded submission; also records the result into
    tables_dict[(sheet, table_name)]."""
    edit_id = st.session_state.get("editing_submission_id")
    key = f"{base_key}_edit_{edit_id}" if edit_id else base_key
    initial = None
    if edit_id and st.session_state.get("edit_source_tables"):
        initial = st.session_state.edit_source_tables.get((sheet, table_name))
    df = entry_grid(key, rows, columns, index_name, initial_data=initial)
    tables_dict[(sheet, table_name)] = df
    return df


def month_options(years_back=1, years_forward=1):
    """['January 2025', 'February 2025', ..., 'December 2027'], newest last."""
    today = date.today()
    start_year = today.year - years_back
    end_year = today.year + years_forward
    opts = []
    for year in range(start_year, end_year + 1):
        for m in range(1, 13):
            opts.append(date(year, m, 1).strftime("%B %Y"))
    return opts


def quarter_options(years_back=1, years_forward=1):
    """['Q1 2025 (Jan-Mar)', ..., 'Q4 2027 (Oct-Dec)']"""
    today = date.today()
    start_year = today.year - years_back
    end_year = today.year + years_forward
    labels = {1: "Jan-Mar", 2: "Apr-Jun", 3: "Jul-Sep", 4: "Oct-Dec"}
    opts = []
    for year in range(start_year, end_year + 1):
        for q in range(1, 5):
            opts.append(f"Q{q} {year} ({labels[q]})")
    return opts


def current_quarter_label():
    today = date.today()
    q = (today.month - 1) // 3 + 1
    labels = {1: "Jan-Mar", 2: "Apr-Jun", 3: "Jul-Sep", 4: "Oct-Dec"}
    return f"Q{q} {today.year} ({labels[q]})"


# ---------------------------------------------------------------------------
# Data quality checks
# ---------------------------------------------------------------------------

def iit_vs_returned_chart(categories, iit_vals, rtt_vals, title, xaxis_title, target_pct=50):
    """Grouped bar of IIT vs Returned (left axis, counts) plus a Return Rate %
    line (right axis, percent) and a dashed target-rate reference line."""
    return_rate = [(r / i * 100) if i else 0 for i, r in zip(iit_vals, rtt_vals)]
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=categories, y=iit_vals, name="IIT (TX_ML)", secondary_y=False)
    fig.add_bar(x=categories, y=rtt_vals, name="Returned (TX_RTT)", secondary_y=False)
    fig.add_trace(
        go.Scatter(
            x=categories, y=return_rate, mode="lines+markers+text", name="Return Rate %",
            text=[f"{r:.0f}%" for r in return_rate], textposition="top center",
            line=dict(color="black", dash="dot"),
        ),
        secondary_y=True,
    )
    fig.add_hline(
        y=target_pct, line_dash="dash", line_color="red", secondary_y=True,
        annotation_text=f"Target: {target_pct}%", annotation_position="top left",
    )
    fig.update_layout(title=title, barmode="group", xaxis_title=xaxis_title)
    fig.update_yaxes(title_text="Clients", secondary_y=False)
    fig.update_yaxes(title_text="Return Rate (%)", secondary_y=True, range=[0, 110])
    return fig


def pct_color(val):
    """Green >= 90%, yellow 84-90%, red < 84%. Used to color-code coverage/rate pivot tables."""
    if val >= 90:
        bg = "#c6efce"  # green
    elif val >= 84:
        bg = "#ffeb9c"  # yellow
    else:
        bg = "#ffc7ce"  # red
    return f"background-color: {bg}"


def styled_pct_table(df, pct_columns, int_columns):
    """Apply pct_color to percentage columns and number formatting, version-safe across pandas."""
    styler = df.style
    style_fn = styler.map if hasattr(styler, "map") else styler.applymap
    fmt = {c: "{:.0f}" for c in int_columns}
    fmt.update({c: "{:.1f}%" for c in pct_columns})
    return style_fn(pct_color, subset=pct_columns).format(fmt)


def run_quality_checks(tables: dict, meta: dict):
    """Returns (errors, warnings). Errors block submission; warnings need confirmation."""
    errors, warnings = [], []

    if not meta["facility"]:
        errors.append("Facility name is required.")
    if not meta["period"]:
        errors.append("Reporting period is required.")
    if not meta["entered_by"]:
        warnings.append("'Entered by' is blank \u2014 consider recording who filled this in.")

    # --- Logical hierarchy: Suppressed <= Tested <= Eligible (TX_PVLS) ---
    try:
        eligible = tables[("TX_PVLS", "Eligible")]
        tested = tables[("TX_PVLS", "Tested")]
        suppressed = tables[("TX_PVLS", "Suppressed Viral Load")]
        for age in eligible.index:
            for sex in ["Female", "Male"]:
                e, t, s = eligible.loc[age, sex], tested.loc[age, sex], suppressed.loc[age, sex]
                if t > e:
                    errors.append(f"TX_PVLS ({sex}, {age}): Tested ({t}) exceeds Eligible ({e}).")
                if s > t:
                    errors.append(f"TX_PVLS ({sex}, {age}): Suppressed ({s}) exceeds Tested ({t}).")
    except KeyError:
        pass

    # --- Logical hierarchy: Suppressed <= Tested <= Eligible (DSDM, per model) ---
    try:
        d_eligible = tables[("DSDM_VLC-VLS", "Eligible")]
        d_tested = tables[("DSDM_VLC-VLS", "Tested")]
        d_suppressed = tables[("DSDM_VLC-VLS", "Suppressed Viral Load")]
        for age in d_eligible.index:
            for model in d_eligible.columns:
                e, t, s = d_eligible.loc[age, model], d_tested.loc[age, model], d_suppressed.loc[age, model]
                if t > e:
                    errors.append(f"DSDM ({model}, {age}): Tested ({t}) exceeds Eligible ({e}).")
                if s > t:
                    errors.append(f"DSDM ({model}, {age}): Suppressed ({s}) exceeds Tested ({t}).")
    except KeyError:
        pass

    # --- Logical hierarchy: Initiated <= Eligible <= Screened (PREP_BF_PREG) ---
    try:
        p_screened = tables[("PREP_BF_PREG", "Screened for PREP")]
        p_eligible = tables[("PREP_BF_PREG", "Screened and eligible for PREP")]
        p_initiated = tables[("PREP_BF_PREG", "Screened, eligible and initiated on PREP")]
        for age in p_screened.index:
            for grp in ["Pregnant", "Breastfeeding"]:
                sc, el, ini = p_screened.loc[age, grp], p_eligible.loc[age, grp], p_initiated.loc[age, grp]
                if el > sc:
                    errors.append(f"PREP_BF_PREG ({grp}, {age}): Eligible ({el}) exceeds Screened ({sc}).")
                if ini > el:
                    errors.append(f"PREP_BF_PREG ({grp}, {age}): Initiated ({ini}) exceeds Eligible ({el}).")
    except KeyError:
        pass

    # --- MMD (ARV dispensing quantity) must reconcile with TX_CURR, by <15/15+ ---
    try:
        txcurr_age = tables[("TX_CURR", "By age and sex")]
        under15_ages = AGE_BANDS_15[:4]   # < 1, 1-4, 5-9, 10-14
        over15_ages = AGE_BANDS_15[4:]    # 15-19 ... 65+

        txcurr_under15 = txcurr_age.loc[under15_ages, ["Female", "Male"]].values.sum()
        txcurr_15plus = txcurr_age.loc[over15_ages, ["Female", "Male"]].values.sum()

        mmd = tables[("TX_CURR", "ARV dispensing quantity")]
        mmd_under15 = mmd.loc[[r for r in mmd.index if "<15yrs" in r], ["Female", "Male"]].values.sum()
        mmd_15plus = mmd.loc[[r for r in mmd.index if "15+yrs" in r], ["Female", "Male"]].values.sum()

        if mmd_under15 != txcurr_under15:
            errors.append(
                f"MMD (<15yrs) total ({int(mmd_under15)}) does not equal TX_CURR <15yrs "
                f"total ({int(txcurr_under15)}). Every client currently on ART should be "
                "counted in exactly one dispensing duration category."
            )
        if mmd_15plus != txcurr_15plus:
            errors.append(
                f"MMD (15+yrs) total ({int(mmd_15plus)}) does not equal TX_CURR 15+yrs "
                f"total ({int(txcurr_15plus)}). Every client currently on ART should be "
                "counted in exactly one dispensing duration category."
            )
    except KeyError:
        pass

    # --- DTG regimen cannot exceed TX_CURR, by <15/15+ ---
    try:
        txcurr_age = tables[("TX_CURR", "By age and sex")]
        under15_ages = AGE_BANDS_15[:4]
        over15_ages = AGE_BANDS_15[4:]
        txcurr_under15 = txcurr_age.loc[under15_ages, ["Female", "Male"]].values.sum()
        txcurr_15plus = txcurr_age.loc[over15_ages, ["Female", "Male"]].values.sum()

        dtg = tables[("TX_CURR", "DTG regimen")]
        dtg_under15 = dtg.loc["<15yrs", ["Female", "Male"]].sum()
        dtg_15plus = dtg.loc["15+yrs", ["Female", "Male"]].sum()

        if dtg_under15 > txcurr_under15:
            errors.append(
                f"DTG regimen (<15yrs) total ({int(dtg_under15)}) exceeds TX_CURR <15yrs "
                f"total ({int(txcurr_under15)}). Clients on DTG are a subset of everyone "
                "currently on ART."
            )
        if dtg_15plus > txcurr_15plus:
            errors.append(
                f"DTG regimen (15+yrs) total ({int(dtg_15plus)}) exceeds TX_CURR 15+yrs "
                f"total ({int(txcurr_15plus)}). Clients on DTG are a subset of everyone "
                "currently on ART."
            )
    except KeyError:
        pass

    # --- Cross-sheet plausibility (soft warnings) ---
    try:
        txcurr_total = tables[("TX_CURR", "By age and sex")].values.sum()
        txnew_total = tables[("TX_NEW", "By age and sex")].values.sum()
        if txnew_total > txcurr_total:
            warnings.append(
                f"TX_NEW total ({int(txnew_total)}) is greater than TX_CURR total "
                f"({int(txcurr_total)}) \u2014 newly enrolled clients are usually a subset "
                "of everyone currently on ART."
            )
        pvls_eligible_total = tables[("TX_PVLS", "Eligible")].values.sum()
        if pvls_eligible_total > txcurr_total:
            warnings.append(
                f"TX_PVLS Eligible total ({int(pvls_eligible_total)}) exceeds TX_CURR total "
                f"({int(txcurr_total)})."
            )
    except KeyError:
        pass

    # --- Completeness: flag sections left entirely at zero ---
    for (sheet, table_name), df in tables.items():
        if df.values.sum() == 0:
            warnings.append(f"{sheet} \u2014 '{table_name}' is all zeros. Confirm this is correct, not just skipped.")

    return errors, warnings


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

LOGO_PATH = "assets/logo.png"

st.set_page_config(page_title="DCT 2026 Data Collection", page_icon=LOGO_PATH, layout="wide")
init_db()

try:
    st.logo(LOGO_PATH, size="large")
except Exception:
    pass

if "auth" not in st.session_state:
    st.session_state.auth = None
if "editing_submission_id" not in st.session_state:
    st.session_state.editing_submission_id = None
if "edit_source_tables" not in st.session_state:
    st.session_state.edit_source_tables = None

# Apply any staged widget values BEFORE those widgets are created this run.
# (Streamlit forbids writing to a widget's session_state key after that
# widget has already been instantiated in the same script run, so the Edit
# button stages values here instead of setting them directly.)
if st.session_state.get("pending_prefill"):
    for _k, _v in st.session_state.pending_prefill.items():
        st.session_state[_k] = _v
    st.session_state.pending_prefill = None

if st.session_state.auth is None:
    st.image(LOGO_PATH, width=220)
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

    if st.session_state.editing_submission_id:
        st.info(f"\u270f\ufe0f Editing submission `{st.session_state.editing_submission_id[:8]}`")
        if st.button("Cancel edit"):
            st.session_state.editing_submission_id = None
            st.session_state.edit_source_tables = None
            st.rerun()
        st.divider()

    st.header("Reporting details")

    facilities_df = load_facilities()
    if facilities_df.empty:
        st.warning(
            "No facilities found in the 'facilities' tab of the Google Sheet. "
            "Add rows there (facility_name, org_unit) then hit Refresh."
        )
        facility = st.text_input("Facility name", key="facility_text")
        org_unit = st.text_input("Org unit / facility code", key="org_unit_text")
    else:
        facility = st.selectbox(
            "Facility name", options=facilities_df["facility_name"].tolist(), key="facility_select"
        )
        matched_unit = facilities_df.loc[
            facilities_df["facility_name"] == facility, "org_unit"
        ]
        org_unit = matched_unit.iloc[0] if not matched_unit.empty else ""
        st.caption(f"Org unit: {org_unit or '(none set)'}")

    if st.button("\U0001F504 Refresh facility list"):
        load_facilities.clear()
        st.rerun()

    period_type = st.radio("Reporting period type", ["Month", "Quarter"],
                            horizontal=True, key="period_type_radio")
    if period_type == "Month":
        months = month_options()
        default_idx = months.index(date.today().strftime("%B %Y"))
        period = st.selectbox("Reporting period", months, index=default_idx, key="period_month_select")
    else:
        quarters = quarter_options()
        default_idx = quarters.index(current_quarter_label())
        period = st.selectbox("Reporting period", quarters, index=default_idx, key="period_quarter_select")

    entered_by = st.text_input("Entered by", key="entered_by_text")
    st.divider()
    nav_options = ["Data entry", "Dashboard", "Submission history"]
    if st.session_state.auth["role"] == "admin":
        nav_options.append("Manage users")
    nav = st.radio("View", nav_options, key="nav_radio")

if nav == "Dashboard":
    st.subheader("Dashboard")
    data = load_all_data()

    if data.empty:
        st.info("No submissions yet. Once reports are submitted, this dashboard will populate.")
        st.stop()

    col_f, col_p = st.columns(2)
    all_facilities = sorted(data["facility"].dropna().unique().tolist())
    all_periods = sorted(data["period"].dropna().unique().tolist())
    facilities_sel = col_f.multiselect("Facility", all_facilities, default=all_facilities)
    periods_sel = col_p.multiselect("Reporting period", all_periods, default=all_periods)

    filtered = data[data["facility"].isin(facilities_sel) & data["period"].isin(periods_sel)]

    if filtered.empty:
        st.warning("No data matches the selected filters.")
        st.stop()

    def total_for(sheet, table_name):
        subset = filtered[(filtered["sheet"] == sheet) & (filtered["table_name"] == table_name)]
        return subset["value"].sum()

    txcurr_total = total_for("TX_CURR", "By age and sex")
    txnew_total = total_for("TX_NEW", "By age and sex")
    txml_total = total_for("TX_ML", "Combined total (all outcomes) by age and sex")
    txrtt_total = total_for("TX_RTT", "By age and sex")
    pvls_eligible = total_for("TX_PVLS", "Eligible")
    pvls_tested = total_for("TX_PVLS", "Tested")
    pvls_suppressed = total_for("TX_PVLS", "Suppressed Viral Load")
    vl_coverage = (pvls_tested / pvls_eligible * 100) if pvls_eligible else 0
    vl_suppression = (pvls_suppressed / pvls_tested * 100) if pvls_tested else 0

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("TX_CURR", f"{int(txcurr_total):,}")
    k2.metric("TX_NEW", f"{int(txnew_total):,}")
    k3.metric("TX_ML (all outcomes)", f"{int(txml_total):,}")
    k4.metric("TX_RTT", f"{int(txrtt_total):,}")
    k5.metric("VL Coverage", f"{vl_coverage:.0f}%")
    k6.metric("VL Suppression", f"{vl_suppression:.0f}%")

    st.divider()

    c1, c2 = st.columns(2)

    with c1:
        trend = (
            filtered[(filtered["sheet"] == "TX_CURR") & (filtered["table_name"] == "By age and sex")]
            .groupby("period")["value"].sum().reset_index()
        )
        if not trend.empty:
            fig = px.line(trend, x="period", y="value", markers=True,
                          title="TX_CURR by reporting period")
            fig.update_layout(yaxis_title="Clients currently on ART", xaxis_title="Period")
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        ml_outcomes = (
            filtered[(filtered["sheet"] == "TX_ML") & (filtered["table_name"].isin(TX_ML_OUTCOMES))]
            .groupby("table_name")["value"].sum().reset_index()
        )
        if not ml_outcomes.empty:
            fig2 = px.bar(ml_outcomes, x="table_name", y="value", title="TX_ML by outcome")
            fig2.update_layout(yaxis_title="Clients", xaxis_title="Outcome")
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown("**Viral Load Cascade by Facility**")
    vl_elig = filtered[(filtered["sheet"] == "TX_PVLS") & (filtered["table_name"] == "Eligible")]
    vl_test = filtered[(filtered["sheet"] == "TX_PVLS") & (filtered["table_name"] == "Tested")]
    vl_supp = filtered[(filtered["sheet"] == "TX_PVLS") & (filtered["table_name"] == "Suppressed Viral Load")]
    if not vl_elig.empty:
        elig_by_fac = vl_elig.groupby("facility")["value"].sum().rename("Eligible")
        test_by_fac = vl_test.groupby("facility")["value"].sum().rename("Tested")
        supp_by_fac = vl_supp.groupby("facility")["value"].sum().rename("Suppressed")
        vl_pivot = pd.concat([elig_by_fac, test_by_fac, supp_by_fac], axis=1).fillna(0)
        vl_pivot["% VL Coverage"] = (
            vl_pivot["Tested"] / vl_pivot["Eligible"] * 100
        ).where(vl_pivot["Eligible"] > 0, 0).round(1)
        vl_pivot["% VL Suppression"] = (
            vl_pivot["Suppressed"] / vl_pivot["Tested"] * 100
        ).where(vl_pivot["Tested"] > 0, 0).round(1)
        vl_pivot.index.name = "Facility"

        vl_totals = vl_pivot[["Eligible", "Tested", "Suppressed"]].sum()
        vl_totals["% VL Coverage"] = (
            round(vl_totals["Tested"] / vl_totals["Eligible"] * 100, 1) if vl_totals["Eligible"] else 0
        )
        vl_totals["% VL Suppression"] = (
            round(vl_totals["Suppressed"] / vl_totals["Tested"] * 100, 1) if vl_totals["Tested"] else 0
        )
        vl_totals.name = "All facilities (total)"
        vl_pivot = pd.concat([vl_pivot, vl_totals.to_frame().T])

        styled_vl = styled_pct_table(
            vl_pivot, ["% VL Coverage", "% VL Suppression"], ["Eligible", "Tested", "Suppressed"]
        )
        st.dataframe(styled_vl, use_container_width=True)
        st.caption(
            "Green \u2265 90% \u00b7 Yellow 84\u201390% \u00b7 Red < 84%. "
            "% VL Coverage = Tested \u00f7 Eligible. % VL Suppression = Suppressed \u00f7 Tested."
        )

    pyramid = (
        filtered[(filtered["sheet"] == "TX_CURR") & (filtered["table_name"] == "By age and sex")]
        .groupby(["row_label", "col_label"])["value"].sum().reset_index()
    )
    if not pyramid.empty:
        pivot = pyramid.pivot(index="row_label", columns="col_label", values="value")
        pivot = pivot.reindex(AGE_BANDS_15).fillna(0)
        fig4 = go.Figure()
        fig4.add_bar(y=pivot.index, x=pivot.get("Female", 0), name="Female", orientation="h")
        fig4.add_bar(y=pivot.index, x=-pivot.get("Male", 0), name="Male", orientation="h")
        fig4.update_layout(barmode="relative", title="TX_CURR \u2014 Age/Sex Pyramid",
                            xaxis_title="Clients")
        st.plotly_chart(fig4, use_container_width=True)

    st.divider()
    st.markdown("### MMD, DTG & DSDM")

    c5, c6 = st.columns(2)

    with c5:
        mmd_df = filtered[(filtered["sheet"] == "TX_CURR") &
                           (filtered["table_name"] == "ARV dispensing quantity")].copy()
        if not mmd_df.empty:
            def mmd_duration(label):
                if "<3 months" in label:
                    return "<3 months"
                if "3-5 months" in label:
                    return "3-5 months"
                return "6+ months"

            mmd_df["duration"] = mmd_df["row_label"].apply(mmd_duration)
            mmd_df["age_group"] = mmd_df["row_label"].apply(
                lambda r: "<15yrs" if "<15yrs" in r else "15+yrs"
            )
            mmd_grouped = mmd_df.groupby(["duration", "age_group"])["value"].sum().reset_index()
            fig5 = px.bar(
                mmd_grouped, x="duration", y="value", color="age_group", barmode="group",
                category_orders={"duration": ["<3 months", "3-5 months", "6+ months"]},
                title="MMD \u2014 ARV Dispensing Duration",
            )
            fig5.update_layout(yaxis_title="Clients", xaxis_title="Dispensing duration")
            st.plotly_chart(fig5, use_container_width=True)

            mmd_6plus = mmd_grouped.loc[mmd_grouped["duration"] == "6+ months", "value"].sum()
            pct_6plus = (mmd_6plus / txcurr_total * 100) if txcurr_total else 0
            st.caption(f"**{pct_6plus:.0f}%** of TX_CURR clients are on 6+ months MMD.")

    with c6:
        dtg_df = filtered[(filtered["sheet"] == "TX_CURR") &
                           (filtered["table_name"] == "DTG regimen")].copy()
        txcurr_age_df = filtered[(filtered["sheet"] == "TX_CURR") &
                                  (filtered["table_name"] == "By age and sex")].copy()
        if not dtg_df.empty and not txcurr_age_df.empty:
            txcurr_age_df["age_group"] = txcurr_age_df["row_label"].apply(
                lambda r: "<15yrs" if r in AGE_BANDS_15[:4] else "15+yrs"
            )
            txcurr_by_group = txcurr_age_df.groupby("age_group")["value"].sum()
            dtg_by_group = dtg_df.groupby("row_label")["value"].sum()

            compare_rows = []
            for grp in ["<15yrs", "15+yrs"]:
                on_dtg = dtg_by_group.get(grp, 0)
                grp_total = txcurr_by_group.get(grp, 0)
                compare_rows.append({"age_group": grp, "status": "On DTG", "value": on_dtg})
                compare_rows.append({"age_group": grp, "status": "Not on DTG",
                                      "value": max(grp_total - on_dtg, 0)})
            dtg_compare = pd.DataFrame(compare_rows)
            fig6 = px.bar(
                dtg_compare, x="age_group", y="value", color="status", barmode="stack",
                title="DTG Coverage by Age Group",
            )
            fig6.update_layout(yaxis_title="Clients", xaxis_title="Age group")
            st.plotly_chart(fig6, use_container_width=True)

            dtg_total_all = dtg_df["value"].sum()
            pct_dtg = (dtg_total_all / txcurr_total * 100) if txcurr_total else 0
            st.caption(f"**{pct_dtg:.0f}%** of TX_CURR clients are on a DTG-based regimen.")

    c6b1, c6b2 = st.columns(2)

    txcurr_age_df2 = filtered[(filtered["sheet"] == "TX_CURR") &
                               (filtered["table_name"] == "By age and sex")].copy()
    if not txcurr_age_df2.empty:
        txcurr_age_df2["age_group"] = txcurr_age_df2["row_label"].apply(
            lambda r: "<15yrs" if r in AGE_BANDS_15[:4] else "15+yrs"
        )
        txcurr_by_grp_sex = (
            txcurr_age_df2.groupby(["age_group", "col_label"])["value"].sum()
            .reset_index().rename(columns={"col_label": "sex"})
        )
        txcurr_by_grp_sex["source"] = "TX_CURR"

    with c6b1:
        mmd_df2 = filtered[(filtered["sheet"] == "TX_CURR") &
                            (filtered["table_name"] == "ARV dispensing quantity")].copy()
        if not mmd_df2.empty and not txcurr_age_df2.empty:
            mmd_df2["age_group"] = mmd_df2["row_label"].apply(
                lambda r: "<15yrs" if "<15yrs" in r else "15+yrs"
            )
            mmd_by_grp_sex = (
                mmd_df2.groupby(["age_group", "col_label"])["value"].sum()
                .reset_index().rename(columns={"col_label": "sex"})
            )
            mmd_by_grp_sex["source"] = "MMD (dispensed)"
            compare_mmd = pd.concat([
                txcurr_by_grp_sex[["age_group", "sex", "value", "source"]],
                mmd_by_grp_sex[["age_group", "sex", "value", "source"]],
            ])
            compare_mmd["group"] = compare_mmd["age_group"] + " \u2014 " + compare_mmd["sex"]
            fig6b = px.bar(
                compare_mmd, x="group", y="value", color="source", barmode="group",
                title="MMD vs TX_CURR, by Age Group and Sex",
            )
            fig6b.update_layout(yaxis_title="Clients", xaxis_title="Age group / sex")
            st.plotly_chart(fig6b, use_container_width=True)

    with c6b2:
        dtg_df2 = filtered[(filtered["sheet"] == "TX_CURR") &
                            (filtered["table_name"] == "DTG regimen")].copy()
        if not dtg_df2.empty and not txcurr_age_df2.empty:
            dtg_by_grp_sex = (
                dtg_df2.groupby(["row_label", "col_label"])["value"].sum()
                .reset_index().rename(columns={"row_label": "age_group", "col_label": "sex"})
            )
            dtg_by_grp_sex["source"] = "DTG"
            compare_dtg = pd.concat([
                txcurr_by_grp_sex[["age_group", "sex", "value", "source"]],
                dtg_by_grp_sex[["age_group", "sex", "value", "source"]],
            ])
            compare_dtg["group"] = compare_dtg["age_group"] + " \u2014 " + compare_dtg["sex"]
            fig6c = px.bar(
                compare_dtg, x="group", y="value", color="source", barmode="group",
                title="DTG vs TX_CURR, by Age Group and Sex",
            )
            fig6c.update_layout(yaxis_title="Clients", xaxis_title="Age group / sex")
            st.plotly_chart(fig6c, use_container_width=True)

    st.markdown("**MMD Coverage by Facility**")
    txcurr_totals = (
        filtered[(filtered["sheet"] == "TX_CURR") & (filtered["table_name"] == "By age and sex")]
        .groupby("facility")["value"].sum().rename("TX_CURR")
    )
    mmd_pivot_df = filtered[(filtered["sheet"] == "TX_CURR") &
                             (filtered["table_name"] == "ARV dispensing quantity")].copy()
    if not txcurr_totals.empty and not mmd_pivot_df.empty:
        mmd_pivot_df["duration"] = mmd_pivot_df["row_label"].apply(
            lambda label: "<3 months" if "<3 months" in label
            else ("3-5 months" if "3-5 months" in label else "6+ months")
        )
        mmd_3plus = (
            mmd_pivot_df[mmd_pivot_df["duration"].isin(["3-5 months", "6+ months"])]
            .groupby("facility")["value"].sum().rename("MMD 3+ months")
        )
        mmd_6plus = (
            mmd_pivot_df[mmd_pivot_df["duration"] == "6+ months"]
            .groupby("facility")["value"].sum().rename("MMD 6+ months")
        )
        pivot_table = pd.concat([txcurr_totals, mmd_3plus, mmd_6plus], axis=1).fillna(0)
        pivot_table["% MMD 3+ months"] = (
            pivot_table["MMD 3+ months"] / pivot_table["TX_CURR"] * 100
        ).where(pivot_table["TX_CURR"] > 0, 0).round(1)
        pivot_table["% MMD 6+ months"] = (
            pivot_table["MMD 6+ months"] / pivot_table["TX_CURR"] * 100
        ).where(pivot_table["TX_CURR"] > 0, 0).round(1)
        pivot_table.index.name = "Facility"

        totals_row = pivot_table[["TX_CURR", "MMD 3+ months", "MMD 6+ months"]].sum()
        totals_row["% MMD 3+ months"] = (
            round(totals_row["MMD 3+ months"] / totals_row["TX_CURR"] * 100, 1)
            if totals_row["TX_CURR"] else 0
        )
        totals_row["% MMD 6+ months"] = (
            round(totals_row["MMD 6+ months"] / totals_row["TX_CURR"] * 100, 1)
            if totals_row["TX_CURR"] else 0
        )
        totals_row.name = "All facilities (total)"
        pivot_table = pd.concat([pivot_table, totals_row.to_frame().T])

        styled = styled_pct_table(
            pivot_table, ["% MMD 3+ months", "% MMD 6+ months"],
            ["TX_CURR", "MMD 3+ months", "MMD 6+ months"],
        )
        st.dataframe(styled, use_container_width=True)
        st.caption(
            "Green \u2265 90% \u00b7 Yellow 84\u201390% \u00b7 Red < 84%. "
            "% MMD 3+ months = (3-5 months + 6+ months dispensed) \u00f7 TX_CURR. "
            "% MMD 6+ months = 6+ months dispensed \u00f7 TX_CURR."
        )

    st.markdown("**MMD Coverage by Facility and Sex**")
    if not txcurr_totals.empty and not mmd_pivot_df.empty:
        txcurr_sex_df = filtered[(filtered["sheet"] == "TX_CURR") &
                                  (filtered["table_name"] == "By age and sex")]
        txcurr_totals_sex = (
            txcurr_sex_df.groupby(["facility", "col_label"])["value"].sum().rename("TX_CURR")
        )
        mmd_3plus_sex = (
            mmd_pivot_df[mmd_pivot_df["duration"].isin(["3-5 months", "6+ months"])]
            .groupby(["facility", "col_label"])["value"].sum().rename("MMD 3+ months")
        )
        mmd_6plus_sex = (
            mmd_pivot_df[mmd_pivot_df["duration"] == "6+ months"]
            .groupby(["facility", "col_label"])["value"].sum().rename("MMD 6+ months")
        )
        pivot_sex = pd.concat([txcurr_totals_sex, mmd_3plus_sex, mmd_6plus_sex], axis=1).fillna(0)
        pivot_sex["% MMD 3+ months"] = (
            pivot_sex["MMD 3+ months"] / pivot_sex["TX_CURR"] * 100
        ).where(pivot_sex["TX_CURR"] > 0, 0).round(1)
        pivot_sex["% MMD 6+ months"] = (
            pivot_sex["MMD 6+ months"] / pivot_sex["TX_CURR"] * 100
        ).where(pivot_sex["TX_CURR"] > 0, 0).round(1)
        pivot_sex.index.names = ["Facility", "Sex"]

        totals_by_sex = pivot_sex.groupby(level="Sex")[["TX_CURR", "MMD 3+ months", "MMD 6+ months"]].sum()
        totals_by_sex["% MMD 3+ months"] = (
            totals_by_sex["MMD 3+ months"] / totals_by_sex["TX_CURR"] * 100
        ).where(totals_by_sex["TX_CURR"] > 0, 0).round(1)
        totals_by_sex["% MMD 6+ months"] = (
            totals_by_sex["MMD 6+ months"] / totals_by_sex["TX_CURR"] * 100
        ).where(totals_by_sex["TX_CURR"] > 0, 0).round(1)
        totals_by_sex.index = pd.MultiIndex.from_tuples(
            [("All facilities (total)", sex) for sex in totals_by_sex.index],
            names=["Facility", "Sex"],
        )
        pivot_sex = pd.concat([pivot_sex, totals_by_sex])

        styled_sex = styled_pct_table(
            pivot_sex, ["% MMD 3+ months", "% MMD 6+ months"],
            ["TX_CURR", "MMD 3+ months", "MMD 6+ months"],
        )
        st.dataframe(styled_sex, use_container_width=True)
        st.caption("Green \u2265 90% \u00b7 Yellow 84\u201390% \u00b7 Red < 84%, same formulas as above, split by sex.")

    c7, c8 = st.columns(2)

    with c7:
        dsdm_active = filtered[(filtered["sheet"] == "DSDM_VLC-VLS") &
                                (filtered["table_name"] == "Active on DSD")]
        if not dsdm_active.empty:
            active_by_model = (
                dsdm_active.groupby("col_label")["value"].sum()
                .reset_index().rename(columns={"col_label": "Model", "value": "Active clients"})
                .sort_values("Active clients")
            )
            fig7 = px.bar(
                active_by_model, x="Active clients", y="Model", orientation="h",
                title="Active on DSD by Service Delivery Model",
            )
            st.plotly_chart(fig7, use_container_width=True)

    with c8:
        d_elig = filtered[(filtered["sheet"] == "DSDM_VLC-VLS") & (filtered["table_name"] == "Eligible")]
        d_test = filtered[(filtered["sheet"] == "DSDM_VLC-VLS") & (filtered["table_name"] == "Tested")]
        d_supp = filtered[(filtered["sheet"] == "DSDM_VLC-VLS") &
                           (filtered["table_name"] == "Suppressed Viral Load")]
        if not d_elig.empty:
            elig_g = d_elig.groupby("col_label")["value"].sum()
            test_g = d_test.groupby("col_label")["value"].sum().reindex(elig_g.index).fillna(0)
            supp_g = d_supp.groupby("col_label")["value"].sum().reindex(elig_g.index).fillna(0)
            vl_by_model = pd.DataFrame({
                "Model": elig_g.index, "Tested": test_g.values, "Suppressed": supp_g.values,
            })
            vl_by_model["Suppression %"] = vl_by_model.apply(
                lambda row: (row["Suppressed"] / row["Tested"] * 100) if row["Tested"] else 0, axis=1
            )
            fig8 = px.bar(
                vl_by_model.sort_values("Suppression %"), x="Suppression %", y="Model",
                orientation="h", title="VL Suppression % by DSD Model",
            )
            fig8.update_layout(xaxis_range=[0, 100])
            st.plotly_chart(fig8, use_container_width=True)

    st.divider()
    st.markdown("### CIRA \u2014 Cycle of Interruption and Return to ART")
    st.caption("Compares clients who interrupted treatment (IIT, from TX_ML) against clients who returned to treatment (TX_RTT).")

    iit_by_facility = (
        filtered[(filtered["sheet"] == "TX_ML") & (filtered["table_name"].isin(IIT_OUTCOMES))]
        .groupby("facility")["value"].sum().reset_index().rename(columns={"value": "IIT (TX_ML)"})
    )
    rtt_by_facility = (
        filtered[(filtered["sheet"] == "TX_RTT") & (filtered["table_name"] == "By age and sex")]
        .groupby("facility")["value"].sum().reset_index().rename(columns={"value": "Returned (TX_RTT)"})
    )
    cira_by_facility = pd.merge(iit_by_facility, rtt_by_facility, on="facility", how="outer").fillna(0)

    iit_total = iit_by_facility["IIT (TX_ML)"].sum() if not iit_by_facility.empty else 0
    rtt_total = rtt_by_facility["Returned (TX_RTT)"].sum() if not rtt_by_facility.empty else 0
    return_rate = (rtt_total / iit_total * 100) if iit_total else 0

    kc1, kc2, kc3 = st.columns(3)
    kc1.metric("Total IIT (TX_ML)", f"{int(iit_total):,}")
    kc2.metric("Total Returned (TX_RTT)", f"{int(rtt_total):,}")
    kc3.metric("Return Rate", f"{return_rate:.0f}%")

    if not cira_by_facility.empty:
        fig9 = iit_vs_returned_chart(
            cira_by_facility["facility"], cira_by_facility["IIT (TX_ML)"],
            cira_by_facility["Returned (TX_RTT)"],
            "IIT vs Returned to Treatment, by Facility", "Facility",
        )
        st.plotly_chart(fig9, use_container_width=True)

    cc1, cc2 = st.columns(2)

    with cc1:
        iit_by_sex = (
            filtered[(filtered["sheet"] == "TX_ML") & (filtered["table_name"].isin(IIT_OUTCOMES))]
            .groupby("col_label")["value"].sum().reset_index()
            .rename(columns={"col_label": "sex", "value": "IIT (TX_ML)"})
        )
        rtt_by_sex = (
            filtered[(filtered["sheet"] == "TX_RTT") & (filtered["table_name"] == "By age and sex")]
            .groupby("col_label")["value"].sum().reset_index()
            .rename(columns={"col_label": "sex", "value": "Returned (TX_RTT)"})
        )
        cira_by_sex = pd.merge(iit_by_sex, rtt_by_sex, on="sex", how="outer").fillna(0)
        if not cira_by_sex.empty:
            fig10 = iit_vs_returned_chart(
                cira_by_sex["sex"], cira_by_sex["IIT (TX_ML)"], cira_by_sex["Returned (TX_RTT)"],
                "IIT vs Returned to Treatment, by Sex", "Sex",
            )
            st.plotly_chart(fig10, use_container_width=True)

    with cc2:
        iit_by_age = (
            filtered[(filtered["sheet"] == "TX_ML") & (filtered["table_name"].isin(IIT_OUTCOMES))]
            .groupby("row_label")["value"].sum().reindex(AGE_BANDS_15).fillna(0).reset_index()
            .rename(columns={"row_label": "age_band", "value": "IIT (TX_ML)"})
        )
        rtt_by_age = (
            filtered[(filtered["sheet"] == "TX_RTT") & (filtered["table_name"] == "By age and sex")]
            .groupby("row_label")["value"].sum().reindex(AGE_BANDS_15).fillna(0).reset_index()
            .rename(columns={"row_label": "age_band", "value": "Returned (TX_RTT)"})
        )
        cira_by_age = pd.merge(iit_by_age, rtt_by_age, on="age_band", how="outer").fillna(0)
        if not cira_by_age.empty:
            fig11 = iit_vs_returned_chart(
                cira_by_age["age_band"], cira_by_age["IIT (TX_ML)"], cira_by_age["Returned (TX_RTT)"],
                "IIT vs Returned to Treatment, by Age Band", "Age band",
            )
            fig11.update_xaxes(categoryorder="array", categoryarray=AGE_BANDS_15)
            st.plotly_chart(fig11, use_container_width=True)

    st.caption(
        "Filters above apply to every chart and KPI on this page. Data refreshes "
        "from the Google Sheet about once a minute."
    )
    st.stop()

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
            col_dl, col_edit = st.columns([1, 1])
            with col_dl:
                st.download_button(
                    "Download this submission as CSV",
                    detail.to_csv(index=False).encode("utf-8"),
                    file_name=f"submission_{chosen[:8]}.csv",
                    mime="text/csv",
                )
            with col_edit:
                if st.button("\u270f\ufe0f Edit this submission", type="primary"):
                    meta_row = subs.loc[subs.submission_id == chosen].iloc[0]

                    st.session_state.editing_submission_id = chosen
                    st.session_state.edit_source_tables = load_submission_tables(chosen)

                    prefill = {"nav_radio": "Data entry", "entered_by_text": meta_row["entered_by"]}

                    # Pre-fill facility, only if it still exists in the current facility list
                    if not facilities_df.empty and meta_row["facility"] in facilities_df["facility_name"].values:
                        prefill["facility_select"] = meta_row["facility"]
                    elif facilities_df.empty:
                        prefill["facility_text"] = meta_row["facility"]
                        prefill["org_unit_text"] = meta_row["org_unit"]

                    # Pre-fill period, only if it's still within the dropdown's date range
                    stored_period = meta_row["period"]
                    if stored_period.startswith("Q") and "(" in stored_period:
                        prefill["period_type_radio"] = "Quarter"
                        if stored_period in quarter_options():
                            prefill["period_quarter_select"] = stored_period
                    else:
                        prefill["period_type_radio"] = "Month"
                        if stored_period in month_options():
                            prefill["period_month_select"] = stored_period

                    st.session_state.pending_prefill = prefill
                    st.rerun()
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
        df = tracked_entry_grid(
            tables, "TX_ML", outcome, f"txml_{outcome}", AGE_BANDS_15, ["Female", "Male"], "Age band"
        )
        outcome_tables[outcome] = df

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
    tracked_entry_grid(
        tables, "TX_CURR", "By age and sex", "txcurr_age", AGE_BANDS_15, ["Female", "Male"], "Age band"
    )
    st.markdown("**ARV dispensing quantity (by sex)**")
    tracked_entry_grid(
        tables, "TX_CURR", "ARV dispensing quantity", "txcurr_dispense",
        TX_CURR_DISPENSING, ["Female", "Male"], "Dispensing category"
    )
    st.markdown("**On DTG-based regimen (by sex)**")
    tracked_entry_grid(
        tables, "TX_CURR", "DTG regimen", "txcurr_dtg", TX_CURR_DTG, ["Female", "Male"], "Age/weight band"
    )

with tabs[2]:
    st.subheader("TX_NEW \u2014 Newly Enrolled on ART")
    tracked_entry_grid(
        tables, "TX_NEW", "By age and sex", "txnew_age", AGE_BANDS_15, ["Female", "Male"], "Age band"
    )

with tabs[3]:
    st.subheader("TX_PVLS \u2014 Viral Load Coverage & Suppression")
    for cat in PVLS_CATEGORIES:
        st.markdown(f"**{cat}**")
        tracked_entry_grid(
            tables, "TX_PVLS", cat, f"pvls_{cat.replace(' ', '_')}",
            AGE_BANDS_10, ["Female", "Male"], "Age band"
        )

with tabs[4]:
    st.subheader("TX_RTT \u2014 Returned to Treatment")
    st.markdown("**By age and sex**")
    tracked_entry_grid(
        tables, "TX_RTT", "By age and sex", "txrtt_age", AGE_BANDS_15, ["Female", "Male"], "Age band"
    )
    st.markdown("**Duration of treatment interruption before returning (by sex)**")
    tracked_entry_grid(
        tables, "TX_RTT", "Duration before returning", "txrtt_duration",
        TX_RTT_DURATION, ["Female", "Male"], "Duration"
    )

with tabs[5]:
    st.subheader("DSDM VLC/VLS \u2014 Differentiated Service Delivery Models")
    st.markdown("**Active on DSD, by age and model**")
    tracked_entry_grid(
        tables, "DSDM_VLC-VLS", "Active on DSD", "dsdm_active", AGE_BANDS_10, DSDM_MODELS, "Age band"
    )
    for cat in PVLS_CATEGORIES:
        st.markdown(f"**{cat}, by age and model**")
        tracked_entry_grid(
            tables, "DSDM_VLC-VLS", cat, f"dsdm_{cat.replace(' ', '_')}",
            AGE_BANDS_10, DSDM_MODELS, "Age band"
        )

with tabs[6]:
    st.subheader("PREP_BF_PREG \u2014 PrEP for Pregnant & Breastfeeding Women")
    st.caption("Enter counts by age band, split between pregnant and breastfeeding women.")
    for i, cat in enumerate(PREP_CATEGORIES):
        st.markdown(f"**{cat}**")
        safe_key = f"prep_{i}_" + "".join(ch if ch.isalnum() else "_" for ch in cat)
        tracked_entry_grid(
            tables, "PREP_BF_PREG", cat, safe_key, AGE_BANDS_10, ["Pregnant", "Breastfeeding"], "Age band"
        )

st.divider()
is_editing = bool(st.session_state.editing_submission_id)
col1, col2 = st.columns([1, 3])
with col1:
    submit = st.button(
        "Update report" if is_editing else "Submit report",
        type="primary", use_container_width=True,
    )

meta = {
    "facility": facility,
    "org_unit": org_unit,
    "period": period,
    "entered_by": entered_by,
}


def _save_or_update(meta, tables):
    if st.session_state.editing_submission_id:
        sub_id = update_submission(st.session_state.editing_submission_id, meta, tables)
        st.session_state.editing_submission_id = None
        st.session_state.edit_source_tables = None
        return sub_id, "updated"
    else:
        sub_id = save_submission(meta, tables)
        return sub_id, "submitted"


if submit:
    errors, warnings = run_quality_checks(tables, meta)
    if errors:
        st.error("Please fix the following before submitting:")
        for e in errors:
            st.markdown(f"- \u274c {e}")
        st.session_state.pending_submission = None
    elif warnings:
        st.session_state.pending_submission = {"meta": meta, "tables": tables}
        st.warning("Some things look worth double-checking:")
        for w in warnings:
            st.markdown(f"- \u26a0\ufe0f {w}")
    else:
        sub_id, verb = _save_or_update(meta, tables)
        st.success(f"Report {verb} and saved (ID: {sub_id[:8]}).")
        st.session_state.pending_submission = None

if st.session_state.get("pending_submission"):
    if st.button("Submit anyway, I've reviewed the warnings above", use_container_width=True):
        pending = st.session_state.pending_submission
        sub_id, verb = _save_or_update(pending["meta"], pending["tables"])
        st.success(f"Report {verb} and saved (ID: {sub_id[:8]}).")
        st.session_state.pending_submission = None