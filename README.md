# DCT 2026 Data Collection Tool

A Streamlit web app for entering the ART program indicators from `DCTs_2026.xlsx`:
**TX_ML, TX_CURR, TX_NEW, TX_PVLS, TX_RTT, and DSDM VLC/VLS**, with the same
age-band, sex, and category disaggregations as the original workbook.

Every submission is saved to a **Google Sheet**, so it survives redeploys and
your team can also open it directly in Sheets if they want.

---

## 1. One-time Google Cloud setup (5-10 min)

1. Go to the [Google Cloud Console](https://console.cloud.google.com/) and
   create a new project (or use an existing one).
2. In the search bar, enable these two APIs for the project:
   - **Google Sheets API**
   - **Google Drive API**
3. Go to **APIs & Services -> Credentials -> Create Credentials -> Service
   account**. Give it any name (e.g. `dct-app`), skip the optional role/access
   steps, and click **Done**.
4. Click into the new service account -> **Keys** tab -> **Add Key -> Create
   new key -> JSON**. This downloads a `.json` file — keep it safe, it's a
   credential.
5. Open that JSON file and copy the `client_email` value (looks like
   `dct-app@your-project-id.iam.gserviceaccount.com`).

## 2. Create and share the Google Sheet

1. Create a new Google Sheet (any name, e.g. "DCT 2026 Submissions"). You can
   leave it empty — the app creates the `submissions` and `entries` tabs
   automatically the first time it runs.
2. Click **Share**, paste in the service account's `client_email` from step
   1.5, and give it **Editor** access.
3. Copy the sheet's key from its URL:
   `https://docs.google.com/spreadsheets/d/`**`THIS_PART_IS_THE_KEY`**`/edit`

## 3. Configure secrets

Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and fill
it in:
- `sheet_key` = the key you copied in step 2.3
- `[gcp_service_account]` = every field from the JSON file you downloaded in
  step 1.4, pasted in as-is (the example file shows the exact shape expected)

**Never commit `secrets.toml` to GitHub.** Add it to `.gitignore`.

## 4. Run it locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Opens at `http://localhost:8501`. The first submission will auto-create the
`submissions` and `entries` tabs in your Google Sheet.

## 5. Share it with people (host it online)

1. Push this folder to a GitHub repo — **excluding** `secrets.toml` (only
   commit `secrets.toml.example`).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click **New app**, point it at your repo and `app.py`.
3. Before/after deploying, open **Settings -> Secrets** in the Streamlit
   Cloud dashboard and paste in the same content as your local
   `secrets.toml`.
4. Deploy. You get a public URL like `https://your-app-name.streamlit.app`
   that anyone can open and submit data through — no install needed on their
   end, and every submission lands straight in your Google Sheet.

---

## How the app works

- **Sign in**: every user needs an account. A `users` tab is auto-created in
  your Google Sheet on first run, seeded with one admin account:
  - username: `admin`
  - password: `Admin@123`
  Log in with that first, then go to the **Manage users** page (visible only
  to admins) to add real accounts for your team and set a new admin password
  (add a new admin user with a password only you know, then delete the
  default `admin` row from the `users` tab in the Sheet).
- **Manage users** (admin only): add username/password/full name/role for
  each person who'll enter data. Roles are `user` or `admin` — only admins
  see the Manage users page. To remove someone's access, delete their row
  from the `users` tab in the Sheet directly.
- **Sidebar**: facility name (dropdown, sourced from a `facilities` tab in your
  Google Sheet), org unit code (auto-filled), reporting period, entered by.
- **Facility list**: the app auto-creates a `facilities` tab in your Sheet the
  first time it runs, with columns `facility_name` and `org_unit`. Add or edit
  rows there any time to change what shows up in the dropdown — no code
  changes needed. The app caches the list for 5 minutes; use the "Refresh
  facility list" button in the sidebar to pick up changes immediately.
- **Data entry tab**: one sub-tab per indicator sheet (TX_ML, TX_CURR, TX_NEW,
  TX_PVLS, TX_RTT, DSDM VLC/VLS). Each disaggregation is an editable
  spreadsheet-style grid.
- **Submit report**: writes one row to the `submissions` tab and one row per
  data point to the `entries` tab in your Google Sheet.
- **Submission history**: browse past submissions, inspect full detail, and
  download any submission as CSV.

## Files

- `app.py` — the full app
- `requirements.txt` — Python dependencies
- `.streamlit/secrets.toml.example` — template for your Google credentials
  (copy to `secrets.toml` and fill in; never commit the real one)

## Data quality checks

On submit, the app validates the entered data before saving:

**Blocks submission (must be fixed):**
- Facility name and reporting period are required
- TX_PVLS: Suppressed cannot exceed Tested, and Tested cannot exceed Eligible
  (checked per age band and sex)
- DSDM VLC/VLS: same Suppressed \u2264 Tested \u2264 Eligible check, per age band and model

**Warns, but lets you confirm and submit anyway:**
- TX_NEW total greater than TX_CURR total (unusual \u2014 newly enrolled clients
  are normally a subset of everyone currently on ART)
- TX_PVLS Eligible total greater than TX_CURR total
- Any table left entirely at zero (easy to miss a whole section by accident)
- "Entered by" left blank

## Security note

Passwords are stored as salted SHA-256 hashes in the `users` tab (never in
plain text), which is reasonable for an internal team tool. It isn't
enterprise-grade authentication (no lockouts, password resets, or MFA), so
keep the Google Sheet itself restricted to people who should have admin-level
trust, and change the default admin password immediately after first login.

## Troubleshooting

- **"Worksheet not found" / permission errors**: double-check you shared the
  Sheet with the exact `client_email` from your JSON key, with Editor access.
- **`gspread.exceptions.SpreadsheetNotFound`**: check `sheet_key` (or
  `sheet_name`) in your secrets matches the real sheet.
- **Works locally but not on Streamlit Cloud**: make sure you pasted secrets
  into the Cloud app's own Secrets panel — it doesn't read your local file.
