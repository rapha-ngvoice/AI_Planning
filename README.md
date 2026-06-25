# Resource Capacity Planning — H2 Replan

A board-ready Streamlit application that forecasts team working-day capacity for
the second half of the year (July–December), with persistent storage in a
Google Sheet. The data model and the Excel export mirror the **Replan H2** tab
of `Teams.xlsx` verbatim.

---

## What it does

- **Employee profiles** — Name, Team (dropdown, with "add new team"), Team Lead,
  Code/work-stream, fully customisable **Capacity %** (100% → `1`, 50% → `0.5`,
  0% → not contributing), Contract type, and explicit Start / End dates.
- **New-starter ramp-up** — automatically applied from the Start Date:
  Month 1 = 33%, Month 2 = 66%, Month 3 = 100% of designated capacity.
- **Max working days** — baseline **15.5 days / person / month** (adjustable).
- **Holiday & absence tracker** — log dated absences per person/month; they
  deduct dynamically from that month's available days.
- **Global sidebar sliders** — Max Working Days, Sickness Allowance, Holiday
  Allowance, Admin / S&C days — all applied across the whole forecast live.
- **Board-ready UI** — `st.metric` headline KPIs, `st.data_editor` inline
  editing, and a "capacity by team per month" chart.
- **Google Sheets persistence** — reads/writes through
  `st.connection("gsheets", type=GSheetsConnection)`. Adding a member or logging
  an absence writes back immediately, so nothing is lost on restart.
- **Excel export** — one click produces a `.xlsx` that reproduces the Replan H2
  layout: per-team blocks, `=SUM` **TOTAL** rows, `=COUNTIF(range,"<>0")`
  **COUNT** (headcount) rows, and 0-day months rendered as `-`.

> **Capacity formula (per person, per month)**
> First derive the net baseline from the assumptions block:
> `Planable = WorkingDays − Sickness − Holiday − Admin/S&C` (e.g.
> `21 − 1 − 2.5 − 2 = 15.5`). Then per person:
> `Planable × ContractFactor × Capacity% × RampFactor − LoggedAbsences`,
> floored at 0 and rounded to 1 dp.
> ContractFactor: 5-day week (FT) = 1.0 (15.5), 4.5-day week = 0.9 (≈14),
> 4-day week = 0.8 (12.4). A result of `0` means "not contributing", exactly as
> on Replan H2. Example — a 4-day-week person at 50% capacity contributes
> `15.5 × 0.8 × 0.5 = 6.2` days/month. The sidebar shows the derived Planable
> days and the per-contract day totals (mirroring the assumptions block).

---

## 1. Install

```bash
pip install -r requirements.txt
```

## 2. Run

```bash
streamlit run app.py
```

On first launch — **before** you configure Google Sheets — the app runs against
an in-session store seeded with the roster extracted from `Teams.xlsx`, so you
can explore it immediately. The sidebar shows the active backend.

## 3. Saving & reopening your data

The app **saves automatically to a durable local file** (`data/team.csv` and
`data/holidays.csv` next to `app.py`) whenever you add a member, log an absence,
or click **Save roster / Save absences**. When you reopen the app it reloads
those saved values — no Google Sheets required. The sidebar shows where the
local file lives.

> On your own machine or a server with a persistent disk, the local file is all
> you need. On **Streamlit Community Cloud** the local disk is wiped on every
> restart/redeploy, so for durable cloud storage (and sharing across people)
> connect a Google Sheet as below — saves then go to *both* the Sheet and the
> local cache.

To start the roster fresh, delete the `data/` folder (or the CSVs inside it);
the app re-seeds from the bundled `Teams.xlsx` roster on next launch.

## 4. Connect a Google Sheet (optional — shared / cloud storage)

### 3a. Create the Sheet
Create a Google Sheet with **two tabs (worksheets)**, named exactly:

- `Team` with header row: `Name | Team | TL | Code | Capacity | Contract | StartDate | EndDate`
- `Holidays` with header row: `Name | Month | Days | Note`

(You can paste the contents of the bundled `seed_roster.json` into the `Team`
tab to start from the real roster. `Capacity` is stored as a fraction —
`1` = 100%, `0.5` = 50%.)

### 3b. Create a Service Account (for read **and** write)
1. In the [Google Cloud Console](https://console.cloud.google.com/) create or
   select a project.
2. Enable the **Google Sheets API** and the **Google Drive API**.
3. Go to **IAM & Admin → Service Accounts → Create service account**.
4. Create a **JSON key** for that service account and download it.
5. Open your Google Sheet → **Share** → share it with the service account's
   `client_email` (give it **Editor** access).

### 3c. Add credentials to Streamlit secrets
Copy the template and fill in the values from the JSON key file:

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Then edit `.streamlit/secrets.toml`:

```toml
[connections.gsheets]
spreadsheet = "https://docs.google.com/spreadsheets/d/XXXX/edit"
type = "service_account"
project_id = "…"
private_key_id = "…"
private_key = "-----BEGIN PRIVATE KEY-----\n…\n-----END PRIVATE KEY-----\n"
client_email = "…@….iam.gserviceaccount.com"
client_id = "…"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "…"
```

Restart the app. The sidebar should now read **"Connected to Google Sheets ✔"**.

> The code reads/writes via `conn.read(worksheet="Team", ttl=0)` and
> `conn.update(worksheet="Team", data=df)` — `ttl=0` ensures fresh reads after
> every write.

### Security notes
- **Never** commit `.streamlit/secrets.toml` or the downloaded JSON key — add
  both to `.gitignore`. The app never asks you to type credentials into the UI.
- Share the Sheet with *only* the service account, and prefer the
  least-privileged scope your workflow allows.
- For Streamlit Community Cloud / a hosted deployment, paste the same TOML into
  the host's **Secrets** manager rather than shipping a file.

---

## Deploying to Streamlit Community Cloud (optional)
Push the repo (without secrets), then in the app's **Settings → Secrets** paste
the `[connections.gsheets]` block. No code changes needed.

---

## Files
- `app.py` — the application.
- `requirements.txt` — dependencies.
- `.streamlit/secrets.toml.example` — credentials template.
- `seed_roster.json` — roster extracted from `Teams.xlsx` (bootstrap data).
