# Canada's Leading Telecom Service Provider Site PM Report Generator

A FastAPI webhook service that automatically generates branded PDF preventive-maintenance reports when a technician submits a survey in AppSheet. The service runs on **Google Cloud Run**, authenticates via a native attached service account (no JSON key files in production), pulls live data from **Google Sheets**, downloads site photos from **Google Drive**, renders a multi-section PDF report with [ReportLab](https://www.reportlab.com/), and uploads the finished file back to the site's Drive folder.

---

## Architecture

```
AppSheet Bot
     │  POST /webhook/survey
     │  X-API-Key: <secret>
     ▼
┌─────────────────────────────────────────────────────────┐
│                    Cloud Run (main.py)                  │
│                                                         │
│  1. Validate X-API-Key header                           │
│  2. Return HTTP 202 immediately (no timeout risk)       │
│  3. Background task:                                    │
│     a. Fetch Surveys CSV  ──► Google Sheets API         │
│     b. Fetch HVAC_Units CSV ► Google Sheets API         │
│     c. Download site images ► Google Drive API          │
│     d. Build PDF ───────────► hvac_report.py            │
│     e. Upload PDF ──────────► Google Drive API          │
└─────────────────────────────────────────────────────────┘
```

Authentication on Cloud Run uses **Application Default Credentials** (ADC) — the service account is attached directly to the Cloud Run instance, so no credential files or env vars are needed in production.

---

## Repository Structure

```
.
├── main.py                  # FastAPI app — webhook receiver & background job orchestrator
├── hvac_report.py           # Core PDF engine: layouts, sections, photo grids, ReportLab logic
├── sitesurvey_report.py     # Standalone CLI: generic site-survey PDF from any AppSheet CSV export
├── drive_service.py         # Google Drive & Sheets helpers (auth, CSV fetch, image download, PDF upload)
├── Dockerfile               # python:3.12-slim image; Cloud Run injects $PORT
├── requirements.txt         # Pinned Python dependencies
├── .env.example             # All required and optional environment variables with comments
└── .gitignore
```

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Python 3.12+ | Matches the Docker base image |
| Google Cloud project | With Cloud Run, Secret Manager, Drive, and Sheets APIs enabled |
| Service account | Needs Drive (`drive`) + Sheets (`spreadsheets.readonly`) OAuth scopes |
| Google Sheets spreadsheet | Two tabs: **Surveys** and **HVAC_Units** |
| Google Drive root folder | `Bell Preventive Maintenance/` — service account must have Editor access |
| AppSheet app | Bot configured to POST to the webhook on survey submission |

---

## Environment Variables

Copy `.env.example` to `.env` for local development. In production these are injected via `--set-env-vars` and `--set-secrets` in the `gcloud run deploy` command.

| Variable | Required | Description |
|---|---|---|
| `API_KEY` | Yes | Shared secret validated against the `X-API-Key` request header. Generate with `python -c "import secrets; print(secrets.token_urlsafe(32))"` |
| `SHEET_ID` | Yes | Google Sheets spreadsheet ID (long string in the sheet URL) |
| `SURVEYS_GID` | Yes | Sheet tab GID for the Surveys table (visible in the URL as `#gid=…`) |
| `HVAC_UNITS_GID` | Yes | Sheet tab GID for the HVAC_Units table |
| `DRIVE_ROOT_FOLDER_ID` | Yes | Drive folder ID for `Bell Preventive Maintenance/` |
| `SERVICE_ACCOUNT_JSON` | Local dev | Full service account key JSON as a single-line string |
| `SERVICE_ACCOUNT_FILE` | Local dev | Path to a service account key JSON file |

`SERVICE_ACCOUNT_JSON` / `SERVICE_ACCOUNT_FILE` are **not needed on Cloud Run** — ADC is used automatically when a service account is attached to the instance.

---

## Local Development

```bash
# 1. Clone and create a virtual environment
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Configure credentials
cp .env.example .env
# Fill in SHEET_ID, SURVEYS_GID, HVAC_UNITS_GID, DRIVE_ROOT_FOLDER_ID,
# and one of SERVICE_ACCOUNT_JSON or SERVICE_ACCOUNT_FILE

# 4. Start the server
uvicorn main:app --reload --port 8080
```

The interactive API docs are available at `http://localhost:8080/docs`.

### Running a Report Locally (without the webhook)

Use `hvac_report.py` directly against local CSV exports and a downloaded images folder:

```bash
python hvac_report.py \
  --surveys  "Bell Site Survey AppSheet Headers - Surveys.csv" \
  --units    "Bell Site Survey AppSheet Headers - HVAC_Units.csv" \
  --photos   "Bell Preventive Maintenance/" \
  --id       <SURVEY_ID> \
  --output   "Site Survey Reports/PM_Report_<LOC>_<DATE>.pdf"
```

Or via the Python API (as used in the scripts above):

```python
import hvac_report as hr
from pathlib import Path

rows     = hr.load_rows("Bell Site Survey AppSheet Headers - Surveys.csv")
units    = hr.load_hvac_units("Bell Site Survey AppSheet Headers - HVAC_Units.csv")
data     = next(r for r in rows if r["Survey ID"] == "<SURVEY_ID>")
my_units = hr.units_for_survey(units, "<SURVEY_ID>")
out_path = Path("Site Survey Reports") / hr.output_filename(data)

hr.build_report(
    data,
    photo_root=".",                                      # root that Drive paths are resolved against
    site_folder="Bell Preventive Maintenance/<SITE_FOLDER_NAME>",
    out_path=str(out_path),
    units=my_units,
)
```

### Generic Site Survey Reports (`sitesurvey_report.py`)

A separate, self-contained CLI for generating a styled PDF from any AppSheet CSV export (not HVAC-specific):

```bash
python sitesurvey_report.py \
  --csv    surveys.csv \
  --photos ./photos \
  --output report.pdf \
  --title  "Q2 Site Surveys"
```

---

## Drive Folder Convention

The service expects site images to live under the Bell Preventive Maintenance root folder using this naming pattern:

```
Bell Preventive Maintenance/
└── <LOC_CODE>_<CUSTOMER>_TS_<SITE_NAME>_<YYYY-MM-DD>/
    └── images/          ← site photos downloaded from here
        ├── <id>.<Column Name>.<HHMMSS>.jpg
        └── ...
```

The `site_folder_name` in the webhook payload must match this folder name exactly. If omitted, it is reconstructed automatically from the survey row as `{Location Code}_{Site Name}_{Report Date}` (spaces replaced with underscores).

The generated PDF is uploaded to the site folder itself (one level above `images/`):

```
Bell Preventive Maintenance/
└── G1032_BELL_TS_Hanover Rogers (Rogers Colocate)_2026-04-29/
    ├── images/
    │   └── *.jpg
    └── PM_Report_G1032_2026-04-29.pdf   ← uploaded here
```

---

## Webhook API

### `GET /health`

Liveness check. Returns `{"status": "ok"}` with HTTP 200.

### `POST /webhook/survey`

Triggers report generation for a submitted survey.

**Headers**

| Header | Value |
|---|---|
| `X-API-Key` | The shared secret matching the `API_KEY` env var |
| `Content-Type` | `application/json` |

**Request body**

```json
{
  "survey_id": "<<[Survey ID]>>",
  "site_name": "<<[Site Name]>>",
  "site_folder_name": "<<[Location Code]>>_<<[Site Name]>>_<<[Report Date]>>"
}
```

`site_name` and `site_folder_name` are optional — they are derived from the Sheets data if omitted.

**Response**

```json
HTTP 202 Accepted
{
  "message": "Report generation queued",
  "survey_id": "<survey_id>"
}
```

The response is returned immediately. Report generation runs asynchronously in the background.

**Error responses**

| Code | Cause |
|---|---|
| `401 Unauthorized` | Missing or invalid `X-API-Key` header |
| `422 Unprocessable Entity` | Malformed JSON body (missing `survey_id`) |

---

## Deployment — Google Cloud Run

### One-time setup

```bash
# Enable required APIs
gcloud services enable run.googleapis.com secretmanager.googleapis.com \
  drive.googleapis.com sheets.googleapis.com --project=<PROJECT_ID>

# Store the API key in Secret Manager
echo -n "<your_api_key>" | \
  gcloud secrets create webhook-api-key --data-file=- --project=<PROJECT_ID>

# Grant the service account access to the secret
gcloud secrets add-iam-policy-binding webhook-api-key \
  --member="serviceAccount:report-generator@<PROJECT_ID>.iam.gserviceaccount.com" \
  --role="roles/secretmanager.secretAccessor" \
  --project=<PROJECT_ID>
```

### Deploy

```bash
gcloud run deploy report-service \
  --source . \
  --region=northamerica-northeast2 \
  --project=<PROJECT_ID> \
  --service-account="report-generator@<PROJECT_ID>.iam.gserviceaccount.com" \
  --set-secrets="API_KEY=webhook-api-key:latest" \
  --set-env-vars="SHEET_ID=<YOUR_SHEET_ID>" \
  --set-env-vars="SURVEYS_GID=0" \
  --set-env-vars="HVAC_UNITS_GID=<YOUR_HVAC_UNITS_GID>" \
  --set-env-vars="DRIVE_ROOT_FOLDER_ID=<YOUR_DRIVE_FOLDER_ID>" \
  --allow-unauthenticated
```

The `--source .` flag uses **Cloud Build** to build and push the Docker image automatically. The `--service-account` flag attaches the service account to the instance, enabling ADC — no JSON key file is required.

### Verify

```bash
# Get the service URL
gcloud run services describe report-service \
  --region=northamerica-northeast2 --project=<PROJECT_ID> \
  --format="value(status.url)"

# Health check
curl https://<SERVICE_URL>/health
```

---

## Authentication Flow

```
Cloud Run instance
       │
       │  google.auth.default(scopes=[...])
       │        │
       │        ▼
       │  Metadata server (169.254.169.254)
       │  Returns short-lived access token for the
       │  attached service account
       │
       ▼
  Google APIs (Drive, Sheets)
```

For local development, `_credentials()` in `drive_service.py` falls back to `SERVICE_ACCOUNT_JSON` or `SERVICE_ACCOUNT_FILE` when ADC is unavailable.

---

## PDF Report Sections

Each generated PM report contains the following sections, in order:

1. **Cover page** — site name, location code, technician, report date, stat summary cards
2. **Timeline** — departure, arrival, work commenced, WNOC clearance, departure time with calculated durations
3. **Survey Overview** — site coordinates, contact person, travel distance, site overview photos
4. **System Identification** — controller type, serial numbers, firmware versions, gateway details with photos
5. **Controller Checks** — active alarms, alarm history, operating mode, DC voltage readings
6. **Arrival Setpoints** — recorded temperature and operational setpoints at arrival
7. **HVAC Unit sections** *(repeated per unit)*
   - **Before State** — pre-maintenance condition, filter state, cleanliness ratings
   - **Maintenance Performed** — filter replacement, coil cleaning, bug screens, gaskets, screws
   - **Unit Testing** — electrical readings (amps/volts per circuit), pressure readings, refrigerant checks
8. **Shared Testing** — lag failover, DC fan failover, temperature controls, sensor validation
9. **Final Inspection** — grills, manuals, TSSA sticker, alarm confirmation, final notes
10. **End of Report** — total photo count, sign-out photo, technician sign-off

---

## Dependencies

| Package | Purpose |
|---|---|
| `fastapi` | ASGI web framework, webhook routing, dependency injection |
| `uvicorn[standard]` | ASGI server (Uvicorn with extras: websockets, http-tools) |
| `pydantic` | Request body validation (`SurveyPayload` model) |
| `reportlab` | Pure-Python PDF generation (no OS-level binaries required) |
| `pillow` | Image dimension detection for correct photo scaling in the PDF |
| `qrcode[pil]` | QR code generation embedded in report pages |
| `google-auth` | ADC (`google.auth.default`) and service account credential support |
| `google-api-python-client` | Drive v3 and Sheets v4 API clients |
| `google-auth-httplib2` | HTTP transport adapter for the Google API client |
| `requests` | HTTP client used for Sheets fast-export CSV fetch |
