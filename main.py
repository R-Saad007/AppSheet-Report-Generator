"""
FastAPI webhook receiver for AppSheet → PDF report generation.

Flow:
  1. AppSheet Bot POSTs a JSON payload when a new survey is submitted.
  2. We validate the X-API-Key header.
  3. We enqueue the job as a background task (responds 202 immediately so
     AppSheet doesn't time out waiting).
  4. Background task:
       a. Fetch Surveys + HVAC_Units CSVs from Google Sheets (fast export URL).
       b. Fetch site images from Google Drive into a temp directory.
       c. Run hvac_report.build_report() to generate the PDF.
       d. Upload the PDF to the site's Drive folder (same folder as images).

Required environment variables (see .env.example):
  API_KEY                   - shared secret sent by AppSheet in X-API-Key header
  SHEET_ID                  - Google Sheets spreadsheet ID
  SURVEYS_GID               - sheet tab GID for Surveys
  HVAC_UNITS_GID            - sheet tab GID for HVAC_Units
  DRIVE_ROOT_FOLDER_ID      - Drive folder ID for "Bell Preventive Maintenance"
  SERVICE_ACCOUNT_JSON      - full JSON of the service account key (or use FILE)
  SERVICE_ACCOUNT_FILE      - path to service account JSON file
"""

import logging
import os
import tempfile
from pathlib import Path

from fastapi import BackgroundTasks, Depends, FastAPI, Header, HTTPException, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel

import drive_service as ds
import hvac_report as hr

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
log = logging.getLogger(__name__)

app = FastAPI(title="Bell PM Report Generator", version="1.0.0")


# ── Security ──────────────────────────────────────────────────────────────────

def _require_api_key(x_api_key: str = Header(..., alias="X-API-Key")):
    expected = os.environ.get("API_KEY", "")
    if not expected:
        raise RuntimeError("API_KEY env var is not set")
    if x_api_key != expected:
        raise HTTPException(status_code=401, detail="Invalid API key")


# ── Payload model ─────────────────────────────────────────────────────────────

class SurveyPayload(BaseModel):
    survey_id: str
    site_name: str | None = None
    site_folder_name: str | None = None  # exact Drive folder name, e.g. "G1053_BELL_TS_Fenelon Falls_2026-04-22"


# ── Background job ────────────────────────────────────────────────────────────

def _generate_report(survey_id: str, site_folder_name: str | None):
    sheet_id = os.environ["SHEET_ID"]
    surveys_gid = os.environ["SURVEYS_GID"]
    units_gid = os.environ["HVAC_UNITS_GID"]
    root_folder_id = os.environ["DRIVE_ROOT_FOLDER_ID"]

    log.info("Starting report generation for survey_id=%s", survey_id)

    with tempfile.TemporaryDirectory(prefix="bell_pm_") as tmp:
        tmp_dir = Path(tmp)

        # 1. Fetch CSVs
        log.info("Fetching Surveys CSV from Sheets (gid=%s)", surveys_gid)
        surveys_csv = ds.fetch_sheet_csv(sheet_id, surveys_gid)
        surveys_path = ds.write_csv_to_temp(surveys_csv, "Surveys.csv", tmp_dir)

        log.info("Fetching HVAC_Units CSV from Sheets (gid=%s)", units_gid)
        units_csv = ds.fetch_sheet_csv(sheet_id, units_gid)
        units_path = ds.write_csv_to_temp(units_csv, "HVAC_Units.csv", tmp_dir)

        # 2. Load rows and find the target survey
        rows = hr.load_rows(surveys_path)
        units = hr.load_hvac_units(units_path)

        matching = [r for r in rows if r.get("Survey ID", "").strip() == survey_id.strip()]
        if not matching:
            log.error("Survey ID '%s' not found in Surveys sheet", survey_id)
            return
        data = matching[0]

        # 3. Determine site folder name (from payload or from Location Code + Site Name)
        if not site_folder_name:
            loc = data.get("Location Code", "").strip()
            site = data.get("Site Name", "").strip()
            report_date = hr.iso_date(data.get("Report Date", "")) or ""
            if loc and site and report_date:
                site_folder_name = f"{loc}_{site}_{report_date}".replace(" ", "_")
            elif loc:
                site_folder_name = loc

        log.info("Resolved site_folder_name=%s", site_folder_name)

        # 4. Download images from Drive
        images_dir = tmp_dir / "images"
        if site_folder_name:
            images_folder_id = ds.resolve_site_images_folder_id(root_folder_id, site_folder_name)
            if images_folder_id:
                log.info("Downloading images from Drive folder %s", images_folder_id)
                ds.download_site_images(images_folder_id, images_dir)
            else:
                log.warning("Could not locate images folder for '%s' in Drive", site_folder_name)
                images_dir.mkdir(parents=True, exist_ok=True)
        else:
            log.warning("No site_folder_name resolved; skipping image download")
            images_dir.mkdir(parents=True, exist_ok=True)

        # 5. Build the PDF
        my_units = hr.units_for_survey(units, survey_id)
        site_folder_local = tmp_dir  # images are in tmp_dir/images, PDF goes here
        out_path = tmp_dir / hr.output_filename(data)

        log.info("Generating PDF: %s (%d units)", out_path.name, len(my_units))
        hr.build_report(data, tmp_dir, site_folder_local, str(out_path), units=my_units)
        log.info("PDF generated: %s (%.0f KB)", out_path.name, out_path.stat().st_size / 1024)

        # 6. Upload PDF to the site's Drive folder
        if site_folder_name:
            site_drive_id = ds.find_site_folder_id(root_folder_id, site_folder_name)
            if not site_drive_id:
                log.warning("Site folder '%s' not found in Drive — uploading to root", site_folder_name)
                site_drive_id = root_folder_id
        else:
            site_drive_id = root_folder_id

        file_id = ds.upload_pdf(out_path, site_drive_id)
        log.info("PDF uploaded to Drive (file_id=%s)", file_id)

    log.info("Report generation complete for survey_id=%s", survey_id)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/webhook/survey", status_code=202, dependencies=[Depends(_require_api_key)])
async def survey_webhook(payload: SurveyPayload, background_tasks: BackgroundTasks):
    """Called by AppSheet when a new survey is submitted."""
    log.info("Received webhook: survey_id=%s site_folder=%s",
             payload.survey_id, payload.site_folder_name)
    background_tasks.add_task(
        _generate_report,
        survey_id=payload.survey_id,
        site_folder_name=payload.site_folder_name,
    )
    return {"message": "Report generation queued", "survey_id": payload.survey_id}
