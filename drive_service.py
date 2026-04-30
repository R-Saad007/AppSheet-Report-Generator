"""
Google Drive / Sheets helper for the Cloud Run webhook service.

Responsibilities:
  - Download Surveys and HVAC_Units CSVs from Google Sheets (fast CSV export URL)
  - Download site images from Google Drive into a local temp directory
  - Upload the finished PDF to the correct site folder in Google Drive

Authentication (in priority order):
  1. Application Default Credentials — used automatically on Cloud Run when a
     service account is attached to the instance (no env vars needed).
  2. SERVICE_ACCOUNT_JSON env var — inline service account key JSON string.
  3. SERVICE_ACCOUNT_FILE env var — path to a service account key JSON file.
"""

import io
import json
import logging
import os
import re
import tempfile
from pathlib import Path

import google.auth
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

log = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]


def _credentials():
    # 1. Application Default Credentials (Cloud Run attached service account)
    try:
        credentials, _ = google.auth.default(scopes=SCOPES)
        return credentials
    except Exception as e:
        log.warning("ADC not available, falling back to env vars: %s", e)

    # 2. Inline service account JSON
    raw = os.environ.get("SERVICE_ACCOUNT_JSON")
    if raw:
        info = json.loads(raw)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    # 3. Service account key file
    path = os.environ.get("SERVICE_ACCOUNT_FILE")
    if path:
        return service_account.Credentials.from_service_account_file(path, scopes=SCOPES)

    raise RuntimeError("Could not establish Google Cloud credentials.")


def _drive():
    return build("drive", "v3", credentials=_credentials(), cache_discovery=False)


# ── Sheet → CSV ───────────────────────────────────────────────────────────────

def fetch_sheet_csv(sheet_id: str, gid: str) -> str:
    """Return the raw CSV text for one sheet tab via the fast export URL.

    The service account must have at least Viewer access to the spreadsheet.
    """
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    creds = _credentials()
    # Refresh token so we can use the access token in the Authorization header.
    import google.auth.transport.requests as gtr
    creds.refresh(gtr.Request())
    resp = requests.get(url, headers={"Authorization": f"Bearer {creds.token}"}, timeout=30)
    resp.raise_for_status()
    return resp.text


def write_csv_to_temp(csv_text: str, filename: str, tmp_dir: Path) -> Path:
    p = tmp_dir / filename
    p.write_text(csv_text, encoding="utf-8")
    return p


# ── Drive image download ───────────────────────────────────────────────────────

def _file_id_from_path(drive, folder_id: str, relative_path: str) -> str | None:
    """Walk relative_path segments inside folder_id and return the final file ID."""
    parts = [p for p in re.split(r"[/\\]+", relative_path) if p]
    current_id = folder_id
    for i, part in enumerate(parts):
        is_last = i == len(parts) - 1
        mime = "application/vnd.google-apps.folder" if not is_last else None
        q_parts = [
            f"'{current_id}' in parents",
            f"name = '{part}'",
            "trashed = false",
        ]
        if mime:
            q_parts.append(f"mimeType = '{mime}'")
        result = drive.files().list(
            q=" and ".join(q_parts),
            fields="files(id, name)",
            pageSize=1,
        ).execute()
        files = result.get("files", [])
        if not files:
            return None
        current_id = files[0]["id"]
    return current_id


def download_file_by_id(drive, file_id: str, dest: Path) -> Path:
    request = drive.files().get_media(fileId=file_id)
    with open(dest, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return dest


def download_site_images(
    site_folder_drive_id: str,
    local_dir: Path,
) -> dict[str, Path]:
    """Download every image in the Drive folder (non-recursive) to local_dir.

    Returns a mapping of Drive filename → local Path.
    """
    drive = _drive()
    local_dir.mkdir(parents=True, exist_ok=True)

    result = drive.files().list(
        q=f"'{site_folder_drive_id}' in parents and trashed = false",
        fields="files(id, name, mimeType)",
        pageSize=1000,
    ).execute()

    image_mimes = {
        "image/jpeg", "image/png", "image/gif",
        "image/bmp", "image/webp",
    }
    mapping: dict[str, Path] = {}
    for f in result.get("files", []):
        if f["mimeType"] not in image_mimes:
            continue
        dest = local_dir / f["name"]
        try:
            download_file_by_id(drive, f["id"], dest)
            mapping[f["name"]] = dest
            log.info("Downloaded image: %s", f["name"])
        except Exception as exc:
            log.warning("Failed to download %s: %s", f["name"], exc)
    return mapping


def resolve_site_images_folder_id(root_folder_id: str, site_folder_name: str) -> str | None:
    """Find the Drive folder ID for Bell Preventive Maintenance/<site_folder_name>/images."""
    drive = _drive()

    # Find site folder inside root
    q = (
        f"'{root_folder_id}' in parents and "
        f"name = '{site_folder_name}' and "
        "mimeType = 'application/vnd.google-apps.folder' and "
        "trashed = false"
    )
    result = drive.files().list(q=q, fields="files(id, name)", pageSize=1).execute()
    files = result.get("files", [])
    if not files:
        log.warning("Site folder '%s' not found in Drive root", site_folder_name)
        return None
    site_id = files[0]["id"]

    # Find images subfolder
    q2 = (
        f"'{site_id}' in parents and "
        "name = 'images' and "
        "mimeType = 'application/vnd.google-apps.folder' and "
        "trashed = false"
    )
    result2 = drive.files().list(q=q2, fields="files(id, name)", pageSize=1).execute()
    files2 = result2.get("files", [])
    if not files2:
        log.info("No 'images' subfolder in site '%s', using site folder directly", site_folder_name)
        return site_id
    return files2[0]["id"]


# ── PDF upload ────────────────────────────────────────────────────────────────

def upload_pdf(local_path: Path, parent_folder_id: str) -> str:
    """Upload PDF to Drive folder, replacing any existing file with the same name.

    Returns the Drive file ID of the uploaded file.
    """
    drive = _drive()
    name = local_path.name

    # Check if file already exists (to update rather than duplicate)
    q = f"'{parent_folder_id}' in parents and name = '{name}' and trashed = false"
    existing = drive.files().list(q=q, fields="files(id)", pageSize=1).execute().get("files", [])

    media = MediaFileUpload(str(local_path), mimetype="application/pdf", resumable=True)
    if existing:
        file_id = existing[0]["id"]
        drive.files().update(fileId=file_id, media_body=media).execute()
        log.info("Updated existing PDF: %s (%s)", name, file_id)
    else:
        meta = {"name": name, "parents": [parent_folder_id]}
        result = drive.files().create(body=meta, media_body=media, fields="id").execute()
        file_id = result["id"]
        log.info("Uploaded new PDF: %s (%s)", name, file_id)

    return file_id


def get_pdf_web_link(file_id: str) -> str:
    drive = _drive()
    f = drive.files().get(fileId=file_id, fields="webViewLink").execute()
    return f.get("webViewLink", "")


def find_site_folder_id(root_folder_id: str, site_folder_name: str) -> str | None:
    """Return the Drive folder ID for Bell Preventive Maintenance/<site_folder_name>."""
    drive = _drive()
    q = (
        f"'{root_folder_id}' in parents and "
        f"name = '{site_folder_name}' and "
        "mimeType = 'application/vnd.google-apps.folder' and "
        "trashed = false"
    )
    result = drive.files().list(q=q, fields="files(id, name)", pageSize=1).execute()
    files = result.get("files", [])
    return files[0]["id"] if files else None
