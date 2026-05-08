import base64
import json
import os
import tempfile
from pathlib import Path

import gspread
from oauth2client.service_account import ServiceAccountCredentials

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]


def _resolve_service_account_dict() -> dict:
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if raw:
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            pass

    b64 = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON_B64")
    if b64:
        try:
            return json.loads(base64.b64decode(b64))
        except Exception as exc:
            raise RuntimeError(
                "GOOGLE_SERVICE_ACCOUNT_JSON_B64 is set but could not be decoded."
            ) from exc

    local_path = Path(__file__).resolve().parent.parent.parent / "google-service-account.json"
    if local_path.exists():
        return json.loads(local_path.read_text())

    raise RuntimeError(
        "No Google service account credentials found. Set "
        "GOOGLE_SERVICE_ACCOUNT_JSON (raw JSON) or "
        "GOOGLE_SERVICE_ACCOUNT_JSON_B64 (base64-encoded JSON)."
    )


def get_gspread_client() -> gspread.Client:
    creds_dict = _resolve_service_account_dict()
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as f:
        json.dump(creds_dict, f)
        path = f.name
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(path, SCOPE)
        return gspread.authorize(creds)
    finally:
        try:
            os.unlink(path)
        except OSError:
            pass
