import base64
import json
import os
from pathlib import Path

import gspread

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
DEFAULT_LOCAL_FILE = REPO_ROOT / "google-service-account.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _resolve_service_account_dict() -> dict:
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if raw:
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "GOOGLE_SERVICE_ACCOUNT_JSON is set but is not valid JSON."
            ) from exc

    b64 = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON_B64")
    if b64:
        try:
            return json.loads(base64.b64decode(b64))
        except Exception as exc:
            raise RuntimeError(
                "GOOGLE_SERVICE_ACCOUNT_JSON_B64 is set but could not be decoded."
            ) from exc

    explicit_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE")
    if explicit_path:
        path = Path(explicit_path).expanduser()
        if not path.exists():
            raise RuntimeError(
                f"GOOGLE_SERVICE_ACCOUNT_FILE points to {path}, which does not exist."
            )
        return json.loads(path.read_text())

    if DEFAULT_LOCAL_FILE.exists():
        return json.loads(DEFAULT_LOCAL_FILE.read_text())

    raise RuntimeError(
        "No Google service account credentials found. Provide one of: "
        "GOOGLE_SERVICE_ACCOUNT_JSON (raw JSON), "
        "GOOGLE_SERVICE_ACCOUNT_JSON_B64 (base64 JSON), "
        "GOOGLE_SERVICE_ACCOUNT_FILE (path to JSON), "
        f"or place google-service-account.json at {REPO_ROOT}."
    )


def get_gspread_client() -> gspread.Client:
    creds_dict = _resolve_service_account_dict()
    return gspread.service_account_from_dict(creds_dict, scopes=SCOPES)
