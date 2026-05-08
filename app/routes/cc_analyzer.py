import base64
import json
from datetime import datetime
from typing import Annotated

from fastapi import APIRouter, Depends, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse

from app.auth import require_auth
from app.services import cc_analyzer as svc
from app.templating import templates

router = APIRouter(
    prefix="/tools/cc-analyzer",
    dependencies=[Depends(require_auth)],
)

_DEFAULT_SHEET_NAME = "Security Bank World CC"


def _parse_cutoff(cutoff_str: str) -> datetime:
    for fmt in ("%Y-%m-%d", "%m/%d/%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(cutoff_str, fmt)
        except ValueError:
            continue
    raise HTTPException(status_code=400, detail=f"Invalid cutoff date: {cutoff_str!r}")


@router.get("", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(
        request,
        "cc_analyzer/index.html",
        {
            "default_sheet_name": _DEFAULT_SHEET_NAME,
            "error": None,
        },
    )


@router.post("/analyze", response_class=HTMLResponse)
async def analyze(
    request: Request,
    sheet_name: Annotated[str, Form()] = _DEFAULT_SHEET_NAME,
    cutoff_date: Annotated[str, Form()] = "",
    pdf_password: Annotated[str, Form()] = "",
    mm_source: Annotated[str, Form()] = "google_sheet",
    pdf_file: UploadFile = File(...),
    mm_file: UploadFile | None = File(None),
):
    try:
        cutoff_dt = _parse_cutoff(cutoff_date)
        billing_period = cutoff_dt.strftime("%B %Y")

        pdf_bytes = await pdf_file.read()
        if not pdf_bytes:
            raise HTTPException(status_code=400, detail="PDF file is empty.")

        if mm_source == "xlsx":
            if mm_file is None or not mm_file.filename:
                raise HTTPException(
                    status_code=400,
                    detail="Money Manager xlsx upload selected but no file was provided.",
                )
            mm_bytes = await mm_file.read()
            mm_entries = svc.load_mm_from_xlsx(mm_bytes, cutoff_dt)
        else:
            mm_entries = svc.load_mm_from_google_sheet(cutoff_dt)

        pdf_extraction = svc.extract_pdf(pdf_bytes, pdf_password or None)
        result = svc.analyse(
            pdf_extraction=pdf_extraction,
            mm_entries=mm_entries,
            sheet_name=sheet_name,
            billing_period=billing_period,
            cutoff_date=cutoff_dt,
        )
    except HTTPException:
        raise
    except ValueError as exc:
        return templates.TemplateResponse(
            request,
            "cc_analyzer/index.html",
            {
                "default_sheet_name": sheet_name or _DEFAULT_SHEET_NAME,
                "error": str(exc),
            },
            status_code=400,
        )
    except Exception as exc:  # noqa: BLE001
        return templates.TemplateResponse(
            request,
            "cc_analyzer/index.html",
            {
                "default_sheet_name": sheet_name or _DEFAULT_SHEET_NAME,
                "error": f"Unexpected error: {exc}",
            },
            status_code=500,
        )

    payload = result.serialise_for_confirm()
    confirm_token = base64.b64encode(json.dumps(payload).encode()).decode()

    return templates.TemplateResponse(
        request,
        "cc_analyzer/result.html",
        {
            "result": result,
            "confirm_token": confirm_token,
        },
    )


@router.post("/confirm", response_class=HTMLResponse)
async def confirm(
    request: Request,
    confirm_token: Annotated[str, Form()],
):
    try:
        payload = json.loads(base64.b64decode(confirm_token))
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"Invalid confirm token: {exc}")

    try:
        sheet_url = svc.push_to_google_sheet(payload)
    except ValueError as exc:
        return templates.TemplateResponse(
            request,
            "cc_analyzer/done.html",
            {
                "error": str(exc),
                "sheet_url": None,
                "billing_period": payload.get("billing_period"),
            },
            status_code=400,
        )
    except Exception as exc:  # noqa: BLE001
        return templates.TemplateResponse(
            request,
            "cc_analyzer/done.html",
            {
                "error": f"Unexpected error: {exc}",
                "sheet_url": None,
                "billing_period": payload.get("billing_period"),
            },
            status_code=500,
        )

    return templates.TemplateResponse(
        request,
        "cc_analyzer/done.html",
        {
            "error": None,
            "sheet_url": sheet_url,
            "billing_period": payload.get("billing_period"),
        },
    )
