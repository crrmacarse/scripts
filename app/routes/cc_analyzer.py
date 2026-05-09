import base64
import json
from datetime import datetime
from typing import Annotated

from fastapi import APIRouter, Depends, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse

from app.auth import require_auth
from app.services import cc_analyzer as svc
from app.templating import templates

router = APIRouter(
    prefix="/tools/cc-analyzer",
    dependencies=[Depends(require_auth)],
)

_DEFAULT_CARD_KEY = svc.CREDIT_CARDS[0].key


def _parse_cutoff(cutoff_str: str) -> datetime:
    for fmt in ("%Y-%m-%d", "%m/%d/%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(cutoff_str, fmt)
        except ValueError:
            continue
    raise HTTPException(status_code=400, detail=f"Invalid cutoff date: {cutoff_str!r}")


def _index_context(card_key: str = _DEFAULT_CARD_KEY, error: str | None = None) -> dict:
    return {
        "credit_cards": svc.CREDIT_CARDS,
        "selected_card_key": card_key,
        "error": error,
    }


def _build_reanalyze_token(
    pdf_extraction: svc.PdfExtraction,
    cutoff_dt: datetime,
    card: svc.CreditCard,
    mm_source: str,
) -> str:
    """Encode parsed PDF + metadata so Re-analyze can rerun without re-upload."""
    payload = {
        "pdf_rows": [
            {
                "tran_date": r.tran_date,
                "post_date": r.post_date,
                "description": r.description,
                "amount": r.amount,
            }
            for r in pdf_extraction.rows
        ],
        "total_amount": pdf_extraction.total_amount,
        "total_cr": pdf_extraction.total_cr,
        "cutoff_date_iso": cutoff_dt.strftime("%Y-%m-%d"),
        "credit_card_key": card.key,
        "mm_source": mm_source,
    }
    return base64.b64encode(json.dumps(payload).encode()).decode()


def _decode_reanalyze_token(
    token: str,
) -> tuple[svc.PdfExtraction, datetime, svc.CreditCard, str]:
    payload = json.loads(base64.b64decode(token))
    card = svc.get_credit_card(payload["credit_card_key"])
    cutoff_dt = datetime.strptime(payload["cutoff_date_iso"], "%Y-%m-%d")
    pdf_extraction = svc.PdfExtraction(
        rows=[svc.PdfRow(**r) for r in payload["pdf_rows"]],
        total_amount=payload["total_amount"],
        total_cr=payload["total_cr"],
    )
    mm_source = payload.get("mm_source", "google_sheet")
    return pdf_extraction, cutoff_dt, card, mm_source


def _render_result(
    request: Request,
    *,
    pdf_extraction: svc.PdfExtraction,
    cutoff_dt: datetime,
    card: svc.CreditCard,
    mm_entries: list[svc.MMEntry],
    mm_source: str,
):
    billing_period = cutoff_dt.strftime("%B %Y")
    result = svc.analyse(
        pdf_extraction=pdf_extraction,
        mm_entries=mm_entries,
        sheet_name=card.sheet_name,
        billing_period=billing_period,
        cutoff_date=cutoff_dt,
        card=card,
    )
    confirm_token = base64.b64encode(
        json.dumps(result.serialise_for_confirm()).encode()
    ).decode()
    reanalyze_token = _build_reanalyze_token(pdf_extraction, cutoff_dt, card, mm_source)

    return templates.TemplateResponse(
        request,
        "cc_analyzer/result.html",
        {
            "result": result,
            "confirm_token": confirm_token,
            "reanalyze_token": reanalyze_token,
            "credit_card": card,
            "mm_source": mm_source,
        },
    )


@router.get("", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(
        request,
        "cc_analyzer/index.html",
        _index_context(),
    )


@router.post("/analyze", response_class=HTMLResponse)
async def analyze(
    request: Request,
    credit_card: Annotated[str, Form()] = _DEFAULT_CARD_KEY,
    cutoff_date: Annotated[str, Form()] = "",
    pdf_password: Annotated[str, Form()] = "",
    mm_source: Annotated[str, Form()] = "google_sheet",
    pdf_file: UploadFile = File(...),
    mm_file: UploadFile | None = File(None),
):
    try:
        card = svc.get_credit_card(credit_card)
        cutoff_dt = _parse_cutoff(cutoff_date)

        pdf_bytes = await pdf_file.read()
        if not pdf_bytes:
            raise HTTPException(status_code=400, detail="PDF file is empty.")

        pdf_extraction = svc.extract_pdf(pdf_bytes, pdf_password or None)
        window_start, window_end = svc.compute_mm_window(cutoff_dt)

        if mm_source == "xlsx":
            if mm_file is None or not mm_file.filename:
                raise HTTPException(
                    status_code=400,
                    detail="Money Manager xlsx upload selected but no file was provided.",
                )
            mm_bytes = await mm_file.read()
            mm_entries = svc.load_mm_from_xlsx(
                mm_bytes, window_start, window_end, card.mm_account
            )
        else:
            mm_entries = svc.load_mm_from_google_sheet(
                window_start, window_end, card.mm_account
            )
    except HTTPException:
        raise
    except ValueError as exc:
        return templates.TemplateResponse(
            request,
            "cc_analyzer/index.html",
            _index_context(card_key=credit_card, error=str(exc)),
            status_code=400,
        )
    except Exception as exc:  # noqa: BLE001
        return templates.TemplateResponse(
            request,
            "cc_analyzer/index.html",
            _index_context(card_key=credit_card, error=f"Unexpected error: {exc}"),
            status_code=500,
        )

    return _render_result(
        request,
        pdf_extraction=pdf_extraction,
        cutoff_dt=cutoff_dt,
        card=card,
        mm_entries=mm_entries,
        mm_source=mm_source,
    )


@router.post("/reanalyze", response_class=HTMLResponse)
async def reanalyze(
    request: Request,
    reanalyze_token: Annotated[str, Form()],
):
    try:
        pdf_extraction, cutoff_dt, card, mm_source = _decode_reanalyze_token(reanalyze_token)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"Invalid re-analyze token: {exc}")

    if mm_source != "google_sheet":
        # We don't have the original xlsx bytes — user needs to re-upload.
        return templates.TemplateResponse(
            request,
            "cc_analyzer/index.html",
            _index_context(
                card_key=card.key,
                error="Re-analyze without re-upload only works for the Google Sheet source. Please re-upload the xlsx.",
            ),
            status_code=400,
        )

    try:
        window_start, window_end = svc.compute_mm_window(cutoff_dt)
        mm_entries = svc.load_mm_from_google_sheet(
            window_start, window_end, card.mm_account
        )
    except ValueError as exc:
        return templates.TemplateResponse(
            request,
            "cc_analyzer/index.html",
            _index_context(card_key=card.key, error=str(exc)),
            status_code=400,
        )
    except Exception as exc:  # noqa: BLE001
        return templates.TemplateResponse(
            request,
            "cc_analyzer/index.html",
            _index_context(card_key=card.key, error=f"Unexpected error: {exc}"),
            status_code=500,
        )

    return _render_result(
        request,
        pdf_extraction=pdf_extraction,
        cutoff_dt=cutoff_dt,
        card=card,
        mm_entries=mm_entries,
        mm_source=mm_source,
    )


@router.post("/confirm", response_class=HTMLResponse)
async def confirm(
    request: Request,
    confirm_token: Annotated[str, Form()],
    shoulder_data: Annotated[str, Form()] = "",
    force_overwrite: Annotated[str, Form()] = "",
):
    try:
        payload = json.loads(base64.b64decode(confirm_token))
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"Invalid confirm token: {exc}")

    if shoulder_data:
        try:
            payload["shoulders"] = json.loads(shoulder_data)
        except Exception:  # noqa: BLE001
            payload["shoulders"] = []
    overwrite = force_overwrite == "1"

    try:
        sheet_url = svc.push_to_google_sheet(payload, force_overwrite=overwrite)
    except svc.WorksheetExistsError:
        return templates.TemplateResponse(
            request,
            "cc_analyzer/overwrite_confirm.html",
            {
                "billing_period": payload.get("billing_period"),
                "sheet_name": payload.get("sheet_name"),
                "confirm_token": confirm_token,
                "shoulder_data": shoulder_data,
            },
        )
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
