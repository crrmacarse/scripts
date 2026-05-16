from fastapi import APIRouter, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse

from app.services import money_manager as svc
from app.templating import templates

router = APIRouter(prefix="/tools/money-manager")


@router.get("", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(
        request,
        "money_manager/index.html",
        {"expected_columns": svc.EXPECTED_COLUMNS, "error": None},
    )


@router.post("/analyze", response_class=HTMLResponse)
async def analyze(request: Request, mm_file: UploadFile = File(...)):
    try:
        contents = await mm_file.read()
        if not contents:
            raise HTTPException(status_code=400, detail="xlsx file is empty.")
        result = svc.analyze(contents)
    except ValueError as exc:
        return templates.TemplateResponse(
            request,
            "money_manager/index.html",
            {"expected_columns": svc.EXPECTED_COLUMNS, "error": str(exc)},
            status_code=400,
        )
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        return templates.TemplateResponse(
            request,
            "money_manager/index.html",
            {
                "expected_columns": svc.EXPECTED_COLUMNS,
                "error": f"Could not parse xlsx: {exc}",
            },
            status_code=500,
        )

    return templates.TemplateResponse(
        request,
        "money_manager/result.html",
        {"result": result, "expected_columns": svc.EXPECTED_COLUMNS},
    )
