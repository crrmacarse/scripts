from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse

from app.auth import is_authed
from app.templating import templates
from app.tools_registry import TOOLS

router = APIRouter()


@router.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(
        request,
        "home.html",
        {"tools": TOOLS, "authed": is_authed(request)},
    )
