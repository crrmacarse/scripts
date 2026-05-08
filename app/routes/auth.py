from fastapi import APIRouter, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse

from app.auth import verify_password
from app.templating import templates

router = APIRouter()


def _safe_next(next_url: str | None) -> str:
    if not next_url:
        return "/"
    if not next_url.startswith("/") or next_url.startswith("//"):
        return "/"
    return next_url


@router.get("/login", response_class=HTMLResponse)
async def login_form(request: Request, next: str | None = None):
    return templates.TemplateResponse(
        request,
        "login.html",
        {"next": _safe_next(next), "error": None},
    )


@router.post("/login")
async def login_submit(
    request: Request,
    password: str = Form(...),
    next: str = Form("/"),
):
    if not verify_password(password):
        return templates.TemplateResponse(
            request,
            "login.html",
            {"next": _safe_next(next), "error": "Wrong password."},
            status_code=401,
        )
    request.session["authed"] = True
    return RedirectResponse(_safe_next(next), status_code=303)


@router.post("/logout")
async def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/", status_code=303)
