import os
import secrets

from fastapi import HTTPException, Request


def _expected_password() -> str | None:
    pw = os.environ.get("TOOLS_PASSWORD")
    return pw if pw else None


def is_authed(request: Request) -> bool:
    return bool(request.session.get("authed"))


def verify_password(submitted: str) -> bool:
    expected = _expected_password()
    if not expected:
        return False
    return secrets.compare_digest(submitted, expected)


def require_auth(request: Request) -> None:
    if is_authed(request):
        return
    if not _expected_password():
        raise HTTPException(
            status_code=503,
            detail="This tool is gated but TOOLS_PASSWORD is not configured on the server.",
        )
    next_url = request.url.path
    if request.url.query:
        next_url = f"{next_url}?{request.url.query}"
    raise HTTPException(
        status_code=303,
        detail="Authentication required",
        headers={"Location": f"/login?next={next_url}"},
    )
