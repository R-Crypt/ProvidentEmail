"""
Shared FastAPI dependencies injected into route handlers.
"""
import logging
from typing import Annotated, Optional

from fastapi import Depends, Header, HTTPException, status
from jose import JWTError
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import settings
from app.core.security import (
    decode_internal_token,
    extract_user_email_from_token,
    verify_ms_token,
)
from app.db.session import get_db
from app.models.schemas import CurrentUser

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Database dependency — re-exported here for a single import point
# ---------------------------------------------------------------------------

DbSession = Annotated[AsyncSession, Depends(get_db)]


# ---------------------------------------------------------------------------
# Authentication dependency
# ---------------------------------------------------------------------------

async def get_current_user(
    authorization: Annotated[Optional[str], Header()] = None,
) -> CurrentUser:
    """
    Extract and validate the Bearer token from the Authorization header.

    Accepts two token types:
    1. Microsoft SSO tokens (RS256, from Office.auth.getAccessTokenAsync)
    2. Internal JWT tokens (HS256, for dev/testing via /api/auth/token)

    Raises HTTP 401 if the token is missing or invalid.
    """
    if settings.BYPASS_AUTH:
        return CurrentUser(
            email="rayaankhaaan@outlook.com",
            tenant_id="00ae596d-5d33-44c2-8df8-2f1933056560",
            display_name="Rayaan Khan",
        )

    credentials_exc = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Could not validate credentials. Please provide a valid Bearer token.",
        headers={"WWW-Authenticate": "Bearer"},
    )

    if not authorization or not authorization.startswith("Bearer "):
        raise credentials_exc

    token = authorization.removeprefix("Bearer ").strip()

    # Try Microsoft SSO token first (production path)
    try:
        payload = await verify_ms_token(token)
        user_email = extract_user_email_from_token(payload)
        return CurrentUser(
            email=user_email,
            tenant_id=payload.get("tid", ""),
            display_name=payload.get("name", user_email),
        )
    except JWTError:
        pass  # Fall through to internal token check

    # Try internal JWT (dev/testing path)
    try:
        payload = decode_internal_token(token)
        return CurrentUser(
            email=payload.get("sub", "unknown@providentpackaging.com"),
            tenant_id=payload.get("tid", ""),
            display_name=payload.get("sub", ""),
        )
    except JWTError:
        logger.warning("Token validation failed for request")
        raise credentials_exc


# Convenience annotated type for use in route signatures
AuthUser = Annotated[CurrentUser, Depends(get_current_user)]
