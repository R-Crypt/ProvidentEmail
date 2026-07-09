"""
Security utilities — JWT creation/validation and Microsoft OAuth token verification.

Auth strategy: Microsoft OAuth (SSO)
The Office.js taskpane calls `Office.auth.getAccessTokenAsync()` to get a signed
JWT issued by Microsoft for the signed-in Outlook user. The backend validates this
token using Microsoft's JWKS endpoint, so we never store passwords.

For local development / API testing, we also support issuing internal JWTs via
the /api/auth/token endpoint with a pre-shared secret.
"""
import logging
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Optional

import httpx
from jose import JWTError, jwt
from jose.exceptions import ExpiredSignatureError

from app.core.config import settings

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Internal JWT (used for service-to-service calls or dev testing)
# ---------------------------------------------------------------------------

def create_access_token(
    subject: str,
    tenant_id: str = "",
    extra_claims: Optional[Dict[str, Any]] = None,
    expires_delta: Optional[timedelta] = None,
) -> str:
    """Create a signed internal JWT token."""
    expire = datetime.now(timezone.utc) + (
        expires_delta or timedelta(minutes=settings.ACCESS_TOKEN_EXPIRE_MINUTES)
    )
    payload: Dict[str, Any] = {
        "sub": subject,
        "tid": tenant_id,
        "exp": expire,
        "iat": datetime.now(timezone.utc),
        "iss": settings.APP_NAME,
    }
    if extra_claims:
        payload.update(extra_claims)
    return jwt.encode(payload, settings.SECRET_KEY, algorithm=settings.ALGORITHM)


def decode_internal_token(token: str) -> Dict[str, Any]:
    """
    Decode and validate an internally-issued JWT.
    Raises jose.JWTError on failure.
    """
    return jwt.decode(token, settings.SECRET_KEY, algorithms=[settings.ALGORITHM])


# ---------------------------------------------------------------------------
# Microsoft OAuth token verification
# ---------------------------------------------------------------------------

# Simple in-memory JWKS cache to avoid hitting Microsoft on every request
_ms_jwks_cache: Optional[Dict] = None
_ms_jwks_fetched_at: Optional[datetime] = None
_MS_JWKS_TTL_SECONDS = 3600  # Re-fetch keys every hour


async def _get_ms_jwks() -> Dict:
    """Fetch and cache Microsoft's public signing keys."""
    global _ms_jwks_cache, _ms_jwks_fetched_at

    now = datetime.now(timezone.utc)
    if (
        _ms_jwks_cache
        and _ms_jwks_fetched_at
        and (now - _ms_jwks_fetched_at).total_seconds() < _MS_JWKS_TTL_SECONDS
    ):
        return _ms_jwks_cache

    async with httpx.AsyncClient(timeout=10) as client:
        resp = await client.get(settings.ms_jwks_uri)
        resp.raise_for_status()
        _ms_jwks_cache = resp.json()
        _ms_jwks_fetched_at = now
        logger.info("Microsoft JWKS keys refreshed")
        return _ms_jwks_cache


async def verify_ms_token(token: str) -> Dict[str, Any]:
    """
    Validate a token issued by Microsoft (from Office.auth.getAccessTokenAsync).
    Returns the decoded payload on success, raises JWTError on failure.

    The token audience is the CLIENT_ID of your Azure AD App Registration.
    """
    try:
        jwks = await _get_ms_jwks()
        payload = jwt.decode(
            token,
            jwks,
            algorithms=["RS256"],
            audience=settings.CLIENT_ID,
            options={"verify_exp": True},
        )
        return payload
    except ExpiredSignatureError:
        logger.warning("Microsoft token expired")
        raise JWTError("Token expired")
    except JWTError as exc:
        logger.warning("Microsoft token validation failed", extra={"error": str(exc)})
        raise


def extract_user_email_from_token(payload: Dict[str, Any]) -> str:
    """
    Extract the user's email from a Microsoft JWT payload.
    Microsoft uses 'preferred_username' or 'upn' for the email.
    """
    return (
        payload.get("preferred_username")
        or payload.get("upn")
        or payload.get("email")
        or payload.get("sub", "unknown@unknown.com")
    )
