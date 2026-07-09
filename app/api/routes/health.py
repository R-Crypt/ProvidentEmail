"""
Health check and general API endpoints.
"""
import os
import logging
from datetime import datetime, timezone
from typing import Optional, Dict

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.responses import FileResponse
from sqlalchemy import text, select
from pydantic import BaseModel

from app.core.config import settings
from app.db.session import AsyncSessionLocal
from app.api.deps import AuthUser, DbSession
from app.models.domain import ProcessedEmail
from app.models.schemas import HealthResponse

logger = logging.getLogger(__name__)
router = APIRouter(tags=["Health / General"])


# Global in-memory auto-reply state
AUTO_REPLY_ENABLED = False


@router.get("/health", response_model=HealthResponse, summary="Liveness check")
async def health() -> HealthResponse:
    """
    Liveness probe — returns 200 as long as the process is running.
    Does NOT check external dependencies (DB, OpenAI).
    """
    return HealthResponse(
        status="ok",
        version=settings.APP_VERSION,
        environment=settings.ENVIRONMENT,
        bypass_auth_active=settings.BYPASS_AUTH,
    )


@router.get("/ready", response_model=HealthResponse, summary="Readiness check")
async def ready() -> HealthResponse:
    """
    Readiness probe — checks that the database is reachable.
    """
    db_ok = False
    details: dict = {}

    try:
        async with AsyncSessionLocal() as session:
            await session.execute(text("SELECT 1"))
        db_ok = True
    except Exception as exc:
        logger.error("DB readiness check failed", extra={"error": str(exc)})
        details["db_error"] = str(exc)

    # Check critical config
    missing = settings.validate_required()
    if missing:
        details["missing_config"] = missing

    status_str = "ok" if (db_ok and not missing) else "degraded"

    return HealthResponse(
        status=status_str,
        version=settings.APP_VERSION,
        environment=settings.ENVIRONMENT,
        db_connected=db_ok,
        bypass_auth_active=settings.BYPASS_AUTH,
        details=details,
    )


# ---------------------------------------------------------------------------
# Serve Outlook Add-in manifest XML file directly from URL
# ---------------------------------------------------------------------------

@router.get(
    "/manifest.xml",
    summary="Serve the Outlook Add-in manifest XML file",
)
async def serve_manifest_xml():
    """
    Exposes the manifest.xml on the root path with the correct XML MIME-type.
    This enables seamless 'Add from URL' installations.
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(current_dir, "..", "..", ".."))
    manifest_path = os.path.join(project_root, "outlook-addin", "manifest.xml")

    if not os.path.exists(manifest_path):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Manifest XML file not found.",
        )

    return FileResponse(manifest_path, media_type="application/xml")


# ---------------------------------------------------------------------------
# General App Settings and Reply Endpoints
# ---------------------------------------------------------------------------

class SettingsRequest(BaseModel):
    auto_reply: bool

class SendReplyRequest(BaseModel):
    message_id: str
    reply_text: str


@router.get("/api/settings", summary="Get global settings")
async def get_settings_endpoint(
    user: AuthUser,
):
    global AUTO_REPLY_ENABLED
    return {"success": True, "auto_reply": AUTO_REPLY_ENABLED}


@router.post("/api/settings", summary="Update global settings")
async def post_settings_endpoint(
    body: SettingsRequest,
    user: AuthUser,
):
    global AUTO_REPLY_ENABLED
    AUTO_REPLY_ENABLED = body.auto_reply
    return {"success": True, "auto_reply": AUTO_REPLY_ENABLED}


@router.post("/api/send_reply", summary="Record a sent reply and advance email status")
async def send_reply_endpoint(
    body: SendReplyRequest,
    db: DbSession,
    user: AuthUser,
):
    stmt = select(ProcessedEmail).where(ProcessedEmail.message_id == body.message_id)
    result = await db.execute(stmt)
    email = result.scalar_one_or_none()
    if not email:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Email record not found.",
        )

    email.reply_sent = True
    email.sent_reply = body.reply_text
    email.reply_sent_at = datetime.now(timezone.utc)

    # Advance status to the reply-triggered status
    from src.database import REPLY_ADVANCE_STATUS
    target_status = REPLY_ADVANCE_STATUS.get(email.category)
    if target_status:
        email.email_status = target_status
        email.status_updated_at = datetime.now(timezone.utc)

    await db.commit()
    return {"success": True, "new_status": email.email_status or ""}
