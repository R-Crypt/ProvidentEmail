"""
Health check endpoints.
Used by Docker healthcheck, load balancers, and uptime monitors.
"""
import logging

from fastapi import APIRouter
from sqlalchemy import text

from app.core.config import settings
from app.db.session import AsyncSessionLocal
from app.models.schemas import HealthResponse

logger = logging.getLogger(__name__)
router = APIRouter(tags=["Health"])


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
    )


@router.get("/ready", response_model=HealthResponse, summary="Readiness check")
async def ready() -> HealthResponse:
    """
    Readiness probe — checks that the database is reachable.
    Load balancers should use this endpoint to decide whether to send traffic.
    Returns 503 if the DB is not reachable.
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
        details=details,
    )
