"""
Background email processing scheduler.
Wired to FastAPI's lifespan so it starts/stops with the app,
rather than running as a separate standalone process.
"""
import logging

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.interval import IntervalTrigger

from app.core.config import settings
from app.db.session import get_db_context
from app.services.email_processor import run_email_batch

logger = logging.getLogger(__name__)

_scheduler = AsyncIOScheduler()


async def _run_batch_job() -> None:
    """Wrapper called by APScheduler — manages its own DB session."""
    logger.info("Scheduled email batch starting")
    try:
        async with get_db_context() as db:
            stats = await run_email_batch(db)
        logger.info("Scheduled email batch complete", extra=stats)
    except Exception as exc:
        logger.error(
            "Scheduled email batch failed",
            extra={"error": str(exc)},
            exc_info=True,
        )


def start_scheduler() -> None:
    """Start the background APScheduler. Called from FastAPI lifespan on startup."""
    missing = settings.validate_required()
    if missing:
        logger.warning(
            "Scheduler not started — missing required config",
            extra={"missing": missing},
        )
        return

    _scheduler.add_job(
        _run_batch_job,
        trigger=IntervalTrigger(minutes=settings.CHECK_INTERVAL_MINUTES),
        id="email_classifier_batch",
        name="Email Classification Batch",
        replace_existing=True,
        max_instances=1,  # Prevent overlapping runs
    )

    _scheduler.start()
    logger.info(
        "Email classifier scheduler started",
        extra={"interval_minutes": settings.CHECK_INTERVAL_MINUTES},
    )


def stop_scheduler() -> None:
    """Stop the scheduler gracefully. Called from FastAPI lifespan on shutdown."""
    if _scheduler.running:
        _scheduler.shutdown(wait=False)
        logger.info("Email classifier scheduler stopped")
