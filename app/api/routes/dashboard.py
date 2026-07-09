"""
Dashboard API routes — all /api/dashboard/* endpoints.
These power the standalone web dashboard and management views.
"""
import logging
from datetime import datetime, timezone
from typing import List

from fastapi import APIRouter, HTTPException, status
from sqlalchemy import func, select

from app.api.deps import AuthUser, DbSession
from app.models.domain import ClassificationStats, ProcessedEmail
from app.models.schemas import (
    AccuracyStats,
    DailyStatPoint,
    DashboardStatsResponse,
    EmailListResponse,
    EmailRecord,
    FeedbackRequest,
    FeedbackResponse,
)
from app.services.email_processor import add_feedback, update_email_category

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/api/dashboard", tags=["Dashboard"])


# ---------------------------------------------------------------------------
# GET /api/dashboard/emails
# ---------------------------------------------------------------------------

@router.get(
    "/emails",
    response_model=EmailListResponse,
    summary="List recent processed emails",
)
async def list_emails(
    db: DbSession,
    user: AuthUser,
    limit: int = 50,
    category: str | None = None,
) -> EmailListResponse:
    """Return most recent classified emails for the authenticated user."""
    stmt = (
        select(ProcessedEmail)
        .where(ProcessedEmail.user_email == user.email)
        .order_by(ProcessedEmail.processed_at.desc())
        .limit(limit)
    )
    if category:
        stmt = stmt.where(ProcessedEmail.category == category)

    result = await db.execute(stmt)
    emails = result.scalars().all()

    return EmailListResponse(
        emails=[EmailRecord.model_validate(e) for e in emails],
        total=len(emails),
    )


# ---------------------------------------------------------------------------
# GET /api/dashboard/stats
# ---------------------------------------------------------------------------

@router.get(
    "/stats",
    response_model=DashboardStatsResponse,
    summary="Classification statistics",
)
async def get_stats(
    db: DbSession,
    user: AuthUser,
    days: int = 7,
) -> DashboardStatsResponse:
    """Return daily classification stats and overall accuracy for the last N days."""

    # Daily stats
    stats_stmt = (
        select(
            ClassificationStats.date,
            ClassificationStats.category,
            func.sum(ClassificationStats.count).label("total"),
            func.avg(ClassificationStats.avg_confidence).label("avg_conf"),
        )
        .where(ClassificationStats.user_email == user.email)
        .group_by(ClassificationStats.date, ClassificationStats.category)
        .order_by(ClassificationStats.date.desc())
        .limit(days * 5)
    )
    stats_result = await db.execute(stats_stmt)
    daily_stats: List[DailyStatPoint] = [
        DailyStatPoint(
            date=r.date,
            category=r.category,
            count=r.total,
            avg_confidence=round(r.avg_conf, 1) if r.avg_conf else 0.0,
        )
        for r in stats_result.all()
    ]

    # Accuracy
    today_start = datetime.now(timezone.utc).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    total_stmt = select(func.count(ProcessedEmail.id)).where(
        ProcessedEmail.user_email == user.email,
        ProcessedEmail.processed_at >= today_start,
    )
    reviewed_stmt = select(func.count(ProcessedEmail.id)).where(
        ProcessedEmail.user_email == user.email,
        ProcessedEmail.is_correct.isnot(None),
    )
    correct_stmt = select(func.count(ProcessedEmail.id)).where(
        ProcessedEmail.user_email == user.email,
        ProcessedEmail.is_correct == True,  # noqa: E712
    )

    total = (await db.execute(total_stmt)).scalar_one()
    reviewed = (await db.execute(reviewed_stmt)).scalar_one()
    correct = (await db.execute(correct_stmt)).scalar_one()
    accuracy = round((correct / reviewed * 100) if reviewed > 0 else 0.0, 1)

    return DashboardStatsResponse(
        daily_stats=daily_stats,
        accuracy=AccuracyStats(
            total_processed=total,
            reviewed=reviewed,
            correct=correct,
            incorrect=reviewed - correct,
            accuracy_percent=accuracy,
        ),
    )


# ---------------------------------------------------------------------------
# POST /api/dashboard/feedback
# ---------------------------------------------------------------------------

@router.post(
    "/feedback",
    response_model=FeedbackResponse,
    summary="Submit correctness feedback",
)
async def submit_feedback(
    body: FeedbackRequest,
    db: DbSession,
    user: AuthUser,
) -> FeedbackResponse:
    """Record user feedback on an AI classification (correct / incorrect)."""
    ok = await add_feedback(db, body.message_id, body.is_correct, body.note)
    if not ok:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Email record not found.",
        )
    return FeedbackResponse(success=True, message_id=body.message_id)


# ---------------------------------------------------------------------------
# POST /api/dashboard/reclassify
# ---------------------------------------------------------------------------

@router.post(
    "/reclassify",
    summary="Manually override a classification",
)
async def dashboard_reclassify(
    body: dict,
    db: DbSession,
    user: AuthUser,
) -> dict:
    """Update the stored category for a message from the dashboard."""
    message_id = body.get("message_id")
    category = body.get("category")
    if not message_id or not category:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="message_id and category are required",
        )

    updated = await update_email_category(db, message_id, category)
    if not updated:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Email record not found.",
        )
    return {"success": True, "message_id": message_id, "new_category": category}
