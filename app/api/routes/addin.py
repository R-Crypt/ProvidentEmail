"""
Add-in API routes — all /api/addin/* endpoints.
These are called directly by taskpane.js running inside Outlook.
All routes require a valid Bearer token (Microsoft SSO or internal JWT).
"""
import logging
from datetime import datetime, timezone

from fastapi import APIRouter, HTTPException, status
from sqlalchemy import func, select
from pydantic import BaseModel

from app.api.deps import AuthUser, DbSession
from app.models.domain import ClassificationStats, ProcessedEmail
from app.models.schemas import (
    CategoryCounts,
    ClassifyRequest,
    ClassifyResponse,
    EmailListResponse,
    EmailRecord,
    FeedbackRequest,
    FeedbackResponse,
    ReclassifyRequest,
    ReclassifyResponse,
    TriageSummaryResponse,
)
from app.services.email_processor import (
    add_feedback,
    get_or_classify,
    update_email_category,
)

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/api/addin", tags=["Add-in"])


# ---------------------------------------------------------------------------
# POST /api/addin/classify
# ---------------------------------------------------------------------------

@router.post(
    "/classify",
    response_model=ClassifyResponse,
    summary="Classify an email (with DB caching)",
)
async def addin_classify(
    body: ClassifyRequest,
    db: DbSession,
    user: AuthUser,
) -> ClassifyResponse:
    """
    Classify an email by subject + body.
    If this message_id was already classified, return the cached result instantly.
    The user's email is taken from their authenticated token (not hardcoded).
    """
    try:
        result, from_cache = await get_or_classify(
            db=db,
            message_id=body.message_id,
            subject=body.subject,
            body=body.body,
            sender=body.sender,
            user_email=user.email,  # Derived from token — no hardcoding
        )

        return ClassifyResponse(
            message_id=result["message_id"],
            category=result["category"],
            confidence=result["confidence"],
            reason=result["reason"],
            extracted_data=result.get("extracted_data"),
            response_draft=result.get("response_draft"),
            processed_at=result.get("processed_at"),
            from_cache=from_cache,
        )

    except Exception as exc:
        logger.error(
            "classify endpoint error",
            extra={"message_id": body.message_id, "error": str(exc)},
            exc_info=True,
        )
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Classification failed: {exc}",
        )


# ---------------------------------------------------------------------------
# POST /api/addin/reclassify
# ---------------------------------------------------------------------------

@router.post(
    "/reclassify",
    response_model=ReclassifyResponse,
    summary="Override the AI's category for an email",
)
async def addin_reclassify(
    body: ReclassifyRequest,
    db: DbSession,
    user: AuthUser,
) -> ReclassifyResponse:
    """
    Manually override the stored category for a message.
    If the record doesn't exist yet, it's created as a manual-override entry.
    """
    from app.core.config import settings

    updated = await update_email_category(db, body.message_id, body.category)

    if not updated:
        # Create a shell record (user overriding before the scheduler classified it)
        from app.models.domain import ProcessedEmail as PE
        outlook_cat = settings.CATEGORIES.get(body.category, {}).get("outlook_category", "General")
        shell = PE(
            message_id=body.message_id,
            user_email=user.email,
            sender="unknown",
            subject="Manual Override",
            body_preview="",
            category=body.category,
            confidence=100.0,
            reason="User manually set category in Outlook Add-in",
            outlook_category_applied=outlook_cat,
            received_at=datetime.now(timezone.utc),
        )
        db.add(shell)

    logger.info(
        "Email reclassified",
        extra={
            "message_id": body.message_id,
            "new_category": body.category,
            "user": user.email,
        },
    )

    return ReclassifyResponse(
        success=True,
        message_id=body.message_id,
        new_category=body.category,
    )


# ---------------------------------------------------------------------------
# GET /api/addin/triage_summary
# ---------------------------------------------------------------------------

@router.get(
    "/triage_summary",
    response_model=TriageSummaryResponse,
    summary="Today's email counts per category",
)
async def addin_triage_summary(
    db: DbSession,
    user: AuthUser,
) -> TriageSummaryResponse:
    """
    Return counts of emails classified today, grouped by category.
    Only returns emails belonging to the authenticated user's mailbox.
    """
    today_start = datetime.now(timezone.utc).replace(
        hour=0, minute=0, second=0, microsecond=0
    )

    stmt = (
        select(ProcessedEmail.category, func.count(ProcessedEmail.id))
        .where(
            ProcessedEmail.processed_at >= today_start,
            ProcessedEmail.user_email == user.email,
        )
        .group_by(ProcessedEmail.category)
    )
    result = await db.execute(stmt)
    rows = result.all()

    counts = CategoryCounts()
    total = 0
    for cat, count in rows:
        if hasattr(counts, cat):
            setattr(counts, cat, count)
            total += count
        else:
            counts.general += count
            total += count

    return TriageSummaryResponse(total=total, counts=counts)


# ---------------------------------------------------------------------------
# GET /api/addin/category_emails
# ---------------------------------------------------------------------------

@router.get(
    "/category_emails",
    response_model=EmailListResponse,
    summary="List processed emails in a category",
)
async def addin_category_emails(
    category: str,
    db: DbSession,
    user: AuthUser,
    limit: int = 50,
) -> EmailListResponse:
    """Return the most recent emails classified under a given category for this user."""
    valid = {"purchase_order", "enquiry", "invoice", "shipping", "general"}
    if category not in valid:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=f"Invalid category. Must be one of: {', '.join(sorted(valid))}",
        )

    stmt = (
        select(ProcessedEmail)
        .where(
            ProcessedEmail.category == category,
            ProcessedEmail.user_email == user.email,
        )
        .order_by(ProcessedEmail.processed_at.desc())
        .limit(limit)
    )
    result = await db.execute(stmt)
    emails = result.scalars().all()

    return EmailListResponse(
        emails=[EmailRecord.model_validate(e) for e in emails],
        total=len(emails),
    )


# ---------------------------------------------------------------------------
# POST /api/addin/feedback
# ---------------------------------------------------------------------------

@router.post(
    "/feedback",
    response_model=FeedbackResponse,
    summary="Submit feedback on a classification",
)
async def addin_feedback(
    body: FeedbackRequest,
    db: DbSession,
    user: AuthUser,
) -> FeedbackResponse:
    """Record whether the AI's classification was correct."""
    ok = await add_feedback(db, body.message_id, body.is_correct, body.note)
    if not ok:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Email record not found.",
        )
    return FeedbackResponse(success=True, message_id=body.message_id)


# ---------------------------------------------------------------------------
# POST /api/addin/update_status
# ---------------------------------------------------------------------------

class UpdateStatusRequest(BaseModel):
    message_id: str
    status: str

@router.post(
    "/update_status",
    summary="Update the lifecycle status of an email",
)
async def addin_update_status(
    body: UpdateStatusRequest,
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

    email.email_status = body.status
    email.status_updated_at = datetime.now(timezone.utc)
    await db.commit()
    return {"success": True}


# ---------------------------------------------------------------------------
# GET /api/addin/status_flows
# ---------------------------------------------------------------------------

@router.get(
    "/status_flows",
    summary="Get status flows for the frontend stepper",
)
async def addin_status_flows(
    user: AuthUser,
):
    from src.database import STATUS_FLOWS, NEXT_STATUS
    return {"success": True, "flows": STATUS_FLOWS, "next_status": NEXT_STATUS}


# ---------------------------------------------------------------------------
# GET /api/addin/stale_alerts
# ---------------------------------------------------------------------------

@router.get(
    "/stale_alerts",
    summary="Get emails requiring attention",
)
async def addin_stale_alerts(
    db: DbSession,
    user: AuthUser,
    hours: int = 24,
):
    from datetime import timedelta
    cutoff = datetime.now(timezone.utc) - timedelta(hours=hours)

    stmt = (
        select(ProcessedEmail)
        .where(
            ProcessedEmail.user_email == user.email,
            ProcessedEmail.status_updated_at < cutoff,
            ProcessedEmail.email_status.notin_([
                "po_closed", "po_cancelled",
                "enq_converted", "enq_lost",
                "inv_paid", "inv_disputed",
                "ship_delivered",
                "gen_read"
            ])
        )
    )
    result = await db.execute(stmt)
    stale_emails = result.scalars().all()

    return {
        "success": True,
        "stale_count": len(stale_emails),
        "stale": [EmailRecord.model_validate(e) for e in stale_emails]
    }
