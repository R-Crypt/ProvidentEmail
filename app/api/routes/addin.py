"""
Add-in API routes — all /api/addin/* endpoints.
These are called directly by taskpane.js running inside Outlook.
All routes require a valid Bearer token (Microsoft SSO or internal JWT).
"""
import logging
from datetime import datetime, timezone
from typing import Optional, List

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
            conversation_id=body.conversation_id,
            source_folder=body.source_folder,
        )

        lifecycle_info = None
        conv_id = body.conversation_id or result.get("conversation_id")
        if conv_id:
            from app.services.email_processor import get_thread_lifecycle_status
            lifecycle_info = await get_thread_lifecycle_status(db, conv_id)

        return ClassifyResponse(
            message_id=result["message_id"],
            category=result["category"],
            confidence=result["confidence"],
            reason=result["reason"],
            extracted_data=result.get("extracted_data"),
            response_draft=result.get("response_draft"),
            processed_at=result.get("processed_at"),
            from_cache=from_cache,
            thread_lifecycle=lifecycle_info,
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
# GET /api/addin/email_detail
# ---------------------------------------------------------------------------

@router.get(
    "/email_detail",
    response_model=ClassifyResponse,
    summary="Get full details for a processed email, including thread lifecycle",
)
async def addin_email_detail(
    message_id: str,
    db: DbSession,
    user: AuthUser,
) -> ClassifyResponse:
    """Retrieve full details of an email by message_id, including its thread lifecycle status."""
    stmt = select(ProcessedEmail).where(
        ProcessedEmail.message_id == message_id,
        ProcessedEmail.user_email == user.email,
    )
    result = await db.execute(stmt)
    existing = result.scalar_one_or_none()
    if not existing:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Email record not found.",
        )

    # Resolve thread lifecycle
    lifecycle_info = None
    if existing.conversation_id:
        from app.services.email_processor import get_thread_lifecycle_status
        lifecycle_info = await get_thread_lifecycle_status(db, existing.conversation_id)

    return ClassifyResponse(
        message_id=existing.message_id,
        category=existing.category,
        confidence=existing.confidence,
        reason=existing.reason,
        extracted_data=existing.extracted_data,
        response_draft=existing.response_draft,
        processed_at=existing.processed_at,
        from_cache=True,
        thread_lifecycle=lifecycle_info,
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

    # 1. Total processed today
    total_stmt = (
        select(func.count(ProcessedEmail.id))
        .where(
            ProcessedEmail.processed_at >= today_start,
            ProcessedEmail.user_email == user.email,
        )
    )
    total_today = await db.scalar(total_stmt) or 0

    # 2. Counts of unresolved emails grouped by category
    TERMINAL_STATUSES = {
        "po_closed", "po_cancelled",
        "enq_converted", "enq_lost",
        "inv_paid", "inv_disputed",
        "ship_delivered", "ship_delayed",
        "gen_read", "junk_deleted"
    }

    counts_stmt = (
        select(ProcessedEmail.category, func.count(ProcessedEmail.id))
        .where(
            ProcessedEmail.user_email == user.email,
            ProcessedEmail.email_status.notin_(TERMINAL_STATUSES)
        )
        .group_by(ProcessedEmail.category)
    )
    result = await db.execute(counts_stmt)
    rows = result.all()

    counts = CategoryCounts()
    for cat, count in rows:
        if hasattr(counts, cat):
            setattr(counts, cat, count)
        else:
            counts.general += count

    return TriageSummaryResponse(total=total_today, counts=counts)


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
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
) -> EmailListResponse:
    """Return the most recent emails classified under a given category for this user, optionally filtered by datetime range."""
    valid = {"purchase_order", "enquiry", "invoice", "shipping", "general", "junk"}
    if category not in valid:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=f"Invalid category. Must be one of: {', '.join(sorted(valid))}",
        )

    start_dt = None
    if start_date:
        try:
            start_dt = datetime.fromisoformat(start_date.replace("Z", "+00:00"))
        except ValueError:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Invalid start_date format. Must be ISO 8601.",
            )

    end_dt = None
    if end_date:
        try:
            end_dt = datetime.fromisoformat(end_date.replace("Z", "+00:00"))
        except ValueError:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Invalid end_date format. Must be ISO 8601.",
            )

    stmt = (
        select(ProcessedEmail)
        .where(
            ProcessedEmail.category == category,
            ProcessedEmail.user_email == user.email,
        )
    )

    if start_dt:
        stmt = stmt.where(ProcessedEmail.received_at >= start_dt)
    if end_dt:
        stmt = stmt.where(ProcessedEmail.received_at <= end_dt)

    stmt = stmt.order_by(ProcessedEmail.processed_at.desc()).limit(limit)
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


# ---------------------------------------------------------------------------
# POST /api/addin/sync
# ---------------------------------------------------------------------------

class SyncResponse(BaseModel):
    success: bool
    processed_count: int
    skipped_count: int
    failed_count: int

@router.post(
    "/sync",
    response_model=SyncResponse,
    summary="Fetch and process recent emails using user's delegated token",
)
async def addin_sync(
    db: DbSession,
    user: AuthUser,
    authorization: Optional[str] = None,
) -> SyncResponse:
    """
    Fetch recent emails from the authenticated user's mailbox using their delegated M365 token,
    classify any new emails, and persist them.
    """
    from fastapi import Header
    import httpx

    # Authorization header is injected via AuthUser dependency check, but we extract the raw token here
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing or invalid Authorization header. Expected 'Bearer <microsoft_access_token>'.",
        )

    token = authorization.removeprefix("Bearer ").strip()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    processed_count = 0
    skipped_count = 0
    failed_count = 0

    from app.services.email_processor import classify_and_save

    async with httpx.AsyncClient(timeout=30.0) as client:
        try:
            # Fetch last 20 emails
            url = "https://graph.microsoft.com/v1.0/me/messages?$top=20&$select=id,subject,body,bodyPreview,from,receivedDateTime,conversationId"
            resp = await client.get(url, headers=headers)
            
            if resp.status_code != 200:
                logger.error(f"Graph API sync fetch failed: {resp.status_code} - {resp.text}")
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail=f"Failed to fetch emails from Microsoft Graph: {resp.text[:200]}"
                )

            messages = resp.json().get("value", [])

            for msg in messages:
                msg_id = msg.get("id")
                if not msg_id:
                    continue

                # Check if already processed
                check_stmt = select(ProcessedEmail).where(ProcessedEmail.message_id == msg_id)
                res = await db.execute(check_stmt)
                exists = res.scalar_one_or_none()
                if exists:
                    skipped_count += 1
                    continue

                try:
                    subject = msg.get("subject") or "(No Subject)"
                    body_content = msg.get("body", {}).get("content") or ""
                    body_preview = msg.get("bodyPreview") or ""
                    
                    from_info = msg.get("from", {}).get("emailAddress", {})
                    sender_name = from_info.get("name", "")
                    sender_address = from_info.get("address", "")
                    
                    if sender_name and sender_address:
                        sender = f"{sender_name} <{sender_address}>"
                    else:
                        sender = sender_address or sender_name or "Unknown Sender <unknown@unknown.com>"

                    received_str = msg.get("receivedDateTime")
                    received_at = None
                    if received_str:
                        try:
                            # Parse UTC datetime from Graph (e.g. 2026-07-16T13:50:00Z)
                            received_at = datetime.fromisoformat(received_str.replace("Z", "+00:00"))
                        except Exception:
                            pass

                    # Run AI classification & save to DB
                    await classify_and_save(
                        db=db,
                        message_id=msg_id,
                        subject=subject,
                        body=body_content,
                        sender=sender,
                        user_email=user.email,
                        received_at=received_at,
                        body_preview=body_preview,
                        conversation_id=msg.get("conversationId"),
                    )
                    processed_count += 1

                except Exception as e:
                    logger.error(f"Failed to process email {msg_id} during sync: {e}", exc_info=True)
                    failed_count += 1

            await db.commit()
            return SyncResponse(
                success=True,
                processed_count=processed_count,
                skipped_count=skipped_count,
                failed_count=failed_count,
            )

        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"Addin sync error: {e}", exc_info=True)
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Sync failed: {e}",
            )
