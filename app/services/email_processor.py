"""
Email processing orchestration service.
Contains the core logic extracted from the original scheduler.py's EmailAutomation class.
Separated from the scheduler so it can be called both by the background job
AND by the API (on-demand classification).
"""
import json
import logging
from datetime import datetime, timedelta, timezone
from typing import Dict, Optional

from sqlalchemy import func, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import settings
from app.models.domain import ClassificationStats, ProcessedEmail
from app.services.classifier import get_classifier
from app.services.graph_client import graph_client

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Core classification + persistence
# ---------------------------------------------------------------------------


async def classify_and_save(
    db: AsyncSession,
    message_id: str,
    subject: str,
    body: str,
    sender: str,
    user_email: str,
    received_at: Optional[datetime] = None,
    body_preview: str = "",
    conversation_id: Optional[str] = None,
    source_folder: Optional[str] = None,
) -> Dict:
    """
    Classify an email with the AI and persist the result to the database.
    Returns the classification result dict.
    """
    classifier = get_classifier()
    result = await classifier.classify(subject, body)

    category = result["category"]
    confidence = result["confidence"]
    reason = result["reason"]
    outlook_cat = settings.CATEGORIES.get(category, {}).get("outlook_category", "General")

    # Map category to initial status ID
    initial_status_map = {
        "purchase_order": "po_received",
        "enquiry":        "enq_new",
        "invoice":        "inv_received",
        "shipping":       "ship_dispatched",
        "general":        "gen_new",
        "junk":           "junk_new",
    }
    initial_status = initial_status_map.get(category, "gen_new")

    # Estimate priority
    from src.priority import compute_priority_score
    p_info = compute_priority_score({
        "received_at": received_at or datetime.now(timezone.utc),
        "estimated_value": result.get("estimated_value", 0.0),
        "subject": subject,
        "body_preview": body_preview[:1000] if body_preview else (body[:200] if body else ""),
        "sender": sender,
        "category": category,
    })

    email_record = ProcessedEmail(
        message_id=message_id,
        user_email=user_email,
        sender=sender,
        subject=subject[:500] if subject else None,
        body_preview=body_preview[:1000] if body_preview else (body[:200] if body else None),
        category=category,
        confidence=confidence,
        reason=reason,
        outlook_category_applied=outlook_cat,
        received_at=received_at or datetime.now(timezone.utc),
        extracted_data=result.get("extracted_data"),
        response_draft=result.get("response_draft"),
        # v2.1 fields
        conversation_id=conversation_id or result.get("conversation_id") or "",
        email_status=initial_status,
        status_updated_at=datetime.now(timezone.utc),
        estimated_value=result.get("estimated_value", 0.0),
        source_folder=source_folder or "Inbox",
        priority_score=p_info["score"],
        priority_tier=p_info["tier"],
        reply_sent=False,
    )

    db.add(email_record)
    await db.flush()  # Get the ID without committing (caller commits via dependency)

    # Update daily stats
    await _update_stats(db, user_email, category, confidence)

    logger.info(
        "Email classified and saved",
        extra={
            "message_id": message_id,
            "category": category,
            "confidence": confidence,
            "user_email": user_email,
        },
    )

    result["message_id"] = message_id
    result["processed_at"] = datetime.now(timezone.utc).isoformat()
    result["outlook_category"] = outlook_cat
    return result


async def get_or_classify(
    db: AsyncSession,
    message_id: str,
    subject: str,
    body: str,
    sender: str,
    user_email: str,
    conversation_id: Optional[str] = None,
    source_folder: Optional[str] = None,
) -> tuple[Dict, bool]:
    """
    Check the DB cache first; only call the AI if the email hasn't been classified yet.
    Returns (result_dict, from_cache).
    """
    stmt = select(ProcessedEmail).where(ProcessedEmail.message_id == message_id)
    result = await db.execute(stmt)
    existing = result.scalar_one_or_none()

    classifier = get_classifier()
    if existing:
        is_fallback = (
            existing.reason.startswith("Keyword match:") or
            existing.reason == "No matching keywords found"
        )
        has_new_body = (
            len(body.strip()) > 10 and 
            (not existing.body_preview or len(existing.body_preview.strip()) <= 5)
        )
        if (is_fallback or has_new_body) and classifier._client is not None:
            logger.info("Upgrading cached email classification with new/AI content", extra={"message_id": message_id})
            await db.delete(existing)
            await db.flush()
            existing = None

    if existing:
        logger.debug("Email found in cache", extra={"message_id": message_id})
        return {
            "message_id": existing.message_id,
            "category": existing.category,
            "confidence": existing.confidence,
            "reason": existing.reason,
            "extracted_data": existing.extracted_data,
            "response_draft": existing.response_draft,
            "processed_at": existing.processed_at.isoformat() if existing.processed_at else None,
        }, True

    data = await classify_and_save(
        db=db,
        message_id=message_id,
        subject=subject,
        body=body,
        sender=sender,
        user_email=user_email,
        conversation_id=conversation_id,
        source_folder=source_folder,
    )
    return data, False


# ---------------------------------------------------------------------------
# Batch processing (called by the scheduler)
# ---------------------------------------------------------------------------


async def run_email_batch(db: AsyncSession) -> Dict:
    """
    Fetch, classify, and tag unread emails across all organisation users.
    Called by the background scheduler every CHECK_INTERVAL_MINUTES minutes.
    Returns batch statistics dict.
    """
    stats: Dict = {"processed": 0, "failed": 0, "skipped": 0, "by_category": {}}

    try:
        users = await graph_client.get_users()
        logger.info("Starting email batch", extra={"user_count": len(users)})

        since = datetime.now(timezone.utc) - timedelta(minutes=settings.LOOKBACK_MINUTES)

        for user in users:
            user_email = user.get("mail") or user.get("userPrincipalName", "")
            user_id = user["id"]

            emails = await graph_client.get_unread_emails(user_id, since)
            if not emails:
                continue

            logger.info(
                "Processing user emails",
                extra={"user_email": user_email, "email_count": len(emails)},
            )

            # Ensure Outlook categories exist for this user
            for cat_id, cat_info in settings.CATEGORIES.items():
                if cat_id == "general":
                    continue
                try:
                    await graph_client.create_outlook_category(
                        user_id, cat_info["outlook_category"], cat_info["color"]
                    )
                except Exception as exc:
                    logger.warning(
                        "Could not create Outlook category",
                        extra={"category": cat_id, "error": str(exc)},
                    )

            for email in emails:
                msg_id = email["id"]

                # Skip already processed
                check = await db.execute(
                    select(ProcessedEmail).where(ProcessedEmail.message_id == msg_id)
                )
                if check.scalar_one_or_none():
                    stats["skipped"] += 1
                    continue

                try:
                    subject = email.get("subject", "(No subject)")
                    body = email.get("body", {}).get("content", "")
                    body_preview = email.get("bodyPreview", "")[:200]
                    sender = (
                        email.get("from", {}).get("emailAddress", {}).get("address", "unknown")
                    )
                    received_at = _parse_datetime(email.get("receivedDateTime"))

                    result = await classify_and_save(
                        db=db,
                        message_id=msg_id,
                        subject=subject,
                        body=body,
                        sender=sender,
                        user_email=user_email,
                        received_at=received_at,
                        body_preview=body_preview,
                        conversation_id=email.get("conversationId"),
                        source_folder="Inbox",
                    )

                    category = result["category"]
                    outlook_cat = settings.CATEGORIES[category]["outlook_category"]
                    await graph_client.apply_category(user_id, msg_id, outlook_cat)

                    stats["processed"] += 1
                    stats["by_category"][category] = stats["by_category"].get(category, 0) + 1

                except Exception as exc:
                    stats["failed"] += 1
                    logger.error(
                        "Failed to process individual email",
                        extra={"message_id": msg_id, "error": str(exc)},
                        exc_info=True,
                    )

        # Commit all DB writes from this batch
        await db.commit()

        logger.info("Email batch complete", extra=stats)
        return stats

    except Exception as exc:
        await db.rollback()
        logger.error("Critical error in email batch", extra={"error": str(exc)}, exc_info=True)
        raise


# ---------------------------------------------------------------------------
# Stats helpers (shared between API routes and scheduler)
# ---------------------------------------------------------------------------


async def update_email_category(
    db: AsyncSession, message_id: str, new_category: str
) -> bool:
    """Update the stored category for an already-classified email."""
    stmt = select(ProcessedEmail).where(ProcessedEmail.message_id == message_id)
    result = await db.execute(stmt)
    email = result.scalar_one_or_none()
    if not email:
        return False
    email.category = new_category
    return True


async def add_feedback(
    db: AsyncSession, message_id: str, is_correct: bool, note: Optional[str] = None
) -> bool:
    """Record user feedback on a classification."""
    stmt = select(ProcessedEmail).where(ProcessedEmail.message_id == message_id)
    result = await db.execute(stmt)
    email = result.scalar_one_or_none()
    if not email:
        return False
    email.is_correct = is_correct
    email.feedback_note = note
    return True


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


async def _update_stats(
    db: AsyncSession, user_email: str, category: str, confidence: float
) -> None:
    """Update the daily classification stats rollup."""
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    stmt = select(ClassificationStats).where(
        ClassificationStats.date == today,
        ClassificationStats.user_email == user_email,
        ClassificationStats.category == category,
    )
    result = await db.execute(stmt)
    stat = result.scalar_one_or_none()

    if stat:
        stat.count += 1
        stat.avg_confidence = (
            (stat.avg_confidence * (stat.count - 1) + confidence) / stat.count
            if stat.avg_confidence
            else confidence
        )
    else:
        stat = ClassificationStats(
            date=today,
            user_email=user_email,
            category=category,
            count=1,
            avg_confidence=confidence,
        )
        db.add(stat)


def _parse_datetime(dt_str: Optional[str]) -> datetime:
    """Parse an ISO datetime string safely."""
    if not dt_str:
        return datetime.now(timezone.utc)
    try:
        return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except (ValueError, AttributeError):
        return datetime.now(timezone.utc)


async def get_thread_lifecycle_status(db: AsyncSession, conversation_id: str) -> Optional[Dict]:
    if not conversation_id:
        return None

    # Query all emails in the same conversation thread sorted by received_at ascending
    stmt = (
        select(ProcessedEmail)
        .where(ProcessedEmail.conversation_id == conversation_id)
        .order_by(ProcessedEmail.received_at.asc())
    )
    result = await db.execute(stmt)
    thread_emails = result.scalars().all()

    if not thread_emails:
        return None

    # Map categories present in the thread
    has_enquiry = any(e.category == "enquiry" for e in thread_emails)
    has_order = any(e.category == "purchase_order" for e in thread_emails)
    has_invoice = any(e.category == "invoice" for e in thread_emails)
    has_shipping = any(e.category == "shipping" for e in thread_emails)

    # Determine current lifecycle stage (highest stage reached)
    current_stage = "general"
    if has_shipping:
        current_stage = "shipping"
    elif has_invoice:
        current_stage = "invoice"
    elif has_order:
        current_stage = "purchase_order"
    elif has_enquiry:
        current_stage = "enquiry"

    # Build milestones representing automatic progressions/transitions
    milestones = []
    seen_categories = set()
    
    # Chronological milestones
    for e in thread_emails:
        if e.category in ["enquiry", "purchase_order", "invoice", "shipping"]:
            if e.category not in seen_categories:
                seen_categories.add(e.category)
                
                # Create a readable progress label
                label = "Enquiry Received"
                if e.category == "purchase_order":
                    label = "Enquiry progressed to Order" if "enquiry" in seen_categories else "Order Placed"
                elif e.category == "invoice":
                    label = "Order progressed to Invoice" if "purchase_order" in seen_categories else "Invoice Received"
                elif e.category == "shipping":
                    label = "Invoice progressed to Shipment" if "invoice" in seen_categories else "Shipped"

                milestones.append({
                    "label": label,
                    "timestamp": e.received_at,
                    "category": e.category
                })

    return {
        "conversation_id": conversation_id,
        "has_enquiry": has_enquiry,
        "has_order": has_order,
        "has_invoice": has_invoice,
        "has_shipping": has_shipping,
        "current_stage": current_stage,
        "milestones": milestones,
    }
