"""
SQLAlchemy ORM models.
These are the same two tables from the original database.py,
adapted for async SQLAlchemy with proper type annotations.
"""
from datetime import datetime, timezone
from typing import Optional

from sqlalchemy import (
    Boolean,
    DateTime,
    Float,
    Integer,
    String,
    Text,
    func,
)
from sqlalchemy.orm import Mapped, mapped_column

from app.db.session import Base


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


class ProcessedEmail(Base):
    """Record of every email that has been classified by the AI."""

    __tablename__ = "processed_emails"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)

    # Unique Microsoft Graph message ID — prevents double-processing
    message_id: Mapped[str] = mapped_column(
        String(255), unique=True, nullable=False, index=True
    )

    # The Outlook mailbox user this email belongs to
    user_email: Mapped[str] = mapped_column(String(255), nullable=False, index=True)

    # Sender info
    sender: Mapped[Optional[str]] = mapped_column(String(255))
    subject: Mapped[Optional[str]] = mapped_column(Text)
    body_preview: Mapped[Optional[str]] = mapped_column(Text)

    # AI classification result
    category: Mapped[str] = mapped_column(String(50), nullable=False, index=True)
    confidence: Mapped[Optional[float]] = mapped_column(Float)
    reason: Mapped[Optional[str]] = mapped_column(Text)

    # The Outlook category name that was applied via Graph API
    outlook_category_applied: Mapped[Optional[str]] = mapped_column(String(100))

    # Feedback from user (None = not reviewed, True = correct, False = wrong)
    is_correct: Mapped[Optional[bool]] = mapped_column(Boolean, default=None)
    feedback_note: Mapped[Optional[str]] = mapped_column(Text)

    # Timestamps
    processed_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
        default=_utcnow,
    )
    received_at: Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True))

    # JSON-serialized extracted structured data and AI-generated reply draft
    extracted_data: Mapped[Optional[str]] = mapped_column(Text)
    response_draft: Mapped[Optional[str]] = mapped_column(Text)

    # v2.1 fields
    conversation_id: Mapped[Optional[str]] = mapped_column(String(255), index=True)
    email_status: Mapped[Optional[str]] = mapped_column(String(50), default=None)
    status_updated_at: Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True), default=None)
    estimated_value: Mapped[Optional[float]] = mapped_column(Float, default=None)
    source_folder: Mapped[Optional[str]] = mapped_column(String(100), default="Inbox")
    priority_score: Mapped[Optional[float]] = mapped_column(Float, default=None)
    priority_tier: Mapped[Optional[str]] = mapped_column(String(10), default=None)
    reply_sent: Mapped[bool] = mapped_column(Boolean, default=False)
    sent_reply: Mapped[Optional[str]] = mapped_column(Text, default=None)
    reply_sent_at: Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True), default=None)


class ClassificationStats(Base):
    """Daily rollup statistics per user per category for reporting."""

    __tablename__ = "classification_stats"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)

    # YYYY-MM-DD string for easy grouping
    date: Mapped[str] = mapped_column(String(10), nullable=False, index=True)
    user_email: Mapped[str] = mapped_column(String(255), nullable=False)
    category: Mapped[str] = mapped_column(String(50), nullable=False)
    count: Mapped[int] = mapped_column(Integer, default=0)
    avg_confidence: Mapped[Optional[float]] = mapped_column(Float)
    incorrect_count: Mapped[int] = mapped_column(Integer, default=0)
