"""
Provident Packaging Email Classifier - Database Layer
Tracks processed emails, classifications, lifecycle status, and priority scoring
"""
import json
from datetime import datetime
from typing import Optional, List, Dict
from sqlalchemy import create_engine, Column, String, Integer, DateTime, Float, Boolean, Text
from sqlalchemy.orm import declarative_base, sessionmaker
from config import config

Base = declarative_base()

# ─────────────────────────────────────────────────────────────────────────────
# Lifecycle status definitions per category
# ─────────────────────────────────────────────────────────────────────────────
STATUS_FLOWS = {
    "purchase_order": [
        {"id": "po_received",     "label": "Received",        "color": "#f59e0b"},
        {"id": "po_acknowledged", "label": "Acknowledged",    "color": "#3b82f6"},
        {"id": "po_in_production","label": "In Production",   "color": "#8b5cf6"},
        {"id": "po_ready",        "label": "Ready to Ship",   "color": "#f97316"},
        {"id": "po_shipped",      "label": "Shipped",         "color": "#10b981"},
        {"id": "po_closed",       "label": "Delivered",       "color": "#059669"},
        {"id": "po_cancelled",    "label": "Cancelled",       "color": "#ef4444"},
    ],
    "enquiry": [
        {"id": "enq_new",       "label": "New Enquiry",        "color": "#f59e0b"},
        {"id": "enq_quoting",   "label": "Quote in Progress",  "color": "#3b82f6"},
        {"id": "enq_quoted",    "label": "Quote Sent",         "color": "#8b5cf6"},
        {"id": "enq_followup",  "label": "Follow-up Needed",   "color": "#f97316"},
        {"id": "enq_converted", "label": "Converted to Order", "color": "#10b981"},
        {"id": "enq_lost",      "label": "Lost",               "color": "#ef4444"},
    ],
    "invoice": [
        {"id": "inv_received", "label": "Invoice Received",    "color": "#f59e0b"},
        {"id": "inv_review",   "label": "Under Review",        "color": "#3b82f6"},
        {"id": "inv_approved", "label": "Approved",            "color": "#8b5cf6"},
        {"id": "inv_paid",     "label": "Payment Sent",        "color": "#10b981"},
        {"id": "inv_disputed", "label": "Disputed",            "color": "#ef4444"},
    ],
    "shipping": [
        {"id": "ship_dispatched", "label": "Dispatched",         "color": "#f59e0b"},
        {"id": "ship_transit",    "label": "In Transit",         "color": "#3b82f6"},
        {"id": "ship_out",        "label": "Out for Delivery",   "color": "#8b5cf6"},
        {"id": "ship_delivered",  "label": "Delivered",          "color": "#10b981"},
        {"id": "ship_delayed",    "label": "Delayed",            "color": "#ef4444"},
    ],
    "general": [
        {"id": "gen_new",  "label": "New",         "color": "#6b7280"},
        {"id": "gen_read", "label": "Acknowledged","color": "#10b981"},
    ],
    "junk": [
        {"id": "junk_new",      "label": "New Junk",       "color": "#6b7280"},
        {"id": "junk_review",   "label": "Under Review",   "color": "#3b82f6"},
        {"id": "junk_flagged",  "label": "Flagged",        "color": "#eab308"},
        {"id": "junk_archived", "label": "Archived",       "color": "#10b981"},
        {"id": "junk_deleted",  "label": "Deleted",        "color": "#ef4444"},
    ],
}

# Default starting status for each category
INITIAL_STATUS = {
    "purchase_order": "po_received",
    "enquiry":        "enq_new",
    "invoice":        "inv_received",
    "shipping":       "ship_dispatched",
    "general":        "gen_new",
    "junk":           "junk_new",
}

# Next status in the normal (non-terminal) flow
NEXT_STATUS = {
    "po_received":     "po_acknowledged",
    "po_acknowledged": "po_in_production",
    "po_in_production":"po_ready",
    "po_ready":        "po_shipped",
    "po_shipped":      "po_closed",
    # terminal: po_closed, po_cancelled

    "enq_new":       "enq_quoting",
    "enq_quoting":   "enq_quoted",
    "enq_quoted":    "enq_followup",
    "enq_followup":  "enq_converted",
    # terminal: enq_converted, enq_lost

    "inv_received":  "inv_review",
    "inv_review":    "inv_approved",
    "inv_approved":  "inv_paid",
    # terminal: inv_paid, inv_disputed

    "ship_dispatched": "ship_transit",
    "ship_transit":    "ship_out",
    "ship_out":        "ship_delivered",
    # terminal: ship_delivered, ship_delayed

    "gen_new":  "gen_read",
    # terminal: gen_read

    "junk_new":     "junk_review",
    "junk_review":   "junk_flagged",
    "junk_flagged":  "junk_archived",
    "junk_archived": "junk_deleted",
    # terminal: junk_deleted
}

# Status IDs that indicate a reply has been sent / work done on the thread
REPLY_ADVANCE_STATUS = {
    "purchase_order": "po_acknowledged",
    "enquiry":        "enq_quoted",
    "invoice":        "inv_review",
    "junk":           "junk_review",
    "shipping":       "ship_delivered",
    "general":        "gen_read",
}


class ProcessedEmail(Base):
    """Record of an email that was classified"""
    __tablename__ = "processed_emails"

    id = Column(Integer, primary_key=True)
    message_id = Column(String(255), unique=True, nullable=False, index=True)
    conversation_id = Column(String(255), index=True)        # MS Graph conversationId
    user_email = Column(String(255), nullable=False)
    sender = Column(String(255))
    subject = Column(Text)
    body_preview = Column(Text)
    category = Column(String(50), nullable=False)
    confidence = Column(Float)
    reason = Column(Text)
    outlook_category_applied = Column(String(100))
    is_correct = Column(Boolean, default=None)
    feedback_note = Column(Text)
    processed_at = Column(DateTime, default=datetime.utcnow)
    received_at = Column(DateTime)
    extracted_data = Column(Text)
    response_draft = Column(Text)

    # v2.1 fields
    email_status = Column(String(50), default=None)           # e.g. "po_received"
    status_updated_at = Column(DateTime, default=None)
    estimated_value = Column(Float, default=None)             # $ extracted from email
    source_folder = Column(String(100), default="Inbox")      # "Inbox","Junk","Archive"...
    priority_score = Column(Float, default=None)              # 0–100
    priority_tier = Column(String(10), default=None)          # "critical","high","medium","low"
    reply_sent = Column(Boolean, default=False)
    sent_reply = Column(Text, default=None)
    reply_sent_at = Column(DateTime, default=None)


class ClassificationStats(Base):
    """Daily statistics for reporting"""
    __tablename__ = "classification_stats"

    id = Column(Integer, primary_key=True)
    date = Column(String(10), nullable=False, index=True)
    user_email = Column(String(255), nullable=False)
    category = Column(String(50), nullable=False)
    count = Column(Integer, default=0)
    avg_confidence = Column(Float)
    incorrect_count = Column(Integer, default=0)


class Database:
    """Database wrapper with v2.1 lifecycle + priority methods"""

    def __init__(self):
        if config.DATABASE_URL.startswith("sqlite:///"):
            db_path = config.DATABASE_URL.replace("sqlite:///", "")
            import os
            db_dir = os.path.dirname(db_path)
            if db_dir and not os.path.exists(db_dir):
                os.makedirs(db_dir, exist_ok=True)

        self.engine = create_engine(config.DATABASE_URL, echo=False)
        # Add new columns to existing DBs via ALTER TABLE (safe migration shim)
        self._migrate_schema()
        Base.metadata.create_all(self.engine)
        self.Session = sessionmaker(bind=self.engine)

    def _migrate_schema(self):
        """Safely add new columns to existing SQLite databases without dropping data."""
        new_cols = [
            ("conversation_id",    "VARCHAR(255)"),
            ("email_status",       "VARCHAR(50)"),
            ("status_updated_at",  "DATETIME"),
            ("estimated_value",    "FLOAT"),
            ("source_folder",      "VARCHAR(100) DEFAULT 'Inbox'"),
            ("priority_score",     "FLOAT"),
            ("priority_tier",      "VARCHAR(10)"),
            ("reply_sent",         "BOOLEAN DEFAULT 0"),
            ("sent_reply",         "TEXT"),
            ("reply_sent_at",      "DATETIME"),
        ]
        with self.engine.connect() as conn:
            for col_name, col_type in new_cols:
                try:
                    conn.execute(
                        __import__("sqlalchemy").text(
                            f"ALTER TABLE processed_emails ADD COLUMN {col_name} {col_type}"
                        )
                    )
                    conn.commit()
                except Exception:
                    pass  # Column already exists — safe to ignore

    # ─── Core write operations ────────────────────────────────────────────────

    def is_processed(self, message_id: str) -> bool:
        session = self.Session()
        try:
            return session.query(ProcessedEmail).filter_by(message_id=message_id).first() is not None
        finally:
            session.close()

    def record_classification(self,
                               message_id: str,
                               user_email: str,
                               sender: str,
                               subject: str,
                               body_preview: str,
                               category: str,
                               confidence: float,
                               reason: str,
                               outlook_category: str,
                               received_at: datetime,
                               extracted_data: str = None,
                               response_draft: str = None,
                               conversation_id: str = None,
                               estimated_value: float = None,
                               source_folder: str = "Inbox") -> None:
        session = self.Session()
        try:
            initial_status = INITIAL_STATUS.get(category, "gen_new")
            email = ProcessedEmail(
                message_id=message_id,
                conversation_id=conversation_id,
                user_email=user_email,
                sender=sender,
                subject=subject[:500] if subject else None,
                body_preview=body_preview[:1000] if body_preview else None,
                category=category,
                confidence=confidence,
                reason=reason,
                outlook_category_applied=outlook_category,
                received_at=received_at,
                extracted_data=extracted_data,
                response_draft=response_draft,
                email_status=initial_status,
                status_updated_at=datetime.utcnow(),
                estimated_value=estimated_value,
                source_folder=source_folder,
            )
            session.add(email)
            session.commit()
            self._update_stats(session, user_email, category, confidence)
        finally:
            session.close()

    def _update_stats(self, session, user_email: str, category: str, confidence: float):
        today = datetime.utcnow().strftime("%Y-%m-%d")
        stat = session.query(ClassificationStats).filter_by(
            date=today, user_email=user_email, category=category
        ).first()
        if stat:
            stat.count += 1
            stat.avg_confidence = (stat.avg_confidence * (stat.count - 1) + confidence) / stat.count
        else:
            stat = ClassificationStats(date=today, user_email=user_email,
                                       category=category, count=1, avg_confidence=confidence)
            session.add(stat)
        session.commit()

    # ─── Status management ────────────────────────────────────────────────────

    def update_status(self, message_id: str, new_status: str) -> bool:
        """Advance the lifecycle status of an email."""
        session = self.Session()
        try:
            email = session.query(ProcessedEmail).filter_by(message_id=message_id).first()
            if email:
                email.email_status = new_status
                email.status_updated_at = datetime.utcnow()
                session.commit()
                return True
            return False
        finally:
            session.close()

    def auto_advance_status_on_reply(self, conversation_id: str, category: str) -> bool:
        """
        Called when a sent-item reply is detected in a conversation thread.
        Advances all emails in that conversation to the reply-triggered status.
        """
        session = self.Session()
        try:
            emails = session.query(ProcessedEmail).filter_by(
                conversation_id=conversation_id, category=category
            ).all()
            if not emails:
                return False
            target_status = REPLY_ADVANCE_STATUS.get(category)
            if not target_status:
                return False
            for email in emails:
                # Only advance if the email hasn't passed this status yet
                current = email.email_status or ""
                flow = STATUS_FLOWS.get(category, [])
                ids = [s["id"] for s in flow]
                current_idx = ids.index(current) if current in ids else -1
                target_idx = ids.index(target_status) if target_status in ids else -1
                if current_idx < target_idx:
                    email.email_status = target_status
                    email.status_updated_at = datetime.utcnow()
            session.commit()
            return True
        finally:
            session.close()

    def get_stale_emails(self, hours: int = 24) -> List[Dict]:
        """Return emails where status hasn't changed in N hours (for attention banner)."""
        session = self.Session()
        try:
            from sqlalchemy import or_
            cutoff = datetime.utcnow() - __import__("datetime").timedelta(hours=hours)
            emails = session.query(ProcessedEmail).filter(
                ProcessedEmail.status_updated_at < cutoff,
                ProcessedEmail.email_status.notin_([
                    "po_closed", "po_cancelled",
                    "enq_converted", "enq_lost",
                    "inv_paid", "inv_disputed",
                    "ship_delivered",
                    "gen_read"
                ])
            ).all()
            return [self._serialize(e) for e in emails]
        finally:
            session.close()

    # ─── Priority ─────────────────────────────────────────────────────────────

    def update_priority(self, message_id: str, score: float, tier: str) -> None:
        session = self.Session()
        try:
            email = session.query(ProcessedEmail).filter_by(message_id=message_id).first()
            if email:
                email.priority_score = score
                email.priority_tier = tier
                session.commit()
        finally:
            session.close()

    # ─── Read operations ──────────────────────────────────────────────────────

    def add_feedback(self, message_id: str, is_correct: bool, note: str = None) -> bool:
        session = self.Session()
        try:
            email = session.query(ProcessedEmail).filter_by(message_id=message_id).first()
            if email:
                email.is_correct = is_correct
                email.feedback_note = note
                session.commit()
                return True
            return False
        finally:
            session.close()

    def get_daily_stats(self, days: int = 7) -> List[Dict]:
        session = self.Session()
        try:
            from sqlalchemy import func
            results = session.query(
                ClassificationStats.date,
                ClassificationStats.category,
                func.sum(ClassificationStats.count).label("total"),
                func.avg(ClassificationStats.avg_confidence).label("avg_conf")
            ).group_by(ClassificationStats.date, ClassificationStats.category).order_by(
                ClassificationStats.date.desc()
            ).limit(days * 5).all()
            return [{"date": r.date, "category": r.category,
                     "count": r.total, "avg_confidence": round(r.avg_conf, 1) if r.avg_conf else 0}
                    for r in results]
        finally:
            session.close()

    def get_accuracy(self, days: int = 7) -> Dict:
        session = self.Session()
        try:
            from sqlalchemy import func
            total = session.query(ProcessedEmail).filter(
                ProcessedEmail.processed_at >= datetime.utcnow().replace(
                    hour=0, minute=0, second=0, microsecond=0)
            ).count()
            reviewed = session.query(ProcessedEmail).filter(ProcessedEmail.is_correct.isnot(None)).count()
            correct = session.query(ProcessedEmail).filter(ProcessedEmail.is_correct == True).count()
            accuracy = (correct / reviewed * 100) if reviewed > 0 else 0
            return {"total_processed": total, "reviewed": reviewed,
                    "correct": correct, "incorrect": reviewed - correct,
                    "accuracy_percent": round(accuracy, 1)}
        finally:
            session.close()

    def get_recent_emails(self, limit: int = 50) -> List[Dict]:
        session = self.Session()
        try:
            from sqlalchemy import func
            valid_folders = ['inbox', 'archive', 'junk', 'junk email', 'junkemail', 'sent items']
            folder_filter = func.lower(func.coalesce(ProcessedEmail.source_folder, 'inbox')).in_(valid_folders)
            emails = session.query(ProcessedEmail).filter(folder_filter).order_by(
                ProcessedEmail.processed_at.desc()
            ).limit(limit).all()
            return [self._serialize(e) for e in emails]
        finally:
            session.close()

    def get_emails_by_category(self, category: str, limit: int = 100) -> List[Dict]:
        """Get emails for a category, sorted by priority score descending."""
        session = self.Session()
        try:
            from sqlalchemy import desc, case, func
            valid_folders = ['inbox', 'archive', 'junk', 'junk email', 'junkemail', 'sent items']
            folder_filter = func.lower(func.coalesce(ProcessedEmail.source_folder, 'inbox')).in_(valid_folders)
            emails = session.query(ProcessedEmail).filter(folder_filter).filter_by(category=category).order_by(
                desc(case(
                    (ProcessedEmail.priority_score.isnot(None), ProcessedEmail.priority_score),
                    else_=0
                )),
                ProcessedEmail.received_at.desc()
            ).limit(limit).all()
            return [self._serialize(e) for e in emails]
        finally:
            session.close()

    def get_emails_by_status(self, category: str, status: str) -> List[Dict]:
        session = self.Session()
        try:
            from sqlalchemy import func
            valid_folders = ['inbox', 'archive', 'junk', 'junk email', 'junkemail', 'sent items']
            folder_filter = func.lower(func.coalesce(ProcessedEmail.source_folder, 'inbox')).in_(valid_folders)
            emails = session.query(ProcessedEmail).filter(folder_filter).filter_by(
                category=category, email_status=status
            ).order_by(ProcessedEmail.received_at.desc()).all()
            return [self._serialize(e) for e in emails]
        finally:
            session.close()

    def update_category(self, message_id: str, category: str) -> bool:
        session = self.Session()
        try:
            email = session.query(ProcessedEmail).filter_by(message_id=message_id).first()
            if email:
                email.category = category
                email.email_status = INITIAL_STATUS.get(category, "gen_new")
                email.status_updated_at = datetime.utcnow()
                session.commit()
                return True
            return False
        finally:
            session.close()

    def _serialize(self, e: ProcessedEmail) -> Dict:
        """Convert ORM record to dict for API responses."""
        return {
            "message_id": e.message_id,
            "conversation_id": e.conversation_id,
            "user_email": e.user_email,
            "sender": e.sender,
            "subject": e.subject,
            "body_preview": e.body_preview,
            "category": e.category,
            "confidence": e.confidence,
            "reason": e.reason,
            "is_correct": e.is_correct,
            "feedback_note": e.feedback_note,
            "processed_at": e.processed_at.isoformat() if e.processed_at else None,
            "received_at": e.received_at.isoformat() if e.received_at else None,
            "extracted_data": e.extracted_data,
            "response_draft": e.response_draft,
            # v2.1
            "email_status": e.email_status,
            "status_updated_at": e.status_updated_at.isoformat() if e.status_updated_at else None,
            "estimated_value": e.estimated_value,
            "source_folder": e.source_folder or "Inbox",
            "priority_score": e.priority_score,
            "priority_tier": e.priority_tier,
            "reply_sent": e.reply_sent,
            "sent_reply": e.sent_reply,
            "reply_sent_at": e.reply_sent_at.isoformat() if e.reply_sent_at else None,
        }

    def record_reply_sent(self, message_id: str, reply_text: str) -> str:
        session = self.Session()
        try:
            email = session.query(ProcessedEmail).filter_by(message_id=message_id).first()
            if email:
                email.reply_sent = True
                email.sent_reply = reply_text
                email.reply_sent_at = datetime.utcnow()
                # Advance status to the reply-triggered status
                target_status = REPLY_ADVANCE_STATUS.get(email.category)
                if target_status:
                    email.email_status = target_status
                    email.status_updated_at = datetime.utcnow()
                session.commit()
                return email.email_status or ""
            return ""
        finally:
            session.close()

    def get_misclassified_examples(self, limit: int = 10) -> List[Dict]:
        session = self.Session()
        try:
            emails = session.query(ProcessedEmail).filter(
                ProcessedEmail.is_correct == False
            ).order_by(ProcessedEmail.processed_at.desc()).limit(limit).all()
            return [{"message_id": e.message_id, "subject": e.subject,
                     "predicted": e.category, "confidence": e.confidence,
                     "feedback": e.feedback_note} for e in emails]
        finally:
            session.close()
