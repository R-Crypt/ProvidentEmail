"""
Pydantic schemas for all request/response bodies.
Separating these from ORM models keeps the API contract stable
even when the DB schema evolves.
"""
from datetime import datetime
from typing import Any, Dict, List, Optional
from pydantic import BaseModel, EmailStr, Field, field_validator


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    expires_in: int


class CurrentUser(BaseModel):
    """Extracted from JWT — passed as a dependency to protected routes."""
    email: str
    tenant_id: str = ""
    display_name: str = ""


# ---------------------------------------------------------------------------
# Add-in API — Classify
# ---------------------------------------------------------------------------

class ClassifyRequest(BaseModel):
    message_id: str = Field(..., min_length=1, max_length=1024)
    subject: str = Field(default="", max_length=2000)
    body: str = Field(default="", max_length=50_000)
    sender: str = Field(default="", max_length=512)

    @field_validator("message_id")
    @classmethod
    def strip_message_id(cls, v: str) -> str:
        return v.strip()


class ClassifyResponse(BaseModel):
    message_id: str
    category: str
    confidence: float
    reason: str
    extracted_data: Optional[str] = None
    response_draft: Optional[str] = None
    processed_at: Optional[datetime] = None
    from_cache: bool = False


# ---------------------------------------------------------------------------
# Add-in API — Reclassify
# ---------------------------------------------------------------------------

class ReclassifyRequest(BaseModel):
    message_id: str = Field(..., min_length=1, max_length=1024)
    category: str = Field(..., min_length=1, max_length=50)

    @field_validator("category")
    @classmethod
    def validate_category(cls, v: str) -> str:
        valid = {"purchase_order", "enquiry", "invoice", "shipping", "general"}
        if v not in valid:
            raise ValueError(f"Category must be one of: {', '.join(sorted(valid))}")
        return v


class ReclassifyResponse(BaseModel):
    success: bool
    message_id: str
    new_category: str


# ---------------------------------------------------------------------------
# Add-in API — Triage Summary
# ---------------------------------------------------------------------------

class CategoryCounts(BaseModel):
    purchase_order: int = 0
    enquiry: int = 0
    invoice: int = 0
    shipping: int = 0
    general: int = 0


class TriageSummaryResponse(BaseModel):
    total: int
    counts: CategoryCounts


# ---------------------------------------------------------------------------
# Dashboard API — Email records
# ---------------------------------------------------------------------------

class EmailRecord(BaseModel):
    message_id: str
    user_email: str
    sender: Optional[str]
    subject: Optional[str]
    body_preview: Optional[str]
    category: str
    confidence: Optional[float]
    reason: Optional[str]
    is_correct: Optional[bool]
    feedback_note: Optional[str]
    processed_at: Optional[datetime]
    received_at: Optional[datetime]
    extracted_data: Optional[str]
    response_draft: Optional[str]

    model_config = {"from_attributes": True}


class EmailListResponse(BaseModel):
    emails: List[EmailRecord]
    total: int


# ---------------------------------------------------------------------------
# Dashboard API — Statistics
# ---------------------------------------------------------------------------

class DailyStatPoint(BaseModel):
    date: str
    category: str
    count: int
    avg_confidence: float


class AccuracyStats(BaseModel):
    total_processed: int
    reviewed: int
    correct: int
    incorrect: int
    accuracy_percent: float


class DashboardStatsResponse(BaseModel):
    daily_stats: List[DailyStatPoint]
    accuracy: AccuracyStats


# ---------------------------------------------------------------------------
# Dashboard API — Feedback
# ---------------------------------------------------------------------------

class FeedbackRequest(BaseModel):
    message_id: str = Field(..., min_length=1)
    is_correct: bool
    note: Optional[str] = Field(default=None, max_length=1000)


class FeedbackResponse(BaseModel):
    success: bool
    message_id: str


# ---------------------------------------------------------------------------
# Health
# ---------------------------------------------------------------------------

class HealthResponse(BaseModel):
    status: str
    version: str
    environment: str
    db_connected: bool = False
    details: Dict[str, Any] = {}
