"""
Provident Packaging — Priority Scoring Engine v2.1
Computes a 0–100 priority score for each processed email based on:
  - Time elapsed since received (40%)
  - Estimated dollar value (35%)
  - Urgency keywords in subject/body (15%)
  - Customer tier — returning vs first-time sender (10%)
"""
import re
from datetime import datetime, timezone
from typing import Dict, List, Optional


# ─── Urgency keyword sets ─────────────────────────────────────────────────────
HIGH_URGENCY_KEYWORDS = [
    "urgent", "urgently", "asap", "as soon as possible", "immediately",
    "critical", "deadline today", "by today", "by end of day", "eod",
    "time sensitive", "priority", "rush", "rush order", "expedite",
]
MEDIUM_URGENCY_KEYWORDS = [
    "soon", "this week", "by friday", "by monday", "waiting",
    "follow up", "follow-up", "reminder", "pending", "overdue",
]


# ─── Score computation ────────────────────────────────────────────────────────

def _time_score(received_at: Optional[datetime]) -> float:
    """
    Score based on how old the email is (0–40 points).
    Newer emails that are still unactioned are scored based on elapsed hours:
      0–6h   → 10 pts (still fresh, low urgency from time alone)
      6–24h  → 25 pts (business day elapsed, needs action)
      24–48h → 35 pts (overdue — one business day without response)
      48h+   → 40 pts (critical — multiple days without action)
    """
    if not received_at:
        return 20.0

    if received_at.tzinfo is None:
        received_at = received_at.replace(tzinfo=timezone.utc)
    now = datetime.now(timezone.utc)
    hours_elapsed = max(0, (now - received_at).total_seconds() / 3600)

    if hours_elapsed < 6:
        return 10.0
    elif hours_elapsed < 24:
        return 25.0
    elif hours_elapsed < 48:
        return 35.0
    else:
        return 40.0


def _value_score(estimated_value: Optional[float]) -> float:
    """
    Score based on estimated dollar value (0–35 points).
    Tiered for a typical packaging company's deal sizes:
      $0           → 0 pts
      $1–$999      → 8 pts  (small / sample run)
      $1k–$4,999   → 18 pts (mid-range order)
      $5k–$14,999  → 27 pts (significant order)
      $15k+        → 35 pts (major order — top priority)
    """
    if not estimated_value or estimated_value <= 0:
        return 0.0
    if estimated_value < 1000:
        return 8.0
    elif estimated_value < 5000:
        return 18.0
    elif estimated_value < 15000:
        return 27.0
    else:
        return 35.0


def _keyword_score(subject: str, body_preview: str) -> float:
    """
    Score based on urgency keyword detection (0–15 points).
    High-urgency keywords found → 15 pts
    Medium-urgency keywords → 8 pts
    No keywords → 0 pts
    """
    text = ((subject or "") + " " + (body_preview or "")).lower()
    for kw in HIGH_URGENCY_KEYWORDS:
        if kw in text:
            return 15.0
    for kw in MEDIUM_URGENCY_KEYWORDS:
        if kw in text:
            return 8.0
    return 0.0


def _customer_tier_score(sender: str, known_senders: set) -> float:
    """
    Score based on whether the sender is a returning customer (0–10 points).
    Returning senders → 10 pts (trusted, high-value relationship)
    First-time senders → 5 pts (still matters, but unknown priority)
    """
    if not sender:
        return 5.0
    domain = sender.split("@")[-1].lower() if "@" in sender else sender.lower()
    return 10.0 if domain in known_senders else 5.0


def _score_to_tier(score: float) -> str:
    """Convert numeric score to human-readable priority tier."""
    if score >= 80:
        return "critical"
    elif score >= 60:
        return "high"
    elif score >= 40:
        return "medium"
    else:
        return "low"


# ─── Public API ───────────────────────────────────────────────────────────────

def compute_priority_score(
    email: Dict,
    known_sender_domains: Optional[set] = None
) -> Dict:
    """
    Compute a full priority score for a single email record dict.
    
    Args:
        email: Dict with fields: received_at (ISO str or datetime), estimated_value,
               subject, body_preview, sender, category.
        known_sender_domains: Set of email domains seen in previous records.
                              If None, customer tier defaults to 5.
    
    Returns:
        {score: float, tier: str, reasons: list[str]}
    """
    if known_sender_domains is None:
        known_sender_domains = set()

    # Parse received_at
    received_at = email.get("received_at")
    if isinstance(received_at, str):
        try:
            received_at = datetime.fromisoformat(received_at.replace("Z", "+00:00"))
        except ValueError:
            received_at = None

    subject = email.get("subject", "")
    body_preview = email.get("body_preview", "")
    sender = email.get("sender", "")
    estimated_value = email.get("estimated_value") or 0.0

    # Component scores
    t_score  = _time_score(received_at)
    v_score  = _value_score(estimated_value)
    k_score  = _keyword_score(subject, body_preview)
    c_score  = _customer_tier_score(sender, known_sender_domains)

    total = min(100.0, t_score + v_score + k_score + c_score)
    tier = _score_to_tier(total)

    # Build human-readable reason list for the UI
    reasons: List[str] = []
    if received_at:
        hours = max(0, (datetime.now(timezone.utc) - received_at.replace(tzinfo=received_at.tzinfo or timezone.utc)).total_seconds() / 3600)
        if hours >= 48:
            reasons.append(f"{int(hours)}h without response")
        elif hours >= 24:
            reasons.append("No response in 24h")
    if estimated_value and estimated_value > 0:
        reasons.append(f"${estimated_value:,.0f} estimated value")
    if k_score >= 15:
        reasons.append("Urgency keyword detected")
    elif k_score >= 8:
        reasons.append("Time-sensitive language")
    if c_score >= 10:
        reasons.append("Returning customer")

    return {"score": round(total, 1), "tier": tier, "reasons": reasons}


def compute_batch_priorities(emails: List[Dict]) -> List[Dict]:
    """
    Compute priority scores for a list of email dicts.
    Extracts the known sender domain set from the batch to weight returning customers.
    Returns the same list with 'priority_score' and 'priority_tier' fields added.
    """
    # Build known domain set from all senders in the batch
    known_domains = set()
    from collections import Counter
    domain_counter = Counter()
    for e in emails:
        sender = e.get("sender", "")
        if "@" in sender:
            domain_counter[sender.split("@")[-1].lower()] += 1
    # A domain seen more than once is a "known/returning" sender
    known_domains = {d for d, cnt in domain_counter.items() if cnt > 1}

    enriched = []
    for email in emails:
        result = compute_priority_score(email, known_domains)
        email["priority_score"] = result["score"]
        email["priority_tier"]  = result["tier"]
        email["priority_reasons"] = result["reasons"]
        enriched.append(email)

    # Sort by score descending
    enriched.sort(key=lambda e: e.get("priority_score", 0), reverse=True)
    return enriched
