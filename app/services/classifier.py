"""
AI Email Classifier — async version of the original classifier.py.
Business logic is 100% preserved; only the OpenAI client is switched to AsyncOpenAI.
"""
import json
import logging
import re
from typing import Dict, Optional

from openai import AsyncOpenAI

from app.core.config import settings

logger = logging.getLogger(__name__)


class EmailClassifier:
    """
    AI-powered email classifier using OpenAI.
    Classifies emails into business categories and extracts structured data.
    """

    def __init__(self) -> None:
        self.categories = settings.CATEGORIES
        self.model = settings.OPENAI_MODEL
        self._client: Optional[AsyncOpenAI] = None

        if settings.OPENAI_API_KEY:
            try:
                self._client = AsyncOpenAI(api_key=settings.OPENAI_API_KEY)
                logger.info("AsyncOpenAI client initialized", extra={"model": self.model})
            except Exception as exc:
                logger.warning("Failed to initialize OpenAI client", extra={"error": str(exc)})
        else:
            logger.warning("OPENAI_API_KEY not set — using keyword fallback classifier")

    async def classify(self, subject: str, body: str) -> Dict:
        """
        Classify an email. Returns dict with keys:
            category, confidence, reason, extracted_data, response_draft
        """
        clean_subject = self._clean_text(subject)
        clean_body = self._clean_text(body)[:3000]

        if not self._client:
            return self._fallback_classify(clean_subject, clean_body)

        system_prompt = self._build_system_prompt()
        user_content = f"Subject: {clean_subject}\n\nBody:\n{clean_body}"

        try:
            response = await self._client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_content},
                ],
                temperature=0,
                max_tokens=600,
                response_format={"type": "json_object"},
            )

            result = json.loads(response.choices[0].message.content)

            category = result.get("category", "general")
            if category not in self.categories:
                category = "general"

            logger.info(
                "Email classified by AI",
                extra={"category": category, "confidence": result.get("confidence")},
            )

            return {
                "category": category,
                "confidence": min(100, max(0, int(result.get("confidence", 80)))),
                "reason": result.get("reason", "AI classification"),
                "extracted_data": json.dumps(result.get("extracted_data", {})),
                "response_draft": result.get("response_draft", ""),
            }

        except Exception as exc:
            logger.warning(
                "AI classification failed, using keyword fallback",
                extra={"error": str(exc)},
            )
            return self._fallback_classify(clean_subject, clean_body)

    async def classify_batch(self, emails: list) -> list:
        """Classify a list of email dicts concurrently."""
        import asyncio
        tasks = [
            self.classify(e.get("subject", ""), e.get("body", {}).get("content", ""))
            for e in emails
        ]
        results = await asyncio.gather(*tasks, return_exceptions=True)

        output = []
        for email, result in zip(emails, results):
            if isinstance(result, Exception):
                logger.error(
                    "classify_batch: individual email failed",
                    extra={"error": str(result)},
                )
                result = self._fallback_classify(email.get("subject", ""), "")

            result["message_id"] = email.get("id")
            result["user_email"] = email.get("user_email")
            result["sender"] = (
                email.get("from", {}).get("emailAddress", {}).get("address", "")
            )
            result["received_at"] = email.get("receivedDateTime")
            result["subject"] = email.get("subject", "")
            output.append(result)

        return output

    # ------------------------------------------------------------------ #
    # Private helpers (identical to original classifier.py)
    # ------------------------------------------------------------------ #

    def _build_system_prompt(self) -> str:
        category_lines = [
            f'- "{cat_id}": {info["description"]}'
            for cat_id, info in self.categories.items()
        ]
        categories_text = "\n".join(category_lines)

        return f"""You are an expert email classifier for Provident Packaging, a packaging company.
Your job is to read incoming business emails and:
1. Classify them into exactly one of these categories:
{categories_text}

2. Extract key operational details depending on the chosen category:
   - For "purchase_order": Extract keys "Customer" (company name), "PO Number", "Quantity", "Items Ordered", "Delivery Date".
   - For "enquiry": Extract keys "Customer", "Requested Item", "Quantity", "Custom Printing (Yes/No)", "Samples Requested (Yes/No)", "Urgency" (Low/Medium/High).
   - For "invoice": Extract keys "Vendor", "Invoice Number", "Amount Due", "Due Date", "Payment Terms".
   - For "shipping": Extract keys "Carrier", "Tracking Number", "Est Delivery Date", "Current Status".
   - For "general": Extract any summary or action items.

3. Generate a professional, concise, ready-to-send draft reply to the email.
   - For "purchase_order": Acknowledge receipt, confirm order details, state we will process it.
   - For "enquiry": Thank them, state pricing is being calculated, promise a quote within 24 hours.
   - For "invoice": Acknowledge receipt, say it has been forwarded to accounting.
   - For "shipping": Thank them for the tracking information.
   - For "general": Polite acknowledgement of receipt.

Respond ONLY with a JSON object in this exact format:
{{
    "category": "category_id",
    "confidence": 85,
    "reason": "Brief explanation of why this category was chosen",
    "extracted_data": {{
        "Key": "Value"
    }},
    "response_draft": "Email reply body here..."
}}"""

    def _fallback_classify(self, subject: str, body: str) -> Dict:
        """Simple keyword-based fallback when AI is unavailable."""
        text = (subject + " " + body).lower()

        keyword_scores = {
            "purchase_order": [
                "purchase order", "po #", "po number", "order confirmation",
                "placed an order", "order details", "quantity", "unit price",
            ],
            "enquiry": [
                "quote", "quotation", "rfq", "request for quote", "price inquiry",
                "pricing", "cost estimate", "how much", "availability", "samples",
            ],
            "invoice": [
                "invoice", "invoice number", "payment due", "outstanding",
                "remittance", "receipt", "billing", "amount due", "paid",
            ],
            "shipping": [
                "shipment", "delivery", "tracking number", "dispatched",
                "in transit", "arrived", "courier", "freight", "logistics",
            ],
        }

        scores = {cat: sum(1 for kw in kws if kw in text) for cat, kws in keyword_scores.items()}

        best = "general"
        confidence = 85
        reason = "No matching keywords found"
        if scores and max(scores.values()) > 0:
            best = max(scores, key=scores.get)  # type: ignore[arg-type]
            confidence = min(95, 60 + scores[best] * 15)
            reason = f"Keyword match: {scores[best]} keyword(s) found"

        fallback_data = {"Subject": subject[:100]}
        drafts = {
            "purchase_order": f"Dear Customer,\n\nThank you for your purchase order regarding '{subject}'. We have received it and are processing it.\n\nBest regards,\nProvident Packaging Operations",
            "enquiry": f"Hi there,\n\nThank you for your inquiry regarding '{subject}'. We are reviewing your requirements and will send pricing details shortly.\n\nBest regards,\nProvident Packaging Sales",
            "invoice": "Hello,\n\nWe have received your invoice. It has been routed to our accounts payable department for verification and processing.\n\nBest regards,\nProvident Packaging Accounts",
            "shipping": f"Hi,\n\nThank you for the shipping update regarding '{subject}'. We will monitor tracking for receipt.\n\nBest regards,\nProvident Packaging Logistics",
            "general": "Hello,\n\nThank you for your email. We have received it and will follow up shortly.\n\nBest regards,\nProvident Packaging Team",
        }

        return {
            "category": best,
            "confidence": confidence,
            "reason": reason,
            "extracted_data": json.dumps(fallback_data),
            "response_draft": drafts.get(best, drafts["general"]),
        }

    @staticmethod
    def _clean_text(text: str) -> str:
        if not text:
            return ""
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'<[^>]+>', '', text)
        return text.strip()


# Singleton — one instance per process (holds the OpenAI client)
_classifier_instance: Optional[EmailClassifier] = None


def get_classifier() -> EmailClassifier:
    """Return the shared classifier singleton."""
    global _classifier_instance
    if _classifier_instance is None:
        _classifier_instance = EmailClassifier()
    return _classifier_instance
