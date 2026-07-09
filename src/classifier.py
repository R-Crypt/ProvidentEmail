"""
Provident Packaging Email Classifier - AI Classification Engine
Uses OpenAI GPT-4o-mini to classify emails into business categories
"""
import json
import re
from typing import Dict, Optional
from openai import OpenAI
from config import config


class EmailClassifier:
    """AI-powered email classifier using OpenAI"""

    def __init__(self):
        self.categories = config.CATEGORIES
        self.model = config.OPENAI_MODEL
        if config.OPENAI_API_KEY:
            try:
                self.client = OpenAI(api_key=config.OPENAI_API_KEY)
            except Exception as e:
                print(f"Warning: Failed to initialize OpenAI client: {e}")
                self.client = None
        else:
            print("Warning: OpenAI API key not configured. Running in keyword fallback mode.")
            self.client = None

    def classify(self, subject: str, body: str) -> Dict:
        """
        Classify an email using OpenAI API
        Returns: {category, confidence, reason}
        """
        # Clean and truncate input
        clean_subject = self._clean_text(subject)
        clean_body = self._clean_text(body)[:3000]  # Limit to 3000 chars

        if not self.client:
            return self._fallback_classify(clean_subject, clean_body)

        # Build the prompt
        system_prompt = self._build_system_prompt()
        user_content = f"Subject: {clean_subject}\n\nBody:\n{clean_body}"

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_content}
                ],
                temperature=0,  # Deterministic output
                max_tokens=500, # Increased tokens to fit response draft
                response_format={"type": "json_object"}
            )

            result = json.loads(response.choices[0].message.content)

            # Validate category
            category = result.get("category", "general")
            if category not in self.categories:
                category = "general"

            # Parse estimated value safely
            raw_val = result.get("estimated_value_usd", 0)
            try:
                estimated_value = float(str(raw_val).replace(",", "").replace("$", ""))
            except (ValueError, TypeError):
                estimated_value = 0.0

            return {
                "category": category,
                "confidence": min(100, max(0, int(result.get("confidence", 80)))),
                "reason": result.get("reason", "AI classification"),
                "estimated_value": estimated_value,
                "extracted_data": json.dumps(result.get("extracted_data", {})),
                "response_draft": result.get("response_draft", "")
            }

        except Exception as e:
            # Fallback to keyword classifier if AI fails
            print(f"AI classification failed: {e}. Using fallback.")
            return self._fallback_classify(clean_subject, clean_body)

    def _build_system_prompt(self) -> str:
        """Build the system prompt for the AI"""
        category_descriptions = []
        for cat_id, cat_info in self.categories.items():
            category_descriptions.append(
                f'- "{cat_id}": {cat_info["description"]}'
            )

        categories_text = "\n".join(category_descriptions)

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

3. Estimate the dollar value (USD) of this email's business transaction:
   - For purchase_order/enquiry: estimate based on quantity × typical unit price for packaging.
   - For invoice: use the dollar amount explicitly stated.
   - For shipping/general: set to 0 if no monetary value is present.
   Return as a number in the "estimated_value_usd" field (e.g. 4500.00). Do NOT include $ signs.

4. Generate a professional, concise, ready-to-send draft reply to the email.
   - For "purchase_order": Acknowledge receipt, confirm the order details (quantities/items), and state we will process it.
   - For "enquiry": Thank them for their interest, state that we are calculating pricing for their requirements (mention items and quantities), and promise a quote back within 24 hours.
   - For "invoice": Acknowledge receipt, and say it has been forwarded to accounting for verification and payment processing.
   - For "shipping": Thank them for the shipment and tracking information.
   - For "general": A polite response acknowledging receipt of their message.

Respond ONLY with a JSON object in this exact format:
{{
    "category": "category_id",
    "confidence": 85,
    "reason": "Brief explanation of why this category was chosen",
    "estimated_value_usd": 4500.00,
    "extracted_data": {{
        "Key": "Value"
    }},
    "response_draft": "Email reply body here..."
}}"""

    def _fallback_classify(self, subject: str, body: str) -> Dict:
        """Simple keyword-based fallback if AI fails"""
        text = (subject + " " + body).lower()

        keyword_scores = {
            "purchase_order": ["purchase order", "po #", "po number", "order confirmation", 
                              "placed an order", "order details", "quantity", "unit price"],
            "enquiry": ["quote", "quotation", "rfq", "request for quote", "price inquiry",
                       "pricing", "cost estimate", "how much", "availability", "samples"],
            "invoice": ["invoice", "invoice number", "payment due", "outstanding",
                       "remittance", "receipt", "billing", "amount due", "paid"],
            "shipping": ["shipment", "delivery", "tracking number", "dispatched",
                        "in transit", "arrived", "courier", "freight", "logistics"]
        }

        scores = {}
        for cat, keywords in keyword_scores.items():
            score = sum(1 for kw in keywords if kw in text)
            scores[cat] = score

        best = "general"
        confidence = 85
        reason = "No matching keywords found"
        if scores and max(scores.values()) > 0:
            best = max(scores, key=scores.get)
            confidence = min(95, 60 + scores[best] * 15)
            reason = f"Keyword match: {scores[best]} keywords found"

        # Generate simple fallback data and draft
        fallback_data = {"Subject": subject[:100]}
        if best == "purchase_order":
            fallback_data["Details"] = "Detected order keywords in subject/body."
            draft = f"Dear Customer,\n\nThank you for your purchase order regarding '{subject}'. We have received it and are processing it.\n\nBest regards,\nProvident Packaging Operations"
        elif best == "enquiry":
            fallback_data["Details"] = "Detected quote/pricing inquiry keywords."
            draft = f"Hi there,\n\nThank you for your inquiry regarding '{subject}'. We are reviewing your requirements and will get back to you with pricing details shortly.\n\nBest regards,\nProvident Packaging Sales"
        elif best == "invoice":
            fallback_data["Details"] = "Detected billing/invoice keywords."
            draft = f"Hello,\n\nWe have received your invoice. It has been routed to our accounts payable department for verification and processing.\n\nBest regards,\nProvident Packaging Accounts"
        elif best == "shipping":
            fallback_data["Details"] = "Detected shipping/tracking details."
            draft = f"Hi,\n\nThank you for the shipping update regarding '{subject}'. We will monitor tracking for receipt.\n\nBest regards,\nProvident Packaging Logistics"
        else:
            draft = f"Hello,\n\nThank you for your email. We have received it and will follow up with you shortly.\n\nBest regards,\nProvident Packaging Team"

        return {
            "category": best,
            "confidence": confidence,
            "reason": reason,
            "estimated_value": 0.0,
            "extracted_data": json.dumps(fallback_data),
            "response_draft": draft
        }

    def _clean_text(self, text: str) -> str:
        """Clean text for processing"""
        if not text:
            return ""
        # Remove excessive whitespace
        text = re.sub(r'\s+', ' ', text)
        # Remove HTML tags if present
        text = re.sub(r'<[^>]+>', '', text)
        return text.strip()

    def classify_batch(self, emails: list) -> list:
        """Classify multiple emails"""
        results = []
        for email in emails:
            body = email.get("bodyPreview") or email.get("body", {}).get("content", "")
            result = self.classify(
                email.get("subject", ""),
                body
            )
            result["message_id"] = email.get("id")
            result["conversation_id"] = email.get("conversationId")
            result["user_email"] = email.get("user_email")
            result["sender"] = email.get("from", {}).get("emailAddress", {}).get("address", "")
            result["received_at"] = email.get("receivedDateTime")
            result["subject"] = email.get("subject", "")
            result["source_folder"] = email.get("source_folder", "Inbox")
            results.append(result)
        return results
