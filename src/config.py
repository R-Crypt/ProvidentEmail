"""
Provident Packaging Email Classifier - Configuration
"""
import os
from dataclasses import dataclass, field
from typing import Dict
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


@dataclass
class Config:
    """Application configuration loaded from environment variables"""

    # Microsoft 365 / Azure AD
    TENANT_ID: str = field(default_factory=lambda: os.getenv("TENANT_ID", ""))
    CLIENT_ID: str = field(default_factory=lambda: os.getenv("CLIENT_ID", ""))
    CLIENT_SECRET: str = field(default_factory=lambda: os.getenv("CLIENT_SECRET", ""))

    # OpenAI
    OPENAI_API_KEY: str = field(default_factory=lambda: os.getenv("OPENAI_API_KEY", ""))
    OPENAI_MODEL: str = "gpt-4o-mini"

    # Scheduling
    CHECK_INTERVAL_MINUTES: int = field(default_factory=lambda: int(os.getenv("CHECK_INTERVAL_MINUTES", "15")))

    # Logging
    LOG_LEVEL: str = field(default_factory=lambda: os.getenv("LOG_LEVEL", "INFO"))

    # Database
    DATABASE_URL: str = field(default_factory=lambda: os.getenv("DATABASE_URL", "sqlite:///data/processed_emails.db"))

    # Sentry (optional error tracking)
    SENTRY_DSN: str = field(default_factory=lambda: os.getenv("SENTRY_DSN", ""))

    # Auto-Reply setting
    AUTO_REPLY_ENABLED: bool = False

    # Email processing
    EMAILS_PER_BATCH: int = 50
    LOOKBACK_MINUTES: int = 30  # How far back to check for new emails

    # Classification categories with Outlook category names
    CATEGORIES: Dict[str, Dict] = field(default_factory=lambda: {
        "purchase_order": {
            "outlook_category": "Purchase Order",
            "display_name": "📦 Purchase Order",
            "color": "preset0",  # Green in Outlook
            "description": "Orders, PO numbers, order confirmations, quantity requests"
        },
        "enquiry": {
            "outlook_category": "Enquiry",
            "display_name": "❓ Enquiry / RFQ",
            "color": "preset1",  # Blue
            "description": "Quote requests, RFQs, pricing questions, product inquiries"
        },
        "invoice": {
            "outlook_category": "Invoice",
            "display_name": "💰 Invoice / Payment",
            "color": "preset2",  # Red
            "description": "Payment requests, invoices, billing, amount due"
        },
        "shipping": {
            "outlook_category": "Shipping",
            "display_name": "🚚 Shipping / Delivery",
            "color": "preset3",  # Yellow
            "description": "Tracking updates, delivery notifications, freight, logistics"
        },
        "general": {
            "outlook_category": "General",
            "display_name": "📧 General / Operations",
            "color": "preset4",  # Gray
            "description": "Meetings, internal emails, general correspondence"
        }
    })

    def validate(self) -> bool:
        """Check that all required config is present"""
        required = [self.TENANT_ID, self.CLIENT_ID, self.CLIENT_SECRET, self.OPENAI_API_KEY]
        missing = [name for name, val in zip(
            ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "OPENAI_API_KEY"], 
            required
        ) if not val]

        if missing:
            print(f"ERROR: Missing required environment variables: {', '.join(missing)}")
            print("Please set them in your .env file or environment.")
            return False
        return True


# Global config instance
config = Config()
