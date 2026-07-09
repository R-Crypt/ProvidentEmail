"""
Provident Packaging — Application Configuration
Uses Pydantic BaseSettings to load and validate all environment variables.
"""
from functools import lru_cache
from typing import Dict, List
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """
    All configuration is driven by environment variables.
    Never hard-code secrets — set them in .env (dev) or your host's secret store (prod).
    """

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=True,
        extra="ignore",
    )

    # ------------------------------------------------------------------ #
    #  App
    # ------------------------------------------------------------------ #
    APP_NAME: str = "Provident Operations Copilot"
    APP_VERSION: str = "2.0.0"
    ENVIRONMENT: str = Field(default="production", description="development | production")
    DEBUG: bool = False

    # Public URL this service is reachable at — used in manifest.xml generation
    # e.g. https://provident-copilot.onrender.com
    PUBLIC_URL: str = Field(default="https://localhost:7071", description="Production HTTPS base URL")

    # ------------------------------------------------------------------ #
    #  Security / JWT
    # ------------------------------------------------------------------ #
    SECRET_KEY: str = Field(default="CHANGE_ME_IN_PRODUCTION_32chars+", min_length=32)
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 60 * 8  # 8 hours
    BYPASS_AUTH: bool = Field(default=False, description="Bypass Azure AD authentication for testing")

    # ------------------------------------------------------------------ #
    #  CORS — list of allowed origins (set to PUBLIC_URL in production)
    # ------------------------------------------------------------------ #
    ALLOWED_ORIGINS: List[str] = Field(
        default=["https://localhost:7071", "https://localhost:3000"]
    )

    # ------------------------------------------------------------------ #
    #  Microsoft 365 / Azure AD
    # ------------------------------------------------------------------ #
    TENANT_ID: str = Field(default="")
    CLIENT_ID: str = Field(default="")
    CLIENT_SECRET: str = Field(default="")

    # Microsoft's OIDC / OAuth2 endpoints (derived from TENANT_ID, set at runtime)
    @property
    def ms_token_url(self) -> str:
        return f"https://login.microsoftonline.com/{self.TENANT_ID}/oauth2/v2.0/token"

    @property
    def ms_jwks_uri(self) -> str:
        return f"https://login.microsoftonline.com/{self.TENANT_ID}/discovery/v2.0/keys"

    # ------------------------------------------------------------------ #
    #  OpenAI
    # ------------------------------------------------------------------ #
    OPENAI_API_KEY: str = Field(default="")
    OPENAI_MODEL: str = "gpt-4o-mini"

    # ------------------------------------------------------------------ #
    #  Database (PostgreSQL in production, SQLite in development)
    # ------------------------------------------------------------------ #
    DATABASE_URL: str = Field(
        default="postgresql+asyncpg://provident:provident@localhost:5432/provident_db",
        description="Async-compatible SQLAlchemy DSN",
    )
    DB_POOL_SIZE: int = 10
    DB_MAX_OVERFLOW: int = 20

    # ------------------------------------------------------------------ #
    #  Scheduling
    # ------------------------------------------------------------------ #
    CHECK_INTERVAL_MINUTES: int = 15
    LOOKBACK_MINUTES: int = 30
    EMAILS_PER_BATCH: int = 50

    # ------------------------------------------------------------------ #
    #  Monitoring
    # ------------------------------------------------------------------ #
    SENTRY_DSN: str = Field(default="")
    LOG_LEVEL: str = "INFO"

    # ------------------------------------------------------------------ #
    #  Email Categories (static — would move to DB for multi-tenant)
    # ------------------------------------------------------------------ #
    @property
    def CATEGORIES(self) -> Dict[str, Dict]:
        return {
            "purchase_order": {
                "outlook_category": "Purchase Order",
                "display_name": "Purchase Order",
                "color": "preset0",
                "description": "Orders, PO numbers, order confirmations, quantity requests",
            },
            "enquiry": {
                "outlook_category": "Enquiry",
                "display_name": "Enquiry / RFQ",
                "color": "preset1",
                "description": "Quote requests, RFQs, pricing questions, product inquiries",
            },
            "invoice": {
                "outlook_category": "Invoice",
                "display_name": "Invoice / Payment",
                "color": "preset2",
                "description": "Payment requests, invoices, billing, amount due",
            },
            "shipping": {
                "outlook_category": "Shipping",
                "display_name": "Shipping / Delivery",
                "color": "preset3",
                "description": "Tracking updates, delivery notifications, freight, logistics",
            },
            "general": {
                "outlook_category": "General",
                "display_name": "General / Operations",
                "color": "preset4",
                "description": "Meetings, internal emails, general correspondence",
            },
            "junk": {
                "outlook_category": "Junk",
                "display_name": "Junk Email",
                "color": "preset5",
                "description": "Spam, marketing, newsletters, or irrelevant emails",
            },
        }

    def validate_required(self) -> list[str]:
        """Return list of missing required env var names."""
        missing = []
        if not self.TENANT_ID:
            missing.append("TENANT_ID")
        if not self.CLIENT_ID:
            missing.append("CLIENT_ID")
        if not self.CLIENT_SECRET:
            missing.append("CLIENT_SECRET")
        if not self.OPENAI_API_KEY:
            missing.append("OPENAI_API_KEY")
        return missing


@lru_cache
def get_settings() -> Settings:
    """Cached settings singleton — call this everywhere instead of importing directly."""
    return Settings()


# Convenience alias for imports
settings = get_settings()
