"""
Structured JSON logging setup.
Every log line is a JSON object so it can be ingested by Datadog, Render logs, etc.
"""
import logging
import sys
import json
from datetime import datetime, timezone
from app.core.config import settings


class JsonFormatter(logging.Formatter):
    """Format log records as single-line JSON objects."""

    def format(self, record: logging.LogRecord) -> str:
        log_obj = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "level": record.levelname,
            "service": "provident-copilot",
            "version": settings.APP_VERSION,
            "environment": settings.ENVIRONMENT,
            "logger": record.name,
            "message": record.getMessage(),
        }

        # Include exception info if present
        if record.exc_info:
            log_obj["exception"] = self.formatException(record.exc_info)

        # Include any extra fields passed via `extra={}` in log calls
        for key, value in record.__dict__.items():
            if key not in (
                "name", "msg", "args", "levelname", "levelno", "pathname",
                "filename", "module", "exc_info", "exc_text", "stack_info",
                "lineno", "funcName", "created", "msecs", "relativeCreated",
                "thread", "threadName", "processName", "process", "message",
                "taskName",
            ):
                log_obj[key] = value

        return json.dumps(log_obj, default=str)


def configure_logging() -> None:
    """Configure root logger. Call once at application startup."""
    level = getattr(logging, settings.LOG_LEVEL.upper(), logging.INFO)

    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(JsonFormatter())

    # Root logger
    root = logging.getLogger()
    root.setLevel(level)
    root.handlers.clear()
    root.addHandler(handler)

    # Silence noisy third-party loggers
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("sqlalchemy.engine").setLevel(logging.WARNING)
    logging.getLogger("apscheduler").setLevel(logging.INFO)

    logging.getLogger(__name__).info(
        "Logging configured",
        extra={"log_level": settings.LOG_LEVEL, "environment": settings.ENVIRONMENT},
    )
