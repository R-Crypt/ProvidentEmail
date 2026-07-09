"""
Provident Operations Copilot — FastAPI Application Factory
"""
import logging
from contextlib import asynccontextmanager
from typing import AsyncGenerator

import sentry_sdk
from fastapi import FastAPI, Request, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from jose import JWTError

from app.core.config import settings
from app.core.logging import configure_logging

# Configure logging before anything else
configure_logging()
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Lifespan — startup / shutdown
# ---------------------------------------------------------------------------

@asynccontextmanager
async def lifespan(app: FastAPI) -> AsyncGenerator[None, None]:
    """
    Manages the full lifecycle of the application.
    Startup: open DB + HTTP clients, run migrations, start scheduler.
    Shutdown: close clients, stop scheduler.
    """
    # --- STARTUP ---
    logger.info("Application starting up", extra={"version": settings.APP_VERSION})

    # Initialize Sentry error tracking (if configured)
    if settings.SENTRY_DSN:
        sentry_sdk.init(
            dsn=settings.SENTRY_DSN,
            environment=settings.ENVIRONMENT,
            release=settings.APP_VERSION,
            traces_sample_rate=0.1,
        )
        logger.info("Sentry initialized")

    # Open the Graph API HTTP client
    from app.services.graph_client import graph_client
    await graph_client.open()

    # Ensure DB tables exist (idempotent; Alembic handles migrations in prod)
    from app.db.session import init_db
    await init_db()

    # Start the background email batch scheduler
    from app.workers.scheduler import start_scheduler
    start_scheduler()

    logger.info("Application startup complete")
    yield  # <-- App runs here

    # --- SHUTDOWN ---
    logger.info("Application shutting down")
    from app.workers.scheduler import stop_scheduler
    stop_scheduler()
    await graph_client.close()
    from app.db.session import close_db
    await close_db()
    logger.info("Application shutdown complete")


# ---------------------------------------------------------------------------
# App factory
# ---------------------------------------------------------------------------

def create_app() -> FastAPI:
    app = FastAPI(
        title=settings.APP_NAME,
        version=settings.APP_VERSION,
        description="AI-powered email classifier and operations copilot for Provident Packaging.",
        docs_url="/docs" if settings.ENVIRONMENT == "development" else None,
        redoc_url="/redoc" if settings.ENVIRONMENT == "development" else None,
        openapi_url="/openapi.json" if settings.ENVIRONMENT == "development" else None,
        lifespan=lifespan,
    )

    # ------------------------------------------------------------------
    # CORS — must be first middleware
    # ------------------------------------------------------------------
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.ALLOWED_ORIGINS,
        allow_credentials=True,
        allow_methods=["GET", "POST", "PATCH", "OPTIONS"],
        allow_headers=["Authorization", "Content-Type"],
        expose_headers=["X-Request-ID"],
    )

    # ------------------------------------------------------------------
    # Global exception handlers
    # ------------------------------------------------------------------

    @app.exception_handler(JWTError)
    async def jwt_error_handler(request: Request, exc: JWTError) -> JSONResponse:
        return JSONResponse(
            status_code=status.HTTP_401_UNAUTHORIZED,
            content={"detail": "Invalid or expired token."},
            headers={"WWW-Authenticate": "Bearer"},
        )

    @app.exception_handler(Exception)
    async def global_exception_handler(request: Request, exc: Exception) -> JSONResponse:
        logger.error(
            "Unhandled exception",
            extra={
                "path": str(request.url),
                "method": request.method,
                "error": str(exc),
            },
            exc_info=True,
        )
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={"detail": "An internal server error occurred."},
        )

    # ------------------------------------------------------------------
    # Routers
    # ------------------------------------------------------------------
    from app.api.routes.health import router as health_router
    from app.api.routes.addin import router as addin_router
    from app.api.routes.dashboard import router as dashboard_router

    app.include_router(health_router)
    app.include_router(addin_router)
    app.include_router(dashboard_router)

    # ------------------------------------------------------------------
    # Static files — serves the Outlook Add-in task pane HTML/CSS/JS
    # Under Nginx in production, Nginx serves these directly for better
    # performance. This mount is the fallback for Render / dev.
    # ------------------------------------------------------------------
    import os
    addin_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "outlook-addin")
    if os.path.isdir(addin_dir):
        app.mount("/addin", StaticFiles(directory=addin_dir, html=True), name="addin")
        logger.info("Outlook add-in static files mounted at /addin")

    logger.info(
        "FastAPI app created",
        extra={
            "environment": settings.ENVIRONMENT,
            "public_url": settings.PUBLIC_URL,
        },
    )

    return app


# Module-level app instance used by Gunicorn / Uvicorn
app = create_app()
