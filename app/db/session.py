"""
Async SQLAlchemy database engine and session factory.
Uses asyncpg driver for PostgreSQL in production.
"""
import logging
from contextlib import asynccontextmanager
from typing import AsyncGenerator

from sqlalchemy.ext.asyncio import (
    AsyncSession,
    async_sessionmaker,
    create_async_engine,
)
from sqlalchemy.orm import DeclarativeBase

from app.core.config import settings

logger = logging.getLogger(__name__)


class Base(DeclarativeBase):
    """Shared declarative base for all ORM models."""
    pass


# ---------------------------------------------------------------------------
# Engine — created once at module import time
# ---------------------------------------------------------------------------
engine = create_async_engine(
    settings.DATABASE_URL,
    pool_size=settings.DB_POOL_SIZE,
    max_overflow=settings.DB_MAX_OVERFLOW,
    echo=settings.DEBUG,          # Log SQL queries only in debug mode
    pool_pre_ping=True,           # Verify connections before use (handles Postgres restarts)
    connect_args={
        "prepared_statement_cache_size": 0  # Required for Supabase PgBouncer (Transaction mode)
    }
)

# Session factory — produces AsyncSession instances
AsyncSessionLocal = async_sessionmaker(
    bind=engine,
    class_=AsyncSession,
    expire_on_commit=False,       # Avoids lazy-load issues after commit
)


async def init_db() -> None:
    """
    Create all tables that don't exist yet.
    In production, prefer running Alembic migrations instead.
    This is kept as a fallback / first-run convenience.
    """
    async with engine.begin() as conn:
        # Import models here so Base.metadata is populated before create_all
        from app.models import domain  # noqa: F401
        await conn.run_sync(Base.metadata.create_all)
    logger.info("Database tables verified/created")


async def close_db() -> None:
    """Dispose the engine connection pool on shutdown."""
    await engine.dispose()
    logger.info("Database engine disposed")


@asynccontextmanager
async def get_db_context() -> AsyncGenerator[AsyncSession, None]:
    """Async context manager — use in background tasks / workers."""
    async with AsyncSessionLocal() as session:
        try:
            yield session
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()


async def get_db() -> AsyncGenerator[AsyncSession, None]:
    """
    FastAPI dependency — yields a database session per request.
    Auto-commits on success, rolls back on exception.

    Usage:
        async def my_route(db: AsyncSession = Depends(get_db)):
            ...
    """
    async with AsyncSessionLocal() as session:
        try:
            yield session
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()
