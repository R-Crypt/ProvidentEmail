#!/bin/bash
# ============================================================
# Container Entrypoint Script
# Runs on every container start. Ensures DB is ready,
# runs Alembic migrations, then starts Gunicorn.
# ============================================================
set -e

echo "=== Provident Operations Copilot ==="
echo "Environment: ${ENVIRONMENT:-production}"
echo "=== Waiting for PostgreSQL... ==="

# Wait for Postgres to be available before starting the app.
# In Docker Compose, the 'depends_on: condition: service_healthy' handles this,
# but this is a belt-and-suspenders fallback.
if [ -n "$DATABASE_URL" ] && [[ "$DATABASE_URL" != sqlite* ]]; then
    # Extract host and port from DATABASE_URL
    DB_HOST=$(echo "$DATABASE_URL" | sed -n 's/.*@\([^:]*\):.*/\1/p')
    DB_PORT=$(echo "$DATABASE_URL" | sed -n 's/.*:\([0-9]*\)\/.*/\1/p')
    DB_PORT=${DB_PORT:-5432}

    MAX_RETRIES=30
    RETRY=0
    until pg_isready -h "$DB_HOST" -p "$DB_PORT" 2>/dev/null || [ $RETRY -ge $MAX_RETRIES ]; do
        RETRY=$((RETRY+1))
        echo "Waiting for DB ($DB_HOST:$DB_PORT)... attempt $RETRY/$MAX_RETRIES"
        sleep 2
    done

    if [ $RETRY -ge $MAX_RETRIES ]; then
        echo "ERROR: Could not reach the database after $MAX_RETRIES attempts."
        exit 1
    fi
    echo "=== Database is ready ==="
fi

echo "=== Starting Gunicorn ==="
exec gunicorn app.main:app \
    --worker-class uvicorn.workers.UvicornWorker \
    --workers ${GUNICORN_WORKERS:-2} \
    --bind 0.0.0.0:8000 \
    --timeout 120 \
    --keepalive 5 \
    --log-level ${LOG_LEVEL:-info} \
    --access-logfile - \
    --error-logfile -
