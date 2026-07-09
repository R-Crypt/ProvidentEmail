# ============================================================
# Multi-stage Dockerfile for Provident Operations Copilot
# Stage 1: Build dependencies (with build tools)
# Stage 2: Lean runtime image
# ============================================================

# -------- Stage 1: Build --------
FROM python:3.12-slim AS builder

WORKDIR /app

# Install build dependencies needed for asyncpg + cryptography
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy and install Python dependencies into a prefix
COPY requirements.txt .
RUN pip install --upgrade pip \
    && pip install --prefix=/install --no-cache-dir -r requirements.txt


# -------- Stage 2: Runtime --------
FROM python:3.12-slim AS runtime

WORKDIR /app

# Copy only installed packages from builder (no build tools in final image)
COPY --from=builder /install /usr/local

# Install runtime-only system libraries
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpq5 \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy application code
COPY app/ ./app/
COPY outlook-addin/ ./outlook-addin/
COPY scripts/entrypoint.sh ./scripts/entrypoint.sh

# Make entrypoint executable
RUN chmod +x ./scripts/entrypoint.sh

# Non-root user for security
RUN groupadd --gid 1001 appuser \
    && useradd --uid 1001 --gid 1001 --no-create-home appuser
USER appuser

# Expose FastAPI port
EXPOSE 8000

# Docker healthcheck using the /health endpoint
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD curl -f http://localhost:8000/health || exit 1

ENTRYPOINT ["./scripts/entrypoint.sh"]
