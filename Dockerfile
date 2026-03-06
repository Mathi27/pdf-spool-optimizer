# ── Build stage ─────────────────────────────────────────────────────────────
FROM python:3.12-slim AS builder

WORKDIR /app

# Install OS-level dependencies required by PyMuPDF
RUN apt-get update && apt-get install -y --no-install-recommends \
        libmupdf-dev \
        mupdf-tools \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies into a prefix so they can be copied cleanly
COPY requirements.txt .
RUN pip install --no-cache-dir --prefix=/install -r requirements.txt

# ── Runtime stage ────────────────────────────────────────────────────────────
FROM python:3.12-slim AS runtime

WORKDIR /app

# Minimal runtime OS libs needed by PyMuPDF at runtime
RUN apt-get update && apt-get install -y --no-install-recommends \
        libmupdf-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy installed packages from builder
COPY --from=builder /install /usr/local

# Copy application source
COPY app.py spool_optimizer.py ./
COPY templates/ templates/
COPY assets/ assets/

# Non-root user for security
RUN useradd --no-create-home --shell /bin/false appuser \
    && chown -R appuser:appuser /app
USER appuser

# Flask / Gunicorn config via env (overridable at runtime)
ENV FLASK_ENV=production \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

EXPOSE 5000

# Use gunicorn for production; single worker is fine for a CPU-bound task
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120", "app:app"]
