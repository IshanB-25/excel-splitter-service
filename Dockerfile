FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies if needed
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY app.py .

# Create non-root user for security
RUN useradd -m -u 1000 appuser && chown -R appuser:appuser /app
USER appuser

# Expose the port
EXPOSE 3070

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:3070/health', timeout=2)" || exit 1

# Gunicorn defaults tuned for memory-heavy workbook processing.
ENV GUNICORN_WORKERS=1
ENV GUNICORN_TIMEOUT=1800
ENV GUNICORN_GRACEFUL_TIMEOUT=1800
ENV GUNICORN_MAX_REQUESTS=25
ENV GUNICORN_MAX_REQUESTS_JITTER=10

# Run with gunicorn for production.
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:3070 --workers ${GUNICORN_WORKERS} --timeout ${GUNICORN_TIMEOUT} --graceful-timeout ${GUNICORN_GRACEFUL_TIMEOUT} --max-requests ${GUNICORN_MAX_REQUESTS} --max-requests-jitter ${GUNICORN_MAX_REQUESTS_JITTER} --worker-class sync --access-logfile - --error-logfile - app:app"]