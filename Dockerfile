# ============================================
# Invoice Generator - Docker Image
# ============================================
# Build:  docker compose build
# Run:    docker compose up -d
# Stop:   docker compose down

FROM python:3.12-slim

# Set working directory
WORKDIR /app

# Update OS packages to patch system vulnerabilities (tar, glibc, etc.)
RUN apt-get update && apt-get upgrade -y && rm -rf /var/lib/apt/lists/*

# Install dependencies first (layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy application source code
COPY api/ ./api/
COPY core/ ./core/
COPY frontend/ ./frontend/
COPY database/ ./default_database/

# Copy and set up entrypoint script
COPY docker_entrypoint.sh /app/
RUN chmod +x /app/docker_entrypoint.sh

# Set Python path so module imports resolve correctly
ENV PYTHONPATH=.

# Expose the FastAPI port
EXPOSE 8080

# Start the application via entrypoint
ENTRYPOINT ["/app/docker_entrypoint.sh"]
CMD ["uvicorn", "api.main:app", "--host", "0.0.0.0", "--port", "8080"]
