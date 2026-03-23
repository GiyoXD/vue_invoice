# ============================================
# Invoice Generator - Docker Image
# ============================================
# Build:  docker compose build
# Run:    docker compose up -d
# Stop:   docker compose down

FROM python:3.12-slim

# Set working directory
WORKDIR /app

# Install dependencies first (layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source code
COPY api/ ./api/
COPY core/ ./core/
COPY frontend/ ./frontend/

# Set Python path so module imports resolve correctly
ENV PYTHONPATH=.

# Expose the FastAPI port
EXPOSE 8000

# Start the application
CMD ["uvicorn", "api.main:app", "--host", "0.0.0.0", "--port", "8000"]
