FROM python:3.13-slim

WORKDIR /app

# Ensure non-buffered Python output
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PPT_STORAGE_DIR=/app/storage

# Install system dependencies (including font utilities if needed)
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Install python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source code and bundled assets
COPY survey_studio/ ./survey_studio/
COPY examples/ ./examples/

# Create persistent storage directories
RUN mkdir -p /app/storage/uploads /app/storage/generated /app/storage/templates

# Expose FastAPI port
EXPOSE 8001

# Run Uvicorn
CMD ["uvicorn", "survey_studio.app:app", "--host", "0.0.0.0", "--port", "8001"]
