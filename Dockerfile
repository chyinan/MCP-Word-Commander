# MCP Word Commander Docker Image
FROM python:3.12-slim

# Set working directory
WORKDIR /app

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# Install system dependencies (if needed for python-docx)
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY server.py .

# Create a directory for documents (can be mounted as volume)
RUN mkdir -p /documents

# Set the documents directory as working directory for file operations
WORKDIR /documents

# Expose no ports - MCP uses stdio transport
# The server communicates via stdin/stdout

# Set the entrypoint
ENTRYPOINT ["python", "/app/server.py"]
