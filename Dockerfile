FROM python:3.11-slim

WORKDIR /app

# System deps for geopandas / GDAL + OCR
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgdal-dev libgeos-dev tesseract-ocr && \
    rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Cloud Run injects PORT env var (default 8080)
ENV PORT=8080
EXPOSE 8080

CMD exec gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
