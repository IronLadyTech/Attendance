# Zoom → Zoho webhook bridge (Railway, Fly, etc.)
FROM python:3.12-slim

WORKDIR /app

COPY requirements-bridge.txt .
RUN pip install --no-cache-dir -r requirements-bridge.txt

COPY . .

ENV PYTHONUNBUFFERED=1

# Railway sets PORT automatically
CMD python tools/zoom_webhook_bridge.py --host 0.0.0.0
