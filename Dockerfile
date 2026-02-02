# --- Stage 1: Build Frontend (Node.js) ---
FROM node:18-alpine as frontend_build

WORKDIR /app/frontend

# Install dependencies (cache optimized)
COPY frontend/package.json ./
# If lockfile exists, copy it too (optional but recommended)
# COPY frontend/package-lock.json ./ 
RUN npm install

# Copy source and build
COPY frontend/ ./
RUN npm run build

# --- Stage 2: Runtime (Python) ---
FROM python:3.11-slim

WORKDIR /app

# Install system dependencies (if needed, e.g. for some python packages)
# RUN apt-get update && apt-get install -y --no-install-recommends gcc && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy Backend Code
COPY backend/ ./backend/

# Copy Built Frontend from Stage 1
# 1. Assets directly to Flask static folder
COPY --from=frontend_build /app/frontend/dist/assets ./backend/static/assets
# 2. HTML to Flask templates folder
COPY --from=frontend_build /app/frontend/dist/index.html ./backend/templates/index.html

# Copy Project Root Files (CSV schema mapping logic requires reading columns, but strictly data_loader uses hardcoded map)
# However, user requests assume CSV in root. 
# We create uploads/outputs directories
RUN mkdir -p uploads
RUN mkdir -p backend/outputs

# FIX: Adjust index.html to load assets from /static/ instead of root /
# Vite builds absolute paths "/assets/..." by default. Flask serves at "/static/assets/...".
RUN sed -i 's|/assets/|/static/assets/|g' ./backend/templates/index.html

# Environment Variables
ENV FLASK_APP=backend/app.py
ENV FLASK_ENV=production
ENV PYTHONUNBUFFERED=1

# Expose Port
EXPOSE 5000

# Start Command using Gunicorn for production stability
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120", "backend.app:app"]
