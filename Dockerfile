# Use the official, multi-architecture Python image.
# This will automatically use the arm64 version on your M-series Mac.
FROM python:3.13-slim-bookworm

# Install Node.js, npm, nginx, and other tools.
# CHANGED: 'chromium-browser' to 'chromium', which is the correct package name in Debian 12.
RUN apt-get update && apt-get install -y --no-install-recommends \
    nodejs \
    npm \
    nginx \
    curl \
    redis-server \
    chromium

# Create a working directory
WORKDIR /app

# Set environment variables for the app
ENV APP_DATA_DIRECTORY=/app/user_data
ENV TEMP_DIRECTORY=/tmp/presenton

# CHANGED: The executable path now points to the binary provided by the 'chromium' package.
ENV PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium

# Install ollama (their script is architecture-aware and should work fine)
RUN curl -fsSL https://ollama.com/install.sh | sh

# Install dependencies for FastAPI
COPY servers/fastapi/requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Install dependencies for Next.js
WORKDIR /app/servers/nextjs
COPY servers/nextjs/package*.json ./
RUN npm ci

# Copy Next.js app
COPY servers/nextjs/ /app/servers/nextjs/

# Build the Next.js app
WORKDIR /app/servers/nextjs
RUN npm run build

# Reset working directory to the root of the app
WORKDIR /app

# Copy FastAPI and other root files
COPY servers/fastapi/ ./servers/fastapi/
COPY start.js LICENSE NOTICE ./
COPY docker-start.sh /app/docker-start.sh
RUN chmod +x /app/docker-start.sh # Make the start script executable

# Copy nginx configuration
COPY nginx.conf /etc/nginx/nginx.conf

# Expose the port Nginx will listen on
EXPOSE 80

# Start the servers
CMD ["/bin/bash", "/app/docker-start.sh"]
