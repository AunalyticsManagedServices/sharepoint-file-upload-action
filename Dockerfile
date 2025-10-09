FROM python:3.11-alpine3.19
HEALTHCHECK NONE
WORKDIR /usr/src/app

# Install Node.js, npm, and required dependencies for mermaid-cli
# Also install Chromium which is needed by Puppeteer (used by mermaid-cli)
RUN apk add --no-cache \
    nodejs \
    npm \
    chromium \
    nss \
    freetype \
    freetype-dev \
    harfbuzz \
    ca-certificates \
    ttf-freefont

# Environment variables matching official mermaid-cli Docker setup
ENV CHROME_BIN="/usr/bin/chromium-browser" \
    PUPPETEER_SKIP_DOWNLOAD="true" \
    PUPPETEER_EXECUTABLE_PATH="/usr/bin/chromium-browser"

# Install mermaid-cli globally with compatible puppeteer version
RUN npm install -g @mermaid-js/mermaid-cli@11.4.2

# Copy and install Python requirements
COPY requirements.txt ./
RUN pip install -r requirements.txt --no-cache-dir

# Create a proper non-root user (following mermaid-cli pattern)
RUN adduser -D -u 1000 sharepoint && \
    mkdir -p /tmp && \
    chmod 777 /tmp

USER sharepoint

# Copy the Python script and puppeteer config
COPY src/send_to_sharepoint.py /usr/src/app/
COPY src/puppeteer-config.json /usr/src/app/
# full path is necessary or it defaults to main branch copy
ENTRYPOINT [ "python", "/usr/src/app/send_to_sharepoint.py" ]