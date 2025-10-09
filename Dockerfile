FROM python:3.11-alpine
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

# Tell Puppeteer to skip installing Chrome. We'll use the installed chromium
ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true \
    PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium-browser

# Install mermaid-cli globally
RUN npm install -g @mermaid-js/mermaid-cli

# Copy and install Python requirements
COPY requirements.txt ./
RUN pip install -r requirements.txt --no-cache-dir

USER 1000

# Copy the Python script and puppeteer config
COPY src/send_to_sharepoint.py /usr/src/app/
COPY src/puppeteer-config.json /usr/src/app/
# full path is necessary or it defaults to main branch copy
ENTRYPOINT [ "python", "/usr/src/app/send_to_sharepoint.py" ]