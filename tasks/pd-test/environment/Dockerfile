FROM python:3.12-slim-bookworm

RUN pip install --no-cache-dir openpyxl==3.0.10 python-docx==0.8.11

RUN apt-get update && apt-get install -y --no-install-recommends \
    curl git ca-certificates \
    && rm -rf /var/lib/apt/lists/*

RUN npm install -g @anthropic-ai/claude-code 2>/dev/null || \
    (curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
     apt-get install -y nodejs && \
     npm install -g @anthropic-ai/claude-code && \
     rm -rf /var/lib/apt/lists/*)

WORKDIR /workspace
