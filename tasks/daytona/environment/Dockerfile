FROM node:20-bookworm-slim AS node_runtime

FROM python:3.12-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PIP_NO_CACHE_DIR=1

# Copy Node.js + npm from official Node image to keep the Python base lean.
COPY --from=node_runtime /usr/local/bin/node /usr/local/bin/node
COPY --from=node_runtime /usr/local/lib/node_modules /usr/local/lib/node_modules
RUN ln -sf /usr/local/lib/node_modules/npm/bin/npm-cli.js /usr/local/bin/npm \
    && ln -sf /usr/local/lib/node_modules/npm/bin/npx-cli.js /usr/local/bin/npx

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
    bash \
    build-essential \
    ca-certificates \
    curl \
    git \
    graphviz \
    ghostscript \
    libfreetype6-dev \
    libcairo2 \
    libgdk-pixbuf-2.0-0 \
    libjpeg62-turbo-dev \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    procps \
    libreoffice-calc \
    libreoffice-writer \
    pandoc \
    poppler-utils \
    wkhtmltopdf \
    zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

RUN python -m pip install --upgrade pip \
    && pip install --no-cache-dir \
    numpy==2.2.6 \
    pandas==2.3.3 \
    matplotlib==3.10.7 \
    pypdf==6.7.0 \
    python-docx==0.8.11 \
    python-pptx==0.6.21 \
    openpyxl==3.0.10 \
    xlrd==2.0.1 \
    PyMuPDF==1.26.5 \
    pdf2image==1.16.3 \
    pdfplumber==0.6.2 \
    pdfkit==0.6.1 \
    pypandoc==1.6.3 \
    docx2txt==0.8 \
    odfpy==1.4.1 \
    pyxlsb==1.0.8 \
    tabula==1.0.5 \
    camelot-py==0.10.1 \
    fpdf2==2.8.3 \
    reportlab==3.6.12 \
    weasyprint==53.3 \
    pdfrw==0.4 \
    graphviz==0.17 \
    pydot==1.4.2 \
    networkx==2.8.8 \
    svglib==1.1.0 \
    svgwrite==1.4.1 \
    cairosvg==2.5.2 \
    trimesh==3.9.29 \
    wordcloud==1.9.4 \
    folium==0.12.1 \
    lxml==5.4.0 \
    defusedxml==0.7.1 \
    Pillow \
    markitdown[pptx] \
    'stirrup[litellm]==0.1.8'

RUN npm install -g --no-audit --no-fund \
    @openai/codex \
    @anthropic-ai/claude-code \
    @google/gemini-cli \
    opencode-ai \
    docx \
    pptxgenjs \
    && npm cache clean --force \
    && rm -rf /root/.npm

# Preinstall Claude's native launcher once to avoid parallel setup OOM kills.
RUN curl -fsSL https://claude.ai/install.sh | bash -s -- latest
ENV PATH="/root/.local/bin:${PATH}"

# python-pptx 0.6.21 expects `collections.abc` to be preloaded on this Python build.
RUN printf '%s\n' \
    'import collections' \
    'import collections.abc' \
    > /usr/local/lib/python3.12/site-packages/sitecustomize.py

COPY data/ /workspace/inputs/
COPY skills/ /workspace/skills/

# Install task-specific requirements when present.
RUN if [ -f /workspace/inputs/requirements.txt ]; then \
      pip install --no-cache-dir -r /workspace/inputs/requirements.txt; \
    fi \
    && if [ -f /workspace/inputs/data/requirements.txt ]; then \
      pip install --no-cache-dir -r /workspace/inputs/data/requirements.txt; \
    fi

WORKDIR /workspace
