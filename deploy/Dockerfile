# code-server (VS Code в браузере) + окружение для сборки отчётов GORD
FROM codercom/code-server:latest

USER root
RUN apt-get update && apt-get install -y --no-install-recommends \
        python3 python3-pip \
        libreoffice poppler-utils fontconfig \
        git curl ca-certificates gnupg \
    && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y nodejs \
    && npm install -g @anthropic-ai/claude-code \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

USER coder
