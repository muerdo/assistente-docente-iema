FROM python:3.11-slim

LABEL maintainer="Ecossistema Docente IEMA TI"
LABEL description="Container enxuto do Assistente Docente Multi-Professor para VPS de 1 GB RAM"

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    TZ=America/Fortaleza \
    MODO_VPS_LEVE=1 \
    DOCENTE_HOME=/app \
    ECOSSISTEMA_DIR=/app \
    ESCOLA_DB=/app/dados/escola.db \
    PASTA_UPLOADS_ISOLADOS=/app/uploads_isolados

RUN apt-get update && apt-get install -y --no-install-recommends \
    sqlite3 \
    curl \
    poppler-utils \
    tzdata \
    ca-certificates \
    && ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY deploy_vps/requirements.vps.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copia módulos centrais leves e segurança
COPY core/webhook_zap.py /app/webhook_zap.py
COPY core/gestor_docentes_rag.py /app/gestor_docentes_rag.py
COPY core/ponte_agente.py /app/ponte_agente.py
COPY core/observador_docente.py /app/observador_docente.py
COPY seguranca/anti_loop.py /app/anti_loop.py
COPY seguranca/contato_restrito.py /app/contato_restrito.py
COPY web /app/web

# Cria diretórios de dados persistentes, uploads isolados e base multi-docente
RUN mkdir -p /app/dados /app/docentes /app/uploads_isolados /app/logs

EXPOSE 5000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -fsS http://127.0.0.1:5000/health || exit 1

CMD ["python3", "-u", "/app/webhook_zap.py"]
