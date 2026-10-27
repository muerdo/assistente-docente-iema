# Assistente Docente — IEMA Açailândia

Template para os alunos montarem o assistente docente (WhatsApp via OpenWA, MCP, rotinas pedagógicas).
Não contém dados pessoais nem credenciais: tudo é configurado por variáveis de ambiente.

## Início rápido
1. `cp .env.example .env` e preencha (OpenWA, números, grupos, LIDs). Identificadores/grupos não preenchidos ficam desativados (nenhum acesso é concedido por padrão).
2. `set -a; source .env; set +a`
3. `pip install -r deploy_vps/requirements.vps.txt`
4. `python3 core/init_escola_db.py` (cria o banco SQLite local em `$DOCENTE_HOME`)
5. `python3 core/webhook_zap.py` (porta 5000) — aponte o webhook do OpenWA para ela.
6. Docker: `cd deploy_vps && cp .env.vps.example .env && docker compose -f docker-compose.vps.yml up -d`

## Estrutura
- `core/` webhook, MCP, cron, observador, orquestrador, RAG docente
- `seguranca/` anti-loop, contatos restritos, pontes
- `cli_gestao/` CLI e kill switch
- `geradores_pedagogicos/` geração de aulas/gabaritos
- `docentes/exemplo/` modelo de perfil de docente
- `web/` portal Flask
- `configs/` OmniRoute e schemas MCP

## Segurança
Nunca faça commit do `.env`, de bancos `.db`, logs ou dados de alunos/docentes.
