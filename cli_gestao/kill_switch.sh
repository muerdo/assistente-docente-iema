#!/bin/bash
# ==============================================================================
# KILL SWITCH & GERENCIADOR CENTRAL — ECOSSISTEMA DOCENTE & LABORATÓRIOS
# ==============================================================================
# Uso:
#   ./kill_switch.sh off      (ou stop/kill) -> DESLIGA TUDO IMEDIATAMENTE (Kill Switch)
#   ./kill_switch.sh on       (ou start)     -> INICIA TUDO COM PROTEÇÃO ATIVA
#   ./kill_switch.sh status                  -> EXIBE PAINEL COMPLETO EM TEMPO REAL
#   ./kill_switch.sh restart                 -> REINICIA O ECOSSISTEMA
# ==============================================================================

SCRIPT_DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
PROJECT_DIR="${DOCENTE_HOME:-$(cd "$SCRIPT_DIR/.." && pwd)}"
VENV_PYTHON="${VENV_PYTHON:-python3}"
LOG_DIR="$PROJECT_DIR/logs"
CRON_BACKUP="$PROJECT_DIR/.crontab_docente.bak"
mkdir -p "$LOG_DIR"

# Cores
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
BOLD='\033[1m'
NC='\033[0m'

check_port() {
    local port=$1
    ss -tulpn 2>/dev/null | grep -q ":$port "
}

kill_process_by_pattern() {
    local pattern=$1
    local name=$2
    local pids=$(pgrep -f "$pattern" 2>/dev/null)
    if [ -n "$pids" ]; then
        echo -e "  ${YELLOW}↳ Matando $name (PIDs: $pids)...${NC}"
        kill -9 $pids 2>/dev/null
    fi
}

kill_switch_off() {
    echo -e "${RED}${BOLD}======================================================${NC}"
    echo -e "${RED}${BOLD}🚨 ACIONANDO KILL SWITCH GERAL (DESLIGANDO TUDO) 🚨${NC}"
    echo -e "${RED}${BOLD}======================================================${NC}"

    # 1. Parar container Docker WhatsApp
    echo -e "\n${CYAN}[1/4] Parando Container WhatsApp (openwa-api)...${NC}"
    if docker ps --format '{{.Names}}' | grep -q "^openwa-api$"; then
        docker stop -t 3 openwa-api 2>/dev/null || docker kill openwa-api 2>/dev/null
        echo -e "  ${GREEN}✔ Container WhatsApp parado (sessão e login preservados intactos no disco).${NC}"
    else
        echo -e "  ${GREEN}✔ Container WhatsApp já estava inativo.${NC}"
    fi

    # 2. Pausar agendamentos do Crontab
    echo -e "\n${CYAN}[2/4] Pausando rotinas do Crontab...${NC}"
    local current_cron=$(crontab -l 2>/dev/null)
    if [ -n "$current_cron" ]; then
        echo "$current_cron" > "$CRON_BACKUP"
        crontab -r 2>/dev/null
        echo -e "  ${GREEN}✔ Crontab salvo em $CRON_BACKUP e desativado.${NC}"
    else
        echo -e "  ${GREEN}✔ Crontab já está pausado.${NC}"
    fi

    # 3. Finalizar todos os processos Python docentes e bots
    echo -e "\n${CYAN}[3/4] Finalizando processos docentes, bots e listeners...${NC}"
    kill_process_by_pattern "webhook_zap.py" "Webhook WhatsApp"
    kill_process_by_pattern "mcp_docente.py" "Servidor MCP Docente"
    kill_process_by_pattern "mcp_cron.py" "MCP Cron"
    kill_process_by_pattern "cron_daemon.py" "Daemon de Alertas"
    kill_process_by_pattern "observador_docente.py" "Observador Docente"
    kill_process_by_pattern "avisos_contato.py" "Despachante de Avisos"
    kill_process_by_pattern "orquestrador_ideias.py" "Orquestrador de Ideias"
    kill_process_by_pattern "testar_docente_noturno.py" "Bateria Noturna"
    kill_process_by_pattern "omniroute serve" "OmniRoute Gateway"
    kill_process_by_pattern "prototipo_mvp" "Servidor Web Ideia 5"

    # 4. Verificação de portas liberadas
    echo -e "\n${CYAN}[4/4] Verificando liberação de portas...${NC}"
    if check_port 5000; then
        local p5000=$(lsof -t -i:5000 2>/dev/null)
        [ -n "$p5000" ] && kill -9 $p5000 2>/dev/null
    fi

    echo -e "\n${GREEN}${BOLD}======================================================${NC}"
    echo -e "${GREEN}${BOLD}🛑 STATUS: SISTEMA TOTALMENTE DESLIGADO COM SUCESSO!${NC}"
    echo -e "${GREEN}${BOLD}Nenhuma mensagem ou rotina será executada.${NC}"
    echo -e "${GREEN}${BOLD}Para religar a qualquer momento: ./kill_switch.sh on${NC}"
    echo -e "${GREEN}${BOLD}======================================================${NC}"
}

kill_switch_on() {
    echo -e "${GREEN}${BOLD}======================================================${NC}"
    echo -e "${GREEN}${BOLD}🚀 LIGANDO ECOSSISTEMA DOCENTE & LABORATÓRIOS 🚀${NC}"
    echo -e "${GREEN}${BOLD}======================================================${NC}"

    # 1. Rclone Mount (Google Drive)
    echo -e "\n${CYAN}[1/5] Verificando montagem do Google Drive...${NC}"
    if ! mountpoint -q "$PROJECT_DIR/mounts/gdrive"; then
        echo "  📂 Montando Google Drive via rclone..."
        mkdir -p "$PROJECT_DIR/mounts/gdrive"
        rclone mount gdrive: "$PROJECT_DIR/mounts/gdrive" --daemon --vfs-cache-mode writes 2>/dev/null
        sleep 1
    fi
    echo -e "  ${GREEN}✔ Google Drive montado em $PROJECT_DIR/mounts/gdrive${NC}"

    # 2. Restaurar Crontab
    echo -e "\n${CYAN}[2/5] Restaurando agendamentos do Crontab...${NC}"
    if [ -f "$CRON_BACKUP" ] && [ -s "$CRON_BACKUP" ]; then
        crontab "$CRON_BACKUP" 2>/dev/null
        echo -e "  ${GREEN}✔ Crontab restaurado com sucesso!${NC}"
    else
        echo -e "  ${YELLOW}⚠ Nenhum backup em $CRON_BACKUP encontrado. Crontab mantido atual.${NC}"
    fi

    # 3. Subir Container WhatsApp (OpenWA)
    echo -e "\n${CYAN}[3/5] Iniciando Container WhatsApp (openwa-api)...${NC}"
    if ! docker ps --format '{{.Names}}' | grep -q "^openwa-api$"; then
        docker start openwa-api > /dev/null 2>&1
        echo -e "  ${GREEN}✔ Container WhatsApp iniciado.${NC}"
    else
        echo -e "  ${GREEN}✔ Container WhatsApp já estava rodando.${NC}"
    fi

    # 4. Iniciar Webhook Zap (com proteções ativas)
    echo -e "\n${CYAN}[4/5] Iniciando Webhook Zap (:5000)...${NC}"
    if check_port 5000; then
        echo -e "  ${GREEN}✔ Webhook já está ativo na porta 5000.${NC}"
    else
        nohup "$VENV_PYTHON" "$PROJECT_DIR/webhook_zap.py" >> "$PROJECT_DIR/webhook_live.log" 2>&1 &
        sleep 2
        if check_port 5000; then
            echo -e "  ${GREEN}✔ Webhook Zap (:5000) online com proteção anti-eco ativa!${NC}"
        else
            echo -e "  ${YELLOW}⚠ Webhook iniciando em background (verifique $PROJECT_DIR/webhook_live.log).${NC}"
        fi
    fi

    # 5. Iniciar MCP Docente
    echo -e "\n${CYAN}[5/5] Iniciando MCP Docente & Auxiliares...${NC}"
    if ! pgrep -f "mcp_docente.py" > /dev/null; then
        nohup "$VENV_PYTHON" "$PROJECT_DIR/mcp_docente.py" > "$LOG_DIR/mcp_docente.log" 2>&1 &
        echo -e "  ${GREEN}✔ MCP Docente inicializado.${NC}"
    else
        echo -e "  ${GREEN}✔ MCP Docente já está rodando.${NC}"
    fi

    echo -e "\n${GREEN}${BOLD}======================================================${NC}"
    echo -e "${GREEN}${BOLD}🎉 ECOSSISTEMA ATIVADO E PROTEGIDO!${NC}"
    echo -e "${GREEN}${BOLD}Para checar status: ./kill_switch.sh status${NC}"
    echo -e "${GREEN}${BOLD}Para desligar tudo: ./kill_switch.sh off${NC}"
    echo -e "${GREEN}${BOLD}======================================================${NC}"
}

show_status() {
    echo -e "${BLUE}${BOLD}======================================================${NC}"
    echo -e "${BLUE}${BOLD}📊 PAINEL DE STATUS: DOCÊNCIA, BOT & LABORATÓRIOS${NC}"
    echo -e "${BLUE}${BOLD}======================================================${NC}"

    # 1. WhatsApp Container
    echo -n "• Container WhatsApp (openwa-api):  "
    if docker ps --format '{{.Names}}' | grep -q "^openwa-api$"; then
        echo -e "${GREEN}🟢 RODANDO${NC}"
    else
        echo -e "${RED}🔴 PARADO${NC}"
    fi

    # 2. Webhook Zap
    echo -n "• Webhook WhatsApp (:5000):          "
    if check_port 5000; then
        local pid_wh=$(pgrep -f "webhook_zap.py" | head -n1)
        echo -e "${GREEN}🟢 ONLINE${NC} (PID: ${pid_wh:-?})"
    else
        echo -e "${RED}🔴 DESLIGADO${NC}"
    fi

    # 3. Crontab
    echo -n "• Rotinas do Crontab:                "
    local cron_lines=$(crontab -l 2>/dev/null | grep -v '^#' | grep -c .)
    if [ "$cron_lines" -gt 0 ]; then
        echo -e "${GREEN}🟢 ATIVO${NC} ($cron_lines rotinas programadas)"
    else
        echo -e "${YELLOW}⏸️  PAUSADO${NC} (Kill Switch acionado ou vazio)"
    fi

    # 4. Servidor MCP Docente
    echo -n "• MCP Docente (Turmas/Notas):        "
    if pgrep -f "mcp_docente.py" > /dev/null; then
        echo -e "${GREEN}🟢 RODANDO${NC}"
    else
        echo -e "${RED}🔴 PARADO${NC}"
    fi

    # 5. OmniRoute Gateway
    echo -n "• OmniRoute LLM (:20128):            "
    if check_port 20128; then
        echo -e "${GREEN}🟢 RODANDO${NC}"
    else
        echo -e "${YELLOW}⚪ INATIVO${NC}"
    fi

    # 6. GDrive
    echo -n "• Google Drive (rclone mount):       "
    if mountpoint -q "$PROJECT_DIR/mounts/gdrive"; then
        echo -e "${GREEN}🟢 MONTADO${NC}"
    else
        echo -e "${RED}🔴 DESMONTADO${NC}"
    fi

    # 7. Banco de Dados Escola
    echo -n "• Banco escola.db (Laboratórios/Aulas):"
    if [ -f "$PROJECT_DIR/escola.db" ]; then
        local db_size=$(du -h "$PROJECT_DIR/escola.db" | cut -f1)
        echo -e " ${GREEN}🟢 OK${NC} ($db_size)"
    else
        echo -e " ${RED}🔴 NÃO ENCONTRADO${NC}"
    fi

    echo -e "${BLUE}${BOLD}======================================================${NC}"
    echo -e "Comandos rápidos:"
    echo -e "  ${RED}./kill_switch.sh off${NC}    -> Desliga tudo imediatamente"
    echo -e "  ${GREEN}./kill_switch.sh on${NC}     -> Liga tudo de forma coordenada"
    echo -e "  ${CYAN}./kill_switch.sh status${NC} -> Atualiza este painel"
    echo -e "${BLUE}${BOLD}======================================================${NC}"
}

case "$1" in
    off|stop|kill)
        kill_switch_off
        ;;
    on|start)
        kill_switch_on
        ;;
    status)
        show_status
        ;;
    restart)
        kill_switch_off
        sleep 2
        kill_switch_on
        ;;
    *)
        echo "Uso: $0 {on|off|status|restart}"
        echo "  $0 off    -> Desliga o container WhatsApp, pausa o crontab e mata os processos"
        echo "  $0 on     -> Sobe o container, restaura o crontab e inicia o webhook protegido"
        echo "  $0 status -> Exibe o status de todos os módulos"
        exit 1
        ;;
esac
