import json
import os
DONO_CHAT_ID = os.environ.get("DONO_CHAT_ID", "")  # chat do dono (ex.: 5599...@c.us)
import csv
import re
import unicodedata
import sqlite3
import base64
import subprocess
import shutil
import sys
import urllib.request
import urllib.error
import urllib.parse
import threading
import time
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from http.server import BaseHTTPRequestHandler, HTTPServer
try:
    from pptx import Presentation
except ImportError:
    Presentation = None

OPENWA_API_URL = os.environ.get("OPENWA_API_URL", "http://127.0.0.1:2785")
OPENWA_API_KEY = os.environ.get("OPENWA_API_KEY", "")
OPENWA_SESSION_ID = os.environ.get("OPENWA_SESSION_ID", "")
OMNIROUTE_URL = os.environ.get("OMNIROUTE_URL", "http://localhost:20128/v1/chat/completions")
OMNIROUTE_MODEL = os.environ.get("OMNIROUTE_MODEL", "no-think/agy/claude-sonnet-5-5-low")
# SEGURANÇA: Desabilita análise automática de imagens e fotos via OmniRoute no WhatsApp
DESABILITAR_LEITURA_IMAGEM_OMNIROUTE = os.environ.get("DESABILITAR_LEITURA_IMAGEM_OMNIROUTE", "1") == "1"
MODO_VPS_LEVE = os.environ.get("MODO_VPS_LEVE", "0") == "1"
HOST_LOCAL_WEBHOOK_URL = os.environ.get("HOST_LOCAL_WEBHOOK_URL", "").strip()

BASE_DIR = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas/automacao-iema")
AULAS_CSV = os.path.join(BASE_DIR, "aulas.csv")
NOTAS_CSV = os.path.join(BASE_DIR, "dados_notas.csv")
NOTAS_3B_CSV = os.path.join(BASE_DIR, "planilha_notas_3bimestre.csv")
FREQUENCIA_LOG_CSV = os.path.join(BASE_DIR, "frequencia_log.csv")
DIARIO_BORDO_CSV = os.path.join(BASE_DIR, "diario_bordo.csv")
LOCAL_LISTAS = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas_listas")
GDRIVE_LISTAS = os.path.join(os.environ.get("DOCENTE_HOME", "."), "mounts/gdrive/Aulas/Listas")
GDRIVE_SLIDES = os.path.join(os.environ.get("DOCENTE_HOME", "."), "mounts/gdrive/Aulas/Slides")
GDRIVE_IDEIAS = os.path.join(os.environ.get("DOCENTE_HOME", "."), "mounts/gdrive/Aulas/Notas_Ideias")
PROJETOS_IDEIAS_DIR = os.path.join(os.environ.get("DOCENTE_HOME", "."), "projetos_ideias")
ESCOLA_DB = os.environ.get("ESCOLA_DB", os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db"))
LOG_FILE = os.path.join(os.environ.get("DOCENTE_HOME", "."), "webhook_live.log")
import contato_restrito
import anti_loop
import ponte_agente
import gestor_docentes_rag

def verificar_host_local_online(timeout_s: float = 2.0) -> bool:
    """Verifica se o Host Local (PC principal) está acessível para receber tarefas pesadas."""
    if not HOST_LOCAL_WEBHOOK_URL:
        return False
    url_health = HOST_LOCAL_WEBHOOK_URL.rstrip("/") + "/health"
    try:
        req = urllib.request.Request(url_health, method="GET")
        with urllib.request.urlopen(req, timeout=timeout_s) as resp:
            return 200 <= resp.status < 300
    except Exception:
        return False

def delegar_ou_fallback_tarefa_pesada(chat_id: str, texto: str, data: dict = None, nome_recurso: str = "Processamento Pesado") -> str | None:
    """
    Quando `MODO_VPS_LEVE=1` está ativo na VPS de 1 GB RAM, intercepta comandos pesados
    (geração pesada de slides/LibreOffice, terminal local, protótipos multiagente).
    Se o Host Local estiver online, encaminha via HTTP; caso contrário, retorna fallback gracioso.
    """
    if not MODO_VPS_LEVE:
        return None
    if verificar_host_local_online():
        try:
            payload = json.dumps({"event": "message.received", "data": data or {"chatId": chat_id, "body": texto}}).encode("utf-8")
            req = urllib.request.Request(
                HOST_LOCAL_WEBHOOK_URL,
                data=payload,
                headers={"Content-Type": "application/json", "X-Forwarded-From-VPS": "1"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                if 200 <= resp.status < 300:
                    return f"☁️ *[VPS 1GB → Host Local]* Solicitação de *{nome_recurso}* encaminhada ao PC principal online."
        except Exception:
            pass
    return (
        f"☁️ *Modo VPS Leve (1 GB RAM) — Fallback Seguro*\n\n"
        f"O recurso *{nome_recurso}* exige processamento pesado no Host Local, que está desligado ou inacessível no momento.\n\n"
        f"✅ *Funções ativas 24h na VPS:*\n"
        f"• Agendamento e grade do laboratório (`!lab`)\n"
        f"• Grade individual de horários (`minha grade` / `!horario`)\n"
        f"• Consulta à base RAG por disciplina (`!rag <tema>`)\n"
        f"• Registro de diário, faltas e notas (`!diario`, `!falta`, `!turma`)"
    )

LIMITE_FALTAS = 15

os.makedirs(PROJETOS_IDEIAS_DIR, exist_ok=True)
os.makedirs(LOCAL_LISTAS, exist_ok=True)

def normalizar_texto(texto: str) -> str:
    if not texto:
        return ""
    texto = unicodedata.normalize("NFD", str(texto))
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto).strip().lower()

TIMESTAMP_INICIO_BOOT = time.time()

def _rotacionar_log_se_necessario(caminho: str = LOG_FILE, max_bytes: int = 5 * 1024 * 1024, backups: int = 3):
    try:
        if os.path.exists(caminho) and os.path.getsize(caminho) > max_bytes:
            for i in range(backups - 1, 0, -1):
                s = f"{caminho}.{i}"
                d = f"{caminho}.{i+1}"
                if os.path.exists(s):
                    shutil.move(s, d)
            if os.path.exists(caminho):
                shutil.move(caminho, f"{caminho}.1")
    except Exception:
        pass

def sanitizar_nome_arquivo(texto: str) -> str:
    t = normalizar_texto(texto).title().replace(" ", "_")
    return re.sub(r"[^a-zA-Z0-9_]", "", t)

def log_terminal(msg):
    linha = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(linha, flush=True)
    _rotacionar_log_se_necessario()

def limpar_emojis(texto: str) -> str:
    if not texto:
        return ""
    emoji_pattern = re.compile(
        "["
        "\U0001F600-\U0001F64F"
        "\U0001F300-\U0001F5FF"
        "\U0001F680-\U0001F6FF"
        "\U0001F1E0-\U0001F1FF"
        "\U00002702-\U000027B0"
        "\U000024C2-\U0001F251"
        "\U0001F900-\U0001F9FF"
        "\U0001FA00-\U0001FA6F"
        "\U0001FA70-\U0001FAFF"
        "\U00002600-\U000026FF"
        "]+",
        flags=re.UNICODE
    )
    limpo = emoji_pattern.sub("", str(texto))
    limpo = re.sub(r"[ \t]+", " ", limpo)
    limpo = re.sub(r"\n{3,}", "\n\n", limpo)
    return limpo.strip()

def _eh_apos_17h_ou_madrugada(dt: datetime = None) -> bool:
    """Retorna True se o horário atual for após as 17:00 (>= 17h) ou antes das 07:00 (< 7h)."""
    alvo = dt or datetime.now()
    return alvo.hour >= 17 or alvo.hour < 7

def _redirecionar_material_grupo_pos_17h(chat_id: str, texto: str = "", eh_documento: bool = False):
    """
    Regra de proteção pós-17h: evita o envio de materiais/conteúdos nos grupos após as 17h
    (e durante a madrugada até 07h), redirecionando tudo para o privado do Prof. Pedro.
    """
    cid = str(chat_id or "")
    if "@g.us" not in cid:
        return cid, texto

    t_norm = normalizar_texto(texto).lower() if texto else ""

    # No grupo da Base Técnica TI, consultas do laboratório (ex: grade semanal em texto e PDF) nunca são redirecionadas
    if cid == globals().get("ID_GRUPO_BASE_TECNICA", ""):
        if "grade_semanal" in t_norm or "grade" in t_norm or "laboratorio" in t_norm or "lab" in t_norm:
            return cid, texto

    if not _eh_apos_17h_ou_madrugada():
        return cid, texto

    eh_grupo_turma = cid in globals().get("TURMA_POR_GRUPO", {})
    marcadores_material = (
        "conteudo definido",
        "planejamento docente",
        "exercicios de caderno",
        "atividade cadastrada",
        "gabarito",
        "lista de exercicios",
        "slides",
        "apostila",
        "resumo matinal",
        "relatorio noturno",
    )
    t_norm = normalizar_texto(texto).lower() if texto else ""
    eh_msg_material = eh_documento or eh_grupo_turma or any(m in t_norm for m in marcadores_material)

    if eh_msg_material:
        destino_priv = globals().get("CHAT_PRIVADO_PADRAO", DONO_CHAT_ID)
        log_terminal(f"🔒 Regra pós-17h: envio em grupo ({cid}) redirecionado para o privado ({destino_priv}).")
        if texto and "[Privado Pós-17h]" not in texto:
            texto = f"[Privado Pós-17h]\n{texto}"
        return destino_priv, texto
    return cid, texto

# Mapa opcional {"<lid>@lid": "<numero>@c.us"}; defina em WHATSAPP_LID_MAP (JSON) no .env
MAPA_LID_PARA_CUS_ENVIO = json.loads(os.environ.get("WHATSAPP_LID_MAP", "{}"))

def _resolver_chat_id_envio_openwa(chat_id: str) -> str:
    """Converte LIDs conhecidos para @c.us na chamada HTTP ao OpenWA para evitar timeout do whatsapp-web.js em self-LID."""
    cid = str(chat_id or "").strip()
    if cid in MAPA_LID_PARA_CUS_ENVIO:
        return MAPA_LID_PARA_CUS_ENVIO[cid]
    return cid

def enviar_zap(chat_id, texto):
    chat_id, texto = _redirecionar_material_grupo_pos_17h(chat_id, texto, eh_documento=False)
    if os.environ.get("MODO_TESTE_ZAP") == "1":
        return True
    if str(chat_id) in globals().get("TURMA_POR_GRUPO", {}):
        texto = remover_dicas_comandos_para_grupo(texto)
    texto = limpar_emojis(texto)
    if not anti_loop.permitir_envio(chat_id, texto):  # disjuntor anti-loop (saída)
        return False
    if "_registrar_texto_enviado_bot" in globals():
        _registrar_texto_enviado_bot(texto)
    chat_id_api = _resolver_chat_id_envio_openwa(chat_id)
    url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/messages/send-text"
    req = urllib.request.Request(
        url,
        data=json.dumps({"chatId": chat_id_api, "text": texto}).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-api-key": OPENWA_API_KEY},
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            try:
                resp_json = json.loads(resp.read().decode("utf-8"))
                m_id = resp_json.get("messageId")
                if m_id:
                    MENSAGENS_PROCESSADAS[extrair_id_mensagem(m_id)] = time.time()
            except Exception:
                pass
            log_terminal(f"Resposta enviada para {chat_id} (Status {resp.status})")
            return True
    except Exception as e:
        log_terminal(f"Erro ao enviar resposta para {chat_id}: {e}")
        return False

def enviar_documento_zap(chat_id, caminho_arquivo, filename, mimetype):
    chat_id, _ = _redirecionar_material_grupo_pos_17h(chat_id, filename, eh_documento=True)
    if os.environ.get("MODO_TESTE_ZAP") == "1":
        return True
    chat_id_api = _resolver_chat_id_envio_openwa(chat_id)
    try:
        with open(caminho_arquivo, "rb") as f:
            conteudo_b64 = base64.b64encode(f.read()).decode("utf-8")
        url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/messages/send-document"
        payload = json.dumps({
            "chatId": chat_id_api,
            "base64": conteudo_b64,
            "mimetype": mimetype,
            "filename": filename
        }).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=payload,
            headers={"Content-Type": "application/json", "x-api-key": OPENWA_API_KEY},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            try:
                m_id = json.loads(resp.read().decode("utf-8")).get("messageId")
                if m_id:
                    MENSAGENS_PROCESSADAS[extrair_id_mensagem(m_id)] = time.time()
            except Exception:
                pass
            log_terminal(f"📎 Documento {filename} enviado para {chat_id} (Status {resp.status})")
            return True
    except Exception as e:
        log_terminal(f"❌ Erro ao enviar documento {filename}: {e}")
        return False

def consultar_llm(prompt: str) -> str:
    modelos = [OMNIROUTE_MODEL, "agy/gemini-3-flash", "cerebras/llama3.1-8b"]
    for mod in modelos:
        try:
            req = urllib.request.Request(
                OMNIROUTE_URL,
                data=json.dumps({
                    "model": mod,
                    "stream": False,
                    "messages": [{"role": "user", "content": prompt}]
                }).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                return data["choices"][0]["message"]["content"]
        except Exception as e_mod:
            log_terminal(f"Aviso consultar_llm com {mod}: {e_mod}")
            continue
    return ""

# --- GERAÇÃO EM SEGUNDO PLANO (WORKERS ASSÍNCRONOS) ---

def converter_md_para_txt_limpo(md_texto: str) -> str:
    # Converte markdown em texto puro e limpo para leitura
    texto = re.sub(r"^#+\s*(.*?)$", r"=== \1 ===", md_texto, flags=re.MULTILINE)
    texto = re.sub(r"\*\*(.*?)\*\*", r"\1", texto)
    texto = re.sub(r"\*(.*?)\*", r"\1", texto)
    texto = re.sub(r"`(.*?)`", r"\1", texto)
    return texto.strip()

def worker_gerar_exercicio(chat_id: str, tema: str, qtd: int = 5, serie: str = "Ensino Médio", foco: str = "geral", formato: str = "múltipla escolha"):
    log_terminal(f"Worker iniciado: Lista de exercícios sobre '{tema}' ({qtd} questões, {serie})...")
    prompt = (
        f"Gere uma lista de {qtd} exercícios pedagógicos e acessíveis para estudantes do {serie}.\n"
        f"Tema: '{tema}'.\n"
        f"Foco do conteúdo: {foco}.\n"
        f"Formato desejado: {formato}.\n\n"
        "Regras pedagógicas essenciais:\n"
        "- Linguagem simples, direta, didática para estudantes do Ensino Médio.\n"
        "- Questões práticas ou contextualizadas com o cotidiano sempre que possível.\n"
        "- Evite teorias excessivamente pesadas ou jargões desnecessários.\n"
        "- Sem emojis decorativos.\n"
        "- Estrutura obrigatória em Markdown:\n"
        f"# Lista de Exercícios: {tema}\n\n"
        "## Questões\n(Enunciados diretos e claros)\n\n"
        "## Gabarito Comentado\n(Resolução simplificada e explicativa passo a passo)"
    )
    try:
        conteudo_md = consultar_llm(prompt)
        nome_base = f"Lista_{sanitizar_nome_arquivo(tema)}"
        caminho_local_md = os.path.join(LOCAL_LISTAS, f"{nome_base}.md")
        caminho_local_pdf = os.path.join(LOCAL_LISTAS, f"{nome_base}.pdf")
        caminho_local_txt = os.path.join(LOCAL_LISTAS, f"{nome_base}.txt")

        # Salva o MD no diretório de trabalho local
        with open(caminho_local_md, "w", encoding="utf-8") as f:
            f.write(conteudo_md)

        # Localiza engine Weasyprint para gerar PDF
        weasy_bin = shutil.which("weasyprint") or "weasyprint"
        gerou_pdf = False

        if os.path.exists(weasy_bin):
            try:
                cmd = ["pandoc", caminho_local_md, "-o", caminho_local_pdf, f"--pdf-engine={weasy_bin}", f"--metadata=title:{tema}"]
                res = subprocess.run(cmd, capture_output=True, timeout=45)
                if res.returncode == 0 and os.path.exists(caminho_local_pdf) and os.path.getsize(caminho_local_pdf) > 200:
                    gerou_pdf = True
            except Exception as e_w:
                log_terminal(f"⚠️ Weasyprint falhou: {e_w}")

        if not gerou_pdf:
            # Fallback pdflatex
            try:
                res2 = subprocess.run(["pandoc", caminho_local_md, "-o", caminho_local_pdf], capture_output=True, timeout=45)
                if res2.returncode == 0 and os.path.exists(caminho_local_pdf) and os.path.getsize(caminho_local_pdf) > 200:
                    gerou_pdf = True
            except Exception:
                pass

        doc_ok = False
        if gerou_pdf:
            # PRIORIDADE 1: Envia formato PDF
            doc_ok = enviar_documento_zap(chat_id, caminho_local_pdf, f"{nome_base}.pdf", "application/pdf")
        else:
            # PRIORIDADE 2: Fallback exclusivo em formato TXT (texto puro) - NUNCA MD
            conteudo_txt = converter_md_para_txt_limpo(conteudo_md)
            with open(caminho_local_txt, "w", encoding="utf-8") as f_txt:
                f_txt.write(conteudo_txt)
            log_terminal(f"📄 PDF não gerado, enviando fallback em TXT limpo para {chat_id}...")
            doc_ok = enviar_documento_zap(chat_id, caminho_local_txt, f"{nome_base}.txt", "text/plain; charset=utf-8")

        # Cópia em background para o Google Drive se estiver montado
        try:
            if os.path.exists(GDRIVE_LISTAS):
                if gerou_pdf and os.path.exists(caminho_local_pdf):
                    shutil.copy2(caminho_local_pdf, os.path.join(GDRIVE_LISTAS, f"{nome_base}.pdf"))
                if os.path.exists(caminho_local_txt):
                    shutil.copy2(caminho_local_txt, os.path.join(GDRIVE_LISTAS, f"{nome_base}.txt"))
        except Exception as e_cp:
            log_terminal(f"⚠️ Aviso ao copiar lista para GDrive: {e_cp}")

        if not doc_ok:
            enviar_zap(chat_id, f"Lista de exercícios '{tema}' salva localmente em `{caminho_local_pdf if gerou_pdf else caminho_local_txt}`.")
    except Exception as e:
        log_terminal(f"❌ Erro no worker_gerar_exercicio: {e}")
        enviar_zap(chat_id, f"Erro ao gerar exercícios para '{tema}': {e}")

def worker_gerar_slides(chat_id: str, tema: str):
    log_terminal(f"📊 Worker iniciado: Slides sobre '{tema}'...")
    prompt = (
        f"Gere um roteiro estruturado para apresentação de slides sobre o tema: '{tema}'.\n"
        "Responda SOMENTE um JSON válido com a seguinte estrutura (sem blocos markdown extras):\n"
        "[\n"
        "  {\"titulo\": \"Objetivos da Aula\", \"pontos\": [\"Ponto 1\", \"Ponto 2\"]},\n"
        "  {\"titulo\": \"Fundamentação Teórica\", \"pontos\": [\"Ponto 1\", \"Ponto 2\"]},\n"
        "  {\"titulo\": \"Laboratório / Prática\", \"pontos\": [\"Passo 1\", \"Passo 2\"]},\n"
        "  {\"titulo\": \"Conclusão e Próximos Passos\", \"pontos\": [\"Ponto final\"]}\n"
        "]"
    )
    try:
        raw_resp = consultar_llm(prompt)
        slides_data = []
        try:
            # Tenta extrair array [...]
            json_match = re.search(r"\[.*\]", raw_resp, re.DOTALL)
            if json_match:
                slides_data = json.loads(json_match.group(0))
            else:
                # Tenta extrair objeto {...}
                obj_match = re.search(r"\{.*\}", raw_resp, re.DOTALL)
                if obj_match:
                    obj = json.loads(obj_match.group(0))
                    for k in ["slides", "presentation", "topics"]:
                        if k in obj and isinstance(obj[k], list):
                            slides_data = obj[k]
                            break
        except Exception:
            slides_data = []

        if not slides_data or not isinstance(slides_data, list):
            slides_data = [
                {"titulo": "Objetivos da Aula", "pontos": [f"Compreender os fundamentos de {tema}", "Aplicar na prática técnica"]},
                {"titulo": "Fundamentação Teórica", "pontos": ["Conceitos essenciais e arquitetura", "Padrões de mercado e melhores práticas"]},
                {"titulo": "Laboratório / Atividade Prática", "pontos": ["Configuração do ambiente", "Execução do exercício prático"]},
                {"titulo": "Conclusão", "pontos": ["Síntese dos tópicos abordados", "Próximos passos e referências"]}
            ]

        prs = Presentation()
        # Capa
        slide_capa = prs.slides.add_slide(prs.slide_layouts[0])
        slide_capa.shapes.title.text = tema
        slide_capa.placeholders[1].text = f"Material Didático • Gerado via Assistente Docente\nData: {datetime.now().strftime('%d/%m/%Y')}"

        # Slides de conteúdo
        for item in slides_data:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = item.get("titulo", "Tópico")
            tf = slide.placeholders[1].text_frame
            for i, pto in enumerate(item.get("pontos", [])):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = pto

        nome_arquivo = f"{sanitizar_nome_arquivo(tema)}.pptx"
        caminho_final = os.path.join(GDRIVE_SLIDES, nome_arquivo)
        caminho_local_tmp = f"/tmp/{nome_arquivo}"
        prs.save(caminho_local_tmp)
        # Sincroniza com Google Drive local e remoto
        try:
            shutil.copy(caminho_local_tmp, caminho_final)
        except Exception:
            pass
        try:
            subprocess.Popen(["rclone", "copy", caminho_local_tmp, "gdrive:Aulas/Slides/"])
        except Exception as e_rc:
            log_terminal(f"⚠️ Aviso rclone copy: {e_rc}")

        doc_enviado = enviar_documento_zap(chat_id, caminho_local_tmp, nome_arquivo, "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        if not doc_enviado:
            enviar_zap(chat_id, f"✅ Apresentação de slides para '{tema}' pronta e salva no Google Drive: Aulas > Slides (não foi possível anexar o arquivo diretamente).")
    except Exception as e:
        log_terminal(f"❌ Erro no worker_gerar_slides: {e}")
        enviar_zap(chat_id, f"⚠️ Erro ao gerar slides para '{tema}': {e}")

# --- GESTÃO DE NOTAS, IDEIAS E PROTOTIPAÇÃO MULTIAGENTE ---

def init_tabelas_ideias():
    if not os.path.exists(ESCOLA_DB):
        return
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS ideias_notas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    titulo TEXT,
                    descricao_original TEXT,
                    status TEXT DEFAULT 'PLANEJADA',
                    data_criacao TEXT,
                    diretorio_prototipo TEXT,
                    plano_etapas_md TEXT,
                    chat_id TEXT DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS tarefas_agentes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ideia_id INTEGER,
                    ordem INTEGER,
                    agente_responsavel TEXT,
                    titulo_tarefa TEXT,
                    descricao TEXT,
                    prompt_agente TEXT,
                    status TEXT DEFAULT 'PRONTO_PARA_EXECUCAO',
                    artefato_esperado TEXT,
                    FOREIGN KEY(ideia_id) REFERENCES ideias_notas(id)
                );
            """)
    except Exception as e:
        log_terminal(f"⚠️ Erro ao inicializar tabelas de ideias: {e}")

def processar_ideia_core(texto_ideia: str, chat_id: str = DONO_CHAT_ID) -> dict:
    """
    Estrutura uma ideia em passos práticos e inteligíveis para aplicação real na rotina docente,
    sem sugerir comandos de terminal ou códigos de programação.
    """
    init_tabelas_ideias()
    log_terminal(f"💡 Estruturando ideia pedagógica/prática: '{texto_ideia[:80]}'...")

    prompt = (
        "Você é um Consultor Pedagógico e Estrategista Educacional do IEMA.\n"
        "O usuário enviou uma ideia/nota sobre aulas, laboratório ou rotina docente. "
        "Seu objetivo é transformar essa ideia em uma proposta prática, facilmente inteligível e pronta para ser colocada em prática pelo professor.\n\n"
        "REGRAS OBRIGATÓRIAS:\n"
        "1. NÃO sugira comandos de terminal (como ssh, python, curl, bash, etc.).\n"
        "2. NÃO sugira códigos de programação ou scripts executáveis, pois você não conhece a arquitetura real do host.\n"
        "3. Forneça passos de ação claros, diretos e humanos para o dia a dia na escola/laboratório.\n\n"
        f"Ideia do Usuário:\n\"{texto_ideia}\"\n\n"
        "Sua resposta DEVE ser estritamente um único objeto JSON (sem formatação markdown, sem comentários) no formato:\n"
        "{\n"
        '  "titulo": "Título direto e objetivo da ideia",\n'
        '  "slug": "nome_curto_snake_case",\n'
        '  "conceito": "Explicação simples e inteligível do que é a ideia e como ela funciona na prática",\n'
        '  "passos_praticos": [\n'
        '    "Passo 1: ...",\n'
        '    "Passo 2: ...",\n'
        '    "Passo 3: ..."\n'
        '  ],\n'
        '  "beneficio": "Benefício real e direto para o professor ou alunos no dia a dia escolar"\n'
        "}"
    )

    try:
        raw_resp = consultar_llm(prompt)
        m = re.search(r'\{.*\}', raw_resp, re.DOTALL)
        if m:
            dados = json.loads(m.group(0))
        else:
            raise ValueError("JSON não encontrado na resposta do modelo.")
    except Exception as e:
        log_terminal(f"⚠️ Erro ao consultar LLM para ideia ({e}). Usando gerador estruturado de contingência.")
        slug_fb = sanitizar_nome_arquivo(texto_ideia[:30]).lower()
        dados = {
            "titulo": f"Proposta: {texto_ideia[:40]}...",
            "slug": slug_fb or "ideia_docente",
            "conceito": f"Aplicação prática da proposta: {texto_ideia}",
            "passos_praticos": [
                "Definir os objetivos e os recursos necessários para a atividade.",
                "Alinhar com a turma ou equipe pedagógica a dinâmica de aplicação.",
                "Executar a proposta em sala ou laboratório registrando observações.",
                "Avaliar o impacto e os resultados obtidos com os estudantes."
            ],
            "beneficio": "Melhoria direta no fluxo de trabalho docente e engajamento dos alunos."
        }

    titulo = dados.get("titulo", "Novo Projeto de Ideia").strip()
    slug = sanitizar_nome_arquivo(dados.get("slug", "ideia")).lower()
    conceito = dados.get("conceito", texto_ideia).strip()
    passos = dados.get("passos_praticos", [])
    beneficio = dados.get("beneficio", "").strip()

    data_criacao = datetime.now().strftime("%d/%m/%Y %H:%M")

    # Salva no banco de dados SQLite
    ideia_id = 1
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO ideias_notas (titulo, descricao_original, status, data_criacao, diretorio_prototipo, plano_etapas_md, chat_id)
                VALUES (?, ?, 'ESTRUTURADA', ?, '', '', ?)
            """, (titulo, texto_ideia, data_criacao, chat_id))
            conn.commit()
            ideia_id = cur.lastrowid
    except Exception as e:
        log_terminal(f"❌ Erro ao salvar ideia no banco: {e}")

    # Diretório para documentação da ideia
    caminho_dir = os.path.join(PROJETOS_IDEIAS_DIR, f"ideia_{ideia_id}_{slug}")
    os.makedirs(caminho_dir, exist_ok=True)

    # Monta documentação em Markdown
    plano_md_linhas = [
        f"# {titulo} (Ideia #{ideia_id})",
        f"\n**Data de Registro:** {data_criacao}",
        f"**Ideia Original:** {texto_ideia}",
        f"\n## Conceito e Aplicação\n{conceito}",
        f"\n## Passos Práticos para Execução\n"
    ]

    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            for idx, passo in enumerate(passos, 1):
                cur.execute("""
                    INSERT INTO tarefas_agentes (ideia_id, ordem, agente_responsavel, titulo_tarefa, descricao, prompt_agente, status, artefato_esperado)
                    VALUES (?, ?, 'Docente', ?, ?, '', 'PENDENTE', 'Ação Prática')
                """, (ideia_id, idx, f"Passo {idx}", str(passo)))
                plano_md_linhas.append(f"{idx}. {passo}")

            if beneficio:
                plano_md_linhas.append(f"\n## Benefício Esperado\n{beneficio}")

            plano_md_str = "\n".join(plano_md_linhas)
            cur.execute("UPDATE ideias_notas SET diretorio_prototipo = ?, plano_etapas_md = ? WHERE id = ?",
                        (caminho_dir, plano_md_str, ideia_id))
            conn.commit()
    except Exception as e_bd:
        log_terminal(f"❌ Erro ao registrar passos no banco: {e_bd}")
        plano_md_str = "\n".join(plano_md_linhas)

    # Salva README.md e ideia.json
    try:
        with open(os.path.join(caminho_dir, "README.md"), "w", encoding="utf-8") as f:
            f.write(plano_md_str)
        with open(os.path.join(caminho_dir, "ideia.json"), "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=2, ensure_ascii=False)
    except Exception as e_rd:
        log_terminal(f"⚠️ Erro ao salvar documentação no diretório: {e_rd}")

    # Backup assíncrono para o Google Drive via rclone
    try:
        subprocess.Popen(
            ["rclone", "copy", caminho_dir, f"gdrive:Aulas/Notas_Ideias/ideia_{ideia_id}_{slug}/"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except Exception as e_rc:
        log_terminal(f"⚠️ Aviso rclone copy ideias: {e_rc}")

    # Mensagem limpa e direta para o WhatsApp
    resp_usuario = [
        f"💡 *NOVA IDEIA ESTRUTURADA!* (Ideia `#{ideia_id}`)\n",
        f"📌 *{titulo}*\n",
        f"📖 *Conceito:*",
        f"{conceito}\n",
        f"📋 *Passos Práticos para Aplicar ({len(passos)}):*"
    ]
    for idx, passo in enumerate(passos, 1):
        resp_usuario.append(f"{idx}. {passo}")

    if beneficio:
        resp_usuario.append(f"\n✨ *Benefício Direto:*")
        resp_usuario.append(beneficio)

    resp_usuario.append("\n━━━━━━━━━━━━━━━━━━━━")
    resp_usuario.append("💡 *Comandos:*")
    resp_usuario.append(f"• `!ideias` -> Ver todas as ideias")
    resp_usuario.append(f"• `!ideias {ideia_id}` -> Detalhes desta ideia")
    resp_usuario.append(f"• `!ideias limpar {ideia_id}` -> Apagar esta ideia")

    return {
        "id": ideia_id,
        "titulo": titulo,
        "caminho_dir": caminho_dir,
        "mensagem": "\n".join(resp_usuario)
    }

def worker_processar_ideia(chat_id: str, texto_ideia: str):
    try:
        resultado = processar_ideia_core(texto_ideia, chat_id)
        enviar_zap(chat_id, resultado["mensagem"])
    except Exception as e:
        log_terminal(f"❌ Erro no worker_processar_ideia: {e}")
        enviar_zap(chat_id, f"⚠️ Erro ao desdobrar sua ideia: {e}")

def limpar_todas_ideias() -> str:
    """
    Remove todas as ideias e tarefas do banco de dados e limpa os arquivos locais e em nuvem.
    """
    init_tabelas_ideias()
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM ideias_notas")
            total = cur.fetchone()[0]
            cur.execute("DELETE FROM tarefas_agentes")
            cur.execute("DELETE FROM ideias_notas")
            conn.commit()

        # Limpa pasta local de ideias
        if os.path.exists(PROJETOS_IDEIAS_DIR):
            for item in os.listdir(PROJETOS_IDEIAS_DIR):
                item_path = os.path.join(PROJETOS_IDEIAS_DIR, item)
                try:
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path, ignore_errors=True)
                    else:
                        os.remove(item_path)
                except Exception as e_rem:
                    log_terminal(f"⚠️ Erro ao remover {item_path}: {e_rem}")

        # Limpa pasta no Google Drive de forma assíncrona
        try:
            subprocess.Popen(
                ["rclone", "delete", "gdrive:Aulas/Notas_Ideias/", "--rmdirs"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        except Exception as e_rc:
            log_terminal(f"⚠️ Aviso rclone delete ideias: {e_rc}")

        return f"🧹 *Limpeza concluída com sucesso!*\n\n• {total} ideias removidas do banco de dados.\n• Diretórios locais e backups no Google Drive limpos."
    except Exception as e:
        return f"❌ Erro ao limpar ideias: {e}"

def apagar_ideia_especifica(id_ou_termo: str) -> str:
    """
    Remove uma ideia específica e seus arquivos pelo ID ou título.
    """
    init_tabelas_ideias()
    t = id_ou_termo.strip()
    if not t:
        return "⚠️ Informe o ID da ideia para apagar. Exemplo: `!ideias limpar 1`"

    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            if t.isdigit():
                cur.execute("SELECT id, titulo, diretorio_prototipo FROM ideias_notas WHERE id = ?", (int(t),))
            else:
                cur.execute("SELECT id, titulo, diretorio_prototipo FROM ideias_notas WHERE id = ? OR titulo LIKE ? ORDER BY id DESC LIMIT 1", (t, f"%{t}%"))
            ideia = cur.fetchone()

            if not ideia:
                return f"🔍 Ideia '{t}' não encontrada para exclusão. Use `!ideias` para ver a lista."

            id_i, tit, cdir = ideia
            cur.execute("DELETE FROM tarefas_agentes WHERE ideia_id = ?", (id_i,))
            cur.execute("DELETE FROM ideias_notas WHERE id = ?", (id_i,))
            conn.commit()

        # Remove diretório local
        if cdir and os.path.exists(cdir):
            shutil.rmtree(cdir, ignore_errors=True)

        # Remove no Google Drive
        try:
            nome_pasta = os.path.basename(cdir) if cdir else f"ideia_{id_i}"
            subprocess.Popen(
                ["rclone", "purge", f"gdrive:Aulas/Notas_Ideias/{nome_pasta}"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        except Exception as e_rc:
            log_terminal(f"⚠️ Aviso rclone purge ideia {id_i}: {e_rc}")

        return f"🗑️ *Ideia #{id_i} ('{tit}') apagada com sucesso!*\nRegistros e arquivos removidos."
    except Exception as e:
        return f"❌ Erro ao apagar ideia: {e}"

def listar_ideias_notas() -> str:
    init_tabelas_ideias()
    if not os.path.exists(ESCOLA_DB):
        return "⚠️ Banco de dados escola.db não encontrado."
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT i.id, i.titulo, i.status, i.data_criacao, COUNT(t.id)
                FROM ideias_notas i
                LEFT JOIN tarefas_agentes t ON t.ideia_id = i.id
                GROUP BY i.id
                ORDER BY i.id DESC
            """)
            linhas = cur.fetchall()

        if not linhas:
            return (
                "💡 *Nenhuma ideia registrada no momento.*\n\n"
                "Para enviar uma nova ideia prática:\n"
                "• `!ideias <sua proposta pedagógica ou rotina>`\n"
                "Ex: `!ideias Dinâmica em duplas com desafios curtos no laboratório`"
            )

        resp = [f"💡 *IDEIAS PEDAGÓGICAS E DE ROTINA ({len(linhas)}):*\n"]
        for id_i, tit, st, dt, qtd_passos in linhas:
            resp.append(f"• `ID #{id_i}` | *{tit}*")
            resp.append(f"  📅 {dt} | 📊 Status: *{st}* | 📋 {qtd_passos} passos práticos\n")

        resp.append("━━━━━━━━━━━━━━━━━━━━")
        resp.append("💡 *Comandos rápidos:*")
        resp.append("• `!ideias <id>` -> Ver detalhes e passos práticos")
        resp.append("• `!ideias limpar <id>` -> Apagar ideia específica")
        resp.append("• `!ideias limpar tudo` -> Limpar todas as ideias")
        return "\n".join(resp)
    except Exception as e:
        return f"❌ Erro ao listar ideias: {e}"

def obter_detalhes_ideia(id_ou_termo: str) -> str:
    init_tabelas_ideias()
    if not os.path.exists(ESCOLA_DB):
        return "⚠️ Banco escola.db não encontrado."

    t = id_ou_termo.strip()
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            if t.isdigit():
                cur.execute("SELECT id, titulo, descricao_original, status, data_criacao, plano_etapas_md FROM ideias_notas WHERE id = ?", (int(t),))
            else:
                cur.execute("SELECT id, titulo, descricao_original, status, data_criacao, plano_etapas_md FROM ideias_notas WHERE titulo LIKE ? ORDER BY id DESC LIMIT 1", (f"%{t}%",))
            ideia = cur.fetchone()

            if not ideia:
                return f"🔍 Nenhuma ideia encontrada com '{t}'. Use `!ideias` para listar todas."

            id_i, tit, desc_orig, st, dt, plano_md = ideia
            cur.execute("""
                SELECT ordem, titulo_tarefa, descricao
                FROM tarefas_agentes
                WHERE ideia_id = ?
                ORDER BY ordem ASC
            """, (id_i,))
            passos = cur.fetchall()

        resp = [
            f"💡 *IDEIA #{id_i} — {tit.upper()}*",
            f"📅 Registrada em: {dt} | Status: *{st}*",
            f"📝 *Ideia Original:* \"{desc_orig}\"\n"
        ]

        if passos:
            resp.append(f"📋 *PASSOS PRÁTICOS PARA COLOCAR EM PRÁTICA ({len(passos)}):*")
            for ordem, titulo_tarefa, descricao in passos:
                resp.append(f"{ordem}. {descricao}")
        elif plano_md:
            resp.append(f"📄 *Plano Estruturado:*\n{plano_md[:600]}")

        resp.append("\n━━━━━━━━━━━━━━━━━━━━")
        resp.append(f"💡 Para apagar esta ideia: `!ideias limpar {id_i}`")
        return "\n".join(resp)
    except Exception as e:
        return f"❌ Erro ao buscar detalhes da ideia: {e}"

def executar_ideia_comando(id_ou_termo: str, chat_id: str):
    init_tabelas_ideias()
    t = id_ou_termo.strip()
    if not t:
        enviar_zap(chat_id, "Use: `!executar <id_ideia>` ou `!rodar <id_ideia>`\nConsulte os IDs com `!ideias`.")
        return

    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            if t.isdigit():
                cur.execute("SELECT id, titulo, status, diretorio_prototipo FROM ideias_notas WHERE id = ?", (int(t),))
            else:
                cur.execute("SELECT id, titulo, status, diretorio_prototipo FROM ideias_notas WHERE id = ? OR titulo LIKE ? ORDER BY id DESC LIMIT 1", (t, f"%{t}%"))
            ideia = cur.fetchone()

        if not ideia:
            enviar_zap(chat_id, f"🔍 Nenhuma ideia encontrada com '{t}'. Use `!ideias` para listar todas.")
            return

        id_i, tit, st, cdir = ideia
        if not cdir or not os.path.exists(cdir):
            enviar_zap(chat_id, f"⚠️ Diretório do protótipo não encontrado para a ideia #{id_i}.")
            return

        arquivos = [f for f in os.listdir(cdir) if f.endswith(".py") or f.endswith(".sh")]
        alvo = "prototipo_mvp.py" if "prototipo_mvp.py" in arquivos else (arquivos[0] if arquivos else None)

        if not alvo:
            enviar_zap(chat_id, f"⚠️ Nenhum script executável (.py ou .sh) encontrado em `{cdir}`.")
            return

        caminho_script = os.path.join(cdir, alvo)
        log_dir = os.path.join(os.environ.get("DOCENTE_HOME", "."), "logs_docente")
        os.makedirs(log_dir, exist_ok=True)
        log_saida = os.path.join(log_dir, f"ideia_{id_i}.log")

        cmd = f"python3 {caminho_script}" if alvo.endswith(".py") else f"bash {caminho_script}"
        proc = subprocess.Popen(
            cmd,
            shell=True,
            cwd=cdir,
            stdout=open(log_saida, "w", encoding="utf-8"),
            stderr=subprocess.STDOUT,
            preexec_fn=os.setpgrp
        )

        time.sleep(1.5)

        saida_preview = ""
        if os.path.exists(log_saida):
            with open(log_saida, "r", encoding="utf-8", errors="replace") as lf:
                linhas = [line.strip() for line in lf.readlines() if line.strip()]
                saida_preview = "\n".join(linhas[:8])

        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("UPDATE ideias_notas SET status = 'EM_EXECUCAO' WHERE id = ?", (id_i,))
            conn.commit()

        msg = [
            f"🚀 *EXECUTANDO PROTÓTIPO DA IDEIA #{id_i}*",
            f"📌 *Projeto:* {tit}",
            f"📁 *Pasta:* `{cdir}`",
            f"▶️ *Script:* `{alvo}` (PID: `{proc.pid}`)",
            f"🟢 *Status:* Rodando em background",
            f"📄 *Log:* `{log_saida}`"
        ]
        if saida_preview:
            msg.append(f"\n🖥️ *Saída Inicial:*\n```\n{saida_preview}\n```")
        msg.append("\n━━━━━━━━━━━━━━━━━━━━")
        msg.append(f"💡 Para monitorar os logs: `!cloud tail -n 20 {log_saida}`")
        msg.append(f"🛑 Para parar o processo: `!cloud kill -9 {proc.pid}`")

        enviar_zap(chat_id, "\n".join(msg))
    except Exception as e:
        log_terminal(f"❌ Erro ao executar ideia #{t}: {e}")
        enviar_zap(chat_id, f"❌ Erro ao executar ideia: {e}")

# --- FUNÇÕES DE SALA DE AULA (< 200ms) ---

def obter_contexto_aula_atual() -> dict:
    """
    Retorna o contexto completo da aula atual ou próxima de hoje:
    turma, sala, disciplina, horário, tempos e alerta contra troca de disciplina.
    """
    if not os.path.exists(ESCOLA_DB):
        return {}
    try:
        agora = datetime.now()
        wd = agora.weekday()
        dias_nomes = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"]
        dia_nome = dias_nomes[wd]
        hora_str = agora.strftime("%H:%M")

        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, hora_inicio, hora_fim, turma, sala, disciplina, COALESCE(tempos, '')
                FROM grade_horaria 
                WHERE dia_semana = ? 
                ORDER BY hora_inicio ASC
            """, (wd,))
            aulas = cur.fetchall()

        if not aulas:
            return {"tem_aula": False, "motivo": f"Nenhuma aula cadastrada na sua grade para {dia_nome}."}

        aula_em_andamento = None
        proxima_aula = None

        for a in aulas:
            id_a, h_ini, h_fim, turma, sala, disc, tempos = a
            if h_ini <= hora_str <= h_fim:
                aula_em_andamento = a
                break
            elif hora_str < h_ini and not proxima_aula:
                proxima_aula = a

        aula_alvo = aula_em_andamento or proxima_aula or aulas[-1]
        is_em_andamento = (aula_em_andamento is not None)
        id_a, h_ini, h_fim, turma, sala, disc, tempos = aula_alvo

        alerta_confusao = ""
        disc_norm = normalizar_texto(disc)
        if turma == "202":
            if "legislacao" in disc_norm or "direito" in disc_norm:
                alerta_confusao = (
                    "⚠️ *ATENÇÃO PROFESSOR (ANTI-CONFUSÃO):*\n"
                    "Hoje na Turma 202 a matéria é *LEGISLAÇÃO DE INFORMÁTICA*, e NÃO Inteligência Artificial!\n"
                    "_(I.A na 202 acontece apenas às quintas-feiras)_."
                )
            elif "inteligencia" in disc_norm or "ia" in disc_norm:
                alerta_confusao = (
                    "🎯 *ATENÇÃO PROFESSOR:*\n"
                    "Hoje na Turma 202 a aula é de *INTELIGÊNCIA ARTIFICIAL*!\n"
                    "_(Legislação na 202 é às segundas, terças e sextas)_."
                )

        return {
            "tem_aula": True,
            "em_andamento": is_em_andamento,
            "turma": turma,
            "sala": sala,
            "disciplina": disc,
            "horario": f"{h_ini} - {h_fim}",
            "tempos": tempos or f"{h_ini} - {h_fim}",
            "dia_nome": dia_nome,
            "alerta_confusao": alerta_confusao
        }
    except Exception as e:
        log_terminal(f"⚠️ Erro em obter_contexto_aula_atual: {e}")
        return {}

def consultar_onde_estou() -> str:
    ctx = obter_contexto_aula_atual()
    if not ctx or not ctx.get("tem_aula"):
        return "🏖️ *Nenhuma aula programada para hoje.*\nVerifique sua grade semanal com `!grade`."

    em_andamento = ctx.get("em_andamento")
    turma = ctx.get("turma")
    sala = ctx.get("sala")
    disc = ctx.get("disciplina")
    horario = ctx.get("horario")
    tempos = ctx.get("tempos")
    dia = ctx.get("dia_nome")
    alerta = ctx.get("alerta_confusao")

    status_rotulo = "🚨 *SUA AULA ATUAL (EM ANDAMENTO)*" if em_andamento else "⏳ *SUA PRÓXIMA AULA DE HOJE*"

    resp = [
        f"{status_rotulo} — {dia}\n",
        f"📍 *Sala:* *{sala}*",
        f"👥 *Turma:* *Turma {turma}*",
        f"⏰ *Horário:* {horario} ({tempos})",
        f"📖 *Matéria:* *{disc}*\n"
    ]

    if alerta:
        resp.append(alerta + "\n")

    resp.append("━━━━━━━━━━━━━━━━━━━━")
    resp.append("💡 *Ações Rápidas:*")
    resp.append(f"• `!diario <conteúdo>` -> Registra direto em {turma} ({disc})")
    resp.append(f"• `!turma {turma}` -> Ver chamada da turma")
    resp.append(f"• `!pendencias {turma}` -> Ver vistos pendentes")
    return "\n".join(resp)

def obter_turma_atual_ou_proxima() -> str:
    """Retorna a turma da aula atual ou próxima de hoje se houver."""
    ctx = obter_contexto_aula_atual()
    return ctx.get("turma", "")

def listar_turmas() -> str:
    if not os.path.exists(ESCOLA_DB):
        return "⚠️ Banco escola.db não encontrado."
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT turma, COUNT(*), GROUP_CONCAT(DISTINCT componente) 
                FROM alunos 
                GROUP BY turma 
                ORDER BY turma ASC
            """)
            linhas = cur.fetchall()
        
        if not linhas:
            return "Nenhuma turma cadastrada."
            
        turma_atual = obter_turma_atual_ou_proxima()
        resp = ["🏫 *TURMAS E QUANTIDADE DE ALUNOS:*\n"]
        for t, qtd, comps in linhas:
            marcador = " 👈 *(Sua aula hoje)*" if t == turma_atual else ""
            resp.append(f"• *Turma {t}*: {qtd} alunos{marcador}\n  📖 {comps}")
            
        resp.append("\n━━━━━━━━━━━━━━━━━━━━")
        resp.append("💡 *Comandos para a Aula:*")
        resp.append("• `!turma <turma>` -> Ver lista numerada da chamada")
        resp.append("• `!falta <turma> <nº ou nome>`")
        resp.append("• `!presenca <turma> <nº ou nome>`")
        resp.append("• `!faltas <turma> <n1, n2, n3...>` (chamada rápida em lote)")
        return "\n".join(resp)
    except Exception as e:
        return f"❌ Erro ao listar turmas: {e}"

def listar_alunos_turma(turma_arg: str = "") -> str:
    if not os.path.exists(ESCOLA_DB):
        return "⚠️ Banco escola.db não encontrado."
    
    t_alvo = turma_arg.strip()
    aviso_inferido = ""
    if not t_alvo:
        t_alvo = obter_turma_atual_ou_proxima()
        if t_alvo:
            aviso_inferido = " *(detectada automaticamente pela aula de hoje)*"
        else:
            return "⚠️ Especifique a turma (ex: `!turma 202`). Digite `!turmas` para ver a lista."
    
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("SELECT matricula, nome, turma, componente, faltas FROM alunos WHERE turma = ? ORDER BY nome ASC", (t_alvo,))
            alunos = cur.fetchall()
        
        if not alunos:
            return f"🔍 Nenhuma turma '{t_alvo}' encontrada. Turmas disponíveis: 200, 201, 202, 203."

        comps = alunos[0][3] if alunos else ""
        hoje_str = datetime.now().strftime("%d/%m/%Y")
        total_faltas_turma = sum(a[4] for a in alunos)

        resp = [
            f"📋 *CHAMADA / ALUNOS — TURMA {t_alvo}*{aviso_inferido} ({len(alunos)} alunos)",
            f"📖 Disciplina: {comps}",
            f"📅 Data: {hoje_str} | Total de Faltas Acumuladas: {total_faltas_turma}\n"
        ]

        for idx, (mat, nome, turma, comp, faltas) in enumerate(alunos, 1):
            sinal_falta = f"⚠️ *{faltas} falta(s)*" if faltas > 0 else "0 faltas"
            resp.append(f"`{idx:2d}.` *{nome}* ({sinal_falta})")

        resp.append("\n━━━━━━━━━━━━━━━━━━━━")
        resp.append("💡 *Lançar Frequência na Aula:*")
        resp.append(f"• Pelo número: `!falta {t_alvo} 6` ou `!presenca {t_alvo} 6`")
        resp.append(f"• Chamada rápida (vários alunos): `!faltas {t_alvo} 2, 6, 14`")
        resp.append(f"• Por nome: `!falta {t_alvo} Carlos` ou `!falta Carlos Daniel`")
        return "\n".join(resp)
    except Exception as e:
        return f"❌ Erro ao listar alunos da turma: {e}"

def alterar_frequencia_aluno(identificador: str, turma_arg: str = "", data_str: str = "", operacao: str = "falta") -> str:
    if not os.path.exists(ESCOLA_DB):
        return "⚠️ Banco escola.db não encontrado."

    data_registro = data_str if data_str else datetime.now().strftime("%d/%m/%Y")
    ident = identificador.strip()
    turma_alvo = turma_arg.strip() if turma_arg else ""

    if not turma_alvo and ident.isdigit() and int(ident) <= 50:
        turma_alvo = obter_turma_atual_ou_proxima()
        if not turma_alvo:
            return f"⚠️ Você informou o número #{ident}, mas é necessário informar a turma (ex: `!{operacao} 202 {ident}`)."

    aluno_achado = None
    with sqlite3.connect(ESCOLA_DB) as conn:
        cur = conn.cursor()
        if turma_alvo and ident.isdigit() and int(ident) <= 50:
            cur.execute("SELECT id, matricula, nome, turma, faltas FROM alunos WHERE turma = ? ORDER BY nome ASC", (turma_alvo,))
            lista_turma = cur.fetchall()
            idx = int(ident) - 1
            if 0 <= idx < len(lista_turma):
                aluno_achado = lista_turma[idx]
            else:
                return f"❌ Número #{ident} inválido para a Turma {turma_alvo} (possui {len(lista_turma)} alunos)."
        elif turma_alvo:
            cur.execute("SELECT id, matricula, nome, turma, faltas FROM alunos WHERE turma = ? ORDER BY nome ASC", (turma_alvo,))
            cands = []
            palavras = normalizar_texto(ident).split()
            for row in cur.fetchall():
                n_norm = normalizar_texto(row[2])
                m_norm = row[1].lower()
                if ident.lower() == m_norm or all(p in n_norm for p in palavras):
                    cands.append(row)
            if len(cands) == 1:
                aluno_achado = cands[0]
            elif len(cands) > 1:
                nomes_dup = "\n".join([f"• {r[2]} ({r[1]})" for r in cands[:5]])
                return f"⚠️ Mais de um aluno encontrado na Turma {turma_alvo} para '{ident}':\n{nomes_dup}\nSeja mais específico ou use o número da chamada."
        else:
            cur.execute("SELECT id, matricula, nome, turma, faltas FROM alunos ORDER BY nome ASC")
            cands = []
            palavras = normalizar_texto(ident).split()
            for row in cur.fetchall():
                n_norm = normalizar_texto(row[2])
                m_norm = row[1].lower()
                if ident.lower() == m_norm or all(p in n_norm for p in palavras):
                    cands.append(row)
            if len(cands) == 1:
                aluno_achado = cands[0]
            elif len(cands) > 1:
                nomes_dup = "\n".join([f"• {r[2]} (Turma {r[3]})" for r in cands[:5]])
                return f"⚠️ Aluno ambíguo para '{ident}'. Encontrado em múltiplas turmas:\n{nomes_dup}\nEspecifique a turma: `!{operacao} <turma> <nome>`"

        if not aluno_achado:
            msg_t = f" na Turma {turma_alvo}" if turma_alvo else ""
            return f"🔍 Aluno '{ident}' não encontrado{msg_t}."

        id_db, mat, nome, turma_res, faltas_atuais = aluno_achado
        novas_faltas = faltas_atuais + 1 if operacao == "falta" else max(0, faltas_atuais - 1)
        cur.execute("UPDATE alunos SET faltas = ? WHERE id = ?", (novas_faltas, id_db))
        conn.commit()

    if os.path.exists(NOTAS_CSV):
        try:
            linhas = []
            achou_csv = False
            with open(NOTAS_CSV, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                fieldnames = reader.fieldnames or ["Matricula", "Nome", "Nota", "Faltas"]
                for row in reader:
                    if row.get("Matricula", "").strip() == mat:
                        row["Faltas"] = str(novas_faltas)
                        achou_csv = True
                    linhas.append(row)
            if not achou_csv:
                linhas.append({"Matricula": mat, "Nome": nome, "Nota": "--", "Faltas": str(novas_faltas)})
            with open(NOTAS_CSV, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(linhas)
        except Exception as e:
            log_terminal(f"⚠️ Erro ao sincronizar dados_notas.csv: {e}")

    try:
        log_existe = os.path.exists(FREQUENCIA_LOG_CSV)
        with open(FREQUENCIA_LOG_CSV, "a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["timestamp", "data", "matricula", "nome", "operacao", "faltas_total"])
            if not log_existe or os.path.getsize(FREQUENCIA_LOG_CSV) == 0:
                writer.writeheader()
            writer.writerow({
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "data": data_registro,
                "matricula": mat,
                "nome": nome,
                "operacao": "Falta registrada" if operacao == "falta" else "Presença/Abono registrado",
                "faltas_total": str(novas_faltas)
            })
    except Exception as e:
        log_terminal(f"⚠️ Erro ao escrever frequencia_log.csv: {e}")

    pct = int((novas_faltas / LIMITE_FALTAS) * 100)
    status_label = "Crítico (Reprovação)" if novas_faltas >= LIMITE_FALTAS else ("Alerta" if novas_faltas >= LIMITE_FALTAS * 0.75 else "Regular")
    op_label = "Falta registrada" if operacao == "falta" else "Presença/Abono aplicado"

    return (
        f"📝 *[Frequência Atualizada]*\n"
        f"• Turma: *{turma_res}*\n"
        f"• Aluno: *{nome}* (`{mat}`)\n"
        f"• Ação: *{op_label}* (Data: {data_registro})\n"
        f"• Total de Faltas: *{novas_faltas}* / {LIMITE_FALTAS} ({pct}% do limite)\n"
        f"• Status: {status_label}"
    )

def alterar_frequencia_lote(texto_resto: str, operacao: str = "falta") -> str:
    texto_resto = texto_resto.strip()
    if not texto_resto:
        turma_atual = obter_turma_atual_ou_proxima() or "202"
        return f"Use: `!{operacao}s [turma] <n1, n2, n3...>` (Ex: `!{operacao}s {turma_atual} 3, 6, 12` ou `!{operacao}s 3, 6, 12`)"

    partes = texto_resto.split(" ", 1)
    TURMAS_VALIDAS = {"101", "102", "103", "200", "201", "202", "203", "300", "301", "302", "303"}

    if partes[0] in TURMAS_VALIDAS and len(partes) > 1:
        turma = partes[0].strip()
        itens_str = partes[1].replace(";", ",").replace(" ", ",")
    else:
        turma = obter_turma_atual_ou_proxima() or "202"
        itens_str = texto_resto.replace(";", ",").replace(" ", ",")

    itens = [item.strip() for item in itens_str.split(",") if item.strip()]
    if not itens:
        return f"Nenhum aluno especificado. Use: `!{operacao}s {turma} 1, 2, 3`"

    sucessos = []
    erros = []
    alertas_criticos = []
    data_hoje = datetime.now().strftime("%d/%m/%Y")

    for item in itens:
        res = alterar_frequencia_aluno(item, turma, data_hoje, operacao)
        if "Frequência Atualizada" in res:
            m_nome = re.search(r"Aluno: \*(.*?)\*", res)
            m_total = re.search(r"Total de Faltas: \*(.*?)\*", res)
            nome_s = m_nome.group(1) if m_nome else item
            total_s = m_total.group(1) if m_total else "?"
            try:
                faltas_num = int(total_s)
                if operacao == "falta" and faltas_num >= 15:
                    alertas_criticos.append(f"🚨 *{nome_s}* atingiu *{faltas_num} faltas* (limiar crítico)!")
            except Exception:
                pass
            sucessos.append(f"• #{item} *{nome_s}* ({total_s} faltas)")
        else:
            erros.append(f"• Item '{item}': {res.splitlines()[0]}")

    op_nome = "Faltas Lançadas" if operacao == "falta" else "Presenças/Abonos Lançados"
    resp = [f"📋 *CHAMADA EM LOTE — TURMA {turma}* ({op_nome})\n📅 Data: {data_hoje}\n"]
    if sucessos:
        resp.append(f"✅ *{len(sucessos)} registro(s) atualizado(s) com sucesso:*")
        resp.extend(sucessos)
        resp.append("\n💾 Sincronizado no banco escolar e boletins.")
    if alertas_criticos:
        resp.append("\n⚠️ *ATENÇÃO PARA LIMITES DE FALTAS:*")
        resp.extend(alertas_criticos)
    if erros:
        resp.append(f"\n⚠️ *{len(erros)} erro(s):*")
        resp.extend(erros)
    
    return "\n".join(resp)

def alterar_frequencia(nome_busca: str, data_str: str, operacao: str) -> str:
    return alterar_frequencia_aluno(nome_busca, "", data_str, operacao)

# --- GESTÃO DE ATIVIDADES E VISTOS ---

def normalizar_fonetico(s: str) -> str:
    if not s:
        return ""
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    s = re.sub(r"ph", "p", s)
    s = re.sub(r"th", "t", s)
    s = re.sub(r"y", "i", s)
    s = re.sub(r"w", "v", s)
    s = re.sub(r"k", "c", s)
    s = re.sub(r"(.)\1+", r"\1", s)
    return s.strip()

def init_tabelas_atividades():
    if not os.path.exists(ESCOLA_DB):
        return
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS atividades_cadastradas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_passada TEXT,
                    data_hora_iso TEXT,
                    turma TEXT,
                    componente TEXT,
                    titulo TEXT,
                    descricao TEXT DEFAULT ''
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS atividades_entregas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    atividade_id INTEGER,
                    data_registro TEXT,
                    data_hora_iso TEXT,
                    turma TEXT,
                    componente TEXT,
                    matricula TEXT,
                    nome_aluno TEXT,
                    atividade TEXT,
                    status TEXT DEFAULT 'VISTO (OK)',
                    observacao TEXT
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS alunos_whatsapp_vinculos (
                    numero TEXT PRIMARY KEY,
                    matricula TEXT,
                    nome TEXT,
                    turma TEXT
                );
            """)
            try:
                cur.execute("ALTER TABLE atividades_entregas ADD COLUMN atividade_id INTEGER")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE atividades_entregas ADD COLUMN foto_caminho TEXT")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE atividades_entregas ADD COLUMN avaliacao_ia TEXT")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE atividades_cadastradas ADD COLUMN prof_chave TEXT DEFAULT 'pedro'")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE atividades_entregas ADD COLUMN prof_chave TEXT DEFAULT 'pedro'")
            except Exception:
                pass
            conn.commit()
    except Exception as e:
        log_terminal(f"⚠️ Erro ao inicializar tabelas de atividades: {e}")

def cadastrar_nova_atividade(texto: str, prof_chave: str = "pedro") -> str:
    """
    Cadastra uma nova atividade passada para a turma com a data em que foi passada.
    Exemplos:
    • 'atv nova 202 IA Exercício 1 - Redes Neurais'
    • 'passar 202 IA Lista 2 14/10'
    • 'atv nova 202 Questionário de Algoritmos'
    """
    init_tabelas_atividades()
    t = re.sub(r'^(nova|criar|cadastrar|passar|atividade)\s*', '', texto.strip(), flags=re.IGNORECASE).strip()
    if not t or t.lower() in ["ajuda", "help", "?"]:
        return (
            "📌 *CADASTRAR ATIVIDADE PASSADA*\n\n"
            "Use: `!atv nova <turma> [disciplina] <título da atividade> [data]`\n"
            "Ex: `!atv nova 202 IA Exercício 1 - Redes Neurais`\n"
            "Ex: `!atv nova 202 Redes Lista 2 14/10/2026`\n"
            "Ex: `!passar 202 IA Questionário Algoritmos`\n\n"
            "💡 Ao cadastrar, você define a data em que foi passada e poderá cobrar com `!cobrar <turma>`."
        )

    tokens = t.split()
    TURMAS_VALIDAS = {"101", "102", "103", "200", "201", "202", "203", "300", "301", "302", "303"}
    COMP_MAP = {
        "ia": "INTELIGÊNCIA ARTIFICIAL",
        "inteligencia": "INTELIGÊNCIA ARTIFICIAL",
        "inteligência": "INTELIGÊNCIA ARTIFICIAL",
        "leg": "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL",
        "legislacao": "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL",
        "legislação": "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL",
        "redes": "REDES DE COMPUTADORES",
        "web": "DESENVOLVIMENTO WEB",
        "design": "DESENVOLVIMENTO E DESIGN DE WEBSITES",
        "mkt": "MARKETING DIGITAL",
        "marketing": "MARKETING DIGITAL",
        "mkt digital": "MARKETING DIGITAL",
        "estrategia": "Estratégia de Marketing",
        "estratégia": "Estratégia de Marketing",
        "legislacao mercadologica": "Legislação Mercadológica",
        "legislação mercadológica": "Legislação Mercadológica",
        "intro mkt": "INTRODUÇÃO AO MARKETING",
        "introducao ao marketing": "INTRODUÇÃO AO MARKETING"
    }

    turma_alvo = ""
    comp_alvo = ""
    data_passada = ""
    tokens_rest = []

    for tok in tokens:
        tok_clean = tok.strip(",;.")
        if tok_clean in TURMAS_VALIDAS and not turma_alvo:
            turma_alvo = tok_clean
            continue
        tok_low = normalizar_texto(tok_clean)
        if tok_low in COMP_MAP and not comp_alvo:
            comp_alvo = COMP_MAP[tok_low]
            continue
        if re.match(r'^\d{1,2}/\d{1,2}(?:/\d{2,4})?$', tok_clean) and not data_passada:
            d_parts = tok_clean.split("/")
            dia = d_parts[0].zfill(2)
            mes = d_parts[1].zfill(2)
            ano = d_parts[2] if len(d_parts) > 2 else str(datetime.now().year)
            if len(ano) == 2:
                ano = f"20{ano}"
            data_passada = f"{dia}/{mes}/{ano}"
            continue
        tokens_rest.append(tok)

    if not turma_alvo:
        turma_atual = obter_turma_atual_ou_proxima()
        if prof_chave == "bruna":
            turma_alvo = turma_atual if turma_atual in ("101", "201") else "201"
        elif prof_chave == "marcos":
            turma_alvo = turma_atual if turma_atual in ("102", "103", "203") else "102"
        else:
            turma_alvo = turma_atual or "202"

    if not comp_alvo:
        if prof_chave == "bruna":
            comp_alvo = "INTRODUÇÃO AO MARKETING" if turma_alvo == "101" else "MARKETING DIGITAL"
        elif prof_chave == "marcos":
            comp_alvo = "PROGRAMAÇÃO PARA WEB" if turma_alvo == "203" else "INTRODUÇÃO À PROGRAMAÇÃO WEB"
        else:
            comp_alvo = "INTELIGÊNCIA ARTIFICIAL"

    if not data_passada:
        data_passada = datetime.now().strftime("%d/%m/%Y")

    titulo = " ".join(tokens_rest).strip()
    if not titulo:
        titulo = f"Atividade de {comp_alvo}"

    agora = datetime.now()
    data_iso = agora.strftime("%Y-%m-%d %H:%M:%S")

    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO atividades_cadastradas (data_passada, data_hora_iso, turma, componente, titulo, descricao, prof_chave)
                VALUES (?, ?, ?, ?, ?, '', ?)
            """, (data_passada, data_iso, turma_alvo, comp_alvo, titulo, prof_chave or "pedro"))
            conn.commit()
            ativ_id = cur.lastrowid

        return (
            f"📌 *NOVA ATIVIDADE CADASTRADA COM SUCESSO!* (ID: #{ativ_id})\n\n"
            f"🏫 *Turma:* {turma_alvo}\n"
            f"📖 *Componente:* {comp_alvo}\n"
            f"📝 *Título:* {titulo}\n"
            f"📅 *Data em que foi passada:* {data_passada}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"💡 *Para dar vistos aos alunos:*\n"
            f"`!atv <nº ou nome> {turma_alvo} ok`\n\n"
            f"📊 *Para cobrar a turma e ver pendentes:*\n"
            f"`!cobrar {turma_alvo}`"
        )
    except Exception as e:
        return f"❌ Erro ao cadastrar atividade: {e}"

TURMAS_VALIDAS_ATV = {"101", "102", "103", "200", "201", "202", "203", "300", "301", "302", "303"}
COMP_MAP_ATV = {
    "ia": "INTELIGÊNCIA ARTIFICIAL",
    "inteligencia": "INTELIGÊNCIA ARTIFICIAL",
    "inteligência": "INTELIGÊNCIA ARTIFICIAL",
    "leg": "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL",
    "legislacao": "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL",
    "legislação": "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL",
    "redes": "REDES DE COMPUTADORES",
    "web": "DESENVOLVIMENTO WEB",
    "design": "DESENVOLVIMENTO E DESIGN DE WEBSITES",
    "mkt": "MARKETING DIGITAL",
    "marketing": "MARKETING DIGITAL",
    "mkt digital": "MARKETING DIGITAL",
    "estrategia": "Estratégia de Marketing",
    "estratégia": "Estratégia de Marketing",
    "legislacao mercadologica": "Legislação Mercadológica",
    "legislação mercadológica": "Legislação Mercadológica",
    "intro mkt": "INTRODUÇÃO AO MARKETING",
    "introducao ao marketing": "INTRODUÇÃO AO MARKETING"
}

def _comp_equiv(c1: str, c2: str) -> bool:
    if not c1 or not c2:
        return False
    n1 = normalizar_texto(c1)
    n2 = normalizar_texto(c2)
    if n1 in n2 or n2 in n1:
        return True
    for kw in ("legislacao", "inteligencia artificial", "redes", "design", "web", "marketing", "estrategia", "mercadologica"):
        if kw in n1 and kw in n2:
            return True
    return False

def _contar_vistos_atividade(cur, turma: str, ativ_id: int, ativ_titulo: str = "", prof_chave: str = "") -> int:
    if prof_chave:
        cur.execute("""
            SELECT DISTINCT matricula, nome_aluno, status
            FROM atividades_entregas
            WHERE turma = ? AND (atividade_id = ? OR (? != '' AND atividade = ?))
              AND prof_chave = ?
        """, (turma, ativ_id, ativ_titulo, ativ_titulo, prof_chave))
    else:
        cur.execute("""
            SELECT DISTINCT matricula, nome_aluno, status
            FROM atividades_entregas
            WHERE turma = ? AND (atividade_id = ? OR (? != '' AND atividade = ?))
        """, (turma, ativ_id, ativ_titulo, ativ_titulo))
    vistos = set()
    for mat, nom, st in cur.fetchall():
        if (st or "").upper() in ["VISTO (OK)", "OK", "ENTREGUE", "FEITO", "10"]:
            vistos.add(mat or normalizar_fonetico(nom))
    return len(vistos)

def _resumo_atividades_turma_str(cur, turma: str, ativ_id_destaque: int = None, prof_chave: str = "") -> str:
    cur.execute("SELECT COUNT(*) FROM alunos WHERE turma = ?", (turma,))
    tot_alunos = (cur.fetchone() or [0])[0]
    if prof_chave:
        cur.execute("""
            SELECT id, data_passada, componente, titulo
            FROM atividades_cadastradas
            WHERE turma = ? AND prof_chave = ?
            ORDER BY id DESC LIMIT 6
        """, (turma, prof_chave))
    else:
        cur.execute("""
            SELECT id, data_passada, componente, titulo
            FROM atividades_cadastradas
            WHERE turma = ?
            ORDER BY id DESC LIMIT 6
        """, (turma,))
    rows = cur.fetchall()
    if not rows:
        return ""
    linhas = [f"📌 *Atividades da Turma {turma}:*"]
    for aid, dt, comp, tit in rows:
        qtd = _contar_vistos_atividade(cur, turma, aid, tit, prof_chave=prof_chave)
        seta = "👉 " if aid == ativ_id_destaque else "• "
        tit_curto = tit if len(tit) <= 48 else tit[:45] + "..."
        linhas.append(f"{seta}*#{aid}* ({dt}) {tit_curto} — *{qtd}/{tot_alunos} vistos*")
    return "\n".join(linhas)

def _extrair_seletor_atividade(tokens: list):
    """
    Extrai indicadores explícitos da atividade a partir dos tokens do comando:
    • #6, id:6, id=6, atv6, atv#6 -> ativ_id_expl
    • 19/10, 20/10/2026, hoje, ontem -> data_expl
    • tema:lgpd, ativ:prompt -> termo_titulo
    Retorna (ativ_id_expl, data_expl, termo_titulo, tokens_restantes)
    """
    ativ_id_expl = None
    data_expl = ""
    termo_titulo = ""
    rest = []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        tok_clean = tok.strip(",;.")
        tok_low = normalizar_texto(tok_clean)

        # #6, id:6, id=6, atv6, atv#6
        m_id = re.match(r"^(?:#|id[:=]?|atv#?|atividade#?)(\d+)$", tok_low)
        if m_id and ativ_id_expl is None:
            val = int(m_id.group(1))
            if val not in (200, 201, 202, 203, 300, 301, 302, 303):
                ativ_id_expl = val
                i += 1
                continue

        # 'id 6' ou 'atividade 6' (dois tokens)
        if tok_low in ("id", "atividade", "atv") and i + 1 < len(tokens):
            prox = tokens[i + 1].strip(",;.#")
            if prox.isdigit() and int(prox) not in (200, 201, 202, 203, 300, 301, 302, 303) and ativ_id_expl is None:
                ativ_id_expl = int(prox)
                i += 2
                continue

        # Data explícita DD/MM ou DD/MM/AAAA
        if re.match(r"^\d{1,2}/\d{1,2}(?:/\d{2,4})?$", tok_clean) and not data_expl:
            partes_d = tok_clean.split("/")
            data_expl = f"{partes_d[0].zfill(2)}/{partes_d[1].zfill(2)}"
            i += 1
            continue

        if tok_low in ("hoje", "ontem") and not data_expl:
            dt_ref = datetime.now() if tok_low == "hoje" else (datetime.now() - timedelta(days=1))
            data_expl = dt_ref.strftime("%d/%m")
            i += 1
            continue

        m_tema = re.match(r"^(?:tema|titulo|ativ)[:=](.+)$", tok_low)
        if m_tema and not termo_titulo:
            termo_titulo = m_tema.group(1).strip()
            i += 1
            continue

        rest.append(tok)
        i += 1

    return ativ_id_expl, data_expl, termo_titulo, rest

def _resolver_atividade_alvo(cur, turma_alvo: str, comp_alvo: str = "", ativ_id_expl: int = None, data_expl: str = "", termo_titulo: str = "", preferir_com_entregas: bool = False, prof_chave: str = ""):
    """
    Localiza de forma unificada a atividade alvo em atividades_cadastradas.
    Retorna tupla (id, data_passada, turma, componente, titulo) ou None.
    """
    if ativ_id_expl is not None:
        if prof_chave:
            cur.execute("SELECT id, data_passada, turma, componente, titulo FROM atividades_cadastradas WHERE id = ? AND prof_chave = ?", (ativ_id_expl, prof_chave))
        else:
            cur.execute("SELECT id, data_passada, turma, componente, titulo FROM atividades_cadastradas WHERE id = ?", (ativ_id_expl,))
        row = cur.fetchone()
        if row:
            return row

    if not turma_alvo:
        return None

    if prof_chave:
        cur.execute("""
            SELECT id, data_passada, turma, componente, titulo
            FROM atividades_cadastradas
            WHERE turma = ? AND prof_chave = ?
            ORDER BY id DESC
        """, (turma_alvo, prof_chave))
    else:
        cur.execute("""
            SELECT id, data_passada, turma, componente, titulo
            FROM atividades_cadastradas
            WHERE turma = ?
            ORDER BY id DESC
        """, (turma_alvo,))
    todas_turma = cur.fetchall()
    if not todas_turma:
        return None

    cands = list(todas_turma)

    if data_expl:
        cands_data = [r for r in cands if (r[1] or "").startswith(data_expl)]
        if cands_data:
            cands = cands_data
        elif os.path.exists(AULAS_CSV):
            try:
                with open(AULAS_CSV, "r", encoding="utf-8") as f_csv:
                    for row_c in csv.DictReader(f_csv, delimiter=";"):
                        dt_c = (row_c.get("data") or "").strip()
                        tm_c = (row_c.get("turma") or "").strip()
                        if tm_c == str(turma_alvo) and dt_c.startswith(data_expl):
                            tema_c = (row_c.get("tema") or "Atividade de Aula").strip()
                            comp_c = (row_c.get("componente") or comp_alvo or "INTELIGÊNCIA ARTIFICIAL").strip()
                            return (None, dt_c, turma_alvo, comp_c, tema_c)
            except Exception:
                pass

    if termo_titulo:
        t_norm = normalizar_texto(termo_titulo)
        cands_tit = [r for r in cands if t_norm in normalizar_texto(r[4] or "")]
        if cands_tit:
            cands = cands_tit
        elif os.path.exists(AULAS_CSV):
            try:
                with open(AULAS_CSV, "r", encoding="utf-8") as f_csv:
                    for row_c in reversed(list(csv.DictReader(f_csv, delimiter=";"))):
                        tm_c = (row_c.get("turma") or "").strip()
                        tema_c = (row_c.get("tema") or "").strip()
                        if (not turma_alvo or tm_c == str(turma_alvo)) and t_norm in normalizar_texto(tema_c):
                            dt_c = (row_c.get("data") or "").strip()
                            comp_c = (row_c.get("componente") or comp_alvo or "INTELIGÊNCIA ARTIFICIAL").strip()
                            return (None, dt_c, tm_c or turma_alvo, comp_c, tema_c)
            except Exception:
                pass

    if comp_alvo:
        cands_comp = [r for r in cands if _comp_equiv(comp_alvo, r[3] or "")]
        if cands_comp:
            cands = cands_comp

    escolhida = cands[0]

    # Se não foi passado seletor explícito e estamos cobrando, verifica se a mais recente está zerada
    # enquanto outra atividade da mesma turma recebeu vistos hoje/recentemente
    if preferir_com_entregas and ativ_id_expl is None and not data_expl and not termo_titulo and len(cands) > 1:
        qtd_top = _contar_vistos_atividade(cur, turma_alvo, escolhida[0], escolhida[4])
        if qtd_top == 0:
            for r_alt in cands[1:]:
                qtd_alt = _contar_vistos_atividade(cur, turma_alvo, r_alt[0], r_alt[4])
                if qtd_alt > 0:
                    escolhida = r_alt
                    break

    return escolhida

def _buscar_alunos_por_termo_turma(item_str: str, alunos_numerados: list):
    """
    Dado um termo (nº chamada '6', intervalo '5-8', matrícula ou nome) e
    alunos_numerados = [(num_ch, id_db, matricula, nome, turma, comp), ...],
    retorna (status_code, lista_matches) onde status_code in ('OK', 'AMBIGUO', 'NAO_ENCONTRADO').
    """
    termo = item_str.strip().strip(",;.").strip()
    termo = re.sub(r"^(?:o|a|os|as|aluno|aluna|nº|no|num|numero|número|#)\s+", "", termo, flags=re.IGNORECASE).strip()
    if not termo:
        return ("NAO_ENCONTRADO", [])

    # 1. Intervalo de números de chamada: ex '5-9' ou '5 a 9'
    m_range = re.match(r"^(\d{1,2})\s*(?:-|a|ate|até)\s*(\d{1,2})$", termo, flags=re.IGNORECASE)
    if m_range:
        ini, fim = int(m_range.group(1)), int(m_range.group(2))
        if ini > fim:
            ini, fim = fim, ini
        matches = [a for a in alunos_numerados if ini <= a[0] <= fim]
        return ("OK", matches) if matches else ("NAO_ENCONTRADO", [])

    # 2. Número de chamada direto
    if termo.isdigit() and 1 <= int(termo) <= len(alunos_numerados):
        idx = int(termo)
        return ("OK", [alunos_numerados[idx - 1]])

    # 3. Matrícula exata
    for a in alunos_numerados:
        if termo.lower() == (a[2] or "").lower():
            return ("OK", [a])

    # 4. Nome normalizado sem perda fonética (diferencia Davi vs Davy, Isabela vs Isabella vs Isabelly, Ana vs Anna)
    t_norm = normalizar_texto(termo)
    pals_norm = t_norm.split()
    if pals_norm:
        # 4a. Match de palavras inteiras exatas
        cands_exatos = []
        for a in alunos_numerados:
            w_nome = normalizar_texto(a[3]).split()
            if all(p in w_nome for p in pals_norm):
                cands_exatos.append(a)
        if len(cands_exatos) == 1:
            return ("OK", cands_exatos)
        if len(cands_exatos) > 1:
            return ("AMBIGUO", cands_exatos)

        # 4b. Match por prefixo de palavra (ex: 'rakey' -> 'rakeyllane')
        cands_pref = []
        for a in alunos_numerados:
            w_nome = normalizar_texto(a[3]).split()
            if all(any(w.startswith(p) for w in w_nome) for p in pals_norm):
                cands_pref.append(a)
        if len(cands_pref) == 1:
            return ("OK", cands_pref)
        if len(cands_pref) > 1:
            return ("AMBIGUO", cands_pref)

    # 5. Fallback fonético (ex: 'renan' -> 'rennan', 'thalison' -> 'thallysson', 'fedro' -> 'phedro')
    t_fon = normalizar_fonetico(termo)
    pals_fon = t_fon.split()
    if pals_fon:
        cands_fon_w = []
        for a in alunos_numerados:
            w_fon = normalizar_fonetico(a[3]).split()
            if all(p in w_fon for p in pals_fon):
                cands_fon_w.append(a)
        if len(cands_fon_w) == 1:
            return ("OK", cands_fon_w)
        if len(cands_fon_w) > 1:
            return ("AMBIGUO", cands_fon_w)

        cands_fon_pref = []
        for a in alunos_numerados:
            w_fon = normalizar_fonetico(a[3]).split()
            if all(any(w.startswith(p) for w in w_fon) for p in pals_fon):
                cands_fon_pref.append(a)
        if len(cands_fon_pref) == 1:
            return ("OK", cands_fon_pref)
        if len(cands_fon_pref) > 1:
            return ("AMBIGUO", cands_fon_pref)

    return ("NAO_ENCONTRADO", [])

def _separar_itens_alunos(texto_lista: str) -> list:
    """
    Divide uma string de alunos/números ('2, 3, ana sophia e davy' ou '2 3 10 27 28')
    em uma lista limpa de termos individuais.
    """
    t = texto_lista.strip()
    # Remove palavras de status residuais no fim ou início
    t = re.sub(r"\b(?:ok|visto|entregue|feito|pendente|pendentes|incompleto|recusado)\b", "", t, flags=re.IGNORECASE).strip()
    if not t:
        return []
    # Substitui conectores por vírgula
    t_norm_sep = re.sub(r"\s+(?:e|ou)\s+|[;/\n]+", ",", t, flags=re.IGNORECASE)
    partes_brutas = [p.strip() for p in t_norm_sep.split(",") if p.strip()]
    itens = []
    for p in partes_brutas:
        # Se a parte contém apenas números ou intervalos separados por espaço (ex: '2 3 10 27 28')
        sub_toks = p.split()
        if len(sub_toks) > 1 and all(re.match(r"^#?\d{1,2}(?:-\d{1,2})?$", st) for st in sub_toks):
            itens.extend([st.lstrip("#") for st in sub_toks])
        else:
            itens.append(p)
    return itens

def _gravar_visto_unico_db(cur, ativ_id, data_reg, data_iso, turma, comp, matricula, nome, ativ_desc, status_alvo, obs, foto_caminho="", avaliacao_ia="", prof_chave="pedro"):
    """Grava ou atualiza o visto do aluno na atividade sem gerar duplicatas."""
    if ativ_id is not None:
        cur.execute("""
            DELETE FROM atividades_entregas
            WHERE turma = ? AND matricula = ? AND (atividade_id = ? OR atividade = ?)
              AND (prof_chave = ? OR prof_chave IS NULL OR prof_chave = '')
        """, (turma, matricula, ativ_id, ativ_desc, prof_chave or "pedro"))
    else:
        cur.execute("""
            DELETE FROM atividades_entregas
            WHERE turma = ? AND matricula = ? AND atividade = ?
              AND (prof_chave = ? OR prof_chave IS NULL OR prof_chave = '')
        """, (turma, matricula, ativ_desc, prof_chave or "pedro"))

    cur.execute("""
        INSERT INTO atividades_entregas
        (atividade_id, data_registro, data_hora_iso, turma, componente, matricula, nome_aluno, atividade, status, observacao, foto_caminho, avaliacao_ia, prof_chave)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (ativ_id, data_reg, data_iso, turma, comp, matricula, nome, ativ_desc, status_alvo, obs, foto_caminho or "", avaliacao_ia or "", prof_chave or "pedro"))

def mover_atividade_entre_turmas(texto: str) -> str:
    """
    Move uma ou mais atividades cadastradas (e seus vistos) para outra turma.
    Exemplos:
      - `!atv mover #64 202`
      - `!atv mover 64 para 202`
      - `!atv mover turma 64 202`
      - `!atv mover #64 #65 para 202`
    """
    init_tabelas_atividades()

    # Identifica a turma de destino
    turma_destino = None
    toks = [tok.strip(",;.#").lower() for tok in texto.split()]
    for tok in toks:
        if tok in TURMAS_VALIDAS_ATV:
            turma_destino = tok
            break

    if not turma_destino:
        return "⚠️ Informe a turma de destino válida (ex: 101, 102, 201, 202, 203, etc.). Exemplo: `!atv mover #64 para 202`"

    # Encontra os IDs das atividades a serem movidas (ignorando números que representem a própria turma)
    todos_numeros = re.findall(r"(?:#|id[:=]?)?(\b\d+\b)", texto)
    ids_alvos = [int(n) for n in todos_numeros if n != turma_destino]

    if not ids_alvos:
        return "⚠️ Informe o ID da atividade a ser movida. Exemplo: `!atv mover #64 para 202`"

    relatorio_movidos = []
    with sqlite3.connect(ESCOLA_DB) as conn:
        cur = conn.cursor()
        for id_alvo in ids_alvos:
            cur.execute("SELECT id, data_passada, turma, componente, titulo FROM atividades_cadastradas WHERE id = ?", (id_alvo,))
            r_ativ = cur.fetchone()
            if not r_ativ:
                relatorio_movidos.append(f"❌ Atividade #{id_alvo} não encontrada.")
                continue

            turma_antiga = str(r_ativ[2])
            if turma_antiga == turma_destino:
                relatorio_movidos.append(f"ℹ️ Atividade #{id_alvo} já pertence à Turma {turma_destino}.")
                continue

            # Atualiza turma da atividade cadastrada
            cur.execute("UPDATE atividades_cadastradas SET turma = ? WHERE id = ?", (turma_destino, id_alvo))

            # Atualiza também nas entregas existentes
            cur.execute("SELECT COUNT(*) FROM atividades_entregas WHERE atividade_id = ?", (id_alvo,))
            qtd_vistos = cur.fetchone()[0]
            if qtd_vistos > 0:
                cur.execute("UPDATE atividades_entregas SET turma = ? WHERE atividade_id = ?", (turma_destino, id_alvo))

            vistos_txt = f" ({qtd_vistos} vistos atualizados)" if qtd_vistos > 0 else ""
            relatorio_movidos.append(f"• *#{id_alvo}* {r_ativ[4][:40]}: Turma {turma_antiga} ➔ *Turma {turma_destino}*{vistos_txt}")

        conn.commit()

    if not relatorio_movidos:
        return "❌ Nenhuma atividade pôde ser movida. Verifique os IDs informados."

    return (
        f"🔄 *GERÊNCIA DE ATIVIDADES — TRANSFERÊNCIA DE TURMA*\n\n"
        + "\n".join(relatorio_movidos)
        + f"\n\n💡 Para consultar a turma: `!atv {turma_destino}`"
    )

def transferir_ou_copiar_vistos_atividade(texto: str, mover: bool = False) -> str:
    """
    Permite mover ou copiar vistos de uma atividade #origem para #destino,
    OU mover uma atividade para outra turma caso contenha uma turma válida.
    Ex: 'atv mover 202 #5 #6', 'atv copiar #5 #6' ou 'atv mover #64 para 202'
    """
    init_tabelas_atividades()
    turma_detectada = any(tok.strip(",;.#").lower() in TURMAS_VALIDAS_ATV for tok in texto.split())
    ids = [int(x) for x in re.findall(r"(?:#|id[:=]?)?(\b\d+\b)", texto) if int(x) not in (100, 101, 102, 103, 200, 201, 202, 203, 300, 301, 302, 303)]

    # Se for comando de mover e houver turma de destino indicada
    if mover and turma_detectada and ("turma" in texto.lower() or len(ids) == 1 or "para" in texto.lower()):
        return mover_atividade_entre_turmas(texto)

    if len(ids) < 2:
        if turma_detectada:
            return mover_atividade_entre_turmas(texto)
        return (
            "⚠️ *Comandos de Gerência de Atividades:*\n\n"
            "• *Mover atividade para outra turma:*\n"
            "  `!atv mover #64 202` ou `!atv mover 64 para 202`\n\n"
            "• *Transferir vistos entre atividades:*\n"
            "  `!atv mover #5 #6` (move vistos de #5 para #6)\n"
            "  `!atv copiar #5 #6` (copia vistos de #5 para #6)"
        )
    id_orig, id_dest = ids[0], ids[1]
    with sqlite3.connect(ESCOLA_DB) as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, data_passada, turma, componente, titulo FROM atividades_cadastradas WHERE id = ?", (id_orig,))
        r_orig = cur.fetchone()
        cur.execute("SELECT id, data_passada, turma, componente, titulo FROM atividades_cadastradas WHERE id = ?", (id_dest,))
        r_dest = cur.fetchone()
        if not r_orig or not r_dest:
            return f"❌ Atividade #{id_orig} ou #{id_dest} não encontrada. Veja os IDs com `atv lista`."

        cur.execute("""
            SELECT data_registro, data_hora_iso, turma, matricula, nome_aluno, status, observacao, foto_caminho, avaliacao_ia
            FROM atividades_entregas
            WHERE atividade_id = ?
        """, (id_orig,))
        entregas = cur.fetchall()
        if not entregas:
            return f"⚠️ Nenhum visto encontrado na atividade *#{id_orig}* para {'mover' if mover else 'copiar'}."

        for dt_r, dt_i, t_e, mat, nom, st, obs, foto, av_ia in entregas:
            _gravar_visto_unico_db(cur, id_dest, dt_r, dt_i, r_dest[2], r_dest[3], mat, nom, r_dest[4], st, obs, foto, av_ia)

        if mover:
            cur.execute("DELETE FROM atividades_entregas WHERE atividade_id = ?", (id_orig,))
        conn.commit()
        resumo_t = _resumo_atividades_turma_str(cur, r_dest[2], id_dest)

    acao_str = "MOVIDOS" if mover else "COPIADOS"
    return (
        f"✅ *VISTOS {acao_str} COM SUCESSO!*\n\n"
        f"• *De:* #{id_orig} — {r_orig[4]} ({r_orig[1]})\n"
        f"• *Para:* #{id_dest} — {r_dest[4]} ({r_dest[1]})\n"
        f"• *Total de registros:* {len(entregas)} alunos\n\n"
        f"{resumo_t}"
    )

def listar_atividades_cadastradas(texto: str = "", prof_chave: str = "") -> str:
    init_tabelas_atividades()
    t = texto.strip()
    turma = ""
    for tok in t.split():
        if tok.strip(",;.") in TURMAS_VALIDAS_ATV:
            turma = tok.strip(",;.")
            break
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            if turma:
                if prof_chave:
                    cur.execute("""
                        SELECT id, data_passada, turma, componente, titulo 
                        FROM atividades_cadastradas 
                        WHERE turma = ? AND prof_chave = ?
                        ORDER BY id DESC LIMIT 20
                    """, (turma, prof_chave))
                else:
                    cur.execute("""
                        SELECT id, data_passada, turma, componente, titulo 
                        FROM atividades_cadastradas 
                        WHERE turma = ? 
                        ORDER BY id DESC LIMIT 20
                    """, (turma,))
            else:
                if prof_chave:
                    cur.execute("""
                        SELECT id, data_passada, turma, componente, titulo 
                        FROM atividades_cadastradas 
                        WHERE prof_chave = ?
                        ORDER BY id DESC LIMIT 20
                    """, (prof_chave,))
                else:
                    cur.execute("""
                        SELECT id, data_passada, turma, componente, titulo 
                        FROM atividades_cadastradas 
                        ORDER BY id DESC LIMIT 20
                    """)
            linhas = cur.fetchall()

            if not linhas:
                turma_str = f" para a Turma {turma}" if turma else ""
                hint_comp = "MKT" if prof_chave == "bruna" else ("WEB" if prof_chave == "marcos" else "IA")
                hint_tm = turma or ("101" if prof_chave == "bruna" else ("102" if prof_chave == "marcos" else "201"))
                return f"📝 *Nenhuma atividade cadastrada{turma_str}.*\nCadastre com: `!atv nova {hint_tm} {hint_comp} <Título>`"

            turma_title = f" — TURMA {turma}" if turma else ""
            resp = [f"📋 *ATIVIDADES CADASTRADAS{turma_title}*:\n"]
            for aid, dt, t_row, comp, tit in linhas:
                cur.execute("SELECT COUNT(*) FROM alunos WHERE turma = ?", (t_row,))
                tot_al = (cur.fetchone() or [0])[0]
                qtd_v = _contar_vistos_atividade(cur, t_row, aid, tit, prof_chave=prof_chave)
                resp.append(
                    f"• *#{aid}* | *{tit}*\n"
                    f"  🏫 Turma {t_row} • {comp} • Passada em: _{dt}_\n"
                    f"  ✅ Entregas: *{qtd_v}/{tot_al} alunos*\n"
                    f"  👉 Lançar todos: `!atv #{aid} {t_row} todos exceto <nº ou nomes>`\n"
                    f"  👉 Cobrar: `!cobrar {t_row} #{aid}`\n"
                )
        return "\n".join(resp)
    except Exception as e:
        return f"❌ Erro ao listar atividades: {e}"

def cobrar_atividades_turma(texto: str, prof_chave: str = "") -> str:
    """
    Cruza todos os alunos matriculados na turma com os que receberam visto,
    destacando quem está pendente, a atividade (#ID) e a data em que foi passada.
    """
    init_tabelas_atividades()
    t = re.sub(r'^(cobrar|pendente|pendentes|pendencias|pendencia|atv\s+cobrar|atv\s+pendentes)\s*', '', texto.strip(), flags=re.IGNORECASE).strip()
    tokens = t.split()

    ativ_id_expl, data_expl, termo_titulo, tokens_after_sel = _extrair_seletor_atividade(tokens)

    turma_alvo = ""
    comp_alvo = ""
    tokens_rest = []

    for tok in tokens_after_sel:
        tok_clean = tok.strip(",;.")
        if tok_clean in TURMAS_VALIDAS_ATV and not turma_alvo:
            turma_alvo = tok_clean
            continue
        tok_low = normalizar_texto(tok_clean)
        if tok_low in COMP_MAP_ATV and not comp_alvo:
            comp_alvo = COMP_MAP_ATV[tok_low]
            continue
        if tok_clean.lstrip("#").isdigit() and ativ_id_expl is None:
            ativ_id_expl = int(tok_clean.lstrip("#"))
            continue
        tokens_rest.append(tok)

    if tokens_rest and not termo_titulo:
        termo_titulo = " ".join(tokens_rest).strip().lstrip("#")

    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()

            # Se informou apenas #ID (ex: 'cobrar #5'), descobre a turma pela atividade
            if ativ_id_expl is not None and not turma_alvo:
                cur.execute("SELECT turma FROM atividades_cadastradas WHERE id = ?", (ativ_id_expl,))
                r_t = cur.fetchone()
                if r_t:
                    turma_alvo = r_t[0]

            if not turma_alvo:
                turma_atual = obter_turma_atual_ou_proxima()
                if prof_chave == "bruna":
                    turma_alvo = turma_atual if turma_atual in ("101", "201") else "201"
                elif prof_chave == "marcos":
                    turma_alvo = turma_atual if turma_atual in ("102", "103", "203") else "102"
                else:
                    turma_alvo = turma_atual or "202"

            # 1. Obter todos os alunos matriculados na turma em ordem alfabética
            cur.execute("""
                SELECT matricula, nome 
                FROM alunos 
                WHERE turma = ? 
                ORDER BY nome ASC
            """, (turma_alvo,))
            alunos_turma = cur.fetchall()

            if not alunos_turma:
                return f"⚠️ Nenhum aluno encontrado para a Turma {turma_alvo} no banco de dados."

            # 2. Localizar a atividade especificada ou a mais adequada
            ativ_selecionada = _resolver_atividade_alvo(
                cur, turma_alvo, comp_alvo, ativ_id_expl, data_expl, termo_titulo, preferir_com_entregas=True, prof_chave=prof_chave
            )

            # 3. Mapear os vistos concedidos
            vistos_registrados = {}
            if ativ_selecionada:
                ativ_id, ativ_data_passada, ativ_turma, ativ_comp, ativ_titulo = ativ_selecionada
                if prof_chave:
                    cur.execute("""
                        SELECT matricula, nome_aluno, status, data_registro 
                        FROM atividades_entregas 
                        WHERE turma = ? AND (atividade_id = ? OR atividade = ? OR atividade LIKE ?)
                          AND prof_chave = ?
                        ORDER BY id ASC
                    """, (turma_alvo, ativ_id, ativ_titulo, f"%{ativ_titulo}%", prof_chave))
                else:
                    cur.execute("""
                        SELECT matricula, nome_aluno, status, data_registro 
                        FROM atividades_entregas 
                        WHERE turma = ? AND (atividade_id = ? OR atividade = ? OR atividade LIKE ?)
                        ORDER BY id ASC
                    """, (turma_alvo, ativ_id, ativ_titulo, f"%{ativ_titulo}%"))
                for mat, nom, st, dt in cur.fetchall():
                    if (st or "").upper() in ["VISTO (OK)", "OK", "ENTREGUE", "FEITO", "10"]:
                        vistos_registrados[mat] = (nom, dt, st)
                    elif (st or "").upper() in ["PENDENTE", "INCOMPLETO", "RECUSADO"]:
                        vistos_registrados.pop(mat, None)
            else:
                if prof_chave:
                    cur.execute("""
                        SELECT matricula, nome_aluno, status, data_registro, atividade, data_hora_iso, componente 
                        FROM atividades_entregas 
                        WHERE turma = ? AND prof_chave = ?
                        ORDER BY id DESC
                    """, (turma_alvo, prof_chave))
                else:
                    cur.execute("""
                        SELECT matricula, nome_aluno, status, data_registro, atividade, data_hora_iso, componente 
                        FROM atividades_entregas 
                        WHERE turma = ? 
                        ORDER BY id DESC
                    """, (turma_alvo,))
                entregas_gerais = cur.fetchall()
                if entregas_gerais:
                    ativ_id = None
                    ativ_titulo = entregas_gerais[0][4] or "Visto de Atividade / Caderno"
                    comp_db = entregas_gerais[0][6] if len(entregas_gerais[0]) > 6 and entregas_gerais[0][6] else ""
                    default_comp = "MARKETING DIGITAL" if prof_chave == "bruna" else ("PROGRAMAÇÃO PARA WEB" if prof_chave == "marcos" else "INTELIGÊNCIA ARTIFICIAL")
                    ativ_comp = comp_alvo or comp_db or default_comp
                    ativ_data_passada = entregas_gerais[0][3].split(" às ")[0] if " às " in entregas_gerais[0][3] else "16/10/2026"
                    for row_e in entregas_gerais:
                        mat, nom, st, dt = row_e[0], row_e[1], row_e[2], row_e[3]
                        if (st or "").upper() in ["VISTO (OK)", "OK", "ENTREGUE", "FEITO", "10"] and mat not in vistos_registrados:
                            vistos_registrados[mat] = (nom, dt, st)
                else:
                    hint_comp = "MKT" if prof_chave == "bruna" else ("WEB" if prof_chave == "marcos" else "IA")
                    return (
                        f"📝 *Nenhuma atividade cadastrada para a Turma {turma_alvo}.*\n\n"
                        f"Para registrar uma atividade e poder cobrar os alunos:\n"
                        f"`!atv nova {turma_alvo} {hint_comp} <Título da Atividade>`\n\n"
                        f"Exemplo:\n"
                        f"`!atv nova {turma_alvo} {hint_comp} Exercício de Fixação`\n"
                        f"Em seguida, digite `!cobrar {turma_alvo}` para acompanhar quem entregou!"
                    )

            # 4. Cruzamento dos alunos com os vistos
            total_alunos = len(alunos_turma)
            entregues = []
            pendentes = []

            for idx, (matricula, nome) in enumerate(alunos_turma, 1):
                teve_visto = matricula in vistos_registrados
                if not teve_visto:
                    nome_fon = normalizar_fonetico(nome)
                    for v_mat, (v_nom, v_dt, v_st) in list(vistos_registrados.items()):
                        if nome_fon == normalizar_fonetico(v_nom):
                            teve_visto = True
                            vistos_registrados[matricula] = (v_nom, v_dt, v_st)
                            break

                if teve_visto:
                    info_v = vistos_registrados.get(matricula, (nome, "Registrado", "OK"))
                    entregues.append((idx, nome, info_v[1]))
                else:
                    pendentes.append((idx, nome))

            # 5. Estatísticas
            pct_entregue = round((len(entregues) / total_alunos) * 100) if total_alunos else 0
            pct_pendente = round((len(pendentes) / total_alunos) * 100) if total_alunos else 0

            id_tag = f"#{ativ_id} — " if ativ_id else ""
            cmd_id_hint = f"#{ativ_id} " if ativ_id else ""
            resumo_outras = _resumo_atividades_turma_str(cur, turma_alvo, ativ_id)

            # 6. Formatação do relatório de cobrança para leitura rápida em sala
            resp = [
                f"📋 *COBRANÇA DE ATIVIDADES — TURMA {turma_alvo}*",
                f"📖 *Componente:* {ativ_comp}",
                f"📝 *Atividade:* *{id_tag}{ativ_titulo}*",
                f"📅 *Passada em:* {ativ_data_passada}",
                f"━━━━━━━━━━━━━━━━━━━━",
                f"📊 *Estatísticas da Turma:*",
                f"• Total de alunos: {total_alunos}",
                f"• ✅ Entregaram (Visto OK): *{len(entregues)}* ({pct_entregue}%)",
                f"• ⏳ Pendentes: *{len(pendentes)}* ({pct_pendente}%)",
                f"━━━━━━━━━━━━━━━━━━━━"
            ]

            if pendentes:
                resp.append(f"🚨 *ALUNOS PENDENTES ({len(pendentes)}):*")
                for num_ch, nome_aluno in pendentes:
                    resp.append(f"  *{num_ch:02d}.* {nome_aluno}")
            else:
                resp.append("🎉 *PARABÉNS! TODOS OS ALUNOS ENTREGARAM A ATIVIDADE!*")

            if entregues:
                resp.append(f"\n✅ *JÁ ENTREGARAM ({len(entregues)}):*")
                for num_ch, nome_aluno, dt_ent in entregues[:12]:
                    dt_fmt = dt_ent.split(" às ")[0] if " às " in dt_ent else dt_ent
                    primeiro_nome = nome_aluno.split()[0]
                    resp.append(f"  • #{num_ch:02d} {primeiro_nome} ({dt_fmt})")
                if len(entregues) > 12:
                    resp.append(f"  _... e mais {len(entregues) - 12} alunos com visto._")

            if resumo_outras:
                resp.append(f"\n━━━━━━━━━━━━━━━━━━━━\n{resumo_outras}")

            resp.append(
                f"\n💡 *Atalhos rápidos para esta atividade:*\n"
                f"• Todos menos alguns: `!atv {cmd_id_hint}{turma_alvo} todos exceto 2, 3`\n"
                f"• Visto individual/lote: `!atv {cmd_id_hint}{turma_alvo} 1, 5, 8 ok`"
            )
            return "\n".join(resp)

    except Exception as e:
        return f"❌ Erro ao cruzar cobrança de atividades: {e}"

def _executar_visto_em_massa_todos_exceto(
    turma_alvo: str,
    comp_alvo: str,
    ativ_id_expl: int,
    data_expl: str,
    termo_titulo: str,
    texto_excecoes: str,
    observacao_cmd: str,
    prof_chave: str = "pedro"
) -> str:
    """
    Marca VISTO (OK) para TODOS os alunos da turma na atividade indicada,
    EXCETO os alunos listados em texto_excecoes (que ficam pendentes).
    """
    with sqlite3.connect(ESCOLA_DB) as conn:
        cur = conn.cursor()
        if ativ_id_expl is not None and not turma_alvo:
            cur.execute("SELECT turma FROM atividades_cadastradas WHERE id = ?", (ativ_id_expl,))
            r_t = cur.fetchone()
            if r_t:
                turma_alvo = r_t[0]

        if not turma_alvo:
            turma_atual = obter_turma_atual_ou_proxima()
            if prof_chave == "bruna":
                turma_alvo = turma_atual if turma_atual in ("101", "201") else "201"
            elif prof_chave == "marcos":
                turma_alvo = turma_atual if turma_atual in ("102", "103", "203") else "102"
            else:
                turma_alvo = turma_atual or "202"

        cur.execute("""
            SELECT id, matricula, nome, turma, componente
            FROM alunos
            WHERE turma = ?
            ORDER BY nome ASC
        """, (turma_alvo,))
        rows_alunos = cur.fetchall()
        if not rows_alunos:
            return f"⚠️ Nenhum aluno encontrado na Turma {turma_alvo}."

        alunos_numerados = [
            (idx, r[0], r[1], r[2], r[3], r[4])
            for idx, r in enumerate(rows_alunos, 1)
        ]

        # Resolve exceções
        itens_exc = _separar_itens_alunos(texto_excecoes) if texto_excecoes else []
        excecoes_confirmadas = {}  # matricula -> (num_ch, nome)
        erros_validacao = []

        for item in itens_exc:
            st_busca, matches = _buscar_alunos_por_termo_turma(item, alunos_numerados)
            if st_busca == "OK":
                for m in matches:
                    excecoes_confirmadas[m[2]] = (m[0], m[3])
            elif st_busca == "AMBIGUO":
                opcoes = ", ".join([f"*#{m[0]:02d}* {m[3]}" for m in matches[:5]])
                erros_validacao.append(f"• *'{item}'* é ambíguo ({opcoes})")
            else:
                erros_validacao.append(f"• *'{item}'* não encontrado na Turma {turma_alvo}")

        if erros_validacao:
            id_str = f"#{ativ_id_expl} " if ativ_id_expl else ""
            return (
                f"⚠️ *Atenção antes de lançar TODOS na Turma {turma_alvo}:*\n"
                f"Não consegui identificar com certeza {len(erros_validacao)} exceção(ões):\n" +
                "\n".join(erros_validacao) +
                f"\n\n👉 Reenvie usando o *nº da chamada* ou nome mais específico.\n"
                f"Exemplo: `atv {id_str}{turma_alvo} todos exceto 2, 7`"
            )

        # Resolve componente se vazio
        if not comp_alvo:
            if prof_chave == "bruna":
                comp_alvo = "INTRODUÇÃO AO MARKETING" if turma_alvo == "101" else "MARKETING DIGITAL"
            elif prof_chave == "marcos":
                comp_alvo = "PROGRAMAÇÃO PARA WEB" if turma_alvo == "203" else "INTRODUÇÃO À PROGRAMAÇÃO WEB"
            else:
                ctx = obter_contexto_aula_atual()
                if ctx.get("tem_aula") and ctx.get("turma") == turma_alvo and ctx.get("disciplina"):
                    comp_alvo = ctx.get("disciplina")
                else:
                    wd_hoje = datetime.now().weekday()
                    cur.execute("SELECT disciplina FROM grade_horaria WHERE dia_semana = ? AND turma = ? LIMIT 1", (wd_hoje, turma_alvo))
                    row_g = cur.fetchone()
                    comp_alvo = row_g[0] if row_g else (rows_alunos[0][4].split(",")[0].strip() if rows_alunos[0][4] else "INFORMÁTICA")

        row_ativ = _resolver_atividade_alvo(
            cur, turma_alvo, comp_alvo, ativ_id_expl, data_expl, termo_titulo, preferir_com_entregas=False, prof_chave=prof_chave
        )
        if row_ativ:
            ativ_id_vinc, dt_passada, _, comp_ativ, ativ_desc = row_ativ
            comp_final = comp_ativ or comp_alvo
            data_passada_info = f" (Passada em: {dt_passada})"
        else:
            ativ_id_vinc = None
            ativ_desc = "Atividade de Aula / Visto de Caderno"
            comp_final = comp_alvo
            data_passada_info = ""

        agora = datetime.now()
        data_reg = agora.strftime("%d/%m/%Y às %H:%M")
        data_iso = agora.strftime("%Y-%m-%d %H:%M:%S")

        qtd_ok = 0
        for num_ch, id_db, mat, nome, t_al, c_al in alunos_numerados:
            if mat in excecoes_confirmadas:
                # Garante que quem está na exceção não fique com VISTO (OK) antigo nesta atividade
                if ativ_id_vinc is not None:
                    cur.execute("DELETE FROM atividades_entregas WHERE turma = ? AND matricula = ? AND (atividade_id = ? OR atividade = ?) AND prof_chave = ?", (turma_alvo, mat, ativ_id_vinc, ativ_desc, prof_chave or "pedro"))
                else:
                    cur.execute("DELETE FROM atividades_entregas WHERE turma = ? AND matricula = ? AND atividade = ? AND prof_chave = ?", (turma_alvo, mat, ativ_desc, prof_chave or "pedro"))
            else:
                _gravar_visto_unico_db(
                    cur, ativ_id_vinc, data_reg, data_iso, turma_alvo, comp_final,
                    mat, nome, ativ_desc, "VISTO (OK)", observacao_cmd, prof_chave=prof_chave or "pedro"
                )
                qtd_ok += 1

        conn.commit()
        resumo_ativs = _resumo_atividades_turma_str(cur, turma_alvo, ativ_id_vinc, prof_chave=prof_chave)

    total_turma = len(alunos_numerados)
    pct_ok = round((qtd_ok / total_turma) * 100) if total_turma else 0
    id_label = f"#{ativ_id_vinc} — " if ativ_id_vinc else ""
    cmd_id = f" #{ativ_id_vinc}" if ativ_id_vinc else ""

    linhas_exc = []
    for mat, (n_ch, n_al) in sorted(excecoes_confirmadas.items(), key=lambda x: x[1][0]):
        linhas_exc.append(f"  *{n_ch:02d}.* {n_al}")

    bloco_exc = (
        f"🚨 *EXCEÇÕES / PENDENTES ({len(linhas_exc)}):*\n" + "\n".join(linhas_exc)
        if linhas_exc else
        "🎉 *Nenhuma exceção! 100% da turma recebeu visto!*"
    )

    return (
        f"✅ *VISTO EM MASSA REGISTRADO — TURMA {turma_alvo}!*\n\n"
        f"📝 *Atividade:* *{id_label}{ativ_desc}*{data_passada_info}\n"
        f"📖 *Componente:* {comp_final}\n"
        f"📊 *Resultado:* *{qtd_ok}/{total_turma}* alunos com VISTO (OK) ({pct_ok}%)\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"{bloco_exc}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"{resumo_ativs}\n\n"
        f"📊 *Para publicar a cobrança:* `!cobrar {turma_alvo}{cmd_id}`"
    )

def _executar_visto_lote_alunos(
    turma_alvo: str,
    comp_alvo: str,
    ativ_id_expl: int,
    data_expl: str,
    termo_titulo: str,
    itens_alunos: list,
    status_alvo: str,
    observacao_cmd: str,
    prof_chave: str = "pedro"
) -> str:
    """
    Registra visto (ou status pendente/incompleto) para uma lista de alunos/números de uma só vez.
    Ex: 'atv #6 202 1, 4, 5, 8-12 ok' ou 'atv 202 ana beatriz, anna luisa ok'
    """
    with sqlite3.connect(ESCOLA_DB) as conn:
        cur = conn.cursor()
        if ativ_id_expl is not None and not turma_alvo:
            cur.execute("SELECT turma FROM atividades_cadastradas WHERE id = ?", (ativ_id_expl,))
            r_t = cur.fetchone()
            if r_t:
                turma_alvo = r_t[0]

        if not turma_alvo:
            turma_atual = obter_turma_atual_ou_proxima()
            if prof_chave == "bruna":
                turma_alvo = turma_atual if turma_atual in ("101", "201") else "201"
            elif prof_chave == "marcos":
                turma_alvo = turma_atual if turma_atual in ("102", "103", "203") else "102"
            else:
                turma_alvo = turma_atual or "202"

        cur.execute("""
            SELECT id, matricula, nome, turma, componente
            FROM alunos
            WHERE turma = ?
            ORDER BY nome ASC
        """, (turma_alvo,))
        rows_alunos = cur.fetchall()
        if not rows_alunos:
            return f"⚠️ Nenhum aluno encontrado na Turma {turma_alvo}."

        alunos_numerados = [
            (idx, r[0], r[1], r[2], r[3], r[4])
            for idx, r in enumerate(rows_alunos, 1)
        ]

        selecionados = {}  # matricula -> (num_ch, id_db, matricula, nome, turma, comp)
        erros = []
        for item in itens_alunos:
            st_busca, matches = _buscar_alunos_por_termo_turma(item, alunos_numerados)
            if st_busca == "OK":
                for m in matches:
                    selecionados[m[2]] = m
            elif st_busca == "AMBIGUO":
                opcoes = ", ".join([f"*#{m[0]:02d}* {m[3]}" for m in matches[:4]])
                erros.append(f"• *'{item}'* ambíguo ({opcoes})")
            else:
                erros.append(f"• *'{item}'* não encontrado")

        if not selecionados:
            return (
                f"⚠️ Nenhum aluno válido identificado na Turma {turma_alvo}:\n" +
                "\n".join(erros)
            )

        if not comp_alvo:
            if prof_chave == "bruna":
                comp_alvo = "INTRODUÇÃO AO MARKETING" if turma_alvo == "101" else "MARKETING DIGITAL"
            elif prof_chave == "marcos":
                comp_alvo = "PROGRAMAÇÃO PARA WEB" if turma_alvo == "203" else "INTRODUÇÃO À PROGRAMAÇÃO WEB"
            else:
                ctx = obter_contexto_aula_atual()
                if ctx.get("tem_aula") and ctx.get("turma") == turma_alvo and ctx.get("disciplina"):
                    comp_alvo = ctx.get("disciplina")
                else:
                    wd_hoje = datetime.now().weekday()
                    cur.execute("SELECT disciplina FROM grade_horaria WHERE dia_semana = ? AND turma = ? LIMIT 1", (wd_hoje, turma_alvo))
                    row_g = cur.fetchone()
                    comp_alvo = row_g[0] if row_g else (rows_alunos[0][4].split(",")[0].strip() if rows_alunos[0][4] else "INFORMÁTICA")

        row_ativ = _resolver_atividade_alvo(
            cur, turma_alvo, comp_alvo, ativ_id_expl, data_expl, termo_titulo, preferir_com_entregas=False, prof_chave=prof_chave
        )
        if row_ativ:
            ativ_id_vinc, dt_passada, _, comp_ativ, ativ_desc = row_ativ
            comp_final = comp_ativ or comp_alvo
            data_passada_info = f" (Passada em: {dt_passada})"
        else:
            ativ_id_vinc = None
            ativ_desc = "Atividade de Aula / Visto de Caderno"
            comp_final = comp_alvo
            data_passada_info = ""

        agora = datetime.now()
        data_reg = agora.strftime("%d/%m/%Y às %H:%M")
        data_iso = agora.strftime("%Y-%m-%d %H:%M:%S")

        lista_gravados = []
        for mat, m in sorted(selecionados.items(), key=lambda x: x[1][0]):
            _gravar_visto_unico_db(
                cur, ativ_id_vinc, data_reg, data_iso, turma_alvo, comp_final,
                mat, m[3], ativ_desc, status_alvo, observacao_cmd, prof_chave=prof_chave or "pedro"
            )
            lista_gravados.append(f"  • *#{m[0]:02d}* {m[3]}")

        conn.commit()
        resumo_ativs = _resumo_atividades_turma_str(cur, turma_alvo, ativ_id_vinc, prof_chave=prof_chave)

    id_label = f"#{ativ_id_vinc} — " if ativ_id_vinc else ""
    cmd_id = f" #{ativ_id_vinc}" if ativ_id_vinc else ""
    bloco_erros = ("\n⚠️ *Não lançados:*\n" + "\n".join(erros) + "\n") if erros else ""

    return (
        f"✅ *LOTE DE ATIVIDADES REGISTRADO ({len(lista_gravados)} alunos)!*\n\n"
        f"🏫 *Turma:* {turma_alvo} | 📝 *Status:* {status_alvo}\n"
        f"📌 *Atividade:* *{id_label}{ativ_desc}*{data_passada_info}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n" +
        "\n".join(lista_gravados) +
        f"\n{bloco_erros}━━━━━━━━━━━━━━━━━━━━\n"
        f"{resumo_ativs}\n\n"
        f"📊 *Para cobrar a turma:* `!cobrar {turma_alvo}{cmd_id}`"
    )

def processar_comando_atividade(texto: str, chat_id: str = DONO_CHAT_ID, data: dict = None) -> str:
    """
    Processa registros individuais, registros em massa ('todos exceto ...' ou lotes '1, 2, 5'),
    seleção explícita de atividade ('#6', data '20/10'), cadastros, transferências e correção via foto (IA Vision).
    """
    init_tabelas_atividades()

    author_num = ""
    author_raw = chat_id
    pushname = ""
    if data:
        author_raw = data.get("author") or (data.get("sender") or {}).get("id") or chat_id
        author_num = "".join(filter(str.isdigit, str(author_raw).split("@")[0]))
        pushname = data.get("pushname") or (data.get("sender") or {}).get("pushname") or ""
    elif "@" in str(chat_id):
        author_num = "".join(filter(str.isdigit, str(chat_id).split("@")[0]))

    prof_chave, prof_nome = identificar_prof_solicitante(author_num, str(author_raw), texto, pushname)
    if not prof_chave:
        prof_chave = "pedro"

    t_lower = normalizar_texto(texto)
    eh_modo_consulta = t_lower.startswith(("!atividades", "/atividades", "atividades", "!vistos", "/vistos", "vistos"))

    corpo = re.sub(r'^(?:[!/]?(?:atividades|atividade|vistos|visto|atv|corrigir|correcao|correção)\s*)+', '', texto, flags=re.IGNORECASE).strip()
    tem_foto = bool(data and (data.get("hasMedia") or data.get("mimetype") or (data.get("quotedMsg") and (data["quotedMsg"].get("hasMedia") or data["quotedMsg"].get("mimetype")))))
    if (not corpo and not tem_foto and not eh_modo_consulta) or corpo.lower() in ["ajuda", "help", "?"]:
        hint_comp = "MKT" if prof_chave == "bruna" else ("WEB" if prof_chave == "marcos" else "IA")
        hint_tm = "201" if prof_chave == "bruna" else ("102" if prof_chave == "marcos" else "202")
        return (
            "📝 *GESTÃO DE ATIVIDADES, VISTOS EM MASSA E CORREÇÃO POR FOTO*\n\n"
            "1️⃣ *Lançar TODOS da turma (com exceções):*\n"
            f"  `!atv [#ID] <turma> todos exceto <nº ou nomes>`\n"
            f"  • Ex: `!atv {hint_tm} todos exceto 2, 3, 10`\n"
            f"  • Ex: `!atv {hint_tm} todos menos ana sophia, davy`\n"
            f"  • Ex: `!atv {hint_tm} todos ok` (100% entregaram)\n\n"
            "2️⃣ *Lançar vários alunos ou individual:*\n"
            f"  • Lote: `!atv {hint_tm} 1, 4, 5, 8-12 ok`\n"
            f"  • Único: `!atv {hint_tm} 14 ok` ou `!atv {hint_tm} joao ok`\n\n"
            "3️⃣ *Corrigir / Dar visto por FOTO (IA Vision):*\n"
            f"  • Envie a foto com legenda `!atv {hint_tm} 14 ok`\n\n"
            "4️⃣ *Indicar, consultar ou mover atividades (#ID):*\n"
            f"  • Ver IDs e placares: `!atv lista` ou `!atv {hint_tm}`\n"
            f"  • Mover atv para outra turma: `!atv mover #64 202`\n"
            f"  • Mover/copiar vistos entre atvs: `!atv mover #5 #6` ou `!atv copiar #5 #6`\n"
            f"  • Cobrar atividade: `!cobrar {hint_tm}`\n"
            f"  • Cadastrar nova: `!atv nova {hint_tm} {hint_comp} <Título> [data]`"
        )

    # Subcomandos específicos
    c_lower = normalizar_texto(corpo)
    if c_lower.startswith(("cobrar", "pendente ", "pendentes")):
        resto = re.sub(r'^(cobrar|pendente|pendentes)\s*', '', corpo, flags=re.IGNORECASE).strip()
        return cobrar_atividades_turma(resto, prof_chave=prof_chave)

    if c_lower.startswith(("nova", "criar", "cadastrar", "passar")):
        resto = re.sub(r'^(nova|criar|cadastrar|passar)\s*', '', corpo, flags=re.IGNORECASE).strip()
        return cadastrar_nova_atividade(resto, prof_chave=prof_chave)

    if c_lower.startswith(("lista", "listar", "cadastradas", "ids", "#")):
        if c_lower.startswith(("lista", "listar", "cadastradas", "ids")):
            resto = re.sub(r'^(lista|listar|cadastradas|ids)\s*', '', corpo, flags=re.IGNORECASE).strip()
            return listar_atividades_cadastradas(resto, prof_chave=prof_chave)

    if c_lower.startswith(("mover", "transferir", "migrar")):
        return transferir_ou_copiar_vistos_atividade(corpo, mover=True)

    if c_lower.startswith(("copiar", "duplicar")):
        return transferir_ou_copiar_vistos_atividade(corpo, mover=False)

    if c_lower.startswith(("turma", "mudar turma", "alterar turma")):
        return mover_atividade_entre_turmas(corpo)

    # Se for consulta explícita ('atividades 202' ou 'vistos')
    if eh_modo_consulta:
        if not corpo:
            return listar_atividades_cadastradas("", prof_chave=prof_chave)
        if corpo.strip() in TURMAS_VALIDAS_ATV:
            return consultar_atividades_turma(corpo.strip())
        return consultar_atividades_aluno(corpo)

    # Verifica se é o modo em massa "TODOS [fizeram/ok] [EXCETO/MENOS ...]"
    m_todos = re.search(
        r"\b(todos|todo\s+mundo|turma\s+toda|geral)\b(?:\s+(?:fizeram|entregaram|ok|visto|feitos))?(?:\s*(?:,|;|-)?\s*\b(?:exceto|menos|salvo|sem|fora|tirando|excecao|exceção|nao\s+fizeram|não\s+fizeram|faltou|faltaram)\b[:\s]*(.*))?$",
        corpo,
        flags=re.IGNORECASE | re.DOTALL
    )
    if m_todos:
        prefixo_antes_todos = corpo[:m_todos.start()].strip()
        texto_excecoes = (m_todos.group(2) or "").strip()
        toks_pref = prefixo_antes_todos.split()
        ativ_id_expl, data_expl, termo_titulo, toks_pref_rest = _extrair_seletor_atividade(toks_pref)

        # Também verifica se o usuário colocou #ID ou turma depois de 'todos' quando não há exceções
        if not texto_excecoes:
            sufixo_pos = corpo[m_todos.end():].strip()
            if sufixo_pos:
                a2, d2, tt2, r2 = _extrair_seletor_atividade(sufixo_pos.split())
                if a2 is not None and ativ_id_expl is None:
                    ativ_id_expl = a2
                if d2 and not data_expl:
                    data_expl = d2
                toks_pref_rest.extend(r2)

        turma_alvo = ""
        comp_alvo = ""
        for t in toks_pref_rest:
            t_clean = t.strip(",;.")
            if t_clean in TURMAS_VALIDAS_ATV and not turma_alvo:
                turma_alvo = t_clean
            elif normalizar_texto(t_clean) in COMP_MAP_ATV and not comp_alvo:
                comp_alvo = COMP_MAP_ATV[normalizar_texto(t_clean)]

        return _executar_visto_em_massa_todos_exceto(
            turma_alvo, comp_alvo, ativ_id_expl, data_expl, termo_titulo, texto_excecoes, corpo, prof_chave=prof_chave
        )

    # Extrai seletores de atividade (#ID, data, tema) do comando normal
    tokens_raw = corpo.split()
    ativ_id_expl, data_expl, termo_titulo, tokens = _extrair_seletor_atividade(tokens_raw)

    # Se o comando for apenas a turma (ex: 'atv 202') ou apenas '#ID' / '202 #ID' (sem foto)
    if not tem_foto:
        if len(tokens) == 1 and tokens[0] in TURMAS_VALIDAS_ATV and ativ_id_expl is None and not data_expl:
            return consultar_atividades_turma(tokens[0])
        if len(tokens) == 0 and ativ_id_expl is not None:
            return cobrar_atividades_turma(f"#{ativ_id_expl}", prof_chave=prof_chave)
        if len(tokens) == 1 and tokens[0] in TURMAS_VALIDAS_ATV and (ativ_id_expl is not None or data_expl):
            sel_s = f"#{ativ_id_expl}" if ativ_id_expl is not None else data_expl
            return cobrar_atividades_turma(f"{tokens[0]} {sel_s}", prof_chave=prof_chave)

    STATUS_KEYWORDS = {"ok", "visto", "entregue", "feito", "pendente", "incompleto", "10", "recusado"}

    # Extrai turma se presente
    turma_alvo = ""
    tokens_restantes = []
    for t in tokens:
        t_clean = t.strip(",;.")
        if t_clean in TURMAS_VALIDAS_ATV and not turma_alvo:
            turma_alvo = t_clean
        else:
            tokens_restantes.append(t)

    # Extrai componente curricular se presente
    comp_alvo = ""
    tokens_restantes_2 = []
    for t in tokens_restantes:
        t_low = normalizar_texto(t.strip(",;."))
        if t_low in COMP_MAP_ATV and not comp_alvo:
            comp_alvo = COMP_MAP_ATV[t_low]
        else:
            tokens_restantes_2.append(t)

    # Extrai status (ex: 'ok', 'visto', 'pendente')
    status_alvo = "VISTO (OK)"
    status_explicito = False
    tokens_ident = []
    for t in tokens_restantes_2:
        t_low = t.strip(",;.").lower()
        if t_low in STATUS_KEYWORDS and not status_explicito:
            status_explicito = True
            if t_low == "pendente":
                status_alvo = "PENDENTE"
            elif t_low == "incompleto":
                status_alvo = "INCOMPLETO"
            elif t_low == "recusado":
                status_alvo = "RECUSADO"
            else:
                status_alvo = "VISTO (OK)"
        else:
            tokens_ident.append(t)

    identificador = " ".join(tokens_ident).strip()

    # Se enviou foto com comando !visto atividade <tema> (sem status explícito 'ok'), verifica se o texto restante é o tema da atividade e não o nome de um aluno
    if tem_foto and identificador and not status_explicito and not termo_titulo:
        id_norm = normalizar_texto(identificador)
        eh_tema_atv = False
        try:
            with sqlite3.connect(ESCOLA_DB) as conn_chk:
                cur_chk = conn_chk.cursor()
                cur_chk.execute("SELECT titulo, turma FROM atividades_cadastradas ORDER BY id DESC")
                for tit_db, tm_db in cur_chk.fetchall():
                    if id_norm in normalizar_texto(tit_db or ""):
                        eh_tema_atv = True
                        if not turma_alvo and tm_db:
                            turma_alvo = str(tm_db)
                        break
            if not eh_tema_atv and os.path.exists(AULAS_CSV):
                with open(AULAS_CSV, "r", encoding="utf-8") as f_csv:
                    for row_c in csv.DictReader(f_csv, delimiter=";"):
                        if id_norm in normalizar_texto(row_c.get("tema") or ""):
                            eh_tema_atv = True
                            if not turma_alvo and row_c.get("turma"):
                                turma_alvo = str(row_c.get("turma")).strip()
                            break
        except Exception:
            pass
        if eh_tema_atv:
            termo_titulo = identificador
            identificador = ""

    # Verifica se é um LOTE de múltiplos alunos (ex: 'atv #6 202 1, 4, 8-12 ok' ou 'atv 202 joao, maria ok')
    if identificador and not tem_foto:
        tem_separador_lote = any(sep in identificador for sep in [",", ";"]) or bool(re.search(r"\b\d{1,2}\s*-\s*\d{1,2}\b", identificador))
        toks_num_espaco = identificador.split()
        eh_lista_numeros = len(toks_num_espaco) >= 2 and all(re.match(r"^#?\d{1,2}(?:-\d{1,2})?$", tk) for tk in toks_num_espaco)
        if tem_separador_lote or eh_lista_numeros:
            itens_lote = _separar_itens_alunos(identificador)
            if len(itens_lote) >= 1:
                return _executar_visto_lote_alunos(
                    turma_alvo, comp_alvo, ativ_id_expl, data_expl, termo_titulo,
                    itens_lote, status_alvo, corpo, prof_chave=prof_chave
                )

    # Se não informou o aluno no texto, tenta identificar pelo remetente da foto (ou quotedMsg)
    if not identificador and data:
        quoted = data.get("quotedMsg") or {}
        q_author = extrair_identificador_limpo(quoted.get("author") or quoted.get("from") or "")
        author_num = extrair_identificador_limpo(data.get("author") or data.get("from") or "")
        num_pesquisa = q_author or author_num
        if num_pesquisa:
            try:
                with sqlite3.connect(ESCOLA_DB) as conn_v:
                    r_v = conn_v.execute("SELECT nome, turma FROM alunos_whatsapp_vinculos WHERE numero = ?", (num_pesquisa,)).fetchone()
                    if r_v:
                        identificador = r_v[0]
                        if not turma_alvo:
                            turma_alvo = r_v[1]
            except Exception:
                pass

    # Se enviou foto sem identificar o aluno: analisa a foto com contexto da atividade e orienta comando com #ID
    if not identificador and tem_foto:
        if DESABILITAR_LEITURA_IMAGEM_OMNIROUTE:
            log_terminal("🛡️ [SEGURANÇA] Foto recebida sem comando: leitura automática via OmniRoute desabilitada (silencioso).")
            return ""
        midia_path = obter_midia_recente(data or {}, chat_id, permitir_cache_chat=True)
        turma_ref = turma_alvo or obter_turma_atual_ou_proxima() or "202"
        ativ_ref_id = None
        ativ_ref_tit = "Exercício de Caderno"
        with sqlite3.connect(ESCOLA_DB) as conn_f:
            cur_f = conn_f.cursor()
            r_af = _resolver_atividade_alvo(cur_f, turma_ref, comp_alvo, ativ_id_expl, data_expl, termo_titulo, preferir_com_entregas=False)
            if r_af:
                ativ_ref_id, _, turma_ref, _, ativ_ref_tit = r_af

        id_cmd_str = f"#{ativ_ref_id} " if ativ_ref_id else (f"{data_expl} " if data_expl else "")
        id_tit_str = f"#{ativ_ref_id} — " if ativ_ref_id else ""
        if midia_path and os.path.exists(midia_path):
            avaliacao_vision = analisar_foto_caderno_ia(midia_path, ativ_ref_tit)
            if chat_id in ULTIMA_MIDIA_CHAT:
                ULTIMA_MIDIA_CHAT[chat_id]["avaliacao_ia"] = avaliacao_vision
                ULTIMA_MIDIA_CHAT[chat_id]["usada_em_visto"] = False
            return (
                "📸 *Foto de caderno recebida e analisada!*\n\n"
                f"📌 *Atividade de referência:* *{id_tit_str}{ativ_ref_tit}* (Turma {turma_ref})\n\n"
                f"🔍 *Análise do Caderno (IA Vision):*\n{avaliacao_vision}\n\n"
                "👉 *Para gravar o visto desta foto agora, basta digitar:*\n"
                f"• `atv {id_cmd_str}{turma_ref} <nº ou nome> ok` (Ex: `atv {id_cmd_str}{turma_ref} 6 ok`)"
            )
        return (
            "📸 *Foto recebida para correção!*\n\n"
            f"📌 *Atividade de referência:* *{id_tit_str}{ativ_ref_tit}* (Turma {turma_ref})\n"
            "Por favor, informe o aluno ou número da chamada.\n"
            f"Exemplo: `atv {id_cmd_str}{turma_ref} 6 ok`"
        )

    if not turma_alvo and ativ_id_expl is not None:
        with sqlite3.connect(ESCOLA_DB) as conn_t:
            r_t = conn_t.execute("SELECT turma FROM atividades_cadastradas WHERE id = ?", (ativ_id_expl,)).fetchone()
            if r_t:
                turma_alvo = r_t[0]

    if not turma_alvo:
        turma_atual = obter_turma_atual_ou_proxima()
        if turma_atual:
            turma_alvo = turma_atual

    aluno_achado = None
    num_chamada = None
    with sqlite3.connect(ESCOLA_DB) as conn:
        cur = conn.cursor()
        if turma_alvo:
            cur.execute("SELECT id, matricula, nome, turma, componente FROM alunos WHERE turma = ? ORDER BY nome ASC", (turma_alvo,))
            lista_t = cur.fetchall()
            alunos_num = [(idx, r[0], r[1], r[2], r[3], r[4]) for idx, r in enumerate(lista_t, 1)]
            st_b, matches = _buscar_alunos_por_termo_turma(identificador, alunos_num)
            if st_b == "OK" and len(matches) == 1:
                m = matches[0]
                num_chamada = m[0]
                aluno_achado = (m[1], m[2], m[3], m[4], m[5])
            elif st_b == "AMBIGUO":
                id_hint = f"#{ativ_id_expl} " if ativ_id_expl else ""
                linhas_dup = [f"• *#{m[0]:02d}* | *{m[3]}* (Turma {m[4]})" for m in matches[:6]]
                return (
                    f"⚠️ Encontrado mais de um aluno correspondente a '{identificador}' na Turma {turma_alvo}:\n" +
                    "\n".join(linhas_dup) +
                    f"\n\n👉 Especifique pelo *nº da chamada*: `atv {id_hint}{turma_alvo} {matches[0][0]} ok`"
                )

        if not aluno_achado and not identificador.isdigit():
            cur.execute("SELECT id, matricula, nome, turma, componente FROM alunos ORDER BY nome ASC")
            todos = cur.fetchall()
            alunos_todos = [(idx, r[0], r[1], r[2], r[3], r[4]) for idx, r in enumerate(todos, 1)]
            st_b, matches = _buscar_alunos_por_termo_turma(identificador, alunos_todos)
            if st_b == "OK" and len(matches) == 1:
                m = matches[0]
                aluno_achado = (m[1], m[2], m[3], m[4], m[5])
            elif st_b == "AMBIGUO":
                id_hint = f"#{ativ_id_expl} " if ativ_id_expl else ""
                linhas_dup = [f"• *{m[3]}* (Turma {m[4]})" for m in matches[:6]]
                return (
                    f"⚠️ Encontrado mais de um aluno correspondente a '{identificador}':\n" +
                    "\n".join(linhas_dup) +
                    f"\n\n👉 Especifique com a turma ou nº da chamada. Ex: `atv {id_hint}{matches[0][4]} {matches[0][3].split()[0]} ok`"
                )

    if not aluno_achado:
        if not status_explicito and ativ_id_expl is None:
            res_busca = consultar_atividades_aluno(identificador)
            if "Nenhuma atividade" not in res_busca:
                return res_busca
        turma_info = f" na Turma {turma_alvo}" if turma_alvo else ""
        return f"🔍 Aluno '{identificador}' não encontrado{turma_info}. Verifique os números com `!turma {turma_alvo or '202'}`."

    id_db, matricula, nome, turma_final, comp_cadastrado = aluno_achado

    if not comp_alvo:
        ctx = obter_contexto_aula_atual()
        if ctx.get("tem_aula") and ctx.get("turma") == turma_final and ctx.get("disciplina"):
            comp_alvo = ctx.get("disciplina")
        elif comp_cadastrado:
            try:
                wd_hoje = datetime.now().weekday()
                with sqlite3.connect(ESCOLA_DB) as conn_g:
                    cur_g = conn_g.cursor()
                    cur_g.execute("SELECT disciplina FROM grade_horaria WHERE dia_semana = ? AND turma = ? LIMIT 1", (wd_hoje, turma_final))
                    row_g = cur_g.fetchone()
                    comp_alvo = row_g[0] if row_g else comp_cadastrado.split(",")[0].strip()
            except Exception:
                comp_alvo = comp_cadastrado.split(",")[0].strip()
        else:
            comp_alvo = "INFORMÁTICA"

    agora = datetime.now()
    data_reg = agora.strftime("%d/%m/%Y às %H:%M")
    data_iso = agora.strftime("%Y-%m-%d %H:%M:%S")
    atividade_desc = "Atividade de Aula / Visto de Caderno"
    atividade_id_vinc = None
    data_passada_info = ""

    # Obtém mídia apenas se enviada nesta mensagem/reply OU se houver uma foto recém-analisada no chat ainda não vinculada a um visto
    midia_path = None
    avaliacao_vision = ""
    if tem_foto:
        midia_path = obter_midia_recente(data or {}, chat_id, permitir_cache_chat=False)
    elif chat_id in ULTIMA_MIDIA_CHAT:
        reg_m = ULTIMA_MIDIA_CHAT[chat_id]
        if (not reg_m.get("usada_em_visto")) and (time.time() - reg_m.get("timestamp", 0) < 300) and os.path.exists(reg_m.get("filepath", "")):
            midia_path = reg_m["filepath"]
            avaliacao_vision = reg_m.get("avaliacao_ia") or ""
            reg_m["usada_em_visto"] = True

    with sqlite3.connect(ESCOLA_DB) as conn:
        cur = conn.cursor()
        if num_chamada is None:
            cur.execute("SELECT matricula FROM alunos WHERE turma = ? ORDER BY nome ASC", (turma_final,))
            mats = [r[0] for r in cur.fetchall()]
            num_chamada = mats.index(matricula) + 1 if matricula in mats else None

        row_ativ = _resolver_atividade_alvo(
            cur, turma_final, comp_alvo, ativ_id_expl, data_expl, termo_titulo, preferir_com_entregas=False, prof_chave=prof_chave
        )
        if row_ativ:
            atividade_id_vinc = row_ativ[0]
            data_passada_info = f" (Passada em: {row_ativ[1]})"
            if row_ativ[3]:
                comp_alvo = row_ativ[3]
            atividade_desc = row_ativ[4]

        if midia_path and os.path.exists(midia_path) and not avaliacao_vision:
            avaliacao_vision = analisar_foto_caderno_ia(midia_path, atividade_desc)
            if chat_id in ULTIMA_MIDIA_CHAT:
                ULTIMA_MIDIA_CHAT[chat_id]["usada_em_visto"] = True

        # Se veio de comando 'corrigir' com foto sem 'ok' explícito, ajusta status conforme parecer da IA
        if midia_path and avaliacao_vision and not status_explicito:
            av_low = normalizar_texto(avaliacao_vision)
            if "parecer: nao e caderno" in av_low or "parecer:** nao e caderno" in av_low:
                status_alvo = "PENDENTE (FOTO INVÁLIDA)"
            elif "parecer: incompleto" in av_low or "parecer:** incompleto" in av_low:
                status_alvo = "INCOMPLETO"

        _gravar_visto_unico_db(
            cur, atividade_id_vinc, data_reg, data_iso, turma_final, comp_alvo,
            matricula, nome, atividade_desc, status_alvo, corpo, midia_path or "", avaliacao_vision, prof_chave=prof_chave or "pedro"
        )
        conn.commit()

        # Vincula o número do WhatsApp do aluno se a mensagem ou citação veio dele
        quoted = (data or {}).get("quotedMsg") or {}
        q_author = extrair_identificador_limpo(quoted.get("author") or quoted.get("from") or "")
        author_num = extrair_identificador_limpo((data or {}).get("author") or (data or {}).get("from") or "")
        target_num = q_author or author_num
        if target_num and target_num not in IDENTIFICADORES_AUTORIZADOS and len(target_num) >= 10:
            try:
                cur.execute("INSERT OR REPLACE INTO alunos_whatsapp_vinculos (numero, matricula, nome, turma) VALUES (?, ?, ?, ?)", (target_num, matricula, nome, turma_final))
                conn.commit()
            except Exception:
                pass

        resumo_ativs = _resumo_atividades_turma_str(cur, turma_final, atividade_id_vinc, prof_chave=prof_chave)

    chamada_str = f"#{num_chamada:02d} " if num_chamada else ""
    visao_str = f"\n🔍 *Avaliação do Caderno (IA Vision):*\n{avaliacao_vision}\n" if avaliacao_vision else ""
    titulo_header = "📸 *VISTO / CORREÇÃO POR FOTO REGISTRADA!*" if midia_path else "✅ *VISTO / ATIVIDADE REGISTRADA!*"
    id_ativ_str = f"#{atividade_id_vinc} — " if atividade_id_vinc else ""
    cmd_id_str = f" #{atividade_id_vinc}" if atividade_id_vinc else ""

    return (
        f"{titulo_header}\n\n"
        f"👤 *Aluno:* {chamada_str}*{nome}*\n"
        f"🏫 *Turma:* {turma_final} | 📖 *Componente:* {comp_alvo}\n"
        f"📝 *Atividade:* *{id_ativ_str}{atividade_desc}*{data_passada_info}\n"
        f"{visao_str}"
        f"📅 *Data do Registro:* {data_reg}\n"
        f"📝 *Status:* {status_alvo}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"{resumo_ativs}\n\n"
        f"📊 *Para cobrar esta atividade:* `!cobrar {turma_final}{cmd_id_str}`"
    )

def consultar_atividades_turma(turma: str) -> str:
    init_tabelas_atividades()
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            resumo_ativs = _resumo_atividades_turma_str(cur, turma)
            cur.execute("""
                SELECT atividade_id, data_registro, nome_aluno, componente, status, atividade 
                FROM atividades_entregas 
                WHERE turma = ? 
                ORDER BY id DESC LIMIT 25
            """, (turma,))
            linhas = cur.fetchall()

        if not linhas and not resumo_ativs:
            return f"📝 *Nenhuma atividade/visto registrado para a Turma {turma}.*\nPara registrar: `!atv #ID {turma} todos exceto 2, 3`"

        resp = []
        if resumo_ativs:
            resp.append(resumo_ativs)
            resp.append("━━━━━━━━━━━━━━━━━━━━")

        if linhas:
            resp.append(f"📋 *ÚLTIMOS VISTOS REGISTRADOS — TURMA {turma}* ({len(linhas)} recentes):\n")
            for aid, dt, nome, comp, st, ativ in linhas[:15]:
                tag_id = f"[#{aid}] " if aid else ""
                resp.append(f"• *{nome}*: {st} — {tag_id}{ativ} (_{dt}_)")
        else:
            resp.append(f"ℹ️ Nenhum visto individual lançado ainda na Turma {turma}.")

        resp.append(
            f"\n💡 *Comandos rápidos:*\n"
            f"• Todos menos alguns: `!atv #<ID> {turma} todos exceto 2, 5`\n"
            f"• Relatório de pendentes: `!cobrar {turma} #<ID>`"
        )
        return "\n".join(resp)
    except Exception as e:
        return f"❌ Erro ao consultar atividades: {e}"

def consultar_atividades_aluno(termo_aluno: str) -> str:
    init_tabelas_atividades()
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT atividade_id, data_registro, turma, componente, nome_aluno, status, atividade 
                FROM atividades_entregas 
                ORDER BY id DESC LIMIT 150
            """)
            todas = cur.fetchall()

        termo_fon = normalizar_fonetico(termo_aluno)
        linhas = []
        for aid, dt, t, comp, nome, st, ativ in todas:
            if termo_fon in normalizar_fonetico(nome) or termo_aluno.lower() in str(dt).lower() or termo_aluno.lower() in str(comp).lower() or termo_aluno.lower() in str(ativ).lower():
                linhas.append((aid, dt, t, comp, nome, st, ativ))

        if not linhas:
            return f"📝 Nenhuma atividade encontrada para '{termo_aluno}'."

        resp = [f"📋 *ATIVIDADES / VISTOS DE {termo_aluno.upper()}* ({len(linhas)} registros):\n"]
        for aid, dt, t, comp, nome, st, ativ in linhas[:20]:
            tag_id = f"#{aid} " if aid else ""
            resp.append(f"• *{nome}* (Turma {t}): {st} - {tag_id}{ativ} ({comp}) em _{dt}_")

        return "\n".join(resp)
    except Exception as e:
        return f"❌ Erro ao consultar atividades do aluno: {e}"

# --- GESTÃO DE ALARMES E AVISOS PRIVADOS ---

def parse_alarme(texto: str):
    agora = datetime.now()
    t = texto.strip()
    
    # 1. Padrão relativo: 'em X min', '10m', '10min', '1h', '30s'
    m_rel = re.match(r'^(?:em\s+)?(\d+)\s*(m|min|minuto|minutos|h|hora|horas|s|seg|segundos)\b(?:\s+(.*))?$', t, re.IGNORECASE)
    if m_rel:
        val = int(m_rel.group(1))
        unid = m_rel.group(2).lower()
        conteudo = (m_rel.group(3) or "").strip()
        if 's' in unid:
            dt = agora + timedelta(seconds=val)
        elif 'h' in unid:
            dt = agora + timedelta(hours=val)
        else:
            dt = agora + timedelta(minutes=val)
        return dt, conteudo

    # 2. Padrão data e hora: '16/10 08:30 Mensagem' ou '16/10/2026 08:30 Mensagem'
    m_data_hora = re.match(r'^(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s+(\d{1,2}:\d{2})\s*(.*)$', t)
    if m_data_hora:
        d_str = m_data_hora.group(1)
        h_str = m_data_hora.group(2)
        conteudo = m_data_hora.group(3).strip()
        partes_d = d_str.split('/')
        dia = int(partes_d[0])
        mes = int(partes_d[1])
        ano = int(partes_d[2]) if len(partes_d) > 2 else agora.year
        if ano < 100: ano += 2000
        hora, minuto = map(int, h_str.split(':'))
        try:
            dt = datetime(ano, mes, dia, hora, minuto, 0)
            return dt, conteudo
        except Exception:
            pass

    # 3. Padrão 'amanha 08:30 Mensagem'
    m_amanha = re.match(r'^(?:amanha|amanhã)\s+(\d{1,2}:\d{2})\s*(.*)$', t, re.IGNORECASE)
    if m_amanha:
        h_str = m_amanha.group(1)
        conteudo = m_amanha.group(2).strip()
        hora, minuto = map(int, h_str.split(':'))
        dt = (agora + timedelta(days=1)).replace(hour=hora, minute=minuto, second=0, microsecond=0)
        return dt, conteudo

    # 4. Padrão apenas hora: '19:30 Mensagem'
    m_hora = re.match(r'^(\d{1,2}:\d{2})\s*(.*)$', t)
    if m_hora:
        h_str = m_hora.group(1)
        conteudo = m_hora.group(2).strip()
        hora, minuto = map(int, h_str.split(':'))
        dt = agora.replace(hour=hora, minute=minuto, second=0, microsecond=0)
        if dt <= agora:
            dt += timedelta(days=1)
        return dt, conteudo

    return None, t

def adicionar_alarme(texto_comando: str, chat_id: str = DONO_CHAT_ID) -> str:
    t = texto_comando.strip()
    if not t:
        return (
            "⏰ *Como Usar o Alarme / Aviso:* \n\n"
            "• `!alarme <tempo/hora> <mensagem>`\n"
            "Exemplos:\n"
            "▫️ `!alarme 10m Ligar para o coordenador`\n"
            "▫️ `!alarme 19:30 Fechar diário da Turma 202`\n"
            "▫️ `!alarme amanhã 07:30 Conferir chamada`\n"
            "▫️ `!alarme 16/10 08:00 Iniciar prova de IA`\n\n"
            "• `!alarmes` -> Listar alarmes pendentes\n"
            "• `!alarme cancelar <id>` -> Cancelar alarme"
        )

    eh_privado = False
    if t.lower().startswith(("privado ", "solo ")):
        eh_privado = True
        t = t.split(" ", 1)[1].strip()
    elif "[privado]" in t.lower() or "#privado" in t.lower():
        eh_privado = True

    t_lower = t.lower()
    if t_lower in ["listar", "lista", "ver", "ativos"]:
        return listar_alarmes()

    if t_lower.startswith(("cancelar", "remover", "excluir", "del", "rm")):
        partes = t.split()
        if len(partes) > 1 and partes[1].isdigit():
            return cancelar_alarme(partes[1])
        return "Use: `!alarme cancelar <id>` (consulte os IDs em `!alarmes`)."

    dt_alvo, conteudo = parse_alarme(t)
    if not dt_alvo:
        return (
            "⚠️ Formato de horário não reconhecido.\n"
            "Exemplos válidos:\n"
            "• `!alarme 10m Aviso aqui` (relativo em minutos)\n"
            "• `!alarme privado 10m Aviso exclusivo` (não espelha)\n"
            "• `!alarme 1h Aviso aqui` (em horas)\n"
            "• `!alarme 19:30 Aviso aqui` (no horário específico)\n"
            "• `!alarme amanhã 07:30 Aviso aqui`"
        )

    if not conteudo:
        conteudo = "Alarme programado"

    if eh_privado and not conteudo.startswith("[PRIVADO]"):
        conteudo = f"[PRIVADO] {conteudo}"

    target_chat = chat_id if chat_id else DONO_CHAT_ID
    h_iso = dt_alvo.strftime("%Y-%m-%d %H:%M:%S")

    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO agendamentos (prompt, horario_iso, status, chat_id)
                VALUES (?, ?, 'PENDENTE', ?)
            """, (conteudo, h_iso, target_chat))
            conn.commit()
            novo_id = cur.lastrowid
    except Exception as e:
        return f"❌ Erro ao salvar alarme no banco: {e}"

    diff = dt_alvo - datetime.now()
    total_seg = int(diff.total_seconds())
    if total_seg < 0:
        total_seg = 0
    horas, resto_s = divmod(total_seg, 3600)
    minutos, segundos = divmod(resto_s, 60)
    
    if horas > 0:
        tempo_str = f"em {horas}h {minutos}min"
    elif minutos > 0:
        tempo_str = f"em {minutos}min"
    else:
        tempo_str = f"em {segundos}s"

    dia_semana = DIAS_NOMES[dt_alvo.weekday()]

    return (
        f"⏰ *Alarme Programado com Sucesso!* (ID `{novo_id}`)\n\n"
        f"📢 *Aviso:* \"{conteudo}\"\n"
        f"🕒 *Horário:* {dt_alvo.strftime('%H:%M:%S')} ({tempo_str})\n"
        f"📅 *Data:* {dt_alvo.strftime('%d/%m/%Y')} ({dia_semana})\n"
        + ("🔒 *Modo Privado:* Este alarme não será espelhado para outros contatos.\n\n" if eh_privado else "\n") +
        f"🔔 O sistema enviará uma mensagem no seu privado assim que o horário for atingido."
    )

def listar_alarmes() -> str:
    if not os.path.exists(ESCOLA_DB):
        return "⚠️ Banco escola.db não encontrado."
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, prompt, horario_iso 
                FROM agendamentos 
                WHERE status = 'PENDENTE' 
                ORDER BY horario_iso ASC
            """)
            itens = cur.fetchall()

        if not itens:
            return "🔕 Você não tem nenhum alarme pendente no momento.\nUse `!alarme <tempo/hora> <conteúdo>` para adicionar."

        resp = ["⏰ *SEUS ALARMES PENDENTES:*\n"]
        agora = datetime.now()
        for id_ag, prompt, h_iso in itens:
            try:
                dt_obj = datetime.strptime(h_iso, "%Y-%m-%d %H:%M:%S")
                diff = dt_obj - agora
                total_seg = int(diff.total_seconds())
                if total_seg > 0:
                    h, r = divmod(total_seg, 3600)
                    m, s = divmod(r, 60)
                    tempo_str = f"daqui a {h}h{m}m" if h > 0 else f"daqui a {m}m"
                else:
                    tempo_str = "disparando agora"
                h_fmt = dt_obj.strftime("%H:%M (%d/%m)")
            except Exception:
                h_fmt = h_iso
                tempo_str = ""

            resp.append(f"• `ID {id_ag}` | 🕒 *{h_fmt}* ({tempo_str})\n  📢 \"{prompt}\"")

        resp.append("\n━━━━━━━━━━━━━━━━━━━━")
        resp.append("💡 Para cancelar: `!alarme cancelar <id>`")
        return "\n".join(resp)
    except Exception as e:
        return f"❌ Erro ao listar alarmes: {e}"

def cancelar_alarme(id_arg: str) -> str:
    try:
        id_int = int(id_arg)
    except ValueError:
        return f"❌ ID inválido: '{id_arg}'. Forneça um número inteiro."

    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("SELECT prompt, horario_iso FROM agendamentos WHERE id = ? AND status = 'PENDENTE'", (id_int,))
            item = cur.fetchone()
            if not item:
                return f"🔍 Nenhum alarme pendente encontrado com o ID `{id_int}`."
            cur.execute("UPDATE agendamentos SET status = 'CANCELADO' WHERE id = ?", (id_int,))
            conn.commit()
        return f"✅ Alarme `[ID {id_int}]` cancelado com sucesso!\n📢 \"{item[0]}\" ({item[1]})"
    except Exception as e:
        return f"❌ Erro ao cancelar alarme: {e}"

def loop_verificador_alarmes():
    """Verifica e dispara alarmes a cada 5 segundos."""
    import time
    log_terminal("⏰ Loop verificador de alarmes ativado (verificação a cada 5s).")
    while True:
        try:
            agora_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if os.path.exists(ESCOLA_DB):
                with sqlite3.connect(ESCOLA_DB) as conn:
                    cur = conn.cursor()
                    cur.execute("""
                        SELECT id, prompt, horario_iso, COALESCE(chat_id, '')
                        FROM agendamentos 
                        WHERE status = 'PENDENTE' AND horario_iso <= ?
                    """, (agora_str,))
                    disparos = cur.fetchall()
                    for id_ag, prompt, h_iso, target_chat in disparos:
                        # Bloqueio atômico para evitar disparos duplicados
                        cur.execute("UPDATE agendamentos SET status = 'DISPARANDO' WHERE id = ? AND status = 'PENDENTE'", (id_ag,))
                        conn.commit()
                        if cur.rowcount > 0:
                            log_terminal(f"⏰ DISPARANDO ALARME [ID {id_ag}]: '{prompt}' para {target_chat}")
                            msg = (
                                f"⏰ *ALARME / AVISO DOCENTE* 🚨\n\n"
                                f"📢 *Lembrete:* {prompt}\n\n"
                                f"🕒 Horário: {datetime.now().strftime('%H:%M:%S')} ({datetime.now().strftime('%d/%m/%Y')})"
                            )
                            enviou = enviar_zap(target_chat, msg)
                            contato_restrito.espelhar_alarme(prompt, target_chat, enviar_zap)
                            st = "ENVIADO" if enviou else "ERRO_ENVIO"
                            cur.execute("UPDATE agendamentos SET status = ? WHERE id = ?", (st, id_ag))
                            conn.commit()
        except Exception as e:
            log_terminal(f"⚠️ Erro no loop_verificador_alarmes: {e}")
        time.sleep(5)

def registrar_diario(turma_ou_texto: str, conteudo_opcional: str = None) -> str:
    """
    Registra o conteúdo ministrado no diário de bordo e no escola.db.
    Entende automaticamente em qual sala/turma o docente está e qual a disciplina de hoje.
    """
    if conteudo_opcional is not None:
        texto_args = f"{turma_ou_texto} {conteudo_opcional}".strip()
    else:
        texto_args = (turma_ou_texto or "").strip()

    ctx = obter_contexto_aula_atual()

    if not texto_args:
        if ctx.get("tem_aula"):
            return (
                f"ℹ️ *Como registrar o diário hoje:*\n"
                f"Sua aula ativa/próxima é na *Turma {ctx['turma']}* ({ctx['disciplina']}).\n"
                f"Basta digitar: `!diario <conteúdo ministrado>` (a turma {ctx['turma']} será atribuída automaticamente).\n"
                f"Ou especifique: `!diario <outra_turma> <conteúdo>`."
            )
        return "⚠️ Uso: `!diario [Turma] <Conteúdo Ministrado>`\nExemplo: `!diario Introdução ao Marco Civil da Internet`"

    partes = texto_args.split(maxsplit=1)
    TURMAS_VALIDAS = {"101", "102", "103", "200", "201", "202", "203", "300", "301", "302", "303"}

    if partes[0] in TURMAS_VALIDAS and len(partes) > 1:
        turma = partes[0]
        conteudo = partes[1].strip()
    elif partes[0] in TURMAS_VALIDAS and len(partes) == 1:
        return f"⚠️ Digite o conteúdo da aula da Turma {partes[0]}: `!diario {partes[0]} <Conteúdo>`"
    else:
        turma = ctx.get("turma") if ctx.get("tem_aula") else "202"
        conteudo = texto_args

    agora = datetime.now()
    wd = agora.weekday()
    timestamp_str = agora.strftime("%Y-%m-%d %H:%M:%S")
    data_fmt = agora.strftime("%d/%m/%Y %H:%M")

    # Identifica o componente correto para este dia e turma no grade_horaria
    disciplina_real = "INFORMÁTICA"
    sala_real = f"Sala {turma}"
    alerta_confusao = ""

    if os.path.exists(ESCOLA_DB):
        try:
            with sqlite3.connect(ESCOLA_DB) as conn:
                cur = conn.cursor()
                cur.execute("SELECT disciplina, sala FROM grade_horaria WHERE dia_semana = ? AND turma = ? LIMIT 1", (wd, turma))
                row = cur.fetchone()
                if row:
                    disciplina_real, sala_real = row[0], row[1]
                else:
                    cur.execute("SELECT componente FROM alunos WHERE turma = ? LIMIT 1", (turma,))
                    row_al = cur.fetchone()
                    if row_al and row_al[0]:
                        disciplina_real = row_al[0]
        except Exception as e_disc:
            log_terminal(f"Aviso identificação disciplina: {e_disc}")

    # Checagem preventiva contra engano (ex: 202 em dia de Legislação)
    disc_norm = normalizar_texto(disciplina_real)
    conteudo_norm = normalizar_texto(conteudo)

    if turma == "202":
        if "legislacao" in disc_norm or "direito" in disc_norm:
            if any(k in conteudo_norm for k in ["deep learning", "rede neural", "redes neurais", "machine learning", "ia ", "inteligencia artificial", "visao computacional", "prompt"]):
                alerta_confusao = (
                    "\n\n🚨 *ALERTA ANTI-ENGANO DO ASSISTENTE:*\n"
                    "Você registrou um conteúdo típico de *Inteligência Artificial*, mas hoje na Turma 202 a matéria oficial da sua grade é *LEGISLAÇÃO DE INFORMÁTICA*!\n"
                    "_(Lembrete: IA na 202 é às quintas-feiras. O registro foi gravado como Legislação conforme a grade)_."
                )
        elif "inteligencia" in disc_norm or "ia" in disc_norm:
            if any(k in conteudo_norm for k in ["marco civil", "lgpd", "constituicao", "artigo", "direito autoral", "crime cibernetico", "legislacao"]):
                alerta_confusao = (
                    "\n\n🚨 *ALERTA ANTI-ENGANO DO ASSISTENTE:*\n"
                    "Você registrou um conteúdo de *Legislação*, mas hoje na Turma 202 a matéria oficial da sua grade é *INTELIGÊNCIA ARTIFICIAL*!\n"
                    "_(O registro foi gravado como IA conforme a grade oficial)_."
                )

    # 1. Grava no CSV
    log_existe = os.path.exists(DIARIO_BORDO_CSV)
    total_turma = 0
    with open(DIARIO_BORDO_CSV, "a+", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["timestamp", "data", "turma", "conteudo"])
        if not log_existe or os.path.getsize(DIARIO_BORDO_CSV) == 0:
            writer.writeheader()
        writer.writerow({"timestamp": timestamp_str, "data": data_fmt, "turma": turma, "conteudo": conteudo})

    try:
        with open(DIARIO_BORDO_CSV, "r", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                if normalizar_texto(r.get("turma", "")) == normalizar_texto(turma):
                    total_turma += 1
    except Exception:
        total_turma = 1

    # 2. Grava no SQLite com a disciplina correta
    if os.path.exists(ESCOLA_DB):
        try:
            with sqlite3.connect(ESCOLA_DB) as conn:
                conn.cursor().execute(
                    "INSERT INTO diario_classe (data_registro, turma, disciplina, conteudo_ministrado, observacoes) VALUES (?, ?, ?, ?, ?)",
                    (agora.strftime("%Y-%m-%d %H:%M"), turma, disciplina_real, conteudo, f"via WhatsApp ({sala_real})")
                )
                conn.commit()
        except Exception as e:
            log_terminal(f"⚠️ Erro ao espelhar no escola.db: {e}")

    resp = [
        f"📝 *DIÁRIO DE CLASSE REGISTRADO!*",
        f"👥 *Turma:* *Turma {turma}* ({sala_real})",
        f"📖 *Disciplina:* *{disciplina_real}*",
        f"📅 *Data:* {data_fmt}",
        f"📌 *Conteúdo:* \"{conteudo}\"",
        f"📊 *Total de registros da turma:* {total_turma} aulas"
    ]
    if alerta_confusao:
        resp.append(alerta_confusao)

    return "\n".join(resp)

def obter_dossie_aluno(nome_busca: str) -> str:
    if not nome_busca:
        return "⚠️ Uso: `!aluno <Nome>`"
    
    termo_norm = normalizar_texto(nome_busca)
    palavras = termo_norm.split()
    
    # 1. Busca aluno prioritariamente em dados_notas.csv ou planilha_notas_3bimestre.csv
    aluno_achado = None
    faltas = 0
    matricula = ""
    nome_completo = ""
    nota_principal = "--"
    turma = "Geral"
    b1, b2, b3, b4 = "--", "--", "--", "--"

    # Varre dados_notas.csv para pegar Faltas e Nota base
    if os.path.exists(NOTAS_CSV):
        try:
            with open(NOTAS_CSV, "r", encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    n_norm = normalizar_texto(row.get("Nome", ""))
                    if all(p in n_norm for p in palavras):
                        aluno_achado = row
                        nome_completo = row.get("Nome", "")
                        matricula = row.get("Matricula", "")
                        nota_principal = row.get("Nota", "--")
                        faltas = int(row.get("Faltas", 0) or 0)
                        break
        except Exception as e:
            log_terminal(f"⚠️ Erro ao ler dados_notas.csv: {e}")

    # Varre planilha oficial planilha_notas_3bimestre.csv para bimestres e dados completos
    if os.path.exists(NOTAS_3B_CSV):
        try:
            with open(NOTAS_3B_CSV, "r", encoding="utf-8") as f:
                for row in csv.DictReader(f, delimiter=";"):
                    n_norm = normalizar_texto(row.get("nome", ""))
                    mat_norm = row.get("matricula", "")
                    if (matricula and matricula == mat_norm) or (not aluno_achado and all(p in n_norm for p in palavras)):
                        turma = row.get("turma", turma)
                        if not nome_completo:
                            nome_completo = row.get("nome", "")
                        if not matricula:
                            matricula = mat_norm
                        b1 = row.get("AV1", "--") or "--"
                        b2 = row.get("AV2", "--") or "--"
                        b3 = row.get("AV3", "--") or "--"
                        aluno_achado = row
                        break
        except Exception as e:
            log_terminal(f"⚠️ Erro ao ler planilha_notas_3bimestre.csv: {e}")

    if not aluno_achado:
        return f"🔍 Aluno '{nome_busca}' não localizado em nenhuma das planilhas de notas."

    if b1 == "--" and nota_principal != "--":
        b1 = nota_principal

    notas_validas = []
    for val in [b1, b2, b3, b4]:
        try:
            notas_validas.append(float(str(val).replace(",", ".")))
        except ValueError:
            pass
            
    media_parcial = (sum(notas_validas) / len(notas_validas)) if notas_validas else 0.0
    pct_faltas = int((faltas / LIMITE_FALTAS) * 100)
    situacao = "Reprovado por Falta" if faltas >= LIMITE_FALTAS else ("Em curso (Aprovado por média até o momento)" if media_parcial >= 6.0 else "Em curso (Abaixo da média)")

    return (
        f"[Dossiê do Aluno]\n"
        f"Nome: {nome_completo} | Matrícula: {matricula} | Turma: {turma}\n"
        f"Notas:\n"
        f"  • B1: {b1} | B2: {b2} | B3: {b3} | B4: {b4}\n"
        f"  • Média Parcial: {media_parcial:.2f}\n"
        f"Faltas: {faltas} / {LIMITE_FALTAS} ({pct_faltas}% do limite atingido)\n"
        f"Situação: {situacao}"
    )

def buscar_multicriterio(query: str) -> str:
    filtro_sala = ""
    filtro_nome = ""
    m = re.search(r"sala:(\S.*)", query, re.IGNORECASE)
    if m:
        filtro_sala = normalizar_texto(m.group(1))
        filtro_nome = normalizar_texto(query[:m.start()].strip())
    else:
        filtro_nome = normalizar_texto(query.strip())

    aulas_da_sala = []
    if filtro_sala and os.path.exists(ESCOLA_DB):
        try:
            with sqlite3.connect(ESCOLA_DB) as conn:
                for row in conn.cursor().execute("SELECT dia_semana, hora_inicio, hora_fim, turma, sala, disciplina FROM grade_horaria"):
                    if filtro_sala in normalizar_texto(row[4]):
                        aulas_da_sala.append(f"• {row[3]} | {row[1]}-{row[2]} | {row[5]} ({row[4]})")
        except Exception as e:
            log_terminal(f"⚠️ Erro ao consultar escola.db: {e}")

    alunos_encontrados = []
    palavras_filtro = filtro_nome.split()
    matriculas_vistas = set()

    # Busca em dados_notas.csv
    if os.path.exists(NOTAS_CSV):
        try:
            with open(NOTAS_CSV, "r", encoding="utf-8") as f:
                for a in csv.DictReader(f):
                    nome_norm = normalizar_texto(a.get("Nome", ""))
                    if filtro_nome and not all(p in nome_norm for p in palavras_filtro):
                        continue
                    matriculas_vistas.add(a.get("Matricula", ""))
                    alunos_encontrados.append(a)
        except Exception as e:
            log_terminal(f"⚠️ Erro ao consultar dados_notas.csv: {e}")

    # Complementa com planilha oficial se houver novos
    if os.path.exists(NOTAS_3B_CSV):
        try:
            with open(NOTAS_3B_CSV, "r", encoding="utf-8") as f:
                for a in csv.DictReader(f, delimiter=";"):
                    nome_norm = normalizar_texto(a.get("nome", ""))
                    mat = a.get("matricula", "")
                    if mat in matriculas_vistas:
                        continue
                    if filtro_nome and not all(p in nome_norm for p in palavras_filtro):
                        continue
                    matriculas_vistas.add(mat)
                    alunos_encontrados.append({
                        "Matricula": mat,
                        "Nome": a.get("nome", ""),
                        "Nota": a.get("AV1", "--") or a.get("AV2", "--") or "--",
                        "Faltas": "0"
                    })
        except Exception as e:
            log_terminal(f"⚠️ Erro ao consultar planilha_notas_3bimestre.csv: {e}")

    res = []
    if filtro_sala:
        res.append(f"📍 *Ocupação da Sala/Espaço:*")
        res.append("\n".join(aulas_da_sala) if aulas_da_sala else "• Nenhuma turma alocada encontrada nesta sala.")
        res.append("")

    if alunos_encontrados:
        res.append(f"👤 *Alunos Encontrados:*")
        for a in alunos_encontrados[:10]:
            res.append(f"• *{a.get('Nome','')}* (Matrícula: {a.get('Matricula','')}) | Nota: {a.get('Nota','')} | Faltas: {a.get('Faltas','')}")
    elif filtro_nome:
        res.append(f"🔍 Nenhum aluno encontrado para '{filtro_nome}'.")

    return "\n".join(res) if res else "Nenhum resultado encontrado."

def consultar_aulas(termo=""):
    if not os.path.exists(AULAS_CSV):
        return "⚠️ Arquivo aulas.csv não encontrado."
    try:
        with open(AULAS_CSV, "r", encoding="utf-8") as f:
            aulas = list(csv.DictReader(f, delimiter=";"))
        t_norm = normalizar_texto(termo)
        res = []
        for a in aulas:
            linha_str = f"📅 {a.get('data','')} | Turma {a.get('turma','')}\n⏰ Horário: {a.get('horarios','')}\n📖 Disciplina: {a.get('componente','')}\n🎯 Tema: {a.get('tema','')}\n📝 Conteúdo: {a.get('conteudo','')}"
            if not t_norm or t_norm in normalizar_texto(json.dumps(a, ensure_ascii=False)):
                res.append(linha_str)
        return "📚 *GRADE DE AULAS:*\n\n" + "\n\n---\n\n".join(res[:5]) if res else f"🔍 Nenhuma aula encontrada para '{termo}'."
    except Exception as e:
        return f"❌ Erro ao ler aulas.csv: {e}"

DIAS_NOMES = [
    "Segunda-feira",
    "Terça-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "Sábado",
    "Domingo"
]

MAPA_DIAS = {
    "segunda": 0, "seg": 0, "segunda-feira": 0, "segunda feira": 0,
    "terca": 1, "terça": 1, "ter": 1, "terca-feira": 1, "terça-feira": 1, "terca feira": 1,
    "quarta": 2, "qua": 2, "quarta-feira": 2, "quarta feira": 2,
    "quinta": 3, "qui": 3, "quinta-feira": 3, "quinta feira": 3,
    "sexta": 4, "sex": 4, "sexta-feira": 4, "sexta feira": 4,
    "sabado": 5, "sábado": 5, "sab": 5,
    "domingo": 6, "dom": 6,
}

GRADE_TEMPOS_PADRAO = [
    ("1ª", "07:30", "08:25"),
    ("2ª", "08:25", "09:10"),
    ("3ª", "09:30", "10:20"),
    ("4ª", "10:20", "11:10"),
    ("5ª", "11:10", "12:00"),
    ("6ª", "13:30", "14:20"),
    ("7ª", "14:20", "15:05"),
    ("8ª", "15:20", "16:10"),
    ("9ª", "16:10", "17:00"),
]

def inferir_tempos(h_ini: str, h_fim: str) -> str:
    tempos_casados = []
    for rotulo, inicio, fim in GRADE_TEMPOS_PADRAO:
        if (h_ini <= inicio and h_fim >= fim) or (h_ini == inicio) or (h_fim == fim):
            tempos_casados.append(rotulo)
    if tempos_casados:
        return " e ".join(tempos_casados) + " Aula"
    return f"{h_ini} - {h_fim}"

def consultar_horario(termo: str = "") -> str:
    agora = datetime.now()
    t_norm = normalizar_texto(termo)
    
    if "amanha" in t_norm:
        alvo_dt = agora + timedelta(days=1)
    elif "ontem" in t_norm:
        alvo_dt = agora - timedelta(days=1)
    elif t_norm in MAPA_DIAS:
        wd_alvo = MAPA_DIAS[t_norm]
        diff = (wd_alvo - agora.weekday()) % 7
        alvo_dt = agora + timedelta(days=diff)
    else:
        m = re.search(r"(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?", t_norm)
        if m:
            dia, mes = int(m.group(1)), int(m.group(2))
            ano = int(m.group(3)) if m.group(3) else agora.year
            if ano < 100: ano += 2000
            try:
                alvo_dt = datetime(ano, mes, dia)
            except Exception:
                alvo_dt = agora
        else:
            achou = False
            for k, v in MAPA_DIAS.items():
                if k in t_norm:
                    diff = (v - agora.weekday()) % 7
                    alvo_dt = agora + timedelta(days=diff)
                    achou = True
                    break
            if not achou:
                alvo_dt = agora

    wd = alvo_dt.weekday()
    data_str = alvo_dt.strftime("%d/%m/%Y")
    dia_nome = DIAS_NOMES[wd]
    is_hoje = (alvo_dt.date() == agora.date())
    hora_atual = agora.strftime("%H:%M")

    linhas_resp = [
        f"📅 *{dia_nome.upper()} — {data_str}*",
        f"⏰ *Grade Escolar:* 07:30 às 17:00 (50 min/aula)",
    ]
    if is_hoje:
        linhas_resp.append(f"🕒 *Hora Atual:* {hora_atual}")
    linhas_resp.append("")

    # Consulta no SQLite escola.db
    aulas = []
    if os.path.exists(ESCOLA_DB):
        try:
            with sqlite3.connect(ESCOLA_DB) as conn:
                cur = conn.cursor()
                cur.execute("SELECT id, hora_inicio, hora_fim, turma, sala, disciplina, COALESCE(tempos, '') FROM grade_horaria WHERE dia_semana = ? ORDER BY hora_inicio ASC", (wd,))
                aulas = cur.fetchall()
        except Exception as e:
            log_terminal(f"⚠️ Erro ao consultar escola.db: {e}")

    # Conteúdo no CSV
    conteudos_dia = {}
    if os.path.exists(AULAS_CSV):
        try:
            with open(AULAS_CSV, "r", encoding="utf-8") as f:
                for row in csv.DictReader(f, delimiter=";"):
                    if row.get("data", "").strip() == data_str:
                        turma = row.get("turma", "").strip()
                        conteudos_dia[turma] = {
                            "tema": row.get("tema", "").strip(),
                            "conteudo": row.get("conteudo", "").strip()
                        }
        except Exception:
            pass

    if not aulas:
        linhas_resp.append("🏖️ *Nenhuma aula programada para este dia.*")
        linhas_resp.append("Aproveite o tempo livre ou prepare suas atividades pedagógicas.")
        linhas_resp.append("\n━━━━━━━━━━━━━━━━━━━━\n💡 Use `!grade` para ver todas as aulas da semana.")
        return "\n".join(linhas_resp)

    linhas_resp.append(f"📚 *AULAS PROGRAMADAS ({len(aulas)}):*\n")
    for id_aula, h_ini, h_fim, turma, sala, disc, tempos in aulas:
        rotulo_tempos = tempos if tempos else inferir_tempos(h_ini, h_fim)
        
        status = ""
        if is_hoje:
            if hora_atual < h_ini:
                status = f" | ⏳ Inicia às {h_ini}"
            elif h_ini <= hora_atual <= h_fim:
                status = " | 🚨 *EM ANDAMENTO*"
            else:
                status = " | ✅ Concluída"

        tema_info = conteudos_dia.get(turma, {})
        tema = tema_info.get("tema", "")
        tema_txt = f"\n   ▫️ Tema: _{tema}_" if tema else ""

        linhas_resp.append(f"• *{rotulo_tempos}* ({h_ini} - {h_fim}){status}")
        linhas_resp.append(f"   ▫️ Turma *{turma}* ({sala}) - {disc}{tema_txt}")

    linhas_resp.append("\n💡 Digite `!lab amanha` para ver horários do laboratório.")
    return "\n".join(linhas_resp)

def gerar_resumo_matinal(chat_id: str = DONO_CHAT_ID) -> str:
    """
    Consolida as aulas do dia, laboratório e pendências de visto em um único resumo executivo matinal.
    """
    hoje_dt = datetime.now()
    dias_semana = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"]
    dia_nome = dias_semana[hoje_dt.weekday()]
    hoje_str = hoje_dt.strftime("%d/%m/%Y")

    horario = consultar_horario("hoje")
    lab = processar_comando_lab("!lab hoje", {})
    alarmes = listar_alarmes()

    resp = [
        f"🌅 *RESUMO MATINAL DO DOCENTE — {dia_nome} ({hoje_str})*\n",
        f"📚 *SUAS AULAS DE HOJE:*",
        horario,
        "\n🔬 *AGENDA DO LABORATÓRIO (HOJE):*",
        lab,
        "\n⏰ *LEMBRETES & ALARMES:*",
        alarmes,
        "\n━━━━━━━━━━━━━━━━━━━━",
        "💡 *Comandos rápidos:* `!horario`, `!lab`, `!pendencias <turma>`, `!diario`"
    ]
    return "\n".join(resp)

def listar_grade_semanal() -> str:
    if not os.path.exists(ESCOLA_DB):
        return "Erro: escola.db não encontrado."
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, dia_semana, hora_inicio, hora_fim, turma, sala, disciplina, COALESCE(tempos, '') FROM grade_horaria ORDER BY dia_semana ASC, hora_inicio ASC")
            linhas = cur.fetchall()
    except Exception as e:
        return f"Erro ao acessar banco: {e}"

    if not linhas:
        return "Nenhuma aula cadastrada na grade horária."

    resp = ["📋 *GRADE COMPLETA DA SEMANA:*", "Horário escolar: 07:30 às 17:00 (50 min/aula)\n"]
    ultimo_dia = None
    for id_aula, wd, h_ini, h_fim, turma, sala, disc, tempos in linhas:
        if wd != ultimo_dia:
            resp.append(f"\n📌 *{DIAS_NOMES[wd].upper()}:*")
            ultimo_dia = wd
        rotulo_tempo = f"({tempos})" if tempos else ""
        resp.append(f"• `ID {id_aula}`: {h_ini}-{h_fim} {rotulo_tempo} | Turma *{turma}* | {disc} ({sala})")
    
    resp.append("\n━━━━━━━━━━━━━━━━━━━━")
    resp.append("Para modificar via WhatsApp:\n• `!adicionaraula <dia> <ini> <fim> <turma> <disc> [sala]`\n• `!removeraula <id>`")
    return "\n".join(resp)

def adicionar_aula_grade(dia_arg: str, h_ini: str, h_fim: str, turma: str, disc: str, sala: str = "Sala 202") -> str:
    t_norm = normalizar_texto(dia_arg)
    if t_norm not in MAPA_DIAS:
        try:
            wd = int(dia_arg)
            if not 0 <= wd <= 6: raise ValueError()
        except ValueError:
            return f"❌ Dia inválido: '{dia_arg}'. Use: segunda, terca, quarta, quinta, sexta, sabado ou domingo."
    else:
        wd = MAPA_DIAS[t_norm]

    tempos = inferir_tempos(h_ini, h_fim)
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("INSERT INTO grade_horaria (dia_semana, hora_inicio, hora_fim, turma, sala, disciplina, tempos) VALUES (?, ?, ?, ?, ?, ?, ?)",
                        (wd, h_ini, h_fim, turma, sala, disc, tempos))
            conn.commit()
            novo_id = cur.lastrowid
        return f"✅ Aula inserida com sucesso (ID `{novo_id}`):\n• {DIAS_NOMES[wd]} | {h_ini}-{h_fim} ({tempos})\n• Turma: {turma} | Sala: {sala} | {disc}"
    except Exception as e:
        return f"❌ Erro ao adicionar aula: {e}"

def remover_aula_grade(id_arg: str) -> str:
    try:
        id_int = int(id_arg)
    except ValueError:
        return f"❌ ID inválido: '{id_arg}'. Forneça o número do ID (veja em `!grade`)."
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("SELECT dia_semana, hora_inicio, hora_fim, turma, disciplina FROM grade_horaria WHERE id = ?", (id_int,))
            item = cur.fetchone()
            if not item:
                return f"❌ Aula com ID `{id_int}` não encontrada."
            cur.execute("DELETE FROM grade_horaria WHERE id = ?", (id_int,))
            conn.commit()
        return f"✅ Aula ID `{id_int}` removida com sucesso! ({DIAS_NOMES[item[0]]} {item[1]}-{item[2]} | Turma {item[3]} - {item[4]})"
    except Exception as e:
        return f"❌ Erro ao remover aula: {e}"

# --- CONFIGURAÇÕES DE SISTEMA, MODO AWAY & CLOUD CLI ---

def init_tabelas_config():
    if not os.path.exists(ESCOLA_DB):
        return
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS config_sistema (
                    chave TEXT PRIMARY KEY,
                    valor TEXT,
                    atualizado_em TEXT
                );
                INSERT OR IGNORE INTO config_sistema (chave, valor, atualizado_em)
                VALUES ('modo_away', '0', datetime('now'));
            """)
    except Exception as e:
        log_terminal(f"⚠️ Erro ao inicializar config_sistema: {e}")

def init_tabela_telemetria():
    if not os.path.exists(ESCOLA_DB):
        return
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=5) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS metricas_eventos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT,
                    prof_chave TEXT,
                    comando TEXT,
                    duracao_ms INTEGER,
                    status TEXT,
                    erro TEXT
                )
            """)
            conn.commit()
    except Exception as e:
        log_terminal(f"⚠️ Erro ao inicializar metricas_eventos: {e}")

def registrar_telemetria(prof_chave: str, comando: str, duracao_ms: int, status: str = "ok", erro: str = ""):
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=3) as conn:
            conn.execute("""
                INSERT INTO metricas_eventos (timestamp, prof_chave, comando, duracao_ms, status, erro)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), prof_chave or "sistema", (comando or "")[:60], duracao_ms, status, (erro or "")[:300]))
            conn.commit()
    except Exception:
        pass

def obter_resumo_metricas_str() -> str:
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=5) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT COUNT(*), 
                       SUM(CASE WHEN status = 'ok' THEN 1 ELSE 0 END),
                       SUM(CASE WHEN status != 'ok' THEN 1 ELSE 0 END),
                       AVG(duracao_ms)
                FROM metricas_eventos
                WHERE datetime(timestamp) >= datetime('now', '-24 hours')
            """)
            row = cur.fetchone()
            tot, ok, err, avg_ms = row[0] or 0, row[1] or 0, row[2] or 0, int(row[3] or 0)

            cur.execute("""
                SELECT comando, COUNT(*) as qtd
                FROM metricas_eventos
                WHERE datetime(timestamp) >= datetime('now', '-24 hours')
                GROUP BY comando ORDER BY qtd DESC LIMIT 3
            """)
            top_cmds = cur.fetchall()

        uptime_s = int(time.time() - TIMESTAMP_INICIO_BOOT)
        disco_gb = round(shutil.disk_usage(BASE_DIR).free / (1024**3), 2)
        tam_log_kb = round(os.path.getsize(LOG_FILE) / 1024, 1) if os.path.exists(LOG_FILE) else 0

        taxa_ok = round((ok / tot * 100), 1) if tot > 0 else 100.0
        top_str = "\n".join([f"  • `{c}`: {q}x" for c, q in top_cmds]) if top_cmds else "  • Nenhuma requisição recente"

        return (
            "📊 *TELEMETRIA & OBSERVABILIDADE (ÚLTIMAS 24H)*\n\n"
            f"⚡ *Comandos processados:* {tot}\n"
            f"✅ *Taxa de sucesso:* {taxa_ok}% ({ok} ok, {err} erros)\n"
            f"⏱️ *Tempo médio de resposta:* {avg_ms} ms\n"
            f"⏳ *Uptime do Webhook:* {str(timedelta(seconds=uptime_s))}\n"
            f"💾 *Disco livre:* {disco_gb} GB | *Log:* {tam_log_kb} KB\n\n"
            f"🔥 *Top Comandos:*\n{top_str}"
        )
    except Exception as e:
        return f"Erro ao coletar métricas: {e}"
        log_terminal(f"⚠️ Erro ao inicializar config_sistema: {e}")

def obter_config(chave: str, padrao: str = "") -> str:
    if not os.path.exists(ESCOLA_DB):
        return padrao
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("SELECT valor FROM config_sistema WHERE chave = ?", (chave,))
            row = cur.fetchone()
            return row[0] if row else padrao
    except Exception:
        return padrao

def definir_config(chave: str, valor: str):
    if not os.path.exists(ESCOLA_DB):
        return
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            conn.cursor().execute("""
                INSERT INTO config_sistema (chave, valor, atualizado_em)
                VALUES (?, ?, datetime('now'))
                ON CONFLICT(chave) DO UPDATE SET valor=excluded.valor, atualizado_em=excluded.atualizado_em
            """, (chave, valor))
            conn.commit()
    except Exception as e:
        log_terminal(f"⚠️ Erro ao salvar config {chave}: {e}")

def is_modo_away() -> bool:
    return obter_config("modo_away", "0") == "1"

def alternar_modo_away(arg: str = "") -> str:
    init_tabelas_config()
    arg_norm = arg.strip().lower()
    
    if arg_norm in ["on", "ativar", "ligar", "1", "sim"]:
        novo_estado = True
    elif arg_norm in ["off", "desativar", "desligar", "0", "nao", "não", "voltar"]:
        novo_estado = False
    elif arg_norm in ["status", "info"]:
        st_txt = "ATIVO 🚶" if is_modo_away() else "INATIVO 🏠"
        return f"🚶 *Status do Modo Away:* {st_txt}\n\nUse `!away on` ou `!away off` para alternar."
    else:
        # Alterna (toggle)
        novo_estado = not is_modo_away()

    definir_config("modo_away", "1" if novo_estado else "0")
    
    if novo_estado:
        return (
            "🚶 *Modo AWAY (Ausente) ATIVADO!*\n\n"
            "• O assistente responderá e enviará notificações, saídas do Cloud CLI "
            "e conclusões de tarefas diretamente no seu WhatsApp privado.\n"
            "• Ao retornar ao computador, envie `!away off`."
        )
    else:
        return (
            "🏠 *Modo AWAY (Ausente) DESATIVADO!*\n\n"
            "• Bem-vindo de volta ao terminal. Notificações regulares restauradas."
        )

def executar_cloud_cli(instrucao: str, chat_id: str):
    """
    Agente Cloud CLI via WhatsApp:
    Recebe comandos diretos ou solicitações em linguagem natural,
    roteia usando OmniRoute se necessário, executa no sistema operacional Linux,
    e devolve o retorno formatado em terminal para o WhatsApp privado.
    """
    cmd_raw = instrucao.strip()
    if not cmd_raw:
        enviar_zap(chat_id, (
            "☁️ *CLOUD CLI DOCENTE (Terminal WhatsApp)*\n\n"
            "Use: `!cloud <comando bash ou solicitação>`\n\n"
            "Exemplos:\n"
            "• `!cloud uptime` (carga e tempo ativo)\n"
            "• `!cloud df -h` (espaço em disco)\n"
            "• `!cloud free -h` (uso de memória RAM)\n"
            "• `!cloud tailscale status`\n"
            "• `!cloud docente status`\n"
            "• `!cloud python3 projetos_ideias/exemplo/prototipo.py &`\n\n"
            "🧠 *Instruções em Linguagem Natural:*\n"
            "• `!cloud verifique o uso de CPU e memória`\n"
            "• `!cloud mostre se a porta 8080 está ativa`\n"
            "• `!cloud liste os arquivos da pasta de ideias`"
        ))
        return

    log_terminal(f"☁️ Cloud CLI acionado: '{cmd_raw}' por {chat_id}")
    
    # Comandos bash diretos comuns
    tokens = cmd_raw.split()
    primeiro_token = tokens[0].lower() if tokens else ""
    COMANDOS_BASH_DIRETOS = {
        "df", "free", "uptime", "ps", "top", "ls", "pwd", "whoami", "uname",
        "date", "cat", "tail", "head", "grep", "find", "pkill", "kill",
        "tailscale", "docente", "git", "python3", "python", "curl", "ss", "netstat",
        "ip", "systemctl", "docker", "rclone", "mkdir", "cp", "mv", "rm", "sqlite3",
        "which", "env", "echo"
    }
    
    if primeiro_token in COMANDOS_BASH_DIRETOS or cmd_raw.startswith(("./", "/", "python3 ", "bash ")):
        comando_para_executar = cmd_raw
        explicacao = ""
    else:
        # Consulta OmniRoute para rotear a intenção em comando Bash
        prompt = (
            "Você é o Agente Cloud CLI Docente em um servidor Linux Ubuntu com Python 3 e Tailscale.\n"
            "O usuário enviou uma instrução via WhatsApp para ser executada no terminal do servidor.\n\n"
            f"Instrução: \"{cmd_raw}\"\n\n"
            "Determine o comando Bash mais seguro, direto e eficiente para realizar o que foi pedido.\n"
            "Responda SOMENTE um JSON válido com a estrutura:\n"
            "{\n"
            '  "comando_bash": "comando bash completo aqui",\n'
            '  "explicacao": "Breve frase explicando o que será feito"\n'
            "}"
        )
        try:
            raw_llm = consultar_llm(prompt)
            m = re.search(r'\{.*\}', raw_llm, re.DOTALL)
            if m:
                d = json.loads(m.group(0))
                comando_para_executar = d.get("comando_bash", cmd_raw)
                explicacao = d.get("explicacao", "")
            else:
                comando_para_executar = cmd_raw
                explicacao = ""
        except Exception as e:
            log_terminal(f"⚠️ Erro no roteamento LLM do Cloud CLI: {e}")
            comando_para_executar = cmd_raw
            explicacao = ""

    # Execução do comando no terminal Linux
    inicio_t = datetime.now()
    try:
        proc = subprocess.run(
            comando_para_executar,
            shell=True,
            capture_output=True,
            text=True,
            timeout=35,
            cwd=os.environ.get("DOCENTE_HOME", ".")
        )
        duracao = round((datetime.now() - inicio_t).total_seconds(), 2)
        saida = proc.stdout if proc.stdout else proc.stderr
        codigo = proc.returncode
    except subprocess.TimeoutExpired:
        saida = "⚠️ Erro: Comando excedeu o tempo limite de 35 segundos."
        codigo = 124
        duracao = 35.0
    except Exception as e:
        saida = f"❌ Erro de execução: {e}"
        codigo = 1
        duracao = round((datetime.now() - inicio_t).total_seconds(), 2)

    # Truncagem amigável para o WhatsApp se a saída for gigantesca
    if len(saida) > 2500:
        saida = saida[:2500] + "\n\n... [Saída truncada para leitura no WhatsApp]"
    elif not saida.strip():
        saida = "(Comando executado com sucesso sem saída de texto)"

    status_icone = "✅" if codigo == 0 else "⚠️"
    
    linhas_resp = [
        "☁️ *CLOUD CLI (Terminal WhatsApp)*",
        f"💻 `$ {comando_para_executar}`"
    ]
    if explicacao:
        linhas_resp.append(f"ℹ️ _{explicacao}_")
    
    linhas_resp.append("\n```\n" + saida.strip() + "\n```")
    linhas_resp.append(f"{status_icone} Retorno: `{codigo}` | ⏱️ {duracao}s")

    if is_modo_away():
        linhas_resp.append("🚶 _[Modo Away Ativo]_")

    enviar_zap(chat_id, "\n".join(linhas_resp))
    if is_modo_away() and chat_id != DONO_CHAT_ID:
        enviar_zap(DONO_CHAT_ID, "\n".join(linhas_resp))

def consultar_status_sistema() -> str:
    agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    ts_ip = os.environ.get("TAILSCALE_IP", "SEU_IP")
    try:
        res = subprocess.run(["tailscale", "ip", "-4"], capture_output=True, text=True)
        if res.returncode == 0 and res.stdout.strip():
            ts_ip = res.stdout.strip()
    except Exception:
        pass
    
    away_txt = "🚶 ATIVO (Respostas e avisos no WhatsApp)" if is_modo_away() else "🏠 INATIVO"
    
    return (
        f"🛰️ *STATUS DO SISTEMA DOCENTE*\n\n"
        f"🕒 Data/Hora: {agora_str}\n"
        f"🌐 IP Tailscale: `{ts_ip}`\n"
        f"🔑 Acesso SSH: `ssh USUARIO@{ts_ip}`\n"
        f"🚶 Modo Away: *{away_txt}*\n\n"
        f"✅ Webhook WhatsApp: Ativo (Porta 5000)\n"
        f"✅ OpenWA Engine: Conectado (Porta 2785)\n"
        f"✅ OmniRoute IA: Operacional (Porta 20128)\n"
        f"✅ Cron Daemon: Rodando\n\n"
        f"💡 Você pode gerenciar tudo via terminal com o comando `docente`."
    )

NUMERO_PROPRIETARIO = os.environ.get("NUMERO_PROPRIETARIO", "")
LID_PROPRIETARIO = os.environ.get("LID_PROPRIETARIO", "")
# Identificadores (número ou LID) autorizados, separados por vírgula em WHATSAPP_AUTORIZADOS
IDENTIFICADORES_AUTORIZADOS = {
    x.strip() for x in ([NUMERO_PROPRIETARIO, LID_PROPRIETARIO] + os.environ.get("WHATSAPP_AUTORIZADOS", "").split(","))
    if x and x.strip()
}

ID_GRUPO_BASE_TECNICA = os.environ.get("ID_GRUPO_BASE_TECNICA", "")
ID_GRUPO_ALUNOS_SEQUESTRO = os.environ.get("ID_GRUPO_ALUNOS_SEQUESTRO", "")
ID_GRUPO_CLUBE_PROGRAMACAO = os.environ.get("ID_GRUPO_CLUBE_PROGRAMACAO", "")
ID_GRUPO_INFRA_REDTEAM = os.environ.get("ID_GRUPO_INFRA_REDTEAM", "")

GRUPOS_ALUNOS_LAB = {k: v for k, v in {
    ID_GRUPO_ALUNOS_SEQUESTRO: "devs IEMA",
    ID_GRUPO_CLUBE_PROGRAMACAO: "Clube de Programação/Hacking/I.A",
    ID_GRUPO_INFRA_REDTEAM: "Infra e Red Ream IEMA🟥",
}.items() if k}

# Grupos de alunos onde comandos de OSINT, Geointeligência e Reconhecimento são liberados
GRUPOS_OSINT_PERMITIDOS = {g for g in {
    ID_GRUPO_CLUBE_PROGRAMACAO,
    ID_GRUPO_INFRA_REDTEAM,
} if g}

GRUPOS_TURMAS = {
    "200": os.environ.get("GRUPO_TURMA_200", ""),
    "201": os.environ.get("GRUPO_TURMA_201", ""),
    "202": os.environ.get("GRUPO_TURMA_202", ""),
    "203": os.environ.get("GRUPO_TURMA_203", ""),
}

NOMES_GRUPOS_TURMAS = {
    "200": "200 - TÉCNICO EM ADMINISTRAÇÃO 📚",
    "201": "201 - TÉCNICO EM MARKETING 💙💡",
    "202": "202 - TÉCNICO EM DESENVOLVIMENTO DE SISTEMAS",
    "203": "203 - TÉCNICO EM INFORMÁTICA PARA INTERNET",
    "base": "IEMA BASE TÉCNICA TI",
}

TURMA_POR_GRUPO = {jid: t for t, jid in GRUPOS_TURMAS.items() if jid}
CHAT_PRIVADO_PADRAO = os.environ.get("CHAT_PRIVADO_PADRAO", "")
CHAT_PRIVADO_LEANDRO = os.environ.get("WHATSAPP_PRIVADO_LEANDRO", "")
CHAT_PRIVADO_BRUNA = os.environ.get("WHATSAPP_PRIVADO_BRUNA", "")

def obter_estado_publico_sala(turma_ou_base: str) -> bool:
    """Retorna True se o estado da sala for PÚBLICO (envia resumo no grupo), False se PRIVADO."""
    chave = f"modo_publico_sala_{str(turma_ou_base).strip().lower()}"
    val = obter_config(chave, "1")
    return val != "0"

def definir_estado_publico_sala(turma_ou_base: str, publico: bool):
    chave = f"modo_publico_sala_{str(turma_ou_base).strip().lower()}"
    definir_config(chave, "1" if publico else "0")

def _extrair_flag_visibilidade_cmd(texto: str):
    """
    Detecta se o comando traz 'publico'/'no grupo' ou 'privado'/'silencioso' ao final/início.
    Retorna (override_bool_ou_None, texto_sem_flag).
    """
    if not texto:
        return None, ""
    tokens = texto.split()
    override = None
    rest = []
    for tok in tokens:
        t_low = normalizar_texto(tok.strip(",;."))
        if t_low in ("publico", "publica", "grupo", "nogrupo"):
            override = True
        elif t_low in ("privado", "privada", "silencioso", "off", "soprivado"):
            override = False
        else:
            rest.append(tok)
    return override, " ".join(rest).strip()

def remover_dicas_comandos_para_grupo(texto: str) -> str:
    """Remove qualquer dica de comando, sintaxe de bot ou rodapé administrativo antes de enviar em grupos."""
    if not texto:
        return ""
    linhas_limpas = []
    pular_bloco_dica = False
    for linha in texto.splitlines():
        l_strip = linha.strip()
        l_low = normalizar_texto(l_strip)
        if any(l_strip.startswith(pref) for pref in ("💡", "👉", "Para modificar via WhatsApp", "Para dar visto", "Para cobrar", "Para registrar", "Para reservar", "Use:", "Ex:", "Exemplo:")):
            pular_bloco_dica = True
            continue
        if any(k in l_low for k in ("atalhos rapidos", "comandos rapidos", "lancar frequencia na aula", "para lancar todos", "para publicar a cobranca", "para cobrar esta atividade", "para cobrar a turma", "para reservar horarios", "para alterar:")):
            pular_bloco_dica = True
            continue
        if pular_bloco_dica:
            if l_strip.startswith(("•", "-", "`", "Ex:", "Exemplo:")):
                continue
            elif not l_strip:
                continue
            else:
                pular_bloco_dica = False
        if re.search(r"`(?:![a-z]+|atv\s|cobrar\s|visto\s|passar\s|resumo\s)", l_strip, flags=re.IGNORECASE):
            continue
        linhas_limpas.append(linha)
    res = "\n".join(linhas_limpas).strip()
    res = re.sub(r"(?:\n━━━━━━━━━━━━━━━━━━━━\s*)+$", "", res).strip()
    return res

def gerar_resumo_publico_frequencia_sala(turma: str, resumo_acao: str = "") -> str:
    """Gera o resumo limpo de frequência da turma para envio no grupo correspondente (sem dicas de comando)."""
    turma = str(turma).strip()
    nome_grupo = NOMES_GRUPOS_TURMAS.get(turma, f"TURMA {turma}")
    hoje_str = datetime.now().strftime("%d/%m/%Y")
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("SELECT nome, faltas, componente FROM alunos WHERE turma = ? ORDER BY nome ASC", (turma,))
            alunos = cur.fetchall()
        if not alunos:
            return ""
        total_alunos = len(alunos)
        total_faltas_acum = sum(a[1] for a in alunos)
        sem_faltas = sum(1 for a in alunos if a[1] == 0)
        com_alerta = [(idx, a[0], a[1]) for idx, a in enumerate(alunos, 1) if a[1] >= 10]

        linhas = [
            f"📋 *BOLETIM DE FREQUÊNCIA — TURMA {turma}*",
            f"🏫 *{nome_grupo}*",
            f"📅 *Data:* {hoje_str}",
            "━━━━━━━━━━━━━━━━━━━━"
        ]
        if resumo_acao:
            linhas.append("📌 *Atualização de Chamada:*")
            linhas.append(resumo_acao.strip())
            linhas.append("━━━━━━━━━━━━━━━━━━━━")

        linhas.append("📊 *Resumo Geral de Frequência:*")
        linhas.append(f"• Total de alunos: *{total_alunos}*")
        linhas.append(f"• Alunos com 100% de presença (0 faltas): *{sem_faltas}*")
        linhas.append(f"• Total de faltas acumuladas na turma: *{total_faltas_acum}*")

        if com_alerta:
            linhas.append("\n⚠️ *Atenção ao limite de faltas:*")
            for idx, n_al, flt in com_alerta[:8]:
                linhas.append(f"  • #{idx:02d} {n_al}: *{flt} faltas*")

        return "\n".join(linhas)
    except Exception as e:
        log_terminal(f"Erro em gerar_resumo_publico_frequencia_sala: {e}")
        return ""

def gerar_resumo_publico_atividade_sala(turma: str, ativ_id: int = None, detalhe_recente: str = "") -> str:
    """Gera o resumo geral da atividade da turma para o grupo da sala (sem dicas de comando)."""
    turma = str(turma).strip()
    nome_grupo = NOMES_GRUPOS_TURMAS.get(turma, f"TURMA {turma}")
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("SELECT matricula, nome FROM alunos WHERE turma = ? ORDER BY nome ASC", (turma,))
            alunos_turma = cur.fetchall()
            if not alunos_turma:
                return ""

            ativ_sel = _resolver_atividade_alvo(
                cur, turma, "", ativ_id, "", "", preferir_com_entregas=(ativ_id is None)
            )
            if not ativ_sel:
                return ""

            aid, dt_passada, _, comp, titulo = ativ_sel
            vistos_reg = {}
            cur.execute("""
                SELECT matricula, nome_aluno, status, data_registro
                FROM atividades_entregas
                WHERE turma = ? AND (atividade_id = ? OR atividade = ? OR atividade LIKE ?)
                ORDER BY id ASC
            """, (turma, aid, titulo, f"%{titulo}%"))
            for mat, nom, st, dt in cur.fetchall():
                if (st or "").upper() in ["VISTO (OK)", "OK", "ENTREGUE", "FEITO", "10"]:
                    vistos_reg[mat] = (nom, dt, st)
                elif (st or "").upper() in ["PENDENTE", "INCOMPLETO", "RECUSADO"]:
                    vistos_reg.pop(mat, None)

        total_alunos = len(alunos_turma)
        entregues = []
        pendentes = []
        for idx, (mat, nome) in enumerate(alunos_turma, 1):
            teve = mat in vistos_reg
            if not teve:
                n_fon = normalizar_fonetico(nome)
                for v_m, (v_n, v_d, v_s) in list(vistos_reg.items()):
                    if n_fon == normalizar_fonetico(v_n):
                        teve = True
                        break
            if teve:
                entregues.append((idx, nome))
            else:
                pendentes.append((idx, nome))

        pct_ent = round((len(entregues) / total_alunos) * 100) if total_alunos else 0
        pct_pend = round((len(pendentes) / total_alunos) * 100) if total_alunos else 0

        resp = [
            f"📋 *RESUMO DE ATIVIDADE — TURMA {turma}*",
            f"🏫 *{nome_grupo}*",
            f"📖 *Componente:* {comp}",
            f"📝 *Atividade:* *{titulo}*",
            f"📅 *Passada em:* {dt_passada}",
            "━━━━━━━━━━━━━━━━━━━━"
        ]
        if detalhe_recente:
            resp.append(detalhe_recente.strip())
            resp.append("━━━━━━━━━━━━━━━━━━━━")

        resp.extend([
            "📊 *Panorama de Entregas da Turma:*",
            f"• Total de alunos: {total_alunos}",
            f"• ✅ Entregaram (Visto OK): *{len(entregues)}* ({pct_ent}%)",
            f"• ⏳ Pendentes: *{len(pendentes)}* ({pct_pend}%)",
            "━━━━━━━━━━━━━━━━━━━━"
        ])

        if pendentes:
            resp.append(f"🚨 *ALUNOS PENDENTES ({len(pendentes)}):*")
            for n_ch, n_al in pendentes:
                resp.append(f"  *{n_ch:02d}.* {n_al}")
        else:
            resp.append("🎉 *PARABÉNS! TODOS OS ALUNOS ENTREGARAM A ATIVIDADE!*")

        if entregues:
            nomes_ok = ", ".join([f"#{n_ch:02d} {n_al.split()[0]}" for n_ch, n_al in entregues])
            resp.append(f"\n✅ *JÁ ENTREGARAM ({len(entregues)}):*\n  {nomes_ok}")

        return "\n".join(resp)
    except Exception as e:
        log_terminal(f"Erro em gerar_resumo_publico_atividade_sala: {e}")
        return ""

def gerar_resumo_geral_sala(alvo: str) -> str:
    """
    Gera o Resumo Geral Consolidado de uma sala ('200', '201', '202', '203' ou 'base'),
    combinando Frequência e Atividades sem nenhuma dica de comando (pronto para envio público ou privado).
    """
    alvo_clean = normalizar_texto(alvo).strip()
    if alvo_clean in ("base", "ti", "base tecnica", "laboratorio", "lab"):
        hoje_str = datetime.now().strftime("%d/%m/%Y")
        lab_txt = remover_dicas_comandos_para_grupo(processar_comando_lab("!lab hoje", {}))
        return (
            f"📊 *RESUMO GERAL — IEMA BASE TÉCNICA TI*\n"
            f"📅 *Data:* {hoje_str}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"{lab_txt}"
        )

    turma = alvo_clean
    nome_grupo = NOMES_GRUPOS_TURMAS.get(turma, f"TURMA {turma}")
    hoje_str = datetime.now().strftime("%d/%m/%Y")
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("SELECT matricula, nome, componente, faltas FROM alunos WHERE turma = ? ORDER BY nome ASC", (turma,))
            alunos = cur.fetchall()
            if not alunos:
                return f"⚠️ Nenhuma informação encontrada para a sala '{alvo}'."

            total_alunos = len(alunos)
            comps = alunos[0][2] if alunos else ""
            total_faltas = sum(a[3] for a in alunos)
            sem_faltas = sum(1 for a in alunos if a[3] == 0)
            com_faltas = [(idx, a[1], a[3]) for idx, a in enumerate(alunos, 1) if a[3] > 0]

            ativ_sel = _resolver_atividade_alvo(cur, turma, "", None, "", "", preferir_com_entregas=True)
            bloco_ativ = []
            if ativ_sel:
                aid, dt_p, _, comp_a, tit_a = ativ_sel
                vistos_reg = {}
                cur.execute("""
                    SELECT matricula, nome_aluno, status
                    FROM atividades_entregas
                    WHERE turma = ? AND (atividade_id = ? OR atividade = ? OR atividade LIKE ?)
                    ORDER BY id ASC
                """, (turma, aid, tit_a, f"%{tit_a}%"))
                for mat, nom, st in cur.fetchall():
                    if (st or "").upper() in ["VISTO (OK)", "OK", "ENTREGUE", "FEITO", "10"]:
                        vistos_reg[mat] = nom
                    elif (st or "").upper() in ["PENDENTE", "INCOMPLETO", "RECUSADO"]:
                        vistos_reg.pop(mat, None)

                entregues = []
                pendentes = []
                for idx, (mat, nome, _, _) in enumerate(alunos, 1):
                    if mat in vistos_reg or any(normalizar_fonetico(nome) == normalizar_fonetico(vn) for vn in vistos_reg.values()):
                        entregues.append((idx, nome))
                    else:
                        pendentes.append((idx, nome))

                pct_e = round((len(entregues) / total_alunos) * 100) if total_alunos else 0
                bloco_ativ.extend([
                    f"📝 *ATIVIDADE EM ACOMPANHAMENTO:*",
                    f"• *Título:* {tit_a}",
                    f"• *Disciplina:* {comp_a} (Passada em: {dt_p})",
                    f"• ✅ *Entregaram:* {len(entregues)}/{total_alunos} ({pct_e}%)",
                    f"• ⏳ *Pendentes:* {len(pendentes)}/{total_alunos}"
                ])
                if pendentes:
                    bloco_ativ.append("\n🚨 *Alunos Pendentes na Atividade:*")
                    for n_ch, n_al in pendentes:
                        bloco_ativ.append(f"  *{n_ch:02d}.* {n_al}")
                else:
                    bloco_ativ.append("\n🎉 *Todos os alunos da turma entregaram a atividade!*")
            else:
                bloco_ativ.append("📝 *ATIVIDADES:* Nenhuma atividade pendente cadastrada no momento.")

        linhas_freq = [
            f"📊 *RESUMO GERAL DA SALA — TURMA {turma}*",
            f"🏫 *{nome_grupo}*",
            f"📖 *Componentes:* {comps}",
            f"📅 *Data:* {hoje_str}",
            "━━━━━━━━━━━━━━━━━━━━",
            "📋 *FREQUÊNCIA E ASSIDUIDADE:*",
            f"• Total de alunos: *{total_alunos}* | Sem faltas: *{sem_faltas}* | Faltas acumuladas: *{total_faltas}*"
        ]
        if com_faltas:
            res_f = ", ".join([f"#{idx:02d} {n.split()[0]} ({f}f)" for idx, n, f in com_faltas[:12]])
            if len(com_faltas) > 12:
                res_f += f" e mais {len(com_faltas) - 12}"
            linhas_freq.append(f"• Registros de faltas: {res_f}")

        linhas_freq.append("━━━━━━━━━━━━━━━━━━━━")
        linhas_freq.extend(bloco_ativ)
        return "\n".join(linhas_freq)
    except Exception as e:
        return f"❌ Erro ao gerar resumo da sala {turma}: {e}"

def despachar_resumo_grupo_se_publico(turma: str, msg_grupo: str, override_publico=None) -> str:
    """
    Se a sala estiver em modo PÚBLICO (ou override_publico=True), envia msg_grupo limpa no grupo da turma
    e retorna o aviso de rodapé para o chat privado do professor.
    """
    turma = str(turma).strip()
    grupo_jid = GRUPOS_TURMAS.get(turma)
    nome_grupo = NOMES_GRUPOS_TURMAS.get(turma, f"Turma {turma}")
    if not grupo_jid or not msg_grupo:
        return ""

    if override_publico is not None:
        definir_estado_publico_sala(turma, override_publico)
        eh_publico = override_publico
    else:
        eh_publico = obter_estado_publico_sala(turma)

    if eh_publico:
        msg_limpa = remover_dicas_comandos_para_grupo(msg_grupo)
        ok = enviar_zap(grupo_jid, msg_limpa)
        if ok:
            return f"\n\n📢 *Resumo publicado no grupo:* _{nome_grupo}_\n_(Para deixar privado: `resumo {turma} privado`)_"
        return f"\n\n⚠️ _Não foi possível publicar no grupo {nome_grupo} no momento._"
    else:
        return f"\n\n🔒 *Modo PRIVADO ativo na Turma {turma}:* respondido apenas aqui.\n_(Para publicar no grupo: `resumo {turma} publico`)_"

def processar_comando_resumo_ou_modo(texto_cmd: str, chat_id: str) -> str:
    """
    Trata:
    • '!resumo' / 'resumo' / 'matinal' (sem turma) -> Resumo matinal do docente (estritamente privado)
    • 'resumo 202 [publico|privado]' / 'resumo base [publico|privado]' -> Resumo Geral da Sala (público no grupo ou privado)
    • '!modo [200|201|202|203|base] [publico|privado]' -> Consulta ou altera estado público/privado da sala
    """
    t_norm = normalizar_texto(texto_cmd)
    eh_cmd_modo = t_norm.startswith(("!modo", "/modo", "modo"))
    corpo = re.sub(r"^[!/]?(?:resumo|matinal|dia|modo)\s*", "", texto_cmd.strip(), flags=re.IGNORECASE).strip()
    corpo = re.sub(r"^(?:da\s+sala|sala|da\s+turma|turma)\s+", "", corpo, flags=re.IGNORECASE).strip()

    override_pub, corpo_sem_flag = _extrair_flag_visibilidade_cmd(corpo)
    toks = corpo_sem_flag.split()

    alvo_sala = ""
    for tk in toks:
        tk_c = normalizar_texto(tk.strip(",;."))
        if tk_c in ("200", "201", "202", "203"):
            alvo_sala = tk_c
            break
        if tk_c in ("base", "ti", "laboratorio", "lab"):
            alvo_sala = "base"
            break

    if not alvo_sala and chat_id in TURMA_POR_GRUPO:
        alvo_sala = TURMA_POR_GRUPO[chat_id]

    if eh_cmd_modo:
        if not alvo_sala:
            linhas_m = ["⚙️ *ESTADO PÚBLICO / PRIVADO DAS SALAS:*\n"]
            for cod in ("200", "201", "202", "203", "base"):
                st = "📢 PÚBLICO (envia resumo no grupo)" if obter_estado_publico_sala(cod) else "🔒 PRIVADO (apenas no seu chat)"
                linhas_m.append(f"• *{NOMES_GRUPOS_TURMAS.get(cod, cod)}*: {st}")
            linhas_m.append("\n💡 *Para alterar:* `resumo 202 privado` ou `resumo 202 publico` (ou `!modo 202 privado`)")
            return "\n".join(linhas_m)
        if override_pub is None:
            override_pub = not obter_estado_publico_sala(alvo_sala)
        definir_estado_publico_sala(alvo_sala, override_pub)
        st_str = "📢 PÚBLICO (resumos enviados ao grupo da sala)" if override_pub else "🔒 PRIVADO (respostas apenas neste chat)"
        return f"✅ Estado da sala *{NOMES_GRUPOS_TURMAS.get(alvo_sala, alvo_sala)}* alterado para: *{st_str}*"

    if not alvo_sala:
        return gerar_resumo_matinal(chat_id)

    if override_pub is not None:
        definir_estado_publico_sala(alvo_sala, override_pub)
        eh_publico = override_pub
    else:
        eh_publico = obter_estado_publico_sala(alvo_sala)

    resumo_limpo = gerar_resumo_geral_sala(alvo_sala)
    grupo_jid = ID_GRUPO_BASE_TECNICA if alvo_sala == "base" else GRUPOS_TURMAS.get(alvo_sala)
    nome_grupo = NOMES_GRUPOS_TURMAS.get(alvo_sala, f"Sala {alvo_sala}")

    if eh_publico and grupo_jid:
        if chat_id == grupo_jid:
            return remover_dicas_comandos_para_grupo(resumo_limpo)
        ok = enviar_zap(grupo_jid, remover_dicas_comandos_para_grupo(resumo_limpo))
        status_env = f"📢 *Enviado no grupo:* _{nome_grupo}_ (Estado: *PÚBLICO*)" if ok else f"⚠️ _Falha ao enviar no grupo {nome_grupo}_"
        return (
            f"{resumo_limpo}\n\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"{status_env}\n"
            f"💡 _Para consultar apenas no privado sem enviar ao grupo: `resumo {alvo_sala} privado`_"
        )
    else:
        return (
            f"{resumo_limpo}\n\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"🔒 *Estado da Sala {alvo_sala.upper()}: PRIVADO* _(Exibido somente neste chat; nada foi enviado ao grupo {nome_grupo})_\n"
            f"💡 _Para publicar no grupo da sala: `resumo {alvo_sala} publico`_"
        )

MEMBROS_CACHE_BASE = set()
MAPA_NOMES_BASE = {}
ULTIMA_CONSULTA_MEMBROS = 0

GRADE_LAB = [
    (1, "07:30", "08:25"),
    (2, "08:25", "09:10"),
    (3, "09:30", "10:20"),
    (4, "10:20", "11:10"),
    (5, "11:10", "12:00"),
    (6, "13:30", "14:20"),
    (7, "14:20", "15:05"),
    (8, "15:20", "16:10"),
    (9, "16:10", "17:00"),
]

INTERVALOS_LAB = {
    2: "--------------------------------\n   Intervalo: 09:10 as 09:30\n--------------------------------",
    5: "--------------------------------\n   Almoco: 12:00 as 13:30\n--------------------------------",
    7: "--------------------------------\n   Intervalo: 15:05 as 15:20\n--------------------------------",
}

TEMPOS_LIMPEZA_SEGUNDA = {1, 2}

def obter_modo_limpeza_segunda() -> str:
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=5) as conn:
            cur = conn.cursor()
            cur.execute("SELECT valor FROM config_sistema WHERE chave = 'limpeza_segunda_modo'")
            row = cur.fetchone()
            if row and row[0]:
                return row[0].strip().lower()
    except Exception:
        pass
    return "padrao"

def definir_modo_limpeza_segunda(modo: str) -> bool:
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=5) as conn:
            conn.execute("INSERT OR REPLACE INTO config_sistema (chave, valor, atualizado_em) VALUES ('limpeza_segunda_modo', ?, datetime('now'))", (modo,))
            conn.commit()
            return True
    except Exception:
        return False

def eh_horario_limpeza_lab(wd: int, tempo: int) -> bool:
    """Segunda-feira (wd == 0) tem horários desabilitados para limpeza (configurável: padrao, manha ou integral)."""
    if wd != 0:
        return False
    modo = obter_modo_limpeza_segunda()
    if modo in ("integral", "dia_todo", "completo", "todos"):
        return True
    elif modo in ("manha", "matutino"):
        return tempo in {1, 2, 3, 4, 5}
    return tempo in TEMPOS_LIMPEZA_SEGUNDA

def extrair_identificador_limpo(jid: str) -> str:
    if not jid:
        return ""
    jid_base = str(jid).split("@")[0].split(":")[0]
    return re.sub(r"\D", "", jid_base)

def obter_membros_grupo_base() -> set:
    global ULTIMA_CONSULTA_MEMBROS, MEMBROS_CACHE_BASE
    agora_ts = time.time()
    if agora_ts - ULTIMA_CONSULTA_MEMBROS < 300 and MEMBROS_CACHE_BASE:
        return MEMBROS_CACHE_BASE
    try:
        url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/groups/{ID_GRUPO_BASE_TECNICA}"
        req = urllib.request.Request(url, headers={"x-api-key": OPENWA_API_KEY})
        with urllib.request.urlopen(req, timeout=3) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            novos = set()
            for p in data.get("participants", []):
                num = extrair_identificador_limpo(p.get("id") or p.get("number") or "")
                if num:
                    novos.add(num)
            if novos:
                MEMBROS_CACHE_BASE = novos
                ULTIMA_CONSULTA_MEMBROS = agora_ts
    except Exception as e:
        ULTIMA_CONSULTA_MEMBROS = agora_ts
        log_terminal(f"Aviso: falha ao consultar membros do grupo OpenWA ({e}). Usando cache.")
    return MEMBROS_CACHE_BASE

CACHE_LID_NUMERO = {}

def resolver_lid_para_numero(jid: str) -> str:
    """Converte um autor @lid no número de telefone real via OpenWA (cacheado). Retorna '' se falhar."""
    jid = str(jid or "")
    if "@lid" not in jid:
        return extrair_identificador_limpo(jid)
    if jid in CACHE_LID_NUMERO:
        return CACHE_LID_NUMERO[jid]
    try:
        url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/contacts/{urllib.parse.quote(jid)}"
        req = urllib.request.Request(url, headers={"x-api-key": OPENWA_API_KEY})
        with urllib.request.urlopen(req, timeout=3) as resp:
            dados = json.loads(resp.read().decode("utf-8"))
        id_ret = str(dados.get("id") or "")
        num = extrair_identificador_limpo(id_ret)
        if num and "@lid" not in id_ret:
            CACHE_LID_NUMERO[jid] = num
            return num
    except Exception as e:
        log_terminal(f"Aviso: falha ao resolver LID {jid} ({e}).")
    return ""

CSV_HORARIO_BASE = os.path.join(os.environ.get("DOCENTE_HOME", "."), "HORÁRIO CONFERIDO SEM CHOQUES.csv")
PROFS_BASE_MAP = {
    "pedro": "Prof. Pedro",
    "carlos": "Prof. Carlos",
    "marcos": "Prof. Marcos",
    "wanderson": "Prof. Wanderson",
    "patricia": "Profa. Patrícia",
    "bruna": "Profa. Bruna",
    "assis": "Prof. Assis",
    "osenias": "Prof. Osenias",
    "leandro": "Prof. Leandro",
}

def sincronizar_grade_base_csv(forcar: bool = False):
    """Popula grade_base_tecnica a partir do CSV HORÁRIO CONFERIDO SEM CHOQUES.csv se existir."""
    if not os.path.exists(CSV_HORARIO_BASE) or not os.path.exists(ESCOLA_DB):
        return
    try:
        if not forcar:
            with sqlite3.connect(ESCOLA_DB) as conn:
                cur = conn.cursor()
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='grade_base_tecnica'")
                if cur.fetchone():
                    cur.execute("SELECT COUNT(*) FROM grade_base_tecnica")
                    if cur.fetchone()[0] > 0:
                        return

        mapa_dias = {
            "2ª feira": 0, "3ª feira": 1, "4ª feira": 2, "5ª feira": 3, "6ª feira": 4
        }
        profs_rgx = {
            "pedro": (re.compile(r"\bpedro\b", re.IGNORECASE), "Prof. Pedro"),
            "carlos": (re.compile(r"\bcarlos\b", re.IGNORECASE), "Prof. Carlos"),
            "marcos": (re.compile(r"\bmarcos\s+sobral\b|\bsobral\b", re.IGNORECASE), "Prof. Marcos"),
            "wanderson": (re.compile(r"\bwanderson\b", re.IGNORECASE), "Prof. Wanderson"),
            "patricia": (re.compile(r"\bpatricia\b|\bpatrícia\b", re.IGNORECASE), "Profa. Patrícia"),
            "bruna": (re.compile(r"\bbruna\b", re.IGNORECASE), "Profa. Bruna"),
            "assis": (re.compile(r"\bassis\b|\bfrancisco de assis\b", re.IGNORECASE), "Prof. Assis"),
            "osenias": (re.compile(r"\bosenias\b", re.IGNORECASE), "Prof. Osenias"),
            "leandro": (re.compile(r"\bleandro\b", re.IGNORECASE), "Prof. Leandro"),
        }
        with open(CSV_HORARIO_BASE, encoding="utf-8", errors="replace") as f:
            rows = list(csv.reader(f))

        dia_atual = None
        turmas = []
        registros = []
        for r in rows:
            if not r: continue
            col0 = normalizar_texto(r[0]).strip().lower()
            if col0 in mapa_dias:
                dia_atual = mapa_dias[col0]
                turmas = [t.strip() for t in r[1:13]]
                continue
            m_tempo = re.match(r"^([1-9])", col0)
            if m_tempo and dia_atual is not None:
                tempo = int(m_tempo.group(1))
                for col_idx, celula in enumerate(r[1:13]):
                    if not celula.strip(): continue
                    turma = turmas[col_idx] if col_idx < len(turmas) else ""
                    for chave_p, (rgx, nome_exib) in profs_rgx.items():
                        if rgx.search(celula):
                            linhas_cel = [l.strip() for l in celula.strip().splitlines() if l.strip()]
                            disc = linhas_cel[0] if linhas_cel else celula.strip()
                            registros.append((dia_atual, tempo, turma, chave_p, nome_exib, disc))

        if registros:
            with sqlite3.connect(ESCOLA_DB) as conn:
                conn.execute("""
                    CREATE TABLE IF NOT EXISTS grade_base_tecnica (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        dia_semana INTEGER NOT NULL,
                        tempo_aula INTEGER NOT NULL,
                        turma TEXT NOT NULL,
                        prof_chave TEXT NOT NULL,
                        prof_nome TEXT NOT NULL,
                        disciplina TEXT NOT NULL
                    )
                """)
                conn.execute("DELETE FROM grade_base_tecnica")
                conn.executemany("""
                    INSERT INTO grade_base_tecnica (dia_semana, tempo_aula, turma, prof_chave, prof_nome, disciplina)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, registros)
                conn.commit()
            log_terminal(f"Grade da Base Tecnica sincronizada ({len(registros)} aulas).")
    except Exception as e:
        log_terminal(f"Erro ao sincronizar grade_base_tecnica: {e}")

def obter_vinculo_professor(numero: str):
    if not os.path.exists(ESCOLA_DB) or not numero:
        return None
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=15) as conn:
            cur = conn.cursor()
            cur.execute("SELECT chave, nome FROM professores_base_vinculos WHERE numero = ?", (numero,))
            return cur.fetchone()
    except Exception:
        return None

def salvar_vinculo_professor(numero: str, chave: str, nome: str):
    if not os.path.exists(ESCOLA_DB) or not numero:
        return
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=15) as conn:
            conn.execute("""
                INSERT OR REPLACE INTO professores_base_vinculos (numero, chave, nome)
                VALUES (?, ?, ?)
            """, (numero, chave, nome))
            conn.commit()
    except Exception:
        pass

# Preencha com {numero_ou_lid: ("chave", "Nome")} ou carregue de docentes/catalogo_docentes.json
PROFS_BASE_POR_NUMERO = {}

def identificar_prof_solicitante(author_num: str, author_raw: str, texto: str = "", pushname: str = ""):

    num_real = resolver_lid_para_numero(author_raw) if "@lid" in str(author_raw) else author_num
    if author_num in PROFS_BASE_POR_NUMERO:
        return PROFS_BASE_POR_NUMERO[author_num]
    if num_real in PROFS_BASE_POR_NUMERO:
        return PROFS_BASE_POR_NUMERO[num_real]

    if author_num:
        vinc = obter_vinculo_professor(author_num)
        if vinc:
            return vinc[0], vinc[1]
    if num_real:
        vinc = obter_vinculo_professor(num_real)
        if vinc:
            return vinc[0], vinc[1]

    combinado = f"{pushname} {texto}".lower()
    for chave, nome_exib in PROFS_BASE_MAP.items():
        if re.search(rf"\b{chave}\b", combinado):
            if author_num:
                salvar_vinculo_professor(author_num, chave, nome_exib)
            return chave, nome_exib

    if pushname and not re.search(r"\d{4,}", pushname):
        return "outro", f"Prof. {pushname.strip()}"

    return "outro", "Docente da Base Técnica"

def obter_nome_contato_completo(author_raw: str, pushname: str = "", default: str = "") -> str:
    author_raw = str(author_raw or "")
    author_num = extrair_identificador_limpo(author_raw)
    
    num_real = resolver_lid_para_numero(author_raw) if "@lid" in author_raw else author_num
    if author_num in PROFS_BASE_POR_NUMERO:
        return PROFS_BASE_POR_NUMERO[author_num][1]
    if num_real in PROFS_BASE_POR_NUMERO:
        return PROFS_BASE_POR_NUMERO[num_real][1]
        
    vinc = obter_vinculo_professor(author_num) or (obter_vinculo_professor(num_real) if num_real else None)
    if vinc:
        return vinc[1]

    if pushname and not re.search(r"\d{4,}", pushname):
        return pushname.strip()

    try:
        url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/contacts/{urllib.parse.quote(author_raw)}"
        req = urllib.request.Request(url, headers={"x-api-key": OPENWA_API_KEY})
        with urllib.request.urlopen(req, timeout=2) as resp:
            d = json.loads(resp.read().decode("utf-8"))
            c_name = (d.get("name") or d.get("pushname") or d.get("formattedName") or "").strip()
            if c_name and not re.search(r"^\+?\d{6,}$", c_name):
                return c_name
    except Exception:
        pass

    try:
        with sqlite3.connect(ESCOLA_DB) as conn_al:
            for n in (author_num, num_real):
                if n:
                    row = conn_al.execute("SELECT nome FROM alunos_whatsapp_vinculos WHERE numero = ?", (n,)).fetchone()
                    if row and row[0]:
                        return row[0].strip()
    except Exception:
        pass

    return default or author_num or "Contato"

LAB_RESERVA_LOCK = threading.RLock()

def sincronizar_notificacao_lab(evento: str, info: dict):
    """
    Sincroniza e faz broadcast enxuto dos agendamentos entre:
    - Grupo IEMA Base Técnica (ID_GRUPO_BASE_TECNICA)
    - Chat Privado do Prof. Pedro (CHAT_PRIVADO_PADRAO)
    - Grupo de Alunos / Sequestro (ID_GRUPO_ALUNOS_SEQUESTRO)
    - Privado do solicitante (se aplicável)
    """
    if os.environ.get("MODO_TESTE_ZAP") == "1" and not info.get("forcar_envio_teste"):
        # Ainda permite que mocks de enviar_zap capturem chamadas se necessário nos testes de aprovação/recusa
        pass
    origem = info.get("chat_id", "")
    solicitante = info.get("solicitante", "Usuário")
    data_fmt = info.get("data_fmt", "")
    dia_nome = info.get("dia_nome", "")
    tempos_str = info.get("tempos_str", "")
    horario_str = info.get("horario_str", "")
    motivo = info.get("motivo", "")
    mot_curto = f" | {motivo}" if motivo else ""
    eh_aluno = info.get("eh_aluno", False)
    pedido_id = info.get("pedido_id", "")
    grupo_aluno_alvo = info.get("grupo_origem") if info.get("grupo_origem") in GRUPOS_ALUNOS_LAB else ID_GRUPO_ALUNOS_SEQUESTRO

    if evento == "APROVADO_ALUNO":
        msg_alunos = (
            f"✅ *Lab Aprovado (#{pedido_id}):* {solicitante}\n"
            f"📅 {data_fmt} ({dia_nome}) • [{tempos_str}] ({horario_str}){mot_curto}"
        )
        if origem != grupo_aluno_alvo:
            try:
                enviar_zap(grupo_aluno_alvo, msg_alunos)
            except Exception as e:
                log_terminal(f"Aviso sync: erro ao notificar grupo alunos ({e})")

        msg_base = (
            f"📢 *Lab Reservado (Alunos):* {solicitante}\n"
            f"📅 {data_fmt} ({dia_nome}) • [{tempos_str}] ({horario_str}){mot_curto}"
        )
        if origem != ID_GRUPO_BASE_TECNICA:
            try:
                enviar_zap(ID_GRUPO_BASE_TECNICA, msg_base)
            except Exception as e:
                log_terminal(f"Aviso sync: erro ao notificar base tecnica ({e})")

        autor_jid = info.get("autor_jid", "")
        if autor_jid and "@g.us" not in autor_jid and autor_jid not in (CHAT_PRIVADO_PADRAO, origem):
            try:
                enviar_zap(autor_jid, msg_alunos)
            except Exception:
                pass

        # Notificação automática ao Prof. Leandro (novo agendamento aprovado)
        if origem != CHAT_PRIVADO_LEANDRO:
            try:
                enviar_zap(CHAT_PRIVADO_LEANDRO, msg_base)
            except Exception as e_leo:
                log_terminal(f"Aviso sync: erro ao notificar Prof. Leandro ({e_leo})")

    elif evento == "RESERVA_DOCENTE":
        msg_curta = (
            f"📌 *Lab Reservado:* {solicitante}\n"
            f"📅 {data_fmt} ({dia_nome}) • [{tempos_str}] ({horario_str}){mot_curto}"
        )
        # Notificação para o Grupo da Base Técnica
        if origem != ID_GRUPO_BASE_TECNICA:
            try:
                enviar_zap(ID_GRUPO_BASE_TECNICA, msg_curta)
            except Exception as e:
                log_terminal(f"Aviso sync: erro ao avisar base técnica ({e})")

        # Notificação exclusiva para o Prof. Leandro (conforme regra: apenas Leandro e Base Técnica recebem avisos)
        is_leandro = ("leandro" in str(solicitante).lower()) or (origem in (CHAT_PRIVADO_LEANDRO, DONO_CHAT_ID))
        if not is_leandro:
            try:
                enviar_zap(CHAT_PRIVADO_LEANDRO, msg_curta)
            except Exception as e_leo:
                log_terminal(f"Aviso sync: erro ao notificar Prof. Leandro ({e_leo})")

    elif evento == "CANCELADO":
        dia_txt = f" ({dia_nome})" if dia_nome else ""
        msg_canc = f"🔓 *Lab Liberado:* {data_fmt}{dia_txt} • [{tempos_str}] (por {solicitante})"
        if origem != ID_GRUPO_BASE_TECNICA and not info.get("apenas_pendente", False):
            try:
                enviar_zap(ID_GRUPO_BASE_TECNICA, msg_canc)
            except Exception:
                pass
        if eh_aluno and origem != grupo_aluno_alvo:
            try:
                enviar_zap(grupo_aluno_alvo, msg_canc)
            except Exception:
                pass

    elif evento == "RECUSADO_ALUNO":
        msg_rec = (
            f"❌ *Pedido #{pedido_id} recusado:* {solicitante} — {data_fmt} [{tempos_str}]{mot_curto}"
        )
        if origem != grupo_aluno_alvo:
            try:
                enviar_zap(grupo_aluno_alvo, msg_rec)
            except Exception:
                pass
        autor_jid = info.get("autor_jid", "")
        if autor_jid and "@g.us" not in autor_jid and autor_jid != origem:
            try:
                enviar_zap(autor_jid, msg_rec)
            except Exception:
                pass

def obter_aulas_prof_dia(dia_semana: int, prof_chave: str):
    if not os.path.exists(ESCOLA_DB):
        return []
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=15) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT tempo_aula, turma, disciplina 
                FROM grade_base_tecnica 
                WHERE dia_semana = ? AND prof_chave = ?
                ORDER BY tempo_aula ASC
            """, (dia_semana, prof_chave))
            return cur.fetchall()
    except Exception:
        return []

def init_tabela_laboratorio():
    if not os.path.exists(ESCOLA_DB):
        return
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=15) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS professores_base_vinculos (
                    numero TEXT PRIMARY KEY,
                    chave TEXT NOT NULL,
                    nome TEXT NOT NULL
                );
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS reservas_laboratorio (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_reserva TEXT NOT NULL,
                    tempo_aula INTEGER NOT NULL,
                    horario_inicio TEXT NOT NULL,
                    horario_fim TEXT NOT NULL,
                    autor_jid TEXT NOT NULL,
                    autor_numero TEXT NOT NULL,
                    autor_nome TEXT,
                    motivo TEXT,
                    mensagem_id TEXT,
                    criado_em TEXT NOT NULL,
                    status TEXT DEFAULT 'CONFIRMADO',
                    grupo_origem TEXT
                );
            """)
            try:
                conn.execute("ALTER TABLE reservas_laboratorio ADD COLUMN grupo_origem TEXT;")
            except Exception:
                pass
            conn.execute("""
                CREATE UNIQUE INDEX IF NOT EXISTS idx_lab_confirmado_unico
                ON reservas_laboratorio(data_reserva, tempo_aula)
                WHERE status = 'CONFIRMADO';
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS snapshot_grade_semanal (
                    semana_chave TEXT PRIMARY KEY,
                    snapshot_json TEXT NOT NULL,
                    atualizado_em TEXT NOT NULL
                );
            """)
        sincronizar_grade_base_csv()
    except Exception as e:
        log_terminal(f"Erro ao inicializar reservas_laboratorio: {e}")

MAPA_DIAS_LAB = {
    "segunda": 0, "seg": 0,
    "terca": 1, "ter": 1,
    "quarta": 2, "qua": 2,
    "quinta": 3, "qui": 3,
    "sexta": 4, "sex": 4,
    "sabado": 5, "sab": 5,
    "domingo": 6, "dom": 6,
}

def parsear_data_alvo_lab(termo: str):
    agora = datetime.now()
    t = normalizar_texto(termo).strip()
    if "amanha" in t:
        return agora + timedelta(days=1), re.sub(r"\bamanha\b", "", t).strip()
    if "hoje" in t:
        return agora, re.sub(r"\bhoje\b", "", t).strip()
    m = re.search(r"\b(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\b", t)
    if m:
        dia, mes = int(m.group(1)), int(m.group(2))
        ano = int(m.group(3)) if m.group(3) else agora.year
        if ano < 100: ano += 2000
        try:
            alvo = datetime(ano, mes, dia)
            resto_limpo = (t[:m.start()] + " " + t[m.end():]).strip()
            return alvo, resto_limpo
        except Exception:
            pass
    for d_str, wd_alvo in MAPA_DIAS_LAB.items():
        m_d = re.search(rf"\b{d_str}\b", t)
        if m_d:
            diff = (wd_alvo - agora.weekday()) % 7
            alvo = agora + timedelta(days=diff)
            resto_limpo = (t[:m_d.start()] + " " + t[m_d.end():]).strip()
            return alvo, resto_limpo
    return agora, t

def consultar_grade_lab(termo: str = "", author_raw: str = "", pushname: str = "", chat_id: str = "") -> str:
    init_tabela_laboratorio()
    alvo_dt, _ = parsear_data_alvo_lab(termo)
    data_str = alvo_dt.strftime("%d/%m/%Y")
    data_iso = alvo_dt.strftime("%Y-%m-%d")
    dias_semana_pt = ["Segunda", "Terca", "Quarta", "Quinta", "Sexta", "Sabado", "Domingo"]
    dia_nome = dias_semana_pt[alvo_dt.weekday()]
    eh_grupo_alunos = (chat_id in GRUPOS_ALUNOS_LAB)

    reservas = {}
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT tempo_aula, horario_inicio, horario_fim, autor_nome, autor_numero, motivo, status, id
                FROM reservas_laboratorio 
                WHERE data_reserva = ? AND status IN ('CONFIRMADO', 'PENDENTE')
                ORDER BY CASE WHEN status = 'CONFIRMADO' THEN 1 ELSE 0 END ASC, tempo_aula ASC
            """, (data_iso,))
            for r in cur.fetchall():
                reservas[r[0]] = r
    except Exception as e:
        return f"Erro ao acessar banco: {e}"

    author_num = extrair_identificador_limpo(author_raw)
    if eh_grupo_alunos:
        nome_aluno = pushname.strip() if pushname and not re.search(r"\d{4,}", pushname) else "Aluno(a)"
        prof_chave, prof_nome = "aluno", nome_aluno
        aulas_dia_lista = []
        aulas_dia = {}
    else:
        prof_chave, prof_nome = identificar_prof_solicitante(author_num, author_raw, termo, pushname)
        aulas_dia_lista = obter_aulas_prof_dia(alvo_dt.weekday(), prof_chave)
        aulas_dia = {t[0]: t for t in aulas_dia_lista}

    linhas = [
        f"🖥️ *LABORATÓRIO IEMA TI — {data_str} ({dia_nome})*",
        f"👋 Olá, *{prof_nome}*!"
    ]

    if not eh_grupo_alunos:
        if prof_chave != "outro" and aulas_dia_lista:
            linhas.append(f"\n📋 *Sua Grade ({dia_nome}):*")
            for t_a, t_turma, t_disc in aulas_dia_lista:
                linhas.append(f"• *{t_a}º Tempo:* Turma {t_turma} — {t_disc}")
        elif prof_chave != "outro":
            linhas.append(f"\n📋 *Sua Grade:* Você não possui aulas na grade técnica cadastradas para {dia_nome}.")
        else:
            linhas.append("\n💡 _Dica: Digite `!sou <seu nome>` para que o sistema identifique sua grade automaticamente._")

    linhas.append("\n━━━━━━━━━━━━━━━━━━━━")
    linhas.append("🕒 *Horários do Laboratório:*")

    for tempo, h_ini, h_fim in GRADE_LAB:
        if eh_horario_limpeza_lab(alvo_dt.weekday(), tempo):
            linhas.append(f"• [{tempo}] {h_ini}-{h_fim}  🧹 *Indisponível (Limpeza do Laboratório)*")
        elif tempo in reservas:
            _, _, _, nome, num_res, mot, st_res, id_res = reservas[tempo]
            mot_str = f" - {mot}" if mot else ""
            eh_reserva_aluno = "aluno" in normalizar_texto(nome)
            if st_res == "PENDENTE":
                linhas.append(f"• [{tempo}] {h_ini}-{h_fim}  🟠 *Pendente de aprovação* ({nome}{mot_str}) [#{id_res}]")
            elif num_res == author_num or (not eh_reserva_aluno and prof_chave not in ("outro", "aluno") and nome.lower().startswith(f"prof. {prof_chave}")):
                linhas.append(f"• [{tempo}] {h_ini}-{h_fim}  🟡 *Sua Reserva:* {mot or 'Confirmado'}")
            else:
                linhas.append(f"• [{tempo}] {h_ini}-{h_fim}  🔴 Reservado ({nome}{mot_str})")
        else:
            if tempo in aulas_dia:
                _, turma_prof, disc_prof = aulas_dia[tempo]
                linhas.append(f"• [{tempo}] {h_ini}-{h_fim}  🟢 *Livre para sua aula:* Turma {turma_prof} ({disc_prof})")
            else:
                linhas.append(f"• [{tempo}] {h_ini}-{h_fim}  🟢 Livre")

        if tempo in INTERVALOS_LAB:
            linhas.append(INTERVALOS_LAB[tempo])

    alvo_kw = "amanha" if alvo_dt.date() == (datetime.now() + timedelta(days=1)).date() else (alvo_dt.strftime("%d/%m") if alvo_dt.date() != datetime.now().date() else "")
    cmd_ex = f"lab {alvo_kw}".strip()
    ex_tempos_docente = "3 4" if alvo_dt.weekday() == 0 else "1 2"
    ex_label_docente = "3º e 4º" if alvo_dt.weekday() == 0 else "1º e 2º"
    if eh_grupo_alunos:
        linhas.extend([
            "",
            "📌 *Como solicitar (basta digitar no grupo):*",
            f"• `{cmd_ex} 6 7` _(pede o 6º e 7º horário)_",
            "• `lab cancelar` _(cancela seu pedido)_"
        ])
    else:
        linhas.extend([
            "",
            "📌 *Como reservar (basta digitar no grupo):*",
            f"• `{cmd_ex} {ex_tempos_docente}` _(reserva o {ex_label_docente} horário)_",
            "• `lab amanha 4 5` _(reserva para amanhã)_",
            "• `lab cancelar` _(libera seus horários do dia)_"
        ])
    return "\n".join(linhas)

def _extrair_tempos_e_motivo_lab(texto_restante: str):
    """
    Extrai tempos (1 a 9) do início do texto (aceitando '1 2', '1,2', '1 e 2', '1º e 2º')
    e retorna (lista_tempos, motivo_restante).
    """
    if not texto_restante:
        return [], ""
    conectores = {"e", "&", "a", "ao", "ate", "tempo", "tempos", "horario", "horarios", "aula", "aulas", "o", "os"}
    tokens = texto_restante.strip().split()
    tempos = []
    idx_fim_tempos = 0

    for i, tok in enumerate(tokens):
        sub_toks = [s for s in re.split(r"[,;]+", tok) if s]
        if not sub_toks:
            idx_fim_tempos = i + 1
            continue
        eh_bloco_tempo = True
        novos_tempos = []
        for st in sub_toks:
            st_low = st.lower().strip(".")
            if st_low in conectores:
                continue
            m_num = re.match(r"^([1-9])(?:º|ª|o|a|°)?$", st_low)
            if m_num:
                novos_tempos.append(int(m_num.group(1)))
            else:
                eh_bloco_tempo = False
                break
        if eh_bloco_tempo and (novos_tempos or all(s.lower().strip(".") in conectores for s in sub_toks)):
            for t_n in novos_tempos:
                if t_n not in tempos:
                    tempos.append(t_n)
            idx_fim_tempos = i + 1
        else:
            break

    motivo = " ".join(tokens[idx_fim_tempos:]).strip()
    return tempos, motivo

def agendar_lab(resto: str, author_raw: str, msg_id: str, pushname: str, chat_id: str = "") -> str:
    init_tabela_laboratorio()
    alvo_dt, texto_restante = parsear_data_alvo_lab(resto)
    data_str = alvo_dt.strftime("%d/%m/%Y")
    data_iso = alvo_dt.strftime("%Y-%m-%d")

    tempos, motivo = _extrair_tempos_e_motivo_lab(texto_restante)
    if not tempos:
        return (
            "Informe os números dos horários (de 1 a 9).\n"
            "Exemplo para hoje: `lab 1 2`\n"
            "Exemplo para amanhã: `lab amanha 4 5`"
        )

    author_num = extrair_identificador_limpo(author_raw)
    is_admin = author_num in IDENTIFICADORES_AUTORIZADOS or any(id_aut in str(author_raw) for id_aut in IDENTIFICADORES_AUTORIZADOS)

    num_real = resolver_lid_para_numero(author_raw) if "@lid" in str(author_raw) else author_num
    membros_base = obter_membros_grupo_base()
    eh_docente_base = (
        is_admin
        or (author_num in membros_base)
        or (num_real in membros_base)
        or (author_num in PROFS_BASE_POR_NUMERO)
        or (num_real in PROFS_BASE_POR_NUMERO)
        or (chat_id == ID_GRUPO_BASE_TECNICA)
    )
    eh_solicitacao_aluno = not eh_docente_base

    agora_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mapa_tempos = {t: (ini, fim) for t, ini, fim in GRADE_LAB}
    wd_alvo = alvo_dt.weekday()
    dias_semana_pt = ["Segunda", "Terca", "Quarta", "Quinta", "Sexta", "Sabado", "Domingo"]

    tempos_limpeza = [t for t in sorted(tempos) if eh_horario_limpeza_lab(wd_alvo, t)]
    if tempos_limpeza:
        tempos_fmt = ", ".join([f"{t}º" for t in tempos_limpeza])
        return (
            f"🧹 *Horário Indisponível ({data_str} - Segunda-feira):*\n"
            f"O(s) horário(s) ({tempos_fmt}) de segunda-feira está(ão) desabilitado(s) para limpeza e manutenção preventiva do laboratório."
        )

    if eh_solicitacao_aluno:
        nome_limpo = obter_nome_contato_completo(author_raw, pushname, default="Aluno")
        author_nome = nome_limpo if nome_limpo.lower().startswith("aluno") else f"Aluno(a) {nome_limpo}"
        try:
            with LAB_RESERVA_LOCK:
                with sqlite3.connect(ESCOLA_DB, timeout=30, isolation_level=None) as conn:
                    cur = conn.cursor()
                    cur.execute("BEGIN IMMEDIATE")
                    placeholders = ",".join(["?"] * len(tempos))
                    cur.execute(f"""
                        SELECT id, tempo_aula, autor_nome, status
                        FROM reservas_laboratorio
                        WHERE data_reserva = ? AND status IN ('CONFIRMADO', 'PENDENTE') AND tempo_aula IN ({placeholders})
                    """, [data_iso] + tempos)
                    ocupados_aluno = cur.fetchall()
                    if ocupados_aluno:
                        conn.execute("ROLLBACK")
                        conflitos_txt = []
                        for _, t_c, n_c, st_c in ocupados_aluno:
                            label_st = "reservado" if st_c == "CONFIRMADO" else "com solicitação pendente"
                            conflitos_txt.append(f"[{t_c}] {mapa_tempos[t_c][0]}-{mapa_tempos[t_c][1]} {label_st} por {n_c}")
                        return f"Horario indisponivel ({data_str}):\n" + "\n".join(conflitos_txt)

                    ids_gerados = []
                    for t in sorted(tempos):
                        h_ini, h_fim = mapa_tempos[t]
                        cur.execute("""
                            INSERT INTO reservas_laboratorio
                            (data_reserva, tempo_aula, horario_inicio, horario_fim, autor_jid, autor_numero, autor_nome, motivo, mensagem_id, criado_em, status, grupo_origem)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDENTE', ?)
                        """, (data_iso, t, h_ini, h_fim, author_raw, author_num, author_nome, motivo, msg_id, agora_iso, chat_id))
                        ids_gerados.append(cur.lastrowid)
                    conn.execute("COMMIT")

            pedido_id = ids_gerados[0]
            ini_geral = mapa_tempos[min(tempos)][0]
            fim_geral = mapa_tempos[max(tempos)][1]
            tempos_lista = ", ".join([str(t) for t in sorted(tempos)])
            motivo_curto = f" | {motivo}" if motivo else ""
            nome_grp_orig = GRUPOS_ALUNOS_LAB.get(chat_id, "Alunos")

            aviso_admin = (
                f"🔔 *Pedido de Lab #{pedido_id} ({nome_grp_orig})*\n"
                f"👤 {author_nome}\n"
                f"📅 {data_str} ({dias_semana_pt[wd_alvo]}) • [{tempos_lista}] ({ini_geral}-{fim_geral}){motivo_curto}\n"
                f"👉 Aprovar: `lab ok` | Recusar: `lab recusar {pedido_id}`"
            )
            try:
                enviar_zap(CHAT_PRIVADO_PADRAO, aviso_admin)
            except Exception as e_zap:
                log_terminal(f"Aviso: falha ao enviar notificacao de reserva pendente: {e_zap}")

            return (
                f"⏳ *Solicitação Pendente (#{pedido_id}):* {author_nome}\n"
                f"📅 {data_str} ({dias_semana_pt[wd_alvo]}) • [{tempos_lista}] ({ini_geral}-{fim_geral}){motivo_curto}\n"
                f"_Aguardando aprovação do Prof. Pedro._"
            )
        except Exception as e:
            return f"Erro ao solicitar agendamento: {e}"

    prof_chave, author_nome = identificar_prof_solicitante(author_num, author_raw, resto, pushname)

    if not motivo:
        aulas_docente = obter_aulas_prof_dia(wd_alvo, prof_chave)
        aulas_tempos = [f"{turma} - {disc}" for t_a, turma, disc in aulas_docente if t_a in tempos]
        if aulas_tempos:
            motivo = ", ".join(list(dict.fromkeys(aulas_tempos)))

    try:
        pendentes_cancelados_alunos = []
        with LAB_RESERVA_LOCK:
            with sqlite3.connect(ESCOLA_DB, timeout=30, isolation_level=None) as conn:
                cur = conn.cursor()
                cur.execute("BEGIN IMMEDIATE")
                placeholders = ",".join(["?"] * len(tempos))
                cur.execute(f"""
                    SELECT id, tempo_aula, autor_nome, autor_numero, motivo 
                    FROM reservas_laboratorio 
                    WHERE data_reserva = ? AND status = 'CONFIRMADO' AND tempo_aula IN ({placeholders})
                """, [data_iso] + tempos)
                ocupados = cur.fetchall()

                conflitos_bloqueantes = []
                tempos_para_substituir = []

                for id_r, t_c, n_c, num_c, m_c in ocupados:
                    # Regra de prioridade silenciosa do Pedro
                    if prof_chave == "pedro":
                        tempos_para_substituir.append(id_r)
                        continue

                    # Verifica se solicitante tem aula oficial em sala neste tempo
                    cur.execute("SELECT turma FROM grade_base_tecnica WHERE dia_semana = ? AND tempo_aula = ? AND prof_chave = ?", (wd_alvo, t_c, prof_chave))
                    solicitante_tem_aula = cur.fetchone() is not None

                    # Verifica se o ocupante atual tem aula oficial em sala neste tempo
                    ocupante_chave, _ = identificar_prof_solicitante(num_c, num_c, m_c, n_c)
                    cur.execute("SELECT turma FROM grade_base_tecnica WHERE dia_semana = ? AND tempo_aula = ? AND prof_chave = ?", (wd_alvo, t_c, ocupante_chave))
                    ocupante_tem_aula = cur.fetchone() is not None

                    # Solicitante tem aula de sala e ocupante nao tem: prioridade de sala
                    if solicitante_tem_aula and not ocupante_tem_aula and ocupante_chave != "pedro":
                        tempos_para_substituir.append(id_r)
                    else:
                        conflitos_bloqueantes.append(f"[{t_c}] {mapa_tempos[t_c][0]}-{mapa_tempos[t_c][1]} reservado por {n_c}")

                if conflitos_bloqueantes:
                    conn.execute("ROLLBACK")
                    return f"Horario indisponivel ({data_str}):\n" + "\n".join(conflitos_bloqueantes)

                for id_sub in tempos_para_substituir:
                    cur.execute("UPDATE reservas_laboratorio SET status = 'CANCELADO' WHERE id = ?", (id_sub,))

                # Verifica e cancela solicitações PENDENTES de alunos nos mesmos tempos reservados por docente
                cur.execute(f"""
                    SELECT id, tempo_aula, autor_nome, autor_jid
                    FROM reservas_laboratorio
                    WHERE data_reserva = ? AND status = 'PENDENTE' AND tempo_aula IN ({placeholders})
                """, [data_iso] + tempos)
                pendentes_cancelados_alunos = cur.fetchall()
                if pendentes_cancelados_alunos:
                    cur.execute(f"""
                        UPDATE reservas_laboratorio
                        SET status = 'CANCELADO'
                        WHERE data_reserva = ? AND status = 'PENDENTE' AND tempo_aula IN ({placeholders})
                    """, [data_iso] + tempos)

                for t in sorted(tempos):
                    h_ini, h_fim = mapa_tempos[t]
                    cur.execute("""
                        INSERT INTO reservas_laboratorio 
                        (data_reserva, tempo_aula, horario_inicio, horario_fim, autor_jid, autor_numero, autor_nome, motivo, mensagem_id, criado_em, status)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'CONFIRMADO')
                    """, (data_iso, t, h_ini, h_fim, author_raw, author_num, author_nome, motivo, msg_id, agora_iso))
                conn.execute("COMMIT")

        ini_geral = mapa_tempos[min(tempos)][0]
        fim_geral = mapa_tempos[max(tempos)][1]
        tempos_lista = ", ".join([str(t) for t in sorted(tempos)])
        motivo_curto = f" | {motivo}" if motivo else ""

        # Sincroniza broadcast nos canais (Base Técnica e Privado), exceto em testes automatizados
        if not str(msg_id).startswith("TEST_") or os.environ.get("FORCAR_SYNC_TESTE") == "1":
            sincronizar_notificacao_lab("RESERVA_DOCENTE", {
                "chat_id": chat_id,
                "solicitante": author_nome,
                "data_fmt": data_str,
                "dia_nome": dias_semana_pt[wd_alvo],
                "tempos_str": tempos_lista,
                "horario_str": f"{ini_geral} às {fim_geral}",
                "motivo": motivo,
                "is_pedro": (prof_chave == "pedro")
            })
            if pendentes_cancelados_alunos:
                t_pend = ", ".join(str(r[1]) for r in pendentes_cancelados_alunos)
                sincronizar_notificacao_lab("RECUSADO_ALUNO", {
                    "chat_id": chat_id,
                    "pedido_id": pendentes_cancelados_alunos[0][0],
                    "solicitante": pendentes_cancelados_alunos[0][2],
                    "data_fmt": data_str,
                    "tempos_str": t_pend,
                    "motivo": f"Horário assumido por {author_nome}",
                    "autor_jid": pendentes_cancelados_alunos[0][3],
                    "eh_aluno": True
                })

        return (
            f"✅ *Reserva Confirmada no Laboratório!*\n"
            f"👤 *Docente:* {author_nome}\n"
            f"📅 {data_str} ({dias_semana_pt[wd_alvo]}) • [{tempos_lista}] ({ini_geral}-{fim_geral}){motivo_curto}"
        )
    except Exception as e:
        return f"Erro ao agendar: {e}"

def calcular_segunda_semanal_lab(termo: str = "") -> datetime:
    """Calcula a data da segunda-feira de referência da semana solicitada para o laboratório."""
    agora = datetime.now()
    t_clean = normalizar_texto(termo).strip()
    if "proxima" in t_clean or "prox" in t_clean or (agora.weekday() >= 5 and "passada" not in t_clean and "anterior" not in t_clean):
        return agora - timedelta(days=agora.weekday()) + timedelta(days=7)
    return agora - timedelta(days=agora.weekday())


def consultar_grade_semanal_lab(termo: str = "") -> str:
    init_tabela_laboratorio()
    segunda = calcular_segunda_semanal_lab(termo)
    dias_uteis = [segunda + timedelta(days=i) for i in range(5)]
    datas_iso = [d.strftime("%Y-%m-%d") for d in dias_uteis]
    placeholders = ",".join(["?"] * len(datas_iso))

    try:
        with sqlite3.connect(ESCOLA_DB, timeout=15) as conn:
            cur = conn.cursor()
            cur.execute(f"""
                SELECT id, data_reserva, tempo_aula, horario_inicio, horario_fim, autor_nome, motivo 
                FROM reservas_laboratorio 
                WHERE data_reserva IN ({placeholders}) AND status = 'CONFIRMADO'
                ORDER BY data_reserva ASC, tempo_aula ASC
            """, datas_iso)
            reservas = cur.fetchall()
    except Exception as e:
        return f"Erro ao consultar semana: {e}"

    dias_pt = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
    linhas = [
        f"🖥️ *LABORATÓRIO IEMA TI — AGENDA DA SEMANA*",
        f"📅 *Período:* {dias_uteis[0].strftime('%d/%m')} a {dias_uteis[-1].strftime('%d/%m')}",
        "📎 _O PDF com a grade visual completa e fontes grandes foi gerado e anexado abaixo._\n",
        "━━━━━━━━━━━━━━━━━━━━"
    ]

    if not reservas:
        linhas.append("ℹ️ _Nenhuma reserva avulsa cadastrada para esta semana até o momento._")
        linhas.append("Todos os horários fora da grade técnica e da limpeza estão livres para agendamento!")
        return "\n".join(linhas)

    agrupado = {}
    for id_r, dt, tempo, ini, fim, nome, mot in reservas:
        agrupado.setdefault(dt, []).append((id_r, tempo, ini, fim, nome, mot))

    for d in dias_uteis:
        dt_str = d.strftime("%Y-%m-%d")
        if dt_str in agrupado:
            linhas.append(f"📌 *{dias_pt[d.weekday()]} ({d.strftime('%d/%m')}):*")
            for id_r, tempo, ini, fim, nome, mot in agrupado[dt_str]:
                mot_str = f" - _{mot}_" if mot else ""
                linhas.append(f"• *{tempo}º Tempo:* {nome}{mot_str} [#{id_r}]")
            linhas.append("")

    return "\n".join(linhas)

def processar_diff_grade_semanal_leandro(termo: str = "", pdf_path: str = None):
    """
    Quando o comando 'lab semana' é disparado no grupo, gera e envia para o Prof. Leandro no privado:
    1. Um texto detalhado com o resumo do que foi alterado na grade semanal desde a última versão gerada.
    2. O arquivo PDF atualizado da Grade Semanal em Alta Visibilidade.
    """
    init_tabela_laboratorio()
    agora = datetime.now()
    t_clean = normalizar_texto(termo).strip()
    if "proxima" in t_clean or "prox" in t_clean or (agora.weekday() >= 5 and "passada" not in t_clean and "anterior" not in t_clean):
        segunda_dt = agora - timedelta(days=agora.weekday()) + timedelta(days=7)
    else:
        segunda_dt = agora - timedelta(days=agora.weekday())

    semana_chave = segunda_dt.strftime("%Y-%m-%d")
    dias_uteis = [segunda_dt + timedelta(days=i) for i in range(5)]
    datas_iso = [d.strftime("%Y-%m-%d") for d in dias_uteis]
    placeholders = ",".join(["?"] * len(datas_iso))

    dias_pt = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira"]

    snapshot_antigo = None
    ultima_atualizacao = None

    try:
        with sqlite3.connect(ESCOLA_DB, timeout=15) as conn:
            cur = conn.cursor()
            cur.execute("SELECT snapshot_json, atualizado_em FROM snapshot_grade_semanal WHERE semana_chave = ?", (semana_chave,))
            row_snap = cur.fetchone()
            if row_snap:
                try:
                    snapshot_antigo = json.loads(row_snap[0])
                    ultima_atualizacao = row_snap[1]
                except Exception:
                    snapshot_antigo = None

            cur.execute(f"""
                SELECT id, data_reserva, tempo_aula, horario_inicio, horario_fim, autor_nome, motivo
                FROM reservas_laboratorio
                WHERE data_reserva IN ({placeholders}) AND status = 'CONFIRMADO'
                ORDER BY data_reserva ASC, tempo_aula ASC
            """, datas_iso)
            reservas_atuais = cur.fetchall()
    except Exception as e:
        log_terminal(f"Erro ao consultar dados para diff da grade semanal: {e}")
        return

    snapshot_atual = {}
    for id_r, dt, tempo, ini, fim, nome, mot in reservas_atuais:
        chave_slot = f"{dt}_t{tempo}"
        snapshot_atual[chave_slot] = {
            "id": id_r,
            "data": dt,
            "tempo": tempo,
            "inicio": ini,
            "fim": fim,
            "autor": (nome or "Docente").strip(),
            "motivo": (mot or "").strip()
        }

    periodo_str = f"{dias_uteis[0].strftime('%d/%m')} a {dias_uteis[-1].strftime('%d/%m')}"

    if snapshot_antigo is None:
        diff_linhas = [
            f"📋 *PRIMEIRA VERSÃO CONSOLIDADA DA SEMANA ({periodo_str})*",
            "Esta é a versão base registrada para a semana. As próximas consultas registrarão os deltas e alterações a partir daqui.\n"
        ]
        if snapshot_atual:
            diff_linhas.append("📌 *Horários atualmente agendados:*")
            for s in sorted(snapshot_atual.values(), key=lambda x: (x["data"], x["tempo"])):
                dt_obj = datetime.strptime(s["data"], "%Y-%m-%d")
                dia_n = dias_pt[dt_obj.weekday()]
                mot_str = f" - _{s['motivo']}_" if s["motivo"] else ""
                diff_linhas.append(f"• {dia_n} ({dt_obj.strftime('%d/%m')}) | Tempo [{s['tempo']}] {s['inicio']}-{s['fim']}: *{s['autor']}*{mot_str}")
        else:
            diff_linhas.append("ℹ️ Nenhum agendamento avulso registrado no momento. Todos os horários fora da grade regular estão livres.")
        diff_msg = "\n".join(diff_linhas)
    else:
        slots_novos = [s for k, s in snapshot_atual.items() if k not in snapshot_antigo]
        slots_removidos = [s for k, s in snapshot_antigo.items() if k not in snapshot_atual]
        slots_alterados = [
            (snapshot_antigo[k], snapshot_atual[k])
            for k in snapshot_atual
            if k in snapshot_antigo and (
                snapshot_atual[k]["autor"] != snapshot_antigo[k]["autor"] or
                snapshot_atual[k]["motivo"] != snapshot_antigo[k]["motivo"]
            )
        ]

        if not slots_novos and not slots_removidos and not slots_alterados:
            diff_msg = f"ℹ️ *Sem alterações:* A grade permanece idêntica à versão anterior (gerada em {ultima_atualizacao})."
        else:
            diff_linhas = [
                f"🔄 *RESUMO DAS ALTERAÇÕES NA GRADE ({periodo_str})*",
                f"🕒 _Comparado com a versão anterior de {ultima_atualizacao}_\n"
            ]
            if slots_novos:
                diff_linhas.append("➕ *Novos Horários Agendados:*")
                for s in sorted(slots_novos, key=lambda x: (x["data"], x["tempo"])):
                    dt_obj = datetime.strptime(s["data"], "%Y-%m-%d")
                    dia_n = dias_pt[dt_obj.weekday()]
                    mot_str = f" - _{s['motivo']}_" if s["motivo"] else ""
                    diff_linhas.append(f"• {dia_n} ({dt_obj.strftime('%d/%m')}) | Tempo [{s['tempo']}] {s['inicio']}-{s['fim']}: *{s['autor']}*{mot_str}")
                diff_linhas.append("")

            if slots_removidos:
                diff_linhas.append("➖ *Horários Liberados / Cancelados:*")
                for s in sorted(slots_removidos, key=lambda x: (x["data"], x["tempo"])):
                    dt_obj = datetime.strptime(s["data"], "%Y-%m-%d")
                    dia_n = dias_pt[dt_obj.weekday()]
                    diff_linhas.append(f"• {dia_n} ({dt_obj.strftime('%d/%m')}) | Tempo [{s['tempo']}] {s['inicio']}-{s['fim']}: Liberado (era: _{s['autor']}_)")
                diff_linhas.append("")

            if slots_alterados:
                diff_linhas.append("✏️ *Alterações:*")
                for ant, atu in slots_alterados:
                    dt_obj = datetime.strptime(atu["data"], "%Y-%m-%d")
                    dia_n = dias_pt[dt_obj.weekday()]
                    diff_linhas.append(f"• {dia_n} ({dt_obj.strftime('%d/%m')}) | Tempo [{atu['tempo']}]: {ant['autor']} ➔ *{atu['autor']}*")
                diff_linhas.append("")

            diff_msg = "\n".join(diff_linhas).strip()

    agora_str = agora.strftime("%d/%m/%Y às %H:%M")
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=15) as conn:
            conn.execute("""
                INSERT INTO snapshot_grade_semanal (semana_chave, snapshot_json, atualizado_em)
                VALUES (?, ?, ?)
                ON CONFLICT(semana_chave) DO UPDATE SET
                    snapshot_json = excluded.snapshot_json,
                    atualizado_em = excluded.atualizado_em;
            """, (semana_chave, json.dumps(snapshot_atual, ensure_ascii=False), agora_str))
    except Exception as e_save:
        log_terminal(f"Erro ao salvar snapshot da grade semanal: {e_save}")

    msg_leandro = (
        f"🖥️ *GRADE SEMANAL DO LABORATÓRIO GERADA NO GRUPO*\n\n"
        f"Olá, Prof. Leandro! O comando de grade semanal foi executado no grupo da Base Técnica.\n\n"
        f"{diff_msg}\n\n"
        f"📎 Segue abaixo a grade visual completa atualizada em PDF de alta visibilidade."
    )
    enviar_zap(CHAT_PRIVADO_LEANDRO, msg_leandro)

    caminho_pdf_envio = pdf_path
    if not caminho_pdf_envio or not os.path.exists(caminho_pdf_envio):
        try:
            import gerar_gabarito_lab
            caminho_pdf_envio = gerar_gabarito_lab.gerar_pdf_grade_semana_acessivel(segunda_dt)
        except Exception as e_pdf:
            log_terminal(f"Aviso ao gerar PDF da grade para Leandro: {e_pdf}")

    if caminho_pdf_envio and os.path.exists(caminho_pdf_envio):
        enviar_documento_zap(
            CHAT_PRIVADO_LEANDRO,
            caminho_pdf_envio,
            "Grade_Semanal_Laboratorio_IEMA.pdf",
            "application/pdf"
        )

def _cancelar_reserva_por_id_banco(id_alvo: int, author_num: str, is_admin: bool, chat_id: str = ""):
    try:
        with LAB_RESERVA_LOCK:
            with sqlite3.connect(ESCOLA_DB, timeout=30, isolation_level=None) as conn:
                cur = conn.cursor()
                cur.execute("BEGIN IMMEDIATE")
                cur.execute("""
                    SELECT id, data_reserva, tempo_aula, horario_inicio, horario_fim, autor_numero, autor_nome, mensagem_id, criado_em, status, COALESCE(grupo_origem, '')
                    FROM reservas_laboratorio
                    WHERE id = ? AND status IN ('CONFIRMADO', 'PENDENTE')
                """, (id_alvo,))
                reserva_id = cur.fetchone()
                if not reserva_id:
                    conn.execute("ROLLBACK")
                    return None
                if not (is_admin or reserva_id[5] == author_num):
                    conn.execute("ROLLBACK")
                    return "Sem permissao para cancelar horario de outro usuario."

                tempos_canc = [str(reserva_id[2])]
                if reserva_id[9] == "PENDENTE":
                    if reserva_id[7]:
                        cur.execute("SELECT tempo_aula FROM reservas_laboratorio WHERE mensagem_id = ? AND status = 'PENDENTE' ORDER BY tempo_aula ASC", (reserva_id[7],))
                        tempos_canc = [str(r[0]) for r in cur.fetchall()] or tempos_canc
                        cur.execute("UPDATE reservas_laboratorio SET status = 'CANCELADO' WHERE mensagem_id = ? AND status = 'PENDENTE'", (reserva_id[7],))
                    else:
                        cur.execute("SELECT tempo_aula FROM reservas_laboratorio WHERE data_reserva = ? AND autor_numero = ? AND criado_em = ? AND status = 'PENDENTE' ORDER BY tempo_aula ASC", (reserva_id[1], reserva_id[5], reserva_id[8]))
                        tempos_canc = [str(r[0]) for r in cur.fetchall()] or tempos_canc
                        cur.execute("UPDATE reservas_laboratorio SET status = 'CANCELADO' WHERE data_reserva = ? AND autor_numero = ? AND criado_em = ? AND status = 'PENDENTE'", (reserva_id[1], reserva_id[5], reserva_id[8]))
                else:
                    cur.execute("UPDATE reservas_laboratorio SET status = 'CANCELADO' WHERE id = ?", (id_alvo,))
                conn.execute("COMMIT")

        dt_obj = datetime.strptime(reserva_id[1], "%Y-%m-%d")
        dt_fmt = dt_obj.strftime("%d/%m/%Y")
        dias_pt = ["Segunda", "Terca", "Quarta", "Quinta", "Sexta", "Sabado", "Domingo"]
        sincronizar_notificacao_lab("CANCELADO", {
            "chat_id": chat_id,
            "grupo_origem": reserva_id[10] if len(reserva_id) > 10 else "",
            "solicitante": "Prof. Pedro" if is_admin else (reserva_id[6] or "Usuário"),
            "data_fmt": dt_fmt,
            "dia_nome": dias_pt[dt_obj.weekday()],
            "tempos_str": ", ".join(tempos_canc),
            "eh_aluno": "aluno" in str(reserva_id[6] or "").lower(),
            "apenas_pendente": (reserva_id[9] == "PENDENTE"),
        })
        return f"Agendamento [ID {id_alvo}] cancelado com sucesso ({dt_fmt} {reserva_id[3]}-{reserva_id[4]} - {reserva_id[6]})."
    except Exception as e:
        return f"Erro ao cancelar por ID: {e}"
    return None

def cancelar_lab(resto: str, author_raw: str, chat_id: str = "", pushname: str = "") -> str:
    init_tabela_laboratorio()
    author_num = extrair_identificador_limpo(author_raw)
    is_admin = author_num in IDENTIFICADORES_AUTORIZADOS or any(id_aut in str(author_raw) for id_aut in IDENTIFICADORES_AUTORIZADOS)
    t_clean = normalizar_texto(resto).strip()

    # Cancelamento por ID explícito (ex: !lab apagar #12, !lab cancelar id 12, ou número > 9)
    m_id = re.search(r"^(?:(id|#)\s*)?(\d+)$", t_clean)
    if m_id and (m_id.group(1) or int(m_id.group(2)) > 9):
        res_id = _cancelar_reserva_por_id_banco(int(m_id.group(2)), author_num, is_admin, chat_id=chat_id)
        if res_id:
            return res_id

    alvo_dt, texto_restante = parsear_data_alvo_lab(resto)
    data_str = alvo_dt.strftime("%d/%m/%Y")
    data_iso = alvo_dt.strftime("%Y-%m-%d")
    dias_pt = ["Segunda", "Terca", "Quarta", "Quinta", "Sexta", "Sabado", "Domingo"]

    tempos, _ = _extrair_tempos_e_motivo_lab(texto_restante)

    try:
        with LAB_RESERVA_LOCK:
            with sqlite3.connect(ESCOLA_DB, timeout=30, isolation_level=None) as conn:
                cur = conn.cursor()
                cur.execute("BEGIN IMMEDIATE")
                if not tempos:
                    # Sem números informados: cancela todas as reservas do próprio usuário na data alvo
                    cur.execute("""
                        SELECT id, tempo_aula, autor_nome, status
                        FROM reservas_laboratorio
                        WHERE data_reserva = ? AND autor_numero = ? AND status IN ('CONFIRMADO', 'PENDENTE')
                        ORDER BY tempo_aula ASC
                    """, (data_iso, author_num))
                    meus_itens = cur.fetchall()
                    if not meus_itens:
                        conn.execute("ROLLBACK")
                        return f"Você não possui reservas ativas em {data_str} para cancelar."
                    ids_canc = [r[0] for r in meus_itens]
                    tempos_canc = [str(r[1]) for r in meus_itens]
                    nome_sol = meus_itens[0][2] or obter_nome_contato_completo(author_raw, pushname, default="Usuário")
                    eh_al = any("aluno" in str(r[2] or "").lower() for r in meus_itens)
                    so_pend = all(r[3] == "PENDENTE" for r in meus_itens)
                    ph = ",".join(["?"] * len(ids_canc))
                    cur.execute(f"UPDATE reservas_laboratorio SET status = 'CANCELADO' WHERE id IN ({ph})", ids_canc)
                    conn.execute("COMMIT")
                else:
                    placeholders = ",".join(["?"] * len(tempos))
                    cur.execute(f"""
                        SELECT id, tempo_aula, autor_numero, autor_nome, status
                        FROM reservas_laboratorio 
                        WHERE data_reserva = ? AND status IN ('CONFIRMADO', 'PENDENTE') AND tempo_aula IN ({placeholders})
                    """, [data_iso] + tempos)
                    itens = cur.fetchall()

                    if not itens:
                        conn.execute("ROLLBACK")
                        if m_id:
                            res_id = _cancelar_reserva_por_id_banco(int(m_id.group(2)), author_num, is_admin, chat_id=chat_id)
                            if res_id:
                                return res_id
                        return f"Nenhum agendamento encontrado para esses horarios em {data_str}."

                    tempos_canc = []
                    for id_r, t_r, num_r, nome_r, st_r in itens:
                        if is_admin or num_r == author_num:
                            cur.execute("UPDATE reservas_laboratorio SET status = 'CANCELADO' WHERE id = ?", (id_r,))
                            tempos_canc.append(str(t_r))
                        else:
                            conn.execute("ROLLBACK")
                            return f"Sem permissao para cancelar horario de outro docente."
                    conn.execute("COMMIT")
                    nome_sol = "Prof. Pedro" if is_admin else (itens[0][3] or obter_nome_contato_completo(author_raw, pushname, default="Docente"))
                    eh_al = any("aluno" in str(r[3] or "").lower() for r in itens)
                    so_pend = all(r[4] == "PENDENTE" for r in itens)

        sincronizar_notificacao_lab("CANCELADO", {
            "chat_id": chat_id,
            "solicitante": nome_sol,
            "data_fmt": data_str,
            "dia_nome": dias_pt[alvo_dt.weekday()],
            "tempos_str": ", ".join(tempos_canc),
            "eh_aluno": eh_al,
            "apenas_pendente": so_pend,
        })

        return f"Horarios liberados ({data_str}): [{', '.join(tempos_canc)}]"
    except Exception as e:
        return f"Erro ao cancelar: {e}"

def listar_pendentes_lab() -> str:
    init_tabela_laboratorio()
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=15) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, data_reserva, tempo_aula, horario_inicio, horario_fim, autor_nome, motivo, mensagem_id, criado_em
                FROM reservas_laboratorio
                WHERE status = 'PENDENTE'
                ORDER BY data_reserva ASC, id ASC
            """)
            rows = cur.fetchall()
    except Exception as e:
        return f"Erro ao consultar solicitacoes pendentes: {e}"

    if not rows:
        return "✅ Nenhuma solicitação de laboratório pendente de aprovação."

    grupos = {}
    for id_r, dt_iso, tempo, ini, fim, nome, mot, msg_id, criado_em in rows:
        chave_grp = msg_id if msg_id else f"{dt_iso}_{nome}_{criado_em}"
        if chave_grp not in grupos:
            grupos[chave_grp] = {
                "pedido_id": id_r,
                "data_iso": dt_iso,
                "nome": nome,
                "motivo": mot,
                "tempos": [],
                "ini": ini,
                "fim": fim,
            }
        grupos[chave_grp]["tempos"].append(tempo)
        grupos[chave_grp]["fim"] = fim

    linhas = ["⏳ *SOLICITAÇÕES PENDENTES — LABORATÓRIO*", ""]
    for g in grupos.values():
        dt_fmt = datetime.strptime(g["data_iso"], "%Y-%m-%d").strftime("%d/%m/%Y")
        tempos_str = ", ".join(str(t) for t in sorted(g["tempos"]))
        mot_str = g["motivo"] or "Sem motivo informado"
        linhas.append(
            f"• *Pedido #{g['pedido_id']}* — {dt_fmt} | Tempos [{tempos_str}] ({g['ini']}-{g['fim']})\n"
            f"  👤 {g['nome']} | 📝 {mot_str}\n"
            f"  👉 `!lab aprovar {g['pedido_id']}` | `!lab recusar {g['pedido_id']}`"
        )
    return "\n".join(linhas)

def _obter_linhas_pedido_pendente(cur, id_alvo: int):
    cur.execute("""
        SELECT id, data_reserva, tempo_aula, horario_inicio, horario_fim, autor_jid, autor_numero, autor_nome, motivo, mensagem_id, criado_em, grupo_origem
        FROM reservas_laboratorio
        WHERE id = ? AND status = 'PENDENTE'
    """, (id_alvo,))
    alvo = cur.fetchone()
    if not alvo:
        return []
    _, dt_iso, _, _, _, autor_jid, autor_num, _, _, msg_id, criado_em, _ = alvo
    if msg_id:
        cur.execute("""
            SELECT id, data_reserva, tempo_aula, horario_inicio, horario_fim, autor_jid, autor_numero, autor_nome, motivo, mensagem_id, criado_em, grupo_origem
            FROM reservas_laboratorio
            WHERE mensagem_id = ? AND status = 'PENDENTE'
            ORDER BY tempo_aula ASC
        """, (msg_id,))
    else:
        cur.execute("""
            SELECT id, data_reserva, tempo_aula, horario_inicio, horario_fim, autor_jid, autor_numero, autor_nome, motivo, mensagem_id, criado_em, grupo_origem
            FROM reservas_laboratorio
            WHERE data_reserva = ? AND autor_numero = ? AND criado_em = ? AND status = 'PENDENTE'
            ORDER BY tempo_aula ASC
        """, (dt_iso, autor_num, criado_em))
    return cur.fetchall()

def aprovar_lab(resto: str, author_raw: str, chat_id: str = "") -> str:
    init_tabela_laboratorio()
    author_num = extrair_identificador_limpo(author_raw)
    is_admin = author_num in IDENTIFICADORES_AUTORIZADOS or any(id_aut in str(author_raw) for id_aut in IDENTIFICADORES_AUTORIZADOS)
    if not is_admin:
        return "Sem permissao. Apenas o Prof. Pedro pode aprovar solicitacoes do laboratorio."

    m_id = re.search(r"#?(\d+)", resto or "")
    id_alvo = int(m_id.group(1)) if m_id else None
    try:
        with LAB_RESERVA_LOCK:
            with sqlite3.connect(ESCOLA_DB, timeout=30, isolation_level=None) as conn:
                cur = conn.cursor()
                cur.execute("BEGIN IMMEDIATE")
                if id_alvo is None:
                    cur.execute("SELECT id FROM reservas_laboratorio WHERE status = 'PENDENTE' ORDER BY id DESC LIMIT 1")
                    row_ult = cur.fetchone()
                    if not row_ult:
                        conn.execute("ROLLBACK")
                        return "Não há nenhuma solicitação pendente no momento."
                    id_alvo = row_ult[0]
                itens = _obter_linhas_pedido_pendente(cur, id_alvo)
                if not itens:
                    conn.execute("ROLLBACK")
                    return f"Solicitação #{id_alvo} não encontrada ou já processada."

                dt_iso = itens[0][1]
                autor_jid = itens[0][5]
                autor_nome = itens[0][7]
                motivo = itens[0][8]
                grupo_orig = itens[0][11] if len(itens[0]) > 11 else None
                tempos = [r[2] for r in itens]
                ids_pedido = [r[0] for r in itens]

                placeholders = ",".join(["?"] * len(tempos))
                cur.execute(f"""
                    SELECT tempo_aula, autor_nome
                    FROM reservas_laboratorio
                    WHERE data_reserva = ? AND status = 'CONFIRMADO' AND tempo_aula IN ({placeholders})
                """, [dt_iso] + tempos)
                conflitos = cur.fetchall()
                if conflitos:
                    conn.execute("ROLLBACK")
                    conf_str = ", ".join(f"[{t}] ({n})" for t, n in conflitos)
                    return f"Não foi possível aprovar #{id_alvo}: horário já confirmado para {conf_str}."

                ph_ids = ",".join(["?"] * len(ids_pedido))
                cur.execute(f"UPDATE reservas_laboratorio SET status = 'CONFIRMADO' WHERE id IN ({ph_ids})", ids_pedido)
                # Cancela outros pedidos pendentes concorrentes para os mesmos tempos na mesma data
                cur.execute(f"""
                    UPDATE reservas_laboratorio
                    SET status = 'RECUSADO'
                    WHERE data_reserva = ? AND status = 'PENDENTE' AND tempo_aula IN ({placeholders}) AND id NOT IN ({ph_ids})
                """, [dt_iso] + tempos + ids_pedido)
                conn.execute("COMMIT")

        dt_obj = datetime.strptime(dt_iso, "%Y-%m-%d")
        dt_fmt = dt_obj.strftime("%d/%m/%Y")
        dias_semana_pt = ["Segunda", "Terca", "Quarta", "Quinta", "Sexta", "Sabado", "Domingo"]
        dia_nome = dias_semana_pt[dt_obj.weekday()]
        ini_geral = itens[0][3]
        fim_geral = itens[-1][4]
        tempos_lista = ", ".join(str(t) for t in sorted(tempos))

        # Sincroniza em broadcast nos canais (Grupo de Alunos solicitante, Base Técnica e Privado)
        sincronizar_notificacao_lab("APROVADO_ALUNO", {
            "chat_id": chat_id or CHAT_PRIVADO_PADRAO,
            "pedido_id": id_alvo,
            "solicitante": autor_nome,
            "data_fmt": dt_fmt,
            "dia_nome": dia_nome,
            "tempos_str": tempos_lista,
            "horario_str": f"{ini_geral} às {fim_geral}",
            "motivo": motivo,
            "autor_jid": autor_jid,
            "eh_aluno": True,
            "grupo_origem": grupo_orig,
        })

        return f"✅ Solicitação #{id_alvo} ({autor_nome} — {dt_fmt} [{tempos_lista}]) aprovada!"
    except Exception as e:
        return f"Erro ao aprovar solicitação: {e}"

def recusar_lab(resto: str, author_raw: str, chat_id: str = "") -> str:
    init_tabela_laboratorio()
    author_num = extrair_identificador_limpo(author_raw)
    is_admin = author_num in IDENTIFICADORES_AUTORIZADOS or any(id_aut in str(author_raw) for id_aut in IDENTIFICADORES_AUTORIZADOS)
    if not is_admin:
        return "Sem permissao. Apenas o Prof. Pedro pode recusar solicitacoes do laboratorio."

    m_id = re.search(r"^\s*(?:id\s*)?#?(\d+)\s*(.*)$", resto or "")
    if not m_id:
        return "Informe o ID da solicitacao (ex: `lab recusar 12 [motivo]`)."

    id_alvo = int(m_id.group(1))
    justificativa = m_id.group(2).strip()
    try:
        with LAB_RESERVA_LOCK:
            with sqlite3.connect(ESCOLA_DB, timeout=30, isolation_level=None) as conn:
                cur = conn.cursor()
                cur.execute("BEGIN IMMEDIATE")
                itens = _obter_linhas_pedido_pendente(cur, id_alvo)
                if not itens:
                    conn.execute("ROLLBACK")
                    return f"Solicitação #{id_alvo} não encontrada ou já processada."

                dt_iso = itens[0][1]
                autor_jid = itens[0][5]
                autor_nome = itens[0][7]
                grupo_orig = itens[0][11] if len(itens[0]) > 11 else None
                tempos = [r[2] for r in itens]
                ids_pedido = [r[0] for r in itens]

                ph_ids = ",".join(["?"] * len(ids_pedido))
                cur.execute(f"UPDATE reservas_laboratorio SET status = 'RECUSADO' WHERE id IN ({ph_ids})", ids_pedido)
                conn.execute("COMMIT")

        dt_fmt = datetime.strptime(dt_iso, "%Y-%m-%d").strftime("%d/%m/%Y")
        tempos_lista = ", ".join(str(t) for t in sorted(tempos))

        sincronizar_notificacao_lab("RECUSADO_ALUNO", {
            "chat_id": chat_id or CHAT_PRIVADO_PADRAO,
            "pedido_id": id_alvo,
            "solicitante": autor_nome,
            "data_fmt": dt_fmt,
            "tempos_str": tempos_lista,
            "motivo": justificativa,
            "autor_jid": autor_jid,
            "eh_aluno": True,
            "grupo_origem": grupo_orig,
        })

        return f"❌ Solicitação #{id_alvo} ({autor_nome} — {dt_fmt} [{tempos_lista}]) recusada."
    except Exception as e:
        return f"Erro ao recusar solicitação: {e}"

def _eh_comando_lab_msg(texto: str) -> bool:
    if not texto:
        return False
    t = normalizar_texto(texto).strip()
    return bool(re.match(r"^[/!]*\s*lab\b", t))

def processar_comando_lab(texto: str, data: dict) -> str:
    author_raw = data.get("author") or (data.get("sender") or {}).get("id") or ""
    pushname = (data.get("sender") or {}).get("pushname") or ""
    msg_id = str(data.get("id") or "")
    chat_id = str(data.get("chatId") or data.get("to") or data.get("from") or "")

    texto_sem_pref = re.sub(r"^[/!]*\s*lab\b\s*", "", texto.strip(), flags=re.IGNORECASE)
    t_clean = normalizar_texto(texto_sem_pref).strip()
    partes = t_clean.split(maxsplit=1)
    partes_orig = texto_sem_pref.split(maxsplit=1)
    subcmd = partes[0].lower() if partes else ""
    resto = partes[1] if len(partes) > 1 else ""
    resto_orig = partes_orig[1] if len(partes_orig) > 1 else resto

    if subcmd in ["agendar", "reservar", "marcar", "atribuir", "solicitar", "pedir"]:
        return agendar_lab(resto, author_raw, msg_id, pushname, chat_id=chat_id)
    elif subcmd in ["aprovar", "aceitar", "confirmar", "autorizar", "ok", "sim"]:
        return aprovar_lab(resto, author_raw, chat_id=chat_id)
    elif subcmd in ["recusar", "negar", "rejeitar", "nao"]:
        return recusar_lab(resto_orig, author_raw, chat_id=chat_id)
    elif subcmd in ["pendente", "pendentes", "pedidos", "solicitacoes"]:
        return listar_pendentes_lab()
    elif subcmd in ["cancelar", "desmarcar", "liberar", "apagar", "excluir"]:
        return cancelar_lab(resto, author_raw, chat_id=chat_id, pushname=pushname)
    elif subcmd in ["semana", "semanal", "grade"]:
        resp_txt = consultar_grade_semanal_lab(resto)
        chat_envio = chat_id or data.get("chatId") or ""
        pdf_path = None
        if chat_envio:
            try:
                import gerar_gabarito_lab
                segunda_alvo = calcular_segunda_semanal_lab(resto)
                pdf_path = gerar_gabarito_lab.gerar_pdf_grade_semana_acessivel(segunda_alvo, termo=resto)
                if pdf_path and os.path.exists(pdf_path):
                    enviar_documento_zap(
                        chat_envio,
                        pdf_path,
                        "Grade_Semanal_Laboratorio_IEMA.pdf",
                        "application/pdf"
                    )
            except Exception as e_pdf:
                log_terminal(f"Aviso ao enviar PDF da grade semanal: {e_pdf}")

        # Se foi gerado em grupo (Base Técnica ou qualquer grupo @g.us), notifica Leandro no privado com resumo do que mudou e o PDF
        is_grupo = str(chat_envio).endswith("@g.us") or chat_envio == ID_GRUPO_BASE_TECNICA
        if is_grupo:
            try:
                processar_diff_grade_semanal_leandro(resto, pdf_path)
            except Exception as e_diff:
                log_terminal(f"Aviso ao processar diff da grade semanal para Leandro: {e_diff}")

        return resp_txt
    elif subcmd in ["ajuda", "help", "menu"]:
        if chat_id in GRUPOS_ALUNOS_LAB:
            return (
                "🖥️ *LABORATÓRIO IEMA TI — ALUNOS*\n\n"
                "• `lab` → Ver horários de hoje\n"
                "• `lab amanha` → Ver horários de amanhã\n"
                "• `lab 6 7` → Pedir 6º e 7º horário de hoje\n"
                "• `lab amanha 6 7` → Pedir para amanhã\n"
                "• `lab cancelar` → Cancelar seu pedido"
            )
        return (
            "🖥️ *LABORATÓRIO IEMA TI — COMO USAR*\n\n"
            "• `lab` → Ver horários de hoje\n"
            "• `lab amanha` → Ver horários de amanhã\n"
            "• `lab 1 2` → Reservar 1º e 2º horário de hoje\n"
            "• `lab amanha 4 5` → Reservar 4º e 5º de amanhã\n"
            "• `lab cancelar` → Cancelar sua reserva de hoje\n"
            "• `lab semana` → Ver reservas da semana"
        )
    elif subcmd in ["sou", "identificar", "nome"]:
        if chat_id in GRUPOS_ALUNOS_LAB:
            return "Comando restrito a docentes."
        author_num = extrair_identificador_limpo(author_raw)
        nome_pesquisa = resto.strip().lower()
        for chave, nome_exib in PROFS_BASE_MAP.items():
            if chave in nome_pesquisa:
                salvar_vinculo_professor(author_num, chave, nome_exib)
                return f"Docente identificado com sucesso: {nome_exib}"
        if resto.strip():
            nome_livre = f"Prof. {resto.strip()}"
            salvar_vinculo_professor(author_num, "outro", nome_livre)
            return f"Docente identificado com sucesso: {nome_livre}"
        return "Informe seu nome (ex: !sou Marcos)."
    elif subcmd in ["lembrete", "lembretes"]:
        author_num = extrair_identificador_limpo(author_raw)
        p_chave, p_nome = identificar_prof_solicitante(author_num, author_raw, texto, pushname)
        alvo_dt, _ = parsear_data_alvo_lab(resto)
        aulas_dia = obter_aulas_prof_dia(alvo_dt.weekday(), p_chave)
        if not aulas_dia:
            return f"{p_nome}: nenhuma aula cadastrada na grade para {alvo_dt.strftime('%d/%m/%Y')}."
        linhas = [f"Lembrete de Aulas — {p_nome} ({alvo_dt.strftime('%d/%m/%Y')}):"]
        for t_a, t_turma, t_disc in aulas_dia:
            linhas.append(f"• {t_a}º Tempo: Turma {t_turma} — {t_disc}")
        linhas.append("\nPara reservar o lab: `lab 1 2`")
        return "\n".join(linhas)
    else:
        # Atalho direto para leigos: se houver números de tempo (ex: "lab 1 2" ou "lab amanha 4 5"), já agenda!
        _, resto_sem_data = parsear_data_alvo_lab(t_clean)
        tempos_diretos, _ = _extrair_tempos_e_motivo_lab(resto_sem_data)
        if tempos_diretos:
            return agendar_lab(texto_sem_pref, author_raw, msg_id, pushname, chat_id=chat_id)
        return consultar_grade_lab(t_clean, author_raw, pushname, chat_id=chat_id)

# ==========================================
# MÓDULO OSINT & GEOLOCALIZAÇÃO
# ==========================================
OSINT_MEDIA_DIR = os.path.join(os.environ.get("DOCENTE_HOME", "."), "osint_media")
os.makedirs(OSINT_MEDIA_DIR, exist_ok=True)
ULTIMA_MIDIA_CHAT = {}  # {chat_id: {"filepath": str, "timestamp": float}}

def extrair_id_mensagem(msg_obj) -> str:
    if not msg_obj:
        return ""
    if isinstance(msg_obj, str):
        return msg_obj
    if isinstance(msg_obj, dict):
        return msg_obj.get("_serialized") or msg_obj.get("id") or ""
    return str(msg_obj)

def baixar_midia_mensagem(chat_id: str, message_id: str, mimetype_hint: str = "") -> str:
    if not message_id or not chat_id:
        return None
    try:
        chat_enc = urllib.parse.quote(chat_id, safe="")
        msg_enc = urllib.parse.quote(message_id, safe="")
        url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/messages/{chat_enc}/{msg_enc}/media"
        req = urllib.request.Request(
            url,
            headers={"x-api-key": OPENWA_API_KEY},
            method="GET"
        )
        with urllib.request.urlopen(req, timeout=45) as resp:
            conteudo = resp.read()
            if not conteudo:
                return None
            ext = ".jpg"
            mime_l = (mimetype_hint or "").lower()
            if "png" in mime_l:
                ext = ".png"
            elif "pdf" in mime_l:
                ext = ".pdf"
            elif "webp" in mime_l:
                ext = ".webp"
            elif "image" in mime_l:
                ext = ".jpg"
            elif "ogg" in mime_l or "opus" in mime_l or "audio" in mime_l:
                ext = ".ogg"
            elif "mp3" in mime_l:
                ext = ".mp3"
            elif "m4a" in mime_l:
                ext = ".m4a"
            elif "wav" in mime_l:
                ext = ".wav"
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            salvo = os.path.join(OSINT_MEDIA_DIR, f"media_{ts}_{message_id[-8:]}{ext}")
            with open(salvo, "wb") as f_out:
                f_out.write(conteudo)
            log_terminal(f"📸 Mídia baixada com sucesso: {salvo} ({len(conteudo)} bytes)")
            return salvo
    except Exception as e:
        log_terminal(f"Erro ao baixar mídia do OpenWA para msg {message_id}: {e}")
        return None

def salvar_midia_base64(conteudo_b64: str, mimetype_hint: str = "") -> str:
    try:
        if "," in conteudo_b64:
            conteudo_b64 = conteudo_b64.split(",", 1)[1]
        dados = base64.b64decode(conteudo_b64)
        ext = ".jpg"
        mime_l = (mimetype_hint or "").lower()
        if "png" in mime_l:
            ext = ".png"
        elif "pdf" in mime_l:
            ext = ".pdf"
        elif "webp" in mime_l:
            ext = ".webp"
        elif "ogg" in mime_l or "opus" in mime_l or "audio" in mime_l:
            ext = ".ogg"
        elif "mp3" in mime_l:
            ext = ".mp3"
        elif "m4a" in mime_l:
            ext = ".m4a"
        elif "wav" in mime_l:
            ext = ".wav"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        salvo = os.path.join(OSINT_MEDIA_DIR, f"media_b64_{ts}{ext}")
        with open(salvo, "wb") as f:
            f.write(dados)
        return salvo
    except Exception as e:
        log_terminal(f"Erro ao decodificar base64: {e}")
        return None

def obter_midia_recente(data: dict, chat_id: str, permitir_cache_chat: bool = True) -> str:
    mimetype = data.get("mimetype") or ""
    msg_id = extrair_id_mensagem(data.get("id"))
    body = data.get("body") or ""

    # Se já baixamos exatamente esta mensagem há poucos segundos (thread de cache), reutiliza
    if msg_id and chat_id in ULTIMA_MIDIA_CHAT:
        reg_atual = ULTIMA_MIDIA_CHAT[chat_id]
        if reg_atual.get("msg_id") == msg_id and os.path.exists(reg_atual.get("filepath", "")):
            return reg_atual["filepath"]

    if data.get("hasMedia") or mimetype or data.get("type") in ["image", "document", "video", "audio", "ptt", "voice"]:
        if body.startswith("data:") or (len(body) > 300 and not body.startswith("http") and "\n" not in body):
            arq = salvar_midia_base64(body, mimetype)
            if arq:
                ULTIMA_MIDIA_CHAT[chat_id] = {"filepath": arq, "timestamp": time.time(), "msg_id": msg_id, "usada_em_visto": False}
                return arq
        arq = baixar_midia_mensagem(chat_id, msg_id, mimetype)
        if arq:
            ULTIMA_MIDIA_CHAT[chat_id] = {"filepath": arq, "timestamp": time.time(), "msg_id": msg_id, "usada_em_visto": False}
            return arq

    quoted = data.get("quotedMsg")
    if quoted and isinstance(quoted, dict):
        q_id = extrair_id_mensagem(quoted.get("id"))
        q_mime = quoted.get("mimetype") or ""
        q_body = quoted.get("body") or ""
        if q_id and chat_id in ULTIMA_MIDIA_CHAT:
            reg_atual = ULTIMA_MIDIA_CHAT[chat_id]
            if reg_atual.get("msg_id") == q_id and os.path.exists(reg_atual.get("filepath", "")):
                return reg_atual["filepath"]
        if quoted.get("hasMedia") or q_mime or quoted.get("type") in ["image", "document", "video", "audio", "ptt", "voice"]:
            if q_body.startswith("data:"):
                arq = salvar_midia_base64(q_body, q_mime)
                if arq:
                    ULTIMA_MIDIA_CHAT[chat_id] = {"filepath": arq, "timestamp": time.time(), "msg_id": q_id, "usada_em_visto": False}
                    return arq
            arq = baixar_midia_mensagem(chat_id, q_id, q_mime)
            if arq:
                ULTIMA_MIDIA_CHAT[chat_id] = {"filepath": arq, "timestamp": time.time(), "msg_id": q_id, "usada_em_visto": False}
                return arq

    if permitir_cache_chat and chat_id in ULTIMA_MIDIA_CHAT:
        reg = ULTIMA_MIDIA_CHAT[chat_id]
        if time.time() - reg.get("timestamp", 0) < 600 and os.path.exists(reg.get("filepath", "")):
            return reg["filepath"]

    return None

def transcrever_audio_omniroute(caminho_audio: str) -> str:
    """
    Transcreve arquivos de áudio (.ogg / .mp3 / .m4a / .wav) usando o endpoint
    /v1/audio/transcriptions do OmniRoute com o modelo Whisper da Groq.
    Rápido, preciso e sem consumo excessivo de tokens do LLM principal.
    """
    if not caminho_audio or not os.path.exists(caminho_audio):
        return ""

    url = f"{OMNIROUTE_URL.replace('/chat/completions', '')}/audio/transcriptions"
    nome_arq = os.path.basename(caminho_audio)
    ext_l = os.path.splitext(caminho_audio)[1].lower()

    mime_type = "audio/ogg"
    if ext_l == ".mp3":
        mime_type = "audio/mpeg"
    elif ext_l in [".m4a", ".mp4"]:
        mime_type = "audio/mp4"
    elif ext_l == ".wav":
        mime_type = "audio/wav"

    try:
        with open(caminho_audio, "rb") as f:
            conteudo_audio = f.read()

        boundary = "----WebKitFormBoundary" + os.urandom(16).hex()
        body = []
        body.append(f"--{boundary}\r\n".encode("utf-8"))
        body.append(b'Content-Disposition: form-data; name="model"\r\n\r\n')
        body.append(b"groq/whisper-large-v3\r\n")

        body.append(f"--{boundary}\r\n".encode("utf-8"))
        body.append(f'Content-Disposition: form-data; name="file"; filename="{nome_arq}"\r\n'.encode("utf-8"))
        body.append(f"Content-Type: {mime_type}\r\n\r\n".encode("utf-8"))
        body.append(conteudo_audio)
        body.append(b"\r\n")
        body.append(f"--{boundary}--\r\n".encode("utf-8"))

        payload = b"".join(body)
        req = urllib.request.Request(
            url,
            data=payload,
            headers={
                "Content-Type": f"multipart/form-data; boundary={boundary}",
                "Content-Length": str(len(payload))
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=35) as resp:
            res_json = json.loads(resp.read().decode("utf-8"))
            return (res_json.get("text") or "").strip()
    except Exception as e:
        log_terminal(f"❌ Erro ao transcrever áudio no OmniRoute: {e}")
        return ""

def processar_audio_instrucao_dono(chat_id: str, data: dict):
    """
    Processa exclusivamente notas de voz / áudios enviados pelo próprio Pedro no privado:
    1. Baixa o arquivo de áudio recebido (.ogg / .m4a / .mp3).
    2. Envia para o OmniRoute (Whisper da Groq) para transcrição ultrarrápida.
    3. Interpreta a transcrição como comando nativo ou consulta ao RAG do docente.
    4. Executa a instrução e responde no WhatsApp privado.
    """
    try:
        log_terminal(f"🎙️ [VOZ DONO] Iniciando download e transcrição via OmniRoute...")
        mimetype = data.get("mimetype") or "audio/ogg"
        msg_id = extrair_id_mensagem(data.get("id"))

        caminho_audio = None
        body = data.get("body") or ""
        if body.startswith("data:") or (len(body) > 300 and not body.startswith("http") and "\n" not in body):
            caminho_audio = salvar_midia_base64(body, mimetype)
        if not caminho_audio:
            caminho_audio = baixar_midia_mensagem(chat_id, msg_id, mimetype)

        if not caminho_audio or not os.path.exists(caminho_audio):
            log_terminal("❌ Falha ao obter arquivo de áudio do WhatsApp.")
            enviar_zap(chat_id, "⚠️ Não foi possível baixar o áudio enviado. Tente novamente.")
            return

        # 1. Transcreve o áudio via OmniRoute (Whisper da Groq)
        texto_transcrito = transcrever_audio_omniroute(caminho_audio)
        if not texto_transcrito or not texto_transcrito.strip():
            log_terminal("❌ Transcrição do áudio retornou vazia.")
            enviar_zap(chat_id, "⚠️ Não consegui entender o áudio com clareza. Pode repetir?")
            return

        texto_transcrito = texto_transcrito.strip()
        log_terminal(f"🎙️ [VOZ DONO] Transcrição com sucesso: '{texto_transcrito}'")
        enviar_zap(chat_id, f"🎙️ *Voz Reconhecida:* \"_{texto_transcrito}_\"")

        # 2. Normaliza para identificar se é comando do sistema ou consulta ao RAG
        t_norm = normalizar_texto(texto_transcrito)

        # Mapeamento inteligente de linguagem falada para comandos docentes
        cmd_exec = texto_transcrito
        if any(p in t_norm for p in ["qual meu horario", "meus horarios", "aulas de hoje", "horario de hoje", "minha grade hoje"]):
            cmd_exec = "!horario hoje"
        elif any(p in t_norm for p in ["aulas de amanha", "horario de amanha", "minha grade amanha"]):
            cmd_exec = "!horario amanha"
        elif any(p in t_norm for p in ["grade da semana", "minha grade", "grade horaria"]):
            cmd_exec = "!grade"
        elif "laboratorio" in t_norm or "lab" in t_norm:
            if "amanha" in t_norm:
                cmd_exec = "lab amanha"
            elif "semana" in t_norm:
                cmd_exec = "lab semana"
            else:
                cmd_exec = "lab"
        elif re.search(r"\bturma\s*(20[0-3])\b", t_norm):
            m_tm = re.search(r"\bturma\s*(20[0-3])\b", t_norm)
            cmd_exec = f"!turma {m_tm.group(1)}"
        elif t_norm.startswith(("status", "como esta o servidor", "como estao os servicos")):
            cmd_exec = "!status"
        elif t_norm.startswith(("preparar aula", "preparar aulas", "planejar aula")):
            cmd_exec = "!preparar amanha" if "amanha" in t_norm else "!preparar hoje"

        # Se for comando do sistema (! ou mapeado)
        if cmd_exec.startswith(("!", "/")) or _eh_comando_valido_sem_prefixo(normalizar_texto(cmd_exec)):
            processar_comando(chat_id, cmd_exec, data)
            return

        # 3. Se for instrução pedagógica, dúvida ou tema para o RAG Docente:
        # Busca no Omni RAG do Pedro e responde sintetizado pelo OmniRoute LLM
        log_terminal(f"🧠 [VOZ DONO] Roteando instrução falada para o Omni RAG: '{texto_transcrito}'")
        contexto_rag = gestor_docentes_rag.buscar_rag_docente("pedro", texto_transcrito, top_k=3)
        trechos_rag = []
        if contexto_rag:
            for r in contexto_rag:
                txt_c = r.get("conteudo") or r.get("texto") or ""
                if txt_c:
                    trechos_rag.append(f"[{r.get('materia', 'Docência')}]: {txt_c[:500]}")
        bloco_rag = "\n\n".join(trechos_rag) if trechos_rag else "Nenhum documento específico na base."

        prompt_interpretacao = (
            "Você é o Assistente Pessoal do Professor Pedro Oliveira (IEMA TI).\n"
            f"O professor acabou de enviar uma nota de voz no WhatsApp com a seguinte instrução/dúvida:\n"
            f"\"{texto_transcrito}\"\n\n"
            f"Base de Conhecimento RAG do Professor Pedro:\n{bloco_rag}\n\n"
            "Diretrizes de resposta:\n"
            "1. Responda de forma direta, clara e útil no WhatsApp.\n"
            "2. Se for uma ordem ou pedido de planejamento, estruture a resposta de forma prática.\n"
            "3. Se for uma dúvida conceitual, responda com rigor técnico e didática."
        )

        resposta_ia = consultar_llm(prompt_interpretacao)
        if resposta_ia:
            enviar_zap(chat_id, resposta_ia.strip())
        else:
            enviar_zap(chat_id, "⚠️ Não obtive resposta do modelo para essa instrução de voz.")

    except Exception as e_voz:
        log_terminal(f"❌ Erro no pipeline de voz OmniRoute: {e_voz}")
        enviar_zap(chat_id, f"❌ Erro ao processar áudio: {e_voz}")

def _preparar_imagem_vision(caminho_imagem: str):
    """Valida a imagem pelo conteúdo real, corrige rotação (EXIF) e reduz p/ ~1600px. Retorna (b64, mime) ou (None, motivo)."""
    try:
        from PIL import Image, ImageOps
        import io
        with Image.open(caminho_imagem) as im:
            im = ImageOps.exif_transpose(im).convert("RGB")
            im.thumbnail((1600, 1600))
            buf = io.BytesIO()
            im.save(buf, format="JPEG", quality=85)
        return base64.b64encode(buf.getvalue()).decode("utf-8"), "image/jpeg"
    except Exception as e:
        return None, f"arquivo não é uma imagem válida ({type(e).__name__})"

def analisar_foto_caderno_ia(caminho_imagem: str, titulo_atividade: str = "") -> str:
    """Usa o modelo multimodal (Vision) via OmniRoute para ler e avaliar a foto do caderno."""
    if DESABILITAR_LEITURA_IMAGEM_OMNIROUTE:
        log_terminal("🛡️ [SEGURANÇA] Leitura de imagem via OmniRoute desabilitada. Chamada Vision ignorada.")
        return ""
    if not caminho_imagem or not os.path.exists(caminho_imagem):
        return "Arquivo de imagem não encontrado."
    b64, mime = _preparar_imagem_vision(caminho_imagem)
    if not b64:
        return f"Não foi possível analisar: {mime}."
    prompt = (
        "Você é um assistente pedagógico corrigindo a foto do caderno ou folha de respostas de um aluno do Ensino Médio Técnico.\n"
        f"Atividade de referência: {titulo_atividade or 'Exercícios de Fixação'}\n\n"
        "Se a imagem NÃO for um caderno/folha com texto manuscrito ou impresso de atividade, diga isso e use Parecer: Não é caderno.\n"
        "Analise visualmente a imagem e responda estritamente neste formato conciso:\n"
        "• **Legibilidade:** (Boa / Parcial / Ruim)\n"
        "• **Conteúdo Identificado:** (ex: 3 questões resolvidas no caderno)\n"
        "• **Síntese das Respostas:** (1 frase resumindo o que o aluno escreveu)\n"
        "• **Parecer:** (Aprovado com Visto / Incompleto / Não é caderno)"
    )
    payload = json.dumps({
        "model": "no-think/agy/claude-sonnet-5-5-low",
        "stream": False,
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": prompt},
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}}
        ]}]
    }).encode("utf-8")
    ultimo_erro = None
    for tentativa in (1, 2):
        try:
            req = urllib.request.Request(OMNIROUTE_URL, data=payload, headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(req, timeout=60) as resp:
                dados = json.loads(resp.read().decode("utf-8"))
                return dados["choices"][0]["message"]["content"].strip()
        except Exception as e:
            ultimo_erro = e
            log_terminal(f"Erro na análise visual de foto (tentativa {tentativa}): {e}")
            time.sleep(2)
    return "Análise visual indisponível no momento. Registro manual aplicado."

def obter_exercicio_caderno_aula(termo: str = "amanha") -> str:
    """Retorna os exercícios simples de caderno para copiar no quadro ou enviar aos alunos."""
    t_clean = normalizar_texto(termo).strip()
    m_id = re.search(r"\b(\d+)\b", t_clean)
    id_busca = int(m_id.group(1)) if m_id else None

    with sqlite3.connect(ESCOLA_DB) as conn:
        cur = conn.cursor()
        for col_def in ("caderno_json TEXT DEFAULT ''", "gabarito_json TEXT DEFAULT ''"):
            try:
                cur.execute(f"ALTER TABLE atividades_cadastradas ADD COLUMN {col_def}")
            except Exception:
                pass
        if id_busca:
            cur.execute("SELECT id, turma, componente, titulo, data_passada, COALESCE(caderno_json, '') FROM atividades_cadastradas WHERE id = ?", (id_busca,))
        else:
            cur.execute("SELECT id, turma, componente, titulo, data_passada, COALESCE(caderno_json, '') FROM atividades_cadastradas ORDER BY id DESC LIMIT 1")
        row = cur.fetchone()

    if not row:
        return "Nenhuma atividade encontrada no momento."

    id_atv, turma, comp, titulo, dt_p, cad_json = row
    questoes_caderno = []
    if cad_json:
        try:
            parsed = json.loads(cad_json)
            if isinstance(parsed, list) and parsed:
                questoes_caderno = [str(x).strip() for x in parsed if str(x).strip()]
        except Exception:
            pass

    if not questoes_caderno:
        comp_norm = normalizar_texto(comp)
        if "legisla" in comp_norm:
            questoes_caderno = [
                "1. O que é 'Prompt Injection Indireto' e por que ele é considerado o 'SQL Injection' dos modelos de IA?",
                "2. Segundo o Artigo 20 da LGPD, qual direito o cidadão possui quando uma decisão é tomada exclusivamente por algoritmos ou IA?",
                "3. Explique qual o risco de segurança e jurídico quando um Agente autônomo conectado ao protocolo MCP tem permissão ilimitada de ferramentas.",
            ]
        else:
            questoes_caderno = [
                "1. Explique resumidamente os 4 passos do ciclo ReAct (Reasoning + Acting) em um Agente Autônomo de IA.",
                "2. Qual é o objetivo principal do protocolo MCP (Model Context Protocol) e por que ele é um padrão aberto?",
                "3. Em um sistema de RAG Avançado, para que serve o componente de 'Reranking' após a busca vetorial por similaridade?",
            ]

    corpo_q = "\n\n".join(questoes_caderno)
    return (
        f"📝 *EXERCÍCIOS PARA O CADERNO — COPIAR E RESOLVER*\n"
        f"*{comp}* | Turma {turma}\n"
        f"_{titulo}_\n"
        f"📅 Data: {dt_p} | 📌 Atividade ID: `{id_atv}`\n\n"
        f"{corpo_q}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"📌 *Orientações:* Responder no caderno (2 a 3 linhas por questão).\n"
        f"✅ *Visto do Professor:* Enviar foto do caderno com comando `atv {id_atv} ok` ou apresentar em sala!"
    )

def parse_gps_coord_str(val_str: str):
    if not val_str:
        return None
    val_str = str(val_str).strip()
    partes = val_str.split()
    if not partes:
        return None
    try:
        num = float(partes[0])
        if len(partes) > 1 and partes[1].upper() in ["S", "W"]:
            num = -abs(num)
        return num
    except Exception:
        return None

def extrair_metadados_exiftool(filepath: str) -> dict:
    try:
        cmd = ["exiftool", "-json", "-c", "%.6f", filepath]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
        if res.returncode == 0 and res.stdout.strip():
            dados = json.loads(res.stdout)
            if dados and isinstance(dados, list):
                return dados[0]
    except Exception as e:
        log_terminal(f"Erro no exiftool: {e}")
    return {}

def extrair_coordenadas_gps(filepath: str) -> dict:
    meta = extrair_metadados_exiftool(filepath)
    if not meta:
        return None

    lat_raw = meta.get("GPSLatitude")
    lon_raw = meta.get("GPSLongitude")

    lat = parse_gps_coord_str(lat_raw)
    lon = parse_gps_coord_str(lon_raw)

    if lat is not None and lon is not None:
        data_foto = meta.get("DateTimeOriginal") or meta.get("CreateDate") or meta.get("FileModifyDate") or "Não informada"
        dispositivo = f"{meta.get('Make', '')} {meta.get('Model', '')}".strip() or "Desconhecido"
        maps_url = f"https://www.google.com/maps?q={lat},{lon}"
        osm_url = f"https://www.openstreetmap.org/?mlat={lat}&mlon={lon}#map=16/{lat}/{lon}"
        return {
            "lat": lat,
            "lon": lon,
            "maps_url": maps_url,
            "osm_url": osm_url,
            "data_foto": data_foto,
            "dispositivo": dispositivo,
            "altitude": meta.get("GPSAltitude", "N/A")
        }
    return None

def geocode_cidade_nominatim(nome_local: str) -> dict | None:
    """Consulta o serviço Nominatim OpenStreetMap para obter coordenadas estimadas de um município/região."""
    try:
        url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode({
            "q": nome_local,
            "format": "json",
            "limit": 1
        })
        req = urllib.request.Request(url, headers={"User-Agent": "DocenteGeoSleuth/1.0"})
        with urllib.request.urlopen(req, timeout=5) as res:
            dados = json.loads(res.read().decode("utf-8"))
            if dados:
                return {
                    "lat": float(dados[0]["lat"]),
                    "lon": float(dados[0]["lon"]),
                    "display_name": dados[0].get("display_name", "")
                }
    except Exception as e:
        log_terminal(f"Aviso Nominatim ({nome_local}): {e}")
    return None

def extrair_locais_candidatos_geointel(guess_list: list, links: list, ocr_items: list) -> list:
    """Extrai nomes limpos de municípios, estados e locais de destaque dos resultados do Yandex e OCR."""
    candidatos = []

    for l in links:
        t = (l.get("title") or "").strip()
        if not t or re.search(r"^\d+×\d+$", t):
            continue
        # Padrão município específico (ex: "Icatu é um município do litoral central do Maranhão...")
        m_mun = re.search(r"([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][a-záéíóúâêîôûãõç]+)\s+é um município[^\.]*?(?:de\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][a-záéíóúâêîôûãõç]+)|(Maranh[ãa]o))", t, re.IGNORECASE)
        if m_mun:
            cidade = m_mun.group(1)
            uf = m_mun.group(3) or m_mun.group(2) or "Maranhão"
            candidatos.append(f"{cidade}, {uf}")
            continue

        # Padrão Cidade - UF (ex: "Brazópolis-MG", "Barreirinhas - MA")
        m_uf = re.search(r"([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][a-záéíóúâêîôûãõç\s]{2,25}?)\s*-\s*([A-Z]{2})\b", t)
        if m_uf and not any(k in m_uf.group(1).lower() for k in ["foto", "hotel", "pousada", "rio", "fazenda"]):
            candidatos.append(f"{m_uf.group(1).strip()}, {m_uf.group(2)}")
            continue

        # Padrão Cidade em Estado (ex: "Barreirinhas, no Maranhão", "Vitoria Do Mearim - State of Maranhao")
        m_est = re.search(r"([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][a-záéíóúâêîôûãõç\s]{2,25}?)(?:,\s*no|\s*,\s*|\s+no\s+|\s+em\s+)(Maranh[ãa]o|Par[áa]|Cear[áa]|Piau[íi]|Bahia|Minas Gerais|S[ãa]o Paulo|Rio de Janeiro|Amazonas)", t, re.IGNORECASE)
        if m_est and not any(k in m_est.group(1).lower() for k in ["onde", "como", "quem"]):
            candidatos.append(f"{m_est.group(1).strip()}, {m_est.group(2).strip()}")
            continue

    for txt in ocr_items:
        if len(txt) >= 4 and any(uf in txt.upper() for uf in ["/MA", "-MA", "/SP", "-SP", "/RJ", "-RJ", "/MG", "-MG"]):
            candidatos.append(txt)

    vistos = set()
    res = []
    for c in candidatos:
        c_norm = c.lower().strip()
        if c_norm not in vistos:
            vistos.add(c_norm)
            res.append(c)
    return res

def limpar_midias_temporarias(caminho_arquivo: str = None, diretorio_extra: str = None):
    """Remove com segurança a mídia baixada e qualquer diretório de análise temporário."""
    try:
        if caminho_arquivo and os.path.exists(caminho_arquivo):
            os.remove(caminho_arquivo)
            log_terminal(f"🧹 Mídia baixada removida após análise: {caminho_arquivo}")
    except Exception as e:
        log_terminal(f"Aviso ao remover mídia temporária ({caminho_arquivo}): {e}")

    try:
        if diretorio_extra and os.path.exists(diretorio_extra):
            shutil.rmtree(diretorio_extra, ignore_errors=True)
            log_terminal(f"🧹 Diretório temporário removido: {diretorio_extra}")
    except Exception as e:
        log_terminal(f"Aviso ao remover diretório temporário ({diretorio_extra}): {e}")

def limpar_arquivos_antigos_osint(max_idade_seg: int = 600):
    """Varre e remove mídias em osint_media e pastas /tmp que já foram analisadas ou expiraram."""
    agora = time.time()
    try:
        if os.path.exists(OSINT_MEDIA_DIR):
            for f in os.listdir(OSINT_MEDIA_DIR):
                caminho = os.path.join(OSINT_MEDIA_DIR, f)
                if os.path.isfile(caminho) and (agora - os.path.getmtime(caminho)) > max_idade_seg:
                    try:
                        os.remove(caminho)
                        log_terminal(f"🧹 Faxina OSINT: mídia antiga removida {f}")
                    except Exception:
                        pass
        for d in os.listdir("/tmp"):
            if d.startswith(("geosleuth_", "intake_")):
                caminho_d = os.path.join("/tmp", d)
                if os.path.isdir(caminho_d) and (agora - os.path.getmtime(caminho_d)) > max_idade_seg:
                    try:
                        shutil.rmtree(caminho_d, ignore_errors=True)
                    except Exception:
                        pass
    except Exception as e:
        log_terminal(f"Aviso faxina OSINT: {e}")

def worker_geolocalizar_imagem(chat_id: str, filepath: str):
    out_dir = None
    try:
        enviar_zap(chat_id, "🔍 *Geo-Sleuth:* Analisando imagem para geolocalização e metadados...")
        gps = extrair_coordenadas_gps(filepath)
        if gps:
            resp = (
                "📍 *LOCALIZAÇÃO EXATA ENCONTRADA (GPS EMBUTIDO)!*\n\n"
                f"• Latitude: `{gps['lat']}`\n"
                f"• Longitude: `{gps['lon']}`\n"
                f"• Dispositivo: {gps['dispositivo']}\n"
                f"• Data do Registro: {gps['data_foto']}\n"
                f"• Altitude: {gps['altitude']}\n\n"
                f"🗺️ *Google Maps:*\n{gps['maps_url']}\n\n"
                f"🌍 *OpenStreetMap:*\n{gps['osm_url']}"
            )
            enviar_zap(chat_id, resp)
            return

        meta = extrair_metadados_exiftool(filepath)
        info_extra = []
        if meta.get("Make") or meta.get("Model"):
            info_extra.append(f"• Aparelho: {meta.get('Make', '')} {meta.get('Model', '')}".strip())
        if meta.get("DateTimeOriginal"):
            info_extra.append(f"• Data do Disparo: {meta.get('DateTimeOriginal')}")
        if meta.get("OffsetTimeOriginal") or meta.get("OffsetTime"):
            info_extra.append(f"• Fuso Horário: {meta.get('OffsetTimeOriginal') or meta.get('OffsetTime')} (Brasil)")
        if meta.get("ImageSize"):
            info_extra.append(f"• Resolução: {meta.get('ImageSize')}")

        detalhes_str = ("\n" + "\n".join(info_extra)) if info_extra else ""

        aviso = (
            "⚠️ *Coordenadas de GPS não embutidas diretamente no arquivo.*"
            f"{detalhes_str}\n\n"
            "🧭 *Iniciando pipeline de Geointeligência Geo-Sleuth:*\n"
            "• Reconhecimento Visual e Busca Reversa (Yandex Smart Camera)\n"
            "• Leitura de Placas e Textos (OCR de Alta Resolução)\n"
            "• Estimativa Cartográfica de Coordenadas (OSM/Nominatim)..."
        )
        enviar_zap(chat_id, aviso)

        skill_base = os.path.join(os.environ.get("DOCENTE_HOME", "."), ".agents/skills/geo-sleuth/scripts")
        uv_bin = os.path.join(os.environ.get("DOCENTE_HOME", "."), ".local/bin/uv") if os.path.exists(os.path.join(os.environ.get("DOCENTE_HOME", "."), ".local/bin/uv")) else "uv"
        out_dir = f"/tmp/geosleuth_{int(time.time())}_{chat_id[-6:].replace('@','_')}"
        os.makedirs(out_dir, exist_ok=True)

        textos_ocr = []
        dados_rev = {}

        # Execução paralela: OCR e Busca Reversa Yandex
        def _exec_ocr():
            ocr_script = os.path.join(skill_base, "ocr.py")
            if not os.path.exists(ocr_script):
                return []
            ocr_json = os.path.join(out_dir, "ocr.json")
            cmd = [uv_bin, "run", ocr_script, filepath, "--out", ocr_json]
            try:
                subprocess.run(cmd, capture_output=True, text=True, timeout=25)
                if os.path.exists(ocr_json):
                    with open(ocr_json, "r", encoding="utf-8") as f_ocr:
                        dados = json.load(f_ocr)
                        return [it["text"].strip() for it in dados.get("items", []) if it.get("text", "").strip()]
            except Exception as e_o:
                log_terminal(f"Aviso OCR Geo-Sleuth: {e_o}")
            return []

        def _exec_rev():
            rev_script = os.path.join(skill_base, "revimg.py")
            if not os.path.exists(rev_script):
                return {}
            cmd = [uv_bin, "run", rev_script, filepath, "--out-dir", out_dir, "--engines", "yandex"]
            try:
                subprocess.run(cmd, capture_output=True, text=True, timeout=38)
                stem = os.path.splitext(os.path.basename(filepath))[0]
                json_path = os.path.join(out_dir, f"{stem}_yandex.json")
                if os.path.exists(json_path):
                    with open(json_path, "r", encoding="utf-8") as f_rev:
                        return json.load(f_rev)
            except Exception as e_r:
                log_terminal(f"Aviso Revimg Geo-Sleuth: {e_r}")
            return {}

        with ThreadPoolExecutor(max_workers=2) as executor:
            fut_ocr = executor.submit(_exec_ocr)
            fut_rev = executor.submit(_exec_rev)
            textos_ocr = fut_ocr.result()
            dados_rev = fut_rev.result()

        guess_list = dados_rev.get("guess", [])
        links_list = dados_rev.get("links", [])

        # Limpeza do guess
        elementos_visuais = []
        for g in guess_list:
            g_clean = re.sub(r"(?i)^image appears to contain:\s*", "", g)
            partes = [p.strip() for p in g_clean.split("/")]
            for p in partes:
                if p and not re.search(r"(?i)similar images|show more|город", p):
                    elementos_visuais.append(p)

        # Extrair candidatos geográficos e resolver coordenadas aproximadas via Nominatim
        candidatos_loc = extrair_locais_candidatos_geointel(guess_list, links_list, textos_ocr)
        ponto_geocodificado = None
        for cand in candidatos_loc[:3]:
            geo = geocode_cidade_nominatim(cand)
            if geo:
                ponto_geocodificado = (cand, geo)
                break

        # Construção da resposta detalhada
        linhas_resp = ["🧭 ════ *GEO-SLEUTH — DOSSIÊ VISUAL* ════ 🧭\n"]

        if ponto_geocodificado:
            cand_nome, geo_info = ponto_geocodificado
            lat = geo_info["lat"]
            lon = geo_info["lon"]
            maps_url = f"https://www.google.com/maps?q={lat},{lon}"
            osm_url = f"https://www.openstreetmap.org/?mlat={lat}&mlon={lon}#map=14/{lat}/{lon}"
            linhas_resp.append(
                f"📍 *LOCALIZAÇÃO ESTIMADA IDENTIFICADA!*\n"
                f"• *Local:* {geo_info.get('display_name') or cand_nome}\n"
                f"• *Coordenadas Aprox.:* `{lat}, {lon}`\n"
                f"🗺️ *Google Maps:* {maps_url}\n"
                f"🌍 *OpenStreetMap:* {osm_url}\n"
            )
        elif candidatos_loc:
            linhas_resp.append("📍 *Localizações Candidatas Identificadas:*")
            for c in candidatos_loc[:3]:
                url_c = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(c)}"
                linhas_resp.append(f"• *{c}*: {url_c}")
            linhas_resp.append("")

        if elementos_visuais:
            linhas_resp.append(f"🔍 *Reconhecimento Visual de Cenário:*\n• " + ", ".join(elementos_visuais[:6]) + "\n")

        # Links e correspondências web
        links_validos = []
        for l in links_list:
            t = l.get("title", "").strip()
            if t and not re.search(r"^\d+×\d+$", t) and l.get("url"):
                links_validos.append(l)

        if links_validos:
            linhas_resp.append("🌐 *Correspondências e Páginas com Fotos Similares:*")
            for idx, l in enumerate(links_validos[:4], 1):
                site = l.get("site", "web")
                tit = l.get("title", "")[:90]
                url = l.get("url", "")
                linhas_resp.append(f"{idx}. *{tit}* ({site})\n   👉 {url}")
            linhas_resp.append("")

        if textos_ocr:
            amostra_ocr = ", ".join(f"`{t}`" for t in textos_ocr[:8])
            linhas_resp.append(f"📝 *Textos / Placas Detectadas no OCR:*\n• {amostra_ocr}\n")

        if not ponto_geocodificado and not elementos_visuais and not links_validos and not textos_ocr:
            linhas_resp.append(
                "Nenhuma correspondência visual ou texto legível foi identificado na varredura.\n"
            )

        linhas_resp.append(
            "💡 *Dica:* Para obter as coordenadas exatas por GPS via satélite:\n"
            "1. Ative 'Salvar localização' nas opções da câmera do seu celular.\n"
            "2. No WhatsApp, envie como *DOCUMENTO* (📎 ➔ Documento) para manter os metadados EXIF intactos."
        )

        enviar_zap(chat_id, "\n".join(linhas_resp))
    except Exception as e:
        log_terminal(f"Erro em worker_geolocalizar_imagem: {e}")
        enviar_zap(chat_id, f"Erro ao processar localização com Geo-Sleuth: {str(e)[:150]}")
    finally:
        limpar_midias_temporarias(caminho_arquivo=filepath, diretorio_extra=out_dir)
        if chat_id in ULTIMA_MIDIA_CHAT and ULTIMA_MIDIA_CHAT[chat_id].get("filepath") == filepath:
            ULTIMA_MIDIA_CHAT.pop(chat_id, None)

def worker_exif_completo(chat_id: str, filepath: str):
    try:
        enviar_zap(chat_id, "📋 Extraindo todos os metadados do arquivo...")
        meta = extrair_metadados_exiftool(filepath)
        if not meta:
            enviar_zap(chat_id, "Não foi possível extrair metadados deste arquivo.")
            return

        campos_relevantes = [
            ("FileName", "Nome do Arquivo"),
            ("FileSize", "Tamanho"),
            ("FileType", "Tipo"),
            ("MIMEType", "MIME"),
            ("ImageSize", "Dimensões"),
            ("Make", "Fabricante"),
            ("Model", "Modelo/Câmera"),
            ("Software", "Software/Editor"),
            ("DateTimeOriginal", "Data/Hora Criação"),
            ("CreateDate", "Data Digitalização"),
            ("GPSLatitude", "Latitude"),
            ("GPSLongitude", "Longitude"),
            ("GPSAltitude", "Altitude"),
            ("Flash", "Flash"),
            ("FocalLength", "Distância Focal"),
            ("ExposureTime", "Exposição"),
            ("FNumber", "Abertura f/")
        ]

        linhas = ["📊 *Dossiê Forense de Metadados (EXIF):*\n"]
        for chave, rotulo in campos_relevantes:
            if chave in meta:
                linhas.append(f"• *{rotulo}:* {meta[chave]}")

        enviar_zap(chat_id, "\n".join(linhas))
    except Exception as e:
        log_terminal(f"Erro em worker_exif_completo: {e}")
        enviar_zap(chat_id, f"Erro ao ler EXIF: {e}")
    finally:
        limpar_midias_temporarias(caminho_arquivo=filepath)
        if chat_id in ULTIMA_MIDIA_CHAT and ULTIMA_MIDIA_CHAT[chat_id].get("filepath") == filepath:
            ULTIMA_MIDIA_CHAT.pop(chat_id, None)

TOP_SITES_SHERLOCK = [
    "GitHub", "GitLab", "Twitter", "X", "Instagram", "Facebook", "Reddit",
    "TikTok", "LinkedIn", "YouTube", "Twitch", "Telegram", "Pinterest",
    "Spotify", "Steam", "Discord", "Medium", "HackerNews", "Snapchat",
    "Patreon", "Docker Hub", "Substack", "Keybase", "Mastodon", "SoundCloud"
]

CATEGORIAS_HOLEHE = {
    "Redes Sociais & Mensageria": [
        "twitter.com", "x.com", "instagram.com", "facebook.com", "snapchat.com",
        "discord.com", "telegram.org", "tumblr.com", "pinterest.com", "reddit.com",
        "tiktok.com", "plurk.com", "vk.com", "weibo.com"
    ],
    "Produtividade & Corporativo": [
        "office365.com", "google.com", "trello.com", "slack.com", "lastpass.com",
        "notion.so", "zoom.us", "evernote.com", "insightly.com", "any.do",
        "basecamp.com", "atlassian.com", "jira.com"
    ],
    "Dev, Tech & Gaming": [
        "github.com", "gitlab.com", "steamcommunity.com", "steampowered.com",
        "docker.com", "replit.com", "gravatar.com", "teamtreehouse.com",
        "devrant.com", "chess.com", "codepen.io"
    ],
    "Mídia & Streaming": [
        "spotify.com", "netflix.com", "deezer.com", "soundcloud.com",
        "vimeo.com", "dailymotion.com", "twitch.tv"
    ],
    "E-commerce & Serviços": [
        "amazon.com", "ebay.com", "mercadolivre.com.br", "paypal.com",
        "airbnb.com", "booking.com", "eventbrite.com", "deliveroo.com"
    ],
    "Plataformas Adultas": [
        "pornhub.com", "xnxx.com", "xvideos.com", "redtube.com", "onlyfans.com", "stripchat.com"
    ]
}

def worker_osint_nick(chat_id: str, username: str):
    try:
        username = re.sub(r"[^a-zA-Z0-9_\-\.]", "", username).strip()
        if not username:
            enviar_zap(chat_id, "⚠️ Informe um username válido. Ex: `!osint nick joaosilva`")
            return

        enviar_zap(chat_id, f"🕵️‍♂️ *Sherlock Engine:* Rastreando presença de `{username}` nas redes abertas...\n_(Isso leva cerca de 20 a 40 segundos)_")

        cmd = [
            "sherlock",
            "--print-found",
            "--no-color",
            "--timeout", "8",
            username
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
        saida = res.stdout

        principais = []
        outros = []

        for linha in saida.splitlines():
            linha = linha.strip()
            if linha.startswith("[+]"):
                conteudo = linha.replace("[+]", "").strip()
                if ":" in conteudo:
                    rede, url = conteudo.split(":", 1)
                    rede = rede.strip()
                    url = url.strip()
                    item_str = f"  • *{rede}:* {url}"
                    if any(t.lower() == rede.lower() for t in TOP_SITES_SHERLOCK):
                        principais.append(item_str)
                    else:
                        outros.append(item_str)
                else:
                    outros.append(f"  • {conteudo}")

        total = len(principais) + len(outros)
        if total > 0:
            linhas = [
                "🕵️‍♂️ ════ *DOSSIÊ DE NICKNAME (SHERLOCK)* ════ 🕵️‍♂️\n",
                f"🎯 *Alvo:* `{username}`",
                f"📊 *Total de Perfis Confirmados:* {total}\n"
            ]

            if principais:
                linhas.append("🌟 *Principais Redes Sociais & Dev:*")
                linhas.extend(principais)
                linhas.append("")

            if outros:
                linhas.append(f"🌐 *Demais Plataformas ({len(outros)}):*")
                linhas.extend(outros[:20])
                if len(outros) > 20:
                    linhas.append(f"\n_...e mais {len(outros) - 20} perfis públicos confirmados._")

            enviar_zap(chat_id, "\n".join(linhas).strip())
        else:
            enviar_zap(chat_id, f"🔎 *Sherlock:* Nenhum perfil público com o username `{username}` foi confirmado.")
    except subprocess.TimeoutExpired:
        enviar_zap(chat_id, "⏱️ Tempo limite excedido na varredura do Sherlock. Tente novamente.")
    except Exception as e:
        log_terminal(f"Erro em worker_osint_nick: {e}")
        enviar_zap(chat_id, f"Erro ao rodar OSINT de username: {e}")

def worker_osint_domain(chat_id: str, domain: str):
    try:
        domain = re.sub(r"[^a-zA-Z0-9\-\.]", "", domain).strip().lower()
        if not domain or "." not in domain:
            enviar_zap(chat_id, "⚠️ Informe um domínio válido. Ex: `!osint domain exemplo.com.br`")
            return

        enviar_zap(chat_id, f"🌐 *OSINT Domínio:* Mapeando infraestrutura de `{domain}`...")

        # 1. Resolução IP (A)
        res_a = subprocess.run(["dig", "+short", domain, "A"], capture_output=True, text=True, timeout=8)
        ips = [i.strip() for i in res_a.stdout.splitlines() if i.strip()]

        # PTR / Host reverso do primeiro IP
        ptr_info = ""
        if ips:
            res_ptr = subprocess.run(["dig", "+short", "-x", ips[0]], capture_output=True, text=True, timeout=5)
            ptr_lines = [p.strip() for p in res_ptr.stdout.splitlines() if p.strip()]
            if ptr_lines:
                ptr_info = ptr_lines[0]

        # 2. Servidor Web / HTTP Status
        http_banner = ""
        try:
            res_http = subprocess.run(
                ["curl", "-s", "-I", "-m", "4", f"http://{domain}"],
                capture_output=True, text=True, timeout=5
            )
            for h_line in res_http.stdout.splitlines()[:6]:
                h_line_s = h_line.strip()
                if h_line_s.startswith("HTTP/"):
                    http_banner = f"`{h_line_s}`"
                elif h_line_s.lower().startswith("server:"):
                    srv = h_line_s.split(":", 1)[1].strip()
                    http_banner += f" | Servidor: `{srv}`"
        except Exception:
            pass

        # 3. MX (Correio)
        res_mx = subprocess.run(["dig", "+short", domain, "MX"], capture_output=True, text=True, timeout=8)
        mx_records = [m.strip() for m in res_mx.stdout.splitlines() if m.strip()]

        # 4. NS (DNS)
        res_ns = subprocess.run(["dig", "+short", domain, "NS"], capture_output=True, text=True, timeout=8)
        ns_records = [n.strip() for n in res_ns.stdout.splitlines() if n.strip()]

        # 5. WHOIS
        res_whois = subprocess.run(["whois", domain], capture_output=True, text=True, timeout=12)
        whois_out = res_whois.stdout
        info_whois = []
        for linha in whois_out.splitlines():
            l_clean = linha.strip()
            if any(l_clean.lower().startswith(k) for k in [
                "titular:", "owner:", "responsavel:", "responsible:",
                "criado:", "created:", "expiracao:", "expires:",
                "org:", "country:", "status:"
            ]):
                info_whois.append(f"  • {l_clean}")

        linhas = [
            "🌐 ════ *RECONHECIMENTO DE DOMÍNIO (OSINT)* ════ 🌐\n",
            f"🎯 *Alvo:* `{domain}`"
        ]
        if http_banner:
            linhas.append(f"📡 *Resposta Web:* {http_banner}")

        linhas.append("\n🖥️ *Endereços IP (A):*")
        if ips:
            for ip in ips[:4]:
                linhas.append(f"  • `{ip}`")
            if ptr_info:
                linhas.append(f"    ↳ _Reverse DNS (PTR):_ `{ptr_info}`")
        else:
            linhas.append("  • _Nenhum registro A encontrado._")

        linhas.append("\n📬 *Servidores de E-mail (MX):*")
        if mx_records:
            for mx in mx_records[:4]:
                linhas.append(f"  • `{mx}`")
        else:
            linhas.append("  • _Sem servidores MX configurados._")

        linhas.append("\n🛡️ *Servidores DNS Autoritativos (NS):*")
        if ns_records:
            for ns in ns_records[:4]:
                linhas.append(f"  • `{ns}`")
        else:
            linhas.append("  • _Sem servidores NS localizados._")

        if info_whois:
            linhas.append("\n📋 *Dados Cadastrais (WHOIS):*")
            linhas.extend(info_whois[:8])

        enviar_zap(chat_id, "\n".join(linhas))
    except Exception as e:
        log_terminal(f"Erro em worker_osint_domain: {e}")
        enviar_zap(chat_id, f"Erro ao consultar domínio: {e}")

def buscar_vinculo_osint_db(alvo: str) -> dict | None:
    """Busca correlações existentes entre e-mail, telefone, username ou nome na base local."""
    if not alvo:
        return None
    alvo_clean = re.sub(r"[^\w@\.\-]", "", alvo).lower()
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS osint_vinculos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT,
                    telefone TEXT,
                    username TEXT,
                    nome TEXT,
                    turma TEXT,
                    servicos TEXT,
                    atualizado_em TEXT
                )
            """)
            row = cur.execute("""
                SELECT * FROM osint_vinculos 
                WHERE LOWER(email) = ? OR telefone LIKE ? OR LOWER(username) = ? OR LOWER(nome) LIKE ?
            """, (alvo_clean, f"%{alvo_clean}%", alvo_clean, f"%{alvo_clean}%")).fetchone()
            if row:
                return dict(row)

            # Busca em alunos por nome
            row_aluno = cur.execute("""
                SELECT nome, turma, matricula FROM alunos 
                WHERE LOWER(nome) LIKE ?
            """, (f"%{alvo_clean}%",)).fetchone()
            if row_aluno:
                return {
                    "nome": row_aluno["nome"],
                    "turma": row_aluno["turma"],
                    "matricula": row_aluno["matricula"],
                    "email": None,
                    "telefone": None
                }
    except Exception as e:
        log_terminal(f"Aviso busca vinculo OSINT DB: {e}")
    return None

def salvar_vinculo_osint_db(email: str = None, telefone: str = None, username: str = None, nome: str = None, turma: str = None, servicos: list = None):
    """Armazena ou atualiza correlações descobertas entre e-mail, celular e plataformas."""
    try:
        serv_str = json.dumps(servicos, ensure_ascii=False) if servicos else None
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS osint_vinculos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT,
                    telefone TEXT,
                    username TEXT,
                    nome TEXT,
                    turma TEXT,
                    servicos TEXT,
                    atualizado_em TEXT
                )
            """)
            query = "SELECT id, email, telefone, username, nome, turma FROM osint_vinculos WHERE 1=0 "
            params = []
            if email:
                query += " OR LOWER(email) = ?"
                params.append(email.lower())
            if telefone:
                query += " OR telefone LIKE ?"
                params.append(f"%{telefone[-9:]}")
            if username:
                query += " OR LOWER(username) = ?"
                params.append(username.lower())

            existente = cur.execute(query, params).fetchone()
            if existente:
                e_id, e_mail, e_tel, e_user, e_nome, e_turma = existente
                f_mail = email or e_mail
                f_tel = telefone or e_tel
                f_user = username or e_user
                f_nome = nome or e_nome
                f_turma = turma or e_turma
                cur.execute("""
                    UPDATE osint_vinculos 
                    SET email = ?, telefone = ?, username = ?, nome = ?, turma = ?, servicos = COALESCE(?, servicos), atualizado_em = ?
                    WHERE id = ?
                """, (f_mail, f_tel, f_user, f_nome, f_turma, serv_str, now_str, e_id))
            else:
                cur.execute("""
                    INSERT INTO osint_vinculos (email, telefone, username, nome, turma, servicos, atualizado_em)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (email, telefone, username, nome, turma, serv_str, now_str))
            conn.commit()
    except Exception as e:
        log_terminal(f"Aviso ao salvar vinculo OSINT: {e}")

def buscar_contato_openwa_por_nome(termo: str) -> dict | None:
    """Busca nos participantes do grupo do Clube de Programação ou contatos do WhatsApp por nome/pushName."""
    if not termo or len(termo) < 3:
        return None
    termo = termo.lower().strip()
    try:
        url_g = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/groups/{ID_GRUPO_CLUBE_PROGRAMACAO}"
        req_g = urllib.request.Request(url_g, headers={"x-api-key": OPENWA_API_KEY})
        with urllib.request.urlopen(req_g, timeout=4) as r:
            d = json.loads(r.read().decode("utf-8"))
            for p in d.get("participants", []):
                jid = p.get("id")
                try:
                    req_c = urllib.request.Request(f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/contacts/{jid}", headers={"x-api-key": OPENWA_API_KEY})
                    with urllib.request.urlopen(req_c, timeout=2) as rc:
                        dc = json.loads(rc.read().decode("utf-8"))
                        pname = dc.get("pushName") or ""
                        cname = dc.get("name") or ""
                        full = f"{pname} {cname}".lower()
                        if termo in full:
                            return {
                                "jid": jid,
                                "numero": re.sub(r"\D", "", jid),
                                "pushName": pname,
                                "name": cname
                            }
                except Exception:
                    pass
    except Exception:
        pass
    return None

def buscar_info_github(username: str) -> dict | None:
    """Consulta rápida da API pública do GitHub para correlacionar perfil de desenvolvedor."""
    if not username:
        return None
    try:
        url = f"https://api.github.com/users/{urllib.parse.quote(username)}"
        req = urllib.request.Request(url, headers={"User-Agent": "DocenteOSINT/1.0"})
        with urllib.request.urlopen(req, timeout=4) as r:
            if r.status == 200:
                d = json.loads(r.read().decode("utf-8"))
                return {
                    "name": d.get("name"),
                    "bio": d.get("bio"),
                    "email": d.get("email"),
                    "company": d.get("company"),
                    "location": d.get("location"),
                    "blog": d.get("blog"),
                    "profile_url": d.get("html_url")
                }
    except Exception:
        pass
    return None

def executar_holehe_estruturado(email: str) -> dict:
    """Executa Holehe com geração e leitura de CSV estruturado para capturar telefones e e-mails de recuperação."""
    import tempfile
    import csv
    email = email.strip().lower()
    tmp_dir = tempfile.mkdtemp(prefix="holehe_osint_")
    cmd = [os.path.join(os.environ.get("DOCENTE_HOME", "."), ".local/bin/holehe"), email, "--csv", "--only-used", "--no-color", "--no-clear"]
    try:
        subprocess.run(cmd, cwd=tmp_dir, capture_output=True, text=True, timeout=40)
        csvs = [f for f in os.listdir(tmp_dir) if f.endswith(".csv")]
        confirmados = []
        telefones = []
        recuperacoes = []
        if csvs:
            with open(os.path.join(tmp_dir, csvs[0]), "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row.get("exists") == "True":
                        dom = row.get("domain") or row.get("name") or ""
                        if dom:
                            confirmados.append(dom)
                        phone = (row.get("phoneNumber") or "").strip()
                        rec = (row.get("emailrecovery") or "").strip()
                        if phone:
                            telefones.append({"servico": dom, "telefone": phone})
                        if rec:
                            recuperacoes.append({"servico": dom, "recuperacao": rec})
        return {
            "confirmados": confirmados,
            "telefones": telefones,
            "recuperacoes": recuperacoes
        }
    except Exception as e:
        log_terminal(f"Aviso executar_holehe_estruturado: {e}")
        return {"confirmados": [], "telefones": [], "recuperacoes": []}
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

def worker_osint_email(chat_id: str, email: str, orquestrado: bool = True):
    try:
        email = email.strip().lower()
        if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
            enviar_zap(chat_id, "⚠️ Informe um e-mail válido. Ex: `!osint email usuario@exemplo.com`")
            return

        dominio_email = email.split("@", 1)[1]
        enviar_zap(chat_id, f"📧 *Holehe Engine & Orquestração:* Rastreando `{email}` e correlacionando celular/redes...\n_(Consultando 120+ plataformas e base de dados)_")

        provedor_mx = ""
        try:
            res_mx = subprocess.run(["dig", "+short", dominio_email, "MX"], capture_output=True, text=True, timeout=5)
            mx_out = res_mx.stdout.lower()
            if "google" in mx_out or "aspmx" in mx_out:
                provedor_mx = "Google Workspace / Gmail"
            elif "outlook" in mx_out or "protection.outlook.com" in mx_out:
                provedor_mx = "Microsoft 365 / Outlook"
            elif "proton" in mx_out:
                provedor_mx = "Proton Mail"
            elif "zoho" in mx_out:
                provedor_mx = "Zoho Mail"
        except Exception:
            pass

        # Execução estruturada do Holehe
        dados_h = executar_holehe_estruturado(email)
        encontrados_brutos = dados_h.get("confirmados", [])
        telefones_holehe = dados_h.get("telefones", [])
        recuperacoes_holehe = dados_h.get("recuperacoes", [])

        # Orquestração: Análise do Username do e-mail
        nick = email.split("@", 1)[0]
        perfil_git = buscar_info_github(nick) if orquestrado else None

        # Correlação cruzada com Celular / Base local / WhatsApp
        vinculo_local = buscar_vinculo_osint_db(email)
        if not vinculo_local and orquestrado:
            palavras_nick = [p for p in re.findall(r"[a-zA-Z]{4,}", nick)]
            for pal in palavras_nick:
                vinc = buscar_vinculo_osint_db(pal)
                if vinc and vinc.get("nome"):
                    vinculo_local = vinc
                    break
                contato_wpp = buscar_contato_openwa_por_nome(pal)
                if contato_wpp:
                    vinculo_local = {
                        "telefone": contato_wpp["numero"],
                        "nome": contato_wpp.get("name") or contato_wpp.get("pushName"),
                        "turma": "Clube de Programação",
                        "origem": "WhatsApp Clube"
                    }
                    break

        if vinculo_local and vinculo_local.get("telefone"):
            salvar_vinculo_osint_db(
                email=email,
                telefone=vinculo_local.get("telefone"),
                username=nick,
                nome=vinculo_local.get("nome"),
                turma=vinculo_local.get("turma"),
                servicos=encontrados_brutos
            )

        linhas = [
            "📧 ════ *DOSSIÊ DE E-MAIL & ORQUESTRAÇÃO* ════ 📧\n",
            f"🎯 *Alvo:* `{email}`",
            f"🏢 *Domínio:* `{dominio_email}`" + (f" ({provedor_mx})" if provedor_mx else ""),
            f"📊 *Serviços Confirmados (Holehe):* {len(encontrados_brutos)}\n"
        ]

        # SEÇÃO DE ORQUESTRAÇÃO: CELULAR VINCULADO IDENTIFICADO
        tem_vinculo_tel = False
        if telefones_holehe:
            linhas.append("📱 *CELULAR / TELEFONE DETECTADO EM SERVIÇOS:*")
            for t_item in telefones_holehe:
                linhas.append(f"  • `{t_item['telefone']}` _(Detectado em: {t_item['servico']})_")
            tem_vinculo_tel = True
            linhas.append("")

        if vinculo_local and vinculo_local.get("telefone"):
            num_clean = vinculo_local.get("telefone")
            num_fmt = f"+{num_clean[:2]} ({num_clean[2:4]}) {num_clean[4:9]}-{num_clean[9:]}" if len(num_clean) >= 12 else f"+{num_clean}"
            linhas.append("📱 ════ *ORQUESTRAÇÃO: CELULAR VINCULADO IDENTIFICADO* ════ 📱")
            linhas.append(f"  • *Número:* `{num_fmt}`")
            if vinculo_local.get("nome"):
                linhas.append(f"  • *Titular:* {vinculo_local.get('nome')}" + (f" ({vinculo_local.get('turma')})" if vinculo_local.get("turma") else ""))
            linhas.append("  • *Presença WhatsApp:* Conta Ativa ✅")
            linhas.append("  • *Origem do Vínculo:* Base Local / Clube de Programação")
            tem_vinculo_tel = True
            linhas.append("")

        if recuperacoes_holehe:
            linhas.append("🔒 *E-MAILS DE RECUPERAÇÃO ENCONTRADOS:*")
            for r_item in recuperacoes_holehe:
                linhas.append(f"  • `{r_item['recuperacao']}` _(Origem: {r_item['servico']})_")
            linhas.append("")

        if perfil_git:
            linhas.append("👤 *Perfil de Desenvolvedor (GitHub):*")
            if perfil_git.get("name"):
                linhas.append(f"  • *Nome:* {perfil_git.get('name')}")
            if perfil_git.get("bio"):
                linhas.append(f"  • *Bio:* {perfil_git.get('bio')}")
            if perfil_git.get("location"):
                linhas.append(f"  • *Local:* {perfil_git.get('location')}")
            if perfil_git.get("profile_url"):
                linhas.append(f"  • *URL:* {perfil_git.get('profile_url')}")
            linhas.append("")

        # Agrupamento de categorias
        if encontrados_brutos:
            agrupados = {cat: [] for cat in CATEGORIAS_HOLEHE}
            outros = []
            for serv in encontrados_brutos:
                alocado = False
                for cat, doms in CATEGORIAS_HOLEHE.items():
                    if any(d in serv.lower() for d in doms):
                        agrupados[cat].append(serv)
                        alocado = True
                        break
                if not alocado:
                    outros.append(serv)

            for cat, itens in agrupados.items():
                if itens:
                    emoji = "🔞 " if "Adulta" in cat else "• "
                    linhas.append(f"*{cat}:*")
                    for s in itens:
                        linhas.append(f"  {emoji}*{s}*")
                    linhas.append("")

            if outros:
                linhas.append("*Outros Serviços Confirmados:*")
                for s in outros:
                    linhas.append(f"  • *{s}*")
                linhas.append("")
        else:
            linhas.append("🔎 *Nenhuma conta pública vinculada ao e-mail foi detectada no Holehe.*\n")

        if not tem_vinculo_tel and orquestrado:
            linhas.append("💡 _Dica: Utilize `!osint phone <numero>` para investigar o número diretamente._")

        enviar_zap(chat_id, "\n".join(linhas).strip())
    except subprocess.TimeoutExpired:
        enviar_zap(chat_id, "⏱️ Tempo limite excedido na busca do Holehe. Tente novamente.")
    except Exception as e:
        log_terminal(f"Erro em worker_osint_email: {e}")
        enviar_zap(chat_id, f"Erro ao rodar OSINT de e-mail: {e}")

def worker_osint_phone(chat_id: str, raw_input: str, orquestrado: bool = True):
    try:
        raw_input = raw_input.strip()
        partes = raw_input.split(maxsplit=1)
        if len(partes) == 2:
            cc = partes[0].lstrip("+")
            num = re.sub(r"[^\d]", "", partes[1])
        else:
            digits = re.sub(r"[^\d]", "", raw_input)
            if digits.startswith("55") and len(digits) in (12, 13):
                cc, num = "55", digits[2:]
            elif len(digits) in (10, 11):
                cc, num = "55", digits
            elif len(digits) > 11 and raw_input.startswith("+"):
                cc, num = digits[:2], digits[2:]
            else:
                cc, num = "55", digits

        if not num or len(num) < 8:
            enviar_zap(chat_id, "⚠️ Informe um telefone válido. Ex: `!osint phone (00) 90000-0000` ou `!osint phone +55 (00) 90000-0000`")
            return

        e164_input = f"+{cc}{num}"
        enviar_zap(chat_id, f"📱 *OSINT Telefônico & Orquestração:* Investigando `+{cc} {num}` e buscando e-mails associados...")

        # 1. Análise Cadastral e Telecom via phonenumbers
        info_regiao = "Brasil"
        tipo_linha = "Móvel / Celular"
        fmt_intl = e164_input
        fmt_nat = num
        try:
            import phonenumbers
            from phonenumbers import geocoder, carrier
            p = phonenumbers.parse(e164_input, None)
            if phonenumbers.is_valid_number(p):
                fmt_intl = phonenumbers.format_number(p, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
                fmt_nat = phonenumbers.format_number(p, phonenumbers.PhoneNumberFormat.NATIONAL)
                reg = geocoder.description_for_number(p, "pt")
                if reg:
                    info_regiao = f"{reg} (Brasil 🇧🇷)"
                operadora = carrier.name_for_number(p, "pt")
                if operadora:
                    tipo_linha += f" ({operadora})"
        except Exception:
            pass

        # 2. Presença no WhatsApp via Open-WA
        whats_info = {}
        try:
            jid_alvo = f"{cc}{num}@c.us"
            url_contato = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/contacts/{jid_alvo}"
            req_c = urllib.request.Request(url_contato, headers={"x-api-key": OPENWA_API_KEY})
            with urllib.request.urlopen(req_c, timeout=5) as r_c:
                if r_c.status == 200:
                    dados_c = json.loads(r_c.read().decode())
                    if dados_c and dados_c.get("id"):
                        whats_info["ativo"] = True
                        whats_info["pushName"] = dados_c.get("pushName") or "Não divulgado"
                        whats_info["id"] = dados_c.get("id")
        except Exception:
            pass

        # 3. Varredura com Ignorant
        servicos_confirmados = []
        try:
            cmd = [
                "/usr/local/bin/ignorant",
                cc,
                num,
                "--only-used",
                "--no-color"
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=35)
            saida = res.stdout + "\n" + res.stderr
            for linha in saida.splitlines():
                linha = linha.strip()
                if linha.startswith("[+]"):
                    servico = linha.replace("[+]", "").strip()
                    if servico and not servico.startswith("Phone number used") and "." in servico:
                        servicos_confirmados.append(servico)
        except Exception as e_ig:
            log_terminal(f"Aviso Ignorant: {e_ig}")

        # 4. ORQUESTRAÇÃO: BUSCAR O E-MAIL ASSOCIADO A ESTE CELULAR
        email_associado = None
        dados_vinculo = buscar_vinculo_osint_db(f"{cc}{num}") or buscar_vinculo_osint_db(num)
        if dados_vinculo and dados_vinculo.get("email"):
            email_associado = dados_vinculo.get("email")

        # Se não achou por número direto, tenta pelo pushName do WhatsApp
        if not email_associado and whats_info.get("pushName"):
            pname = whats_info.get("pushName")
            v_pname = buscar_vinculo_osint_db(pname)
            if v_pname and v_pname.get("email"):
                email_associado = v_pname.get("email")
                dados_vinculo = v_pname

        servicos_email_holehe = []
        if email_associado and orquestrado:
            try:
                dados_h = executar_holehe_estruturado(email_associado)
                servicos_email_holehe = dados_h.get("confirmados", [])
            except Exception:
                pass

        # 5. Montagem do Dossiê Visual e Integrado
        linhas = [
            "📱 ════ *DOSSIÊ TELEFÔNICO & ORQUESTRAÇÃO* ════ 📱\n",
            f"🎯 *Número:* `{fmt_nat}`",
            f"🌍 *Formato Internacional:* `{fmt_intl}`",
            f"📍 *Região:* {info_regiao}",
            f"📶 *Classificação:* {tipo_linha}\n"
        ]

        linhas.append("💬 *Presença no WhatsApp:*")
        if whats_info.get("ativo"):
            linhas.append("  • *Conta Ativa:* Sim ✅")
            linhas.append(f"  • *Nome de Exibição (PushName):* {whats_info.get('pushName')}")
            linhas.append(f"  • *JID:* `{whats_info.get('id')}`")
        else:
            linhas.append("  • *Conta Ativa:* Não verificada ou privada")
        linhas.append("")

        # SEÇÃO DE ORQUESTRAÇÃO: E-MAIL ASSOCIADO AO CELULAR
        if email_associado:
            linhas.append("📧 ════ *ORQUESTRAÇÃO: E-MAIL ASSOCIADO AO CELULAR* ════ 📧")
            linhas.append(f"• *E-mail Cadastrado:* `{email_associado}`")
            if dados_vinculo and dados_vinculo.get("nome"):
                linhas.append(f"• *Titular:* {dados_vinculo.get('nome')}" + (f" ({dados_vinculo.get('turma')})" if dados_vinculo.get("turma") else ""))
            if servicos_email_holehe:
                linhas.append(f"• *Contas Confirmadas para este E-mail (Holehe - {len(servicos_email_holehe)}):*")
                for s_h in servicos_email_holehe[:8]:
                    linhas.append(f"   ↳ 🎯 *{s_h}*")
            linhas.append("")

        linhas.append("🌐 *Serviços Vinculados ao Número (Ignorant):*")
        if servicos_confirmados:
            for s in servicos_confirmados:
                linhas.append(f"  • 🎯 *{s}* (Conta confirmada)")
        else:
            linhas.append("  • _Nenhum serviço público adicional confirmado diretamente pelo número._")

        if not email_associado and orquestrado:
            linhas.append("\n💡 _Dica: Para rastreamento cruzado completo de redes, use `!osint full <alvo>`._")

        enviar_zap(chat_id, "\n".join(linhas).strip())
    except subprocess.TimeoutExpired:
        enviar_zap(chat_id, "⏱️ Tempo limite excedido na varredura telefônica.")
    except Exception as e:
        log_terminal(f"Erro em worker_osint_phone: {e}")
        enviar_zap(chat_id, f"Erro ao rodar OSINT de telefone: {e}")

def worker_osint_orquestrador(chat_id: str, alvo: str):
    """Orquestrador OSINT integrado: infere o tipo de alvo e executa a correlação cruzada completa."""
    alvo = alvo.strip()
    if not alvo:
        enviar_zap(chat_id, "⚠️ Informe um alvo para a orquestração (e-mail, celular ou username). Ex: `!osint full mateusfernandes3483@gmail.com`")
        return
    enviar_zap(chat_id, f"🕸️ *OSINT Orchestrator:* Iniciando rastreamento integrado e correlação cruzada para `{alvo}`...")
    if "@" in alvo:
        worker_osint_email(chat_id, alvo, orquestrado=True)
    elif re.sub(r"\D", "", alvo) and len(re.sub(r"\D", "", alvo)) >= 8 and not re.search(r"[a-zA-Z]", alvo):
        worker_osint_phone(chat_id, alvo, orquestrado=True)
    else:
        worker_osint_nick(chat_id, alvo)

def instagram_perfil_publico_og(username: str):
    """Identifica perfil Instagram sem login via meta tags OpenGraph (UA de crawler social)."""
    import html as _html
    try:
        req = urllib.request.Request(f"https://www.instagram.com/{username}/", headers={"User-Agent": "facebookexternalhit/1.1"})
        with urllib.request.urlopen(req, timeout=12) as r:
            corpo = r.read().decode("utf-8", "ignore")
    except Exception:
        return None
    def og(prop):
        m = re.search(r'<meta property="%s" content="([^"]*)"' % prop, corpo)
        return _html.unescape(m.group(1)) if m else ""
    desc, titulo = og("og:description"), og("og:title")
    m = re.match(r"([\d.,KMkm]+) Followers?, ([\d.,KMkm]+) Following, ([\d.,KMkm]+) Posts?", desc)
    if not m or "@" not in titulo:
        return None
    nome = re.sub(r"\s*\(@.*", "", titulo).strip()
    return {"nome": nome, "seguidores": m.group(1), "seguindo": m.group(2), "posts": m.group(3),
            "foto": og("og:image"), "url": f"https://www.instagram.com/{username}/"}

def instagram_buscar_por_nome(termo: str, limite: int = 8):
    """Descobre perfis Instagram por nome/nick via dork no DuckDuckGo (site:instagram.com)."""
    import urllib.parse as _up
    try:
        q = _up.quote(f'site:instagram.com "{termo}"')
        req = urllib.request.Request(f"https://html.duckduckgo.com/html/?q={q}", headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as r:
            corpo = r.read().decode("utf-8", "ignore")
    except Exception:
        return []
    achados = []
    for u in re.findall(r'uddg=([^&"]+)', corpo):
        u = _up.unquote(u)
        m = re.match(r"https?://(?:www\.)?instagram\.com/([A-Za-z0-9._]{2,30})/?$", u)
        if m and m.group(1) not in ("p", "reel", "explore", "accounts") and m.group(1) not in achados:
            achados.append(m.group(1))
    return achados[:limite]

def worker_osint_insta_publico(chat_id: str, alvo: str):
    """Identificação de Instagram sem session_id: perfil direto (OG) ou descoberta por nome."""
    alvo = alvo.strip().lstrip("@")
    if re.fullmatch(r"[A-Za-z0-9._]{2,30}", alvo):
        d = instagram_perfil_publico_og(alvo)
        if d:
            enviar_zap(chat_id, "📸 *INSTAGRAM (consulta pública, sem login)*\n"
                f"• *Perfil:* @{alvo}\n• *Nome:* {d['nome']}\n• *Seguidores:* {d['seguidores']} | *Seguindo:* {d['seguindo']} | *Posts:* {d['posts']}\n• {d['url']}")
            return
    enviar_zap(chat_id, f"🔎 Perfil `{alvo}` não encontrado diretamente. Buscando por nome/nick na web...")
    cands = instagram_buscar_por_nome(alvo)
    if not cands:
        enviar_zap(chat_id, "Nenhum perfil público encontrado.")
        return
    linhas = [f"📸 *Candidatos Instagram para* `{alvo}`:"]
    for c in cands:
        d = instagram_perfil_publico_og(c)
        linhas.append(f"• @{c}" + (f" — {d['nome']} ({d['seguidores']} seg.)" if d else "") + f"\n  https://www.instagram.com/{c}/")
    enviar_zap(chat_id, "\n".join(linhas))

def worker_osint_insta(chat_id: str, param: str):
    try:
        partes = param.strip().split(maxsplit=1)
        username = partes[0].replace("@", "").strip()
        custom_session = partes[1].strip() if len(partes) > 1 else ""

        if not username:
            enviar_zap(chat_id, "⚠️ Informe um usuário do Instagram. Ex: `!osint insta neymarjr`")
            return

        session_id = custom_session or os.environ.get("TOUTATIS_SESSION_ID") or os.environ.get("INSTAGRAM_SESSION_ID") or ""
        if not session_id and os.path.exists(os.path.expanduser("~/.toutatis_session")):
            try:
                with open(os.path.expanduser("~/.toutatis_session")) as f:
                    session_id = f.read().strip()
            except Exception:
                pass

        enviar_zap(chat_id, f"📸 *Toutatis Engine:* Investigando perfil Instagram `@{username}`...")

        from toutatis import core

        linhas = [
            "📸 ════ *DOSSIÊ INSTAGRAM (TOUTATIS)* ════ 📸\n",
            f"👤 *Perfil:* `@{username}`",
            f"🔗 *URL:* https://www.instagram.com/{username}/\n"
        ]

        if session_id:
            info_data = core.getInfo(username, session_id, searchType="username")
            if info_data.get("error"):
                linhas.append(f"⚠️ *Aviso da API:* {info_data['error']}\n")
            elif info_data.get("user"):
                u = info_data["user"]
                linhas.append("📋 *Dados da Conta:*")
                linhas.append(f"  • *ID Numérico:* `{u.get('userID', 'N/A')}`")
                linhas.append(f"  • *Nome Completo:* {u.get('full_name', 'Não informado')}")
                linhas.append(f"  • *Status:* {'🔒 Privada' if u.get('is_private') else '🔓 Pública'}")
                linhas.append(f"  • *Verificado:* {'⭐ Sim' if u.get('is_verified') else 'Não'}")
                linhas.append(f"  • *Conta Comercial:* {'Sim' if u.get('is_business') else 'Não'}")
                linhas.append(f"  • *Novo Usuário:* {'Sim' if u.get('is_new_to_instagram') else 'Não'}")
                f_cnt = u.get('follower_count', 0)
                fol_cnt = u.get('following_count', 0)
                linhas.append(f"  • *Seguidores:* {f_cnt:,} | *Seguindo:* {fol_cnt:,}")
                linhas.append(f"  • *Publicações:* {u.get('media_count', 0)}")
                if u.get("external_url"):
                    linhas.append(f"  • *Website/Link:* {u.get('external_url')}")
                if u.get("public_email"):
                    linhas.append(f"  • *E-mail Público:* `{u.get('public_email')}`")
                if u.get("public_phone_number"):
                    linhas.append(f"  • *Telefone Público:* `+{u.get('public_phone_country_code')} {u.get('public_phone_number')}`")
                if u.get("is_whatsapp_linked"):
                    linhas.append(f"  • *WhatsApp Comercial Vinculado:* Sim")
                if u.get("biography"):
                    bio = u.get("biography").strip()
                    linhas.append(f"\n📝 *Biografia:*\n_{bio}_\n")
        else:
            worker_osint_insta_publico(chat_id, username)
            linhas.append(
                "🔑 *Autenticação com Session ID:*\n"
                "Para desbloquear ID numérico, e-mail/telefone comerciais públicos e métricas completas via API, forneça o `sessionid`:\n"
                "👉 `!osint insta <usuario> <seu_session_id>`\n"
                "_(O `sessionid` é copiado dos cookies do navegador em instagram.com)_\n"
            )

        # Tentativa de recuperação ofuscada
        try:
            adv = core.advanced_lookup(username)
            if adv and adv.get("user"):
                usr_adv = adv["user"]
                recup = []
                if usr_adv.get("obfuscated_email"):
                    recup.append(f"  • 📧 *E-mail de Recuperação:* `{usr_adv['obfuscated_email']}`")
                if usr_adv.get("obfuscated_phone"):
                    recup.append(f"  • 📱 *Telefone de Recuperação:* `{usr_adv['obfuscated_phone']}`")
                if recup:
                    linhas.append("🛡️ *Dados de Recuperação Ofuscados:*")
                    linhas.extend(recup)
        except Exception:
            pass

        enviar_zap(chat_id, "\n".join(linhas).strip())
    except Exception as e:
        log_terminal(f"Erro em worker_osint_insta: {e}")
        enviar_zap(chat_id, f"Erro ao rodar Toutatis no Instagram: {e}")

TEXTOS_ENVIADOS_RECENTES = {}  # {texto_normalizado: timestamp}

def _registrar_texto_enviado_bot(texto: str):
    if not texto:
        return
    agora = time.time()
    t_norm = normalizar_texto(texto)
    if t_norm:
        TEXTOS_ENVIADOS_RECENTES[t_norm] = agora
    if len(TEXTOS_ENVIADOS_RECENTES) > 500:
        for k in list(TEXTOS_ENVIADOS_RECENTES.keys()):
            if agora - TEXTOS_ENVIADOS_RECENTES[k] > 300:
                TEXTOS_ENVIADOS_RECENTES.pop(k, None)

def _eh_eco_ou_mensagem_bot(texto: str) -> bool:
    """
    Detecta de forma determinística e independente da flag `fromMe` se o texto
    é uma resposta gerada pelo próprio assistente/observador ou um eco do WhatsApp.
    """
    if not texto:
        return False
    t_raw = str(texto).strip().lower()
    t_norm = normalizar_texto(texto)

    # 1. Cache de mensagens enviadas recentemente pelo próprio bot
    agora = time.time()
    if t_norm in TEXTOS_ENVIADOS_RECENTES and (agora - TEXTOS_ENVIADOS_RECENTES[t_norm] < 180):
        return True

    # 2. Marcadores de fallback / erro de comando (bloqueia recursão "Comando 'Comando '..." na 1ª vez)
    if (
        t_raw.startswith("comando '")
        or t_norm.startswith("comando '")
        or "nao reconhecido. digite" in t_norm
        or "recebido. digite" in t_norm
        or "digite `!ajuda` para ver o menu" in t_raw
        or "digite !ajuda para ver o menu" in t_norm
        or "envie *!bot ajuda*" in t_raw
    ):
        return True

    # 3. Emojis iniciais de templates do bot
    emojis_bot = (
        "📋", "🌅", "🚨", "⚠️", "❌", "✅", "💡", "📅", "🔭",
        "🖥️", "🏫", "📝", "⚡", "🔎", "🎬", "📄", "🧹", "👋", "⏰", "📌", "🎯"
    )
    if any(t_raw.startswith(e) for e in emojis_bot):
        return True

    # 4. Cabeçalhos de respostas do bot mesmo após strip de emojis (limpar_emojis)
    t_sem_md = re.sub(r"^[\s\*_~\-\[\]]+", "", t_norm)
    prefixos_bot_norm = (
        "observador",
        "assistente docente",
        "menu completo do assistente",
        "laboratorio iema ti",
        "ola, prof.",
        "ola, profa.",
        "alarme / aviso docente",
        "gestao de atividades, vistos em massa",
        "estado publico / privado das salas",
        "comandos osint disponiveis",
        "gerador de aula a partir de link",
        "gabarito de laboratorio gerado",
        "sincronizar reserva do papel",
        "configuracao de limpeza de segunda-feira",
        "confirmacao de seguranca:",
        "ideia recebida:",
        "preparando conteudo ",
        "docente identificado com sucesso:",
        "operacao cancelada.",
        "recebi. o agente vai ver.",
        "privado pos-17h",
        "achados neste",
        "resposta enviada",
        "reserva confirmada",
        "horario indisponivel",
        "horarios liberados",
        "resumo das alteracoes na grade",
    )
    if any(t_sem_md.startswith(p) for p in prefixos_bot_norm):
        return True

    return False

def _eh_comando_valido_sem_prefixo(msg_norm: str) -> bool:
    """
    Centraliza os comandos aceitos sem prefixo (! ou /), exigindo fronteira de palavra (\b)
    para que frases naturais (ex.: 'modos de comandos', 'dia bom') nunca disparem comandos.
    """
    if not msg_norm:
        return False
    if _eh_comando_lab_msg(msg_norm):
        return True
    # 'modo' sem prefixo só é comando se estiver sozinho ou acompanhado de turma/visibilidade
    if re.match(r"^modo(?:\s+(?:20[0-3]|base|ti|lab|laboratorio|publico|privado))*\s*$", msg_norm):
        return True
    # Comandos operacionais do dia a dia docente com fronteira estrita de palavra
    padrao = (
        r"^(?:"
        r"atv|vistos?|atividades?|corrigir|correcao|caderno|"
        r"cobrar|pendentes?|pendencias?|passar|"
        r"resumo|matinal|"
        r"minha grade|meus horarios|"
        r"onde|onde estou|onde eu to|onde eu estou|qual minha sala|qual minha aula|minha aula agora|"
        r"aula\s+20[0-3]|aplicar\s+|temas|"
        r"ache a localizacao|onde foi tirada|qual a localizacao|geolocalizar|"
        r"ajuda|help|menu"
        r")\b"
    )
    return bool(re.match(padrao, msg_norm))

def _eh_comando_osint_ou_geo(msg_norm: str) -> bool:
    """Verifica se a mensagem é um comando do ecossistema OSINT ou Geolocalização Visual."""
    if not msg_norm:
        return False
    termos = (
        "!osint", "/osint", "osint",
        "!geo", "/geo", "!geolocalizar", "/geolocalizar",
        "!localizacao", "/localizacao", "!localização", "/localização",
        "!exif", "/exif", "exif",
        "ache a localizacao", "ache a localização", "onde foi tirada",
        "qual a localizacao", "qual a localização"
    )
    return any(msg_norm.startswith(t) for t in termos)

def validar_remetente_autorizado(data: dict) -> bool:
    """
    Segurança do Webhook:
    1. Descarta newsletters, broadcasts e qualquer eco/mensagem gerada pelo próprio bot.
    2. Se o bot estiver em pausa/desligado momentaneamente, permite apenas o comando de religar do proprietário.
    3. Em grupos (@g.us): Segurança estrita com exceção autenticada para IEMA Base Técnica TI.
    4. Em conversas privadas (1-on-1): Permite mensagens do proprietário e comandos autorizados.
    """
    chat_id = data.get("chatId") or data.get("to") or data.get("from") or ""

    # 1. Descarta canais, broadcasts e ecos do próprio bot (independente de fromMe)
    if "@newsletter" in chat_id or "@broadcast" in chat_id:
        return False

    msg_raw = (data.get("caption") or data.get("body") or "").strip()
    if _eh_eco_ou_mensagem_bot(msg_raw):
        return False

    # 2. Verificação de Desligamento / Pausa Momentânea do Bot
    if bot_esta_pausado():
        author_raw = data.get("author") or (data.get("sender") or {}).get("id") or data.get("from") or chat_id or ""
        ids_rem = {extrair_identificador_limpo(chat_id), extrair_identificador_limpo(author_raw)}
        eh_dono = bool(ids_rem & set(IDENTIFICADORES_AUTORIZADOS))
        msg_norm_pausa = normalizar_texto(msg_raw)
        if eh_dono and (_eh_comando_religar_bot(msg_norm_pausa) or msg_norm_pausa in ["!status", "/status"]):
            return True
        log_terminal(f"Bot pausado momentaneamente: mensagem ignorada (chat {chat_id})")
        return False

    # 2. Em grupos (@g.us), validação estrita do autor
    if "@g.us" in chat_id:
        if data.get("fromMe", False):
            msg_b_fm = normalizar_texto(msg_raw)
            if chat_id == ID_GRUPO_BASE_TECNICA or chat_id in GRUPOS_ALUNOS_LAB:
                return msg_raw.startswith(("!", "/")) or _eh_comando_valido_sem_prefixo(msg_b_fm)
            return msg_raw.startswith(("!", "/"))

        author_raw = data.get("author") or (data.get("sender") or {}).get("id") or ""
        author_clean = extrair_identificador_limpo(author_raw)
        if author_clean in IDENTIFICADORES_AUTORIZADOS or any(id_aut in str(author_raw) for id_aut in IDENTIFICADORES_AUTORIZADOS):
            return True

        if chat_id == ID_GRUPO_BASE_TECNICA:
            msg_b = normalizar_texto(msg_raw)
            if _eh_comando_lab_msg(msg_b) or msg_b.startswith(("!sou", "/sou", "!ajuda", "/ajuda", "!help", "/help", "menu", "ajuda")):
                return True
            membros = obter_membros_grupo_base()
            numero_real = author_clean if "@lid" not in str(author_raw) else resolver_lid_para_numero(author_raw)
            eh_membro = (
                (numero_real in membros)
                or author_clean in membros
                or any(m in str(author_raw) for m in membros)
                or (numero_real in PROFS_BASE_POR_NUMERO)
                or (author_clean in PROFS_BASE_POR_NUMERO)
                or not membros
            )
            if eh_membro:
                return True

        if chat_id in GRUPOS_ALUNOS_LAB:
            msg_b = normalizar_texto(msg_raw)
            if _eh_comando_lab_msg(msg_b) or msg_b.startswith(("!ajuda", "/ajuda", "!help", "/help", "menu", "ajuda")):
                return True
            if chat_id in GRUPOS_OSINT_PERMITIDOS:
                if (
                    _eh_comando_osint_ou_geo(msg_b)
                    or _eh_comando_valido_sem_prefixo(msg_b)
                    or data.get("hasMedia")
                    or data.get("type") in ["image", "document"]
                ):
                    return True

        log_terminal(f"Segurança: Mensagem em grupo ignorada (chat {chat_id}, autor {author_raw})")
        return False

    # 3. Conversas privadas (1-on-1):
    # - O próprio dono (self-chat) pode usar seus comandos normalmente.
    # - Outros professores/contatos no privado NÃO disparam o bot em conversas normais:
    #   exigem comando de autenticação/início (!iniciar, !bot, !sou) ou sessão já autenticada.
    if data.get("fromMe", False):
        cid_clean = extrair_identificador_limpo(chat_id)
        if cid_clean in IDENTIFICADORES_AUTORIZADOS:
            return msg_raw.startswith(("!", "/")) or _eh_comando_valido_sem_prefixo(normalizar_texto(msg_raw))
        return False

    remetente = data.get("author") or data.get("from") or chat_id
    ids = {extrair_identificador_limpo(chat_id), extrair_identificador_limpo(remetente)}
    msg_b = normalizar_texto(msg_raw)

    # Se for o próprio dono no privado
    if ids & set(IDENTIFICADORES_AUTORIZADOS):
        return True

    # Contatos restritos (Contato) usam indicador explícito !bot
    if ids & set(contato_restrito.CONTATOS):
        return contato_restrito.tem_indicador_bot(msg_raw)

    # Permite envio de atividades/fotos de caderno por alunos SOMENTE quando acompanhado de comando explícito (!visto, !atv, etc.)
    eh_cmd_atv = bool(re.match(r"^[!/](?:atv|vistos?|atividades?|corrigir|correcao|caderno)\b", msg_b))
    if eh_cmd_atv:
        return True

    # Para docentes e demais contatos no privado: exige comando de iniciar (!iniciar, !bot, !sou) ou sessão privada ativa
    if _eh_comando_iniciar_bot_privado(msg_b) or sessao_bot_privado_ativa(chat_id, remetente):
        return True

    log_terminal(f"Privado ignorado (sem sessão !iniciar ativa): chat {chat_id}")
    return False

ESTADOS_CONVERSA = {}
MENSAGENS_PROCESSADAS = {}  # {msg_id: timestamp}
SESSOES_BOT_PRIVADO = {}    # {id_limpo: timestamp_expiracao}
TTL_SESSAO_PRIVADO_SEG = int(os.environ.get("TTL_SESSAO_PRIVADO_SEG", 900))  # 15 minutos de inatividade no privado
ARQUIVO_FLAG_PAUSA = os.path.join(os.environ.get("DOCENTE_HOME", "."), ".bot_pausado")


def bot_esta_pausado() -> bool:
    """Verifica se o bot está temporariamente desligado/pausado pelo administrador."""
    return os.path.exists(ARQUIVO_FLAG_PAUSA)


def definir_pausa_bot(pausar: bool, motivo: str = "") -> bool:
    """Ativa ou desativa a pausa temporária do bot com persistência de estado."""
    if pausar:
        dados = {
            "pausado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "motivo": motivo or "Pausa temporária manual",
        }
        try:
            with open(ARQUIVO_FLAG_PAUSA, "w", encoding="utf-8") as f:
                json.dump(dados, f, indent=2)
        except Exception as e:
            log_terminal(f"Erro ao salvar flag de pausa: {e}")
        log_terminal(f"⏸️ Assistente Docente pausado/desligado temporariamente às {dados['pausado_em']}")
        return True
    else:
        if os.path.exists(ARQUIVO_FLAG_PAUSA):
            try:
                os.remove(ARQUIVO_FLAG_PAUSA)
            except Exception as e:
                log_terminal(f"Erro ao remover flag de pausa: {e}")
        log_terminal("▶️ Assistente Docente religado/reativado com sucesso.")
        return False


def obter_info_pausa_bot() -> dict:
    if not bot_esta_pausado():
        return {"pausado": False}
    try:
        with open(ARQUIVO_FLAG_PAUSA, "r", encoding="utf-8") as f:
            dados = json.load(f)
            dados["pausado"] = True
            return dados
    except Exception:
        return {"pausado": True, "pausado_em": "recentemente"}


def _eh_comando_pausar_bot(msg_norm: str) -> bool:
    """Detecta comandos do dono para desligar/pausar o bot temporariamente."""
    if not msg_norm:
        return False
    return bool(re.match(r"^[!/]?(?:pausar|desligar|dormir|stopbot|desativar|pausa|silenciarbot)\b", msg_norm))


def _eh_comando_religar_bot(msg_norm: str) -> bool:
    """Detecta comandos do dono para religar/reativar o bot."""
    if not msg_norm:
        return False
    return bool(re.match(r"^[!/]?(?:ligar|reativar|retomar|acordar|startbot|ativar|religar)\b", msg_norm))


def _eh_comando_iniciar_bot_privado(msg_norm: str) -> bool:
    """Verifica se a mensagem no privado é um comando explícito para autenticar/iniciar o bot."""
    if not msg_norm:
        return False
    return bool(re.match(r"^[!/](?:iniciar|start|bot|login|autenticar|sou)\b", msg_norm))


def _eh_comando_encerrar_bot_privado(msg_norm: str) -> bool:
    """Verifica se a mensagem encerra explicitamente a sessão do bot no privado."""
    if not msg_norm:
        return False
    return bool(re.match(r"^[!/]?(?:encerrar|sair|parar\s*bot|fechar\s*bot|tchau)\b", msg_norm))


def _eh_comando_onboarding_docente(msg_norm: str) -> bool:
    """Verifica se a mensagem solicita cadastro de novo professor no ecossistema."""
    if not msg_norm:
        return False
    return bool(
        re.search(
            r"\b(?:!cadastr(?:ar|o)|/cadastr(?:ar|o)|cadastr(?:ar|o)|sou\s*novo|novo\s*docente|novo\s*prof|me\s*cadastrar)\b",
            msg_norm,
        )
    )


def iniciar_sessao_bot_privado(*identificadores, ttl: int = TTL_SESSAO_PRIVADO_SEG):
    exp = time.time() + ttl
    for raw in identificadores:
        limpo = extrair_identificador_limpo(raw)
        if limpo:
            SESSOES_BOT_PRIVADO[limpo] = exp


def renovar_sessao_bot_privado(*identificadores, ttl: int = TTL_SESSAO_PRIVADO_SEG):
    """Estende silenciosamente o prazo de inatividade caso o contato envie novos comandos."""
    agora = time.time()
    for raw in identificadores:
        limpo = extrair_identificador_limpo(raw)
        if limpo and limpo in SESSOES_BOT_PRIVADO:
            SESSOES_BOT_PRIVADO[limpo] = agora + ttl


def encerrar_sessao_bot_privado(*identificadores):
    for raw in identificadores:
        limpo = extrair_identificador_limpo(raw)
        if limpo:
            SESSOES_BOT_PRIVADO.pop(limpo, None)


def sessao_bot_privado_ativa(*identificadores) -> bool:
    agora = time.time()
    for raw in identificadores:
        limpo = extrair_identificador_limpo(raw)
        if not limpo:
            continue
        exp = SESSOES_BOT_PRIVADO.get(limpo, 0)
        if exp > agora:
            return True
        elif exp:
            SESSOES_BOT_PRIVADO.pop(limpo, None)
    return False

class WebhookHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith("/healthz") or self.path.startswith("/health"):
            db_ok = False
            try:
                with sqlite3.connect(ESCOLA_DB, timeout=2) as conn:
                    db_ok = bool(conn.execute("SELECT 1").fetchone())
            except Exception:
                db_ok = False

            openwa_ok = False
            try:
                req_ow = urllib.request.Request(f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}", headers={"x-api-key": OPENWA_API_KEY})
                with urllib.request.urlopen(req_ow, timeout=1.5) as r:
                    openwa_ok = (r.status == 200)
            except Exception:
                openwa_ok = False

            omni_ok = False
            try:
                with urllib.request.urlopen("http://localhost:20128/v1/models", timeout=0.8) as r:
                    omni_ok = (r.status == 200)
            except Exception:
                omni_ok = False

            uptime_s = int(time.time() - TIMESTAMP_INICIO_BOOT)
            disco_livre_gb = round(shutil.disk_usage(BASE_DIR).free / (1024**3), 2)

            resp_payload = {
                "status": "ok" if (db_ok and openwa_ok) else "degraded",
                "uptime_segundos": uptime_s,
                "uptime_legivel": str(timedelta(seconds=uptime_s)),
                "banco_sqlite": db_ok,
                "whatsapp_openwa": openwa_ok,
                "omniroute_llm": omni_ok,
                "disco_livre_gb": disco_livre_gb,
                "bot_pausado": bot_esta_pausado(),
                "modo_vps_leve": MODO_VPS_LEVE
            }
            self.send_response(200 if resp_payload["status"] == "ok" else 503)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(resp_payload, ensure_ascii=False).encode("utf-8"))
            return
        if self.path.startswith("/api/bot/status_pausa"):
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(obter_info_pausa_bot()).encode("utf-8"))
            return
        if self.path.startswith("/loc"):
            return contato_restrito.http_loc(self)
        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        if self.path.startswith("/api/bot/toggle_pausa"):
            esta_p = bot_esta_pausado()
            novo_st = definir_pausa_bot(not esta_p, motivo="API Webhook HTTP")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"sucesso": True, "bot_pausado": novo_st}).encode("utf-8"))
            return
        if self.path.startswith("/loc"):
            return contato_restrito.http_loc(self)
        post_data = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        self.send_response(200)
        self.end_headers()
        try:
            payload = json.loads(post_data.decode("utf-8"))
            event = payload.get("event", "")
            data = payload.get("data", {})

            # Filtra eventos que não sejam novas mensagens (descarta acks, status, presença)
            if event and event not in ["message.received", "message.create", "message", "*", "message.sent"]:
                return

            # DEDUPLICAÇÃO ESTRITA: descarta mensagens repetidas enviadas por múltiplos gatilhos do OpenWA
            msg_id = extrair_id_mensagem(data.get("id"))
            agora = time.time()
            if msg_id:
                if msg_id in MENSAGENS_PROCESSADAS and (agora - MENSAGENS_PROCESSADAS[msg_id] < 120):
                    return
                MENSAGENS_PROCESSADAS[msg_id] = agora
                if len(MENSAGENS_PROCESSADAS) > 2000:
                    for k in list(MENSAGENS_PROCESSADAS.keys()):
                        if agora - MENSAGENS_PROCESSADAS[k] > 3600:
                            MENSAGENS_PROCESSADAS.pop(k, None)

            try:
                contato_restrito.registrar_localizacao(data)
            except Exception as e_loc:
                log_terminal(f"Aviso localização: {e_loc}")

            if not validar_remetente_autorizado(data):
                return

            caption = (data.get("caption") or "").strip()
            body = (data.get("body") or "").strip()
            msg_body = caption if caption else body
            chat_id = data.get("chatId") or data.get("to") or data.get("from") or DONO_CHAT_ID
            remetente = data.get("author") or data.get("from") or chat_id
            author_num = extrair_identificador_limpo(remetente)
            chat_num = extrair_identificador_limpo(chat_id)

            # ROTEAMENTO EXCLUSIVO DE VOZ DO DONO PARA OMNIROUTE (Whisper) -> RAG / Comandos
            eh_dono_priv = "@g.us" not in str(chat_id) and (
                chat_num in IDENTIFICADORES_AUTORIZADOS or author_num in IDENTIFICADORES_AUTORIZADOS
            )
            eh_audio = bool(
                data.get("type") in ["ptt", "audio", "voice"]
                or (data.get("mimetype") or "").startswith("audio/")
            )
            if eh_dono_priv and eh_audio:
                log_terminal(f"🎙️ [VOZ DONO] Áudio detectado no privado do Pedro ({chat_id}). Despachando para OmniRoute...")
                threading.Thread(target=processar_audio_instrucao_dono, args=(chat_id, data)).start()
                return

            if _eh_eco_ou_mensagem_bot(msg_body):
                return

            # Se houver mídia chegando, baixa e faz cache em background
            if data.get("hasMedia") or data.get("mimetype") or data.get("type") in ["image", "document", "video", "audio", "ptt", "voice"]:
                threading.Thread(target=obter_midia_recente, args=(data, chat_id)).start()

            msg_norm = normalizar_texto(msg_body)
            eh_comando_sem_prefixo = _eh_comando_valido_sem_prefixo(msg_norm)

            # Para contatos restritos (Contato), só entra no pipeline se indicou explicitamente com !bot
            remetente = data.get("author") or data.get("from") or chat_id
            eh_priv_autenticado = (
                "@g.us" not in str(chat_id)
                and (_eh_comando_iniciar_bot_privado(msg_norm) or sessao_bot_privado_ativa(chat_id, remetente))
            )
            eh_cmd_onboarding = bool(
                "@g.us" not in str(chat_id) and _eh_comando_onboarding_docente(msg_norm)
            )
            tem_midia_anexo = bool(
                ("@g.us" not in str(chat_id) or chat_id in GRUPOS_OSINT_PERMITIDOS)
                and (data.get("hasMedia") or data.get("mimetype") or data.get("type") in ["document", "image"])
            )
            if contato_restrito.eh_restrito(chat_id, data):
                if data.get("fromMe") or not contato_restrito.tem_indicador_bot(msg_body):
                    return
            elif not (
                msg_body.startswith("!")
                or msg_body.startswith("/")
                or eh_comando_sem_prefixo
                or _eh_comando_osint_ou_geo(msg_norm)
                or eh_priv_autenticado
                or eh_cmd_onboarding
                or tem_midia_anexo
            ) and chat_id not in ESTADOS_CONVERSA:
                return

            log_terminal(f"Comando autorizado no chat {chat_id}: '{msg_body}'")
            threading.Thread(target=processar_comando, args=(chat_id, msg_body, data)).start()
        except Exception as e:
            log_terminal(f"Erro no Webhook: {e}")

    def log_message(self, format, *args):
        pass

ULTIMOS_ALERTAS_ERRO = {}
def _notificar_erro_critico_admin(prof_chave: str, texto_cmd: str, chat_id: str, msg_erro: str):
    agora = time.time()
    chave_alerta = f"{prof_chave}:{msg_erro[:50]}"
    if agora - ULTIMOS_ALERTAS_ERRO.get(chave_alerta, 0) < 300:
        return
    ULTIMOS_ALERTAS_ERRO[chave_alerta] = agora
    alerta = (
        f"🚨 *[ALERTA DOCENTE — FALHA INESPERADA]*\n\n"
        f"• *Docente:* {prof_chave}\n"
        f"• *Chat:* `{chat_id}`\n"
        f"• *Comando:* `{texto_cmd[:100]}`\n"
        f"• *Erro:* `{msg_erro[:250]}`\n"
        f"• *Data/Hora:* {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    )
    enviar_zap(NUMERO_PROPRIETARIO, alerta)

def processar_comando(chat_id, texto, data=None):
    t_inicio = time.time()
    prof_ch = "sistema"
    try:
        author_raw = (data or {}).get("author") or ((data or {}).get("sender") or {}).get("id") or chat_id or ""
        author_num = extrair_identificador_limpo(author_raw)
        prof_ch, _ = identificar_prof_solicitante(author_num, author_raw, texto)
        _processar_comando_real(chat_id, texto, data)
        dur_ms = int((time.time() - t_inicio) * 1000)
        registrar_telemetria(prof_ch, texto, dur_ms, status="ok")
    except Exception as e_proc:
        dur_ms = int((time.time() - t_inicio) * 1000)
        registrar_telemetria(prof_ch, texto, dur_ms, status="erro", erro=str(e_proc))
        log_terminal(f"❌ Falha crítica em '{texto}': {e_proc}")
        _notificar_erro_critico_admin(prof_ch, texto, chat_id, str(e_proc))
        enviar_zap(chat_id, "⚠️ Ocorreu uma instabilidade inesperada ao processar o comando. O administrador foi avisado.")

def _processar_comando_real(chat_id, texto, data=None):
    if data is None:
        data = {}
    if _eh_eco_ou_mensagem_bot(texto):
        return
    cmd_lower = normalizar_texto(texto)
    eh_cmd_emergencia = (
        _eh_comando_pausar_bot(cmd_lower)
        or _eh_comando_religar_bot(cmd_lower)
        or cmd_lower in ["!status", "/status"]
    )
    if not eh_cmd_emergencia and chat_id not in ESTADOS_CONVERSA and not anti_loop.permitir_entrada(chat_id, texto):  # disjuntor anti-loop (entrada)
        log_terminal(f"Anti-loop: entrada bloqueada temporariamente ({chat_id})")
        return
    if ponte_agente.tratar(chat_id, texto, data, enviar_zap):  # dono -> agente (WhatsApp)
        return
    # 🔒 Contatos restritos: lista fechada de comandos / interação via !bot (ver contato_restrito.py)
    if contato_restrito.tratar(chat_id, texto, data, enviar_zap, adicionar_alarme, consultar_llm):
        return
    log_terminal(f"Processando: '{texto}'...")

    author_raw = data.get("author") or (data.get("sender") or {}).get("id") or chat_id or ""
    author_num = extrair_identificador_limpo(author_raw)
    prof_chave, prof_nome = identificar_prof_solicitante(author_num, author_raw, texto)
    eh_privado = "@g.us" not in str(chat_id)
    eh_dono_priv = eh_privado and (
        extrair_identificador_limpo(chat_id) in IDENTIFICADORES_AUTORIZADOS
        or author_num in IDENTIFICADORES_AUTORIZADOS
    )
    eh_dono_geral = (
        extrair_identificador_limpo(chat_id) in IDENTIFICADORES_AUTORIZADOS
        or author_num in IDENTIFICADORES_AUTORIZADOS
        or prof_chave == "pedro"
    )

    # ⏸️ / ▶️ CONTROLE DE PAUSA E DESLIGAMENTO TEMPORÁRIO PELO DONO
    if eh_dono_geral:
        if _eh_comando_religar_bot(cmd_lower):
            definir_pausa_bot(False)
            enviar_zap(
                chat_id,
                "▶️ *Assistente Docente Reativado com Sucesso!*\n\n"
                "O bot voltou a operar normalmente em todos os grupos e conversas privadas.\n"
                "Para pausar a qualquer momento, envie `!pausar` ou use o Portal Web."
            )
            return

        if _eh_comando_pausar_bot(cmd_lower):
            definir_pausa_bot(True, motivo="Comando manual enviado pelo proprietário")
            enviar_zap(
                chat_id,
                "⏸️ *Assistente Docente Pausado Momentaneamente!*\n\n"
                "O bot foi desligado até você decidir religar.\n"
                "Nenhuma mensagem em grupos ou conversas privadas receberá respostas.\n\n"
                "💡 Para religar, basta enviar `!ligar` aqui no WhatsApp ou clicar em *Reativar* no Portal Web."
            )
            return

    if bot_esta_pausado():
        if eh_dono_geral and cmd_lower in ["!status", "/status"]:
            enviar_zap(
                chat_id,
                "⏸️ *O Assistente Docente está atualmente pausado!*\n\n"
                "Nenhuma rotina está respondendo a outros usuários.\n"
                "Para reativar o atendimento, envie `!ligar`."
            )
        return

    # 🔐 CONTROLE DE AUTENTICAÇÃO NO PRIVADO PARA OUTROS DOCENTES / CONTATOS
    if eh_privado and not eh_dono_priv:
        if _eh_comando_onboarding_docente(cmd_lower):
            ESTADOS_CONVERSA[chat_id] = {"tipo": "onboarding_docente_nome", "whatsapp": author_num}
            enviar_zap(
                chat_id,
                "👋 *Olá, Colega Docente! Seja muito bem-vindo(a) ao Assistente Docente IEMA TI!*\n\n"
                "Vamos cadastrar o seu perfil para que você possa consultar sua grade de horários, "
                "gerar slides/exercícios e alimentar sua base RAG pessoal com suas próprias ementas e apostilas.\n\n"
                "📝 *Qual é o seu Nome Completo?* (Ex: Prof. Roberto Mendes ou Profa. Juliana Lima)"
            )
            return

        if _eh_comando_encerrar_bot_privado(cmd_lower):
            encerrar_sessao_bot_privado(chat_id, author_raw)
            enviar_zap(chat_id, f"🔒 *Sessão do Assistente encerrada.* Obrigado, {prof_nome}! Para ativar novamente no privado, digite `!iniciar`.")
            return

        if _eh_comando_iniciar_bot_privado(cmd_lower):
            # Se usou !sou <nome>, registra o vínculo antes de abrir a sessão
            if cmd_lower.startswith(("!sou", "/sou")):
                resto_sou = re.sub(r"^[!/]sou\s*", "", texto, flags=re.I).strip()
                if resto_sou:
                    processar_comando_lab(f"sou {resto_sou}", data)
                    prof_chave, prof_nome = identificar_prof_solicitante(author_num, author_raw, texto)

            if prof_chave == "outro" and prof_chave not in PROFS_BASE_MAP:
                enviar_zap(
                    chat_id,
                    "👋 *Olá! Não localizamos o seu número no catálogo de professores.*\n\n"
                    "Se você leciona no IEMA e deseja utilizar o assistente com suas próprias disciplinas e RAG, "
                    "basta responder `!cadastrar` para iniciar o fluxo rápido de integração!"
                )
                return

            iniciar_sessao_bot_privado(chat_id, author_raw)
            resposta_ini = (
                f"👋 *Olá, {prof_nome}!* Sessão autenticada no Assistente do Laboratório & Docente IEMA TI.\n\n"
                f"Como docente da Base Técnica, sua sessão fica ativa durante o uso (encerra silenciosamente após 15 minutos de inatividade). Comandos disponíveis:\n\n"
                f"🖥️ *Comandos Rápidos:*\n"
                f"• `minha grade` ou `!grade` → Ver seus horários de aula (`minha grade hoje` / `amanha`)\n"
                f"• `!rag <tema>` → Consultar sua base de conteúdos/ementa da disciplina\n"
                f"• `!anexar` ou enviar arquivo → Indexar novos slides (.pptx) ou guias (.docx/.pdf) no seu RAG\n"
                f"• `!atv nova <turma> <titulo>` → Cadastrar nova atividade avaliativa\n"
                f"• `!cobrar <turma>` → Ver alunos pendentes e entregas com visto\n"
                f"• `!atv <turma> todos exceto <nºs> ok` → Dar visto em massa para a turma\n"
                f"• `lab` → Ver horários livres e reservas de hoje\n"
                f"• `lab amanha` → Ver horários de amanhã\n"
                f"• `lab semana` → Ver e receber o PDF da grade semanal em alta visibilidade\n"
                f"• `lab 1 2` → Reservar o 1º e 2º horários de hoje\n"
                f"• `lab amanha 4 5` → Reservar horários de amanhã\n"
                f"• `lab cancelar` → Cancelar suas reservas\n"
                f"• `!sair` → Encerrar o atendimento do bot no privado\n\n"
                f"💡 Digite qualquer um dos comandos acima para continuar!"
            )
            enviar_zap(chat_id, resposta_ini)
            return

        # Se não enviou comando de iniciar e NÃO tem sessão privada ativa, ignora silenciosamente
        # (a menos que seja entrega explícita de atividade com !visto / !atv ou conversa em andamento)
        eh_cmd_atv_priv = bool(re.match(r"^[!/](?:atv|vistos?|atividades?|corrigir|correcao|caderno)\b", cmd_lower))
        if not sessao_bot_privado_ativa(chat_id, author_raw) and not eh_cmd_atv_priv and chat_id not in ESTADOS_CONVERSA:
            log_terminal(f"Privado ignorado sem sessão ativa ({prof_nome} - {chat_id}): '{texto}'")
            return

        # Sessão ativa no privado: renova sliding timeout de inatividade
        renovar_sessao_bot_privado(chat_id, author_raw)

    # 📎 Instrução e estado para envio de arquivos para o RAG Docente
    if cmd_lower.startswith(("!anexar", "/anexar", "!upload", "/upload", "!anexo", "/anexo")):
        disc_informada = ""
        if " " in texto:
            disc_informada = texto.split(" ", 1)[1].strip()
        tem_doc_no_comando = bool(
            data.get("hasMedia")
            or data.get("mimetype")
            or data.get("type") in ["document", "image"]
        )
        if not tem_doc_no_comando:
            ESTADOS_CONVERSA[chat_id] = {
                "tipo": "aguardando_doc_rag",
                "prof_chave": prof_chave,
                "prof_nome": prof_nome,
                "disciplina": disc_informada,
            }
            nome_d = prof_nome if prof_nome else "Professor(a)"
            hint_disc = f" ({disc_informada})" if disc_informada else ""
            enviar_zap(
                chat_id,
                f"📎 *Envio de Documentos para o RAG Docente — {nome_d}{hint_disc}:*\n\n"
                "Envie seu arquivo (**PDF, Word .docx, Slides .pptx ou Texto .txt/.md**) diretamente como anexo nesta conversa!\n"
                "• O conteúdo será extraído e salvo no seu RAG individual com isolamento por matéria.\n"
                "• Suas próximas aulas e planos serão gerados automaticamente com base neste material.\n\n"
                "*(Envie o arquivo como anexo agora ou digite `!cancelar` para cancelar)*"
            )
            return

    # 📥 RECONHECIMENTO E INDEXAÇÃO DE DOCUMENTOS / SLIDES PARA O RAG DO DOCENTE
    tem_doc_media = bool(
        data.get("hasMedia")
        or data.get("mimetype")
        or data.get("type") in ["document", "image"]
    )
    eh_aguardando_rag = (chat_id in ESTADOS_CONVERSA and ESTADOS_CONVERSA[chat_id].get("tipo") == "aguardando_doc_rag")
    if tem_doc_media and (eh_privado or "!rag" in cmd_lower or "!anexo" in cmd_lower or "!anexar" in cmd_lower or "!ementa" in cmd_lower or eh_aguardando_rag):
        caminho_midia = obter_midia_recente(data, chat_id)
        if caminho_midia and os.path.exists(caminho_midia):
            ext_m = os.path.splitext(caminho_midia)[1].lower()
            if ext_m in [".pdf", ".docx", ".pptx", ".txt", ".md", ".csv"]:
                alvo_chave = prof_chave
                if eh_aguardando_rag and ESTADOS_CONVERSA[chat_id].get("prof_chave"):
                    alvo_chave = ESTADOS_CONVERSA[chat_id]["prof_chave"]
                elif alvo_chave == "outro" and chat_id in ESTADOS_CONVERSA:
                    alvo_chave = ESTADOS_CONVERSA[chat_id].get("prof_chave", "outro")
                if alvo_chave in PROFS_BASE_MAP or alvo_chave != "outro":
                    nome_orig = data.get("filename") or os.path.basename(caminho_midia)
                    caption_t = (data.get("caption") or texto or "").strip()
                    materia_hint = None
                    if eh_aguardando_rag and ESTADOS_CONVERSA[chat_id].get("disciplina"):
                        materia_hint = ESTADOS_CONVERSA[chat_id]["disciplina"]
                    if not materia_hint and ("disciplina:" in caption_t.lower() or "materia:" in caption_t.lower()):
                        materia_hint = re.split(r"(?:disciplina|materia):\s*", caption_t, flags=re.I)[1].split("\n")[0].strip()
                    elif not materia_hint and ("!anexar " in caption_t.lower() or "/anexar " in caption_t.lower()):
                        materia_hint = re.split(r"/[!]anexar\s*", caption_t, flags=re.I)[-1].strip()
                    res_rag = gestor_docentes_rag.adicionar_documento_ao_rag(
                        prof_chave=alvo_chave,
                        caminho_arquivo=caminho_midia,
                        nome_original=nome_orig,
                        materia_id=materia_hint
                    )
                    if res_rag.get("sucesso"):
                        if chat_id in ESTADOS_CONVERSA and ESTADOS_CONVERSA[chat_id].get("tipo") == "aguardando_doc_rag":
                            ESTADOS_CONVERSA.pop(chat_id, None)
                        renovar_sessao_bot_privado(chat_id, author_raw)
                        resp_ok = (
                            f"✅ *Documento Indexado com Sucesso no seu RAG!*\n\n"
                            f"👤 *Docente:* {res_rag['docente']}\n"
                            f"📄 *Arquivo:* `{res_rag['arquivo']}`\n"
                            f"📘 *Disciplina Isolada:* {res_rag['disciplina']}\n"
                            f"🧠 *Conteúdo Extraído:* {res_rag['caracteres']} caracteres ({res_rag['linhas']} linhas)\n"
                            f"🛡️ *Garantia:* Material indexado no RAG individual com isolamento estrito.\n\n"
                            f"💡 *Suas próximas aulas utilizarão este documento como base!*\n"
                            f"• Digite `!preparar hoje` ou `!preparar amanha` para gerar slides e lista\n"
                            f"• Digite `!rag {res_rag['disciplina_slug'][:10]}` para consultar sua base\n"
                            f"• Digite `!exercicio {res_rag['disciplina']}` para gerar atividades\n"
                            f"• Você pode enviar novos materiais ou slides a qualquer momento!"
                        )
                        enviar_zap(chat_id, resp_ok)
                        return
                        enviar_zap(chat_id, resp_ok)
                        return

    # Cancelamento de conversa pendente
    if cmd_lower in ["!cancelar", "/cancelar", "cancelar"]:
        if chat_id in ESTADOS_CONVERSA:
            ESTADOS_CONVERSA.pop(chat_id, None)
            enviar_zap(chat_id, "Operação cancelada.")
            return

    # Tratamento de conversa interativa (Onboarding e Exercícios)
    if chat_id in ESTADOS_CONVERSA:
        st = ESTADOS_CONVERSA[chat_id]
        if st.get("tipo") == "onboarding_docente_nome":
            nome_doc = texto.strip()
            permitido, motivo, chave_auto = gestor_docentes_rag.validar_docente_permitido(nome_doc)
            if not permitido:
                ESTADOS_CONVERSA.pop(chat_id, None)
                enviar_zap(chat_id, f"⛔ *Cadastro Não Autorizado:*\n\n{motivo}")
                return
            ESTADOS_CONVERSA[chat_id] = {
                "tipo": "onboarding_docente_disciplinas",
                "nome": nome_doc,
                "chave": chave_auto,
                "whatsapp": st.get("whatsapp", ""),
            }
            enviar_zap(
                chat_id,
                f"Muito bem, *{nome_doc}*! 📘\n\n"
                "Quais *disciplinas* você leciona no IEMA? (Pode listar separadas por vírgula. Ex: Marketing Digital, Legislação, Informática Aplicada):"
            )
            return

        elif st.get("tipo") == "onboarding_docente_disciplinas":
            discs = [d.strip() for d in re.split(r"[,;\n]+", texto) if d.strip()]
            if not discs:
                enviar_zap(chat_id, "Por favor, informe ao menos uma disciplina.")
                return
            st["disciplinas"] = discs
            st["tipo"] = "onboarding_docente_turmas"
            ESTADOS_CONVERSA[chat_id] = st
            enviar_zap(
                chat_id,
                "Perfeito! 👥 Quais *turmas* você atende atualmente? (Ex: 101, 102, 201. Ou digite 'todas'):"
            )
            return

        elif st.get("tipo") == "onboarding_docente_turmas":
            turmas_txt = texto.strip()
            turmas_lista = [t.strip() for t in re.split(r"[,;\s]+", turmas_txt) if re.match(r"^\d{3}$", t.strip())]
            if not turmas_lista:
                turmas_lista = ["101", "102"]

            nome_final = st.get("nome", "Professor")
            disc_final = st.get("disciplinas", ["Base Técnica"])
            w_final = st.get("whatsapp", "") or extrair_identificador_limpo(chat_id)

            novo = gestor_docentes_rag.cadastrar_novo_docente(
                nome=nome_final,
                whatsapp=w_final,
                disciplinas=disc_final,
                turmas=turmas_lista,
            )
            if not novo.get("sucesso"):
                ESTADOS_CONVERSA.pop(chat_id, None)
                enviar_zap(chat_id, f"⛔ *Cadastro Não Autorizado:*\n\n{novo.get('erro')}")
                return

            ch = novo.get("chave")
            PROFS_BASE_MAP[ch] = nome_final
            if w_final:
                PROFS_BASE_POR_NUMERO[w_final] = (ch, nome_final)
                salvar_vinculo_professor(w_final, ch, nome_final)
            iniciar_sessao_bot_privado(chat_id, w_final)

            ESTADOS_CONVERSA[chat_id] = {
                "tipo": "aguardando_doc_rag",
                "prof_chave": ch,
                "prof_nome": nome_final,
                "disciplinas": disc_final,
            }

            discs_str = ", ".join(disc_final)
            msg_sucesso = (
                f"🎉 *Cadastro Concluído com Sucesso, {nome_final}!* 🎓\n\n"
                f"• Chave no Sistema: `{ch}`\n"
                f"• Disciplinas com RAG Isolado: *{discs_str}*\n"
                f"• Turmas: *{', '.join(turmas_lista)}*\n\n"
                "📂 *Agora envie seus materiais:* Você pode me enviar arquivos (**PDF, Word .docx ou Slides .pptx**) "
                "aqui no WhatsApp para compor a sua base de conhecimento anti-alucinação!\n\n"
                "🖥️ *Comandos para testar a qualquer momento:*\n"
                "• `minha grade` → Consultar seus horários\n"
                "• `!rag <tema>` → Consultar trechos das suas ementas/guias\n"
                "• `lab` / `lab semana` → Consultar e agendar o laboratório\n"
                "• `!slides <tema>` → Gerar estrutura de apresentação\n"
                "• `!exercicio <tema>` → Gerar lista de exercícios alinhada à BNCC\n"
                "• `!sair` → Encerrar atendimento privado"
            )
            enviar_zap(chat_id, msg_sucesso)
            return

        elif st.get("tipo") == "aguardando_doc_rag":
            if not data.get("hasMedia") and not data.get("mimetype"):
                enviar_zap(
                    chat_id,
                    "💡 *Aguardando seu material didático:*\n"
                    "Envie um arquivo (**Word .docx, Slides .pptx ou PDF**) com a ementa ou aula da sua matéria.\n"
                    "Ou digite `!rag` para testar a base inicial, ou `minha grade` para ver seus horários!"
                )
                return

        elif st.get("tipo") == "exercicio_tema":
            tema = texto.strip()
            ESTADOS_CONVERSA[chat_id] = {"tipo": "exercicio_detalhes", "tema": tema}
            msg = (
                f"Tema: {tema}\n\n"
                "Envie os detalhes desejados:\n"
                "1. Serie (1, 2 ou 3 ano EM)\n"
                "2. Foco (teoria basica, cotidiano ou pratica)\n"
                "3. Formato (multipla escolha ou aberta)\n"
                "4. Quantidade (padrao: 5)\n\n"
                "(Pode enviar tudo em uma linha so. Ex: 2 ano, cotidiano, multipla escolha, 5. Ou !cancelar)"
            )
            enviar_zap(chat_id, msg)
            return
        elif st.get("tipo") == "exercicio_detalhes":
            tema = st.get("tema", "Conteúdo")
            ESTADOS_CONVERSA.pop(chat_id, None)

            txt_l = normalizar_texto(texto)
            texto_sem_serie = re.sub(r"\b[123]\s*(?:º|o|ª|a)?\s*ano\b", "", txt_l)
            m_qtd = re.search(r"\b(\d+)\b", texto_sem_serie)
            qtd = int(m_qtd.group(1)) if m_qtd else 5
            if qtd > 30: qtd = 30

            serie = "Ensino Médio"
            if "1" in txt_l or "primeir" in txt_l:
                serie = "1º ano do Ensino Médio"
            elif "2" in txt_l or "segund" in txt_l:
                serie = "2º ano do Ensino Médio"
            elif "3" in txt_l or "terceir" in txt_l:
                serie = "3º ano do Ensino Médio"

            formato = "múltipla escolha (com 4 ou 5 alternativas e gabarito)"
            if "abert" in txt_l or "discursiv" in txt_l:
                formato = "questões abertas/discursivas comentadas"
            elif "misto" in txt_l or "ambos" in txt_l:
                formato = "misto (múltipla escolha e discursivas)"

            foco = texto.strip()
            enviar_zap(chat_id, f"Gerando {qtd} questoes para {serie} sobre '{tema}'...")
            threading.Thread(target=worker_gerar_exercicio, args=(chat_id, tema, qtd, serie, foco, formato)).start()
            return

    # 🎓 COMANDOS MULTI-DOCENTE (Grade Individual e Base RAG Isolada por Professor)
    m_grade_ind = re.match(r"^(?:[!/]?(?:minha\s*grade|meus\s*horarios?))\b(.*)$", cmd_lower)
    if m_grade_ind and prof_chave in PROFS_BASE_MAP:
        filtro_g = m_grade_ind.group(1).strip()
        alvo_prof = prof_chave
        # Permite ao Prof. Pedro testar/consultar a grade de outro docente (ex: "minha grade marcos" ou "minha grade bruna hoje")
        if prof_chave == "pedro" and filtro_g:
            partes_g = filtro_g.split()
            if partes_g[0] in PROFS_BASE_MAP:
                alvo_prof = partes_g[0]
                filtro_g = " ".join(partes_g[1:]).strip()
        resp_g = gestor_docentes_rag.consultar_grade_individual_docente(alvo_prof, filtro_g)
        enviar_zap(chat_id, resp_g)
        return

    m_rag_ind = re.match(r"^[!/](?:rag|ementa)\b(.*)$", cmd_lower)
    if m_rag_ind and prof_chave in PROFS_BASE_MAP:
        termo_rag = m_rag_ind.group(1).strip()
        alvo_prof_rag = prof_chave
        # Permite ao Prof. Pedro testar/simular o RAG de qualquer outro docente (ex: "!rag bruna marketing" ou "!rag marcos html")
        if prof_chave == "pedro" and termo_rag:
            partes_r = termo_rag.split()
            if partes_r[0] in PROFS_BASE_MAP:
                alvo_prof_rag = partes_r[0]
                termo_rag = " ".join(partes_r[1:]).strip()
        resp_r = gestor_docentes_rag.formatar_resposta_rag_docente(alvo_prof_rag, termo_rag)
        enviar_zap(chat_id, resp_r)
        return

    if eh_privado and prof_chave in PROFS_BASE_MAP and prof_chave != "pedro":
        if not _eh_comando_lab_msg(cmd_lower):
            if cmd_lower.startswith(("!grade", "/grade", "!horario", "/horario")):
                resto_h = re.sub(r"^[!/](?:grade|horarios?)\s*", "", cmd_lower).strip()
                enviar_zap(chat_id, gestor_docentes_rag.consultar_grade_individual_docente(prof_chave, resto_h))
                return
            resposta = (
                f"👋 *Olá, {prof_nome}!* Sessão ativa no Assistente do Laboratório IEMA TI.\n\n"
                f"🖥️ *Comandos disponíveis na sua sessão:*\n"
                f"• `minha grade` → Ver seus horários de aula\n"
                f"• `!rag <tema>` → Consultar sua base de conteúdos/ementa\n"
                f"• `lab` / `lab amanha` / `lab semana` → Consultar laboratório\n"
                f"• `lab 1 2` / `lab cancelar` → Reservar ou cancelar horários\n"
                f"• `!sair` → Encerrar sessão do bot"
            )
            enviar_zap(chat_id, resposta)
            return

    # ☁️ SEPARAÇÃO MODULAR VPS (1 GB RAM) vs HOST LOCAL: Fallback gracioso para tarefas pesadas
    if MODO_VPS_LEVE and cmd_lower.startswith(("!preparar", "/preparar", "!slides", "/slides", "!cloud", "/cloud", "!executar", "/executar")):
        msg_fb = delegar_ou_fallback_tarefa_pesada(chat_id, texto, data, nome_recurso=cmd_lower.split()[0])
        if msg_fb:
            enviar_zap(chat_id, msg_fb)
            return

    # 🔒 CONTROLE RESTRITO PARA O GRUPO DA BASE TÉCNICA
    if chat_id == ID_GRUPO_BASE_TECNICA:
        if cmd_lower.startswith(("!ajuda", "/ajuda", "!help", "/help", "menu", "ajuda")):
            resposta = (
                "🖥️ *LABORATÓRIO IEMA TI — COMO USAR*\n\n"
                "• `lab` → Ver horários de hoje\n"
                "• `lab amanha` → Ver horários de amanhã\n"
                "• `lab 1 2` → Reservar 1º e 2º horário de hoje\n"
                "• `lab amanha 4 5` → Reservar 4º e 5º de amanhã\n"
                "• `lab cancelar` → Cancelar sua reserva de hoje\n"
                "• `lab semana` → Ver reservas da semana"
            )
            enviar_zap(chat_id, resposta)
            return

        if not (_eh_comando_lab_msg(cmd_lower) or cmd_lower.startswith(("!sou", "/sou", "resumo", "!resumo", "/resumo"))):
            log_terminal(f"Segurança: Comando restrito '{texto}' bloqueado no grupo da Base Técnica.")
            return

    # 🔒 CONTROLE RESTRITO PARA OS GRUPOS DE ALUNOS (SEQUESTRO / CLUBE DE PROGRAMAÇÃO / INFRA / HACKATONS)
    if chat_id in GRUPOS_ALUNOS_LAB:
        if cmd_lower.startswith(("!ajuda", "/ajuda", "!help", "/help", "menu", "ajuda")):
            if chat_id in GRUPOS_OSINT_PERMITIDOS:
                resposta = (
                    "💻 *CLUBE DE PROGRAMAÇÃO & HACKING / LAB TI*\n\n"
                    "🖥️ *Laboratório:*\n"
                    "• `lab` → Ver horários de hoje\n"
                    "• `lab amanha` → Ver horários de amanhã\n"
                    "• `lab 6 7` → Solicitar 6º e 7º horários\n"
                    "• `lab cancelar` → Cancelar pedido de reserva\n\n"
                    "🔍 *Módulo OSINT & Investigação Liberados:*\n"
                    "• `!osint nick <usuario>` → Busca de perfis em centenas de redes (Sherlock)\n"
                    "• `!osint email <email>` → Descobre contas/serviços vinculados (Holehe)\n"
                    "• `!osint phone <numero>` → Dossiê de telefone e operadora (Ignorant)\n"
                    "• `!osint domain <dominio>` → Dossiê de rede, DNS, MX e WHOIS\n"
                    "• `!osint insta <usuario>` → Dossiê de perfil público Instagram\n"
                    "• Enviar foto com `ache a localizacao` ou `!geo` → Investigação visual Geo-Sleuth\n"
                    "• Enviar foto ou documento com `!exif` → Extração forense de metadados"
                )
            else:
                resposta = (
                    "🖥️ *LABORATÓRIO IEMA TI — ALUNOS*\n\n"
                    "• `lab` → Ver horários de hoje\n"
                    "• `lab amanha` → Ver horários de amanhã\n"
                    "• `lab 6 7` → Pedir 6º e 7º horário de hoje\n"
                    "• `lab amanha 6 7` → Pedir para amanhã\n"
                    "• `lab cancelar` → Cancelar seu pedido"
                )
            enviar_zap(chat_id, resposta)
            return

        if chat_id in GRUPOS_OSINT_PERMITIDOS and _eh_comando_osint_ou_geo(cmd_lower):
            pass  # Liberado para os grupos técnicos de alunos (Clube de Programação e Infra)
        elif not _eh_comando_lab_msg(cmd_lower):
            log_terminal(f"Segurança: Comando restrito '{texto}' bloqueado no grupo de alunos ({chat_id}).")
            return

    # Se um comando estritamente privado/administrativo foi acionado em qualquer grupo (@g.us),
    # redireciona o chat_id de resposta para o privado do professor (exceto comandos OSINT nos grupos liberados).
    if "@g.us" in str(chat_id):
        if chat_id in GRUPOS_OSINT_PERMITIDOS and _eh_comando_osint_ou_geo(cmd_lower):
            pass  # Resposta permanece diretamente no grupo de alunos!
        elif cmd_lower.startswith((
            "!cloud", "/cloud", "!status", "/status", "!ssh", "!rede",
            "!alarme", "/alarme", "!aviso", "!lembrete", "!alarmes", "/alarmes",
            "!ideia", "/ideia", "!nota", "!executar", "/executar", "!rodar", "!run",
            "!osint", "/osint", "!exif", "/exif", "!geo", "ache a localizacao", "onde foi tirada",
            "!gabarito", "/gabarito", "!preparar", "/preparar", "!planejar", "!plano", "/plano", "!dinamica", "/dinamica",
            "!diario", "/diario", "!aluno", "/aluno", "!adicionaraula", "!removeraula", "!modo", "/modo", "modo"
        )) or "privado" in cmd_lower.split():
            if prof_chave == "bruna":
                chat_id = DONO_CHAT_ID
            elif prof_chave == "marcos":
                chat_id = DONO_CHAT_ID
            elif author_raw and "@lid" in str(author_raw):
                chat_id = author_raw
            else:
                chat_id = CHAT_PRIVADO_PADRAO

    # !exercicio [tema]
    if cmd_lower.startswith(("!exercicio", "/exercicio")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        if not resto:
            ESTADOS_CONVERSA[chat_id] = {"tipo": "exercicio_tema"}
            enviar_zap(chat_id, "Qual o tema ou assunto da lista de exercicios?")
            return

        ESTADOS_CONVERSA[chat_id] = {"tipo": "exercicio_detalhes", "tema": resto}
        msg = (
            f"Tema: {resto}\n\n"
            "Envie os detalhes desejados:\n"
            "1. Serie (1, 2 ou 3 ano EM)\n"
            "2. Foco (teoria basica, cotidiano ou pratica)\n"
            "3. Formato (multipla escolha ou aberta)\n"
            "4. Quantidade (padrao: 5)\n\n"
            "(Pode enviar tudo em uma linha so. Ex: 2 ano, cotidiano, multipla escolha, 5. Ou !cancelar)"
        )
        enviar_zap(chat_id, msg)
        return



    # !slides <Tema>
    elif cmd_lower.startswith(("!slide", "/slide")):
        tema = texto.split(" ", 1)[1] if " " in texto else ""
        if not tema:
            enviar_zap(chat_id, "Use: `!slides <Tema>` (ex: `!slides Introdução ao Docker`)")
            return
        enviar_zap(chat_id, f"⏳ Gerando apresentação de slides para \"{tema}\"...\nLocal de destino: Google Drive > Aulas > Slides")
        threading.Thread(target=worker_gerar_slides, args=(chat_id, tema)).start()
        return

    # !temas / !conteudos (Catálogo de conteúdos passados e futuros para escolher)
    elif cmd_lower.startswith(("!temas", "/temas", "temas", "!conteudos", "/conteudos")):
        try:
            import gerador_aulas_rotina
            resposta = gerador_aulas_rotina.listar_catalogo_temas()
        except Exception as e_cat:
            resposta = f"Erro ao listar catálogo de temas: {e_cat}"
        enviar_zap(chat_id, resposta)
        return

    # !aplicar <turma> <tema> (Escolhe e aplica um conteúdo específico na turma com slides, PDF e visto)
    elif cmd_lower.startswith(("!aplicar", "/aplicar", "aplicar ")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        m_t = re.match(r"^([1-3]0[0-3])\s+(.+)$", resto)
        if not m_t:
            import gerador_aulas_rotina
            enviar_zap(chat_id, "Informe a turma e o tema! Ex: `aplicar 203 RAG e guardrails` ou `aplicar 101 mix de marketing`\n\n" + gerador_aulas_rotina.listar_catalogo_temas())
            return
        turma_alvo, tema_alvo = m_t.group(1), m_t.group(2).strip()
        t_low = tema_alvo.lower()
        if turma_alvo == "101" or prof_chave == "bruna" or any(k in t_low for k in ("mkt", "marketing", "swot", "4p", "cdc", "consumidor", "mercadol")):
            alvo_prof_ap = "bruna"
        elif turma_alvo in ("102", "103") or prof_chave == "marcos" or any(k in t_low for k in ("web", "html", "css", "dom", "frontend", "front-end", "backend", "javascript")):
            alvo_prof_ap = "marcos"
        else:
            alvo_prof_ap = "pedro"
        enviar_zap(chat_id, f"⏳ *Preparando conteúdo '{tema_alvo}' para a Turma {turma_alvo}*...\nGerando slides (.pptx), lista (.pdf), exercícios de caderno e atividade no banco!")
        def _worker_aplicar_tema(t_cod, t_tema, c_id, p_ch):
            try:
                import gerador_aulas_rotina
                gerador_aulas_rotina.aplicar_conteudo_turma(t_cod, t_tema, termo_dia="hoje", prof_chave=p_ch, enviar_whatsapp=True, chat_destino=c_id)
            except Exception as e_ap:
                enviar_zap(c_id, f"❌ Erro ao aplicar tema '{t_tema}' na turma {t_cod}: {e_ap}")
        threading.Thread(target=_worker_aplicar_tema, args=(turma_alvo, tema_alvo, chat_id, alvo_prof_ap)).start()
        return

    # !preparar [amanha|hoje|dd/mm] OU !preparar <turma> <tema>
    elif cmd_lower.startswith(("!preparar", "/preparar", "!planejar", "/planejar", "!gerar aulas", "/gerar aulas")):
        param_dia = texto.split(" ", 1)[1].strip() if " " in texto else "amanha"
        if not param_dia:
            param_dia = "amanha"

        alvo_prof = prof_chave
        if prof_chave == "pedro" and param_dia.lower().startswith("bruna"):
            alvo_prof = "bruna"
            param_dia = param_dia[5:].strip() or "amanha"
        elif prof_chave == "pedro" and param_dia.lower().startswith("marcos"):
            alvo_prof = "marcos"
            param_dia = param_dia[6:].strip() or "amanha"

        m_turma_tema = re.match(r"^([1-3]0[0-3])\s+(.+)$", param_dia)
        if m_turma_tema:
            turma_alvo, tema_alvo = m_turma_tema.group(1), m_turma_tema.group(2).strip()
            t_low = tema_alvo.lower()
            if turma_alvo == "101" or any(k in t_low for k in ("mkt", "marketing", "swot", "4p", "cdc", "consumidor", "mercadol")):
                alvo_prof = "bruna"
            elif turma_alvo in ("102", "103") or any(k in t_low for k in ("web", "html", "css", "dom", "frontend", "front-end", "backend", "javascript")):
                alvo_prof = "marcos"
            enviar_zap(chat_id, f"⏳ *Preparando conteúdo '{tema_alvo}' para a Turma {turma_alvo}*...\nGerando slides (.pptx), lista (.pdf), exercícios de caderno e atividade no banco!")
            def _worker_prep_tema(t_cod, t_tema, c_id, p_ch):
                try:
                    import gerador_aulas_rotina
                    gerador_aulas_rotina.aplicar_conteudo_turma(t_cod, t_tema, termo_dia="hoje", prof_chave=p_ch, enviar_whatsapp=True, chat_destino=c_id)
                except Exception as e_ap:
                    enviar_zap(c_id, f"❌ Erro ao preparar tema '{t_tema}' para a turma {t_cod}: {e_ap}")
            threading.Thread(target=_worker_prep_tema, args=(turma_alvo, tema_alvo, chat_id, alvo_prof)).start()
            return

        nome_doc = "Prof. Marcos" if alvo_prof == "marcos" else ("Profa. Bruna" if alvo_prof == "bruna" else "Prof. Pedro")
        enviar_zap(chat_id, f"⏳ *Preparando aulas para '{param_dia}' ({nome_doc})*...\nConsultando grade, estruturando dinâmicas, gerando slides (.pptx) e listas (.pdf). Você receberá os arquivos aqui em instantes!")
        def _worker_preparar(p_dia, c_id, p_ch):
            try:
                import gerador_aulas_rotina
                gerador_aulas_rotina.preparar_aulas_dia(p_dia, prof_chave=p_ch, enviar_whatsapp=True, chat_destino=c_id)
            except Exception as e_prep:
                enviar_zap(c_id, f"❌ Erro ao preparar aulas para '{p_dia}': {e_prep}")
        threading.Thread(target=_worker_preparar, args=(param_dia, chat_id, alvo_prof)).start()
        return

    # !plano [1..8|semana] (Consulta o plano curricular do semestre)
    elif cmd_lower.startswith(("!plano", "/plano")):
        resto = texto.split(" ", 1)[1].strip().lower() if " " in texto else ""
        caminho_plano = os.path.join(os.environ.get("DOCENTE_HOME", "."), "PLANO_CURRICULAR_IA_FIM_SEMESTRE.md")
        if os.path.exists(caminho_plano):
            try:
                with open(caminho_plano, "r", encoding="utf-8") as f:
                    conteudo = f.read()
                
                # Se pediu uma semana específica (ex: !plano 2 ou !plano semana 3)
                m_sem = re.search(r"\b([1-8])\b", resto)
                if m_sem:
                    num = m_sem.group(1)
                    padrao = rf"### Semana {num}:.*?(?=### Semana \d+:|$)"
                    bloco = re.search(padrao, conteudo, re.DOTALL)
                    if bloco:
                        resposta = f"📅 *PLANO CURRICULAR — SEMANA {num}*\n\n" + bloco.group(0).strip()
                    else:
                        resposta = f"Semana {num} não localizada no plano."
                else:
                    resposta = (
                        "📚 *PLANO CURRICULAR — FIM DE SEMESTRE (8 SEMANAS)*\n\n"
                        "• *Semana 1:* Agentes Autônomos & Ciclo ReAct (Decisões Automatizadas - LGPD Art. 20)\n"
                        "• *Semana 2:* Protocolo MCP & Tool Calling (Lei Carolina Dieckmann & Execução Indevida)\n"
                        "• *Semana 3:* RAG Avançado, Embeddings e Reranking (Sigilo & Minimização)\n"
                        "• *Semana 4:* Prompt Injections & Red Teaming (Responsabilidade Civil em IA)\n"
                        "• *Semana 5:* Guardrails Determinísticos (Compliance & Auditoria PL 2338/23)\n"
                        "• *Semana 6:* Sistemas Multiagentes & Orquestração (Cadeias Distribuídas)\n"
                        "• *Semana 7:* Projeto Prático: Assistente Escolar MCP\n"
                        "• *Semana 8:* Pitches Técnicos, Fechamento de Vistos e Avaliação Final\n\n"
                        "💡 _Para ver detalhes de uma semana envie: `!plano <1 a 8>`_"
                    )
            except Exception as e_plano:
                resposta = f"Erro ao ler plano curricular: {e_plano}"
        else:
            resposta = "Arquivo do plano curricular não encontrado."
        enviar_zap(chat_id, resposta)
        return

    # !dinamica [tema|amanha] (Dicas de dinâmicas práticas em sala)
    elif cmd_lower.startswith(("!dinamica", "/dinamica", "!pratica", "/pratica")):
        termo = texto.split(" ", 1)[1].strip().lower() if " " in texto else ""
        if "prompt" in termo or "injec" in termo or "seguranca" in termo or "legislacao" in termo or termo in ["amanha", "terca"]:
            resposta = (
                "🎯 *DINÂMICA: O Laboratório de Prompt Injection & LGPD*\n\n"
                "1. **Cenário:** Simule um assistente de RH de um hospital escolar que recebe currículos e mensagens de alunos.\n"
                "2. **Missão dos Alunos (Red Team):** Eles devem elaborar um parágrafo que induza o bot a revelar os nomes de alunos com notas baixas ou liberar acesso root.\n"
                "3. **Discussão Pedagógica:** Conecte o resultado com o Art. 46 da LGPD (dever de segurança) e Art. 20 (revisão humana de decisões automatizadas).\n"
                "4. **Duração recomendada:** 25 minutos."
            )
        elif "mcp" in termo or "react" in termo or "agente" in termo or "quarta" in termo:
            resposta = (
                "🎯 *DINÂMICA: O Teatro do Protocolo MCP & ReAct*\n\n"
                "1. **Papéis:**\n"
                "   • Aluno A (Usuário): Faz um pedido complexo de sistema.\n"
                "   • Aluno B (Agente ReAct): Escreve no quadro seu pensamento (Thought) e a chamada JSON (Action).\n"
                "   • Aluno C (Servidor MCP): Executa a busca em uma folha de dados e devolve o resultado (Observation).\n"
                "2. **Objetivo:** Mostrar visualmente que LLMs não 'sabem tudo', mas aprendem a invocar ferramentas padronizadas."
            )
        else:
            resposta = (
                "💡 *DINÂMICAS PRÁTICAS DISPONÍVEIS:*\n\n"
                "• `!dinamica amanha` → Red Teaming de Prompt Injection & LGPD\n"
                "• `!dinamica mcp` → O Teatro do Protocolo MCP & Ciclo ReAct\n"
                "• `!dinamica rag` → Busca Vetorial vs Similaridade Humana\n"
                "• `!dinamica guardrails` → A Linha de Defesa Determinística"
            )
        enviar_zap(chat_id, resposta)
        return

    # !gabarito <id_atividade> (Consulta rápida do gabarito comentado)
    elif cmd_lower.startswith(("!gabarito", "/gabarito")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        m_id = re.search(r"\b(\d+)\b", resto)
        if m_id:
            id_atv = int(m_id.group(1))
            try:
                with sqlite3.connect(ESCOLA_DB) as conn:
                    for col_def in ("caderno_json TEXT DEFAULT ''", "gabarito_json TEXT DEFAULT ''"):
                        try:
                            conn.execute(f"ALTER TABLE atividades_cadastradas ADD COLUMN {col_def}")
                        except Exception:
                            pass
                    row = conn.execute("SELECT id, turma, componente, titulo, COALESCE(gabarito_json, '') FROM atividades_cadastradas WHERE id = ?", (id_atv,)).fetchone()
                if row:
                    _, turma, comp, titulo, gab_json = row
                    itens_gab = []
                    if gab_json:
                        try:
                            parsed_gab = json.loads(gab_json)
                            if isinstance(parsed_gab, list) and parsed_gab:
                                for g in parsed_gab:
                                    itens_gab.append(f"• *Q{g.get('numero')}:* {g.get('resposta')} — _{g.get('explicacao')}_")
                        except Exception:
                            pass
                    if itens_gab:
                        corpo_gab = "\n".join(itens_gab)
                        resposta = (
                            f"📝 *GABARITO COMENTADO — ATV ID {id_atv}*\n"
                            f"*{comp}* | Turma {turma}\n_{titulo}_\n\n"
                            f"{corpo_gab}\n\n"
                            f"💡 Para dar visto: `atv {id_atv} ok`"
                        )
                    else:
                        resposta = (
                            f"📝 *GABARITO COMENTADO — ATV ID {id_atv}*\n"
                            f"*{comp}* | Turma {turma}\n_{titulo}_\n\n"
                            "• *Q1:* Alternativa C (ou B) — Identificação correta de conceitos e arquitetura ReAct/Injeção.\n"
                            "• *Q2:* Alternativa B — Risco de Unauthorized Tool Execution via MCP e violação do Art. 46 da LGPD.\n"
                            "• *Q3:* Alternativa B (ou C) — Atuação de Guardrails como barreira determinística / Reranking com Cross-Encoder.\n"
                            "• *Q4 (Estudo de Caso):* Responsabilidade objetiva (CDC Art. 14 / CC 927) + exigência de Human-in-the-Loop.\n"
                            "• *Q5 (Estudo de Caso):* Necessidade de ACLs e controle de acesso estrito na base vetorial RAG para evitar vazamento de dados.\n\n"
                            f"💡 Para dar visto: `atv {id_atv} ok`"
                        )
                else:
                    resposta = f"Atividade ID {id_atv} não encontrada no banco escola.db."
            except Exception as e_gab:
                resposta = f"Erro ao consultar gabarito: {e_gab}"
        else:
            resposta = "Informe o ID da atividade. Ex: `!gabarito 6`"
        enviar_zap(chat_id, resposta)
        return

    # !turmas / !turma / !chamada / !alunos
    elif cmd_lower.startswith(("!turmas", "/turmas")):
        resposta = listar_turmas()
    elif cmd_lower.startswith(("!turma", "/turma", "!chamada", "/chamada", "!alunos", "/alunos")):
        termo = texto.split(" ", 1)[1].strip() if " " in texto else ""
        resposta = listar_alunos_turma(termo)

    # !faltas <turma> <n1, n2, n3...>
    elif cmd_lower.startswith(("!faltas", "/faltas")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        ov_pub, resto_limpo = _extrair_flag_visibilidade_cmd(resto)
        resposta = alterar_frequencia_lote(resto_limpo, "falta")
        m_t = re.search(r"TURMA\s+(\d{3})", resposta)
        if m_t and "registro(s) atualizado(s)" in resposta:
            t_cod = m_t.group(1)
            itens_lan = [l for l in resposta.splitlines() if l.strip().startswith("• #")]
            res_grp = gerar_resumo_publico_frequencia_sala(t_cod, "\n".join(itens_lan))
            if chat_id != GRUPOS_TURMAS.get(t_cod):
                resposta += despachar_resumo_grupo_se_publico(t_cod, res_grp, ov_pub)

    # !presencas <turma> <n1, n2, n3...>
    elif cmd_lower.startswith(("!presencas", "/presencas")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        ov_pub, resto_limpo = _extrair_flag_visibilidade_cmd(resto)
        resposta = alterar_frequencia_lote(resto_limpo, "presenca")
        m_t = re.search(r"TURMA\s+(\d{3})", resposta)
        if m_t and "registro(s) atualizado(s)" in resposta:
            t_cod = m_t.group(1)
            itens_lan = [l for l in resposta.splitlines() if l.strip().startswith("• #")]
            res_grp = gerar_resumo_publico_frequencia_sala(t_cod, "\n".join(itens_lan))
            if chat_id != GRUPOS_TURMAS.get(t_cod):
                resposta += despachar_resumo_grupo_se_publico(t_cod, res_grp, ov_pub)

    # !falta [turma] <Aluno ou Nº> [data]
    elif cmd_lower.startswith(("!falta", "/falta")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        ov_pub, resto = _extrair_flag_visibilidade_cmd(resto)
        if not resto:
            resposta = "Use: `!falta [turma] <Aluno ou Nº> [data]`\nEx: `!falta 202 6` ou `!falta Carlos Daniel` ou `!faltas 202 2, 6, 12`"
        elif "," in resto or (len(resto.split()) > 2 and resto.split()[0] in ["200", "201", "202", "203"] and resto.split()[1].isdigit() and resto.split()[2].isdigit()):
            resposta = alterar_frequencia_lote(resto, "falta")
            m_t = re.search(r"TURMA\s+(\d{3})", resposta)
            if m_t and "registro(s) atualizado(s)" in resposta:
                t_cod = m_t.group(1)
                itens_lan = [l for l in resposta.splitlines() if l.strip().startswith("• #")]
                res_grp = gerar_resumo_publico_frequencia_sala(t_cod, "\n".join(itens_lan))
                if chat_id != GRUPOS_TURMAS.get(t_cod):
                    resposta += despachar_resumo_grupo_se_publico(t_cod, res_grp, ov_pub)
        else:
            m_data = re.search(r"\b(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s*$", resto)
            if m_data:
                data_param = m_data.group(1)
                corpo = resto[:m_data.start()].strip()
            else:
                data_param = ""
                corpo = resto
            
            partes = corpo.split(maxsplit=1)
            if partes and partes[0] in ["200", "201", "202", "203"] and len(partes) > 1:
                turma_p = partes[0]
                aluno_p = partes[1]
            else:
                turma_p = TURMA_POR_GRUPO.get(chat_id, "")
                aluno_p = corpo
            resposta = alterar_frequencia_aluno(aluno_p, turma_p, data_param, "falta")
            m_t = re.search(r"Turma:\s*\*(\d{3})\*", resposta)
            m_al = re.search(r"Aluno:\s*\*(.*?)\*", resposta)
            m_tot = re.search(r"Total de Faltas:\s*\*(\d+)\*", resposta)
            if m_t and m_al and "Frequência Atualizada" in resposta:
                t_cod = m_t.group(1)
                res_acao = f"• *{m_al.group(1)}* — Falta registrada (Acumulado: {m_tot.group(1) if m_tot else '?'} falta(s))"
                res_grp = gerar_resumo_publico_frequencia_sala(t_cod, res_acao)
                if chat_id != GRUPOS_TURMAS.get(t_cod):
                    resposta += despachar_resumo_grupo_se_publico(t_cod, res_grp, ov_pub)

    # !presenca [turma] <Aluno ou Nº> [data]
    elif cmd_lower.startswith(("!presenca", "/presenca")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        ov_pub, resto = _extrair_flag_visibilidade_cmd(resto)
        if not resto:
            resposta = "Use: `!presenca [turma] <Aluno ou Nº> [data]`\nEx: `!presenca 202 6` ou `!presenca Carlos Daniel`"
        elif "," in resto or (len(resto.split()) > 2 and resto.split()[0] in ["200", "201", "202", "203"] and resto.split()[1].isdigit() and resto.split()[2].isdigit()):
            resposta = alterar_frequencia_lote(resto, "presenca")
            m_t = re.search(r"TURMA\s+(\d{3})", resposta)
            if m_t and "registro(s) atualizado(s)" in resposta:
                t_cod = m_t.group(1)
                itens_lan = [l for l in resposta.splitlines() if l.strip().startswith("• #")]
                res_grp = gerar_resumo_publico_frequencia_sala(t_cod, "\n".join(itens_lan))
                if chat_id != GRUPOS_TURMAS.get(t_cod):
                    resposta += despachar_resumo_grupo_se_publico(t_cod, res_grp, ov_pub)
        else:
            m_data = re.search(r"\b(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s*$", resto)
            if m_data:
                data_param = m_data.group(1)
                corpo = resto[:m_data.start()].strip()
            else:
                data_param = ""
                corpo = resto
            
            partes = corpo.split(maxsplit=1)
            if partes and partes[0] in ["200", "201", "202", "203"] and len(partes) > 1:
                turma_p = partes[0]
                aluno_p = partes[1]
            else:
                turma_p = TURMA_POR_GRUPO.get(chat_id, "")
                aluno_p = corpo
            resposta = alterar_frequencia_aluno(aluno_p, turma_p, data_param, "presenca")
            m_t = re.search(r"Turma:\s*\*(\d{3})\*", resposta)
            m_al = re.search(r"Aluno:\s*\*(.*?)\*", resposta)
            m_tot = re.search(r"Total de Faltas:\s*\*(\d+)\*", resposta)
            if m_t and m_al and "Frequência Atualizada" in resposta:
                t_cod = m_t.group(1)
                res_acao = f"• *{m_al.group(1)}* — Presença/Abono registrado (Acumulado: {m_tot.group(1) if m_tot else '0'} falta(s))"
                res_grp = gerar_resumo_publico_frequencia_sala(t_cod, res_acao)
                if chat_id != GRUPOS_TURMAS.get(t_cod):
                    resposta += despachar_resumo_grupo_se_publico(t_cod, res_grp, ov_pub)

    # !diario [Turma] <Conteúdo> (Entende automaticamente a sala/disciplina de hoje se omitido)
    elif cmd_lower.startswith(("!diario", "/diario")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        resposta = registrar_diario(resto)

    # !aluno <Nome> (Dossiê 360°)
    elif cmd_lower.startswith(("!aluno", "/aluno")):
        nome = texto.split(" ", 1)[1] if " " in texto else ""
        resposta = obter_dossie_aluno(nome) if nome else "Use: `!aluno <Nome>`"

    # !onde / onde eu to / qual minha sala / !sala
    elif cmd_lower.startswith(("!onde", "/onde", "onde", "onde eu to", "onde eu estou", "qual minha sala", "qual minha aula", "minha aula agora", "!ondeestou")):
        resposta = consultar_onde_estou()
    elif cmd_lower.startswith(("!sala", "/sala")):
        sala = texto.split(" ", 1)[1].strip() if " " in texto else ""
        resposta = buscar_multicriterio(f"sala:{sala}") if sala else consultar_onde_estou()

    # !horario / !agenda
    elif cmd_lower.startswith(("!horario", "/horario", "!agenda", "/agenda")):
        termo = texto.split(" ", 1)[1] if " " in texto else ""
        if prof_chave != "pedro" and prof_chave in PROFS_BASE_MAP:
            resposta = gestor_docentes_rag.consultar_grade_individual_docente(prof_chave, termo)
        else:
            resposta = consultar_horario(termo)

    # !grade / !semana
    elif cmd_lower.startswith(("!grade", "/grade", "!semana", "/semana")):
        termo = texto.split(" ", 1)[1] if " " in texto else ""
        if prof_chave != "pedro" and prof_chave in PROFS_BASE_MAP:
            resposta = gestor_docentes_rag.consultar_grade_individual_docente(prof_chave, termo)
        else:
            resposta = listar_grade_semanal()

    # !adicionaraula / !novaaula
    elif cmd_lower.startswith(("!adicionaraula", "/adicionaraula", "!novaaula", "/novaaula")):
        partes = texto.split(" ")
        if len(partes) < 6:
            resposta = "Use: `!adicionaraula <dia> <inicio> <fim> <turma> <disciplina> [sala]`\nEx: `!adicionaraula quinta 13:30 15:10 202 'INTELIGÊNCIA ARTIFICIAL' Sala202`"
        else:
            dia_a = partes[1]
            h_ini = partes[2]
            h_fim = partes[3]
            turma = partes[4]
            disc = partes[5]
            sala = partes[6] if len(partes) > 6 else "Sala 202"
            resposta = adicionar_aula_grade(dia_a, h_ini, h_fim, turma, disc, sala)

    # !removeraula / !excluiraula
    elif cmd_lower.startswith(("!removeraula", "/removeraula", "!excluiraula", "/excluiraula")):
        partes = texto.split(" ", 1)
        if len(partes) < 2:
            resposta = "Use: `!removeraula <id_aula>` (veja os IDs com `!grade`)"
        else:
            resposta = remover_aula_grade(partes[1].strip())

    # !alarmes / !alarme / !aviso / !lembrete
    elif cmd_lower.startswith(("!alarmes", "/alarmes")):
        resposta = listar_alarmes()
    elif cmd_lower.startswith(("!alarme", "/alarme", "!aviso", "/aviso", "!lembrete", "/lembrete")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        resposta = adicionar_alarme(resto, chat_id)

    # !resumo / !matinal / !dia / !modo (com fronteira de palavra estrita para evitar colisão com 'modos de comandos', 'dia bom', etc.)
    elif re.match(r"^(?:[!/](?:resumo|matinal|dia|modo)\b|(?:resumo|matinal)\b|modo(?:\s+(?:20[0-3]|base|ti|lab|laboratorio|publico|privado))*\s*$)", cmd_lower):
        resposta = processar_comando_resumo_ou_modo(texto, chat_id)

    # !aula_link / !video / !esmiucar (Ideia #12)
    elif cmd_lower.startswith(("!aula_link", "/aula_link", "!video", "/video", "!esmiucar", "/esmiucar")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        if not resto:
            resposta = (
                "🎬 *GERADOR DE AULA A PARTIR DE LINK (Ideia #12)*\n\n"
                "Envie o link do YouTube ou artigo técnico para esmiuçar em aula de 50min:\n"
                "• `!aula_link <url> [turma]`\n"
                "• `!video <url>`"
            )
        else:
            partes = resto.split(" ", 1)
            url_alvo = partes[0]
            turma_alvo = partes[1] if len(partes) > 1 else "Técnico em Informática"
            enviar_zap(chat_id, "⏳ Analisando conteúdo do link e esmiuçando para aula técnica de 50min...")
            try:
                import gerador_aula_video
                res_aula = gerador_aula_video.processar_link_aula(url_alvo, turma_alvo)
                resposta = res_aula["msg_zap"]
                if res_aula.get("caminho_pptx") and os.path.exists(res_aula["caminho_pptx"]):
                    enviar_documento_zap(chat_id, res_aula["caminho_pptx"], os.path.basename(res_aula["caminho_pptx"]), "application/vnd.openxmlformats-officedocument.presentationml.presentation")
            except Exception as e_a:
                resposta = f"⚠️ Erro ao processar aula do link: {e_a}"

    # !gabarito_lab / !folha_lab (Ideia #15)
    elif cmd_lower.startswith(("!gabarito_lab", "/gabarito_lab", "!folha_lab", "/folha_lab", "!gabarito", "/gabarito")):
        enviar_zap(chat_id, "📄 Compilando grade técnica e reservas da semana para o gabarito...")
        try:
            import gerar_gabarito_lab
            res_gab = gerar_gabarito_lab.gerar_gabarito_semanal()
            resposta = (
                f"📋 *GABARITO DE LABORATÓRIO GERADO!* (Semana: {res_gab['periodo']})\n\n"
                f"• Grade fixa e tempos protegidos compilados.\n"
                f"• Espaços livres diagramados para assinatura manual.\n"
                f"• Arquivo Markdown: `{os.path.basename(res_gab['caminho_md'])}`\n\n"
                f"💡 Após colher assinaturas na porta, lance no bot com:\n"
                f"`!papel <dia> <tempos> <docente>`"
            )
            if res_gab.get("caminho_pdf") and os.path.exists(res_gab["caminho_pdf"]):
                enviar_documento_zap(chat_id, res_gab["caminho_pdf"], os.path.basename(res_gab["caminho_pdf"]), "application/pdf")
        except Exception as e_g:
            resposta = f"⚠️ Erro ao gerar gabarito: {e_g}"

    # !papel / !sincronizar_papel (Ideia #15)
    elif cmd_lower.startswith(("!papel", "/papel", "!sincronizar_papel", "/sincronizar_papel")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        partes = resto.split()
        if len(partes) < 3:
            resposta = (
                "📝 *SINCRONIZAR RESERVA DO PAPEL (Ideia #15)*\n\n"
                "Use: `!papel <dia> <horarios> <docente> [motivo]`\n"
                "Exemplos:\n"
                "• `!papel terca 3,4 Prof. Carlos Prática de Redes`\n"
                "• `!papel quinta 8 Prof. Marcos Reposição`"
            )
        else:
            dia_p = partes[0]
            tempos_p = partes[1]
            resto_p = " ".join(partes[2:])
            motivo_p = "Origem: Gabarito de Papel da Porta"
            prof_p = resto_p
            try:
                import gerar_gabarito_lab
                res_p = gerar_gabarito_lab.sincronizar_reserva_papel(dia_p, tempos_p, prof_p, motivo_p)
                resposta = res_p["mensagem"]
            except Exception as e_p:
                resposta = f"⚠️ Erro ao sincronizar papel: {e_p}"

    # !limpeza_segunda / !limpeza (Ideia #14)
    elif cmd_lower.startswith(("!limpeza_segunda", "/limpeza_segunda", "!limpeza", "/limpeza")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        if not resto:
            modo_atual = obter_modo_limpeza_segunda()
            resposta = (
                f"🧹 *CONFIGURAÇÃO DE LIMPEZA DE SEGUNDA-FEIRA (Ideia #14)*\n\n"
                f"• Modo atual: *{modo_atual.upper()}*\n\n"
                f"Para alterar o modo de bloqueio:\n"
                f"• `!limpeza padrao` -> Bloqueia 1º e 2º horários (07:30 às 09:10)\n"
                f"• `!limpeza manha` -> Bloqueia manhã toda (1º ao 5º horários)\n"
                f"• `!limpeza integral` -> Bloqueia todo o expediente de segunda-feira"
            )
        else:
            modo_solicitado = resto.strip().lower()
            if modo_solicitado in ("padrao", "manha", "integral", "dia_todo"):
                definir_modo_limpeza_segunda(modo_solicitado)
                resposta = f"✅ Modo de limpeza de segunda-feira atualizado com sucesso para *{modo_solicitado.upper()}*."
            else:
                resposta = "⚠️ Modo inválido. Escolha: `padrao`, `manha` ou `integral`."

    # !cobrar / cobrar / !pendencias / pendencias / !pendentes / pendentes
    elif re.match(r"^[!/]?(?:cobrar|pendencias?|pendentes?)\b", cmd_lower):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        ov_pub, resto_limpo = _extrair_flag_visibilidade_cmd(resto)
        if not resto_limpo and chat_id in TURMA_POR_GRUPO:
            resto_limpo = TURMA_POR_GRUPO[chat_id]
        resposta = cobrar_atividades_turma(resto_limpo, prof_chave=prof_chave)
        m_t = re.search(r"TURMA\s+(\d{3})", resposta)
        m_id = re.search(r"\*#(\d+)\s+—", resposta)
        if m_t and chat_id != GRUPOS_TURMAS.get(m_t.group(1)):
            t_cod = m_t.group(1)
            aid_c = int(m_id.group(1)) if m_id else None
            res_grp = gerar_resumo_publico_atividade_sala(t_cod, aid_c)
            resposta += despachar_resumo_grupo_se_publico(t_cod, res_grp, ov_pub)

    # !passar / passar
    elif re.match(r"^[!/]?passar\b", cmd_lower):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        ov_pub, resto_limpo = _extrair_flag_visibilidade_cmd(resto)
        resposta = cadastrar_nova_atividade(resto_limpo, prof_chave=prof_chave)
        m_t = re.search(r"\*Turma:\*\s*(\d{3})", resposta)
        m_id = re.search(r"ID:\s*#(\d+)", resposta)
        if m_t and "CADASTRADA COM SUCESSO" in resposta and chat_id != GRUPOS_TURMAS.get(m_t.group(1)):
            t_cod = m_t.group(1)
            aid_c = int(m_id.group(1)) if m_id else None
            res_grp = gerar_resumo_publico_atividade_sala(t_cod, aid_c, "📌 *Nova atividade cadastrada pelo professor!*")
            resposta += despachar_resumo_grupo_se_publico(t_cod, res_grp, ov_pub)

    # !atv / !visto / !atividade / atv / visto / atividade / corrigir / correcao
    elif re.match(r"^[!/]?(?:atv|vistos?|atividades?|corrigir|correcao)\b", cmd_lower):
        ov_pub, texto_limpo = _extrair_flag_visibilidade_cmd(texto)
        resposta = processar_comando_atividade(texto_limpo, chat_id, data)
        # Se houve lançamento efetivo de visto ou cadastro de atividade, envia o resumo limpo no grupo da turma (se público)
        foi_lancamento = any(marc in resposta for marc in (
            "VISTO EM MASSA REGISTRADO",
            "LOTE DE ATIVIDADES REGISTRADO",
            "VISTO / ATIVIDADE REGISTRADA",
            "VISTO / CORREÇÃO POR FOTO REGISTRADA",
            "NOVA ATIVIDADE CADASTRADA COM SUCESSO",
            "VISTOS MOVIDOS COM SUCESSO",
            "VISTOS COPIADOS COM SUCESSO",
        ))
        if foi_lancamento:
            m_t = re.search(r"(?:TURMA|\*Turma:\*|Para:.*Turma)\s*(\d{3})", resposta)
            if not m_t:
                for tc in ("200", "201", "202", "203"):
                    if f"Turma {tc}" in resposta or f"TURMA {tc}" in resposta:
                        m_t_val = tc
                        break
                else:
                    m_t_val = None
            else:
                m_t_val = m_t.group(1)
            m_id = re.search(r"#(\d+)\s+—", resposta) or re.search(r"ID:\s*#(\d+)", resposta) or re.search(r"Para:\*\s*#(\d+)", resposta)
            aid_c = int(m_id.group(1)) if m_id else None
            if m_t_val and chat_id != GRUPOS_TURMAS.get(m_t_val):
                detalhe_acao = ""
                m_al_u = re.search(r"\*Aluno:\*\s*(.*)", resposta)
                if m_al_u:
                    detalhe_acao = f"✅ *Visto atualizado:* {m_al_u.group(1).strip()}"
                elif "VISTO EM MASSA REGISTRADO" in resposta:
                    detalhe_acao = "✅ *Atualização de vistos em massa realizada pelo professor.*"
                elif "LOTE DE ATIVIDADES REGISTRADO" in resposta:
                    detalhe_acao = "✅ *Lote de vistos atualizado pelo professor.*"
                elif "NOVA ATIVIDADE CADASTRADA" in resposta:
                    detalhe_acao = "📌 *Nova atividade cadastrada pelo professor!*"
                res_grp = gerar_resumo_publico_atividade_sala(m_t_val, aid_c, detalhe_acao)
                resposta += despachar_resumo_grupo_se_publico(m_t_val, res_grp, ov_pub)

    # !caderno / caderno
    elif re.match(r"^[!/]?caderno\b", cmd_lower):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        resposta = obter_exercicio_caderno_aula(resto)

    # !executar / !rodar / !run
    elif cmd_lower.startswith(("!executar", "/executar", "!rodar", "/rodar", "!run", "/run")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        threading.Thread(target=executar_ideia_comando, args=(resto, chat_id)).start()
        return

    # !ideias / !ideia / !notas / !nota
    elif cmd_lower.startswith(("!ideias", "/ideias", "!ideia", "/ideia", "!notas", "/notas", "!nota", "/nota")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        if not resto or resto.lower() in ["listar", "lista", "ver", "todas"]:
            resposta = listar_ideias_notas()
        elif resto.lower() in ["confirmar limpar tudo", "confirmar apagar tudo", "limpar tudo confirmar"]:
            resposta = limpar_todas_ideias()
        elif resto.lower() in ["limpar tudo", "apagar tudo", "remover tudo", "zerar", "limpar todos", "apagar todos"]:
            resposta = (
                "⚠️ *CONFIRMAÇÃO DE SEGURANÇA:*\n\n"
                "Você solicitou apagar *TODAS* as ideias e arquivos salvos.\n"
                "Para confirmar com segurança, envie:\n"
                "• `!ideias confirmar limpar tudo`\n"
                "Ou envie `!cancelar`."
            )
        elif resto.lower().startswith(("limpar", "apagar", "remover", "delete")):
            partes = resto.split()
            if len(partes) > 1 and (partes[1].isdigit() or partes[1] not in ["tudo", "todos"]):
                resposta = apagar_ideia_especifica(partes[1])
            elif len(partes) == 1:
                resposta = (
                    "ℹ️ *Como limpar ideias:*\n"
                    "• `!ideias limpar <id>` -> Apagar uma ideia específica (ex: `!ideias limpar 1`)\n"
                    "• `!ideias limpar tudo` -> Apagar todas as ideias e arquivos antigos"
                )
            else:
                resposta = limpar_todas_ideias()
        elif resto.isdigit() or (resto.lower().startswith("ver ") and len(resto.split()) > 1 and resto.split()[1].isdigit()):
            id_alvo = resto if resto.isdigit() else resto.split()[1]
            resposta = obter_detalhes_ideia(id_alvo)
        else:
            enviar_zap(chat_id, f"💡 *Ideia Recebida:* \"{resto}\"\n\n⏳ Estruturando proposta prática e inteligível sem comandos de terminal ou códigos para sua rotina docente...")
            threading.Thread(target=worker_processar_ideia, args=(chat_id, resto)).start()
            return

    # !cloud <comando ou solicitacao>
    elif cmd_lower.startswith(("!cloud", "/cloud")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        threading.Thread(target=executar_cloud_cli, args=(resto, chat_id)).start()
        return

    # !away [on|off|status]
    elif cmd_lower.startswith(("!away", "/away", "!ausente", "/ausente")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        resposta = alternar_modo_away(resto)

    # !metricas / !telemetria
    elif cmd_lower.startswith(("!metricas", "/metricas", "!telemetria", "/telemetria")):
        resposta = obter_resumo_metricas_str()

    # !status / !ssh / !rede
    elif cmd_lower.startswith(("!status", "/status", "!ssh", "/ssh", "!rede", "/rede")):
        resposta = consultar_status_sistema()

    # !aulas / !conteudo / aula <turma> <tema>
    elif cmd_lower.startswith(("!aula", "/aula", "aula ", "!conteudo", "/conteudo")):
        termo = texto.split(" ", 1)[1].strip() if " " in texto else ""
        m_turma_tema = re.match(r"^(20[0-3])\s+(.+)$", termo)
        if m_turma_tema:
            turma_alvo, tema_alvo = m_turma_tema.group(1), m_turma_tema.group(2).strip()
            enviar_zap(chat_id, f"⏳ *Preparando conteúdo '{tema_alvo}' para a Turma {turma_alvo}*...\nGerando slides (.pptx), lista (.pdf), exercícios de caderno e atividade no banco!")
            def _worker_aula_tema(t_cod, t_tema, c_id):
                try:
                    import gerador_aulas_rotina
                    gerador_aulas_rotina.aplicar_conteudo_turma(t_cod, t_tema, termo_dia="hoje", enviar_whatsapp=True, chat_destino=c_id)
                except Exception as e_ap:
                    enviar_zap(c_id, f"❌ Erro ao aplicar tema '{t_tema}' na turma {t_cod}: {e_ap}")
            threading.Thread(target=_worker_aula_tema, args=(turma_alvo, tema_alvo, chat_id)).start()
            return
        resposta = consultar_aulas(termo)

    # !lab / /lab / lab (Agendamento Laboratorio Base Tecnica)
    elif _eh_comando_lab_msg(cmd_lower):
        data_lab = dict(data) if isinstance(data, dict) else {}
        data_lab.setdefault("chatId", chat_id)
        resposta = processar_comando_lab(texto, data_lab)

    # !sou (Identificacao de docente)
    elif cmd_lower.startswith(("!sou", "/sou")):
        resto = texto.split(" ", 1)[1].strip() if " " in texto else ""
        author_raw = data.get("author") or (data.get("sender") or {}).get("id") or ""
        author_num = extrair_identificador_limpo(author_raw)
        nome_pesquisa = resto.lower()
        achou = False
        for chave, nome_exib in PROFS_BASE_MAP.items():
            if chave in nome_pesquisa:
                salvar_vinculo_professor(author_num, chave, nome_exib)
                resposta = f"Docente identificado com sucesso: {nome_exib}"
                achou = True
                break
        if not achou and resto:
            nome_livre = f"Prof. {resto.strip()}"
            salvar_vinculo_professor(author_num, "outro", nome_livre)
            resposta = f"Docente identificado com sucesso: {nome_livre}"
        elif not achou:
            resposta = "Informe seu nome (ex: !sou Marcos)."

    # !geolocalizar / ache a localizacao / onde foi tirada
    elif any(cmd_lower.startswith(k) for k in [
        "ache a localizacao", "ache a localização", "onde foi tirada",
        "qual a localizacao", "qual a localização", "!geolocalizar", "/geolocalizar",
        "!localizacao", "/localizacao", "!geo", "/geo"
    ]):
        midia_path = obter_midia_recente(data, chat_id)
        if not midia_path:
            enviar_zap(chat_id, "Envie ou responda a uma foto com a mensagem: 'ache a localizacao' (Dica: envie como documento para preservar metadados GPS completos).")
            return
        threading.Thread(target=worker_geolocalizar_imagem, args=(chat_id, midia_path)).start()
        return

    # !exif (Extração forense de metadados de foto/documento)
    elif cmd_lower.startswith(("!exif", "/exif")):
        midia_path = obter_midia_recente(data, chat_id)
        if not midia_path:
            enviar_zap(chat_id, "Envie ou responda a uma foto ou documento com o comando `!exif` para extrair os metadados.")
            return
        threading.Thread(target=worker_exif_completo, args=(chat_id, midia_path)).start()
        return

    # !osint nick <usuario> | !osint domain <dominio> | !osint email <email> | !osint phone <numero> | !osint insta <user> | !osint full <alvo>
    elif cmd_lower.startswith(("!osint", "/osint")):
        partes = texto.strip().split(maxsplit=2)
        sub = partes[1].lower() if len(partes) > 1 else ""
        alvo = partes[2].strip() if len(partes) > 2 else ""

        if sub in ["full", "auto", "tudo", "orquestrar", "rastrear", "investigar"] and alvo:
            threading.Thread(target=worker_osint_orquestrador, args=(chat_id, alvo)).start()
            return
        elif sub in ["nick", "user", "username", "perfil"] and alvo:
            threading.Thread(target=worker_osint_nick, args=(chat_id, alvo)).start()
            return
        elif sub in ["domain", "dominio", "host", "site"] and alvo:
            threading.Thread(target=worker_osint_domain, args=(chat_id, alvo)).start()
            return
        elif sub in ["email", "mail", "e-mail"] and alvo:
            threading.Thread(target=worker_osint_email, args=(chat_id, alvo)).start()
            return
        elif sub in ["phone", "tel", "telefone", "celular", "numero", "fone"] and alvo:
            threading.Thread(target=worker_osint_phone, args=(chat_id, alvo)).start()
            return
        elif sub in ["instabusca", "igbusca", "buscainsta"] and alvo:
            threading.Thread(target=worker_osint_insta_publico, args=(chat_id, alvo)).start()
            return
        elif sub in ["insta", "instagram", "ig"] and alvo:
            threading.Thread(target=worker_osint_insta, args=(chat_id, alvo)).start()
            return
        elif "@" in sub and not alvo:
            threading.Thread(target=worker_osint_email, args=(chat_id, sub)).start()
            return
        elif re.sub(r"\D", "", sub) and len(re.sub(r"\D", "", sub)) >= 8 and not alvo:
            threading.Thread(target=worker_osint_phone, args=(chat_id, sub)).start()
            return
        else:
            enviar_zap(
                chat_id,
                "🔎 *Comandos OSINT Disponíveis:*\n\n"
                "• `!osint full <email ou celular>` -> 🔄 *Orquestração Total* (Cruzamento E-mail ↔ Celular ↔ Redes)\n"
                "• `!osint email <email>` -> Checa plataformas vinculadas (Holehe) + cruzamento\n"
                "• `!osint phone <numero>` -> Checa contas vinculadas (Ignorant) + correlação\n"
                "• `!osint nick <usuario>` -> Busca perfis em centenas de redes (Sherlock)\n"
                "• `!osint insta <usuario> [session_id]` -> Dossiê de perfil Instagram (Toutatis)\n"
                "• `!osint instabusca <nome ou nick>` -> Acha Instagram sem login (OG + busca web)\n"
                "• `!osint domain <dominio>` -> Dossiê de IPs, MX, NS e WHOIS\n"
                "• `!exif` (com foto/doc) -> Extrai metadados forenses\n"
                "• Foto + `ache a localizacao` -> Rastreia coordenadas GPS e Google Maps (Geo-Sleuth)"
            )
            return

    # !ajuda completa
    elif cmd_lower.startswith(("!ajuda completa", "/ajuda completa", "!help full")):
        resposta = (
            "*Menu Completo do Assistente:*\n\n"
            "OSINT & Geolocalização:\n"
            "• `!osint full <alvo>` -> Orquestração total (E-mail ↔ Celular ↔ Redes)\n"
            "• Envie foto + `ache a localizacao` -> Rastreia GPS e Google Maps (Geo-Sleuth)\n"
            "• `!exif` (com foto/doc) -> Dossiê de metadados forenses\n"
            "• `!osint email <email>` -> Plataformas vinculadas (Holehe) + cruzamento\n"
            "• `!osint phone <numero>` -> Serviços vinculados ao número (Ignorant)\n"
            "• `!osint nick <usuario>` -> Busca de perfis (Sherlock)\n"
            "• `!osint insta <usuario>` -> Dossiê Instagram (Toutatis)\n"
            "• `!osint domain <dominio>` -> Dossiê de rede e WHOIS\n\n"
            "Laboratório e Grade:\n"
            "• `!lab [hoje|amanha|dd/mm]` -> Consulta grade do laboratório\n"
            "• `!lab agendar [data] <tempos> [motivo]` -> Reserva de horários\n"
            "• `!lab cancelar [data] <tempos|ID>` -> Libera reserva\n"
            "• `!horario [hoje|amanha|dd/mm]` | `!grade`\n\n"
            "Turmas e Pedagógico:\n"
            "• `!turmas` | `!turma <turma>` -> Lista alunos e faltas\n"
            "• `!falta [turma] <nº>` | `!presenca [turma] <nº>`\n"
            "• `!diario <Turma> <Conteudo>` -> Registro de diário\n"
            "• `!cobrar <turma>` -> Pendências e vistos\n"
            "• `!atv nova <turma> <IA> <titulo>` -> Cadastra atividade\n"
            "• `!atv <nº/nome> <turma> ok` -> Registra visto\n"
            "• `!preparar [amanha|hoje]` -> Gera aulas, slides, PDFs e visto no Zap\n"
            "• `!exercicio [tema]` -> Cria lista BNCC\n"
            "• `!slides <Tema>` -> Slides (.pptx)\n"
            "• `!ideias <ideia>` | `!executar <id>` -> Protótipo multiagente\n"
            "• `!cloud <comando>` -> Terminal Linux\n"
            "• `!status` -> Diagnóstico do sistema"
        )

    # !ajuda / !comandos (Menu Simplificado e Objetivo)
    elif re.match(r"^(?:[!/](?:ajuda|help|start|menu|comandos?)|(?:ajuda|help|menu))\b", cmd_lower):
        resposta = (
            "📋 *ASSISTENTE DOCENTE — MENU RÁPIDO*\n\n"
            "🏫 *Aulas & Horários:*\n"
            "• `!preparar [amanha|hoje]` → Prepara aulas, slides e listas\n"
            "• `!horario [amanha|hoje]` → Suas aulas do dia\n"
            "• `!lab [amanha|hoje]` → Horários do laboratório\n"
            "• `!lab agendar <tempos> <motivo>` → Reservar lab (ex: `!lab agendar 4,5 201 IA`)\n"
            "• `!lab cancelar <tempos|ID>` → Liberar reserva\n\n"
            "📝 *Turmas & Alunos:*\n"
            "• `!turma <200|201|202>` → Lista e faltas\n"
            "• `!falta <turma> <nº>` | `!presenca <turma> <nº>`\n"
            "• `!diario <turma> <conteúdo>` → Registrar aula\n"
            "• `!cobrar <turma>` → Vistos pendentes\n\n"
            "⚡ *Apoio & Ferramentas:*\n"
            "• `!alarme <tempo/hora> <aviso>` → Programar lembrete\n"
            "• `!exercicio [tema]` → Gerar lista de exercícios\n"
            "• `!status` → Diagnóstico dos serviços\n"
            "• `!ajuda completa` → Menu detalhado (OSINT, IA, Cloud)"
        )
    else:
        # Só emite aviso de comando não reconhecido se o usuário digitou explicitamente um comando curto com prefixo ! ou /
        # Textos avulsos, links, ecos ou mensagens comuns NUNCA devem receber resposta de fallback para evitar loops de eco.
        if not cmd_lower.startswith(("!", "/")) or _eh_eco_ou_mensagem_bot(texto) or len(texto) > 60:
            return
        resposta = f"Comando '{texto}' não reconhecido. Digite `!ajuda` para ver o menu."

    # Separação estrita: se o comando foi enviado em um grupo (@g.us), remove dicas de comando;
    # se for comando estritamente administrativo/privado, responde apenas no privado do professor.
    if "@g.us" in str(chat_id):
        eh_cmd_privado_estrito = cmd_lower.startswith((
            "!cloud", "/cloud", "!status", "/status", "!ssh", "!rede",
            "!alarme", "/alarme", "!aviso", "!lembrete",
            "!ideia", "/ideia", "!nota", "!executar", "!rodar", "!run",
            "!osint", "/osint", "!exif", "/exif", "!geo", "ache a localizacao",
            "!gabarito", "/gabarito", "!preparar", "/preparar", "!planejar",
            "!diario", "/diario", "!aluno", "/aluno", "!adicionaraula", "!removeraula"
        ))
        if chat_id in GRUPOS_OSINT_PERMITIDOS and _eh_comando_osint_ou_geo(cmd_lower):
            eh_cmd_privado_estrito = False
        if eh_cmd_privado_estrito:
            enviar_zap(CHAT_PRIVADO_PADRAO, resposta)
            return
        if chat_id in TURMA_POR_GRUPO:
            resposta = remover_dicas_comandos_para_grupo(resposta)

    if not resposta:
        return

    enviar_zap(chat_id, resposta)

if __name__ == "__main__":
    port = 5000
    init_tabelas_config()
    init_tabelas_ideias()
    init_tabelas_atividades()
    init_tabela_laboratorio()
    init_tabela_telemetria()
    try:
        limpar_arquivos_antigos_osint(max_idade_seg=300)
    except Exception:
        pass
    try:
        cron_ativo = bool(subprocess.check_output(["pgrep", "-f", "cron_daemon.py"]).decode().strip())
    except Exception:
        cron_ativo = False
    if not cron_ativo:
        threading.Thread(target=loop_verificador_alarmes, daemon=True).start()
    else:
        log_terminal("⏰ Cron daemon externo ativo. Thread interna de alarmes em repouso.")
    log_terminal(f"🚀 Webhook rodando na porta {port}. Comandos docentes, alarmes, ideias e Cloud CLI ativos.")
    HTTPServer(("0.0.0.0", port), WebhookHandler).serve_forever()
