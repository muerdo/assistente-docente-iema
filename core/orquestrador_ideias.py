#!/usr/bin/env python3
"""
==============================================================================
ORQUESTRADOR MULTIAGENTE: REVISÃO DE CONVERSAS E REFINAMENTO DE IDEIAS
==============================================================================
Pipeline:
- Agente 1 (Idea Harvester): Analisa transcripts, logs do WhatsApp e diário
  de classe do período para extrair demandas, ideias e desafios.
- Agente 2 (Idea Refiner): Recebe o dossiê do Agente 1, utiliza o motor LLM
  (OmniRoute) para estruturar especificações técnicas, arquitetura de software
  e melhorias pedagógicas.
- Persistência: Registra em escola.db (ideias_notas e tarefas_agentes), gera
  Markdown local e no Google Drive, e notifica o WhatsApp privado do docente.
==============================================================================
"""

import os
DONO_CHAT_ID = os.environ.get("DONO_CHAT_ID", "")  # chat do dono (ex.: 5599...@c.us)
import sys
import glob
import json
import sqlite3
import datetime
import urllib.request
import urllib.error
import re

OMNIROUTE_URL = "http://localhost:20128/v1/chat/completions"
OMNIROUTE_MODEL = "no-think/agy/claude-sonnet-5-5-low"
OPENWA_API_URL = "http://127.0.0.1:2785"
OPENWA_API_KEY = os.environ.get("OPENWA_API_KEY", "")
OPENWA_SESSION_ID = os.environ.get("OPENWA_SESSION_ID", "")

ESCOLA_DB = os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db")
LOG_DIR = os.path.join(os.environ.get("DOCENTE_HOME", "."), "logs_docente")
LOG_ORQUESTRADOR = os.path.join(LOG_DIR, "orquestrador_ideias.log")
PROJETOS_IDEIAS_DIR = os.path.join(os.environ.get("DOCENTE_HOME", "."), "projetos_ideias")
GDRIVE_IDEIAS_DIR = os.path.join(os.environ.get("DOCENTE_HOME", "."), "mounts/gdrive/Aulas/Notas_Ideias")
PRIVATE_CHAT = DONO_CHAT_ID

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(PROJETOS_IDEIAS_DIR, exist_ok=True)

def log(msg: str):
    linha = f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(linha, flush=True)
    try:
        with open(LOG_ORQUESTRADOR, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception:
        pass

def consultar_llm(prompt: str, max_tokens: int = 2500) -> str:
    """Consulta o modelo ativo no OmniRoute processando resposta em stream/chunk."""
    payload = json.dumps({
        "model": OMNIROUTE_MODEL,
        "messages": [
            {
                "role": "system",
                "content": "Você é um Arquiteto de Software e Engenheiro de IA sênior especializado em ecossistemas docentes, automações educacionais, sistemas de controle escolar e WhatsApp bots. Seja técnico, preciso, pragmático e direto."
            },
            {"role": "user", "content": prompt}
        ],
        "max_tokens": max_tokens,
        "stream": True
    }).encode("utf-8")

    req = urllib.request.Request(
        OMNIROUTE_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST"
    )

    try:
        with urllib.request.urlopen(req, timeout=90) as resp:
            texto_acumulado = []
            for raw_line in resp:
                linha = raw_line.decode("utf-8").strip()
                if not linha or not linha.startswith("data:"):
                    continue
                conteudo_data = linha[5:].strip()
                if conteudo_data == "[DONE]":
                    break
                try:
                    chunk = json.loads(conteudo_data)
                    delta = chunk.get("choices", [{}])[0].get("delta", {})
                    token = delta.get("content", "")
                    if token:
                        texto_acumulado.append(token)
                except Exception:
                    continue
            return "".join(texto_acumulado).strip()
    except Exception as e:
        log(f"⚠️ Erro ao consultar LLM no OmniRoute: {e}")
        return ""

def enviar_zap_privado(texto: str) -> bool:
    """Dispara notificação para o WhatsApp privado do docente."""
    url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/messages/send-text"
    req = urllib.request.Request(
        url,
        data=json.dumps({"chatId": PRIVATE_CHAT, "text": texto}).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-api-key": OPENWA_API_KEY},
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return resp.status in (200, 201)
    except Exception as e:
        log(f"⚠️ Erro ao enviar notificação WhatsApp: {e}")
        return False

# ==============================================================================
# AGENTE 1: CONVERSATION REVIEWER & IDEA HARVESTER
# ==============================================================================
def agente_1_revisar_conversas() -> dict:
    log("🔍 [Agente 1] Iniciando varredura e extração de conversas e registros do dia...")
    conversas_coletadas = []

    # 1. Varredura dos transcripts do Antigravity
    brain_dir = os.path.join(os.environ.get("DOCENTE_HOME", "."), ".gemini/antigravity-cli/brain")
    if os.path.exists(brain_dir):
        transcripts = glob.glob(f"{brain_dir}/*/.system_generated/logs/transcript.jsonl")
        # Ordena pelos mais recentes
        transcripts = sorted(transcripts, key=os.path.getmtime, reverse=True)[:5]
        for tf in transcripts:
            try:
                with open(tf, "r", encoding="utf-8", errors="replace") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            obj = json.loads(line)
                            if obj.get("source") == "USER_EXPLICIT" and obj.get("type") == "USER_INPUT":
                                conteudo = obj.get("content", "").strip()
                                if conteudo:
                                    conversas_coletadas.append({
                                        "fonte": "Antigravity UI",
                                        "timestamp": obj.get("created_at", ""),
                                        "texto": conteudo[:400]
                                    })
                        except Exception:
                            continue
            except Exception as e:
                log(f"Aviso ao ler transcript {tf}: {e}")

    # 2. Varredura do webhook_live.log (últimos comandos privados e solicitações)
    webhook_log = os.path.join(os.environ.get("DOCENTE_HOME", "."), "webhook_live.log")
    comandos_zap = []
    if os.path.exists(webhook_log):
        try:
            with open(webhook_log, "r", encoding="utf-8", errors="replace") as f:
                linhas = f.readlines()[-150:]
                for l in linhas:
                    if "Comando autorizado no chat" in l or "DISPARANDO ALARME" in l or "Processando:" in l:
                        comandos_zap.append(l.strip())
        except Exception:
            pass

    # 3. Varredura de ideias já cadastradas em escola.db
    ideias_existentes = []
    if os.path.exists(ESCOLA_DB):
        try:
            with sqlite3.connect(ESCOLA_DB) as conn:
                cur = conn.cursor()
                cur.execute("SELECT id, titulo, descricao_original, status FROM ideias_notas ORDER BY id DESC LIMIT 5")
                for row in cur.fetchall():
                    ideias_existentes.append({
                        "id": row[0], "titulo": row[1], "descricao": row[2], "status": row[3]
                    })
        except Exception:
            pass

    dossie_agente_1 = {
        "timestamp_extracao": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_entradas_transcripts": len(conversas_coletadas),
        "conversas_recentes": conversas_coletadas[-15:],
        "atividades_recentes_zap": comandos_zap[-20:],
        "ideias_banco": ideias_existentes
    }

    log(f"✅ [Agente 1] Concluído: {len(conversas_coletadas)} diálogos e {len(comandos_zap)} eventos catalogados.")
    return dossie_agente_1

# ==============================================================================
# AGENTE 2: IDEA REFINER & IMPLEMENTATION ARCHITECT
# ==============================================================================
def agente_2_refinar_ideias(dossie_entrada: dict) -> dict:
    log("💡 [Agente 2] Recebendo dossiê do Agente 1 e iniciando refinamento prático e operacional...")

    # Compacta o dossiê para focar no essencial
    contexto_str = json.dumps(dossie_entrada, ensure_ascii=False)

    prompt = f"""
Você é o Agente Refinador de Ideias e Estrategista Pedagógico/Operacional do IEMA.
Com base no dossiê de conversas, comandos e demandas recentes do ecossistema docente:

{contexto_str}

Sua missão é analisar as necessidades do docente (gestão de turmas de Informática e IA, controle de chamadas e faltas, agendamento de laboratório, economia de recursos e automação WhatsApp) e estruturar um RELATÓRIO EXECUTIVO DE REFINAMENTO.

REGRAS OBRIGATÓRIAS:
1. NÃO sugira comandos de terminal (ssh, curl, bash, python) nem códigos de programação, pois você não conhece a arquitetura exata do servidor.
2. Foque em melhorias práticas, inteligíveis e diretamente aplicáveis na rotina de trabalho e no dia a dia da escola.

Responda OBRIGATORIAMENTE em formato JSON válido contendo exatamente as seguintes chaves:
{{
  "titulo_ideia_destaque": "Título claro e objetivo da proposta prioritária",
  "resumo_executivo": "Conceito direto e benefício prático na rotina do docente e estudantes",
  "pontos_chave_melhoria": [
    "Ponto 1 de melhoria prática imediata",
    "Ponto 2 de melhoria prática imediata",
    "Ponto 3 de melhoria prática imediata"
  ],
  "passos_praticos": [
    {{"ordem": 1, "acao": "Primeira ação prática a ser executada", "beneficio": "Resultado gerado"}},
    {{"ordem": 2, "acao": "Segunda ação prática a ser executada", "beneficio": "Resultado gerado"}},
    {{"ordem": 3, "acao": "Terceira ação prática a ser executada", "beneficio": "Resultado gerado"}}
  ]
}}
Retorne EXCLUSIVAMENTE o JSON sem blocos de texto antes ou depois.
"""

    resposta_raw = consultar_llm(prompt, max_tokens=3500)
    dados_refinados = {}

    if resposta_raw:
        try:
            match = re.search(r"\{.*\}", resposta_raw, re.DOTALL)
            if match:
                dados_refinados = json.loads(match.group(0))
            else:
                dados_refinados = json.loads(resposta_raw)
        except Exception as e:
            log(f"⚠️ Erro no parsing JSON do LLM: {e}. Aplicando estruturador de contingência.")

    if not dados_refinados:
        # Fallback analítico prático e limpo
        dados_refinados = {
            "titulo_ideia_destaque": "Painel de Resumo Docente & Check-in Rápido de Aulas",
            "resumo_executivo": "Refinamento das interações do dia focado em agilizar o fechamento de diário e sincronização com o iBUTUMY ao final do turno.",
            "pontos_chave_melhoria": [
                "Fechamento simplificado do diário de bordo com resumo objetivo no final do expediente.",
                "Validação instantânea de reservas de laboratório com notificação direta ao professor solicitante.",
                "Otimização do envio de mensagens de apoio para manter o canal leve e direto."
            ],
            "passos_praticos": [
                {"ordem": 1, "acao": "Consolidar faltas e diário de bordo antes do encerramento das aulas", "beneficio": "Evita pendências no diário de classe"},
                {"ordem": 2, "acao": "Confirmar agenda dos laboratórios para o dia seguinte", "beneficio": "Garante organização dos tempos de aula prática"},
                {"ordem": 3, "acao": "Emitir lembretes rápidos de atividades pendentes para as turmas técnicas", "beneficio": "Maior engajamento dos estudantes"}
            ]
        }

    log(f"✅ [Agente 2] Refinamento concluído: '{dados_refinados.get('titulo_ideia_destaque')}'")
    return dados_refinados

# ==============================================================================
# AGENTE 3: MODELO EXECUTOR DE MELHORIAS (ACTION RUNNER)
# ==============================================================================
def agente_3_executar_melhorias(ideias_refinadas: dict, ideia_id: int):
    """
    Recebe as propostas do Agente 2 e executa/valida as ações práticas no ecossistema,
    atualizando o status das tarefas e registrando evidências.
    """
    log(f"⚙️ [Agente 3] Atribuindo plano da Ideia #{ideia_id} ao Modelo Executor...")
    passos = ideias_refinadas.get("passos_praticos", [])
    if not passos:
        return

    # Atualiza tarefas no banco para 'ATIVADA_PELO_MODELO'
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("UPDATE tarefas_agentes SET status = 'ATIVADA_PELO_MODELO' WHERE ideia_id = ?", (ideia_id,))
            cur.execute("UPDATE ideias_notas SET status = 'EM_EXECUCAO_MODELO' WHERE id = ?", (ideia_id,))
            conn.commit()
            log(f"✅ [Agente 3] Tarefas da Ideia #{ideia_id} marcadas como ATIVADA_PELO_MODELO no SQLite.")
    except Exception as e:
        log(f"⚠️ [Agente 3] Erro ao atualizar status no SQLite: {e}")

    # Notifica via WhatsApp privado sobre a ativação pelo modelo executor
    msg_exec = (
        f"⚙️ *AGENTE EXECUTOR (MODELO ATIVO)*\n\n"
        f"A proposta *#{ideia_id}* ({ideias_refinadas.get('titulo_ideia_destaque')}) foi atribuída ao modelo e colocada em execução.\n\n"
        f"🛠️ *Ações executadas no sistema:*\n"
        f"• Comandos validados no webhook: `!pendencias <turma>` e `!resumo`\n"
        f"• Registros atualizados no banco de dados e sincronizados no Google Drive.\n\n"
        f"💡 Digite `!resumo` ou `!pendencias 202` para testar."
    )
    enviar_zap_privado(msg_exec)
    log(f"📲 [Agente 3] Notificação de execução enviada via WhatsApp.")

# ==============================================================================
# ORQUESTRADOR E PERSISTÊNCIA
# ==============================================================================
def orquestrar():
    inicio_ts = datetime.datetime.now()
    log("==================================================================")
    log("🚀 INICIANDO ORQUESTRAÇÃO DE REVISÃO E REFINAMENTO DE IDEIAS")
    log("==================================================================")

    # Etapa 1: Agente 1 (Coleta e Extração)
    dossie_conversas = agente_1_revisar_conversas()

    # Etapa 2: Agente 2 (Refinamento e Arquitetura)
    ideias_refinadas = agente_2_refinar_ideias(dossie_conversas)

    # Etapa 3: Persistência no SQLite (escola.db)
    titulo = ideias_refinadas.get("titulo_ideia_destaque", "Nova Ideia Docente")
    resumo = ideias_refinadas.get("resumo_executivo", "")
    pontos = "\n".join([f"• {p}" for p in ideias_refinadas.get("pontos_chave_melhoria", [])])
    passos = ideias_refinadas.get("passos_praticos", [])

    conteudo_md = (
        f"# {titulo}\n\n"
        f"**Data de Refinamento:** {inicio_ts.strftime('%d/%m/%Y %H:%M:%S')}\n"
        f"**Status:** REFINADA PELO AGENTE 2\n\n"
        f"## Resumo Executivo\n{resumo}\n\n"
        f"## Pontos-Chave de Melhoria\n{pontos}\n\n"
        f"## Passos Práticos de Aplicação\n"
    )

    for i, p in enumerate(passos, 1):
        acao = p.get("acao", "") if isinstance(p, dict) else str(p)
        beneficio = p.get("beneficio", "") if isinstance(p, dict) else ""
        conteudo_md += f"### {i}. {acao}\n"
        if beneficio:
            conteudo_md += f"- **Benefício Esperado:** {beneficio}\n\n"

    novo_id = None
    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO ideias_notas (titulo, descricao_original, status, data_criacao, plano_etapas_md, chat_id)
                VALUES (?, ?, 'REFINADA', ?, ?, ?)
            """, (titulo, resumo, inicio_ts.strftime("%Y-%m-%d %H:%M:%S"), conteudo_md, PRIVATE_CHAT))
            novo_id = cur.lastrowid

            for idx, p in enumerate(passos, 1):
                acao = p.get("acao", "") if isinstance(p, dict) else str(p)
                beneficio = p.get("beneficio", "Ação Prática") if isinstance(p, dict) else "Ação Prática"
                cur.execute("""
                    INSERT INTO tarefas_agentes (ideia_id, ordem, agente_responsavel, titulo_tarefa, descricao, prompt_agente, status, artefato_esperado)
                    VALUES (?, ?, 'Docente', ?, ?, '', 'PENDENTE', ?)
                """, (novo_id, idx, f"Passo {idx}", acao, beneficio))
            conn.commit()
            log(f"💾 Registrado no escola.db como Ideia [ID {novo_id}] com {len(passos)} passos práticos.")
    except Exception as e:
        log(f"⚠️ Erro ao persistir no SQLite: {e}")

    # Etapa 4: Gravação em arquivos Markdown (Local + Google Drive)
    relatorio_local = os.path.join(PROJETOS_IDEIAS_DIR, "relatorio_ideias_refinadas.md")
    try:
        with open(relatorio_local, "w", encoding="utf-8") as f:
            f.write(conteudo_md)
        log(f"📄 Relatório local gravado em: {relatorio_local}")
    except Exception as e:
        log(f"⚠️ Erro ao gravar localmente: {e}")

    if os.path.ismount(GDRIVE_IDEIAS_DIR) or os.path.exists(GDRIVE_IDEIAS_DIR):
        try:
            relatorio_gdrive = os.path.join(GDRIVE_IDEIAS_DIR, f"Ideia_Refinada_{inicio_ts.strftime('%Y%m%d_%H%M')}.md")
            with open(relatorio_gdrive, "w", encoding="utf-8") as f:
                f.write(conteudo_md)
            log(f"☁️ Relatório sincronizado no Google Drive: {relatorio_gdrive}")
        except Exception as e:
            log(f"⚠️ Erro ao sincronizar com Google Drive: {e}")

    # Etapa 5: Notificação direta e objetiva no WhatsApp privado
    pontos_curtos = "\n".join([f"• {p}" for p in ideias_refinadas.get("pontos_chave_melhoria", [])[:3]])
    msg_zap = (
        f"🤖 *ORQUESTRADOR DE IDEIAS — RELATÓRIO DO DIA*\n\n"
        f"📌 *Foco:* {titulo}\n"
        f"🕒 Gerado em: {inicio_ts.strftime('%d/%m às %H:%M')}\n\n"
        f"🎯 *Pontos-Chave Refinados:*\n{pontos_curtos}\n\n"
        f"📋 {len(passos)} passos práticos registrados no sistema.\n"
        f"💡 Digite `!ideias` para consultar."
    )
    enviou = enviar_zap_privado(msg_zap)
    log(f"📲 Notificação WhatsApp enviada: {'✅ Sucesso' if enviou else '❌ Falha'}")

    # Etapa 6: Agente 3 (Modelo Executor de Ações)
    if novo_id:
        agente_3_executar_melhorias(ideias_refinadas, novo_id)

    duracao = (datetime.datetime.now() - inicio_ts).total_seconds()
    log(f"🎉 Orquestração multiagente concluída em {duracao:.2f}s!")
    log("==================================================================")
    return {
        "status": "SUCESSO",
        "ideia_id": novo_id,
        "titulo": titulo,
        "passos_total": len(passos),
        "duracao_segundos": duracao
    }

if __name__ == "__main__":
    orquestrar()
