#!/usr/bin/env python3
import os
DONO_CHAT_ID = os.environ.get("DONO_CHAT_ID", "")  # chat do dono (ex.: 5599...@c.us)
import sys
import re
import json
import shutil
import sqlite3
import subprocess
import urllib.request
import urllib.parse
import yaml
from datetime import datetime
from unittest.mock import patch
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify
from pptx import Presentation

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ECOSSISTEMA_DIR = os.path.dirname(BASE_DIR)
HOME_DIR = os.environ.get("DOCENTE_HOME", ".")

for p in (HOME_DIR, os.path.join(ECOSSISTEMA_DIR, "core"), os.path.join(ECOSSISTEMA_DIR, "seguranca")):
    if p not in sys.path:
        sys.path.insert(0, p)

import gestor_docentes_rag
import webhook_zap

DADOS_DIR = os.path.join(BASE_DIR, "dados", "docentes")
ESCOLA_DB_PRINCIPAL = os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db")
OPENWA_API_URL = webhook_zap.OPENWA_API_URL
OPENWA_SESSION_ID = webhook_zap.OPENWA_SESSION_ID
OPENWA_API_KEY = webhook_zap.OPENWA_API_KEY

app = Flask(__name__)
app.secret_key = "iema_assistente_docente_secret_key"


def extrair_texto_arquivo(caminho):
    """Extrai texto completo de arquivos de base de conhecimento (PDF, Word DOCX, PowerPoint PPTX, TXT, MD, CSV)."""
    if not os.path.exists(caminho):
        return ""
    try:
        txt = gestor_docentes_rag.extrair_texto_documento(caminho)
        if txt:
            return txt
    except Exception as e:
        print(f"Aviso ao extrair via gestor_docentes_rag: {e}")
    ext = os.path.splitext(caminho)[1].lower()
    try:
        if ext in [".txt", ".md", ".json", ".csv"]:
            with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()[:20000]
        elif ext == ".pdf":
            out = subprocess.check_output(["/usr/bin/pdftotext", caminho, "-"], stderr=subprocess.DEVNULL)
            return out.decode("utf-8", errors="ignore")[:20000]
    except Exception as e:
        print(f"Erro ao extrair {caminho}: {e}")
    return ""


def sincronizar_docentes_portal_e_bot():
    """
    Garante que todos os 9 docentes do `gestor_docentes_rag` existam também no Portal Web
    (`web/dados/docentes/<id>/perfil.yaml`) e que a base RAG esteja unificada.
    """
    os.makedirs(DADOS_DIR, exist_ok=True)
    catalogo = gestor_docentes_rag.inicializar_estrutura_docentes()

    for chave, info_seed in gestor_docentes_rag.DOCENTES_SEED.items():
        doc_web_dir = os.path.join(DADOS_DIR, chave)
        os.makedirs(doc_web_dir, exist_ok=True)
        os.makedirs(os.path.join(doc_web_dir, "aulas_geradas", "slides"), exist_ok=True)
        os.makedirs(os.path.join(doc_web_dir, "aulas_geradas", "questionarios"), exist_ok=True)

        # Para o Prof. Pedro, garante vínculo direto com o banco principal dados/escola.db
        db_prof = os.path.join(doc_web_dir, "escola.db")
        if chave == "pedro" and os.path.exists(ESCOLA_DB_PRINCIPAL):
            if os.path.exists(db_prof) and not os.path.islink(db_prof):
                try:
                    os.remove(db_prof)
                except Exception:
                    pass
            if not os.path.exists(db_prof):
                try:
                    os.symlink(ESCOLA_DB_PRINCIPAL, db_prof)
                except Exception:
                    shutil.copy2(ESCOLA_DB_PRINCIPAL, db_prof)

        caminho_yaml = os.path.join(doc_web_dir, "perfil.yaml")
        info_cat = catalogo.get("docentes", {}).get(chave, {})
        turmas_prof = info_cat.get("turmas") or ["200", "201", "202"]
        disciplinas_prof = info_cat.get("disciplinas") or info_seed.get("disciplinas", ["Base Técnica"])
        w_nums = info_seed.get("identificadores_whatsapp", [])
        zap_principal = next((n for n in w_nums if n.startswith("55")), w_nums[0] if w_nums else "")

        if not os.path.exists(caminho_yaml):
            lista_disc_yaml = []
            for d_nome in disciplinas_prof:
                d_slug = re.sub(r"[^a-z0-9]+", "_", gestor_docentes_rag._normalizar(d_nome)).strip("_")[:24] or "geral"
                pasta_m = os.path.join(doc_web_dir, "materias", d_slug)
                os.makedirs(pasta_m, exist_ok=True)
                # Copia a ementa_base.md do RAG para a pasta da matéria se estiver vazia
                ementa_rag = os.path.join(gestor_docentes_rag.DOCENTES_DIR, chave, "rag_conteudos", "ementa_base.md")
                arq_destino = os.path.join(pasta_m, "ementa_base.md")
                if os.path.exists(ementa_rag) and not os.path.exists(arq_destino):
                    shutil.copy2(ementa_rag, arq_destino)
                lista_disc_yaml.append({
                    "id": d_slug,
                    "nome": d_nome,
                    "turmas": turmas_prof,
                    "ementa_topicos": [
                        f"Fundamentos de {d_nome}",
                        "Prática de Laboratório e Projetos",
                        "Avaliação Formativa BNCC",
                    ],
                })

            perfil_yaml = {
                "id": chave,
                "nome": info_seed["nome"],
                "email": f"{chave}@exemplo.edu.br",
                "telefone_whatsapp": zap_principal,
                "nivel_acesso": info_seed.get("nivel_acesso", "essencial"),
                "eixo_tecnologico": "Informação e Comunicação" if chave in ("pedro", "carlos", "marcos", "assis", "wanderson") else "Gestão, Negócios & Base Técnica",
                "instituicao": {"nome": "IEMA Pleno Açailândia", "cidade": "Açailândia - MA"},
                "turmas": [{"id": t, "nome": f"Turma {t}", "disciplina": disciplinas_prof[0]} for t in turmas_prof],
                "disciplinas": lista_disc_yaml,
                "regras_avaliacao": {"limite_faltas_critico": 15, "media_aprovacao": 6.0},
            }
            with open(caminho_yaml, "w", encoding="utf-8") as f:
                yaml.safe_dump(perfil_yaml, f, allow_unicode=True)


def listar_docentes():
    sincronizar_docentes_portal_e_bot()
    docentes = []
    if os.path.exists(DADOS_DIR):
        for nome in sorted(os.listdir(DADOS_DIR)):
            caminho_yaml = os.path.join(DADOS_DIR, nome, "perfil.yaml")
            if os.path.isfile(caminho_yaml):
                with open(caminho_yaml, "r", encoding="utf-8") as f:
                    try:
                        p = yaml.safe_load(f)
                        docentes.append({
                            "id": p.get("id", nome),
                            "nome": p.get("nome", nome),
                            "eixo": p.get("eixo_tecnologico", "Geral"),
                            "whatsapp": p.get("telefone_whatsapp", ""),
                            "nivel_acesso": p.get("nivel_acesso", "completo" if nome == "pedro" else "essencial"),
                        })
                    except Exception:
                        docentes.append({"id": nome, "nome": nome, "eixo": "Geral", "whatsapp": "", "nivel_acesso": "essencial"})
    # Coloca pedro no topo
    docentes.sort(key=lambda d: (0 if d["id"] == "pedro" else 1, d["nome"]))
    return docentes


def carregar_perfil(docente_id):
    caminho = os.path.join(DADOS_DIR, docente_id, "perfil.yaml")
    if not os.path.exists(caminho):
        return None
    with open(caminho, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def obter_db(docente_id):
    if docente_id == "pedro" and os.path.exists(ESCOLA_DB_PRINCIPAL):
        conn = sqlite3.connect(ESCOLA_DB_PRINCIPAL)
        conn.row_factory = sqlite3.Row
        return conn
    db_path = os.path.join(DADOS_DIR, docente_id, "escola.db")
    if not os.path.exists(db_path):
        if os.path.exists(ESCOLA_DB_PRINCIPAL):
            conn = sqlite3.connect(ESCOLA_DB_PRINCIPAL)
            conn.row_factory = sqlite3.Row
            return conn
        return None
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def obter_status_servicos_ecossistema():
    """Consulta em tempo real o status dos bots/serviços do ecossistema."""
    status = {
        "webhook_5000": False,
        "openwa_2785": False,
        "openwa_phone": "5500000000000",
        "omniroute_20128": False,
        "modo_vps_leve": webhook_zap.MODO_VPS_LEVE,
        "bot_pausado": webhook_zap.bot_esta_pausado(),
        "pausa_info": webhook_zap.obter_info_pausa_bot(),
        "grupos_alunos_lab": list(webhook_zap.GRUPOS_ALUNOS_LAB.values()),
    }
    try:
        with urllib.request.urlopen("http://127.0.0.1:5000/health", timeout=1.5) as r:
            status["webhook_5000"] = (r.status == 200)
    except Exception:
        pass
    try:
        req = urllib.request.Request(f"{OPENWA_API_URL}/api/sessions", headers={"x-api-key": OPENWA_API_KEY})
        with urllib.request.urlopen(req, timeout=1.5) as r:
            sess = json.loads(r.read().decode("utf-8"))
            if sess and isinstance(sess, list):
                status["openwa_2785"] = (sess[0].get("status") == "ready")
                status["openwa_phone"] = sess[0].get("phone") or "5500000000000"
    except Exception:
        pass
    try:
        with urllib.request.urlopen("http://127.0.0.1:20128/v1/models", timeout=1.2) as r:
            status["omniroute_20128"] = (r.status == 200)
    except Exception:
        pass
    return status


def obter_dados_laboratorio_e_grade(docente_id):
    """Obtém a grade individual do professor e as reservas ativas/pendentes do laboratório."""
    grade_individual = []
    grade_path = os.path.join(gestor_docentes_rag.DOCENTES_DIR, docente_id, "grade_horarios.json")
    if os.path.exists(grade_path):
        try:
            with open(grade_path, "r", encoding="utf-8") as f:
                grade_individual = json.load(f).get("aulas", [])
        except Exception:
            pass

    reservas_confirmadas = []
    reservas_pendentes = []
    if os.path.exists(ESCOLA_DB_PRINCIPAL):
        try:
            webhook_zap.init_tabela_laboratorio()
            with sqlite3.connect(ESCOLA_DB_PRINCIPAL) as conn:
                conn.row_factory = sqlite3.Row
                hoje_iso = datetime.now().strftime("%Y-%m-%d")
                reservas_confirmadas = conn.execute(
                    """
                    SELECT id, data_reserva, tempo_aula, horario_inicio, horario_fim, autor_nome, motivo, status
                    FROM reservas_laboratorio
                    WHERE status = 'CONFIRMADO' AND data_reserva >= ?
                    ORDER BY data_reserva ASC, tempo_aula ASC LIMIT 40
                    """,
                    (hoje_iso,),
                ).fetchall()
                reservas_pendentes = conn.execute(
                    """
                    SELECT id, data_reserva, tempo_aula, horario_inicio, horario_fim, autor_nome, motivo, COALESCE(grupo_origem, '') as grupo_origem
                    FROM reservas_laboratorio
                    WHERE status = 'PENDENTE'
                    ORDER BY id DESC LIMIT 20
                    """
                ).fetchall()
        except Exception as e:
            print("Erro ao ler reservas_laboratorio:", e)

    return grade_individual, reservas_confirmadas, reservas_pendentes


def obter_ultimas_linhas_log(n=25):
    log_path = os.path.join(os.environ.get("DOCENTE_HOME", "."), "webhook_live.log")
    if not os.path.exists(log_path):
        return []
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            linhas = f.readlines()
            return [l.strip() for l in linhas[-n:] if l.strip()]
    except Exception:
        return []


@app.route("/")
def index():
    docentes = listar_docentes()
    docente_id = request.args.get("docente", "pedro")
    turma_filtro = request.args.get("turma_filtro", "")
    aba_ativa = request.args.get("aba", "")
    if not any(d["id"] == docente_id for d in docentes) and docentes:
        docente_id = docentes[0]["id"]

    perfil = carregar_perfil(docente_id)
    if not perfil:
        return f"Perfil do docente '{docente_id}' não encontrado.", 404

    conn = obter_db(docente_id)
    total_alunos = 0
    atividades = []
    alunos = []
    turmas_cadastradas = []
    if conn:
        try:
            total_alunos = conn.execute("SELECT COUNT(*) FROM alunos").fetchone()[0]
            atividades = conn.execute("""
                SELECT a.id, a.turma, a.componente, a.titulo, a.data_passada,
                       (SELECT COUNT(*) FROM atividades_entregas e WHERE e.atividade_id = a.id) as total_entregas
                FROM atividades_cadastradas a
                WHERE a.prof_chave = ? OR (a.prof_chave IS NULL AND ? = 'pedro')
                ORDER BY a.id DESC LIMIT 30
            """, (docente_id, docente_id)).fetchall()

            if turma_filtro:
                alunos = conn.execute("SELECT * FROM alunos WHERE turma = ? ORDER BY nome ASC", (turma_filtro,)).fetchall()
            else:
                alunos = conn.execute("SELECT * FROM alunos ORDER BY turma ASC, nome ASC LIMIT 100").fetchall()

            t_rows = conn.execute("SELECT DISTINCT turma FROM alunos WHERE turma IS NOT NULL AND turma != '' ORDER BY turma").fetchall()
            turmas_cadastradas = [r[0] for r in t_rows]
        except Exception as e:
            print("Erro ao consultar banco:", e)
        finally:
            conn.close()

    # Materiais e Base de Conhecimento com Isolamento Estrito por Disciplina (Anti-Alucinação)
    materias_info = []
    doc_dir = os.path.join(DADOS_DIR, docente_id)
    pasta_rag_base = os.path.join(gestor_docentes_rag.DOCENTES_DIR, docente_id, "rag_conteudos")
    os.makedirs(pasta_rag_base, exist_ok=True)

    for disc in perfil.get("disciplinas", []):
        disc_id = disc.get("id")
        disc_nome = disc.get("nome", disc_id)
        disc_slug = gestor_docentes_rag.slug_disciplina(disc_id)
        pasta_disc_web = os.path.join(doc_dir, "materias", disc_id)
        pasta_disc_rag = os.path.join(pasta_rag_base, disc_slug)
        os.makedirs(pasta_disc_web, exist_ok=True)
        os.makedirs(pasta_disc_rag, exist_ok=True)

        # Arquivos exclusivos desta disciplina (sem vazamento de outros componentes)
        nomes_arqs = set()
        if os.path.exists(pasta_disc_web):
            nomes_arqs |= set(f for f in os.listdir(pasta_disc_web) if os.path.isfile(os.path.join(pasta_disc_web, f)))
        if os.path.exists(pasta_disc_rag):
            nomes_arqs |= set(f for f in os.listdir(pasta_disc_rag) if os.path.isfile(os.path.join(pasta_disc_rag, f)))

        detalhes_arquivos = []
        for nome_arq in sorted(nomes_arqs):
            caminho_arq = os.path.join(pasta_disc_web, nome_arq)
            if not os.path.exists(caminho_arq):
                caminho_arq = os.path.join(pasta_disc_rag, nome_arq)
            sz = os.path.getsize(caminho_arq) if os.path.exists(caminho_arq) else 0
            txt = extrair_texto_arquivo(caminho_arq)
            detalhes_arquivos.append({
                "nome": nome_arq,
                "tamanho_kb": round(sz / 1024, 1),
                "total_chars": len(txt),
                "snippet": (txt[:180] + "...") if len(txt) > 180 else txt,
            })

        materias_info.append({
            "id": disc_id,
            "slug": disc_slug,
            "nome": disc_nome,
            "turmas": disc.get("turmas", []),
            "topicos": disc.get("ementa_topicos", []),
            "arquivos": detalhes_arquivos,
            "total_docs": len(detalhes_arquivos),
        })

    # Aulas e Quizzes gerados
    pasta_slides = os.path.join(doc_dir, "aulas_geradas", "slides")
    pasta_quizzes = os.path.join(doc_dir, "aulas_geradas", "questionarios")
    os.makedirs(pasta_slides, exist_ok=True)
    os.makedirs(pasta_quizzes, exist_ok=True)
    slides_gerados = sorted(os.listdir(pasta_slides), reverse=True)
    quizzes_gerados = sorted(os.listdir(pasta_quizzes), reverse=True)

    grade_individual, reservas_confirmadas, reservas_pendentes = obter_dados_laboratorio_e_grade(docente_id)
    status_servicos = obter_status_servicos_ecossistema()
    logs_recentes = obter_ultimas_linhas_log(25)

    return render_template(
        "index.html",
        docentes=docentes,
        docente_atual=docente_id,
        perfil=perfil,
        total_alunos=total_alunos,
        atividades=atividades,
        alunos=alunos,
        turmas_cadastradas=turmas_cadastradas,
        turma_filtro=turma_filtro,
        materias=materias_info,
        slides_gerados=slides_gerados,
        quizzes_gerados=quizzes_gerados,
        grade_individual=grade_individual,
        reservas_confirmadas=reservas_confirmadas,
        reservas_pendentes=reservas_pendentes,
        status_servicos=status_servicos,
        logs_recentes=logs_recentes,
        aba_ativa=aba_ativa,
    )


# ==============================================================================
# CONTROLE DO LABORATÓRIO E INTERAÇÃO COM O BOT PELO PORTAL WEB
# ==============================================================================
@app.route("/lab/<docente_id>/agendar", methods=["POST"])
def agendar_lab_web(docente_id):
    dia_ou_data = request.form.get("dia_ou_data", "hoje").strip()
    tempos = request.form.get("tempos", "").strip()
    motivo = request.form.get("motivo", "").strip()
    if not tempos:
        flash("Informe ao menos um horário (1 a 9) para reservar o laboratório.", "warning")
        return redirect(url_for("index", docente=docente_id, aba="lab"))

    perfil = carregar_perfil(docente_id) or {}
    zap_num = perfil.get("telefone_whatsapp") or "5500000000000"
    nome_doc = perfil.get("nome") or docente_id.title()
    comando_resto = f"{dia_ou_data} {tempos} {motivo}".strip()

    resp = webhook_zap.agendar_lab(
        comando_resto,
        f"{zap_num}@c.us",
        f"WEB_PORTAL_{datetime.now().strftime('%H%M%S')}",
        nome_doc,
        chat_id=webhook_zap.ID_GRUPO_BASE_TECNICA,
    )
    flash(resp, "success" if "Confirmada" in resp else "warning")
    return redirect(url_for("index", docente=docente_id, aba="lab"))


@app.route("/lab/<docente_id>/cancelar", methods=["POST"])
def cancelar_lab_web(docente_id):
    alvo = request.form.get("alvo", "").strip()
    perfil = carregar_perfil(docente_id) or {}
    zap_num = perfil.get("telefone_whatsapp") or "5500000000000"
    nome_doc = perfil.get("nome") or docente_id.title()
    resp = webhook_zap.cancelar_lab(alvo, f"{zap_num}@c.us", chat_id=webhook_zap.ID_GRUPO_BASE_TECNICA, pushname=nome_doc)
    flash(resp, "info")
    return redirect(url_for("index", docente=docente_id, aba="lab"))


@app.route("/lab/<docente_id>/decidir_pendente", methods=["POST"])
def decidir_pendente_lab_web(docente_id):
    acao = request.form.get("acao", "aprovar")
    pedido_id = request.form.get("pedido_id", "").strip()
    motivo = request.form.get("motivo", "").strip()
    if acao == "aprovar":
        resp = webhook_zap.aprovar_lab(pedido_id, DONO_CHAT_ID)
        flash(resp, "success")
    else:
        arg_rec = f"{pedido_id} {motivo}".strip()
        resp = webhook_zap.recusar_lab(arg_rec, DONO_CHAT_ID)
        flash(resp, "warning")
    return redirect(url_for("index", docente=docente_id, aba="lab"))


@app.route("/bot/<docente_id>/executar_comando", methods=["POST"])
def executar_comando_bot_web(docente_id):
    comando = (request.form.get("comando") or "").strip()
    enviar_real = request.form.get("enviar_real") == "1"
    if not comando:
        return jsonify({"ok": False, "resposta": "Digite um comando (ex: minha grade, lab, !rag lgpd)."})

    perfil = carregar_perfil(docente_id) or {}
    zap_num = perfil.get("telefone_whatsapp") or "5500000000000"
    chat_alvo = f"{zap_num}@c.us"
    respostas_capturadas = []

    def _capturar_envio(cid, txt):
        respostas_capturadas.append(txt)
        if enviar_real:
            webhook_zap.enviar_zap(cid, txt)
        return True

    data_sim = {
        "chatId": chat_alvo,
        "author": chat_alvo,
        "from": chat_alvo,
        "sender": {"id": chat_alvo, "pushname": perfil.get("nome", docente_id)},
        "body": comando,
        "id": f"WEB_CMD_{datetime.now().strftime('%H%M%S%f')}",
    }

    with patch("webhook_zap.enviar_zap", side_effect=_capturar_envio), patch(
        "webhook_zap.enviar_documento_zap", return_value=True
    ):
        if docente_id != "pedro" and not webhook_zap._eh_comando_iniciar_bot_privado(webhook_zap.normalizar_texto(comando)):
            webhook_zap.iniciar_sessao_bot_privado(chat_alvo)
        webhook_zap.processar_comando(chat_alvo, comando, data_sim)

    texto_final = "\n\n---\n\n".join(respostas_capturadas) if respostas_capturadas else "Comando processado (sem saída textual direta)."
    return jsonify({
        "ok": True,
        "comando": comando,
        "docente": perfil.get("nome", docente_id),
        "enviado_whatsapp": enviar_real,
        "resposta": texto_final,
    })


@app.route("/bot/status_pausa")
def bot_status_pausa():
    return jsonify(webhook_zap.obter_info_pausa_bot())


@app.route("/bot/toggle_pausa", methods=["POST"])
def bot_toggle_pausa():
    esta_p = webhook_zap.bot_esta_pausado()
    novo_st = webhook_zap.definir_pausa_bot(not esta_p, motivo="Alternado pelo painel web")
    msg = "Assistente Docente pausado/desligado momentaneamente." if novo_st else "Assistente Docente religado com sucesso!"
    flash(msg, "warning" if novo_st else "success")
    if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.is_json:
        return jsonify({"pausado": novo_st, "mensagem": msg})
    docente_id = request.args.get("docente", "pedro")
    return redirect(url_for("index", docente=docente_id, aba="bot"))


# ==============================================================================
# GESTÃO DE ALUNOS
# ==============================================================================
@app.route("/alunos/<docente_id>/adicionar", methods=["POST"])
def adicionar_aluno(docente_id):
    nome = request.form.get("nome", "").strip()
    turma = request.form.get("turma", "").strip()
    matricula = request.form.get("matricula", "").strip()
    componente = request.form.get("componente", "").strip()
    whatsapp = request.form.get("whatsapp", "").strip()

    if not nome or not turma:
        flash("Nome e Turma são obrigatórios para cadastrar o aluno.", "danger")
        return redirect(url_for("index", docente=docente_id, aba="alunos"))

    if not matricula:
        matricula = f"{turma}_{datetime.now().strftime('%m%d%H%M%S')}"

    conn = obter_db(docente_id)
    if conn:
        try:
            conn.execute(
                "INSERT OR REPLACE INTO alunos (matricula, nome, turma, componente, faltas) VALUES (?, ?, ?, ?, 0)",
                (matricula, nome, turma, componente)
            )
            if whatsapp:
                conn.execute(
                    "INSERT OR REPLACE INTO alunos_whatsapp_vinculos (numero, matricula, nome, turma) VALUES (?, ?, ?, ?)",
                    (whatsapp, matricula, nome, turma)
                )
            conn.commit()
            flash(f"Aluno '{nome}' cadastrado com sucesso na Turma {turma}!", "success")
        except Exception as e:
            flash(f"Erro ao salvar aluno: {e}", "danger")
        finally:
            conn.close()

    return redirect(url_for("index", docente=docente_id, turma_filtro=turma, aba="alunos"))


@app.route("/alunos/<docente_id>/remover/<int:aluno_id>", methods=["POST"])
def remover_aluno(docente_id, aluno_id):
    turma = request.form.get("turma", "")
    conn = obter_db(docente_id)
    if conn:
        try:
            conn.execute("DELETE FROM alunos WHERE id = ?", (aluno_id,))
            conn.commit()
            flash("Aluno removido do banco de dados.", "info")
        except Exception as e:
            flash(f"Erro ao remover: {e}", "danger")
        finally:
            conn.close()
    return redirect(url_for("index", docente=docente_id, turma_filtro=turma, aba="alunos"))


@app.route("/alunos/<docente_id>/falta/<int:aluno_id>/<acao>", methods=["POST"])
def alterar_falta(docente_id, aluno_id, acao):
    delta = 1 if acao == "mais" else -1
    turma = request.form.get("turma", "")
    conn = obter_db(docente_id)
    if conn:
        try:
            conn.execute("UPDATE alunos SET faltas = MAX(0, faltas + ?) WHERE id = ?", (delta, aluno_id))
            conn.commit()
        except Exception as e:
            print("Erro ao atualizar falta:", e)
        finally:
            conn.close()
    return redirect(url_for("index", docente=docente_id, turma_filtro=turma, aba="alunos"))


# ==============================================================================
# BASE DE CONHECIMENTO (ISOLAMENTO ESTRITO POR DISCIPLINA & ANTI-ALUCINAÇÃO)
# ==============================================================================
@app.route("/upload_material/<docente_id>/<materia_id>", methods=["POST"])
def upload_material(docente_id, materia_id):
    if "arquivo" not in request.files:
        flash("Nenhum arquivo enviado.", "warning")
        return redirect(url_for("index", docente=docente_id, aba="materias"))
    arquivo = request.files["arquivo"]
    if arquivo.filename == "":
        flash("Arquivo sem nome.", "warning")
        return redirect(url_for("index", docente=docente_id, aba="materias"))

    disc_slug = gestor_docentes_rag.slug_disciplina(materia_id)
    pasta_destino = os.path.join(DADOS_DIR, docente_id, "materias", materia_id)
    os.makedirs(pasta_destino, exist_ok=True)
    caminho = os.path.join(pasta_destino, arquivo.filename)
    arquivo.save(caminho)

    # Ingestão universal com isolamento estrito via gestor_docentes_rag
    # Suporta .pdf, .docx, .pptx (slides), .txt, .md, .csv
    res = gestor_docentes_rag.adicionar_documento_ao_rag(
        prof_chave=docente_id,
        caminho_arquivo=caminho,
        nome_original=arquivo.filename,
        materia_id=materia_id,
    )
    if res.get("sucesso"):
        flash(
            f"✅ Documento '{arquivo.filename}' indexado com sucesso no RAG da disciplina '{materia_id}' ({res.get('caracteres', 0)} caracteres extraídos)!",
            "success",
        )
    else:
        flash(
            f"⚠️ Arquivo salvo, mas houve aviso na indexação RAG: {res.get('erro', 'formato não suportado')}",
            "warning",
        )
    return redirect(url_for("index", docente=docente_id, aba="materias"))


@app.route("/material/<docente_id>/<materia_id>/preview/<path:nome_arquivo>")
def preview_material(docente_id, materia_id, nome_arquivo):
    """Retorna o conteúdo textual completo e metadados de um material anexado para auditoria anti-alucinação."""
    disc_slug = gestor_docentes_rag.slug_disciplina(materia_id)
    caminho1 = os.path.join(DADOS_DIR, docente_id, "materias", materia_id, nome_arquivo)
    caminho2 = os.path.join(gestor_docentes_rag.DOCENTES_DIR, docente_id, "rag_conteudos", disc_slug, nome_arquivo)

    caminho = caminho1 if os.path.exists(caminho1) else caminho2
    if not os.path.exists(caminho):
        return jsonify({"erro": "Arquivo não encontrado."}), 404

    tamanho = os.path.getsize(caminho)
    texto = extrair_texto_arquivo(caminho)
    return jsonify({
        "nome": nome_arquivo,
        "materia_id": materia_id,
        "disc_slug": disc_slug,
        "tamanho_bytes": tamanho,
        "tamanho_kb": round(tamanho / 1024, 1),
        "total_caracteres": len(texto),
        "total_linhas": len(texto.splitlines()),
        "conteudo": texto
    })


@app.route("/material/<docente_id>/<materia_id>/remover/<path:nome_arquivo>", methods=["POST"])
def remover_material(docente_id, materia_id, nome_arquivo):
    """Remove o material de ambos os repositórios (web e RAG bot) preservando a integridade da disciplina."""
    disc_slug = gestor_docentes_rag.slug_disciplina(materia_id)
    caminho_web = os.path.join(DADOS_DIR, docente_id, "materias", materia_id, nome_arquivo)
    caminho_rag = os.path.join(gestor_docentes_rag.DOCENTES_DIR, docente_id, "rag_conteudos", disc_slug, nome_arquivo)
    
    # Se for PDF, checa também o arquivo _extraido.md correspondente
    nome_extraido = os.path.splitext(nome_arquivo)[0] + "_extraido.md"
    caminho_rag_ext = os.path.join(gestor_docentes_rag.DOCENTES_DIR, docente_id, "rag_conteudos", disc_slug, nome_extraido)

    removido = False
    for p in [caminho_web, caminho_rag, caminho_rag_ext]:
        if os.path.exists(p):
            try:
                os.remove(p)
                removido = True
            except Exception as e:
                print(f"Erro ao remover {p}:", e)

    if removido:
        flash(f"Arquivo '{nome_arquivo}' removido da base da disciplina.", "info")
    else:
        flash(f"Arquivo '{nome_arquivo}' não encontrado.", "warning")
    return redirect(url_for("index", docente=docente_id, aba="materias"))


# ==============================================================================
# GERAÇÃO DE CONTEÚDO (SLIDES + QUESTIONÁRIO COM BASE DE CONHECIMENTO / RAG)
# ==============================================================================
@app.route("/gerar_aula/<docente_id>", methods=["POST"])
def gerar_aula(docente_id):
    materia_id = request.form.get("materia")
    turma = request.form.get("turma", "201")
    topico = request.form.get("topico", "").strip()
    if not topico:
        flash("Informe o tema/tópico da aula.", "danger")
        return redirect(url_for("index", docente=docente_id))

    perfil = carregar_perfil(docente_id)
    instituicao = perfil.get("instituicao", {}).get("nome", "IEMA Pleno Açailândia")
    eixo = perfil.get("eixo_tecnologico", "Informação e Comunicação")
    nome_docente = perfil.get("nome", "Docente Responsável")

    materia_nome = materia_id.replace("_", " ").title()
    for d in perfil.get("disciplinas", []):
        if d.get("id") == materia_id:
            materia_nome = d.get("nome", materia_nome)
            break

    doc_dir = os.path.join(DADOS_DIR, docente_id)
    pasta_mat = os.path.join(doc_dir, "materias", materia_id)

    # 1. Consulta RAG com isolamento estrito por disciplina para evitar alucinação cruzada
    contexto_guia = ""
    fontes_usadas = []
    blocos_rag = gestor_docentes_rag.buscar_rag_docente(
        docente_id, topico, top_k=3, materia_id=materia_id, exigir_match_consulta=False
    )
    for b in blocos_rag:
        contexto_guia += f"\n[RAG {b['disciplina']} - {b['arquivo']}]:\n{b['trecho'][:800]}\n"
        fontes_usadas.append(f"{b['arquivo']} ({b['disciplina']})")

    if os.path.exists(pasta_mat):
        for arq in os.listdir(pasta_mat):
            c_arq = os.path.join(pasta_mat, arq)
            t_ext = extrair_texto_arquivo(c_arq)
            if t_ext:
                contexto_guia += f"\n[Trecho do Guia '{arq}']:\n{t_ext[:1200]}\n"
                if arq not in fontes_usadas:
                    fontes_usadas.append(arq)

    data_str = datetime.now().strftime("%Y%m%d_%H%M")
    slug = "".join([c if c.isalnum() else "_" for c in topico])[:30]

    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = topico
    s1.placeholders[1].text = (
        f"{instituicao}\n"
        f"Eixo Tecnológico: {eixo}\n"
        f"Componente: {materia_nome} • Turma: {turma}\n"
        f"Docente: {nome_docente} • Ano Letivo 2026"
    )

    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "1. Problematização & Mundo Real"
    tf2 = s2.placeholders[1].text_frame
    p = tf2.add_paragraph(); p.text = f"• O que motiva o aprendizado de {topico}?"
    p = tf2.add_paragraph(); p.text = f"• Aplicação direta no Eixo de {eixo} e no mercado."
    p = tf2.add_paragraph(); p.text = "• Desafio disparador para a formação técnica do estudante."

    s3 = prs.slides.add_slide(prs.slide_layouts[1])
    s3.shapes.title.text = f"2. Fundamentos Técnicos ({materia_nome})"
    tf3 = s3.placeholders[1].text_frame
    p = tf3.add_paragraph(); p.text = "• Conceitos fundamentados na ementa oficial do componente."
    if blocos_rag:
        resumo_rag_curto = blocos_rag[0]["trecho"].splitlines()[-1][:140]
        p = tf3.add_paragraph(); p.text = f"• RAG Ementa: {resumo_rag_curto}"
    else:
        p = tf3.add_paragraph(); p.text = "• Normas técnicas, especificações e boas práticas aplicadas."
    if fontes_usadas:
        p = tf3.add_paragraph(); p.text = f"• Fonte indexada: {fontes_usadas[0]}"

    s4 = prs.slides.add_slide(prs.slide_layouts[1])
    s4.shapes.title.text = "3. Prática de Bancada / Laboratório"
    tf4 = s4.placeholders[1].text_frame
    p = tf4.add_paragraph(); p.text = "• Execução prática individual ou em duplas no laboratório."
    p = tf4.add_paragraph(); p.text = "• Validação de resultados no console / terminal."
    p = tf4.add_paragraph(); p.text = "• Registro das evidências e códigos no caderno ou GitHub."

    s5 = prs.slides.add_slide(prs.slide_layouts[1])
    s5.shapes.title.text = "4. Checkpoint, Visto & Avaliação"
    tf5 = s5.placeholders[1].text_frame
    p = tf5.add_paragraph(); p.text = "• Apresente sua resolução para registro do visto presencial."
    p = tf5.add_paragraph(); p.text = "• Ou envie a foto/código no WhatsApp com a tag correspondente."

    nome_pptx = f"Slides_{slug}_{turma}_{data_str}.pptx"
    caminho_pptx = os.path.join(doc_dir, "aulas_geradas", "slides", nome_pptx)
    prs.save(caminho_pptx)

    fontes_lista_str = "\n".join([f"- {f}" for f in fontes_usadas]) if fontes_usadas else "- Ementa padrão cadastrada"

    quiz_conteudo = f"""# {instituicao}
### EIXO TECNOLÓGICO: {eixo.upper()}
**Componente Curricular:** {materia_nome} | **Turma:** {turma}
**Docente Responsável:** {nome_docente} | **Data:** {datetime.now().strftime('%d/%m/%Y')}
**Tema da Aula:** {topico}
**Fontes RAG Auditadas (Anti-Alucinação):**
{fontes_lista_str}

---

### PARTE 1 — VERSÃO DO ESTUDANTE (QUESTIONÁRIO FORMATIVO)

1. **[Conceitual - Fácil]** Com base na introdução à aula de **{topico}** e nas diretrizes de **{materia_nome}**, defina o conceito essencial e como ele é aplicado em rotinas técnicas de {eixo}.
   *(Espaço para resposta: 4 linhas)*

2. **[Interpretação - Médio]** Sobre a aplicação prática de **{topico}**, assinale a alternativa tecnicamente correta:
   - [ ] A) É aplicável exclusivamente em sistemas isolados sem possibilidade de conexão em rede.
   - [ ] B) Estabelece padronização de fluxos, garantindo eficiência, escalabilidade e qualidade técnica no escopo de {materia_nome}.
   - [ ] C) Não requer testes de validação ou cuidados com integridade de dados.
   - [ ] D) Trata-se de um protocolo legado sem relevância no mercado profissional atual.

3. **[Prática / Laboratório - Difícil]** Descreva as etapas do roteiro de laboratório executado hoje para consolidar **{topico}**, indicando o resultado esperado da compilação ou execução do sistema.

---

### PARTE 2 — GABARITO COMENTADO (EXCLUSIVO DO PROFESSOR)

* **Questão 1:** O estudante deve associar o conceito à sua utilidade prática na área técnica ({materia_nome}), demonstrando compreensão funcional além da mera memorização.
* **Questão 2:** **Alternativa B (Correta)**. Justificativa: A padronização e escalabilidade no escopo de {materia_nome} são os pilares centrais desse conteúdo.
* **Questão 3:** Espera-se a descrição da sequência lógica configurada no laboratório e a validação do comportamento esperado conforme o guia prático do componente.
"""
    nome_quiz = f"Questionario_{slug}_{turma}_{data_str}.md"
    caminho_quiz = os.path.join(doc_dir, "aulas_geradas", "questionarios", nome_quiz)
    with open(caminho_quiz, "w", encoding="utf-8") as f:
        f.write(quiz_conteudo)

    conn = obter_db(docente_id)
    if conn:
        try:
            conn.execute(
                "INSERT INTO atividades_cadastradas (turma, componente, titulo, data_passada, ativa) VALUES (?, ?, ?, ?, 1)",
                (turma, materia_nome, topico, datetime.now().strftime("%d/%m/%Y"))
            )
            conn.commit()
        except Exception as e:
            print("Erro ao registrar atividade:", e)
        finally:
            conn.close()

    flash(f"Aula sobre '{topico}' gerada com sucesso usando a base RAG de {nome_docente}!", "success")
    return redirect(url_for("index", docente=docente_id))


@app.route("/download/<docente_id>/<tipo>/<nome_arquivo>")
def download(docente_id, tipo, nome_arquivo):
    sub = "slides" if tipo == "slides" else "questionarios"
    caminho = os.path.join(DADOS_DIR, docente_id, "aulas_geradas", sub, nome_arquivo)
    if os.path.exists(caminho):
        return send_file(caminho, as_attachment=True)
    flash("Arquivo não encontrado.", "danger")
    return redirect(url_for("index", docente=docente_id))


@app.route("/novo_docente", methods=["POST"])
def novo_docente():
    nome = request.form.get("nome", "").strip()
    email = request.form.get("email", "").strip()
    whatsapp = request.form.get("whatsapp", "").strip()
    eixo = request.form.get("eixo", "Informação e Comunicação").strip()
    disciplina_nome = request.form.get("disciplina", "Informática Básica").strip()
    turmas_str = request.form.get("turmas", "101, 102").strip()

    id_novo = "".join([c.lower() for c in nome if c.isalnum()])[:15]
    if not id_novo:
        flash("Nome inválido.", "danger")
        return redirect(url_for("index"))

    permitido, motivo, chave_auto = gestor_docentes_rag.validar_docente_permitido(nome)
    if not permitido:
        flash(f"⛔ {motivo}", "danger")
        return redirect(url_for("index"))

    if chave_auto:
        id_novo = chave_auto

    novo_dir = os.path.join(DADOS_DIR, id_novo)
    turmas_lista = [t.strip() for t in turmas_str.split(",") if t.strip()]
    disc_slug = "".join([c.lower() if c.isalnum() else "_" for c in disciplina_nome]).strip("_")[:20]

    pasta_mat = os.path.join(novo_dir, "materias", disc_slug)
    os.makedirs(pasta_mat, exist_ok=True)
    os.makedirs(os.path.join(novo_dir, "aulas_geradas", "slides"), exist_ok=True)
    os.makedirs(os.path.join(novo_dir, "aulas_geradas", "questionarios"), exist_ok=True)

    pasta_rag_novo = os.path.join(gestor_docentes_rag.DOCENTES_DIR, id_novo, "rag_conteudos", disc_slug)
    os.makedirs(pasta_rag_novo, exist_ok=True)

    if "arquivos_base" in request.files:
        for file in request.files.getlist("arquivos_base"):
            if file and file.filename:
                c_salvo = os.path.join(pasta_mat, file.filename)
                file.save(c_salvo)
                gestor_docentes_rag.adicionar_documento_ao_rag(
                    prof_chave=id_novo,
                    caminho_arquivo=c_salvo,
                    nome_original=file.filename,
                    materia_id=disciplina_nome,
                )

    novo_perfil = {
        "id": id_novo,
        "nome": nome,
        "email": email,
        "telefone_whatsapp": whatsapp,
        "nivel_acesso": "essencial",
        "eixo_tecnologico": eixo,
        "instituicao": {"nome": "IEMA Pleno Açailândia", "cidade": "Açailândia - MA"},
        "turmas": [{"id": t, "nome": f"Turma {t}", "disciplina": disciplina_nome} for t in turmas_lista],
        "disciplinas": [{
            "id": disc_slug,
            "nome": disciplina_nome,
            "turmas": turmas_lista,
            "ementa_topicos": ["Fundamentos e Introdução", "Prática Inicial", "Avaliação Formativa"]
        }],
        "regras_avaliacao": {"limite_faltas_critico": 15, "media_aprovacao": 6.0}
    }

    with open(os.path.join(novo_dir, "perfil.yaml"), "w", encoding="utf-8") as f:
        yaml.safe_dump(novo_perfil, f, allow_unicode=True)

    # Vincula no banco SQLite e no gestor_docentes_rag para reconhecimento imediato no WhatsApp
    num_limpo = re.sub(r"\D", "", whatsapp)
    if num_limpo:
        webhook_zap.salvar_vinculo_professor(num_limpo, id_novo, nome)
    gestor_docentes_rag.DOCENTES_SEED[id_novo] = {
        "chave": id_novo,
        "nome": nome,
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [] if num_limpo else [],
        "disciplinas": [disciplina_nome],
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "ementa_inicial": f"# Base RAG — {nome}\n\n## {disciplina_nome}\n- Conteúdos programáticos de {disciplina_nome}.\n",
    }
    gestor_docentes_rag.inicializar_estrutura_docentes()

    flash(f"Docente '{nome}' ({eixo}) cadastrado e integrado ao Bot WhatsApp e RAG!", "success")
    return redirect(url_for("index", docente=id_novo))


@app.route("/disparar_whatsapp/<docente_id>", methods=["POST"])
def disparar_whatsapp(docente_id):
    turma = request.form.get("turma")
    titulo = request.form.get("titulo")
    prazo = request.form.get("prazo", "próxima aula")
    destino = request.form.get("destino", "privado")

    perfil = carregar_perfil(docente_id) or {}
    docente_nome = perfil.get("nome", "Professor")
    eixo = perfil.get("eixo_tecnologico", "Educação Técnica")

    card_zap = (
        f"📢 *NOVA ATIVIDADE — IEMA PLENO AÇAILÂNDIA*\n"
        f"• *Eixo:* {eixo}\n"
        f"• *Docente:* {docente_nome}\n"
        f"• *Turma:* {turma}\n"
        f"• *Tema:* {titulo}\n"
        f"• *Prazo de Entrega:* {prazo}\n\n"
        f"📝 *Como Entregar:* Envie a foto nítida do seu caderno ou o arquivo no chat citando seu número de chamada."
    )

    chat_alvo = f"{perfil.get('telefone_whatsapp', '5500000000000')}@c.us"
    if destino == "grupo_turma" and str(turma) in webhook_zap.GRUPOS_TURMAS:
        chat_alvo = webhook_zap.GRUPOS_TURMAS[str(turma)]

    sucesso_envio = webhook_zap.enviar_zap(chat_alvo, card_zap)
    link_wa = f"https://api.whatsapp.com/send?text={urllib.parse.quote(card_zap)}"
    return jsonify({"card": card_zap, "link_wa": link_wa, "sucesso_api": sucesso_envio, "destino": chat_alvo})


@app.route("/atividades/<docente_id>/cadastrar", methods=["POST"])
def cadastrar_atividade_web(docente_id):
    turma = request.form.get("turma", "").strip()
    componente = request.form.get("componente", "").strip()
    titulo = request.form.get("titulo", "").strip()
    data_passada = request.form.get("data_passada", "").strip() or datetime.now().strftime("%d/%m/%Y")
    
    if not turma or not titulo:
        flash("Turma e Título são obrigatórios.", "danger")
        return redirect(url_for("index", docente=docente_id, aba="atividades"))
    
    conn = obter_db(docente_id)
    if conn:
        try:
            data_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute("""
                INSERT INTO atividades_cadastradas (data_passada, data_hora_iso, turma, componente, titulo, descricao, prof_chave)
                VALUES (?, ?, ?, ?, ?, '', ?)
            """, (data_passada, data_iso, turma, componente, titulo, docente_id))
            conn.commit()
            flash(f"Atividade '{titulo}' da Turma {turma} cadastrada com sucesso!", "success")
        except Exception as e:
            flash(f"Erro ao salvar atividade: {e}", "danger")
        finally:
            conn.close()
    return redirect(url_for("index", docente=docente_id, aba="atividades"))


@app.route("/atividades/<docente_id>/lancar_visto", methods=["POST"])
def lancar_visto_web(docente_id):
    atividade_id = request.form.get("atividade_id", "").strip()
    alunos_input = request.form.get("alunos_input", "").strip()
    turma = request.form.get("turma", "").strip()
    
    if not atividade_id or not alunos_input:
        flash("Informe a atividade e os alunos para lançar vistos.", "danger")
        return redirect(url_for("index", docente=docente_id, aba="atividades"))
        
    cmd = f"!atv {alunos_input} {turma} ok" if turma else f"!atv {alunos_input} ok"
    
    perfil = carregar_perfil(docente_id) or {}
    zap_num = perfil.get("telefone_whatsapp") or "5500000000000"
    chat_alvo = f"{zap_num}@c.us"
    data_sim = {
        "chatId": chat_alvo,
        "author": chat_alvo,
        "from": chat_alvo,
        "sender": {"id": chat_alvo, "pushname": perfil.get("nome", docente_id)},
        "body": cmd,
        "id": f"WEB_ATV_{datetime.now().strftime('%H%M%S%f')}",
    }
    respostas = []
    def _capturar(cid, txt):
        respostas.append(txt)
        return True
    with patch("webhook_zap.enviar_zap", side_effect=_capturar), patch("webhook_zap.enviar_documento_zap", return_value=True):
        webhook_zap.processar_comando(chat_alvo, cmd, data_sim)
    
    msg_res = "\n".join(respostas) if respostas else "Comando de vistos executado."
    flash(f"Resultado do lançamento de vistos:\n{msg_res[:300]}", "info")
    return redirect(url_for("index", docente=docente_id, aba="atividades"))


@app.route("/atividades/<docente_id>/mover_turma", methods=["POST"])
def mover_turma_web(docente_id):
    atividade_id = request.form.get("atividade_id", "").strip()
    nova_turma = request.form.get("nova_turma", "").strip()
    
    if not atividade_id or not nova_turma:
        flash("Informe o ID da atividade e a nova turma.", "danger")
        return redirect(url_for("index", docente=docente_id, aba="atividades"))
    
    conn = obter_db(docente_id)
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("SELECT titulo, turma FROM atividades_cadastradas WHERE id = ?", (atividade_id,))
            r = cur.fetchone()
            turma_antiga = r[1] if r else "anterior"
            cur.execute("UPDATE atividades_cadastradas SET turma = ? WHERE id = ?", (nova_turma, atividade_id))
            cur.execute("UPDATE atividades_entregas SET turma = ? WHERE atividade_id = ?", (nova_turma, atividade_id))
            conn.commit()
            flash(f"Atividade #{atividade_id} transferida da Turma {turma_antiga} para a Turma {nova_turma} com sucesso!", "success")
        except Exception as e:
            flash(f"Erro ao mover atividade: {e}", "danger")
        finally:
            conn.close()
    return redirect(url_for("index", docente=docente_id, aba="atividades"))



if __name__ == "__main__":
    sincronizar_docentes_portal_e_bot()
    app.run(host="0.0.0.0", port=8000, debug=False)
