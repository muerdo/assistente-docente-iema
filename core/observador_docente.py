#!/usr/bin/env python3
"""
Observador do ecossistema docente (somente leitura sobre o código dos outros agentes).

Ciclo: OBSERVAR -> DIAGNOSTICAR -> DELEGAR -> APRENDER
  1. Sondas de saúde (serviços, sintaxe dos scripts, cron, DB, orquestrador).
  2. Colhe feedback das conversas diárias (transcripts agy + webhook_live.log).
  3. Cada achado vira registro deduplicado em `observador_achados`; reincidência
     sobe a severidade (aprendizado) e achados acionáveis viram tarefa delegada
     em `tarefas_agentes` + arquivo-prompt em projetos_ideias/observador/fila/.
  4. Relatório em projetos_ideias/observador/ultimo_relatorio.md; WhatsApp só
     se houver achado CRÍTICO novo.

Uso: observador_docente.py [--no-zap]
"""
import ast, datetime, glob, hashlib, json, os, re, socket, sqlite3, subprocess, sys

HOME = os.environ.get("DOCENTE_HOME", ".")
DB = f"{HOME}/escola.db"
OUT = f"{HOME}/projetos_ideias/observador"
FILA = f"{OUT}/fila"
LOG_ORQ = f"{HOME}/logs_docente/orquestrador_ideias.log"
WEBHOOK_LOG = f"{HOME}/webhook_live.log"
BRAIN = f"{HOME}/.gemini/antigravity-cli/brain"
SCRIPTS = ["webhook_zap.py", "cron_daemon.py", "orquestrador_ideias.py",
           "mcp_docente.py", "docente_cli.py", "testar_docente_noturno.py"]
PORTAS = {2785: "OpenWA API", 20128: "OmniRoute", 5000: "Webhook Zap"}
SEV = ["BAIXA", "MEDIA", "ALTA", "CRITICA"]
FRICCAO = re.compile(
    r"(n[aã]o funcion|deu erro|erro|quebr|travou|de novo|outra vez|repet|demor|"
    r"lento|n[aã]o chegou|n[aã]o respond|sumiu|errado|bug|falhou)", re.I)

agora = datetime.datetime.now()
achados = []  # (categoria, severidade, titulo, evidencia, acionavel)


def achado(cat, sev, titulo, evid="", acionavel=True):
    achados.append((cat, sev, titulo, evid[:600], acionavel))


# ---------------------------------------------------------------- 1. SAÚDE
def sondas():
    for porta, nome in PORTAS.items():
        with socket.socket() as s:
            s.settimeout(1)
            if s.connect_ex(("127.0.0.1", porta)) != 0:
                achado("servico", "CRITICA" if porta != 20128 else "ALTA",
                       f"{nome} fora do ar (porta {porta})",
                       "Subir com ./gerenciar_docente.sh start; investigar log antes.")
    if subprocess.run(["pgrep", "-f", "cron_daemon.py"], capture_output=True).returncode:
        achado("servico", "ALTA", "cron_daemon.py (alarmes/agendamentos) não está rodando")
    for f in SCRIPTS:
        p = f"{HOME}/{f}"
        if not os.path.exists(p):
            achado("codigo", "ALTA", f"{f} ausente"); continue
        try:
            ast.parse(open(p, encoding="utf-8").read(), p)
        except SyntaxError as e:
            achado("codigo", "CRITICA", f"Erro de sintaxe em {f}:{e.lineno}", str(e))
    cron = subprocess.run(["crontab", "-l"], capture_output=True, text=True).stdout
    for f in ("orquestrador_ideias.py", "testar_docente_noturno.py"):
        if f not in cron:
            achado("cron", "MEDIA", f"{f} não está agendado no crontab")


def sonda_db():
    if not os.path.exists(DB):
        achado("db", "CRITICA", "escola.db não existe"); return
    c = sqlite3.connect(DB)
    pend = c.execute("select count(*) from tarefas_agentes where status='PRONTO_PARA_EXECUCAO' "
                     "and titulo_tarefa not like '[Observador]%'").fetchone()[0]
    if pend > 5:
        achado("pipeline", "MEDIA",
               "Tarefas de ideias paradas em PRONTO_PARA_EXECUCAO (backlog do orquestrador)",
               f"{pend} tarefas sem execução; revisar e aprovar/descartar (despachante só executa as do Observador).")
    falhas = c.execute("select count(*) from tarefas_agentes where status='FALHOU' "
                       "and titulo_tarefa like '[Observador]%'").fetchone()[0]
    if falhas:
        achado("pipeline", "ALTA", "Tarefas delegadas pelo Observador falharam no despachante",
               f"{falhas} em status FALHOU; ver fila/saida_N.md.")
    err = c.execute("select count(*) from agendamentos where status='ERRO_ENVIO'").fetchone()[0]
    if err:
        achado("zap", "MEDIA", f"{err} agendamento(s) com ERRO_ENVIO", "Reenviar/checar OpenWA.")
    stuck = c.execute("select count(*) from agendamentos where status='DISPARANDO'").fetchone()[0]
    if stuck:
        achado("zap", "MEDIA", f"{stuck} agendamento(s) presos em DISPARANDO (daemon caiu no meio)")
    c.close()


def sonda_loop():
    """Lê eventos do disjuntor anti-loop (últimas 24h)."""
    p = f"{HOME}/logs_docente/loop_events.log"
    if not os.path.exists(p):
        return
    lim = (agora - datetime.timedelta(hours=24)).strftime("%Y-%m-%d %H:%M:%S")
    ev = [l.strip() for l in open(p, encoding="utf-8", errors="replace").read().splitlines()[-200:]
          if l[1:20] >= lim]
    if ev:
        achado("loop", "CRITICA", "Loop de mensagens detectado pelo disjuntor (chat bloqueado)",
               f"{len(ev)} evento(s). Último: {ev[-1]}")


def sonda_orquestrador():
    if not os.path.exists(LOG_ORQ):
        return
    linhas = open(LOG_ORQ, encoding="utf-8", errors="replace").read().splitlines()[-200:]
    for chave, titulo in (("parsing JSON do LLM", "Orquestrador: LLM devolve JSON truncado (cai no fallback genérico)"),
                          ("Erro ao persistir", "Orquestrador: falha ao gravar no escola.db"),
                          ("Erro ao consultar LLM", "Orquestrador: OmniRoute/LLM indisponível")):
        hits = [l for l in linhas if chave in l]
        if hits:
            achado("orquestrador", "ALTA" if "persistir" in chave else "MEDIA", titulo,
                   hits[-1] + f" ({len(hits)}x nas últimas 200 linhas)")


# ------------------------------------------------------------ 2. FEEDBACK
def feedback():
    hoje = agora.strftime("%Y-%m-%d")
    msgs = []
    for tf in sorted(glob.glob(f"{BRAIN}/*/.system_generated/logs/transcript.jsonl"),
                     key=os.path.getmtime, reverse=True)[:8]:
        try:
            for line in open(tf, encoding="utf-8", errors="replace"):
                o = json.loads(line)
                if (o.get("type") == "USER_INPUT" and o.get("source") == "USER_EXPLICIT"
                        and str(o.get("created_at", "")).startswith(hoje)):
                    msgs.append(o.get("content", ""))
        except Exception:
            continue
    for m in msgs:
        m = re.sub(r"<[^>]+>", "", m).strip()
        if FRICCAO.search(m) and len(m) < 600:
            achado("feedback", "MEDIA", "Fricção relatada pelo docente hoje", m, True)
    if os.path.exists(WEBHOOK_LOG):
        tail = open(WEBHOOK_LOG, encoding="utf-8", errors="replace").read().splitlines()[-400:]
        tb = [l for l in tail if "Traceback" in l or "Error" in l]
        if tb:
            achado("zap", "ALTA", f"webhook_live.log com {len(tb)} erro(s) recentes", tb[-1])


# ------------------------------------------- 3. PERSISTIR / DELEGAR / APRENDER
def persistir():
    c = sqlite3.connect(DB)
    c.execute("""create table if not exists observador_achados(
        id integer primary key autoincrement, hash text unique, categoria text,
        severidade text, titulo text, evidencia text, primeira_vez text,
        ultima_vez text, ocorrencias integer default 1, status text default 'ABERTO',
        tarefa_id integer)""")
    novos_criticos, delegados, novos_all, resolvidos = [], [], [], []
    vistos = set()
    for cat, sev, titulo, evid, acion in achados:
        h = hashlib.sha1(f"{cat}|{titulo[:80]}".encode()).hexdigest()
        vistos.add(h)
        row = c.execute("select id,ocorrencias,severidade,tarefa_id,status from observador_achados where hash=?", (h,)).fetchone()
        ts = agora.strftime("%Y-%m-%d %H:%M:%S")
        if row:
            aid, n, sev_old, tid, st = row
            n += 1
            # aprendizado: reincidência (>=3) sobe um nível de severidade
            sev_new = SEV[min(SEV.index(sev_old) + (1 if n == 3 else 0), 3)]
            c.execute("update observador_achados set ocorrencias=?,ultima_vez=?,severidade=?,evidencia=?,status='ABERTO' where id=?",
                      (n, ts, sev_new, evid, aid))
        else:
            cur = c.execute("insert into observador_achados(hash,categoria,severidade,titulo,evidencia,primeira_vez,ultima_vez) values(?,?,?,?,?,?,?)",
                            (h, cat, sev, titulo, evid, ts, ts))
            aid, tid, sev_new = cur.lastrowid, None, sev
            novos_all.append(f"[{sev}] {titulo}")
            if sev == "CRITICA":
                novos_criticos.append(titulo)
        if acion and not tid and sev_new in ("ALTA", "CRITICA", "MEDIA") and (sev_new != "MEDIA" or (row and row[1] + 1 >= 3)):
            os.makedirs(FILA, exist_ok=True)
            prompt = (f"# Tarefa delegada pelo Observador (achado #{aid}, {sev_new})\n\n"
                      f"**Problema:** {titulo}\n\n**Evidência:**\n{evid}\n\n"
                      "**Regras:** corrija a causa raiz com a menor mudança possível, "
                      "rode a verificação (py_compile/teste) e registre o resultado em "
                      f"`observador_achados` (status='RESOLVIDO' id={aid}).\n")
            arq = f"{FILA}/achado_{aid}.md"
            open(arq, "w", encoding="utf-8").write(prompt)
            t = c.execute("insert into tarefas_agentes(ideia_id,ordem,agente_responsavel,titulo_tarefa,descricao,prompt_agente,status,artefato_esperado) values(NULL,1,?,?,?,?, 'PRONTO_PARA_EXECUCAO',?)",
                          (f"Agente {cat}", f"[Observador] {titulo[:90]}", evid, prompt, arq))
            c.execute("update observador_achados set tarefa_id=? where id=?", (t.lastrowid, aid))
            delegados.append((aid, titulo))
    # auto-resolução: achado aberto que não reapareceu neste ciclo
    for aid, h, titulo in c.execute("select id,hash,titulo from observador_achados where status='ABERTO'").fetchall():
        if h not in vistos:
            c.execute("update observador_achados set status='RESOLVIDO',ultima_vez=? where id=?",
                      (agora.strftime("%Y-%m-%d %H:%M:%S"), aid))
            resolvidos.append(titulo)
    c.commit()
    rec = c.execute("select titulo,ocorrencias,severidade from observador_achados where ocorrencias>=3 and status='ABERTO' order by ocorrencias desc limit 5").fetchall()
    res = c.execute("select count(*) from observador_achados where status='RESOLVIDO'").fetchone()[0]
    c.close()
    return novos_criticos, delegados, rec, res, novos_all, resolvidos


def relatorio(delegados, rec, res):
    os.makedirs(OUT, exist_ok=True)
    L = [f"# Relatório do Observador — {agora:%d/%m/%Y %H:%M}\n"]
    L.append(f"Achados neste ciclo: **{len(achados)}** | resolvidos (histórico): **{res}**\n")
    for cat, sev, t, ev, _ in sorted(achados, key=lambda a: -SEV.index(a[1])):
        L.append(f"- **[{sev}]** ({cat}) {t}" + (f"\n  - _{ev[:200]}_" if ev else ""))
    if delegados:
        L.append("\n## Delegado neste ciclo")
        L += [f"- achado #{a}: {t} → `fila/achado_{a}.md`" for a, t in delegados]
    if rec:
        L.append("\n## Reincidentes (o sistema deve aprender com estes)")
        L += [f"- {t} — {n}x [{s}]" for t, n, s in rec]
    if not achados:
        L.append("Tudo saudável.")
    txt = "\n".join(L) + "\n"
    open(f"{OUT}/ultimo_relatorio.md", "w", encoding="utf-8").write(txt)
    return txt


def main():
    zap = "--no-zap" not in sys.argv
    sondas(); sonda_db(); sonda_loop(); sonda_orquestrador(); feedback()
    crit, deleg, rec, res, novos, resolv = persistir()
    txt = relatorio(deleg, rec, res)
    print(txt)
    if zap and (novos or deleg or resolv):
        try:
            sys.path.insert(0, HOME)
            from orquestrador_ideias import enviar_zap_privado
            L = ["🔭 *Observador*"]
            L += [f"🆕 {n}" for n in novos[:5]]
            L += [f"📤 delegado: {t[:70]}" for _, t in deleg[:5]]
            L += [f"✅ resolvido: {t[:70]}" for t in resolv[:5]]
            enviar_zap_privado("\n".join(L))
        except Exception as e:
            print("zap falhou:", e)


if __name__ == "__main__":
    main()
