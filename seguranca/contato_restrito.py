"""
Contato restrito (ex.: 'Contato'): acesso por lista fechada de comandos, respostas curtas,
sem expor configuração do sistema. Qualquer comando fora da lista é ignorado/negado.

Ganchos no webhook_zap.py:
  - processar_comando(): `if contato_restrito.tratar(...): return`
  - do_POST(): `contato_restrito.registrar_localizacao(data)` (antes do filtro de prefixo)
  - loop_verificador_alarmes(): `contato_restrito.espelhar_alarme(...)`
"""
import os
import re, sqlite3, time
from datetime import datetime

DB = os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db")
# número/lid (só dígitos) -> apelido
CONTATOS = {}  # {"<lid_ou_numero>": "Nome"} contatos com acesso restrito
DONO_IDS = tuple(x for x in (os.environ.get("NUMERO_PROPRIETARIO", ""), os.environ.get("LID_PROPRIETARIO", "")) if x)
DIAS = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]


def _digitos(jid) -> str:
    return re.sub(r"\D", "", str(jid or "").split("@")[0].split(":")[0])


def eh_restrito(chat_id, data=None) -> bool:
    ids = {_digitos(chat_id)}
    if data:
        ids.add(_digitos(data.get("author") or (data.get("sender") or {}).get("id")))
    return bool(ids & set(CONTATOS))


def _con():
    c = sqlite3.connect(DB, timeout=30)
    c.execute("create table if not exists localizacao_atual(id integer primary key, lat real, lng real, ts text, ao_vivo integer default 0)")
    return c


# ---------------------------------------------------------------- localização
def _extrair_coords(data):
    loc = data.get("location") if isinstance(data.get("location"), dict) else {}
    lat = data.get("lat") or data.get("latitude") or loc.get("latitude") or loc.get("lat")
    lng = data.get("lng") or data.get("lon") or data.get("longitude") or loc.get("longitude") or loc.get("lng")
    try:
        return float(lat), float(lng)
    except (TypeError, ValueError):
        return None


def registrar_localizacao(data) -> bool:
    """Guarda a última localização enviada pelo dono (inclusive atualizações de localização ao vivo)."""
    tipo = str(data.get("type", "")).lower()
    if "location" not in tipo and "location" not in data and "lat" not in data and "latitude" not in data:
        return False
    autor = _digitos(data.get("author") or data.get("from") or data.get("chatId"))
    if not (data.get("fromMe") or autor in DONO_IDS):
        return False
    xy = _extrair_coords(data)
    try:
        open(os.path.join(os.environ.get("DOCENTE_HOME", "."), "logs_docente/localizacao_raw.log"), "a").write(
            f"{datetime.now():%F %T} xy={xy} keys={ {k: (str(v)[:80]) for k, v in data.items() if k != 'body'} }\n")
    except Exception:
        pass
    if not xy:
        return False
    salvar_localizacao(xy[0], xy[1], 1 if "live" in tipo else 0)
    return True


def salvar_localizacao(lat, lng, ao_vivo=0):
    """Persiste a posição atual (1 linha) e o histórico (para trajeto/auditoria)."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c = _con()
    c.execute("create table if not exists localizacao_historico(id integer primary key autoincrement, lat real, lng real, ts text, fonte text)")
    c.execute("insert or replace into localizacao_atual(id,lat,lng,ts,ao_vivo) values(1,?,?,?,?)", (lat, lng, ts, ao_vivo))
    c.execute("insert into localizacao_historico(lat,lng,ts,fonte) values(?,?,?,?)", (lat, lng, ts, "live" if ao_vivo else "gps"))
    c.execute("delete from localizacao_historico where ts < datetime('now','localtime','-30 days')")
    c.commit(); c.close()


def _token():
    import secrets
    c = _con()
    c.execute("create table if not exists config_sistema(chave text primary key, valor text, atualizado_em text)")
    r = c.execute("select valor from config_sistema where chave='loc_token'").fetchone()
    if not r:
        t = secrets.token_urlsafe(16)
        c.execute("insert into config_sistema values('loc_token',?,?)", (t, datetime.now().isoformat()))
        c.commit()
    else:
        t = r[0]
    c.close()
    return t


def http_loc(h) -> None:
    """Endpoint /loc?k=TOKEN (GET lat/lng ou POST JSON: OwnTracks, Tasker, Termux, GPSLogger)."""
    from urllib.parse import urlparse, parse_qs
    import json, hmac
    u = urlparse(h.path)
    q = {k: v[0] for k, v in parse_qs(u.query).items()}
    body = {}
    if h.command == "POST":
        try:
            body = json.loads(h.rfile.read(int(h.headers.get("Content-Length", 0))).decode() or "{}")
        except Exception:
            body = {}
    ok = hmac.compare_digest(str(q.get("k") or h.headers.get("X-Token") or ""), _token())
    d = {**q, **(body if isinstance(body, dict) else {})}
    try:
        lat = float(d.get("lat") or d.get("latitude"))
        lng = float(d.get("lon") or d.get("lng") or d.get("longitude"))
        assert -90 <= lat <= 90 and -180 <= lng <= 180
    except Exception:
        lat = lng = None
    code = 200 if (ok and lat is not None) else (403 if not ok else 400)
    if code == 200:
        salvar_localizacao(lat, lng, 1)
    h.send_response(code)
    h.send_header("Content-Type", "application/json")
    h.end_headers()
    h.wfile.write(b"[]" if code == 200 else b"{}")


def _onde() -> str:
    c = _con()
    r = c.execute("select lat,lng,ts from localizacao_atual where id=1").fetchone()
    c.close()
    if not r:
        return "Sem localização recente. Peça ao Pedro para compartilhar."
    min_ = int((datetime.now() - datetime.strptime(r[2], "%Y-%m-%d %H:%M:%S")).total_seconds() // 60)
    ha = "agora" if min_ < 1 else f"há {min_} min" if min_ < 120 else f"há {min_ // 60} h"
    return f"📍 Última posição ({ha}):\nhttps://maps.google.com/?q={r[0]},{r[1]}"


# -------------------------------------------------------------------- agenda
def _norm(t):
    import unicodedata
    return unicodedata.normalize("NFD", t.lower()).encode("ascii", "ignore").decode()


GRADE_TEMPOS = {1: ("07:30", "08:20"), 2: ("08:20", "09:10"), 3: ("09:30", "10:20"), 4: ("10:20", "11:10"),
                5: ("11:10", "12:00"), 6: ("13:30", "14:20"), 7: ("14:20", "15:10"), 8: ("15:20", "16:10"),
                9: ("16:10", "17:00")}
PROF_DONO = "pedro"


def _aulas(wd):
    """Aulas do dono no dia, da grade_base_tecnica (mesma fonte do agendamento do lab).
    Tempos consecutivos da mesma turma/disciplina viram um bloco. -> [(ini, fim, turma, disc, sala)]"""
    c = _con()
    rs = c.execute("select tempo_aula,turma,disciplina from grade_base_tecnica where dia_semana=? and prof_chave=? order by tempo_aula",
                   (wd, PROF_DONO)).fetchall()
    if not rs:  # fallback: grade manual
        rs2 = c.execute("select hora_inicio,hora_fim,turma,disciplina,sala from grade_horaria where dia_semana=? order by hora_inicio", (wd,)).fetchall()
        c.close()
        return rs2
    c.close()
    blocos = []
    for t, turma, disc in rs:
        if blocos and blocos[-1][2] == turma and blocos[-1][3] == disc and blocos[-1][4] == t - 1 and t != 6:
            blocos[-1][4] = t
        else:
            blocos.append([t, None, turma, disc, t])
    return [(GRADE_TEMPOS[b[0]][0], GRADE_TEMPOS[b[4]][1], b[2], b[3], f"Sala {b[2]}") for b in blocos]


def _linha_dia(wd, titulo=None):
    aulas = _aulas(wd)
    cab = titulo or DIAS[wd].capitalize()
    if not aulas:
        return f"{cab}: sem aulas."
    return f"*{cab}*\n" + "\n".join(f"• {a}-{b} · T{t} · {d[:26]}" for a, b, t, d, _ in aulas)


def _alvo_dia(arg):
    """Converte 'hoje|amanha|depois|segunda..sexta|dd/mm' em (weekday, rotulo) ou None."""
    from datetime import timedelta
    a = _norm(arg).strip()
    hoje = datetime.now()
    if a in ("", "hoje"):
        return hoje.weekday(), "Hoje"
    if a.startswith("amanha"):
        d = hoje + timedelta(days=1)
        return d.weekday(), "Amanhã"
    if a.startswith("depois"):
        d = hoje + timedelta(days=2)
        return d.weekday(), "Depois de amanhã"
    for i, n in enumerate(DIAS):
        if a.startswith(_norm(n)[:3]):
            return i, DIAS[i].capitalize()
    m = re.match(r"(\d{1,2})/(\d{1,2})", a)
    if m:
        try:
            d = datetime(hoje.year, int(m.group(2)), int(m.group(1)))
            return d.weekday(), d.strftime("%d/%m")
        except ValueError:
            return None
    return None


def _agenda(arg="") -> str:
    a = _norm(arg).strip()
    if a.startswith(("semana", "toda", "grade")):
        return "📅 *Semana*\n" + "\n".join(_linha_dia(w) for w in range(5))
    alvo = _alvo_dia(a)
    if not alvo:
        return "Use: !agenda [hoje|amanhã|segunda…sexta|semana]"
    wd, rot = alvo
    txt = _linha_dia(wd, f"{rot} ({DIAS[wd]})" if rot in ("Hoje", "Amanhã", "Depois de amanhã") else rot)
    if not _aulas(wd) and rot == "Hoje":
        for k in range(1, 7):
            n = (wd + k) % 7
            if _aulas(n):
                return txt + f"\nPróxima: {DIAS[n]}.\n" + _linha_dia(n)
    return txt


def _proxima() -> str:
    now = datetime.now()
    for k in range(0, 7):
        wd = (now.weekday() + k) % 7
        for ini, fim, turma, disc, sala in _aulas(wd):
            if k > 0 or ini > now.strftime("%H:%M"):
                quando = "hoje" if k == 0 else "amanhã" if k == 1 else DIAS[wd]
                return f"Próxima aula: {quando} {ini} · T{turma} · {sala}"
    return "Sem aulas cadastradas."


def _meus_alarmes(chat_id) -> str:
    c = _con()
    rs = c.execute("select id,horario_iso,prompt from agendamentos where status='PENDENTE' and chat_id=? order by horario_iso limit 10", (chat_id,)).fetchall()
    c.close()
    if not rs:
        return "Nenhum alarme seu."
    return "⏰ Seus alarmes:\n" + "\n".join(f"{i}) {h[5:16].replace('-', '/')} {p[:40]}" for i, h, p in rs)


def _cancelar(chat_id, arg) -> str:
    if not arg.strip().isdigit():
        return "Use: !alarme cancelar <id>"
    c = _con()
    cur = c.execute("update agendamentos set status='CANCELADO' where id=? and chat_id=? and status='PENDENTE'", (int(arg), chat_id))
    c.commit(); c.close()
    return "Cancelado." if cur.rowcount else "Alarme não encontrado."


AJUDA = (
    "🤖 *Assistente do Pedro*\n\n"
    "Para falar comigo, inicie com *!bot*:\n\n"
    "• `!bot onde` → Onde o Pedro está (localização)\n"
    "• `!bot agenda` → Aulas de hoje ou da semana\n"
    "• `!bot proxima` → Próxima aula na escola\n"
    "• `!bot alarme <tempo> <texto>` → Agendar lembrete\n"
    "• `!bot alarmes` → Ver seus lembretes ativos\n"
    "• `!bot alarme cancelar <id>` → Cancelar lembrete\n"
    "• `!bot <sua mensagem>` → Falar direto comigo"
)


def tem_indicador_bot(texto: str) -> bool:
    """Verifica se a mensagem contém o indicador explícito !bot (ou /bot)."""
    if not texto:
        return False
    return bool(re.search(r"(?i)(?:^|\s)[!/]bot\b", texto.strip()))


def extrair_comando_bot(texto: str) -> str:
    """Remove o prefixo/indicador !bot do texto."""
    if not texto:
        return ""
    limpo = re.sub(r"(?i)(?:^|\s)[!/]bot\b[:,\s]*", " ", texto.strip(), count=1)
    return limpo.strip()


def tratar(chat_id, texto, data, enviar, adicionar_alarme, consultar_llm=None) -> bool:
    """Retorna True se o remetente é restrito (mensagem totalmente tratada aqui)."""
    if not eh_restrito(chat_id, data):
        return False
    if (data or {}).get("fromMe"):
        return True  # eco do próprio bot/dono neste chat: ignora (evita loop)

    t = (texto or "").strip()
    # 🔒 Só processa se o contato indicou explicitamente com !bot
    if not tem_indicador_bot(t):
        return True  # Não responde nada e não interfere na conversa privada

    corpo = extrair_comando_bot(t)
    if not corpo:
        enviar(chat_id, AJUDA)
        return True

    low = corpo.lower().strip()
    cmd, _, resto = low.partition(" ")
    cmd = _norm(cmd)

    if cmd in ("ajuda", "help", "menu", "comandos"):
        r = AJUDA
    elif cmd in ("agenda", "hoje", "horario", "horarios", "aulas", "grade", "semana", "amanha"):
        arg = {"hoje": "hoje", "amanha": "amanha", "semana": "semana", "grade": "semana", "horario": "semana", "horarios": "semana", "aulas": resto or "semana"}.get(cmd, resto)
        r = _agenda(arg)
    elif cmd in ("proxima", "proximo", "proximas"):
        r = _proxima()
    elif cmd in ("onde", "local", "localizacao", "cade", "pedro"):
        r = _onde()
    elif cmd == "alarmes":
        r = _meus_alarmes(chat_id)
    elif cmd == "alarme" and resto.startswith("cancelar"):
        r = _cancelar(chat_id, resto[8:])
    elif cmd == "alarme":
        r = adicionar_alarme(re.sub(r"(?i)^alarme\s*", "", corpo), chat_id)
        m = re.search(r"Horário:\* ([^\n(]+)\(([^)]*)\)", r or "")
        r = f"⏰ Ok, {m.group(1).strip()} ({m.group(2)})." if m else "Use: !bot alarme 10m texto"
    else:
        if callable(consultar_llm):
            try:
                prompt_ia = (
                    "Você é a assistente virtual pessoal do professor Pedro no WhatsApp, conversando com o contato.\n"
                    "Responda com gentileza, presteza e de forma bastante concisa e direta (máximo 2 a 3 frases curtas).\n"
                    "Não use termos técnicos do sistema nem mencione scripts, código ou banco de dados.\n"
                    f"Mensagem do contato: {corpo}"
                )
                r = consultar_llm(prompt_ia).strip()
            except Exception:
                r = "Desculpe, não consegui processar essa mensagem agora. Para ajuda envie: !bot ajuda"
        else:
            r = "Comando não reconhecido. Envie *!bot ajuda* para ver as opções."

    enviar(chat_id, r)
    return True



# ------------------------------------------------------------ espelho de alarmes
def espelhar_alarme(prompt, target_chat, enviar) -> None:
    """Copia alarmes do dono para os contatos restritos (texto curto), exceto se for privado."""
    if _digitos(target_chat) not in DONO_IDS:
        return
    p_low = (prompt or "").lower()
    if "[privado]" in p_low or "#privado" in p_low or "privado" in p_low.split() or "solo" in p_low.split():
        return
    for num in CONTATOS:
        try:
            enviar(f"{num}@lid", f"⏰ {prompt}")
        except Exception:
            pass
