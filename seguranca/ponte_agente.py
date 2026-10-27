"""
Ponte WhatsApp <-> agente (dono apenas).
  !r <texto> | !agente <texto> | !ia <texto>  -> grava em `ponte_agente` (fila de entrada)
O agente (sessão agy) bloqueia em ponte_espera.py, acorda quando chega mensagem,
trabalha e responde com avisar_zap.py. Texto de resposta ao dono: curto.
"""
import os
import re, sqlite3
from datetime import datetime

DB = os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db")
DONO_IDS = tuple(x for x in (os.environ.get("NUMERO_PROPRIETARIO", ""), os.environ.get("LID_PROPRIETARIO", "")) if x)


def _con():
    c = sqlite3.connect(DB, timeout=30)
    c.execute("create table if not exists ponte_agente(id integer primary key autoincrement, ts text, origem text, texto text, status text default 'NOVA')")
    return c


def _dig(j):
    return re.sub(r"\D", "", str(j or "").split("@")[0].split(":")[0])


def tratar(chat_id, texto, data, enviar) -> bool:
    t = (texto or "").strip()
    m = re.match(r"(?i)^[!/](?:r|agente|ia)\s+(\S.*)$", t, re.S)
    if not m:
        return False
    d = data or {}
    if _dig(chat_id) not in DONO_IDS:   # só o chat do dono (nunca contatos restritos/grupos)
        return False
    corpo = m.group(1).strip()
    c = _con()
    c.execute("insert into ponte_agente(ts,origem,texto) values(?,?,?)", (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), chat_id, corpo))
    c.commit(); c.close()
    enviar(chat_id, "✅ Recebi. O agente vai ver.")
    return True
