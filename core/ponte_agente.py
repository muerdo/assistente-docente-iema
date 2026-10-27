"""
Ponte WhatsApp <-> agente (dono apenas).
  !r <texto> | !agente <texto> | !ia <texto>  -> grava em `ponte_agente` (fila de entrada)
O agente (sessão agy) bloqueia em ponte_espera.py, acorda quando chega mensagem,
trabalha e responde com avisar_zap.py. Texto de resposta ao dono: curto.
"""
import json, os, re, sqlite3, urllib.request
from datetime import datetime

DB = os.environ.get("ESCOLA_DB", os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db"))
DONO_IDS = tuple(x for x in (os.environ.get("NUMERO_PROPRIETARIO", ""), os.environ.get("LID_PROPRIETARIO", "")) if x)


def _con():
    c = sqlite3.connect(DB, timeout=30)
    c.execute("create table if not exists ponte_agente(id integer primary key autoincrement, ts text, origem text, texto text, status text default 'NOVA')")
    return c


def _dig(j):
    return re.sub(r"\D", "", str(j or "").split("@")[0].split(":")[0])


def _encaminhar_host_local(chat_id: str, texto: str, data: dict) -> bool:
    """Se estiver na VPS (MODO_VPS_LEVE=1) e houver HOST_LOCAL_WEBHOOK_URL, tenta delegar ao host local."""
    host_url = os.environ.get("HOST_LOCAL_WEBHOOK_URL", "").strip()
    if not host_url:
        return False
    try:
        payload = json.dumps({"event": "message.received", "data": data or {"chatId": chat_id, "body": texto}}).encode("utf-8")
        req = urllib.request.Request(
            host_url,
            data=payload,
            headers={"Content-Type": "application/json", "X-Forwarded-From-VPS": "1"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=3) as resp:
            return 200 <= resp.status < 300
    except Exception:
        return False


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

    if os.environ.get("MODO_VPS_LEVE") == "1":
        if _encaminhar_host_local(chat_id, texto, d):
            enviar(chat_id, "✅ Recebi na VPS e encaminhei ao agente no Host Local (online).")
        else:
            enviar(
                chat_id,
                "☁️ *Modo VPS Leve (1 GB RAM)*: Host local está desligado/inacessível no momento.\n"
                "📥 Sua solicitação foi salva na fila (`ponte_agente`) e será processada assim que o PC principal ficar online.",
            )
        return True

    enviar(chat_id, "✅ Recebi. O agente vai ver.")
    return True
