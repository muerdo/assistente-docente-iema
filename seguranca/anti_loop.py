"""
Detector/disjuntor de loop de mensagens, leve em memória.

Memória limitada: no máx. MAX_CHATS chats, cada um com deque(maxlen=N) de
(timestamp, hash_int) — alguns KB no total. Sem armazenar texto.

Dispara (e bloqueia o chat por BLOQUEIO_S) quando, numa janela curta:
  - rajada: mais de LIM_RAJADA mensagens (entrada ou saída) em JANELA_S, ou
  - repetição: o mesmo texto >= LIM_REPETE vezes em JANELA_REP_S.
Eventos vão para logs_docente/loop_events.log (o Observador lê e delega).

API: permitir_entrada(chat, texto) / permitir_envio(chat, texto) -> bool
"""
import os
import time
from collections import deque, OrderedDict
from datetime import datetime

MAX_CHATS = 50
JANELA_S, LIM_RAJADA = 30, 8
JANELA_REP_S, LIM_REPETE = 60, 4
BLOQUEIO_S = 600
LOG = os.path.join(os.environ.get("DOCENTE_HOME", "."), "logs_docente/loop_events.log")

_hist = OrderedDict()   # (dir, chat) -> deque[(ts, hash)]
_bloq = {}              # (dir, chat) -> ate_ts


def _evento(msg):
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%F %T}] {msg}\n")
    except Exception:
        pass


def _checar(direcao, chat, texto) -> bool:
    agora = time.time()
    k = (direcao, str(chat))
    if _bloq.get(k, 0) > agora:
        return False
    if k in _bloq:
        del _bloq[k]

    txt = str(texto or "").strip()
    txt_low = txt.lower()
    # Trava imediata de recursão / eco em cascata (atua já na 1ª ocorrência de eco de fallback)
    if (
        "comando 'comando '" in txt_low
        or txt_low.startswith("comando '")
        or "recebido. digite `!ajuda`" in txt_low
        or "não reconhecido. digite `!ajuda`" in txt_low
        or "nao reconhecido. digite `!ajuda`" in txt_low
    ):
        if direcao == "in":
            _evento(f"ECO/LOOP BLOQUEADO NA ENTRADA chat={chat}: {txt[:80]}")
            return False
        if "comando 'comando '" in txt_low or txt_low.count("comando '") >= 2:
            _bloq[k] = agora + BLOQUEIO_S
            _evento(f"LOOP RECURSIVO BLOQUEADO chat={chat}; bloqueado {BLOQUEIO_S}s")
            return False

    dq = _hist.pop(k, None) or deque(maxlen=LIM_RAJADA + 2)
    _hist[k] = dq                      # move para o fim (LRU)
    while len(_hist) > MAX_CHATS:
        _hist.popitem(last=False)
    h = hash(texto) if texto else 0
    dq.append((agora, h))
    rajada = sum(1 for t, _ in dq if agora - t <= JANELA_S)
    repete = sum(1 for t, x in dq if x == h and agora - t <= JANELA_REP_S)
    if rajada > LIM_RAJADA or repete >= LIM_REPETE:
        _bloq[k] = agora + BLOQUEIO_S
        dq.clear()
        motivo = f"rajada {rajada}/{JANELA_S}s" if rajada > LIM_RAJADA else f"repetição {repete}x/{JANELA_REP_S}s"
        _evento(f"LOOP {direcao.upper()} chat={chat} {motivo}; bloqueado {BLOQUEIO_S}s")
        return False
    return True


def permitir_entrada(chat, texto="") -> bool:
    return _checar("in", chat, texto)


def permitir_envio(chat, texto="") -> bool:
    return _checar("out", chat, texto)
