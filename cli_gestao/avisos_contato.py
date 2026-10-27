#!/usr/bin/env python3
"""
Avisos automáticos para contatos restritos (rodar a cada minuto via cron).
  - aviso de aula (10 min antes)
  - agenda do dia (seg-sex 06:30)
  - localização nova do dono (no máx. 1 envio a cada 2 min)
Deduplica em `avisos_enviados`. Texto sempre curto.
"""
import os
import json, sqlite3, sys, urllib.request
from datetime import datetime, timedelta

sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
import contato_restrito as cr
from orquestrador_ideias import OPENWA_API_URL, OPENWA_API_KEY, OPENWA_SESSION_ID

ANTECEDENCIA_MIN = 10


def enviar(chat, texto) -> bool:
    req = urllib.request.Request(
        f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/messages/send-text",
        data=json.dumps({"chatId": chat, "text": texto}).encode(),
        headers={"Content-Type": "application/json", "x-api-key": OPENWA_API_KEY}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.status in (200, 201)
    except Exception as e:
        print("falha envio:", e)
        return False


def main():
    agora = datetime.now()
    c = cr._con()
    c.execute("create table if not exists avisos_enviados(chave text primary key, ts text)")

    def unico(chave):
        try:
            c.execute("insert into avisos_enviados values(?,?)", (chave, agora.isoformat()))
            c.commit()
            return True
        except sqlite3.IntegrityError:
            return False

    hoje = agora.strftime("%Y-%m-%d")
    for num in cr.CONTATOS:
        chat = f"{num}@lid"
        # 1) aviso de aula
        for ini, _fim, turma, disc, sala in cr._aulas(agora.weekday()):
            h = datetime.strptime(f"{hoje} {ini}", "%Y-%m-%d %H:%M")
            if timedelta(0) <= h - agora <= timedelta(minutes=ANTECEDENCIA_MIN) and unico(f"aula|{num}|{hoje}|{ini}|{turma}"):
                enviar(chat, f"🔔 Aula às {ini} · T{turma} · {sala}")
        # 2) agenda do dia
        if agora.weekday() < 5 and agora.strftime("%H:%M") >= "06:30" and agora.strftime("%H:%M") < "09:00" \
                and unico(f"agenda|{num}|{hoje}"):
            enviar(chat, cr._agenda())
        # 3) localização nova
        r = c.execute("select ts from localizacao_atual where id=1").fetchone()
        if r:
            ult = c.execute("select max(ts) from avisos_enviados where chave like ?", (f"loc|{num}|%",)).fetchone()[0]
            if not ult or (agora - datetime.fromisoformat(ult)) >= timedelta(minutes=2):
                if unico(f"loc|{num}|{r[0]}"):
                    enviar(chat, cr._onde())
    c.close()


if __name__ == "__main__":
    main()
