#!/usr/bin/env python3
import sqlite3
import datetime
import time
import json
import urllib.request
import os
import sys
sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
import contato_restrito

DB_FILE = os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db")
OPENWA_API_URL = "http://127.0.0.1:2785"
OPENWA_API_KEY = os.environ.get("OPENWA_API_KEY", "")
OPENWA_SESSION_ID = os.environ.get("OPENWA_SESSION_ID", "")

def enviar_zap(chat_id, texto):
    url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/messages/send-text"
    req = urllib.request.Request(
        url,
        data=json.dumps({"chatId": chat_id, "text": texto}).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-api-key": OPENWA_API_KEY},
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return True
    except Exception as e:
        print(f"Erro ao enviar zap: {e}")
        return False

def loop():
    while True:
        try:
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            agora_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            c.execute("""
                SELECT id, prompt, COALESCE(chat_id, '') 
                FROM agendamentos 
                WHERE status = 'PENDENTE' AND horario_iso <= ?
            """, (agora_iso,))
            pendentes = c.fetchall()
            
            for id_agenda, prompt, chat_id in pendentes:
                c.execute("UPDATE agendamentos SET status = 'DISPARANDO' WHERE id = ? AND status = 'PENDENTE'", (id_agenda,))
                conn.commit()
                if c.rowcount > 0:
                    print(f"[{agora_iso}] Disparando alarme {id_agenda}: {prompt}")
                    msg = (
                        f"⏰ *ALARME / AVISO DOCENTE* 🚨\n\n"
                        f"📢 *Lembrete:* {prompt}\n\n"
                        f"🕒 Disparado em: {datetime.datetime.now().strftime('%H:%M:%S (%d/%m/%Y)')}"
                    )
                    enviou = enviar_zap(chat_id, msg)
                    try:
                        contato_restrito.espelhar_alarme(prompt, chat_id, enviar_zap)
                    except Exception as e:
                        print(f"Erro ao espelhar alarme: {e}")
                    st = "ENVIADO" if enviou else "ERRO_ENVIO"
                    c.execute("UPDATE agendamentos SET status = ? WHERE id = ?", (st, id_agenda))
                    conn.commit()
            
            conn.close()
        except Exception as e:
            print("Erro no daemon:", e)
        
        time.sleep(5)

if __name__ == "__main__":
    loop()
