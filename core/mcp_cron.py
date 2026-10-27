import os
import sys
import json
import sqlite3
import datetime

DB_FILE = os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db")

def setup_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS agendamentos (
                 id INTEGER PRIMARY KEY AUTOINCREMENT,
                 prompt TEXT,
                 horario_iso TEXT,
                 status TEXT DEFAULT 'PENDENTE')''')
    conn.commit()
    conn.close()

def registrar_agendamento(prompt, horario_iso):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO agendamentos (prompt, horario_iso) VALUES (?, ?)", (prompt, horario_iso))
    conn.commit()
    conn.close()
    return f"Agendado com sucesso para {horario_iso}."

def mcp_server():
    setup_db()
    for line in sys.stdin:
        try:
            req = json.loads(line)
            method = req.get("method")
            req_id = req.get("id")
            if method == "initialize":
                resp = {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "result": {
                        "protocolVersion": req.get("params", {}).get("protocolVersion", "2024-11-05"),
                        "capabilities": {"tools": {}},
                        "serverInfo": {"name": "agendamentos", "version": "1.0.0"}
                    }
                }
                sys.stdout.write(json.dumps(resp) + "\n")
                sys.stdout.flush()
            elif method == "notifications/initialized":
                pass
            elif method == "tools/list":
                resp = {"jsonrpc": "2.0", "id": req_id, "result": {"tools": [
                    {"name": "agendar_tarefa", "description": "Agenda um prompt para ser executado (enviado ao usuário) em um horário futuro.", "inputSchema": {"type": "object", "properties": {"prompt": {"type": "string"}, "horario_iso": {"type": "string", "description": "Formato YYYY-MM-DD HH:MM:00"}}, "required": ["prompt", "horario_iso"]}}
                ]}}
                sys.stdout.write(json.dumps(resp) + "\n")
                sys.stdout.flush()
            elif method == "tools/call":
                args = req["params"]["arguments"]
                res = registrar_agendamento(args["prompt"], args["horario_iso"])
                resp = {"jsonrpc": "2.0", "id": req_id, "result": {"content": [{"type": "text", "text": res}]}}
                sys.stdout.write(json.dumps(resp) + "\n")
                sys.stdout.flush()
        except Exception as e:
            pass

if __name__ == "__main__":
    mcp_server()
