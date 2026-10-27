#!/usr/bin/env python3
"""Bloqueia até chegar mensagem do WhatsApp para o agente; imprime e marca como LIDA. Sai em seguida
(o término acorda o agente). Timeout padrão 2h."""
import os
import sys, time
sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
import ponte_agente as p
limite = time.time() + (int(sys.argv[1]) if len(sys.argv) > 1 else 7200)
while time.time() < limite:
    c = p._con()
    rs = c.execute("select id,ts,texto from ponte_agente where status='NOVA' order by id").fetchall()
    if rs:
        c.executemany("update ponte_agente set status='LIDA' where id=?", [(r[0],) for r in rs])
        c.commit(); c.close()
        print("MENSAGENS DO WHATSAPP (dono):")
        for _, ts, tx in rs:
            print(f"[{ts}] {tx}")
        sys.exit(0)
    c.close()
    time.sleep(3)
print("sem mensagens (timeout)")
