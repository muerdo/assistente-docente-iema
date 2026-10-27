#!/usr/bin/env python3
"""Envia texto curto ao WhatsApp do dono. Uso: avisar_zap.py "mensagem" """
import os
DONO_CHAT_ID = os.environ.get("DONO_CHAT_ID", "")  # chat do dono (ex.: 5599...@c.us)
import sys
sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
from avisos_contato import enviar
sys.exit(0 if enviar(DONO_CHAT_ID, " ".join(sys.argv[1:])) else 1)
