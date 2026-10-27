#!/usr/bin/env python3
"""
Disparo Automático do Resumo Matinal do Docente
Envia um briefing consolidado para o WhatsApp privado às 06:45 (segunda a sexta).
"""
import sys
import os
DONO_CHAT_ID = os.environ.get("DONO_CHAT_ID", "")  # chat do dono (ex.: 5599...@c.us)

sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
from webhook_zap import gerar_resumo_matinal, enviar_zap, log_terminal

DESTINATARIO = DONO_CHAT_ID

def main():
    try:
        log_terminal("🌅 Disparando Resumo Matinal automático do docente...")
        texto = gerar_resumo_matinal(DESTINATARIO)
        sucesso = enviar_zap(DESTINATARIO, texto)
        if sucesso:
            log_terminal("✅ Resumo Matinal enviado com sucesso ao WhatsApp privado.")
        else:
            log_terminal("❌ Falha no envio do Resumo Matinal.")
    except Exception as e:
        log_terminal(f"❌ Erro ao disparar Resumo Matinal: {e}")

if __name__ == "__main__":
    main()
