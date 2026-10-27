import os
DONO_CHAT_ID = os.environ.get("DONO_CHAT_ID", "")  # chat do dono (ex.: 5599...@c.us)
#!/usr/bin/env python3
import json
import base64
import time
import urllib.request

OPENWA_API_URL = "http://127.0.0.1:2785"
OPENWA_API_KEY = os.environ.get("OPENWA_API_KEY", "")
OPENWA_SESSION_ID = os.environ.get("OPENWA_SESSION_ID", "")
DESTINATARIO = DONO_CHAT_ID

def enviar_zap(chat_id, texto):
    url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/messages/send-text"
    payload = json.dumps({"chatId": chat_id, "text": texto}).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=payload,
        headers={"Content-Type": "application/json", "x-api-key": OPENWA_API_KEY},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status in (200, 201)

def enviar_documento_zap(chat_id, caminho_arquivo, filename, mimetype):
    with open(caminho_arquivo, "rb") as f:
        conteudo_b64 = base64.b64encode(f.read()).decode("utf-8")
    url = f"{OPENWA_API_URL}/api/sessions/{OPENWA_SESSION_ID}/messages/send-document"
    payload = json.dumps({
        "chatId": chat_id,
        "base64": conteudo_b64,
        "mimetype": mimetype,
        "filename": filename
    }).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=payload,
        headers={"Content-Type": "application/json", "x-api-key": OPENWA_API_KEY},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.status in (200, 201)

def main():
    print(f"🚀 Iniciando envio dos materiais de aula para {DESTINATARIO}...")

    mensagem_resumo = (
        "📚 *PLANEJAMENTO DAS AULAS DE SEGUNDA-FEIRA*\n\n"
        "Prof. Pedro, aqui estão os materiais preparados para você avaliar e ministrar amanhã:\n\n"
        "🕒 *10:20 - 12:00 | Turma 201 | Sala 201*\n"
        "• *Disciplina:* Inteligência Artificial (4ª e 5ª aulas)\n"
        "• *Tema:* Algoritmos de Recomendação, RAG, Guardrails & Orquestração Multi-Agente\n"
        "• *Visto:* `atv 3 ok`\n\n"
        "🕒 *13:30 - 15:10 | Turma 200 | Sala 200*\n"
        "• *Disciplina:* Inteligência Artificial (6ª e 7ª aulas)\n"
        "• *Tema:* Algoritmos de Recomendação, RAG, Guardrails & Orquestração Multi-Agente\n"
        "• *Visto:* `atv 4 ok`\n\n"
        "🕒 *16:10 - 17:00 | Turma 202 | Sala 202*\n"
        "• *Disciplina:* Legislação de Informática (9ª aula)\n"
        "• *Tema:* Recomendação, Profiling, Art. 20 da LGPD & Viés Algorítmico\n"
        "• *Visto:* `atv 5 ok`\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "Abaixo estou enviando os slides (.pptx) e as listas com gabarito (.pdf) para sua avaliação. 👇"
    )

    if enviar_zap(DESTINATARIO, mensagem_resumo):
        print("✅ Mensagem de resumo enviada.")
    else:
        print("❌ Falha ao enviar mensagem de resumo.")
    time.sleep(2)

    # 1. Slides IA
    pptx_ia = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas_slides/Algoritmos_Recomendacao_Agentes_RAG.pptx")
    if enviar_documento_zap(DESTINATARIO, pptx_ia, "Algoritmos_Recomendacao_Agentes_RAG.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"):
        print("✅ Slides IA enviados com sucesso.")
    else:
        print("❌ Falha ao enviar slides IA.")
    time.sleep(2)

    # 2. Lista IA (PDF)
    pdf_ia = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas_listas/Atividade_Fixacao_Recomendacao_Agentes_RAG.pdf")
    if enviar_documento_zap(DESTINATARIO, pdf_ia, "Atividade_Fixacao_Recomendacao_Agentes_RAG.pdf", "application/pdf"):
        print("✅ Lista IA (PDF) enviada com sucesso.")
    else:
        print("❌ Falha ao enviar lista IA.")
    time.sleep(2)

    # 3. Slides Legislação
    pptx_leg = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas_slides/Recomendacao_IA_e_Legislacao_LGPD.pptx")
    if enviar_documento_zap(DESTINATARIO, pptx_leg, "Recomendacao_IA_e_Legislacao_LGPD.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"):
        print("✅ Slides Legislação enviados com sucesso.")
    else:
        print("❌ Falha ao enviar slides Legislação.")
    time.sleep(2)

    # 4. Lista Legislação (PDF)
    pdf_leg = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas_listas/Atividade_Fixacao_Legislacao_Algoritmos_LGPD.pdf")
    if enviar_documento_zap(DESTINATARIO, pdf_leg, "Atividade_Fixacao_Legislacao_Algoritmos_LGPD.pdf", "application/pdf"):
        print("✅ Lista Legislação (PDF) enviada com sucesso.")
    else:
        print("❌ Falha ao enviar lista Legislação.")

    print("🎉 Envio finalizado!")

if __name__ == "__main__":
    main()
