#!/usr/bin/env python3
"""
CLI de Gestão Docente para Acesso Remoto (SSH via Tailscale)
Permite consultar e modificar horários, grade de aulas e verificar status do sistema.
"""

import sys
import os
DONO_CHAT_ID = os.environ.get("DONO_CHAT_ID", "")  # chat do dono (ex.: 5599...@c.us)
import sqlite3
import csv
import json
import subprocess
import urllib.request
from datetime import datetime, timedelta

ESCOLA_DB = os.path.expanduser("~/escola.db")
BASE_AULAS = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas/automacao-iema")
AULAS_CSV = os.path.join(BASE_AULAS, "aulas.csv")
LOG_FILE = os.path.join(os.environ.get("DOCENTE_HOME", "."), "webhook_live.log")

DIAS_NOMES = [
    "Segunda-feira",
    "Terça-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "Sábado",
    "Domingo"
]

MAPA_DIAS = {
    "segunda": 0, "seg": 0, "segunda-feira": 0,
    "terca": 1, "terça": 1, "ter": 1, "terca-feira": 1, "terça-feira": 1,
    "quarta": 2, "qua": 2, "quarta-feira": 2,
    "quinta": 3, "qui": 3, "quinta-feira": 3,
    "sexta": 4, "sex": 4, "sexta-feira": 4,
    "sabado": 5, "sábado": 5, "sab": 5,
    "domingo": 6, "dom": 6,
}

GRADE_TEMPOS_PADRAO = [
    ("1ª", "07:30", "08:20"),
    ("2ª", "08:20", "09:10"),
    ("3ª", "09:30", "10:20"),
    ("4ª", "10:20", "11:10"),
    ("5ª", "11:10", "12:00"),
    ("6ª", "13:30", "14:20"),
    ("7ª", "14:20", "15:10"),
    ("8ª", "15:20", "16:10"),
    ("9ª", "16:10", "17:00"),
]

def normalizar_str(texto: str) -> str:
    import unicodedata, re
    if not texto: return ""
    texto = unicodedata.normalize("NFD", str(texto))
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto).strip().lower()

def inferir_tempos(h_ini: str, h_fim: str) -> str:
    tempos_casados = []
    for rotulo, inicio, fim in GRADE_TEMPOS_PADRAO:
        if (h_ini <= inicio and h_fim >= fim) or (h_ini == inicio) or (h_fim == fim):
            tempos_casados.append(rotulo)
    if tempos_casados:
        return " e ".join(tempos_casados) + " Aula"
    return f"{h_ini} - {h_fim}"

def cmd_horario(termo: str = ""):
    agora = datetime.now()
    t_norm = normalizar_str(termo)
    
    if "amanha" in t_norm:
        alvo_dt = agora + timedelta(days=1)
    elif "ontem" in t_norm:
        alvo_dt = agora - timedelta(days=1)
    elif t_norm in MAPA_DIAS:
        wd_alvo = MAPA_DIAS[t_norm]
        diff = (wd_alvo - agora.weekday()) % 7
        alvo_dt = agora + timedelta(days=diff)
    else:
        # Verifica se data no formato dd/mm
        import re
        m = re.search(r"(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?", t_norm)
        if m:
            dia, mes = int(m.group(1)), int(m.group(2))
            ano = int(m.group(3)) if m.group(3) else agora.year
            if ano < 100: ano += 2000
            try:
                alvo_dt = datetime(ano, mes, dia)
            except Exception:
                alvo_dt = agora
        else:
            alvo_dt = agora

    wd = alvo_dt.weekday()
    data_str = alvo_dt.strftime("%d/%m/%Y")
    dia_nome = DIAS_NOMES[wd]
    is_hoje = (alvo_dt.date() == agora.date())
    hora_atual = agora.strftime("%H:%M")

    print("\n" + "="*50)
    print(f"📅  {dia_nome.upper()} — {data_str}")
    print("⏰  Grade Escolar: 07:30 às 17:00 (50 min/aula)")
    if is_hoje:
        print(f"🕒  Hora Atual: {hora_atual}")
    print("="*50)

    # Busca em escola.db
    aulas = []
    if os.path.exists(ESCOLA_DB):
        conn = sqlite3.connect(ESCOLA_DB)
        cur = conn.cursor()
        cur.execute("SELECT id, hora_inicio, hora_fim, turma, sala, disciplina, COALESCE(tempos, '') FROM grade_horaria WHERE dia_semana = ? ORDER BY hora_inicio ASC", (wd,))
        aulas = cur.fetchall()
        conn.close()

    # Busca conteúdo no aulas.csv
    conteudos_dia = {}
    if os.path.exists(AULAS_CSV):
        try:
            with open(AULAS_CSV, "r", encoding="utf-8") as f:
                for row in csv.DictReader(f, delimiter=";"):
                    if row.get("data", "").strip() == data_str:
                        turma = row.get("turma", "").strip()
                        conteudos_dia[turma] = {
                            "tema": row.get("tema", "").strip(),
                            "conteudo": row.get("conteudo", "").strip()
                        }
        except Exception:
            pass

    if not aulas:
        print("\n🏖️  Nenhuma aula programada para este dia.")
        print("   Aproveite o tempo livre ou prepare suas atividades pedagógicas.\n")
        return

    print(f"\n📚  AULAS PROGRAMADAS ({len(aulas)}):")
    for id_aula, h_ini, h_fim, turma, sala, disc, tempos in aulas:
        if not tempos:
            tempos = inferir_tempos(h_ini, h_fim)
        
        status = ""
        if is_hoje:
            if hora_atual < h_ini:
                status = f"⏳ Próxima (inicia às {h_ini})"
            elif h_ini <= hora_atual <= h_fim:
                status = f"🚨 EM ANDAMENTO (termina às {h_fim})"
            else:
                status = "✅ Concluída"

        tema_info = conteudos_dia.get(turma, {})
        tema = tema_info.get("tema", "")

        print("-" * 50)
        print(f"🔹 [ID {id_aula}] {tempos} | {h_ini} às {h_fim}")
        print(f"   👥 Turma: {turma}  |  📍 Sala/Local: {sala}")
        print(f"   📖 Disciplina: {disc}")
        if tema:
            print(f"   🎯 Tema Previsto: {tema}")
        if status:
            print(f"   ⏱️ Status: {status}")
    print("-" * 50 + "\n")

def cmd_grade():
    if not os.path.exists(ESCOLA_DB):
        print(f"Erro: {ESCOLA_DB} não encontrado.")
        return
    conn = sqlite3.connect(ESCOLA_DB)
    cur = conn.cursor()
    cur.execute("SELECT id, dia_semana, hora_inicio, hora_fim, turma, sala, disciplina, COALESCE(tempos, '') FROM grade_horaria ORDER BY dia_semana ASC, hora_inicio ASC")
    linhas = cur.fetchall()
    conn.close()

    print("\n" + "="*60)
    print("📋  GRADE COMPLETA DA SEMANA (escola.db)")
    print("="*60)
    
    ultimo_dia = None
    for id_aula, wd, h_ini, h_fim, turma, sala, disc, tempos in linhas:
        if wd != ultimo_dia:
            print(f"\n📌 {DIAS_NOMES[wd].upper()}:")
            ultimo_dia = wd
        rotulo_tempo = f"({tempos})" if tempos else ""
        print(f"   [ID {id_aula:2d}] {h_ini}-{h_fim} {rotulo_tempo:12s} | Turma {turma:4s} | {disc} ({sala})")
    print("\n" + "="*60 + "\n")

def cmd_adicionar(dia_arg: str, h_ini: str, h_fim: str, turma: str, disciplina: str, sala: str = "Sala 202"):
    t_norm = normalizar_str(dia_arg)
    if t_norm not in MAPA_DIAS:
        try:
            wd = int(dia_arg)
            if not 0 <= wd <= 6: raise ValueError()
        except ValueError:
            print(f"❌ Dia inválido: '{dia_arg}'. Use: segunda, terca, quarta, quinta, sexta, sabado, domingo (ou 0 a 6).")
            return
    else:
        wd = MAPA_DIAS[t_norm]

    tempos = inferir_tempos(h_ini, h_fim)
    conn = sqlite3.connect(ESCOLA_DB)
    cur = conn.cursor()
    cur.execute("INSERT INTO grade_horaria (dia_semana, hora_inicio, hora_fim, turma, sala, disciplina, tempos) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (wd, h_ini, h_fim, turma, sala, disciplina, tempos))
    conn.commit()
    novo_id = cur.lastrowid
    conn.close()

    print(f"✅ Aula [ID {novo_id}] inserida com sucesso!")
    print(f"   {DIAS_NOMES[wd]} | {h_ini}-{h_fim} ({tempos}) | Turma {turma} | {disciplina} | {sala}\n")

def cmd_remover(id_arg: str):
    try:
        id_int = int(id_arg)
    except ValueError:
        print(f"❌ ID inválido: '{id_arg}'. Forneça um número inteiro.")
        return
    conn = sqlite3.connect(ESCOLA_DB)
    cur = conn.cursor()
    cur.execute("SELECT dia_semana, hora_inicio, hora_fim, turma, disciplina FROM grade_horaria WHERE id = ?", (id_int,))
    item = cur.fetchone()
    if not item:
        print(f"❌ Aula ID {id_int} não encontrada.")
        conn.close()
        return
    cur.execute("DELETE FROM grade_horaria WHERE id = ?", (id_int,))
    conn.commit()
    conn.close()
    print(f"✅ Aula [ID {id_int}] removida com sucesso! ({DIAS_NOMES[item[0]]} {item[1]}-{item[2]} Turma {item[3]} {item[4]})\n")

def cmd_status():
    print("\n" + "="*50)
    print("🛰️  STATUS DO SISTEMA DOCENTE & CONEXÕES")
    print("="*50)
    print(f"🕒  Horário Local: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Tailscale
    try:
        res = subprocess.run(["tailscale", "ip", "-4"], capture_output=True, text=True)
        ts_ip = res.stdout.strip() if res.returncode == 0 else "Não detectado"
    except Exception:
        ts_ip = "Tailscale CLI indisponível"
    print(f"🌐  Tailscale IP: {ts_ip}")
    print(f"🔑  SSH Port: 22 (Acesso: ssh USUARIO@{ts_ip})")

    # Serviços locais
    print("\n🔍  Verificação de Serviços:")
    
    # 1. OpenWA (WhatsApp)
    try:
        req = urllib.request.Request(f"http://127.0.0.1:2785/api/sessions/{os.environ.get('OPENWA_SESSION_ID', '')}",
                                     headers={"x-api-key": os.environ.get("OPENWA_API_KEY", "")})
        with urllib.request.urlopen(req, timeout=3) as resp:
            data = json.loads(resp.read().decode())
            st = data.get("status", "desconhecido")
            print(f"   • OpenWA (WhatsApp API): ✅ Ativo (Status: {st})")
    except Exception as e:
        print(f"   • OpenWA (WhatsApp API): ❌ Indisponível ({e})")

    # 2. Omniroute (LLM Gateway)
    try:
        with urllib.request.urlopen("http://localhost:20128/v1/models", timeout=3) as resp:
            print("   • OmniRoute (IA Gateway): ✅ Ativo (Porta 20128)")
    except Exception:
        # Tenta checar porta
        res = subprocess.run(["pgrep", "-f", "omniroute"], capture_output=True)
        if res.returncode == 0:
            print("   • OmniRoute (IA Gateway): ✅ Processo Rodando")
        else:
            print("   • OmniRoute (IA Gateway): ❌ Inativo")

    # 3. Webhook Zap
    res = subprocess.run(["pgrep", "-f", "webhook_zap.py"], capture_output=True, text=True)
    if res.returncode == 0:
        pids = res.stdout.strip().replace("\n", ", ")
        print(f"   • Webhook Zap (Porta 5000): ✅ Rodando (PID: {pids})")
    else:
        print("   • Webhook Zap (Porta 5000): ❌ Parado")

    # 4. Cron Daemon
    res_c = subprocess.run(["pgrep", "-f", "cron_daemon.py"], capture_output=True, text=True)
    if res_c.returncode == 0:
        pids_c = res_c.stdout.strip().replace("\n", ", ")
        print(f"   • Cron Daemon (Agendamentos): ✅ Rodando (PID: {pids_c})")
    else:
        print("   • Cron Daemon: ❌ Parado")

    print("="*50 + "\n")

def cmd_restart():
    print("🔄 Reiniciando serviços docentes...")
    subprocess.run(["pkill", "-f", "webhook_zap.py"])
    import time
    time.sleep(1)
    # Inicia webhook em background
    venv_py = os.path.join(os.environ.get("DOCENTE_HOME", "."), "mcp-audit-toolkit/venv/bin/python3")
    py_bin = venv_py if os.path.exists(venv_py) else "/usr/bin/python3"
    subprocess.Popen([py_bin, os.path.join(os.environ.get("DOCENTE_HOME", "."), "webhook_zap.py")],
                     stdout=open(LOG_FILE, "a"),
                     stderr=open(LOG_FILE, "a"),
                     start_new_session=True)
    print("✅ Webhook reiniciado!")
    time.sleep(1)
    cmd_status()

def cmd_logs():
    if not os.path.exists(LOG_FILE):
        print("Arquivo de log vazio ou não existe.")
        return
    res = subprocess.run(["tail", "-n", "30", LOG_FILE], capture_output=True, text=True)
    print(res.stdout)

def cmd_turmas():
    conn = sqlite3.connect(ESCOLA_DB)
    cur = conn.cursor()
    cur.execute("SELECT turma, COUNT(*), GROUP_CONCAT(DISTINCT componente) FROM alunos GROUP BY turma ORDER BY turma ASC")
    linhas = cur.fetchall()
    conn.close()
    
    print("\n" + "="*50)
    print("🏫  TURMAS E QUANTIDADE DE ALUNOS")
    print("="*50)
    for t, qtd, comps in linhas:
        print(f"• Turma {t}: {qtd} alunos  |  {comps}")
    print("\n💡 Dica: Use 'docente turma <turma>' para listar a chamada.")
    print("="*50 + "\n")

def cmd_alunos(turma_arg: str = ""):
    conn = sqlite3.connect(ESCOLA_DB)
    cur = conn.cursor()
    t_alvo = turma_arg.strip()
    if not t_alvo:
        agora = datetime.now()
        wd = agora.weekday()
        hora_str = agora.strftime("%H:%M")
        cur.execute("SELECT hora_inicio, hora_fim, turma FROM grade_horaria WHERE dia_semana = ? ORDER BY hora_inicio ASC", (wd,))
        for h_ini, h_fim, t in cur.fetchall():
            if h_ini <= hora_str <= h_fim or hora_str < h_ini:
                t_alvo = t
                break
    
    if not t_alvo:
        print("❌ Especifique a turma (ex: docente turma 202). Use 'docente turmas' para listar.")
        conn.close()
        return

    cur.execute("SELECT matricula, nome, turma, componente, faltas FROM alunos WHERE turma = ? ORDER BY nome ASC", (t_alvo,))
    alunos = cur.fetchall()
    conn.close()

    if not alunos:
        print(f"❌ Turma '{t_alvo}' não encontrada. Turmas disponíveis: 200, 201, 202, 203.")
        return

    print("\n" + "="*65)
    print(f"📋  CHAMADA / ALUNOS — TURMA {t_alvo} ({len(alunos)} alunos)")
    print(f"📖  {alunos[0][3]}")
    print("="*65)
    for idx, (mat, nome, turma, comp, faltas) in enumerate(alunos, 1):
        sinal = f"[⚠️ {faltas} falta(s)]" if faltas > 0 else "[0 faltas]"
        print(f"  {idx:2d}. {nome:<42} {sinal:>14}  ({mat})")
    print("="*65)
    print(f"💡 Para lançar falta/presença na aula:")
    print(f"   docente falta {t_alvo} <número ou nome>")
    print(f"   docente presenca {t_alvo} <número ou nome>")
    print(f"   docente faltas {t_alvo} 2, 6, 14\n")

def cmd_frequencia(turma_arg: str, aluno_arg: str, operacao: str = "falta"):
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    res = webhook_zap.alterar_frequencia_aluno(aluno_arg, turma_arg, "", operacao)
    print("\n" + res + "\n")

def cmd_frequencia_lote(turma_arg: str, itens_arg: str, operacao: str = "falta"):
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    res = webhook_zap.alterar_frequencia_lote(f"{turma_arg} {itens_arg}", operacao)
    print("\n" + res + "\n")

def cmd_alarme(texto_arg: str):
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    res = webhook_zap.adicionar_alarme(texto_arg)
    print("\n" + res + "\n")

def cmd_alarmes():
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    res = webhook_zap.listar_alarmes()
    print("\n" + res + "\n")

def cmd_ideias(texto_arg: str = ""):
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    t = texto_arg.strip()
    if not t:
        print("\n" + webhook_zap.listar_ideias_notas() + "\n")
        return
    if t.isdigit() or t.startswith("ver "):
        id_alvo = t if t.isdigit() else t.split()[1]
        print("\n" + webhook_zap.obter_detalhes_ideia(id_alvo) + "\n")
        return

    print(f"\n💡 Processando ideia com IA e decompondo para múltiplos agentes...")
    print(f"Ideia: \"{t}\"\n")
    res = webhook_zap.processar_ideia_core(t)
    print(res["mensagem"] + "\n")

def cmd_away(arg: str = ""):
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    res = webhook_zap.alternar_modo_away(arg)
    print("\n" + res + "\n")

def cmd_cloud(instrucao: str = ""):
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    webhook_zap.executar_cloud_cli(instrucao, DONO_CHAT_ID)

def imprimir_ajuda():
    print("""
Uso do comando 'docente':
  # Horários e Grade:
  docente horario [hoje|amanha|sexta|dd/mm]  - Consulta dia da semana e horários de aula
  docente grade                              - Exibe grade horária semanal completa com IDs
  docente adicionar <dia> <ini> <fim> <turma> <disciplina> [sala]
  docente remover <id>                       - Remove uma aula da grade pelo ID

  # Chamada & Alunos (Sala de Aula):
  docente turmas                             - Lista todas as turmas cadastradas
  docente turma [200|201|202|203]            - Lista numerada de alunos com faltas
  docente falta <turma> <nº ou nome>         - Atribui falta a um aluno
  docente presenca <turma> <nº ou nome>      - Abona falta / confirma presença
  docente faltas <turma> <n1, n2, n3...>     - Atribui faltas em lote na chamada

  # Ideias & Multiagentes:
  docente ideias                             - Lista todas as ideias e protótipos registrados
  docente ideias <ideia>                     - Decompõe ideia em etapas para subagentes e gera protótipo MVP
  docente ideias <id>                        - Exibe detalhes técnicos, prompts dos agentes e diretório
  docente ideia <ideia ou id>                - (Sinônimo)

  # Cloud CLI & Modo Away (Terminal WhatsApp):
  docente cloud <comando ou instrução>       - Executa no terminal e envia resultado ao WhatsApp
  docente away [on|off|status]               - Ativa/desativa modo ausente (notificações no WhatsApp)

  # Alarmes & Lembretes Privados (WhatsApp):
  docente alarme <tempo/hora> <mensagem>     - Programa aviso privado no WhatsApp
  docente alarmes                            - Lista alarmes pendentes
  docente alarme cancelar <id>               - Cancela um alarme

  # Infraestrutura & Serviços:
  docente status                             - Exibe status da rede (Tailscale, SSH, Modo Away) e serviços
  docente logs                               - Exibe os últimos logs do webhook
  docente restart                            - Reinicia o webhook em background
""")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        cmd_horario()
        sys.exit(0)

    arg = sys.argv[1].lower()
    if arg in ["horario", "horários", "horarios", "aulas"]:
        termo = sys.argv[2] if len(sys.argv) > 2 else ""
        cmd_horario(termo)
    elif arg in ["grade", "semana"]:
        cmd_grade()
    elif arg in ["turmas"]:
        cmd_turmas()
    elif arg in ["turma", "chamada", "alunos"]:
        t_arg = sys.argv[2] if len(sys.argv) > 2 else ""
        cmd_alunos(t_arg)
    elif arg in ["falta"]:
        if len(sys.argv) < 3:
            print("❌ Uso: docente falta [turma] <aluno ou número>")
        elif len(sys.argv) == 3:
            # Pode ser: docente falta 202 6 ou docente falta "Carlos Daniel"
            arg2 = sys.argv[2]
            if arg2 in ["200", "201", "202", "203"]:
                print("❌ Especifique o aluno ou número: docente falta 202 <nº>")
            else:
                cmd_frequencia("", arg2, "falta")
        else:
            t_arg = sys.argv[2]
            aluno_arg = " ".join(sys.argv[3:])
            cmd_frequencia(t_arg, aluno_arg, "falta")
    elif arg in ["presenca", "presença", "abono"]:
        if len(sys.argv) < 3:
            print("❌ Uso: docente presenca [turma] <aluno ou número>")
        elif len(sys.argv) == 3:
            arg2 = sys.argv[2]
            if arg2 in ["200", "201", "202", "203"]:
                print("❌ Especifique o aluno ou número: docente presenca 202 <nº>")
            else:
                cmd_frequencia("", arg2, "presenca")
        else:
            t_arg = sys.argv[2]
            aluno_arg = " ".join(sys.argv[3:])
            cmd_frequencia(t_arg, aluno_arg, "presenca")
    elif arg in ["faltas"]:
        if len(sys.argv) < 4:
            print("❌ Uso: docente faltas <turma> <n1, n2, n3...>")
            print("   Ex: docente faltas 202 2, 6, 14")
        else:
            t_arg = sys.argv[2]
            itens = " ".join(sys.argv[3:])
            cmd_frequencia_lote(t_arg, itens, "falta")
    elif arg in ["presencas", "presenças"]:
        if len(sys.argv) < 4:
            print("❌ Uso: docente presencas <turma> <n1, n2, n3...>")
        else:
            t_arg = sys.argv[2]
            itens = " ".join(sys.argv[3:])
            cmd_frequencia_lote(t_arg, itens, "presenca")
    elif arg in ["adicionar", "add", "inserir"]:
        if len(sys.argv) < 7:
            print("❌ Uso: docente adicionar <dia> <hora_inicio> <hora_fim> <turma> <disciplina> [sala]")
            print("   Ex: docente adicionar quinta 13:30 15:10 202 'INTELIGÊNCIA ARTIFICIAL' 'Sala 202'")
        else:
            dia_a = sys.argv[2]
            h_ini = sys.argv[3]
            h_fim = sys.argv[4]
            turma = sys.argv[5]
            disc = sys.argv[6]
            sala = sys.argv[7] if len(sys.argv) > 7 else "Sala 202"
            cmd_adicionar(dia_a, h_ini, h_fim, turma, disc, sala)
    elif arg in ["remover", "rm", "del", "excluir"]:
        if len(sys.argv) < 3:
            print("❌ Uso: docente remover <id_aula>")
        else:
            cmd_remover(sys.argv[2])
    elif arg in ["alarme", "aviso", "lembrete"]:
        texto_arg = " ".join(sys.argv[2:])
        cmd_alarme(texto_arg)
    elif arg in ["alarmes", "avisos", "lembretes"]:
        cmd_alarmes()
    elif arg in ["ideias", "ideia", "nota", "notas"]:
        texto_arg = " ".join(sys.argv[2:])
        cmd_ideias(texto_arg)
    elif arg in ["cloud", "terminal"]:
        instrucao = " ".join(sys.argv[2:])
        cmd_cloud(instrucao)
    elif arg in ["away", "ausente"]:
        away_arg = sys.argv[2] if len(sys.argv) > 2 else ""
        cmd_away(away_arg)
    elif arg in ["status", "rede", "info"]:
        cmd_status()
    elif arg in ["restart", "reiniciar"]:
        cmd_restart()
    elif arg in ["logs", "log"]:
        cmd_logs()
    else:
        imprimir_ajuda()
