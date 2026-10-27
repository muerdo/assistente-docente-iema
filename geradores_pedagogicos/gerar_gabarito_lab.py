#!/usr/bin/env python3
"""
==============================================================================
GERADOR DE GABARITO DE RESERVAS DE LABORATÓRIO IMPRESSO & SINCRONIZADOR (IDEIA #15)
==============================================================================
Compatibiliza a cultura escolar de anotações físicas na porta do laboratório com
o ecossistema digital do assistente docente.

Funcionalidades:
1. Gera a Folha Semanal de Gabarito do Laboratório de Informática em HTML/PDF,
   Markdown e TXT formatada para impressão (papel A4), contendo horários fixos
   da Grade Técnica, blocos de limpeza e caixas de assinatura para os horários livres.
2. Fornece API/CLI para sincronização rápida das marcações de papel para o banco de
   dados (reservas_laboratorio) evitando divergências entre o papel e o bot.
==============================================================================
"""

import os
import sys
import re
import sqlite3
import subprocess
import shutil
from datetime import datetime, timedelta

HOME = os.environ.get("DOCENTE_HOME", ".")
ESCOLA_DB = f"{HOME}/escola.db"
PROJETOS_IDEIAS_DIR = f"{HOME}/projetos_ideias"
DIR_IDEIA_15 = f"{PROJETOS_IDEIAS_DIR}/ideia_15_de_reservar_por_papel_escrito"
DIR_GABARITOS = f"{DIR_IDEIA_15}/gabaritos_gerados"
DIR_GDRIVE_LAB = f"{HOME}/mounts/gdrive/Aulas/Laboratorio/Gabaritos"

os.makedirs(DIR_GABARITOS, exist_ok=True)

GRADE_LAB = [
    (1, "07:30", "08:25"),
    (2, "08:25", "09:10"),
    (3, "09:30", "10:20"),
    (4, "10:20", "11:10"),
    (5, "11:10", "12:00"),
    (6, "13:30", "14:20"),
    (7, "14:20", "15:05"),
    (8, "15:20", "16:10"),
    (9, "16:10", "17:00"),
]

DIAS_SEMANA = [
    ("Segunda", 0),
    ("Terça", 1),
    ("Quarta", 2),
    ("Quinta", 3),
    ("Sexta", 4),
]


def obter_segunda_feira_referencia(data_ref: datetime = None, termo: str = "") -> datetime:
    if data_ref is None:
        data_ref = datetime.now()
    t_clean = str(termo or "").lower()
    # Se uma data de segunda-feira explícita for passada (weekday == 0):
    if data_ref.weekday() == 0 and "proxima" not in t_clean and "prox" not in t_clean:
        return data_ref
    # Nos fins de semana (sábado ou domingo), por padrão a semana de interesse é a que começa na segunda-feira seguinte
    if "proxima" in t_clean or "prox" in t_clean or (data_ref.weekday() >= 5 and "passada" not in t_clean and "anterior" not in t_clean):
        return data_ref - timedelta(days=data_ref.weekday()) + timedelta(days=7)
    return data_ref - timedelta(days=data_ref.weekday())


def carregar_grade_base(conn) -> dict:
    """Carrega as aulas fixas da grade base técnica (dia_semana, tempo) -> texto_aula."""
    grade = {}
    try:
        cur = conn.cursor()
        cur.execute("SELECT dia_semana, tempo_aula, turma, disciplina, prof_nome FROM grade_base_tecnica")
        for dia, tempo, turma, disc, prof in cur.fetchall():
            prof_curto = prof.split()[-1] if prof else ""
            grade[(dia, tempo)] = f"Turma {turma} ({disc[:14]}) - {prof_curto}"
    except Exception:
        pass
    return grade


def carregar_reservas_semana(conn, data_inicio_iso: str, data_fim_iso: str) -> dict:
    """Carrega reservas ativas para a semana (data_iso, tempo) -> autor_nome."""
    reservas = {}
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT data_reserva, tempo_aula, autor_nome, motivo
            FROM reservas_laboratorio
            WHERE data_reserva BETWEEN ? AND ? AND status IN ('CONFIRMADO', 'PENDENTE')
        """, (data_inicio_iso, data_fim_iso))
        for d_iso, tempo, autor, motivo in cur.fetchall():
            reservas[(d_iso, tempo)] = (autor or "Docente", motivo or "")
    except Exception:
        pass
    return reservas


def obter_modo_limpeza(conn) -> str:
    try:
        cur = conn.cursor()
        cur.execute("SELECT valor FROM config_sistema WHERE chave = 'limpeza_segunda_modo'")
        row = cur.fetchone()
        return row[0].strip().lower() if row and row[0] else "padrao"
    except Exception:
        return "padrao"


def eh_limpeza(dia_idx: int, tempo: int, modo_limpeza: str) -> bool:
    if dia_idx != 0:
        return False
    if modo_limpeza in ("integral", "dia_todo", "completo"):
        return True
    if modo_limpeza in ("manha", "matutino"):
        return tempo in {1, 2, 3, 4, 5}
    return tempo in {1, 2}


def gerar_gabarito_semanal(data_ref: datetime = None) -> dict:
    """Gera gabarito semanal em Markdown, HTML e PDF."""
    segunda_dt = obter_segunda_feira_referencia(data_ref)
    sexta_dt = segunda_dt + timedelta(days=4)

    segunda_iso = segunda_dt.strftime("%Y-%m-%d")
    sexta_iso = sexta_dt.strftime("%Y-%m-%d")

    periodo_str = f"{segunda_dt.strftime('%d/%m/%Y')} a {sexta_dt.strftime('%d/%m/%Y')}"
    timestamp_geracao = datetime.now().strftime("%d/%m/%Y às %H:%M")

    with sqlite3.connect(ESCOLA_DB) as conn:
        grade_base = carregar_grade_base(conn)
        reservas = carregar_reservas_semana(conn, segunda_iso, sexta_iso)
        modo_limp = obter_modo_limpeza(conn)

    # Monta matriz de dados: [tempo][dia_idx] = texto_celula
    matriz = []
    for tempo, ini, fim in GRADE_LAB:
        linha_tempo = {"tempo": tempo, "horario": f"{ini} - {fim}", "dias": []}
        for nome_dia, dia_idx in DIAS_SEMANA:
            data_dia = segunda_dt + timedelta(days=dia_idx)
            data_dia_iso = data_dia.strftime("%Y-%m-%d")

            celula = {
                "dia": nome_dia,
                "data": data_dia.strftime("%d/%m"),
                "tipo": "livre",
                "texto": "",
                "subtexto": ""
            }

            if eh_limpeza(dia_idx, tempo, modo_limp):
                celula["tipo"] = "limpeza"
                celula["texto"] = "🧹 LIMPEZA / MANUTENÇÃO"
                celula["subtexto"] = "Horário Técnico Protegido"
            elif (dia_idx, tempo) in grade_base:
                celula["tipo"] = "grade_fixa"
                celula["texto"] = f"📚 {grade_base[(dia_idx, tempo)]}"
                celula["subtexto"] = "Aula Regular Fixa"
            elif (data_dia_iso, tempo) in reservas:
                autor, mot = reservas[(data_dia_iso, tempo)]
                celula["tipo"] = "reservado"
                celula["texto"] = f"🔒 {autor}"
                celula["subtexto"] = mot if mot else "Reserva Confirmada"
            else:
                celula["tipo"] = "livre"
                celula["texto"] = "[  ] Ass: ____________________"
                celula["subtexto"] = "Disponível para Reserva"

            linha_tempo["dias"].append(celula)
        matriz.append(linha_tempo)

    # 1. Gera Markdown
    md_linhas = [
        "# INSTITUTO ESTADUAL DE EDUCAÇÃO, CIÊNCIA E TECNOLOGIA DO MARANHÃO — IEMA",
        "## Controle e Gabarito de Ocupação do Laboratório de Informática",
        f"**Semana de Referência:** {periodo_str} | **Emitido em:** {timestamp_geracao}",
        "\n> **INSTRUÇÃO:** Assine no campo correspondente para reservar horários livres. Horários marcados com `[🧹]` ou `[📚]` não podem ser sobrescritos sem autorização da coordenação.",
        "\n| Horário | Segunda | Terça | Quarta | Quinta | Sexta |",
        "| :--- | :--- | :--- | :--- | :--- | :--- |"
    ]

    for row in matriz:
        cols = [f"**{row['tempo']}º ({row['horario']})**"]
        for c in row["dias"]:
            if c["tipo"] == "livre":
                cols.append("`[  ]` Ass: __________")
            else:
                cols.append(c["texto"])
        md_linhas.append("| " + " | ".join(cols) + " |")

    md_texto = "\n".join(md_linhas)

    # 2. Gera HTML estilizado para A4 Paisagem (Landscape)
    html_template = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Gabarito de Laboratório - {periodo_str}</title>
<style>
  @page {{
    size: A4 landscape;
    margin: 10mm 12mm 10mm 12mm;
  }}
  body {{
    font-family: 'Helvetica Neue', Arial, sans-serif;
    color: #222;
    margin: 0;
    padding: 0;
    font-size: 11px;
  }}
  .cabecalho {{
    text-align: center;
    border-bottom: 2px solid #0d47a1;
    padding-bottom: 6px;
    margin-bottom: 8px;
  }}
  .cabecalho h1 {{
    margin: 0 0 2px 0;
    font-size: 15px;
    color: #0d47a1;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }}
  .cabecalho h2 {{
    margin: 0;
    font-size: 12px;
    color: #444;
    font-weight: normal;
  }}
  .meta {{
    display: flex;
    justify-content: space-between;
    font-size: 10px;
    color: #666;
    margin-bottom: 6px;
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 4px;
    page-break-inside: avoid;
  }}
  th, td {{
    border: 1px solid #aaa;
    padding: 5px 6px;
    text-align: left;
    vertical-align: middle;
  }}
  th {{
    background-color: #0d47a1;
    color: #fff;
    font-weight: 600;
    font-size: 11px;
    text-align: center;
  }}
  .col-tempo {{
    width: 11%;
    background-color: #f0f4f8;
    font-weight: bold;
    text-align: center;
    font-size: 10px;
  }}
  .col-dia {{
    width: 17.8%;
  }}
  .celula-livre {{
    background-color: #ffffff;
    height: 38px;
  }}
  .celula-livre .box-ass {{
    color: #333;
    font-size: 10px;
    border-bottom: 1px dotted #888;
    padding-bottom: 2px;
    display: block;
    margin-top: 2px;
  }}
  .celula-limpeza {{
    background-color: #fbe9e7;
    color: #c62828;
    font-weight: bold;
    font-size: 10px;
    text-align: center;
  }}
  .celula-grade {{
    background-color: #e8f5e9;
    color: #2e7d32;
    font-size: 10px;
    font-weight: 600;
  }}
  .celula-reservado {{
    background-color: #fff8e1;
    color: #f57f17;
    font-size: 10px;
    font-weight: 600;
  }}
  .subtexto {{
    font-size: 8.5px;
    color: #555;
    font-weight: normal;
    display: block;
  }}
  .rodape {{
    margin-top: 8px;
    font-size: 9px;
    color: #666;
    border-top: 1px solid #ccc;
    padding-top: 4px;
    display: flex;
    justify-content: space-between;
  }}
</style>
</head>
<body>
  <div class="cabecalho">
    <h1>IEMA • Instituto Estadual de Educação, Ciência e Tecnologia do Maranhão</h1>
    <h2>Gabarito Oficial de Controle e Reserva do Laboratório de Informática</h2>
  </div>
  <table style="width:100%; border:none; margin-bottom:4px;">
    <tr style="border:none;">
      <td style="border:none; padding:0; font-size:10px; font-weight:bold;">Período: {periodo_str}</td>
      <td style="border:none; padding:0; text-align:right; font-size:10px; color:#555;">Documento gerado em: {timestamp_geracao}</td>
    </tr>
  </table>
  <table>
    <thead>
      <tr>
        <th class="col-tempo">Horário</th>
        <th class="col-dia">Segunda ({segunda_dt.strftime('%d/%m')})</th>
        <th class="col-dia">Terça ({(segunda_dt + timedelta(days=1)).strftime('%d/%m')})</th>
        <th class="col-dia">Quarta ({(segunda_dt + timedelta(days=2)).strftime('%d/%m')})</th>
        <th class="col-dia">Quinta ({(segunda_dt + timedelta(days=3)).strftime('%d/%m')})</th>
        <th class="col-dia">Sexta ({sexta_dt.strftime('%d/%m')})</th>
      </tr>
    </thead>
    <tbody>
"""

    for row in matriz:
        html_template += f"""      <tr>
        <td class="col-tempo">
          {row['tempo']}º Tempo<br>
          <span style="font-weight:normal; font-size:9px;">{row['horario']}</span>
        </td>
"""
        for c in row["dias"]:
            classe_css = f"celula-{c['tipo']}"
            if c["tipo"] == "livre":
                conteudo_td = '<span class="box-ass">Ass: ___________________________</span><span class="subtexto">Docente / Disciplina</span>'
            else:
                conteudo_td = f"{c['texto']}<span class=\"subtexto\">{c['subtexto']}</span>"

            html_template += f'        <td class="{classe_css}">{conteudo_td}</td>\n'
        html_template += "      </tr>\n"

    html_template += f"""    </tbody>
  </table>
  <div class="rodape">
    <span>📌 <b>Regras:</b> Horários assinalados manualmente na folha física devem ser sincronizados no bot via <code>!papel</code>.</span>
    <span>Sincronização WhatsApp Bot: (99) 8165-0963</span>
  </div>
</body>
</html>
"""

    # Nomes dos arquivos de saída
    slug_data = segunda_dt.strftime("%Y%m%d")
    caminho_md = f"{DIR_GABARITOS}/Gabarito_Lab_Semana_{slug_data}.md"
    caminho_html = f"{DIR_GABARITOS}/Gabarito_Lab_Semana_{slug_data}.html"
    caminho_pdf = f"{DIR_GABARITOS}/Gabarito_Lab_Semana_{slug_data}.pdf"

    with open(caminho_md, "w", encoding="utf-8") as f:
        f.write(md_texto)

    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(html_template)

    # Gera PDF com WeasyPrint se disponível
    weasy_bin = shutil.which("weasyprint") or "weasyprint"
    gerou_pdf = False
    if os.path.exists(weasy_bin):
        try:
            cmd = [weasy_bin, caminho_html, caminho_pdf]
            subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True, timeout=20)
            gerou_pdf = True
        except Exception as e:
            print(f"Aviso weasyprint: {e}", file=sys.stderr)

    return {
        "periodo": periodo_str,
        "caminho_md": caminho_md,
        "caminho_html": caminho_html,
        "caminho_pdf": caminho_pdf if gerou_pdf else None,
        "matriz": matriz
    }


def sincronizar_reserva_papel(dia_str: str, tempos_str: str, prof_nome: str, motivo: str = "Origem: Gabarito de Papel da Porta") -> dict:
    """
    Lança no banco de dados reservas anotadas manualmente na folha de papel.
    Exemplo: sincronizar_reserva_papel('terca', '3,4', 'Prof. Marcos', 'Prática de Redes')
    """
    mapa_dias = {
        "seg": 0, "segunda": 0, "segunda-feira": 0,
        "ter": 1, "terca": 1, "terça": 1, "terca-feira": 1, "terça-feira": 1,
        "qua": 2, "quarta": 2, "quarta-feira": 2,
        "qui": 3, "quinta": 3, "quinta-feira": 3,
        "sex": 4, "sexta": 4, "sexta-feira": 4,
    }

    dia_limpo = dia_str.strip().lower()
    if dia_limpo not in mapa_dias:
        return {"status": "ERRO", "mensagem": f"Dia da semana inválido: '{dia_str}'. Use segunda, terca, quarta, quinta ou sexta."}

    dia_idx = mapa_dias[dia_limpo]
    segunda_dt = obter_segunda_feira_referencia()
    data_alvo = segunda_dt + timedelta(days=dia_idx)
    data_iso = data_alvo.strftime("%Y-%m-%d")
    data_formatada = data_alvo.strftime("%d/%m/%Y")

    # Extrai tempos
    tempos = [int(t) for t in re.findall(r"\b[1-9]\b", tempos_str)]
    if not tempos:
        return {"status": "ERRO", "mensagem": "Informe os horários (de 1 a 9). Ex: 3,4 ou 3 4."}

    with sqlite3.connect(ESCOLA_DB) as conn:
        cur = conn.cursor()
        modo_limp = obter_modo_limpeza(conn)

        # Checa limpeza
        for t in tempos:
            if eh_limpeza(dia_idx, t, modo_limp):
                return {
                    "status": "CONFLITO",
                    "mensagem": f"⚠️ O {t}º horário na segunda-feira está bloqueado para limpeza técnica do laboratório."
                }

        # Checa conflitos existentes
        placeholders = ",".join(["?"] * len(tempos))
        cur.execute(f"""
            SELECT tempo_aula, autor_nome FROM reservas_laboratorio
            WHERE data_reserva = ? AND status IN ('CONFIRMADO', 'PENDENTE') AND tempo_aula IN ({placeholders})
        """, [data_iso] + tempos)
        conflitos = cur.fetchall()
        if conflitos:
            detalhes = ", ".join([f"{t}º horário por {autor}" for t, autor in conflitos])
            return {
                "status": "CONFLITO",
                "mensagem": f"⚠️ Conflito no papel: O horário já consta reservado no sistema: {detalhes}."
            }

        # Insere reservas do papel
        agora_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        inseridos = []
        for t in tempos:
            ini, fim = GRADE_LAB[t - 1][1], GRADE_LAB[t - 1][2]
            cur.execute("""
                INSERT INTO reservas_laboratorio (
                    data_reserva, tempo_aula, horario_inicio, horario_fim,
                    autor_jid, autor_numero, autor_nome, mensagem_id, motivo, criado_em, status
                ) VALUES (?, ?, ?, ?, 'papel@escola.local', 'PAPEL_FISICO', ?, ?, ?, ?, 'CONFIRMADO')
            """, (data_iso, t, ini, fim, prof_nome, f"PAPEL_{data_iso}_{t}", motivo, agora_iso))
            inseridos.append(t)
        conn.commit()

    tempos_txt = ", ".join([f"{t}º" for t in inseridos])
    return {
        "status": "SUCESSO",
        "mensagem": f"✅ Gabarito sincronizado: {data_formatada} ({dia_str.title()}) nos horários ({tempos_txt}) confirmados para *{prof_nome}*."
    }


def gerar_pdf_grade_semana_acessivel(data_ref: datetime = None, termo: str = "") -> str:
    """
    Gera um PDF em formato A4 Paisagem (Landscape) de ALTA VISIBILIDADE com:
    - Letras e fontes grandes
    - Cores chamativas e alto contraste
    - Pensado para usuários de mais idade e leitura instantânea no celular ou impresso
    - Destaque claro de horários livres, ocupados, aulas e limpeza
    """
    segunda_dt = obter_segunda_feira_referencia(data_ref, termo=termo)
    sexta_dt = segunda_dt + timedelta(days=4)
    hoje_iso = (data_ref or datetime.now()).strftime("%Y-%m-%d")

    segunda_iso = segunda_dt.strftime("%Y-%m-%d")
    sexta_iso = sexta_dt.strftime("%Y-%m-%d")

    periodo_str = f"{segunda_dt.strftime('%d/%m/%Y')} a {sexta_dt.strftime('%d/%m/%Y')}"
    timestamp_geracao = datetime.now().strftime("%d/%m/%Y às %H:%M")

    with sqlite3.connect(ESCOLA_DB) as conn:
        grade_base = carregar_grade_base(conn)
        reservas = carregar_reservas_semana(conn, segunda_iso, sexta_iso)
        modo_limp = obter_modo_limpeza(conn)

    intervalos_map = {
        2: "☕ INTERVALO DA MANHÃ (09:10 às 09:30)",
        5: "🍽️ HORÁRIO DE ALMOÇO (12:00 às 13:30)",
        7: "☕ INTERVALO DA TARDE (15:05 às 15:20)",
    }

    # Cabeçalho da tabela com os dias
    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Grade Semanal do Laboratório - Alta Visibilidade</title>
<style>
  @page {{
    size: A4 landscape;
    margin: 6mm 8mm;
  }}
  body {{
    font-family: 'Segoe UI', Arial, Helvetica, sans-serif;
    color: #111;
    margin: 0;
    padding: 0;
    font-size: 13px;
    background-color: #ffffff;
  }}
  .cabecalho {{
    text-align: center;
    background: linear-gradient(135deg, #0d47a1 0%, #1565c0 100%);
    background-color: #0d47a1;
    color: #ffffff;
    padding: 8px 12px;
    border-radius: 8px;
    margin-bottom: 5px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.15);
  }}
  .cabecalho h1 {{
    margin: 0;
    font-size: 21px;
    letter-spacing: 0.5px;
    font-weight: 900;
    text-transform: uppercase;
  }}
  .cabecalho .sub {{
    font-size: 13.5px;
    font-weight: 700;
    margin-top: 2px;
    color: #ffeb3b;
  }}
  table {{
    width: 100%;
    border-collapse: separate;
    border-spacing: 3px;
    margin-top: 3px;
    page-break-inside: avoid;
  }}
  th {{
    padding: 7px 5px;
    font-size: 14px;
    font-weight: 900;
    text-align: center;
    border-radius: 5px;
  }}
  .th-tempo {{
    width: 11%;
    background-color: #212121;
    color: #ffffff;
  }}
  .th-dia {{
    width: 17.8%;
    background-color: #0d47a1;
    color: #ffffff;
    border: 2px solid #083273;
  }}
  .th-dia-hoje {{
    background-color: #e65100 !important;
    border: 2.5px solid #bf360c !important;
    color: #ffffff !important;
  }}
  td {{
    border-radius: 6px;
    padding: 5px 6px;
    vertical-align: middle;
    text-align: center;
    box-sizing: border-box;
  }}
  .td-tempo {{
    background-color: #eceff1;
    color: #212121;
    font-weight: 900;
    font-size: 12px;
    border: 2px solid #b0bec5;
  }}
  .card-livre {{
    background-color: #e8f5e9;
    border: 2.5px solid #2e7d32;
  }}
  .card-livre .badge {{
    font-size: 13.5px;
    font-weight: 900;
    color: #1b5e20;
    letter-spacing: 0.3px;
  }}
  .card-livre .sub-txt {{
    font-size: 11px;
    font-weight: 700;
    color: #2e7d32;
    margin-top: 1px;
  }}
  .card-reservado {{
    background-color: #fff8e1;
    border: 3px solid #ff8f00;
  }}
  .card-reservado .badge {{
    font-size: 12.5px;
    font-weight: 900;
    color: #c62828;
    text-transform: uppercase;
  }}
  .card-reservado .prof {{
    font-size: 12px;
    font-weight: 800;
    color: #212121;
    margin-top: 1px;
  }}
  .card-reservado .motivo {{
    font-size: 10.5px;
    font-weight: 700;
    color: #424242;
    margin-top: 1px;
  }}
  .card-grade {{
    background-color: #e3f2fd;
    border: 2.5px solid #1976d2;
  }}
  .card-grade .badge {{
    font-size: 12px;
    font-weight: 800;
    color: #0d47a1;
  }}
  .card-grade .prof {{
    font-size: 11.5px;
    font-weight: 800;
    color: #212121;
  }}
  .card-limpeza {{
    background-color: #ffebee;
    border: 2.5px solid #d32f2f;
  }}
  .card-limpeza .badge {{
    font-size: 12.5px;
    font-weight: 900;
    color: #c62828;
  }}
  .card-limpeza .sub-txt {{
    font-size: 10.5px;
    font-weight: 700;
    color: #b71c1c;
  }}
  .row-intervalo td {{
    background-color: #37474f !important;
    color: #ffffff !important;
    font-size: 12px !important;
    font-weight: 900 !important;
    padding: 3px 6px !important;
    text-align: center !important;
    letter-spacing: 0.8px;
    border: 1px solid #263238 !important;
  }}
  .legenda {{
    margin-top: 5px;
    padding: 5px 10px;
    background-color: #f5f5f5;
    border: 2px solid #cfd8dc;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 800;
    text-align: center;
  }}
  .legenda-item {{
    display: inline-block;
    padding: 2px 10px;
    border-radius: 4px;
    margin: 0 6px;
    border: 1.5px solid;
  }}
  .leg-livre {{ background:#e8f5e9; border-color:#2e7d32; color:#1b5e20; }}
  .leg-res {{ background:#fff8e1; border-color:#ff8f00; color:#bf360c; }}
  .leg-grade {{ background:#e3f2fd; border-color:#1976d2; color:#0d47a1; }}
  .leg-limp {{ background:#ffebee; border-color:#d32f2f; color:#c62828; }}
</style>
</head>
<body>
  <div class="cabecalho">
    <h1>🖥️ LABORATÓRIO IEMA TI — GRADE SEMANAL DE HORÁRIOS</h1>
    <div class="sub">📅 SEMANA: {periodo_str} &nbsp;|&nbsp; EMISSÃO: {timestamp_geracao}</div>
  </div>

  <table>
    <thead>
      <tr>
        <th class="th-tempo">HORÁRIO</th>
"""

    for nome_dia, dia_idx in DIAS_SEMANA:
        dt_col = segunda_dt + timedelta(days=dia_idx)
        dt_col_iso = dt_col.strftime("%Y-%m-%d")
        classe_th = "th-dia th-dia-hoje" if dt_col_iso == hoje_iso else "th-dia"
        hoje_aviso = " (HOJE)" if dt_col_iso == hoje_iso else ""
        html += f'        <th class="{classe_th}">{nome_dia.upper()}<br><span style="font-size:12px; font-weight:700;">{dt_col.strftime("%d/%m")}{hoje_aviso}</span></th>\n'

    html += """      </tr>
    </thead>
    <tbody>
"""

    for tempo, ini, fim in GRADE_LAB:
        html += f"""      <tr>
        <td class="td-tempo">
          <span style="font-size:13.5px;">{tempo}º TEMPO</span><br>
          <span style="font-size:11px; font-weight:700; color:#37474f;">{ini} às {fim}</span>
        </td>
"""
        for nome_dia, dia_idx in DIAS_SEMANA:
            data_dia = segunda_dt + timedelta(days=dia_idx)
            data_dia_iso = data_dia.strftime("%Y-%m-%d")

            if eh_limpeza(dia_idx, tempo, modo_limp):
                td_content = (
                    '<td class="card-limpeza">'
                    '<div class="badge">🧹 LIMPEZA</div>'
                    '<div class="sub-txt">Manutenção Técnica</div>'
                    '</td>'
                )
            elif (data_dia_iso, tempo) in reservas:
                autor, mot = reservas[(data_dia_iso, tempo)]
                mot_html = f'<div class="motivo">{mot}</div>' if mot else ''
                td_content = (
                    f'<td class="card-reservado">'
                    f'<div class="badge">🔒 RESERVADO</div>'
                    f'<div class="prof">{autor}</div>'
                    f'{mot_html}'
                    f'</td>'
                )
            elif (dia_idx, tempo) in grade_base:
                aula_txt = grade_base[(dia_idx, tempo)]
                td_content = (
                    f'<td class="card-grade">'
                    f'<div class="badge">📚 AULA REGULAR</div>'
                    f'<div class="prof">{aula_txt}</div>'
                    f'</td>'
                )
            else:
                td_content = (
                    '<td class="card-livre">'
                    '<div class="badge">🟢 LIVRE</div>'
                    '<div class="sub-txt">Disponível</div>'
                    '</td>'
                )
            html += f"        {td_content}\n"

        html += "      </tr>\n"

        # Linha destacada de intervalo/almoço após determinados tempos
        if tempo in intervalos_map:
            html += f"""      <tr class="row-intervalo">
        <td colspan="6">{intervalos_map[tempo]}</td>
      </tr>
"""

    html += """    </tbody>
  </table>

  <div class="legenda">
    <span style="margin-right:8px; color:#212121;">🎨 <b>GUIA VISUAL DE CORES:</b></span>
    <span class="legenda-item leg-livre">🟢 VERDE = HORÁRIO LIVRE</span>
    <span class="legenda-item leg-res">🟡 LARANJA = HORÁRIO RESERVADO</span>
    <span class="legenda-item leg-grade">🔵 AZUL = AULA REGULAR DA GRADE</span>
    <span class="legenda-item leg-limp">🔴 VERMELHO = LIMPEZA TÉCNICA</span>
  </div>
</body>
</html>
"""

    slug_data = segunda_dt.strftime("%Y%m%d")
    caminho_html = f"{DIR_GABARITOS}/Grade_Semana_Alta_Visibilidade_{slug_data}.html"
    caminho_pdf = f"{DIR_GABARITOS}/Grade_Semana_Alta_Visibilidade_{slug_data}.pdf"
    caminho_publico_pdf = f"{HOME}/aulas_listas/Grade_Semanal_Laboratorio_IEMA.pdf"

    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(html)

    weasy_bin = shutil.which("weasyprint") or "weasyprint"
    if os.path.exists(weasy_bin):
        try:
            cmd = [weasy_bin, caminho_html, caminho_pdf]
            subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True, timeout=25)
            shutil.copy(caminho_pdf, caminho_publico_pdf)
            return caminho_publico_pdf
        except Exception as e:
            print(f"Erro ao compilar weasyprint: {e}", file=sys.stderr)

    return caminho_html


if __name__ == "__main__":
    print("🚀 Gerando Gabarito e Grade Acessível de Alta Visibilidade...")
    res = gerar_gabarito_semanal()
    pdf_acessivel = gerar_pdf_grade_semana_acessivel()
    print(f"Período: {res['periodo']}")
    print(f"PDF Acessível Alta Visibilidade: {pdf_acessivel}")
