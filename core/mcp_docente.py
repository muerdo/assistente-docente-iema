#!/usr/bin/env python3
import os, sqlite3
from datetime import datetime
from mcp.server.mcpserver import MCPServer
from pptx import Presentation

mcp = MCPServer("TeacherAssistantMCP")
DB_PATH = os.path.expanduser("~/escola.db")
GDRIVE_AULAS_DIR = os.path.expanduser("~/mounts/gdrive/Aulas")

@mcp.tool()
def consultar_sala_e_agenda() -> str:
    if not os.path.exists(DB_PATH): return "Erro: Base escolar (escola.db) não encontrada."
    agora = datetime.now()
    dia_semana, hora_atual = agora.weekday(), agora.strftime("%H:%M")
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT hora_inicio, hora_fim, turma, sala, disciplina FROM grade_horaria WHERE dia_semana = ? ORDER BY hora_inicio ASC", (dia_semana,))
    aulas = cur.fetchall()
    conn.close()
    if not aulas: return f"Você não tem aulas cadastradas para hoje."
    for h_ini, h_fim, turma, sala, disc in aulas:
        if h_ini <= hora_atual <= h_fim: return f"🚨 [AULA ATUAL EM ANDAMENTO]\n• Horário: {h_ini} às {h_fim}\n• Turma: {turma}\n• Sala/Local: {sala}\n• Disciplina: {disc}"
        elif hora_atual < h_ini: return f"⏳ [PRÓXIMA AULA]\n• Horário: {h_ini} às {h_fim}\n• Turma: {turma}\n• Sala/Local: {sala}\n• Disciplina: {disc}"
    resumo = "\n".join([f"- {h_ini}-{h_fim}: {turma} ({sala}) - {disc}" for h_ini, h_fim, turma, sala, disc in aulas])
    return f"Aulas do dia já encerradas.\nQuadro de hoje:\n{resumo}"

@mcp.tool()
def registrar_diario_classe(turma: str, disciplina: str, conteudo: str, observacoes: str = "") -> str:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    data_hoje = datetime.now().strftime("%Y-%m-%d %H:%M")
    cur.execute("INSERT INTO diario_classe (data_registro, turma, disciplina, conteudo_ministrado, observacoes) VALUES (?, ?, ?, ?, ?)", (data_hoje, turma, disciplina, conteudo, observacoes))
    conn.commit()
    conn.close()
    return f"✅ Registro concluído com sucesso para a {turma} ({disciplina}) em {data_hoje}."

@mcp.tool()
def gerar_apresentacao_powerpoint(tema: str, turma: str, slides_json: list) -> dict:
    os.makedirs(GDRIVE_AULAS_DIR, exist_ok=True)
    nome_arquivo = f"Aula_{tema.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
    caminho_final = os.path.join(GDRIVE_AULAS_DIR, nome_arquivo)
    prs = Presentation()
    slide_capa = prs.slides.add_slide(prs.slide_layouts[0])
    slide_capa.shapes.title.text = tema
    slide_capa.placeholders[1].text = f"IEMA Pleno Açailândia • Material de Aula • {turma}\nGerado em: {datetime.now().strftime('%d/%m/%Y')}"
    for slide_data in slides_json:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = slide_data.get("titulo", "Tópico")
        tf = slide.placeholders[1].text_frame
        for i, ponto in enumerate(slide_data.get("pontos", [])):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = ponto
    prs.save(caminho_final)
    return {"status": "criado", "caminho_local": caminho_final, "nome_arquivo": nome_arquivo}

import csv

@mcp.tool()
def atualizar_aulas_csv_ibutumy(aulas: list) -> str:
    """
    Adiciona novas aulas ao aulas.csv do script de automação do Ibutumy ou atualiza se já existir.
    O argumento 'aulas' deve ser uma lista de dicionários com as chaves:
    registrar, data, componente, turma, horarios, modalidade, tema, conteudo, recursos.
    """
    arquivo = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas/automacao-iema/aulas.csv")
    import os
    linhas_existentes = []
    if os.path.exists(arquivo):
        with open(arquivo, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            linhas_existentes = list(reader)
    
    mapa_aulas = {}
    for linha in linhas_existentes:
        chave = (linha.get('data',''), linha.get('turma',''), linha.get('horarios',''))
        mapa_aulas[chave] = linha
        
    for nova in aulas:
        chave = (nova.get('data',''), nova.get('turma',''), nova.get('horarios',''))
        mapa_aulas[chave] = nova

    with open(arquivo, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['registrar','data','componente','turma','horarios','modalidade','tema','conteudo','recursos'], delimiter=';')
        writer.writeheader()
        for v in mapa_aulas.values():
            writer.writerow(v)
            
    return f"✅ {len(aulas)} aula(s) processada(s) (inseridas ou atualizadas sem duplicar) no {arquivo}. Quando quiser fechar o diário, rode 'python registrar_aulas.py' no terminal da sua máquina."

@mcp.tool()
def atualizar_notas_csv_ibutumy(notas: list, sobrescrever: bool = True) -> str:
    """
    Sobrescreve ou atualiza o dados_notas.csv do Ibutumy.
    O argumento 'notas' deve ser uma lista de dicionários com as chaves:
    Matricula, Nome, Nota, Faltas.
    """
    arquivo = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas/automacao-iema/dados_notas.csv")
    modo = 'w' if sobrescrever else 'a'
    with open(arquivo, modo, newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['Matricula','Nome','Nota','Faltas'])
        if sobrescrever or os.path.getsize(arquivo) == 0:
            writer.writeheader()
        for nota in notas:
            writer.writerow(nota)
    return f"✅ {len(notas)} registro(s) salvo(s) no {arquivo}. Rode 'python bot.py' no seu terminal para preencher as notas e faltas no sistema."

@mcp.tool()
def consultar_aulas_csv_ibutumy() -> str:
    """
    Lê o aulas.csv para listar os horários, turmas, temas e conteúdos cadastrados para o Ibutumy.
    """
    arquivo = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas/automacao-iema/aulas.csv")
    if not os.path.exists(arquivo): return "Erro: aulas.csv não encontrado."
    with open(arquivo, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=';')
        aulas = list(reader)
    if not aulas: return "O arquivo aulas.csv está vazio."
    resumo = "\n".join([f"- {a['data']} ({a['horarios']}): {a['turma']} | Tema: {a['tema']} | Cont: {a['conteudo']}" for a in aulas])
    return f"Conteúdos e Turmas no CSV:\n{resumo}"

@mcp.tool()
def consultar_notas_csv_ibutumy(nome_aluno: str = "") -> str:
    """
    Lê o dados_notas.csv para consultar a nota e as faltas de um aluno (ou listar todos se não especificado).
    """
    arquivo = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas/automacao-iema/dados_notas.csv")
    if not os.path.exists(arquivo): return "Erro: dados_notas.csv não encontrado."
    with open(arquivo, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        notas = list(reader)
    if not notas: return "Nenhuma nota/falta cadastrada no CSV."
    
    if nome_aluno:
        filtrado = [n for n in notas if nome_aluno.lower() in n.get('Nome', '').lower()]
        if not filtrado: return f"Nenhum aluno encontrado contendo '{nome_aluno}'."
        notas = filtrado
        
    resumo = "\n".join([f"- {n.get('Nome', '')}: Nota {n.get('Nota', '')}, Faltas {n.get('Faltas', '')}" for n in notas])
    return f"Resultado da consulta de notas:\n{resumo}"

@mcp.tool()
def listar_alunos_turma(turma: str) -> str:
    """
    Lista todos os alunos de uma turma específica (200, 201, 202, 203) com matrículas e faltas atuais.
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT matricula, nome, turma, componente, faltas FROM alunos WHERE turma = ? ORDER BY nome ASC", (turma.strip(),))
    alunos = cur.fetchall()
    conn.close()
    if not alunos: return f"Turma {turma} não encontrada ou sem alunos cadastrados."
    linhas = [f"{idx:2d}. {nome} ({mat}) - Faltas: {faltas}" for idx, (mat, nome, t, c, faltas) in enumerate(alunos, 1)]
    return f"Alunos da Turma {turma} ({len(alunos)} total):\n" + "\n".join(linhas)

@mcp.tool()
def registrar_frequencia(turma: str, aluno_ou_num: str, operacao: str = "falta", data: str = "") -> str:
    """
    Atribui falta ou presença/abono para um aluno de uma turma específica (pelo número da chamada ou nome).
    """
    import sys
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    return webhook_zap.alterar_frequencia_aluno(aluno_ou_num, turma, data, operacao)

@mcp.tool()
def registrar_visto_atividade(comando_texto: str) -> str:
    """
    Registra entrega de atividade ou visto de caderno (individual, lote ou 'todos exceto').
    Suporta indicação explícita da atividade via #<ID>, data (DD/MM) ou tema:<palavra>.
    Exemplos de entrada:
    'atv #6 202 todos ok exceto 2, 3, 10, 27, 28'
    'atv #5 202 todos menos ana sophia, davy'
    'atv #6 202 1, 4, 8-12 ok'
    'atv #6 202 pedro ok'
    'atv mover #5 #6'
    """
    import sys
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    return webhook_zap.processar_comando_atividade(comando_texto)

@mcp.tool()
def consultar_vistos_atividades(turma_ou_aluno: str) -> str:
    """
    Consulta os vistos e atividades entregues de uma turma específica (ex: '202') ou de um aluno (ex: 'pedro').
    """
    import sys
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    t = turma_ou_aluno.strip()
    if t in ["200", "201", "202", "203", "300", "301", "302", "303"]:
        return webhook_zap.consultar_atividades_turma(t)
    return webhook_zap.consultar_atividades_aluno(t)

@mcp.tool()
def cadastrar_atividade(turma: str, componente: str, titulo: str, data_passada: str = "") -> str:
    """
    Cadastra uma nova atividade passada para a turma com título e a data em que foi passada.
    """
    import sys
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    cmd = f"{turma} {componente} {titulo} {data_passada}".strip()
    return webhook_zap.cadastrar_nova_atividade(cmd)

@mcp.tool()
def cobrar_atividades(turma: str, atividade: str = "") -> str:
    """
    Cruza todos os alunos matriculados na turma com os que receberam visto,
    destacando quem está pendente, a atividade e a data em que foi passada.
    """
    import sys
    sys.path.insert(0, os.environ.get("DOCENTE_HOME", "."))
    import webhook_zap
    cmd = f"{turma} {atividade}".strip()
    return webhook_zap.cobrar_atividades_turma(cmd)

if __name__ == "__main__": mcp.run()

