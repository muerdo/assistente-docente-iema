#!/usr/bin/env python3
import sqlite3, os, csv

DB_PATH = os.path.expanduser("~/escola.db")
NOTAS_CSV = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas/automacao-iema/dados_notas.csv")
PLANILHA_CSV = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas/automacao-iema/planilha_notas_3bimestre.csv")

def setup_database():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS grade_horaria (id INTEGER PRIMARY KEY AUTOINCREMENT, dia_semana INTEGER, hora_inicio TEXT, hora_fim TEXT, turma TEXT, sala TEXT, disciplina TEXT, tempos TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS plano_curso (id INTEGER PRIMARY KEY AUTOINCREMENT, disciplina TEXT, turma TEXT, ordem_aula INTEGER, topico TEXT, detalhes TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS diario_classe (id INTEGER PRIMARY KEY AUTOINCREMENT, data_registro TEXT, turma TEXT, disciplina TEXT, conteudo_ministrado TEXT, observacoes TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS agendamentos (id INTEGER PRIMARY KEY AUTOINCREMENT, prompt TEXT, horario_iso TEXT, status TEXT DEFAULT 'PENDENTE', chat_id TEXT DEFAULT '')")
    cur.execute("""
    CREATE TABLE IF NOT EXISTS alunos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        matricula TEXT,
        nome TEXT,
        turma TEXT,
        componente TEXT,
        faltas INTEGER DEFAULT 0,
        av1 TEXT,
        av2 TEXT,
        av3 TEXT,
        UNIQUE(matricula, turma)
    )
    """)
    
    cur.execute("PRAGMA table_info(grade_horaria)")
    cols = [col[1] for col in cur.fetchall()]
    if "tempos" not in cols:
        cur.execute("ALTER TABLE grade_horaria ADD COLUMN tempos TEXT DEFAULT ''")

    cur.execute("DELETE FROM grade_horaria")
    grade_oficial = [
        (0, "10:20", "12:00", "201", "Sala 201", "INTELIGÊNCIA ARTIFICIAL", "4ª, 5ª"),
        (0, "13:30", "15:10", "200", "Sala 200", "INTELIGÊNCIA ARTIFICIAL", "6ª, 7ª"),
        (0, "16:10", "17:00", "202", "Sala 202", "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL", "9ª"),
        (1, "15:20", "17:00", "202", "Sala 202", "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL", "8ª, 9ª"),
        (2, "15:20", "17:00", "203", "Sala 203", "INTELIGÊNCIA ARTIFICIAL", "8ª, 9ª"),
        (3, "13:30", "15:10", "202", "Sala 202", "INTELIGÊNCIA ARTIFICIAL", "6ª, 7ª"),
        (4, "09:30", "11:10", "202", "Sala 202", "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL", "3ª, 4ª"),
    ]
    cur.executemany("INSERT INTO grade_horaria (dia_semana, hora_inicio, hora_fim, turma, sala, disciplina, tempos) VALUES (?, ?, ?, ?, ?, ?, ?)", grade_oficial)

    # Popula alunos se tabela estiver vazia
    cur.execute("SELECT COUNT(*) FROM alunos")
    if cur.fetchone()[0] == 0 and os.path.exists(PLANILHA_CSV):
        faltas_map = {}
        if os.path.exists(NOTAS_CSV):
            with open(NOTAS_CSV, "r", encoding="utf-8") as f:
                for r in csv.DictReader(f):
                    faltas_map[r["Matricula"].strip()] = int(r.get("Faltas", 0) or 0)

        alunos_unicos = {}
        with open(PLANILHA_CSV, "r", encoding="utf-8") as f:
            for r in csv.DictReader(f, delimiter=";"):
                mat = r["matricula"].strip()
                turma = r["turma"].strip()
                nome = r["nome"].strip()
                comp = r["componente"].strip()
                chave = (turma, mat)
                if chave not in alunos_unicos:
                    alunos_unicos[chave] = {
                        "matricula": mat, "nome": nome, "turma": turma, "componente": comp,
                        "faltas": faltas_map.get(mat, 0),
                        "av1": r.get("AV1", "").strip(),
                        "av2": r.get("AV2", "").strip(),
                        "av3": r.get("AV3", "").strip()
                    }
                else:
                    if comp not in alunos_unicos[chave]["componente"]:
                        alunos_unicos[chave]["componente"] += f", {comp}"

        for a in alunos_unicos.values():
            cur.execute("""
                INSERT OR REPLACE INTO alunos (matricula, nome, turma, componente, faltas, av1, av2, av3)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (a["matricula"], a["nome"], a["turma"], a["componente"], a["faltas"], a["av1"], a["av2"], a["av3"]))

    conn.commit()
    conn.close()
    print(f"[+] Banco pedagógico atualizado com sucesso em: {DB_PATH}")

if __name__ == "__main__":
    setup_database()
