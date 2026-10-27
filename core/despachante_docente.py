#!/usr/bin/env python3
"""
Despachante: pega tarefas do Observador (fila/achado_N.md) em tarefas_agentes
com status PRONTO_PARA_EXECUCAO e dispara um agente `agy` não interativo por tarefa.

  despachante_docente.py                 # dry-run: lista o que seria executado
  despachante_docente.py --run [--max N] [--id TAREFA_ID] [--auto]

Sem --auto o agente roda em modo accept-edits (edições de arquivo ok, comandos
pedem permissão -> em print mode ficam negados). Com --auto usa
--dangerously-skip-permissions (use só para tarefas confiáveis).
O fechamento do ciclo é do Observador: se o achado não reaparecer, vira RESOLVIDO.
"""
import argparse, datetime, os, sqlite3, subprocess

HOME = os.environ.get("DOCENTE_HOME", ".")
DB = f"{HOME}/escola.db"
FILA = f"{HOME}/projetos_ideias/observador/fila"
ap = argparse.ArgumentParser()
ap.add_argument("--run", action="store_true")
ap.add_argument("--max", type=int, default=1)
ap.add_argument("--id", type=int)
ap.add_argument("--auto", action="store_true")
ap.add_argument("--timeout", type=int, default=600)
a = ap.parse_args()

c = sqlite3.connect(DB, timeout=30)
q = ("select id,titulo_tarefa,prompt_agente,artefato_esperado from tarefas_agentes "
     "where status='PRONTO_PARA_EXECUCAO' and titulo_tarefa like '[Observador]%'")
if a.id:
    q += f" and id={int(a.id)}"
tarefas = c.execute(q + " order by id limit ?", (a.max if a.run else 50,)).fetchall()
print(f"{len(tarefas)} tarefa(s) do Observador na fila.")

for tid, titulo, prompt, arq in tarefas:
    print(f"- #{tid} {titulo}")
    if not a.run:
        continue
    # reserva atômica: evita duas execuções simultâneas da mesma tarefa
    cur = c.execute("update tarefas_agentes set status='EM_EXECUCAO' where id=? and status='PRONTO_PARA_EXECUCAO'", (tid,))
    c.commit()
    if cur.rowcount == 0:
        continue
    cmd = ["agy", "-p", prompt, "--mode", "accept-edits"]
    if a.auto:
        cmd.append("--dangerously-skip-permissions")
    saida = f"{FILA}/saida_{tid}.md"
    t0 = datetime.datetime.now()
    rc = None
    # saída em arquivo (não pipe): filhos em background (ex.: daemons iniciados pelo agente)
    # não seguram o despachante; no timeout mata só o agente, não os filhos.
    tmp = f"{FILA}/.saida_{tid}.tmp"
    try:
        with open(tmp, "w") as fo:
            p = subprocess.Popen(cmd, cwd=HOME, stdin=subprocess.DEVNULL, stdout=fo,
                                 stderr=subprocess.STDOUT, start_new_session=True)
            try:
                rc = p.wait(timeout=a.timeout)
                status = "CONCLUIDA" if rc == 0 else "FALHOU"
            except subprocess.TimeoutExpired:
                p.kill()
                status = "FALHOU"
        corpo = open(tmp, errors="replace").read() + (f"\n[timeout {a.timeout}s]" if status == "FALHOU" and rc is None else "")
    except Exception as e:
        status, corpo = "FALHOU", str(e)
    open(saida, "w", encoding="utf-8").write(
        f"# Saída da tarefa #{tid} ({status}) — {t0:%d/%m %H:%M}, {(datetime.datetime.now()-t0).seconds}s\n\n{corpo}\n")
    c.execute("update tarefas_agentes set status=? where id=?", (status, tid))
    c.commit()
    print(f"  -> {status}, saída em {saida}")
    try:
        subprocess.run(["python3", f"{HOME}/avisar_zap.py", f"🤖 Tarefa #{tid} {status}: {titulo[:70]}"], timeout=30)
    except Exception:
        pass
c.close()
