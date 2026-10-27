import csv
import os

arquivo = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas/automacao-iema/aulas.csv")

def upsert_aulas(novas_aulas):
    linhas_existentes = []
    if os.path.exists(arquivo):
        with open(arquivo, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            linhas_existentes = list(reader)
    
    # Criar dict por chave única
    mapa_aulas = {}
    for linha in linhas_existentes:
        chave = (linha.get('data',''), linha.get('turma',''), linha.get('horarios',''))
        mapa_aulas[chave] = linha
        
    # Atualiza ou insere as novas
    for nova in novas_aulas:
        chave = (nova.get('data',''), nova.get('turma',''), nova.get('horarios',''))
        mapa_aulas[chave] = nova

    # Salva tudo de volta
    with open(arquivo, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['registrar','data','componente','turma','horarios','modalidade','tema','conteudo','recursos'], delimiter=';')
        writer.writeheader()
        for v in mapa_aulas.values():
            writer.writerow(v)
