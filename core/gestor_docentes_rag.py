"""
Gestor Multi-Docente e Motor RAG com Isolamento Estrito por Disciplina (Anti-Alucinação).

Organiza os perfis, grades de horários individuais e bases de conhecimento (RAG)
de múltiplos professores em `docentes/<chave>/`:
- Cada disciplina possui seu próprio subdiretório isolado em:
  `docentes/<chave>/rag_conteudos/<materia_id>/`
- Evita contaminação cruzada entre disciplinas do mesmo professor (ex.: IA vs Legislação)
  e garante que a geração de aulas use estritamente os guias/ementas anexados àquela matéria.
"""

import json
import os
import re
import sqlite3
import unicodedata
import subprocess
import shutil
from datetime import datetime, timedelta

BASE_ECOSSISTEMA_DIR = os.environ.get(
    "ECOSSISTEMA_DIR", os.path.join(os.environ.get("DOCENTE_HOME", "."), "ecossistema_docente")
)
DOCENTES_DIR = os.path.join(BASE_ECOSSISTEMA_DIR, "docentes")
CATALOGO_PATH = os.path.join(DOCENTES_DIR, "catalogo_docentes.json")
ESCOLA_DB = os.environ.get("ESCOLA_DB", os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db"))
PASTA_UPLOADS_ISOLADOS = os.environ.get(
    "PASTA_UPLOADS_ISOLADOS",
    os.path.join(BASE_ECOSSISTEMA_DIR, "uploads_isolados")
)

BASE_TECNICA_PERMITIDOS = {
    "pedro": ["pedro", "prof pedro", "pedro lucas", "professor pedro"],
    "leandro": ["leandro", "prof leandro", "professor leandro"],
    "bruna": ["bruna", "brunna", "profa bruna", "professora bruna", "brunna tavares"],
    "marcos": ["marcos", "marcos sobral", "prof marcos", "professor marcos"],
    "assis": ["assis", "francisco de assis", "prof assis", "professor assis"],
    "patricia": ["patricia", "profa patricia", "professora patricia"],
    "wanderson": ["wanderson", "prof wanderson", "professor wanderson"],
    "osenias": ["osenias", "prof osenias", "professor osenias"],
}

DOCENTES_BLOQUEADOS = {
    "carlos": "Prof. Carlos (bloqueado por determinação administrativa)",
}


def validar_docente_permitido(nome_ou_chave: str) -> tuple[bool, str, str | None]:
    """
    Regra Estrita de Segurança e Conformidade:
    - Somente professores da Base Técnica previamente cadastrados são autorizados.
    - O Prof. Carlos está expressamente bloqueado/excluído por determinação administrativa.
    - É estritamente vedada a criação de qualquer outro docente externo.
    """
    n_norm = _normalizar(nome_ou_chave)
    if not n_norm:
        return False, "Nome não informado.", None

    for b in DOCENTES_BLOQUEADOS:
        if b in n_norm:
            return (
                False,
                "⛔ O cadastro para o Prof. Carlos está bloqueado por determinação administrativa.",
                "carlos",
            )

    for ch, aliases in BASE_TECNICA_PERMITIDOS.items():
        if ch in n_norm or any(a in n_norm for a in aliases):
            return True, "Docente autorizado da Base Técnica.", ch

    return (
        False,
        "⛔ Cadastro não permitido. O sistema está restrito exclusivamente aos docentes homologados da Base Técnica já cadastrados no ecossistema (exceto Prof. Carlos). Novos professores externos não podem ser criados.",
        None,
    )

HORARIOS_AULA_PADRAO = {
    1: ("07:30", "08:25"),
    2: ("08:25", "09:10"),
    3: ("09:30", "10:20"),
    4: ("10:20", "11:10"),
    5: ("11:10", "12:00"),
    6: ("13:30", "14:20"),
    7: ("14:20", "15:05"),
    8: ("15:20", "16:10"),
    9: ("16:10", "17:00"),
}

DIAS_SEMANA_NOMES = {
    0: "Segunda-feira",
    1: "Terça-feira",
    2: "Quarta-feira",
    3: "Quinta-feira",
    4: "Sexta-feira",
    5: "Sábado",
    6: "Domingo",
}


def _normalizar(texto: str) -> str:
    if not texto:
        return ""
    t = unicodedata.normalize("NFD", str(texto))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", t).strip().lower()


def slug_disciplina(nome_disc: str) -> str:
    s = re.sub(r"[^a-z0-9]+", "_", _normalizar(nome_disc)).strip("_")
    return s[:28] or "disciplina_geral"


DOCENTES_SEED = {
    "pedro": {
        "chave": "pedro",
        "nome": "Prof. Pedro",
        "nivel_acesso": "completo",
        "identificadores_whatsapp": [],
        "disciplinas": ["IA", "Legislação de Informática"],
        "funcoes_ativas": [
            "grade_horarios",
            "laboratorio_admin",
            "rag_consulta",
            "planejamento_aulas",
            "diario_classe",
            "frequencia_faltas",
            "vistos_caderno",
            "agente_cli",
        ],
        "ementas_por_disciplina": {
            "IA": (
                "# Guia e Ementa — Inteligência Artificial (IA)\n"
                "**Docente:** Prof. Pedro | **Escopo Estrito:** Inteligência Artificial\n\n"
                "## 1. Fundamentos de IA e Aprendizado de Máquina\n"
                "- Conceitos de Inteligência Artificial simbólica vs conexionista.\n"
                "- Machine Learning supervisionado (classificação e regressão) e não supervisionado (clusterização).\n"
                "- Redes Neurais Artificiais, Perceptron e funções de ativação.\n\n"
                "## 2. IA Generativa, LLMs e Arquitetura RAG\n"
                "- Grandes Modelos de Linguagem (LLMs), tokenização, janela de contexto e Engenharia de Prompt.\n"
                "- Arquitetura RAG (Retrieval-Augmented Generation): indexação documental, busca semântica/lexical e prevenção de alucinações.\n"
                "- Protocolo MCP (Model Context Protocol) e agentes autônomos com uso de ferramentas.\n\n"
                "## 3. Ética, Segurança e Prática em Python\n"
                "- Viés algorítmico, alucinação em modelos generativos, uso responsável e verificação de fontes.\n"
            ),
            "Legislação de Informática": (
                "# Guia e Ementa — Legislação de Informática e Direito Digital\n"
                "**Docente:** Prof. Pedro | **Escopo Estrito:** Legislação de Informática\n\n"
                "## 1. Marco Civil da Internet (Lei nº 12.965/2014)\n"
                "- Princípios fundamentais: neutralidade da rede, liberdade de expressão e privacidade.\n"
                "- Guarda de registros de conexão e de acesso a aplicações de internet.\n\n"
                "## 2. LGPD — Lei Geral de Proteção de Dados (Lei nº 13.709/2018)\n"
                "- Fundamentos, dados pessoais, dados pessoais sensíveis e dados anonimizados.\n"
                "- Bases legais para tratamento de dados, direitos do titular e papel da ANPD.\n"
                "- Segurança da informação, resposta a incidentes e sanções administrativas.\n\n"
                "## 3. Propriedade Intelectual e Crimes Cibernéticos\n"
                "- Lei do Software (Lei nº 9.609/1998), licenças de software livre/aberto (GPL, MIT, Apache) vs proprietário.\n"
                "- Tipificação de crimes cibernéticos no Código Penal (invasão de dispositivo informático — Lei nº 12.737/2012 e Lei nº 14.155/2021).\n"
            ),
        },
    },
    "leandro": {
        "chave": "leandro",
        "nome": "Prof. Leandro",
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [],
        "disciplinas": ["FÍSICA", "AULÃO DE FÍSICA", "PRÁTICA"],
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "ementas_por_disciplina": {
            "FÍSICA": (
                "# Guia e Ementa — Física\n"
                "**Docente:** Prof. Leandro | **Escopo Estrito:** Física\n\n"
                "## 1. Mecânica: Cinemática e Dinâmica\n"
                "- Movimento uniforme (MU), movimento uniformemente variado (MUV) e queda livre.\n"
                "- As três Leis de Newton, força de atrito, trabalho, potência e conservação de energia mecânica.\n\n"
                "## 2. Eletricidade e Termologia\n"
                "- Corrente elétrica, primeira e segunda Leis de Ohm, associação de resistores e potência elétrica.\n"
                "- Escalas termométricas, calorimetria e leis da termodinâmica.\n"
            ),
            "AULÃO DE FÍSICA": (
                "# Guia e Ementa — Aulão de Física (Preparatório ENEM/Vestibulares)\n"
                "**Docente:** Prof. Leandro | **Escopo Estrito:** Aulão de Física\n\n"
                "## Resolução de Questões e Eixos Cognitivos ENEM\n"
                "- Interpretação de gráficos cinemáticos, consumo de energia elétrica residencial (kWh) e fenômenos ondulatórios.\n"
            ),
            "PRÁTICA": (
                "# Guia e Ementa — Prática Experimental de Física\n"
                "**Docente:** Prof. Leandro | **Escopo Estrito:** Prática Experimental\n\n"
                "## Roteiros de Bancada e Medições\n"
                "- Uso de multímetro, montagem de circuitos em série/paralelo e aferição experimental das Leis de Newton.\n"
            ),
        },
    },
    "bruna": {
        "chave": "bruna",
        "nome": "Profa. Bruna",
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [],
        "disciplinas": [
            "Estratégia de Marketing",
            "INTRODUÇÃO AO MARKETING",
            "Legislação Mercadológica",
            "MARKETING DIGITAL",
        ],
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "ementas_por_disciplina": {
            "INTRODUÇÃO AO MARKETING": (
                "# Guia e Ementa — Introdução ao Marketing\n"
                "**Docente:** Profa. Bruna | **Escopo Estrito:** Introdução ao Marketing\n\n"
                "## Fundamentos e Composto Mercadológico\n"
                "- Conceitos centrais: necessidade, desejo, demanda, valor e satisfação.\n"
                "- Os 4 Ps do Marketing (Produto, Preço, Praça e Promoção) e evolução para os 8 Ps.\n"
            ),
            "Estratégia de Marketing": (
                "# Guia e Ementa — Estratégia de Marketing\n"
                "**Docente:** Profa. Bruna | **Escopo Estrito:** Estratégia de Marketing\n\n"
                "## Planejamento Estratégico e Posicionamento\n"
                "- Matriz SWOT (FOFA), Matriz BCG, 5 Forças de Porter, segmentação de mercado, público-alvo e construção de persona.\n"
            ),
            "MARKETING DIGITAL": (
                "# Guia e Ementa — Marketing Digital\n"
                "**Docente:** Profa. Bruna | **Escopo Estrito:** Marketing Digital\n\n"
                "## Inbound Marketing, Tráfego e Métricas Digitais\n"
                "- Funil de vendas digital (Topo, Meio e Fundo de Funil), Inbound Marketing, SEO e Marketing de Conteúdo.\n"
                "- Indicadores de performance: CAC (Custo de Aquisição de Cliente), LTV (Lifetime Value), ROI e taxa de conversão.\n"
            ),
            "Legislação Mercadológica": (
                "# Guia e Ementa — Legislação Mercadológica\n"
                "**Docente:** Profa. Bruna | **Escopo Estrito:** Legislação Mercadológica\n\n"
                "## Direito do Consumidor e Ética Publicitária\n"
                "- Código de Defesa do Consumidor (CDC — Lei nº 8.078/1990): direitos básicos, publicidade enganosa vs abusiva.\n"
                "- Autorregulamentação publicitária (CONAR), direito de arrependimento no e-commerce (Art. 49 do CDC) e LGPD no marketing.\n"
            ),
        },
    },
    "carlos": {
        "chave": "carlos",
        "nome": "Prof. Carlos",
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [],
        "disciplinas": ["Proj. de Aplicativos Mobile", "Proj. de Sist e Orient.  a Objeto com UML"],
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "ementas_por_disciplina": {
            "Proj. de Sist e Orient.  a Objeto com UML": (
                "# Guia e Ementa — Projeto de Sistemas e Orientação a Objetos com UML\n"
                "**Docente:** Prof. Carlos | **Escopo Estrito:** POO e UML\n\n"
                "## 1. Programação Orientada a Objetos (POO)\n"
                "- Classes, objetos, atributos, métodos, abstração, encapsulamento, herança e polimorfismo.\n\n"
                "## 2. Modelagem com UML\n"
                "- Levantamento de requisitos funcionais e não funcionais.\n"
                "- Diagrama de Casos de Uso, Diagrama de Classes, Diagrama de Sequência e Diagrama de Atividades.\n"
            ),
            "Proj. de Aplicativos Mobile": (
                "# Guia e Ementa — Projeto de Aplicativos Mobile\n"
                "**Docente:** Prof. Carlos | **Escopo Estrito:** Desenvolvimento Mobile\n\n"
                "## Arquitetura Mobile, UI/UX e Consumo de APIs\n"
                "- Componentização de telas, navegação entre telas, estado reativo, armazenamento local e consumo de APIs REST em JSON.\n"
            ),
        },
    },
    "marcos": {
        "chave": "marcos",
        "nome": "Prof. Marcos",
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [],
        "disciplinas": ["Int. a Prog Web", "Progamação para Web"],
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "ementas_por_disciplina": {
            "Int. a Prog Web": (
                "# Guia e Ementa — Introdução à Programação Web\n"
                "**Docente:** Prof. Marcos | **Escopo Estrito:** Front-end Essencial (HTML5/CSS3/JS)\n\n"
                "## HTML5 Semântico e Estilização com CSS3\n"
                "- Tags semânticas HTML5 (`header`, `nav`, `main`, `section`, `article`, `footer`), formulários e acessibilidade.\n"
                "- Box Model, seletores CSS, layouts responsivos com Flexbox e CSS Grid, e introdução ao DOM em JavaScript.\n"
            ),
            "Progamação para Web": (
                "# Guia e Ementa — Programação para Web\n"
                "**Docente:** Prof. Marcos | **Escopo Estrito:** Desenvolvimento Web Dinâmico & Backend\n\n"
                "## Aplicações Web Dinâmicas, Eventos, Fetch API e Persistência\n"
                "- Manipulação avançada do DOM, requisições assíncronas (`async/await`, `fetch`), rotas HTTP, validação de formulários e integração com banco de dados.\n"
            ),
        },
    },
    "assis": {
        "chave": "assis",
        "nome": "Prof. Assis",
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [],
        "disciplinas": ["INT. a Algorítmos"],
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "ementas_por_disciplina": {
            "INT. a Algorítmos": (
                "# Guia e Ementa — Introdução a Algoritmos\n"
                "**Docente:** Prof. Assis | **Escopo Estrito:** Lógica de Programação e Algoritmos\n\n"
                "## Construção de Algoritmos e Estruturas de Controle\n"
                "- Variáveis, constantes, tipos primitivos, entrada/saída e operadores lógicos/relacionais.\n"
                "- Estruturas de decisão (`se/senão`, `escolha`) e repetição (`enquanto`, `para`), vetores e modularização.\n"
            ),
        },
    },
    "patricia": {
        "chave": "patricia",
        "nome": "Profa. Patrícia",
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [],
        "disciplinas": ["Economia de Mercado", "Gestão Fin. e Econômica", "Gestão de Pessoas", "TGA"],
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "ementas_por_disciplina": {
            "Economia de Mercado": (
                "# Guia e Ementa — Economia de Mercado\n"
                "**Docente:** Profa. Patrícia | **Escopo Estrito:** Economia de Mercado\n\n"
                "## Microeconomia e Macroeconomia Básica\n"
                "- Oferta, demanda, ponto de equilíbrio, estruturas de mercado, inflação, taxa Selic e PIB.\n"
            ),
            "Gestão Fin. e Econômica": (
                "# Guia e Ementa — Gestão Financeira e Econômica\n"
                "**Docente:** Profa. Patrícia | **Escopo Estrito:** Gestão Financeira\n\n"
                "## Fluxo de Caixa, Custos e Viabilidade\n"
                "- Custos fixos e variáveis, margem de contribuição, ponto de equilíbrio contábil, fluxo de caixa e capital de giro.\n"
            ),
            "Gestão de Pessoas": (
                "# Guia e Ementa — Gestão de Pessoas\n"
                "**Docente:** Profa. Patrícia | **Escopo Estrito:** Gestão de Pessoas\n\n"
                "## Recrutamento, Treinamento e Clima Organizacional\n"
                "- Subprocessos de RH: agregar, aplicar, recompensar, desenvolver, manter e monitorar pessoas; liderança e motivação.\n"
            ),
            "TGA": (
                "# Guia e Ementa — Teoria Geral da Administração (TGA)\n"
                "**Docente:** Profa. Patrícia | **Escopo Estrito:** TGA\n\n"
                "## Escolas Administrativas e Funções PODC\n"
                "- Administração Científica (Taylor), Teoria Clássica (Fayol), Relações Humanas, Burocracia e ciclo PODC (Planejar, Organizar, Dirigir, Controlar).\n"
            ),
        },
    },
    "wanderson": {
        "chave": "wanderson",
        "nome": "Prof. Wanderson",
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [],
        "disciplinas": ["IA"],
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "ementas_por_disciplina": {
            "IA": (
                "# Guia e Ementa — Inteligência Artificial\n"
                "**Docente:** Prof. Wanderson | **Escopo Estrito:** Inteligência Artificial\n\n"
                "## Fundamentos de IA e Automação Inteligente\n"
                "- Conceitos de IA, resolução de problemas, aprendizado supervisionado e ferramentas de IA generativa na prática.\n"
            ),
        },
    },
    "osenias": {
        "chave": "osenias",
        "nome": "Prof. Osenias",
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [],
        "disciplinas": ["Base Técnica"],
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "ementas_por_disciplina": {
            "Base Técnica": (
                "# Guia e Ementa — Componentes da Base Técnica\n"
                "**Docente:** Prof. Osenias | **Escopo Estrito:** Base Técnica\n\n"
                "## Formação Técnica e Prática Profissional\n"
                "- Projetos integradores, normas técnicas e desenvolvimento de competências práticas.\n"
            ),
        },
    },
}


def _limpar_id(raw: str) -> str:
    return re.sub(r"\D", "", str(raw or "").split("@")[0].split(":")[0])


def _carregar_grade_sqlite_por_prof() -> dict:
    grades = {}
    if not os.path.exists(ESCOLA_DB):
        return grades
    try:
        with sqlite3.connect(ESCOLA_DB, timeout=15) as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='grade_base_tecnica'"
            )
            if not cur.fetchone():
                return grades
            rows = cur.execute(
                """
                SELECT prof_chave, dia_semana, tempo_aula, turma, disciplina
                FROM grade_base_tecnica
                ORDER BY prof_chave, dia_semana, tempo_aula
                """
            ).fetchall()
            for pk, dia, tempo, turma, disc in rows:
                ini, fim = HORARIOS_AULA_PADRAO.get(int(tempo), ("", ""))
                grades.setdefault(pk, []).append(
                    {
                        "dia_semana": int(dia),
                        "dia_nome": DIAS_SEMANA_NOMES.get(int(dia), f"Dia {dia}"),
                        "tempo_aula": int(tempo),
                        "horario_inicio": ini,
                        "horario_fim": fim,
                        "turma": turma,
                        "disciplina": disc,
                    }
                )
    except Exception:
        pass
    return grades


def inicializar_estrutura_docentes(forcar_atualizacao_grade: bool = False) -> dict:
    """
    Cria e sincroniza a estrutura multi-docente com isolamento estrito por disciplina:
    - `docentes/<chave>/rag_conteudos/<materia_slug>/ementa_<materia_slug>.md`
    - Remove arquivos misturados antigos na raiz de `rag_conteudos` se já existirem pastas isoladas.
    """
    os.makedirs(DOCENTES_DIR, exist_ok=True)
    grades_db = _carregar_grade_sqlite_por_prof()

    catalogo = {
        "versao": "2.0-strict-rag",
        "atualizado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "docentes": {},
        "indice_whatsapp": {},
    }

    for chave, info in DOCENTES_SEED.items():
        pasta_prof = os.path.join(DOCENTES_DIR, chave)
        pasta_rag_raiz = os.path.join(pasta_prof, "rag_conteudos")
        pasta_planos = os.path.join(pasta_prof, "planos_aula")
        os.makedirs(pasta_rag_raiz, exist_ok=True)
        os.makedirs(pasta_planos, exist_ok=True)

        # Remove arquivo legado `ementa_base.md` solto na raiz que misturava disciplinas
        legado_raiz = os.path.join(pasta_rag_raiz, "ementa_base.md")
        if os.path.exists(legado_raiz) and info.get("ementas_por_disciplina"):
            try:
                os.remove(legado_raiz)
            except Exception:
                pass

        aulas_prof = grades_db.get(chave, [])
        turmas_prof = sorted({a["turma"] for a in aulas_prof if a.get("turma")})
        ementas_map = info.get("ementas_por_disciplina", {})
        disciplinas_finais = list(ementas_map.keys()) if ementas_map else info.get("disciplinas", ["Base Técnica"])

        mapa_pastas_disciplinas = {}
        for d_nome in disciplinas_finais:
            d_slug = slug_disciplina(d_nome)
            pasta_disc_rag = os.path.join(pasta_rag_raiz, d_slug)
            os.makedirs(pasta_disc_rag, exist_ok=True)
            arq_ementa = os.path.join(pasta_disc_rag, f"guia_ementa_{d_slug}.md")
            texto_ementa = ementas_map.get(d_nome) or (
                f"# Guia e Ementa — {d_nome}\n"
                f"**Docente:** {info['nome']} | **Escopo Estrito:** {d_nome}\n\n"
                f"## Conteúdos Programáticos de {d_nome}\n"
                f"- Fundamentos, prática dirigida e avaliação formativa de {d_nome}.\n"
            )
            if not os.path.exists(arq_ementa):
                with open(arq_ementa, "w", encoding="utf-8") as f_em:
                    f_em.write(texto_ementa)
            mapa_pastas_disciplinas[d_slug] = {
                "id": d_slug,
                "nome": d_nome,
                "pasta_rag": pasta_disc_rag,
            }

        perfil_path = os.path.join(pasta_prof, "perfil.json")
        perfil_data = {
            "chave": chave,
            "nome": info["nome"],
            "nivel_acesso": info["nivel_acesso"],
            "identificadores_whatsapp": info["identificadores_whatsapp"],
            "disciplinas": disciplinas_finais,
            "disciplinas_isoladas": mapa_pastas_disciplinas,
            "turmas": turmas_prof,
            "funcoes_ativas": info["funcoes_ativas"],
            "pasta_rag": pasta_rag_raiz,
            "pasta_planos_aula": pasta_planos,
        }
        with open(perfil_path, "w", encoding="utf-8") as f:
            json.dump(perfil_data, f, ensure_ascii=False, indent=2)

        grade_path = os.path.join(pasta_prof, "grade_horarios.json")
        if forcar_atualizacao_grade or not os.path.exists(grade_path) or aulas_prof:
            grade_data = {
                "prof_chave": chave,
                "prof_nome": info["nome"],
                "total_aulas_semanais": len(aulas_prof),
                "turmas": turmas_prof,
                "disciplinas": disciplinas_finais,
                "aulas": aulas_prof,
            }
            with open(grade_path, "w", encoding="utf-8") as f:
                json.dump(grade_data, f, ensure_ascii=False, indent=2)

        catalogo["docentes"][chave] = {
            "nome": info["nome"],
            "nivel_acesso": info["nivel_acesso"],
            "disciplinas": disciplinas_finais,
            "disciplinas_isoladas": list(mapa_pastas_disciplinas.keys()),
            "turmas": turmas_prof,
            "total_aulas": len(aulas_prof),
            "pasta": pasta_prof,
        }
        for wid in info["identificadores_whatsapp"]:
            catalogo["indice_whatsapp"][_limpar_id(wid)] = chave

    if os.path.exists(ESCOLA_DB):
        try:
            with sqlite3.connect(ESCOLA_DB, timeout=10) as conn:
                rows = conn.execute(
                    "SELECT numero, chave, nome FROM professores_base_vinculos"
                ).fetchall()
                for num, ch, _ in rows:
                    num_limpo = _limpar_id(num)
                    if num_limpo and ch in catalogo["docentes"]:
                        catalogo["indice_whatsapp"][num_limpo] = ch
        except Exception:
            pass

    with open(CATALOGO_PATH, "w", encoding="utf-8") as f:
        json.dump(catalogo, f, ensure_ascii=False, indent=2)

    return catalogo


def identificar_docente_por_whatsapp(*identificadores) -> dict | None:
    if not os.path.exists(CATALOGO_PATH):
        inicializar_estrutura_docentes()

    try:
        with open(CATALOGO_PATH, "r", encoding="utf-8") as f:
            catalogo = json.load(f)
    except Exception:
        catalogo = inicializar_estrutura_docentes()

    indice = catalogo.get("indice_whatsapp", {})
    chave_encontrada = None

    for raw in identificadores:
        limpo = _limpar_id(raw)
        if not limpo:
            continue
        if limpo in indice:
            chave_encontrada = indice[limpo]
            break
        if len(limpo) >= 8:
            sufixo = limpo[-8:]
            for k_id, ch in indice.items():
                if len(k_id) >= 8 and k_id.endswith(sufixo):
                    chave_encontrada = ch
                    break
        if chave_encontrada:
            break

    if not chave_encontrada:
        return None

    perfil_path = os.path.join(DOCENTES_DIR, chave_encontrada, "perfil.json")
    if os.path.exists(perfil_path):
        try:
            with open(perfil_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return None


def consultar_grade_individual_docente(prof_chave: str, filtro_texto: str = "") -> str:
    if not os.path.exists(DOCENTES_DIR):
        inicializar_estrutura_docentes()

    grade_path = os.path.join(DOCENTES_DIR, prof_chave, "grade_horarios.json")
    if not os.path.exists(grade_path):
        inicializar_estrutura_docentes(forcar_atualizacao_grade=True)

    if not os.path.exists(grade_path):
        return "Grade individual não encontrada para este perfil docente."

    with open(grade_path, "r", encoding="utf-8") as f:
        grade = json.load(f)

    nome = grade.get("prof_nome", f"Prof. {prof_chave.title()}")
    aulas = grade.get("aulas", [])
    if not aulas:
        return (
            f"📅 *GRADE DE HORÁRIOS — {nome}*\n\n"
            "Nenhuma aula cadastrada na planilha base no momento."
        )

    t_norm = _normalizar(filtro_texto)
    agora = datetime.now()
    dia_filtro = None
    titulo_dia = "SEMANA COMPLETA"

    mapa_dias_filtro = {
        "segunda": 0,
        "seg": 0,
        "terca": 1,
        "ter": 1,
        "quarta": 2,
        "qua": 2,
        "quinta": 3,
        "qui": 3,
        "sexta": 4,
        "sex": 4,
    }

    if "hoje" in t_norm:
        dia_filtro = agora.weekday()
        titulo_dia = f"HOJE ({DIAS_SEMANA_NOMES.get(dia_filtro, '')})"
    elif "amanha" in t_norm:
        amanha = agora + timedelta(days=1)
        dia_filtro = amanha.weekday()
        titulo_dia = f"AMANHÃ ({DIAS_SEMANA_NOMES.get(dia_filtro, '')})"
    else:
        for nome_d, idx_d in mapa_dias_filtro.items():
            if re.search(rf"\b{nome_d}\b", t_norm):
                dia_filtro = idx_d
                titulo_dia = DIAS_SEMANA_NOMES.get(idx_d, "").upper()
                break

    if dia_filtro is not None:
        if dia_filtro >= 5:
            return f"📅 *Grade — {nome} ({titulo_dia}):* Sem aulas regulares no fim de semana."
        aulas_filtradas = [a for a in aulas if a["dia_semana"] == dia_filtro]
        if not aulas_filtradas:
            return f"📅 *Grade — {nome} ({titulo_dia}):* Nenhuma aula registrada."
        linhas = [f"📅 *Grade — {nome} ({titulo_dia})*"]
        for a in aulas_filtradas:
            linhas.append(f"• *{a['tempo_aula']}º Tempo:* Turma {a['turma']} — {a['disciplina']}")
        return "\n".join(linhas)

    linhas = [f"📅 *Grade Semanal — {nome}*"]
    for dia_idx in range(5):
        aulas_dia = [a for a in aulas if a["dia_semana"] == dia_idx]
        if not aulas_dia:
            continue
        linhas.append(f"\n*{DIAS_SEMANA_NOMES[dia_idx]}:*")
        for a in aulas_dia:
            linhas.append(f"• {a['tempo_aula']}º Tempo: Turma {a['turma']} — {a['disciplina']}")

    return "\n".join(linhas).strip()


def buscar_rag_docente(
    prof_chave: str,
    consulta: str,
    top_k: int = 3,
    materia_id: str | None = None,
    exigir_match_consulta: bool = False,
) -> list[dict]:
    """
    Motor RAG com Isolamento Estrito por Disciplina (Anti-Alucinação):
    - Quando `materia_id` é informado, vasculha **exclusivamente** a subpasta
      `docentes/<prof_chave>/rag_conteudos/<materia_id>/` (nunca mistura conteúdos de outra matéria).
    - Quando `exigir_match_consulta=True` e a consulta tem termos, só retorna trechos que
      efetivamente possuam match (`score > 0`) nos documentos anexados daquela matéria,
      evitando inventar ou extrapolar conteúdos fora do guia/ementa.
    """
    if not os.path.exists(DOCENTES_DIR):
        inicializar_estrutura_docentes()

    pasta_rag_base = os.path.join(DOCENTES_DIR, prof_chave, "rag_conteudos")
    if not os.path.exists(pasta_rag_base):
        return []

    if materia_id:
        m_slug = slug_disciplina(materia_id)
        pasta_alvo = os.path.join(pasta_rag_base, m_slug)
        if not os.path.exists(pasta_alvo):
            # Tenta match parcial caso o id legado seja ligeiramente diferente (ex: inteligencia_artificial -> ia)
            subpastas = [
                d for d in os.listdir(pasta_rag_base)
                if os.path.isdir(os.path.join(pasta_rag_base, d))
            ]
            encontrada = None
            for sp in subpastas:
                if sp in m_slug or m_slug in sp or (sp == "ia" and "inteligencia" in m_slug):
                    encontrada = os.path.join(pasta_rag_base, sp)
                    break
            pasta_alvo = encontrada or pasta_alvo
        pastas_escopo = [pasta_alvo] if os.path.exists(pasta_alvo) else []
    else:
        pastas_escopo = [pasta_rag_base]

    termos = [
        t
        for t in re.findall(r"[a-z0-9]{2,}", _normalizar(consulta))
        if t not in {"de", "da", "do", "das", "dos", "para", "com", "em", "na", "no", "sobre", "aula", "rag"}
    ]

    blocos = []
    for p_esc in pastas_escopo:
        for root, _, files in os.walk(p_esc):
            rel_disc = os.path.basename(root) if root != pasta_rag_base else "geral"
            for fname in sorted(files):
                if not fname.endswith((".md", ".txt", ".csv")):
                    continue
                fpath = os.path.join(root, fname)
                try:
                    with open(fpath, "r", encoding="utf-8", errors="replace") as f:
                        conteudo = f.read()
                except Exception:
                    continue

                raw_secoes = re.split(r"\n(?=##+\s)|\n{3,}", conteudo)
                secoes = []
                buffer = ""
                for s in raw_secoes:
                    s_limpa = s.strip()
                    if not s_limpa:
                        continue
                    if len(s_limpa) < 160 and not buffer:
                        buffer = s_limpa
                        continue
                    if buffer:
                        s_limpa = buffer + "\n\n" + s_limpa
                        buffer = ""
                    secoes.append(s_limpa)
                if buffer:
                    secoes.append(buffer)

                for sec in secoes:
                    sec_limpa = sec.strip()
                    if not sec_limpa:
                        continue
                    sec_norm = _normalizar(sec_limpa)
                    score = 0
                    termos_encontrados = []
                    for termo in termos:
                        if termo in sec_norm:
                             pts = 3 if re.search(rf"\b{re.escape(termo)}\b", sec_norm) else 1
                             score += pts
                             termos_encontrados.append(termo)
                    cons_limpa = _normalizar(consulta)
                    if len(termos) > 1 and cons_limpa in sec_norm:
                        score += 10
                    if len(termos) > 1 and all(t in rel_disc for t in termos):
                        score += 5
                    blocos.append(
                        {
                            "disciplina_slug": rel_disc,
                            "arquivo": fname,
                            "caminho": fpath,
                            "trecho": sec_limpa,
                            "score": score,
                            "termos_matched": termos_encontrados,
                        }
                    )

    if not blocos:
        return []

    blocos_ordenados = sorted(blocos, key=lambda x: x["score"], reverse=True)
    if termos:
        com_match = [b for b in blocos_ordenados if b["score"] > 0]
        if com_match:
            return com_match[:top_k]
        if exigir_match_consulta:
            return []
    return blocos_ordenados[:top_k]


def formatar_resposta_rag_docente(prof_chave: str, consulta: str = "", materia_id: str | None = None) -> str:
    perfil_path = os.path.join(DOCENTES_DIR, prof_chave, "perfil.json")
    if not os.path.exists(perfil_path):
        inicializar_estrutura_docentes()
    try:
        with open(perfil_path, "r", encoding="utf-8") as f:
            perfil = json.load(f)
    except Exception:
        return "Perfil docente não encontrado."

    resultados = buscar_rag_docente(prof_chave, consulta, top_k=3, materia_id=materia_id)
    nome = perfil.get("nome", prof_chave)
    disciplinas = ", ".join(perfil.get("disciplinas", []))

    linhas = [
        f"🧠 *BASE DE CONHECIMENTO (RAG ESTRITO) — {nome}*",
        f"📘 *Disciplinas vinculadas:* {disciplinas}\n",
    ]
    if consulta.strip():
        linhas.append(f"🔎 *Filtro consultado:* `{consulta.strip()}`\n")

    if not resultados:
        linhas.append("⚠️ Nenhum trecho encontrado com lastro nos guias/ementas desta disciplina.")
    else:
        for idx, item in enumerate(resultados, 1):
            linhas.append(f"📄 *[{idx}] Disciplina: `{item['disciplina_slug']}` • Fonte: `{item['arquivo']}`*")
            linhas.append(item["trecho"])
            linhas.append("")

    if perfil.get("nivel_acesso") == "essencial":
        linhas.append(
            "ℹ️ _Seu perfil está no modo Assistente Essencial (Grade de Horários + Laboratório + Base RAG Isolada por Disciplina)._"
        )
    return "\n".join(linhas).strip()


def extrair_texto_documento(caminho: str) -> str:
    """
    Extrator universal de documentos didáticos (PDF, Word DOCX, PowerPoint PPTX, Markdown, Texto e CSV).
    Preserva a estrutura textual, tabelas e marcadores de tópicos essenciais para o RAG.
    """
    if not os.path.exists(caminho):
        return ""
    ext = os.path.splitext(caminho)[1].lower()
    
    # 1. Arquivos de texto puro, markdown e dados tabulares simples
    if ext in [".txt", ".md", ".json", ".csv"]:
        try:
            with open(caminho, "r", encoding="utf-8", errors="replace") as f:
                return f.read()
        except Exception:
            return ""

    # 2. Documentos Microsoft Word (.docx)
    if ext == ".docx":
        try:
            import docx
            doc = docx.Document(caminho)
            linhas = []
            for p in doc.paragraphs:
                t = p.text.strip()
                if t:
                    linhas.append(t)
            for t in doc.tables:
                for row in t.rows:
                    celulas = [c.text.strip().replace("\n", " ") for c in row.cells if c.text.strip()]
                    u = []
                    for c in celulas:
                        if not u or c != u[-1]:
                            u.append(c)
                    if u:
                        linhas.append(" | ".join(u))
            return "\n".join(linhas)
        except Exception as e:
            return f"Erro ao extrair Word (.docx): {e}"

    # 3. Apresentações em Slides Microsoft PowerPoint (.pptx)
    if ext == ".pptx":
        try:
            import pptx
            prs = pptx.Presentation(caminho)
            slides_txt = []
            for idx, slide in enumerate(prs.slides, 1):
                textos = []
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            t = paragraph.text.strip()
                            if t and t not in textos:
                                textos.append(t)
                if textos:
                    slides_txt.append(f"### Slide {idx}\n" + "\n".join(f"- {t}" for t in textos))
            return "\n\n".join(slides_txt)
        except Exception as e:
            return f"Erro ao extrair Slides (.pptx): {e}"

    # 4. Arquivos PDF (.pdf)
    if ext == ".pdf":
        try:
            if os.path.exists("/usr/bin/pdftotext"):
                out = subprocess.check_output(["/usr/bin/pdftotext", caminho, "-"], stderr=subprocess.DEVNULL)
                txt = out.decode("utf-8", errors="replace").strip()
                if txt:
                    return txt
        except Exception:
            pass
        try:
            from pypdf import PdfReader
            reader = PdfReader(caminho)
            paginas = [p.extract_text() or "" for p in reader.pages]
            return "\n\n".join([p for p in paginas if p.strip()])
        except Exception as e:
            return f"Erro ao extrair PDF: {e}"

    return ""


def adicionar_documento_ao_rag(
    prof_chave: str,
    caminho_arquivo: str,
    nome_original: str = "",
    materia_id: str | None = None,
) -> dict:
    """
    Ingesta um documento (.docx, .pptx, .pdf, .txt, .md) diretamente na base RAG do docente,
    garantindo o isolamento estrito por disciplina e sincronização entre Bot e Portal Web.
    """
    if not os.path.exists(DOCENTES_DIR):
        inicializar_estrutura_docentes()

    perfil_path = os.path.join(DOCENTES_DIR, prof_chave, "perfil.json")
    if not os.path.exists(perfil_path):
        return {"sucesso": False, "erro": f"Docente '{prof_chave}' não encontrado no sistema."}

    with open(perfil_path, "r", encoding="utf-8") as f:
        perfil = json.load(f)

    nome_arquivo = nome_original or os.path.basename(caminho_arquivo)
    nome_sanitizado = re.sub(r"[^a-zA-Z0-9._-]", "_", os.path.basename(nome_arquivo))
    ext = os.path.splitext(nome_arquivo)[1].lower()

    # Identifica ou normaliza a disciplina alvo
    disciplinas_docente = perfil.get("disciplinas", [])
    if not materia_id:
        # Tenta adivinhar pelo nome do arquivo ou pega a primeira disciplina do professor
        fn_norm = _normalizar(nome_arquivo)
        encontrada = None
        for d in disciplinas_docente:
            dn = _normalizar(d)
            if any(termo in fn_norm for termo in dn.split() if len(termo) > 3):
                encontrada = d
                break
        materia_nome = encontrada or (disciplinas_docente[0] if disciplinas_docente else "Geral")
    else:
        materia_nome = materia_id

    disc_slug = slug_disciplina(materia_nome)
    pasta_rag_disc = os.path.join(DOCENTES_DIR, prof_chave, "rag_conteudos", disc_slug)
    os.makedirs(pasta_rag_disc, exist_ok=True)

    # 🛡️ Armazenamento Isolado para VPS / Produção
    pasta_uploads = os.path.join(PASTA_UPLOADS_ISOLADOS, prof_chave, disc_slug)
    os.makedirs(pasta_uploads, exist_ok=True)
    caminho_bruto_isolado = os.path.join(pasta_uploads, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{nome_sanitizado}")
    try:
        shutil.copy2(caminho_arquivo, caminho_bruto_isolado)
    except Exception:
        pass

    # Copia o arquivo original para o diretório da disciplina
    caminho_copia = os.path.join(pasta_rag_disc, nome_arquivo)
    if os.path.abspath(caminho_arquivo) != os.path.abspath(caminho_copia):
        shutil.copy2(caminho_arquivo, caminho_copia)

    # Extrai o texto formatado para indexação em Markdown
    texto = extrair_texto_documento(caminho_copia)
    if not texto.strip():
        return {
            "sucesso": False,
            "erro": f"Não foi possível extrair texto do documento '{nome_arquivo}' (formato não suportado ou vazio).",
        }

    nome_md = os.path.splitext(nome_arquivo)[0] + "_extraido.md"
    caminho_md = os.path.join(pasta_rag_disc, nome_md)
    cabecalho = (
        f"# {nome_arquivo} — Documento RAG\n"
        f"**Docente:** {perfil.get('nome', prof_chave)} | **Disciplina:** {materia_nome}\n"
        f"**Data de Indexação:** {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
    )
    with open(caminho_md, "w", encoding="utf-8") as f:
        f.write(cabecalho + texto + "\n")

    # Sincroniza também com o Portal Web
    pasta_web = os.path.join(BASE_ECOSSISTEMA_DIR, "web", "dados", prof_chave, "materias", disc_slug)
    os.makedirs(pasta_web, exist_ok=True)
    try:
        shutil.copy2(caminho_copia, os.path.join(pasta_web, nome_arquivo))
        shutil.copy2(caminho_md, os.path.join(pasta_web, nome_md))
    except Exception:
        pass

    # Garante que a disciplina conste no perfil se for uma nova
    if materia_nome not in disciplinas_docente:
        disciplinas_docente.append(materia_nome)
        perfil["disciplinas"] = disciplinas_docente
        if "disciplinas_isoladas" not in perfil or isinstance(perfil["disciplinas_isoladas"], list):
            perfil["disciplinas_isoladas"] = {}
        perfil["disciplinas_isoladas"][disc_slug] = {
            "id": disc_slug,
            "nome": materia_nome,
            "pasta_rag": pasta_rag_disc,
        }
        with open(perfil_path, "w", encoding="utf-8") as f:
            json.dump(perfil, f, ensure_ascii=False, indent=2)

    return {
        "sucesso": True,
        "docente": perfil.get("nome", prof_chave),
        "chave": prof_chave,
        "disciplina": materia_nome,
        "disciplina_slug": disc_slug,
        "arquivo": nome_arquivo,
        "arquivo_md": nome_md,
        "caracteres": len(texto),
        "linhas": len(texto.splitlines()),
        "caminho_rag": caminho_md,
    }


def cadastrar_novo_docente(
    nome: str,
    whatsapp: str,
    disciplinas: list[str],
    turmas: list[str] | None = None,
    eixo: str = "Base Técnica",
) -> dict:
    """
    Cadastra ou atualiza um docente no ecossistema com vinculação completa:
    - Aplica verificação estrita da Whitelist da Base Técnica (Carlos bloqueado, outros proibidos)
    - Cria diretório isolado RAG em `docentes/<chave>/`
    - Registra número no catálogo e índice de WhatsApp
    - Vincula no banco SQLite `escola.db` (tabela professores_base_vinculos)
    - Cria perfil no Portal Web (`web/dados/<chave>/perfil.yaml`)
    """
    nome_limpo = re.sub(r"\s+", " ", str(nome or "").strip())
    if not nome_limpo:
        return {"sucesso": False, "erro": "Nome do docente não pode ser vazio."}

    # 🛡️ Verificação Estrita da Whitelist da Base Técnica (Carlos bloqueado, outros proibidos)
    permitido, motivo, chave_homologada = validar_docente_permitido(nome_limpo)
    if not permitido:
        return {"sucesso": False, "erro": motivo}

    # Usa a chave homologada da lista oficial
    chave = chave_homologada or slug_disciplina(nome_limpo)

    num_limpo = _limpar_id(whatsapp)
    turmas_lista = turmas or ["101", "102"]
    disc_lista = [d.strip() for d in disciplinas if d.strip()] or ["Componente Curricular Técnico"]

    pasta_prof = os.path.join(DOCENTES_DIR, chave)
    pasta_rag = os.path.join(pasta_prof, "rag_conteudos")
    pasta_planos = os.path.join(pasta_prof, "planos_aula")
    os.makedirs(pasta_rag, exist_ok=True)
    os.makedirs(pasta_planos, exist_ok=True)

    # Cria pastas isoladas para cada disciplina
    mapa_disc = {}
    for d in disc_lista:
        d_slug = slug_disciplina(d)
        p_disc = os.path.join(pasta_rag, d_slug)
        os.makedirs(p_disc, exist_ok=True)
        ementa_arq = os.path.join(p_disc, f"guia_ementa_{d_slug}.md")
        if not os.path.exists(ementa_arq):
            with open(ementa_arq, "w", encoding="utf-8") as f:
                f.write(
                    f"# Guia e Ementa — {d}\n"
                    f"**Docente:** {nome_limpo} | **Instituição:** IEMA Pleno Açailândia\n\n"
                    f"## Objetivos de Aprendizagem e Competências\n"
                    f"- Fundamentos teóricos e atividades práticas de {d}.\n"
                )
        mapa_disc[d_slug] = {"id": d_slug, "nome": d, "pasta_rag": p_disc}

    # Salva perfil.json
    perfil_data = {
        "chave": chave,
        "nome": nome_limpo,
        "nivel_acesso": "essencial",
        "identificadores_whatsapp": [] if num_limpo else [],
        "disciplinas": disc_lista,
        "disciplinas_isoladas": mapa_disc,
        "turmas": turmas_lista,
        "funcoes_ativas": ["grade_horarios", "laboratorio", "rag_consulta"],
        "pasta_rag": pasta_rag,
        "pasta_planos_aula": pasta_planos,
    }
    with open(os.path.join(pasta_prof, "perfil.json"), "w", encoding="utf-8") as f:
        json.dump(perfil_data, f, ensure_ascii=False, indent=2)

    # Atualiza catalogo_docentes.json
    try:
        cat = {}
        if os.path.exists(CATALOGO_PATH):
            with open(CATALOGO_PATH, "r", encoding="utf-8") as f:
                cat = json.load(f)
        cat.setdefault("docentes", {})[chave] = {
            "nome": nome_limpo,
            "nivel_acesso": "essencial",
            "disciplinas": disc_lista,
            "disciplinas_isoladas": list(mapa_disc.keys()),
            "turmas": turmas_lista,
            "total_aulas": 0,
            "pasta": pasta_prof,
        }
        if num_limpo:
            cat.setdefault("indice_whatsapp", {})[num_limpo] = chave
        with open(CATALOGO_PATH, "w", encoding="utf-8") as f:
            json.dump(cat, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Aviso ao atualizar catálogo: {e}")

    # Registra no banco SQLite escola.db
    if os.path.exists(ESCOLA_DB) and num_limpo:
        try:
            with sqlite3.connect(ESCOLA_DB, timeout=10) as conn:
                conn.execute(
                    "INSERT OR REPLACE INTO professores_base_vinculos (numero, chave, nome) VALUES (?, ?, ?);",
                    (num_limpo, chave, nome_limpo),
                )
                conn.commit()
        except Exception as e:
            print(f"Aviso ao registrar vínculo no SQLite: {e}")

    # Cria estrutura web/dados/<chave>/perfil.yaml
    web_dir = os.path.join(BASE_ECOSSISTEMA_DIR, "web", "dados", chave)
    os.makedirs(web_dir, exist_ok=True)
    os.makedirs(os.path.join(web_dir, "aulas_geradas", "slides"), exist_ok=True)
    os.makedirs(os.path.join(web_dir, "aulas_geradas", "questionarios"), exist_ok=True)
    for d in disc_lista:
        d_slug = slug_disciplina(d)
        os.makedirs(os.path.join(web_dir, "materias", d_slug), exist_ok=True)

    try:
        import yaml
        perfil_yaml = {
            "id": chave,
            "nome": nome_limpo,
            "email": f"{chave}@exemplo.edu.br",
            "telefone_whatsapp": num_limpo,
            "nivel_acesso": "essencial",
            "eixo_tecnologico": eixo,
            "instituicao": {"nome": "IEMA Pleno Açailândia", "cidade": "Açailândia - MA"},
            "turmas": [{"id": t, "nome": f"Turma {t}", "disciplina": disc_lista[0]} for t in turmas_lista],
            "disciplinas": [
                {"id": slug_disciplina(d), "nome": d, "turmas": turmas_lista, "ementa_topicos": [f"Fundamentos de {d}"]}
                for d in disc_lista
            ],
            "regras_avaliacao": {"limite_faltas_critico": 15, "media_aprovacao": 6.0},
        }
        with open(os.path.join(web_dir, "perfil.yaml"), "w", encoding="utf-8") as f:
            yaml.safe_dump(perfil_yaml, f, allow_unicode=True)
    except Exception:
        pass

    return {
        "sucesso": True,
        "chave": chave,
        "nome": nome_limpo,
        "disciplinas": disc_lista,
        "turmas": turmas_lista,
        "whatsapp": num_limpo,
    }

