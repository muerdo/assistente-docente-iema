#!/usr/bin/env python3
"""
==============================================================================
GERADOR DE CONTEÚDO DE AULA TÉCNICA A PARTIR DE VÍDEO OU LINK (IDEIA #12)
==============================================================================
Recebe uma URL (YouTube ou artigo técnico na web) e esmiúça o conteúdo para
estruturar uma aula prática completa de 50 minutos (ou bloco duplo) para o IEMA.

Gera:
1. Plano de Aula minucioso em Markdown cronometrado (00-10m, 10-25m, 25-45m, 45-50m).
2. Atividade prática de laboratório com critérios de avaliação e visto.
3. Apresentação de slides em PowerPoint (.pptx).
4. Resumo direto e conciso para envio no WhatsApp.
==============================================================================
"""

import os
import sys
import re
import json
import subprocess
import urllib.request
import urllib.parse
from datetime import datetime
from html.parser import HTMLParser

try:
    from pptx import Presentation
except ImportError:
    Presentation = None

HOME = os.environ.get("DOCENTE_HOME", ".")
PROJETOS_IDEIAS_DIR = f"{HOME}/projetos_ideias"
DIR_IDEIA_12 = f"{PROJETOS_IDEIAS_DIR}/ideia_12_gerar_comteudo_pra_aula_com_ba"
DIR_AULAS_GERADAS = f"{DIR_IDEIA_12}/aulas_geradas"
DIR_AULAS_LISTAS = f"{HOME}/aulas_listas"
DIR_GDRIVE_SLIDES = f"{HOME}/mounts/gdrive/Aulas/Slides"

os.makedirs(DIR_AULAS_GERADAS, exist_ok=True)
os.makedirs(DIR_AULAS_LISTAS, exist_ok=True)


class ExtratorTextoHTML(HTMLParser):
    def __init__(self):
        super().__init__()
        self.ignorar = False
        self.tags_ignorar = {"script", "style", "noscript", "header", "footer", "nav"}
        self.titulo = ""
        self.capturando_titulo = False
        self.texto = []

    def handle_starttag(self, tag, attrs):
        if tag.lower() in self.tags_ignorar:
            self.ignorar = True
        elif tag.lower() == "title":
            self.capturando_titulo = True

    def handle_endtag(self, tag):
        if tag.lower() in self.tags_ignorar:
            self.ignorar = False
        elif tag.lower() == "title":
            self.capturando_titulo = False

    def handle_data(self, data):
        if self.capturando_titulo and not self.titulo:
            self.titulo = data.strip()
        elif not self.ignorar:
            txt = data.strip()
            if txt:
                self.texto.append(txt)

    def obter_texto_limpo(self, max_chars=4000):
        corpo = " ".join(self.texto)
        corpo = re.sub(r"\s+", " ", corpo).strip()
        return self.titulo, corpo[:max_chars]


def sanitizar_nome(texto: str) -> str:
    limpo = re.sub(r"[^\w\s-]", "", texto).strip()
    return re.sub(r"[-\s]+", "_", limpo).capitalize()[:40]


def extrair_metadados_url(url: str) -> dict:
    """Extrai informações cruciais do link (YouTube ou artigo web)."""
    resultado = {
        "url": url,
        "titulo": "",
        "descricao": "",
        "duracao_segundos": 0,
        "fonte": "web"
    }

    eh_youtube = bool(re.search(r"(youtube\.com|youtu\.be)", url, re.I))

    if eh_youtube:
        resultado["fonte"] = "youtube"
        try:
            cmd = ["yt-dlp", "--dump-json", "--skip-download", "--no-warnings", url]
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
            if proc.returncode == 0 and proc.stdout:
                dados = json.loads(proc.stdout)
                resultado["titulo"] = dados.get("title", "Vídeo Técnico")
                resultado["descricao"] = dados.get("description", "")[:1200]
                resultado["duracao_segundos"] = dados.get("duration", 0)
                return resultado
        except Exception:
            pass

    # Fallback ou página web genérica via HTTP
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) DocenteAssist/1.0"}
        )
        with urllib.request.urlopen(req, timeout=12) as resp:
            raw_html = resp.read().decode("utf-8", errors="replace")
            parser = ExtratorTextoHTML()
            parser.feed(raw_html)
            tit, desc = parser.obter_texto_limpo()
            resultado["titulo"] = tit or "Recurso Técnico para Aula"
            resultado["descricao"] = desc
    except Exception as e:
        resultado["titulo"] = "Recurso Técnico Externo"
        resultado["descricao"] = f"Conteúdo acessível no link fornecido: {url} (Erro ao extrair: {e})"

    return resultado


def esmiucar_para_aula(metadados: dict, turma: str = "Técnico em Informática", tempo_minutos: int = 50) -> dict:
    """
    Estrutura a aula de 50 minutos de forma didática, prática e aplicada,
    distribuindo os tempos rigorosamente.
    """
    titulo_base = metadados.get("titulo", "Tópico Técnico")
    titulo_base = re.sub(r"[\(\[].*?[\)\]]", "", titulo_base).strip()
    url = metadados.get("url", "")
    descricao = metadados.get("descricao", "")

    # Identifica pontos técnicos na descrição
    palavras_chave = []
    padroes_tecnicos = re.findall(r"\b([A-Z][a-zA-Z0-9_\+\#]{2,}|python|javascript|sql|linux|docker|api|git)\b", descricao, re.I)
    for p in padroes_tecnicos:
        if p.lower() not in [x.lower() for x in palavras_chave]:
            palavras_chave.append(p)
        if len(palavras_chave) >= 5:
            break

    tags_str = ", ".join(palavras_chave) if palavras_chave else "Conceitos Fundamentais, Prática Guiada"

    roteiro_minutagem = [
        {
            "bloco": "00 a 10 min",
            "etapa": "Acolhida & Desafio Disparador",
            "detalhes": (
                f"Apresentação da problematização do vídeo/link. Exibição do trecho-chave (3 a 5 min) "
                f"ou discussão do problema real: 'Por que precisamos dominar {titulo_base}?'"
            ),
            "acao_docente": "Conectar a aplicação real do mercado de trabalho com a rotina dos alunos.",
            "acao_estudante": "Anotar o problema central no caderno ou terminal."
        },
        {
            "bloco": "10 a 25 min",
            "etapa": "Fundamentação & Desconstrução Técnica",
            "detalhes": (
                f"Esmiuçamento do conteúdo em 3 conceitos essenciais: 1) Arquitetura/Fundamento, "
                f"2) Sintaxe/Configuração prática, 3) Cuidados e boas práticas recomendadas."
            ),
            "acao_docente": "Explicar no quadro ou projetor passo a passo, demonstrando um exemplo mínimo.",
            "acao_estudante": "Acompanhar a demonstração e tirar dúvidas sobre a estrutura técnica."
        },
        {
            "bloco": "25 a 45 min",
            "etapa": "Prática Mão na Massa em Laboratório",
            "detalhes": (
                f"Laboratório dirigido: Cada estudante implementa e testa o conceito demonstrado no vídeo. "
                f"Execução de roteiro prático curto com validação no próprio computador."
            ),
            "acao_docente": "Circular pelo laboratório, tirando dúvidas individuais e avaliando o progresso.",
            "acao_estudante": "Executar o laboratório prático, corrigir erros e registrar a saída esperada."
        },
        {
            "bloco": "45 a 50 min",
            "etapa": "Fechamento & Checkpoint de Aprendizagem",
            "detalhes": (
                "Revisão rápida de 5 minutos: 1 pergunta oral de fixação, registro de visto para quem "
                "concluiu a prática e orientação sobre o próximo passo."
            ),
            "acao_docente": "Carimbar/dar visto nos computadores/cadernos e consolidar o aprendizado.",
            "acao_estudante": "Mostrar a prática concluída ao professor e salvar o trabalho."
        }
    ]

    exercicio_pratico = [
        f"1. Explique com suas palavras a função principal de '{titulo_base}' demonstrada na introdução.",
        "2. Reproduza no laboratório o exemplo prático com as suas próprias configurações/dados.",
        "3. Descreva qual foi o maior desafio ou erro encontrado durante a execução e como foi resolvido."
    ]

    return {
        "titulo_aula": f"Aula Prática: {titulo_base}",
        "turma": turma,
        "tempo_total": f"{tempo_minutos} minutos (1 Tempo de Aula)",
        "url_origem": url,
        "competencias": "Compreensão de arquitetura de sistemas, raciocínio lógico e aplicação técnica prática.",
        "tags": tags_str,
        "roteiro": roteiro_minutagem,
        "exercicio": exercicio_pratico
    }


def gerar_artefatos(aula: dict) -> tuple:
    """Gera o Markdown e o PowerPoint (.pptx)."""
    nome_slug = sanitizar_nome(aula["titulo_aula"])
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    nome_base = f"Aula_{nome_slug}_{ts}"

    caminho_md = f"{DIR_AULAS_GERADAS}/{nome_base}.md"
    caminho_pptx = f"{DIR_AULAS_GERADAS}/{nome_base}.pptx"
    caminho_lista_md = f"{DIR_AULAS_LISTAS}/{nome_base}.md"

    # Monta Markdown
    md_linhas = [
        f"# {aula['titulo_aula']}",
        f"**Público-Alvo / Turma:** {aula['turma']} | **Duração:** {aula['tempo_total']}",
        f"**Data de Planejamento:** {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        f"**Material de Origem:** [{aula['url_origem']}]({aula['url_origem']})",
        f"**Palavras-Chave:** `{aula['tags']}`",
        "\n---",
        "## 🎯 Objetivos de Aprendizagem",
        f"- {aula['competencias']}",
        "- Domínio prático em bancada / laboratório de informática.",
        "\n---",
        "## ⏱️ Roteiro de Aula Detalhado (50 minutos)",
        ""
    ]

    for item in aula["roteiro"]:
        md_linhas.append(f"### 🕒 [{item['bloco']}] — {item['etapa']}")
        md_linhas.append(f"• **Foco Pedagógico:** {item['detalhes']}")
        md_linhas.append(f"• **Papel do Docente:** {item['acao_docente']}")
        md_linhas.append(f"• **Atividade do Aluno:** {item['acao_estudante']}\n")

    md_linhas.append("---")
    md_linhas.append("## 💻 Atividade Prática / Laboratório de Fixação")
    for q in aula["exercicio"]:
        md_linhas.append(f"{q}")

    md_linhas.append("\n---")
    md_linhas.append("## 📝 Critérios para Visto e Avaliação")
    md_linhas.append("• Execução correta do ambiente de laboratório.")
    md_linhas.append("• Resposta objetiva às questões práticas no caderno ou repositório.")
    md_linhas.append("• Participação ativa no encerramento da aula.")

    conteudo_md = "\n".join(md_linhas)

    with open(caminho_md, "w", encoding="utf-8") as f:
        f.write(conteudo_md)

    try:
        with open(caminho_lista_md, "w", encoding="utf-8") as f:
            f.write(conteudo_md)
    except Exception:
        pass

    # Gera PowerPoint se biblioteca pptx estiver disponível
    gerou_pptx = False
    if Presentation:
        try:
            prs = Presentation()
            # Slide 1: Capa
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = aula["titulo_aula"]
            slide.placeholders[1].text = f"IEMA • Educação Técnica e Profissional\n{aula['turma']} • {aula['tempo_total']}"

            # Slide 2: Roteiro e Tempos
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = "Cronograma da Aula (50 min)"
            tf = slide.placeholders[1].text_frame
            for item in aula["roteiro"]:
                p = tf.add_paragraph()
                p.text = f"• [{item['bloco']}] {item['etapa']}"

            # Slide 3: Prática de Laboratório
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = "Mão na Massa: Desafio Prático"
            tf = slide.placeholders[1].text_frame
            for ex in aula["exercicio"]:
                p = tf.add_paragraph()
                p.text = f"• {ex}"

            # Slide 4: Conclusão & Visto
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = "Fechamento & Avaliação"
            tf = slide.placeholders[1].text_frame
            p = tf.add_paragraph()
            p.text = "• Apresente a saída do seu terminal / código para validação."
            p2 = tf.add_paragraph()
            p2.text = f"• Fonte e referências: {aula['url_origem']}"

            prs.save(caminho_pptx)
            gerou_pptx = True
        except Exception as e:
            print(f"Aviso ao gerar pptx: {e}", file=sys.stderr)

    return caminho_md, caminho_pptx if gerou_pptx else None


def formatar_zap(aula: dict, caminho_md: str, caminho_pptx: str = None) -> str:
    """Gera mensagem limpa para visualização imediata no WhatsApp."""
    linhas = [
        f"🎬 *PLANO DE AULA ESTRUTURADO A PARTIR DE LINK*",
        f"📌 *{aula['titulo_aula']}*",
        f"⏱️ *Duração:* {aula['tempo_total']} | Turma: {aula['turma']}",
        f"🔗 *Fonte:* {aula['url_origem']}\n",
        f"🕒 *Distribuição Cronometrada (50 min):*"
    ]

    for item in aula["roteiro"]:
        linhas.append(f"• *{item['bloco']}* — {item['etapa']}")

    linhas.append(f"\n💻 *Atividade de Fixação:* 3 questões práticas prontas.")
    linhas.append(f"📄 *Plano Completo:* `{os.path.basename(caminho_md)}`")
    if caminho_pptx:
        linhas.append(f"📊 *Slides .pptx:* Salvos e prontos para apresentação.")

    return "\n".join(linhas)


def processar_link_aula(url: str, turma: str = "Técnico em Informática") -> dict:
    metadados = extrair_metadados_url(url)
    aula = esmiucar_para_aula(metadados, turma)
    caminho_md, caminho_pptx = gerar_artefatos(aula)
    msg_zap = formatar_zap(aula, caminho_md, caminho_pptx)
    return {
        "status": "SUCESSO",
        "aula": aula,
        "caminho_md": caminho_md,
        "caminho_pptx": caminho_pptx,
        "msg_zap": msg_zap
    }


if __name__ == "__main__":
    url_teste = sys.argv[1] if len(sys.argv) > 1 else "https://www.youtube.com/watch?v=dQw4w9WgXcQ"
    turma_teste = sys.argv[2] if len(sys.argv) > 2 else "2º Ano Técnico - Turma 202"
    print(f"🚀 Processando link: {url_teste}...")
    res = processar_link_aula(url_teste, turma_teste)
    print("\n" + res["msg_zap"])
    print(f"\nSalvo em: {res['caminho_md']}")
