#!/usr/bin/env python3
"""
==============================================================================
GERADOR DE MATERIAL PEDAGÓGICO COMPLETO PARA AS AULAS DE SEGUNDA-FEIRA (IEMA)
==============================================================================
Aulas Atendidas:
1. 10:20 - 12:00 | Turma 201 | INTELIGÊNCIA ARTIFICIAL
2. 13:30 - 15:10 | Turma 200 | INTELIGÊNCIA ARTIFICIAL
3. 16:10 - 17:00 | Turma 202 | LEGISLAÇÃO DE INFORMÁTICA NO BRASIL

Artefatos Gerados:
- Slides (.pptx) para IA (Recomendação, Agentes, RAG, Guardrails, Orquestração)
- Slides (.pptx) para Legislação (Algoritmos, Perfilamento, LGPD, Marco Civil)
- Atividades de Fixação em PDF e Markdown com gabarito comentado
- Sincronização automática com o Google Drive
- Cadastro no banco de dados escola.db (atividades_cadastradas)
==============================================================================
"""

import os
import sys
import sqlite3
import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable

SLIDES_DIR = os.path.join(os.environ.get("DOCENTE_HOME", "."), "mounts/gdrive/Aulas/Slides")
LISTAS_DIR = os.path.join(os.environ.get("DOCENTE_HOME", "."), "mounts/gdrive/Aulas/Listas")
LOCAL_SLIDES = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas_slides")
LOCAL_LISTAS = os.path.join(os.environ.get("DOCENTE_HOME", "."), "aulas_listas")
ESCOLA_DB = os.path.join(os.environ.get("DOCENTE_HOME", "."), "escola.db")

os.makedirs(SLIDES_DIR, exist_ok=True)
os.makedirs(LISTAS_DIR, exist_ok=True)
os.makedirs(LOCAL_SLIDES, exist_ok=True)
os.makedirs(LOCAL_LISTAS, exist_ok=True)

# Cores tema para slides
COR_NAVY = RGBColor(15, 32, 67)       # #0F2043
COR_AZUL_IEMA = RGBColor(0, 102, 204) # #0066CC
COR_CINZA_TEXTO = RGBColor(50, 50, 50)
COR_BRANCO = RGBColor(255, 255, 255)
COR_DESTAQUE = RGBColor(220, 53, 69)

def criar_slide(prs, titulo, subtitulo_ou_pontos, eh_capa=False):
    slide_layout = prs.slide_layouts[6] # layout em branco
    slide = prs.slides.add_slide(slide_layout)
    
    # Barra de topo
    topo = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.18))
    topo.fill.solid()
    topo.fill.fore_color.rgb = COR_AZUL_IEMA
    topo.line.color.rgb = COR_AZUL_IEMA

    if eh_capa:
        # Fundo elegante
        bg = slide.shapes.add_shape(1, Inches(0.5), Inches(0.8), prs.slide_width - Inches(1.0), Inches(5.8))
        bg.fill.solid()
        bg.fill.fore_color.rgb = COR_NAVY
        bg.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), prs.slide_width - Inches(1.6), Inches(2.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = titulo
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = COR_BRANCO
        p.alignment = PP_ALIGN.LEFT

        p2 = tf.add_paragraph()
        p2.text = subtitulo_ou_pontos[0] if isinstance(subtitulo_ou_pontos, list) else subtitulo_ou_pontos
        p2.font.size = Pt(20)
        p2.font.color.rgb = RGBColor(200, 220, 255)
        p2.alignment = PP_ALIGN.LEFT

        tb_foot = slide.shapes.add_textbox(Inches(0.8), Inches(5.2), Inches(8), Inches(0.8))
        tf_foot = tb_foot.text_frame
        p_foot = tf_foot.paragraphs[0]
        p_foot.text = "IEMA — Instituto Estadual de Educação, Ciência e Tecnologia do Maranhão\nProf. Pedro • Ensino Médio Técnico"
        p_foot.font.size = Pt(13)
        p_foot.font.color.rgb = RGBColor(180, 190, 210)
        return slide

    # Slide de conteúdo normal
    # Título do slide
    tb_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), prs.slide_width - Inches(1.6), Inches(1.1))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = titulo
    p_t.font.size = Pt(26)
    p_t.font.bold = True
    p_t.font.color.rgb = COR_NAVY

    # Linha divisória
    linha = slide.shapes.add_shape(1, Inches(0.8), Inches(1.4), Inches(8.4), Inches(0.04))
    linha.fill.solid()
    linha.fill.fore_color.rgb = COR_AZUL_IEMA
    linha.line.color.rgb = COR_AZUL_IEMA

    # Corpo com tópicos
    tb_c = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), prs.slide_width - Inches(1.6), Inches(5.0))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True

    pontos = subtitulo_ou_pontos if isinstance(subtitulo_ou_pontos, list) else [subtitulo_ou_pontos]
    for idx, pt_txt in enumerate(pontos):
        p = tf_c.paragraphs[0] if idx == 0 else tf_c.add_paragraph()
        p.text = f"•  {pt_txt}"
        p.font.size = Pt(18)
        p.font.color.rgb = COR_CINZA_TEXTO
        p.space_after = Pt(14)
        p.line_spacing = 1.25

    return slide

# ==============================================================================
# 1. GERAÇÃO DOS SLIDES DE IA (Turmas 201 e 200)
# ==============================================================================
def gerar_slides_ia():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5) # formato 4:3 / padrão legível

    # Slide 1: Capa
    criar_slide(prs, "Além do Chatbot:\nSistemas de Recomendação, RAG e Agentes de IA",
                "Como a Inteligência Artificial toma decisões e molda o seu feed", eh_capa=True)

    # Slide 2: O Desafio do Feed
    criar_slide(prs, "O Enigma do Feed: Como o Algoritmo Sabe?", [
        "Você abre o TikTok, YouTube ou Spotify e o conteúdo perfeito aparece na tela.",
        "Como a máquina aprende o que você gosta sem ler a sua mente?",
        "Toda interação é um sinal: tempo de tela, curtidas, cliques, pausas e pulos.",
        "O computador não vê vídeos nem ouve áudio diretamente: ele vê números e vetores.",
        "Objetivo da aula: Desvendar a engrenagem matemática por trás da recomendação."
    ])

    # Slide 3: Filtragem Colaborativa vs Baseada em Conteúdo
    criar_slide(prs, "Motores de Recomendação: As Duas Grandes Estratégias", [
        "1. Baseada em Conteúdo: 'Se você gostou do Filme X (Ação + Espionagem), vai gostar de Y com as mesmas tags'.",
        "2. Filtragem Colaborativa: 'Pessoas com histórico igual ao seu também assistiram Z'.",
        "Matriz de Afinidade: Linhas são Usuários, Colunas são Itens e Células são Notas/Interações.",
        "Similaridade de Cosseno: Medição matemática de proximidade entre os perfis de gosto.",
        "O problema do 'Início a Frio' (Cold Start): O que sugerir para um usuário recém-cadastrado?"
    ])

    # Slide 4: O Que São Agentes de IA?
    criar_slide(prs, "Agentes Autônomos: Muito Além de um Prompt", [
        "LLMs tradicionais (ChatGPT) são estáticos: respondem apenas à entrada imediata.",
        "Agentes de IA possuem 4 pilares fundamentais:",
        "  ▫️ Percepção: Observa o ambiente, mensagens e dados que chegam.",
        "  ▫️ Raciocínio & Planejamento: Quebra grandes objetivos em tarefas lógicas.",
        "  ▫️ Memória (Curto e Longo Prazo): Lembra do contexto de conversas anteriores.",
        "  ▫️ Ação & Ferramentas (Tools): Executa consultas, APIs, banco de dados e arquivos."
    ])

    # Slide 5: RAG - Retrieval-Augmented Generation
    criar_slide(prs, "RAG: Conectando a IA com o Mundo Real", [
        "O Grande Problema: IAs puras sofrem de 'Alucinação' (inventam respostas com convicção).",
        "A Solução RAG: Buscar em documentos reais ANTES de gerar a resposta final.",
        "Passo a Passo do RAG:",
        "  1. A pergunta do usuário é convertida em um vetor (Embedding).",
        "  2. O banco vetorial encontra os trechos oficiais mais próximos semanticamente.",
        "  3. O modelo sintetiza a resposta final usando apenas os trechos verdadeiros.",
        "Resultado: Respostas confiáveis, atualizadas e sem invenções!"
    ])

    # Slide 6: Guardrails - A Muralha de Segurança
    criar_slide(prs, "Guardrails: Barreiras de Segurança e Ética", [
        "Guardrail = Trilho de proteção que impede que o sistema descarrilhe.",
        "Filtro de Entrada: Bloqueia ataques de Jailbreak, injeção de prompt e dados maliciosos.",
        "Filtro de Saída: Impede vazamento de dados sigilosos (CPF, senhas), ofensas e desvios.",
        "Controle de Escopo: Garante que um bot escolar só fale sobre escola e recuse outros temas.",
        "Analogia: O Guardrail é o 'árbitro rigoroso' que analisa cada jogada da IA."
    ])

    # Slide 7: Orquestração Multiagente
    criar_slide(prs, "Orquestração Multiagente: Trabalho em Equipe", [
        "Em vez de um único agente tentar fazer tudo, especializamos as funções:",
        "  ▫️ Agente Coletor: Varre registros, conversas e desafios do dia a dia.",
        "  ▫️ Agente Refinador: Estrutura os conceitos, planeja passos práticos e benefícios.",
        "  ▫️ Agente Executor: Aplica os comandos, grava no banco e dispara no WhatsApp.",
        "O Orquestrador gerencia a fila e garante que nenhuma etapa se perca.",
        "Arquitetura modular = mais estabilidade, facilidade de manutenção e rapidez."
    ])

    # Slide 8: Dinâmica em Sala
    criar_slide(prs, "Dinâmica Prática: O Algoritmo Humano em Ação", [
        "Atividade Desplugada em Equipe (15 minutos):",
        "  1. Cada aluno recebe uma 'Ficha de Vetor' com notas de 1 a 5 para 4 estilos musicais.",
        "  2. Formar duplas e calcular a similaridade de gosto com o parceiro.",
        "  3. Recomendar 1 música inédita com base na afinidade calculada.",
        "Desafio dos Guardrails:",
        "  • Criar 1 regra de filtro que barraria uma recomendação inadequada para a escola."
    ])

    # Slide 9: Conclusão & Atividade
    criar_slide(prs, "Conclusão: O Engenheiro de IA do Futuro", [
        "Inteligência Artificial moderna não é apenas decorar comandos no chat.",
        "É arquitetura: saber combinar vetores, documentos reais (RAG) e regras de segurança.",
        "Próximo Passo: Resolver a Lista de Fixação de Hoje (individual ou em dupla).",
        "A entrega vale visto no diário de classe da disciplina!",
        "Dúvidas? Pergunte agora ou no laboratório!"
    ])

    caminho_local = os.path.join(LOCAL_SLIDES, "Algoritmos_Recomendacao_Agentes_RAG.pptx")
    caminho_gdrive = os.path.join(SLIDES_DIR, "Algoritmos_Recomendacao_Agentes_RAG.pptx")
    prs.save(caminho_local)
    prs.save(caminho_gdrive)
    print(f"✅ Slides IA salvos em: {caminho_gdrive}")

# ==============================================================================
# 2. GERAÇÃO DOS SLIDES DE LEGISLAÇÃO (Turma 202)
# ==============================================================================
def gerar_slides_legislacao():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Slide 1: Capa
    criar_slide(prs, "Algoritmos de Recomendação, IA e a Lei:\nLGPD, Perfilamento e Direitos Digitais",
                "Turma 202 • Legislação de Informática no Brasil • Prof. Pedro", eh_capa=True)

    # Slide 2: O Problema do Perfilamento
    criar_slide(prs, "O Algoritmo Sabe Demais: O Problema do Perfilamento", [
        "Perfilamento (Profiling): Análise automatizada de dados para avaliar a personalidade humana.",
        "O que as plataformas sabem sobre você?",
        "  ▫️ Seus interesses políticos, hábitos de sono, localização em tempo real e poder aquisitivo.",
        "Risco social: Bolhas de desinformação, discriminação de preços e manipulação de consumo.",
        "A pergunta jurídica: Até onde as empresas têm o direito de mapear o seu comportamento?"
    ])

    # Slide 3: LGPD e Dados Pessoais
    criar_slide(prs, "A LGPD (Lei 13.709/2018) no Contexto da IA", [
        "Dado Pessoal: Qualquer informação que identifique ou possa identificar uma pessoa viva.",
        "Dado Sensível: Origem racial, convicção religiosa, saúde, biometria e filiação sindical.",
        "Princípios Fundamentais para Algoritmos de IA:",
        "  ▫️ Finalidade: Coletar apenas para o objetivo prometido ao usuário.",
        "  ▫️ Necessidade (Minimização): Usar o mínimo de dados possível para o modelo funcionar.",
        "  ▫️ Transparência: O usuário tem o direito de saber quais critérios o algoritmo usa."
    ])

    # Slide 4: Artigo 20 da LGPD - O Direito à Explicação
    criar_slide(prs, "Artigo 20 da LGPD: A 'Caixa Preta' da IA", [
        "Art. 20: 'O titular dos dados tem direito a solicitar revisão de decisões tomadas unicamente com base em tratamento automatizado'.",
        "Aplicações reais:",
        "  • Recusa de crédito bancário por algoritmo.",
        "  • Descarte de currículo em processos seletivos por IA.",
        "  • Bloqueio ou moderação de conta em redes sociais.",
        "Direito à Explicação: A empresa é obrigada a fornecer informações claras sobre os critérios."
    ])

    # Slide 5: Viés Algorítmico e Discriminação
    criar_slide(prs, "Viés Algorítmico: Quando o Código Herda Preconceitos", [
        "A IA aprende com dados do passado: se o passado teve desigualdade, a IA a reproduzirá.",
        "Casos Históricos Notórios:",
        "  ▫️ Reconhecimento facial errando muito mais em rostos negros e mulheres.",
        "  ▫️ Algoritmo de recrutamento da Amazon que descartava candidatas mulheres para vagas de TI.",
        "Responsabilidade Civil: A empresa desenvolvedora responde objetivamente pelos danos causados."
    ])

    # Slide 6: Marco Civil da Internet (Lei 12.965/2014)
    criar_slide(prs, "O Marco Civil da Internet e a Neutralidade", [
        "Marco Civil da Internet: A 'Constituição da Internet' no Brasil.",
        "Neutralidade da Rede (Art. 9º): Provedores não podem privilegiar nem degradar pacotes.",
        "Liberdade de Expressão e Responsabilidade de Provedores (Art. 19):",
        "  • Plataformas respondem por conteúdos de terceiros somente após ordem judicial específica.",
        "O debate atual: A recomendação algorítmica ativa de conteúdos tóxicos deve mudar essa regra?"
    ])

    # Slide 7: Guardrails Jurídicos na Prática
    criar_slide(prs, "Guardrails Jurídicos: Compliance no Código", [
        "O profissional de TI não pode programar sem conhecer a lei.",
        "Privacy by Design (Privacidade desde a Concepção):",
        "  • Implementar consentimento explícito no formulário de cadastro.",
        "  • Não guardar senhas nem logs com dados sensíveis sem criptografia.",
        "  • Criar mecanismos de anonimização nos dados de treino de modelos de recomendação.",
        "Programar com ética é o que diferencia um técnico comum de um profissional de elite."
    ])

    # Slide 8: Estudo de Caso Clássico
    criar_slide(prs, "Estudo de Caso: A Loja Target e a Menina Grávida", [
        "Caso Real (EUA):",
        "A rede de lojas Target analisou padrões de compras (loções sem cheiro, vitaminas) e previu a gravidez de uma adolescente.",
        "A loja enviou cupons de fraldas para a casa dela antes mesmo que o pai dela soubesse da gestação.",
        "Reflexão para a turma:",
        "  • Houve violação de privacidade?",
        "  • Como a LGPD brasileira puniria essa situação hoje?",
        "  • Qual o limite entre utilidade de negócio e invasão de intimidade?"
    ])

    # Slide 9: Fechamento da Aula
    criar_slide(prs, "Fechamento & Atividade de Fixação", [
        "A tecnologia avança na velocidade dos algoritmos; o Direito avança na velocidade dos humanos.",
        "O bom técnico de informática domina a técnica E protege os direitos fundamentais.",
        "Atividade Individual de Fixação: Responder a Lista de Casos Práticos de Legislação.",
        "Entrega e visto ao final da aula para cômputo no diário de bordo!"
    ])

    caminho_local = os.path.join(LOCAL_SLIDES, "Recomendacao_IA_e_Legislacao_LGPD.pptx")
    caminho_gdrive = os.path.join(SLIDES_DIR, "Recomendacao_IA_e_Legislacao_LGPD.pptx")
    prs.save(caminho_local)
    prs.save(caminho_gdrive)
    print(f"✅ Slides Legislação salvos em: {caminho_gdrive}")

# ==============================================================================
# 3. GERAÇÃO DAS LISTAS DE EXERCÍCIOS EM PDF E MARKDOWN
# ==============================================================================
def gerar_pdf_atividade(arquivo_pdf, titulo_doc, subtitulo_doc, questoes, gabarito):
    doc = SimpleDocTemplate(arquivo_pdf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()

    # Estilos customizados
    style_inst = ParagraphStyle('Inst', parent=styles['Normal'], fontSize=9, leading=12, alignment=1, textColor=colors.HexColor('#0F2043'))
    style_tit = ParagraphStyle('Tit', parent=styles['Heading1'], fontSize=15, leading=18, alignment=1, textColor=colors.HexColor('#0066CC'), spaceAfter=4)
    style_sub = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, leading=13, alignment=1, textColor=colors.HexColor('#555555'), spaceAfter=10)
    style_q_tit = ParagraphStyle('QTit', parent=styles['Normal'], fontSize=10, leading=14, fontName='Helvetica-Bold', textColor=colors.HexColor('#0F2043'), spaceBefore=8, spaceAfter=4)
    style_q_txt = ParagraphStyle('QTxt', parent=styles['Normal'], fontSize=9, leading=13, textColor=colors.HexColor('#222222'), spaceAfter=6)
    style_gab_tit = ParagraphStyle('GabTit', parent=styles['Heading2'], fontSize=12, leading=15, textColor=colors.HexColor('#0F2043'), spaceBefore=12, spaceAfter=6)
    style_gab_txt = ParagraphStyle('GabTxt', parent=styles['Normal'], fontSize=8.5, leading=12, textColor=colors.HexColor('#444444'), spaceAfter=4)

    story = []

    # Cabeçalho da Escola
    story.append(Paragraph("<b>GOVERNO DO ESTADO DO MARANHÃO — SEDUC</b>", style_inst))
    story.append(Paragraph("<b>IEMA — INSTITUTO ESTADUAL DE EDUCAÇÃO, CIÊNCIA E TECNOLOGIA DO MARANHÃO</b>", style_inst))
    story.append(Paragraph("CURSO TÉCNICO EM INFORMÁTICA • PROF. PEDRO OLIVEIRA", style_inst))
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0066CC'), spaceBefore=2, spaceAfter=8))

    story.append(Paragraph(titulo_doc, style_tit))
    story.append(Paragraph(subtitulo_doc, style_sub))

    # Campo de identificação do estudante
    ident_data = [
        [Paragraph("<b>Aluno(a):</b> ________________________________________________", style_q_txt),
         Paragraph("<b>Nº:</b> ______", style_q_txt),
         Paragraph("<b>Data:</b> ___/___/2026", style_q_txt)]
    ]
    t_ident = Table(ident_data, colWidths=[330, 80, 110])
    t_ident.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(t_ident)
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor('#CCCCCC'), spaceBefore=2, spaceAfter=8))

    # Questões
    for q in questoes:
        story.append(Paragraph(f"<b>Questão {q['numero']}:</b> {q['enunciado']}", style_q_tit))
        if q.get('opcoes'):
            for op in q['opcoes']:
                story.append(Paragraph(op, style_q_txt))
        if q.get('linhas_resposta', 0) > 0:
            story.append(Spacer(1, 4))
            for _ in range(q['linhas_resposta']):
                story.append(Paragraph("__________________________________________________________________________________________", style_q_txt))
        story.append(Spacer(1, 4))

    # Quebra de página para o Gabarito Comentado (Uso do Professor)
    story.append(Spacer(1, 14))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#0066CC'), spaceBefore=4, spaceAfter=8))
    story.append(Paragraph("<b>GABARITO COMENTADO (USO DO PROFESSOR / CORREÇÃO EM SALA)</b>", style_gab_tit))
    for g in gabarito:
        story.append(Paragraph(f"<b>Questão {g['numero']}:</b> <b>{g['resposta']}</b> — <i>{g['justificativa']}</i>", style_gab_txt))

    doc.build(story)
    print(f"📄 PDF gerado com sucesso: {arquivo_pdf}")

def salvar_markdown_lista(caminho_md, titulo, questoes, gabarito):
    linhas = [
        f"# {titulo}",
        "\n**IEMA — Instituto Estadual de Educação, Ciência e Tecnologia do Maranhão**",
        "**Docente:** Prof. Pedro | **Ano Letivo:** 2026",
        "\n---\n"
    ]
    for q in questoes:
        linhas.append(f"### Questão {q['numero']}: {q['enunciado']}\n")
        if q.get('opcoes'):
            for op in q['opcoes']:
                linhas.append(f"- {op}")
        if q.get('linhas_resposta', 0) > 0:
            linhas.append("\n*(Espaço para resposta discursiva do aluno)*\n")
        linhas.append("")

    linhas.append("\n---\n## Gabarito Comentado para o Docente\n")
    for g in gabarito:
        linhas.append(f"- **Questão {g['numero']}:** {g['resposta']}")
        linhas.append(f"  - *Justificativa:* {g['justificativa']}\n")

    with open(caminho_md, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))
    print(f"📄 Markdown gerado com sucesso: {caminho_md}")

# ==============================================================================
# DEFINIÇÃO DAS ATIVIDADES DE FIXAÇÃO
# ==============================================================================
def criar_atividades():
    # 1. Atividade de Inteligência Artificial (Turmas 201 e 200)
    questoes_ia = [
        {
            "numero": 1,
            "enunciado": "As plataformas de streaming utilizam algoritmos de recomendação para engajar os usuários. Sobre a técnica de <b>Filtragem Colaborativa</b>, é CORRETO afirmar:",
            "opcoes": [
                "(A) Ela analisa unicamente o texto e os metadados dos vídeos, sem comparar usuários entre si.",
                "(B) Ela encontra correlações e padrões entre usuários que tiveram históricos de consumo semelhantes para indicar novos itens.",
                "(C) Ela exige que o usuário preencha um questionário formal com todos os seus gostos antes de receber qualquer sugestão.",
                "(D) Ela só funciona se o modelo de IA tiver acesso irrestrito às fotos e mensagens privadas dos usuários.",
                "(E) Ela gera resultados aleatórios toda vez que a página do aplicativo é atualizada."
            ]
        },
        {
            "numero": 2,
            "enunciado": "Modelos de linguagem (LLMs) frequentemente sofrem de 'alucinação', gerando respostas convincentes, porém incorretas. A arquitetura <b>RAG (Retrieval-Augmented Generation)</b> soluciona esse problema ao:",
            "opcoes": [
                "(A) Aumentar a velocidade do processador da máquina para responder em menos de 100 milissegundos.",
                "(B) Desativar completamente a rede neural e utilizar apenas buscas simples por palavras-chave no Google.",
                "(C) Recuperar trechos de documentos oficiais em um banco vetorial confiável ANTES de formular a resposta final.",
                "(D) Limitar a resposta da IA a no máximo cinco palavras por interação.",
                "(E) Reiniciar o treinamento da inteligência artificial do zero a cada pergunta feita pelo usuário."
            ]
        },
        {
            "numero": 3,
            "enunciado": "No ecossistema de agentes inteligentes, um componente fundamental é o <b>Guardrail</b>. Qual é a principal função prática de um Guardrail em um sistema de IA?",
            "opcoes": [
                "(A) Fazer a renderização gráfica dos botões na tela do smartphone.",
                "(B) Bloquear ataques maliciosos (como Jailbreak), impedir vazamento de dados sensíveis e barrar alucinações de saída.",
                "(C) Compactar o banco de dados para que ele ocupe menos memória no disco rígido.",
                "(D) Traduzir automaticamente qualquer código Python para a linguagem HTML.",
                "(E) Eliminar a necessidade de o sistema possuir qualquer conexão com a internet."
            ]
        },
        {
            "numero": 4,
            "enunciado": "Diferencie um <b>Chatbot Tradicional</b> de um <b>Agente de IA Autônomo</b>, explicando o papel das <i>Ferramentas (Tools)</i> e da <i>Memória</i> na tomada de decisão.",
            "linhas_resposta": 4
        },
        {
            "numero": 5,
            "enunciado": "<b>Estudo de Caso Prático:</b> Um hospital quer criar um assistente de IA para ajudar médicos a encontrar informações sobre prontuários de pacientes. Explique como o <b>RAG</b> e os <b>Guardrails</b> devem ser combinados para garantir que o assistente seja preciso e não viole o sigilo médico.",
            "linhas_resposta": 5
        }
    ]

    gabarito_ia = [
        {"numero": 1, "resposta": "Alternativa (B)", "justificativa": "A Filtragem Colaborativa baseia-se na máxima de que pessoas que concordaram no passado tendem a concordar no futuro. Ela cruza padrões de interações entre múltiplos usuários na matriz de afinidade."},
        {"numero": 2, "resposta": "Alternativa (C)", "justificativa": "O RAG ancora a resposta em fatos recuperados semanticamente de uma base de conhecimento validada (Embeddings + Vector Search), reduzindo drasticamente a alucinação generativa."},
        {"numero": 3, "resposta": "Alternativa (B)", "justificativa": "Guardrails atuam como filtros bilaterais de segurança e conformidade, inspecionando prompts de entrada para barrar injeções e analisando saídas para evitar vazamentos e respostas fora do escopo."},
        {"numero": 4, "resposta": "Discursiva", "justificativa": "O chatbot tradicional apenas mapeia entradas e saídas estáticas baseadas em prompt. O Agente Autônomo possui ciclo de Percepção-Raciocínio-Ação, é capaz de manter memória de contexto e chamar ferramentas externas (APIs, bancos, scripts) para executar ações no mundo real."},
        {"numero": 5, "resposta": "Discursiva", "justificativa": "O RAG é indispensável para extrair estritamente os fatos do histórico clínico real do paciente sem inventar remédios ou diagnósticos falsos. O Guardrail garante que nenhum dado pessoal de outro paciente vaze na resposta e restringe o acesso de cada médico apenas aos seus respectivos pacientes autorizados."}
    ]

    # 2. Atividade de Legislação de Informática (Turma 202)
    questoes_leg = [
        {
            "numero": 1,
            "enunciado": "A Lei Geral de Proteção de Dados (LGPD - Lei 13.709/2018) estabelece regras claras sobre o tratamento de dados pessoais. No contexto de algoritmos de perfilamento e recomendação, o <b>Princípio da Necessidade (Minimização)</b> exige que:",
            "opcoes": [
                "(A) As empresas coletem todos os dados possíveis do celular do usuário para melhorar o algoritmo.",
                "(B) O aplicativo colete apenas os dados estritamente necessários e pertinentes para a finalidade declarada do serviço.",
                "(C) O usuário seja obrigado a ceder seus dados biométricos para poder navegar na internet.",
                "(D) Os desenvolvedores guardem os dados para sempre, mesmo após o usuário cancelar a sua conta.",
                "(E) Toda inteligência artificial seja desenvolvida exclusivamente pelo poder público."
            ]
        },
        {
            "numero": 2,
            "enunciado": "O <b>Artigo 20 da LGPD</b> é uma das ferramentas mais importantes de defesa do cidadão na era da IA. Ele garante que o titular dos dados tem o direito de:",
            "opcoes": [
                "(A) Receber gratuitamente um computador novo da empresa caso o seu dispositivo quebre.",
                "(B) Solicitar a revisão e explicações sobre decisões tomadas unicamente com base em tratamento automatizado de seus dados.",
                "(C) Exigir que todos os filmes e séries sejam disponibilizados sem anúncios na internet.",
                "(D) Proibir o uso de qualquer computador em escolas e universidades do país.",
                "(E) Registrar códigos de programas de computador em cartório sem pagar taxas."
            ]
        },
        {
            "numero": 3,
            "enunciado": "Sobre o <b>Marco Civil da Internet (Lei 12.965/2014)</b> e a responsabilidade das plataformas digitais por conteúdos de terceiros (Artigo 19), é CORRETO afirmar:",
            "opcoes": [
                "(A) As plataformas são punidas criminalmente de forma automática por qualquer postagem feita por usuários.",
                "(B) A internet no Brasil deve ter velocidade controlada e reduzida durante o horário escolar.",
                "(C) Como regra geral, o provedor só responde civilmente por danos de conteúdo de terceiros se, após ordem judicial específica, não tomar providências para indisponibilizá-lo.",
                "(D) Qualquer cidadão pode solicitar a prisão de um programador caso discorde de uma linha de código.",
                "(E) É proibido usar criptografia em redes sociais que operam no território brasileiro."
            ]
        },
        {
            "numero": 4,
            "enunciado": "Explique o conceito de <b>Viés Algorítmico</b> e analise por que um sistema de IA treinado com dados históricos da sociedade pode perpetuar e amplificar discriminações no mercado de trabalho ou no sistema de justiça.",
            "linhas_resposta": 4
        },
        {
            "numero": 5,
            "enunciado": "<b>Estudo de Caso:</b> Um aplicativo de seleção de candidatos para vagas de estágio em empresas de tecnologia utilizou um algoritmo que descartava automaticamente currículos de pessoas moradoras de determinadas regiões periféricas. Avalie essa prática à luz do <i>Art. 20 da LGPD</i> e dos <i>direitos à não-discriminação e transparência</i>.",
            "linhas_resposta": 5
        }
    ]

    gabarito_leg = [
        {"numero": 1, "resposta": "Alternativa (B)", "justificativa": "O princípio da necessidade (Art. 6º, III da LGPD) limita o tratamento ao mínimo necessário para a realização das finalidades, vedando a coleta abusiva de dados excedentes."},
        {"numero": 2, "resposta": "Alternativa (B)", "justificativa": "O Art. 20 garante ao titular o direito de pedir revisão humana/explicativa de decisões automatizadas que afetem seus interesses, como perfis de crédito ou seleção profissional."},
        {"numero": 3, "resposta": "Alternativa (C)", "justificativa": "O Art. 19 do Marco Civil visa assegurar a liberdade de expressão e evitar a censura privada, exigindo prévia ordem judicial específica para responsabilização secundária do provedor de aplicações."},
        {"numero": 4, "resposta": "Discursiva", "justificativa": "O viés algorítmico surge quando os dados de treinamento refletem preconceitos históricos, disparidades sociais ou amostragens incompletas. Como a IA busca otimizar correlações estatísticas sobre o passado, ela tenderá a reproduzir e institucionalizar essas exclusões se não houver auditoria e guardrails de equidade."},
        {"numero": 5, "resposta": "Discursiva", "justificativa": "A prática é manifestamente ilegal. Viola o princípio da não-discriminação (Art. 6º, IX da LGPD) por adotar critério abusivo e discriminatório geográfico, e fere o Art. 20, pois os candidatos têm direito à revisão humana e à explicação dos critérios que motivaram o descarte automatizado de suas candidaturas."}
    ]

    # Gera arquivos de IA
    pdf_ia = os.path.join(LISTAS_DIR, "Atividade_Fixacao_Recomendacao_Agentes_RAG.pdf")
    md_ia = os.path.join(LISTAS_DIR, "Atividade_Fixacao_Recomendacao_Agentes_RAG.md")
    pdf_ia_local = os.path.join(LOCAL_LISTAS, "Atividade_Fixacao_Recomendacao_Agentes_RAG.pdf")
    md_ia_local = os.path.join(LOCAL_LISTAS, "Atividade_Fixacao_Recomendacao_Agentes_RAG.md")

    gerar_pdf_atividade(pdf_ia, "ATIVIDADE DE FIXAÇÃO: ALGORITMOS DE RECOMENDAÇÃO & AGENTES DE IA",
                        "Componente: Inteligência Artificial • Turmas 201 e 200", questoes_ia, gabarito_ia)
    gerar_pdf_atividade(pdf_ia_local, "ATIVIDADE DE FIXAÇÃO: ALGORITMOS DE RECOMENDAÇÃO & AGENTES DE IA",
                        "Componente: Inteligência Artificial • Turmas 201 e 200", questoes_ia, gabarito_ia)
    salvar_markdown_lista(md_ia, "Atividade de Fixação: Algoritmos de Recomendação & Agentes de IA", questoes_ia, gabarito_ia)
    salvar_markdown_lista(md_ia_local, "Atividade de Fixação: Algoritmos de Recomendação & Agentes de IA", questoes_ia, gabarito_ia)

    # Gera arquivos de Legislação
    pdf_leg = os.path.join(LISTAS_DIR, "Atividade_Fixacao_Legislacao_Algoritmos_LGPD.pdf")
    md_leg = os.path.join(LISTAS_DIR, "Atividade_Fixacao_Legislacao_Algoritmos_LGPD.md")
    pdf_leg_local = os.path.join(LOCAL_LISTAS, "Atividade_Fixacao_Legislacao_Algoritmos_LGPD.pdf")
    md_leg_local = os.path.join(LOCAL_LISTAS, "Atividade_Fixacao_Legislacao_Algoritmos_LGPD.md")

    gerar_pdf_atividade(pdf_leg, "ATIVIDADE DE FIXAÇÃO: IA, ALGORITMOS E A LEGISLAÇÃO BRASILEIRA",
                        "Componente: Legislação de Informática no Brasil • Turma 202", questoes_leg, gabarito_leg)
    gerar_pdf_atividade(pdf_leg_local, "ATIVIDADE DE FIXAÇÃO: IA, ALGORITMOS E A LEGISLAÇÃO BRASILEIRA",
                        "Componente: Legislação de Informática no Brasil • Turma 202", questoes_leg, gabarito_leg)
    salvar_markdown_lista(md_leg, "Atividade de Fixação: IA, Algoritmos e Legislação Brasileira (LGPD)", questoes_leg, gabarito_leg)
    salvar_markdown_lista(md_leg_local, "Atividade de Fixação: IA, Algoritmos e Legislação Brasileira (LGPD)", questoes_leg, gabarito_leg)

# ==============================================================================
# 4. CADASTRO DAS ATIVIDADES NO BANCO (ESCOLA.DB)
# ==============================================================================
def registrar_atividades_banco():
    if not os.path.exists(ESCOLA_DB):
        return
    agora_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    hoje_data = "19/10/2026"

    atvs = [
        ("201", "INTELIGÊNCIA ARTIFICIAL", "Atividade Prática: Algoritmos de Recomendação, RAG e Guardrails",
         "Resolução da lista de fixação sobre motores de recomendação, vetores de similaridade e agentes autônomos."),
        ("200", "INTELIGÊNCIA ARTIFICIAL", "Atividade Prática: Algoritmos de Recomendação, RAG e Guardrails",
         "Resolução da lista de fixação sobre motores de recomendação, vetores de similaridade e agentes autônomos."),
        ("202", "LEGISLAÇÃO DE INFORMÁTICA NO BRASIL", "Atividade Prática: LGPD, Decisões Automatizadas e Viés Algorítmico",
         "Estudo de casos reais sobre Art. 20 da LGPD, perfilamento algorítmico e responsabilidade civil em IA.")
    ]

    try:
        with sqlite3.connect(ESCOLA_DB) as conn:
            cur = conn.cursor()
            for turma, comp, titulo, desc in atvs:
                # Evita duplicar se já foi cadastrada hoje
                cur.execute("SELECT id FROM atividades_cadastradas WHERE turma = ? AND titulo = ?", (turma, titulo))
                if not cur.fetchone():
                    cur.execute("""
                        INSERT INTO atividades_cadastradas (data_passada, data_hora_iso, turma, componente, titulo, descricao)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (hoje_data, agora_iso, turma, comp, titulo, desc))
            conn.commit()
            print("💾 Atividades cadastradas com sucesso no escola.db para as turmas 201, 200 e 202!")
    except Exception as e:
        print(f"⚠️ Erro ao registrar atividades no banco: {e}")

if __name__ == "__main__":
    print("🚀 Iniciando geração do material didático para as aulas de segunda-feira...")
    gerar_slides_ia()
    gerar_slides_legislacao()
    criar_atividades()
    registrar_atividades_banco()
    print("🎉 Todo o material didático foi gerado e sincronizado com sucesso!")
