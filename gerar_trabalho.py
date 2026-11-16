#!/usr/bin/env python3
"""Gerador e construtor de modelos acadêmicos do SCEN (Seminário Cristão Evangélico do Norte).

Gera os 12 modelos padronizados em `.docx` e `.pdf`, além da pasta
`Informacoes_Essenciais/` com as fichas de dados essenciais (`dados_essenciais.json`),
versões em `.docx` e `.pdf` de cada trabalho, e permite gerar trabalhos finalizados
em `.docx` e `.pdf` em `Trabalhos Gerados/`.
"""

from __future__ import annotations

import argparse
import html
import json
from pathlib import Path
from typing import Any

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
try:
    import weasyprint
except (ImportError, OSError):
    weasyprint = None

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


BASE_DIR = Path(__file__).resolve().parent
MODELOS_DIR = BASE_DIR / "Modelos"
INFO_ESSENCIAIS_DIR = BASE_DIR / "Informacoes_Essenciais"
GERADOS_DIR = BASE_DIR / "Trabalhos Gerados"

INSTITUICAO_PADRAO = "SEMINÁRIO CRISTÃO EVANGÉLICO DO NORTE"
CIDADE_PADRAO = "São Luís"


# ==============================================================================
# HELPERS DOCX
# ==============================================================================


def _configurar_documento() -> docx.Document:
    """Cria um documento A4 com margens (2,54 cm) e fonte Times New Roman 12 extraídas dos modelos base."""
    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.01)
    sec.page_height = Cm(29.69)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(2.54)
    sec.right_margin = Cm(2.54)
    sec.header_distance = Cm(1.27)
    sec.footer_distance = Cm(1.27)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    return doc


def _add_p(
    doc: docx.Document,
    texto: str = "",
    *,
    align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.LEFT,
    bold: bool = False,
    italic: bool = False,
    size: int = 12,
    line_spacing: float = 1.5,
    first_line_indent: float | None = None,
    left_indent: float | None = None,
    space_after: float = 0,
    space_before: float = 0,
) -> docx.text.paragraph.Paragraph:
    """Adiciona um parágrafo formatado em Arial."""
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    if first_line_indent is not None:
        pf.first_line_indent = Cm(first_line_indent)
    if left_indent is not None:
        pf.left_indent = Cm(left_indent)

    if texto:
        run = p.add_run(texto)
        run.font.name = "Times New Roman"
        run.font.size = Pt(size)
        run.bold = bold
        run.italic = italic
    return p


def _add_linhas_vazias(doc: docx.Document, qtd: int, line_spacing: float = 1.5) -> None:
    """Adiciona quebras de linha com espaçamento controlado."""
    for _ in range(qtd):
        _add_p(doc, "", line_spacing=line_spacing)


def _add_capa_e_folha_rosto(
    doc: docx.Document,
    dados: dict[str, Any],
    natureza_texto: str,
) -> None:
    """Constrói a Capa e a Folha de Rosto com posicionamento e estilo exatos dos modelos base."""
    instituicao = dados.get("instituicao", INSTITUICAO_PADRAO)
    aluno = dados.get("aluno", "{{NOME_DO_ALUNO}}")
    titulo = dados.get("titulo", "{{TITULO_DO_TRABALHO}}")
    subtitulo = dados.get("subtitulo", "")
    cidade = dados.get("cidade", CIDADE_PADRAO)
    ano = str(dados.get("ano", "{{ANO}}"))
    professor = dados.get("professor", "{{NOME_DO_PROFESSOR}}")

    # === CAPA ===
    _add_p(doc, instituicao.upper(), align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, line_spacing=1.0)
    _add_linhas_vazias(doc, 4, line_spacing=1.0)
    _add_p(doc, aluno, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    _add_linhas_vazias(doc, 9, line_spacing=1.0)
    if subtitulo:
        _add_p(doc, f"{titulo}:", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, line_spacing=1.0)
        _add_p(doc, subtitulo, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, line_spacing=1.0)
    else:
        _add_p(doc, titulo, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, line_spacing=1.0)
    _add_linhas_vazias(doc, 12, line_spacing=1.0)
    _add_p(doc, cidade, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    _add_p(doc, ano, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    doc.add_page_break()

    # === FOLHA DE ROSTO ===
    _add_p(doc, aluno, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    _add_linhas_vazias(doc, 12, line_spacing=1.0)
    if subtitulo:
        _add_p(doc, f"{titulo}:", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, line_spacing=1.0)
        _add_p(doc, subtitulo, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, line_spacing=1.0)
    else:
        _add_p(doc, titulo, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, line_spacing=1.0)
    _add_linhas_vazias(doc, 2, line_spacing=1.0)
    _add_p(
        doc,
        natureza_texto,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        size=10,
        line_spacing=1.0,
        left_indent=8.0,
        space_after=6,
    )
    _add_p(
        doc,
        f"Professor(a): {professor}",
        align=WD_ALIGN_PARAGRAPH.LEFT,
        size=10,
        line_spacing=1.0,
        left_indent=8.0,
    )
    _add_linhas_vazias(doc, 10, line_spacing=1.0)
    _add_p(doc, cidade, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    _add_p(doc, ano, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
    doc.add_page_break()


def _set_cell_border(cell: docx.table._Cell) -> None:
    """Aplica borda simples preta às células da tabela."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        borders.append(b)
    tc_pr.append(borders)


def _set_cell_shading(cell: docx.table._Cell, fill_hex: str) -> None:
    """Aplica cor de fundo (shading) à célula da tabela."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def _format_cell(
    cell: docx.table._Cell,
    texto: str,
    bold: bool = False,
    align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER,
    fill_hex: str | None = None,
) -> None:
    """Formata célula da tabela em Times New Roman 12."""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _set_cell_border(cell)
    if fill_hex:
        _set_cell_shading(cell, fill_hex)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(texto)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.bold = bold


def _add_referencias(doc: docx.Document, refs: list[str]) -> None:
    """Adiciona seção REFERÊNCIAS no padrão ABNT/SCEN."""
    doc.add_page_break()
    _add_p(doc, "REFERÊNCIAS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    _add_linhas_vazias(doc, 1)
    for ref in refs:
        _add_p(doc, ref, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.0, space_after=12)


def _add_secoes(doc: docx.Document, secoes: list[dict[str, Any]]) -> None:
    """Adiciona seções numeradas com título em negrito e parágrafos com recuo de 1,5 cm."""
    for sec in secoes:
        _add_p(
            doc,
            sec["titulo"],
            align=WD_ALIGN_PARAGRAPH.LEFT,
            bold=True,
            space_before=6,
            space_after=6,
        )
        for paragrafo in sec.get("paragrafos", [sec.get("conteudo", "")]):
            if paragrafo:
                _add_p(
                    doc,
                    paragrafo,
                    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                    first_line_indent=1.5,
                    space_after=6,
                )


# ==============================================================================
# CONSTRUTORES DOCX POR TIPO
# ==============================================================================


def _titulo_completo(dados: dict[str, Any]) -> str:
    titulo = dados.get("titulo", "{{TITULO_DO_TRABALHO}}")
    subtitulo = dados.get("subtitulo", "")
    return f"{titulo}: {subtitulo}" if subtitulo else titulo


def _cronograma_padrao(dados: dict[str, Any]) -> list[dict[str, str]]:
    return dados.get(
        "cronograma",
        [
            {"secao": "Prefácio / Introdução", "titulo": "{{TITULO_SECAO}}", "datas": "{{DD/MM/AAAA}}", "horas": "{{XX minutos}}"},
            {"secao": "Capítulo 1", "titulo": "{{TITULO_CAPITULO_1}}", "datas": "{{DD/MM/AAAA}}", "horas": "{{XX minutos}}"},
            {"secao": "Capítulo 2", "titulo": "{{TITULO_CAPITULO_2}}", "datas": "{{DD/MM/AAAA}}", "horas": "{{XX minutos}}"},
            {"secao": "Conclusão", "titulo": "{{TITULO_CONCLUSAO}}", "datas": "{{DD/MM/AAAA}}", "horas": "{{XX minutos}}"},
        ],
    )


def _secoes_sermao(dados: dict[str, Any], genero: str) -> list[dict[str, Any]]:
    passagem = dados.get("passagem", "{{TEXTO_BIBLICO}}")
    especificos = {
        "Narrativa": [
            {"titulo": "1 DADOS GERAIS DO SERMÃO (NARRATIVA BÍBLICA)", "paragrafos": [f"Texto Base: {passagem}", "Tema / Ideia Central: {{IDEA_CENTRAL}}"]},
            {"titulo": "2 CONTEXTO NARRATIVO (CENÁRIO, PERSONAGENS E CONFLITO)", "paragrafos": ["{{CONTEXTO_NARRATIVO}}"]},
            {"titulo": "3 INTRODUÇÃO", "paragrafos": ["{{INTRODUCAO}}"]},
            {"titulo": "4 DESENVOLVIMENTO DA NARRATIVA (MOVIMENTOS E CLÍMAX)", "paragrafos": ["{{DESENVOLVIMENTO}}"]},
            {"titulo": "5 APLICAÇÃO E CONEXÃO COM A REDENÇÃO", "paragrafos": ["{{APLICACAO}}"]},
            {"titulo": "6 CONCLUSÃO E APELO", "paragrafos": ["{{CONCLUSAO}}"]},
        ],
        "Evangelho": [
            {"titulo": "1 DADOS GERAIS DO SERMÃO (EVANGELHOS)", "paragrafos": [f"Texto Base: {passagem}", "Tema / Ideia Central: {{IDEA_CENTRAL}}"]},
            {"titulo": "2 CONTEXTO DO MINISTÉRIO DE JESUS E AUDIÊNCIA ORIGINAL", "paragrafos": ["{{CONTEXTO_EVANGELHO}}"]},
            {"titulo": "3 INTRODUÇÃO", "paragrafos": ["{{INTRODUCAO}}"]},
            {"titulo": "4 EXPOSIÇÃO DO TEXTO (PONTOS PRINCIPAIS)", "paragrafos": ["{{DESENVOLVIMENTO}}"]},
            {"titulo": "5 FOCO CRISTOLÓGICO E APLICAÇÃO PRÁTICA", "paragrafos": ["{{APLICACAO}}"]},
            {"titulo": "6 CONCLUSÃO", "paragrafos": ["{{CONCLUSAO}}"]},
        ],
        "Parábola": [
            {"titulo": "1 DADOS GERAIS DO SERMÃO (PARÁBOLA)", "paragrafos": [f"Texto Base: {passagem}", "Tema / Ideia Central: {{IDEA_CENTRAL}}"]},
            {"titulo": "2 OCASIÃO DA PARÁBOLA E ELEMENTOS CULTURAIS", "paragrafos": ["{{OCASIAO_PARABOLA}}"]},
            {"titulo": "3 INTRODUÇÃO", "paragrafos": ["{{INTRODUCAO}}"]},
            {"titulo": "4 INTERPRETAÇÃO DO PONTO CENTRAL E CONTRASTES", "paragrafos": ["{{DESENVOLVIMENTO}}"]},
            {"titulo": "5 APLICAÇÃO AO REINO DE DEUS HOJE", "paragrafos": ["{{APLICACAO}}"]},
            {"titulo": "6 CONCLUSÃO", "paragrafos": ["{{CONCLUSAO}}"]},
        ],
    }
    return dados.get("secoes", especificos[genero])


def construir_declaracao_atual(doc: docx.Document, dados: dict[str, Any]) -> None:
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    professor = dados.get("professor", "{{PROFESSOR}}")
    livro = _titulo_completo(dados)
    referencia = dados.get("referencia_abnt", "{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}")

    natureza = (
        f"Declaração de leitura apresentada para aprovação na disciplina {disciplina} "
        f"do Seminário Cristão Evangélico do Norte, no {curso}."
    )
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_p(
        doc,
        f"Declaro para fins de aprovação na disciplina {disciplina} do {curso} do "
        f"Seminário Cristão Evangélico do Norte, ministrada pelo/a professor/a {professor}, "
        f"que li a obra {livro}. Minha rotina de leitura seguiu o fluxo registrado no cronograma abaixo:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        space_after=12,
    )

    tabela = doc.add_table(rows=2, cols=4)
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Linha 0 mesclada com a referência ABNT e fundo #F1F3F4 (igual ao Modelo de Declaração Atual.docx)
    merged = tabela.rows[0].cells[0].merge(tabela.rows[0].cells[3])
    _format_cell(merged, referencia, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, fill_hex="F1F3F4")
    # Linha 1 cabeçalhos
    for idx, texto in enumerate(["Seção", "Título", "Datas de início e de conclusão das partes da obra", "Horas empreendidas na parte da obra"]):
        _format_cell(tabela.rows[1].cells[idx], texto, bold=True)
    for item in _cronograma_padrao(dados):
        row = tabela.add_row()
        _format_cell(row.cells[0], item.get("secao", ""))
        _format_cell(row.cells[1], item.get("titulo", ""))
        _format_cell(row.cells[2], item.get("datas", ""))
        _format_cell(row.cells[3], item.get("horas", ""))


def construir_declaracao_antiga(doc: docx.Document, dados: dict[str, Any]) -> None:
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    aluno = dados.get("aluno", "_________________________")
    livro = _titulo_completo(dados)
    cidade = dados.get("cidade", CIDADE_PADRAO)
    data_assinatura = dados.get("data_assinatura", "___ / ___ / _____")

    natureza = (
        f"Declaração de leitura apresentada para aprovação na disciplina de {disciplina} "
        f"do Seminário Cristão Evangélico do Norte, no {curso}."
    )
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_p(doc, "DECLARAÇÃO DE LEITURA", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, space_after=18)
    _add_p(
        doc,
        f'Eu, {aluno}, declaro, para fins de crédito, que li o livro: "{livro}" '
        f"como leitura exigida para obtenção de nota da disciplina {disciplina} do {curso} do SCEN.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=1.5,
        space_after=18,
    )
    _add_p(
        doc,
        "9 Não mintais uns aos outros, pois que já vos despistes do velho homem com os seus feitos "
        "10 e vos vestistes do novo, que se renova para o conhecimento, segundo a imagem daquele que o criou; "
        "(Colossenses 3.9-10).",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        italic=True,
        space_after=24,
    )
    _add_p(doc, f"{cidade}, {data_assinatura}.", align=WD_ALIGN_PARAGRAPH.RIGHT)


def construir_relatorio_leitura(doc: docx.Document, dados: dict[str, Any]) -> None:
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    referencia = dados.get("referencia_abnt", "{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}")

    natureza = (
        f"Relatório de leitura apresentado ao Seminário Cristão Evangélico do Norte como requisito "
        f"para obtenção de nota na disciplina {disciplina} do {curso}."
    )
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_p(doc, referencia, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.0, space_after=12)

    perguntas = [
        ("1. Qual o tema da obra?", dados.get("q1_tema", "{{RESPOSTA_1_TEMA_DA_OBRA}}")),
        ("2. Qual a ideia central e o propósito da obra?", dados.get("q2_ideia_central", "{{RESPOSTA_2_IDEIA_CENTRAL_E_PROPOSITO}}")),
        ("3. Quais os principais argumentos apresentados pelo autor?", dados.get("q3_argumentos", "{{RESPOSTA_3_PRINCIPAIS_ARGUMENTOS}}")),
        ("4. Quais os pontos de problematização e elementos contrastantes presentes da obra?", dados.get("q4_problematizacao", "{{RESPOSTA_4_PROBLEMATIZACAO_E_CONTRASTES}}")),
        ("5. Faça uma síntese do livro.", dados.get("q5_sintese", "{{RESPOSTA_5_SINTESE_DO_LIVRO}}")),
        ("6. Apresente sua avaliação pessoal sobre a obra e diga de que forma este conteúdo poderá ser útil em seu ministério.", dados.get("q6_avaliacao_ministerio", "{{RESPOSTA_6_AVALIACAO_E_UTILIDADE_MINISTERIAL}}")),
        ("7. Escreva a conclusão da leitura.", dados.get("q7_conclusao", "{{RESPOSTA_7_CONCLUSAO}}")),
    ]
    for pergunta, resposta in perguntas:
        _add_p(doc, pergunta, align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, space_before=6, space_after=4)
        _add_p(doc, resposta, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=1.5, space_after=8)


def construir_paper(doc: docx.Document, dados: dict[str, Any]) -> None:
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    sinopse = dados.get("sinopse", "{{TEXTO_DA_SINOPSE}}")
    secoes = dados.get(
        "secoes",
        [
            {"titulo": "1 INTRODUÇÃO", "paragrafos": ["{{TEXTO_INTRODUCAO}}"]},
            {"titulo": "2 DESENVOLVIMENTO TEÓRICO", "paragrafos": ["{{TEXTO_DESENVOLVIMENTO}}"]},
            {"titulo": "3 CONCLUSÃO", "paragrafos": ["{{TEXTO_CONCLUSAO}}"]},
        ],
    )
    refs = dados.get("referencias", ["{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}"])

    natureza = f"Paper apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_p(doc, f"SINOPSE: {sinopse}", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=12)
    _add_secoes(doc, secoes)
    _add_referencias(doc, refs)


def construir_exegese_nt(doc: docx.Document, dados: dict[str, Any]) -> None:
    disciplina = dados.get("disciplina", "Exegese do Novo Testamento")
    curso = dados.get("curso", "{{CURSO}}")
    passagem = dados.get("passagem", "{{LIVRO CAPITULO:VERSICULOS}}")
    secoes = dados.get(
        "secoes",
        [
            {"titulo": f"PASSO 1 — DELIMITAÇÃO DA PERÍCOPE E CONTEXTO HISTÓRICO-LITERÁRIO ({passagem})", "paragrafos": ["{{PASSO_1_CONTEUDO}}"]},
            {"titulo": "PASSO 2 — CRÍTICA TEXTUAL E TRADUÇÃO LITERAL", "paragrafos": ["{{PASSO_2_CONTEUDO}}"]},
            {"titulo": "PASSO 3 — ANÁLISE GRAMATICAL, SINTÁTICA E LÉXICA", "paragrafos": ["{{PASSO_3_CONTEUDO}}"]},
            {"titulo": "PASSO 4 — ANÁLISE ESTRUTURAL E LITERÁRIA", "paragrafos": ["{{PASSO_4_CONTEUDO}}"]},
            {"titulo": "PASSO 5 — CONTEXTO TEOLÓGICO E BÍBLICO", "paragrafos": ["{{PASSO_5_CONTEUDO}}"]},
            {"titulo": "PASSO 6 — IDEIA CENTRAL EXEGÉTICA (SÍNTESE)", "paragrafos": ["{{PASSO_6_CONTEUDO}}"]},
            {"titulo": "PASSO 7 — APLICAÇÃO CONTEMPORÂNEA E HOMILÉTICA", "paragrafos": ["{{PASSO_7_CONTEUDO}}"]},
        ],
    )
    refs = dados.get("referencias", ["{{SOBRENOME, Nome. Comentário exegético. Cidade: Editora, Ano.}}"])

    natureza = f"Trabalho exegético apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_p(doc, f"Texto Bíblico Base: {passagem}", align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, space_after=12)
    _add_secoes(doc, secoes)
    _add_referencias(doc, refs)


def construir_sermao(doc: docx.Document, dados: dict[str, Any], genero: str) -> None:
    disciplina = dados.get("disciplina", "Prática da Pregação")
    curso = dados.get("curso", "{{CURSO}}")
    natureza = f"Esboço de sermão em {genero.lower()} apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_secoes(doc, _secoes_sermao(dados, genero))


def construir_fichamento(doc: docx.Document, dados: dict[str, Any]) -> None:
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    tipo_fichamento = dados.get("tipo_fichamento", "Resumo / Transcrição / Comentário")
    referencia = dados.get("referencia_abnt", "{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}")
    fichas = dados.get(
        "fichas",
        [
            {"cabecalho": "Ficha 1 — {{TEMA_OU_CAPITULO}}", "conteudo": "{{CONTEUDO_DA_FICHA_1}}"},
            {"cabecalho": "Ficha 2 — {{TEMA_OU_CAPITULO}}", "conteudo": "{{CONTEUDO_DA_FICHA_2}}"},
        ],
    )
    natureza = f"Fichamento de {tipo_fichamento} apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_p(doc, referencia, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.0, space_after=12)
    for ficha in fichas:
        _add_p(doc, ficha["cabecalho"], align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, space_before=6, space_after=4)
        _add_p(doc, ficha["conteudo"], align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=1.5, space_after=8)


def construir_resumo(doc: docx.Document, dados: dict[str, Any]) -> None:
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    referencia = dados.get("referencia_abnt", "{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}")
    texto = dados.get("texto_resumo", "{{TEXTO_DO_RESUMO_EM_PARAGRAFO_UNICO_150_A_500_PALAVRAS}}")
    palavras_chave = dados.get("palavras_chave", "{{palavra1; palavra2; palavra3}}")

    natureza = f"Resumo apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_p(doc, referencia, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.0, space_after=12)
    _add_p(doc, texto, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=12)
    _add_p(doc, f"Palavras-chave: {palavras_chave}.", align=WD_ALIGN_PARAGRAPH.LEFT)


def construir_resenha(doc: docx.Document, dados: dict[str, Any]) -> None:
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    referencia = dados.get("referencia_abnt", "{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}")
    paragrafos = dados.get(
        "paragrafos",
        [
            "{{PARAGRAFO_1_APRESENTACAO_DA_OBRA_E_AUTOR}}",
            "{{PARAGRAFO_2_SINTESE_DAS_IDEIAS_PRINCIPAIS}}",
            "{{PARAGRAFO_3_ANALISE_CRITICA_E_RECOMENDACAO}}",
        ],
    )
    natureza = f"Resenha apresentada para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_p(doc, referencia, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.0, space_after=12)
    for par in paragrafos:
        _add_p(doc, par, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=1.5, space_after=6)


def construir_artigo(doc: docx.Document, dados: dict[str, Any]) -> None:
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    resumo = dados.get("resumo", "{{RESUMO_DO_ARTIGO}}")
    palavras_chave = dados.get("palavras_chave", "{{palavra1; palavra2; palavra3}}")
    secoes = dados.get(
        "secoes",
        [
            {"titulo": "1 INTRODUÇÃO", "paragrafos": ["{{TEXTO_INTRODUCAO}}"]},
            {"titulo": "2 REFERENCIAL TEÓRICO", "paragrafos": ["{{TEXTO_REFERENCIAL}}"]},
            {"titulo": "3 MÉTODO", "paragrafos": ["{{TEXTO_METODO}}"]},
            {"titulo": "4 RESULTADOS E DISCUSSÃO", "paragrafos": ["{{TEXTO_RESULTADOS}}"]},
            {"titulo": "5 CONSIDERAÇÕES FINAIS", "paragrafos": ["{{TEXTO_CONSIDERACOES_FINAIS}}"]},
        ],
    )
    refs = dados.get("referencias", ["{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}"])

    natureza = f"Artigo científico apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
    _add_capa_e_folha_rosto(doc, dados, natureza)
    _add_p(doc, "RESUMO", align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, space_after=4)
    _add_p(doc, resumo, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6)
    _add_p(doc, f"Palavras-chave: {palavras_chave}.", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=12)
    _add_secoes(doc, secoes)
    _add_referencias(doc, refs)


# ==============================================================================
# GERADOR DE PDF (HTML + WEASYPRINT NO PADRÃO SCEN 2025)
# ==============================================================================


def _esc(val: Any) -> str:
    return html.escape(str(val))


def _render_capa_e_folha_html(dados: dict[str, Any], natureza: str) -> str:
    inst = _esc(dados.get("instituicao", INSTITUICAO_PADRAO).upper())
    aluno = _esc(dados.get("aluno", "{{NOME_DO_ALUNO}}"))
    titulo = _esc(_titulo_completo(dados))
    cidade = _esc(dados.get("cidade", CIDADE_PADRAO))
    ano = _esc(dados.get("ano", "{{ANO}}"))
    professor = _esc(dados.get("professor", "{{NOME_DO_PROFESSOR}}"))
    return f"""
    <section class="capa page-break">
      <div class="center bold">{inst}</div>
      <div class="center mt-2">{aluno}</div>
      <div class="center bold titulo-meio">{titulo}</div>
      <div class="center rodape">{cidade}<br/>{ano}</div>
    </section>
    <section class="folha-rosto page-break">
      <div class="center">{aluno}</div>
      <div class="center bold titulo-folha">{titulo}</div>
      <div class="natureza">
        <p>{_esc(natureza)}</p>
        <p>Professor(a): {professor}</p>
      </div>
      <div class="center rodape">{cidade}<br/>{ano}</div>
    </section>
    """


def _render_secoes_html(secoes: list[dict[str, Any]]) -> str:
    partes: list[str] = []
    for sec in secoes:
        partes.append(f'<h2 class="sec-titulo">{_esc(sec["titulo"])}</h2>')
        for p in sec.get("paragrafos", [sec.get("conteudo", "")]):
            if p:
                partes.append(f'<p class="recuo">{_esc(p)}</p>')
    return "\n".join(partes)


def _render_refs_html(refs: list[str]) -> str:
    itens = "\n".join(f'<p class="ref">{_esc(r)}</p>' for r in refs)
    return f"""
    <section class="page-break-before">
      <h2 class="center bold">REFERÊNCIAS</h2>
      <div class="mt-2">{itens}</div>
    </section>
    """


def _construir_html_pdf(tipo: str, dados: dict[str, Any]) -> str:
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    professor = dados.get("professor", "{{PROFESSOR}}")
    livro = _titulo_completo(dados)
    referencia = dados.get("referencia_abnt", "{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}")

    if tipo == "declaracao_atual":
        nat = f"Declaração de leitura apresentada para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
        linhas = "\n".join(
            f"<tr><td>{_esc(i.get('secao',''))}</td><td>{_esc(i.get('titulo',''))}</td>"
            f"<td>{_esc(i.get('datas',''))}</td><td>{_esc(i.get('horas',''))}</td></tr>"
            for i in _cronograma_padrao(dados)
        )
        corpo = f"""
        <p>Declaro para fins de aprovação na disciplina {_esc(disciplina)} do {_esc(curso)} do Seminário Cristão Evangélico do Norte, ministrada pelo/a professor/a {_esc(professor)}, que li a obra {_esc(livro)}. Minha rotina de leitura seguiu o fluxo registrado no cronograma abaixo:</p>
        <table>
          <thead>
            <tr><th colspan="4" class="ref-row">{_esc(referencia)}</th></tr>
            <tr><th>Seção</th><th>Título</th><th>Datas de início e de conclusão das partes da obra</th><th>Horas empreendidas na parte da obra</th></tr>
          </thead>
          <tbody>{linhas}</tbody>
        </table>
        """
    elif tipo == "declaracao_antiga":
        nat = f"Declaração de leitura apresentada para aprovação na disciplina de {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
        aluno = dados.get("aluno", "_________________________")
        cidade = dados.get("cidade", CIDADE_PADRAO)
        data_ass = dados.get("data_assinatura", "___ / ___ / _____")
        corpo = f"""
        <h2 class="center bold">DECLARAÇÃO DE LEITURA</h2>
        <p class="recuo mt-2">Eu, {_esc(aluno)}, declaro, para fins de crédito, que li o livro: "{_esc(livro)}" como leitura exigida para obtenção de nota da disciplina {_esc(disciplina)} do {_esc(curso)} do SCEN.</p>
        <p class="italic mt-2">9 Não mintais uns aos outros, pois que já vos despistes do velho homem com os seus feitos 10 e vos vestistes do novo, que se renova para o conhecimento, segundo a imagem daquele que o criou; (Colossenses 3.9-10).</p>
        <p class="right mt-3">{_esc(cidade)}, {_esc(data_ass)}.</p>
        """
    elif tipo == "relatorio_leitura":
        nat = f"Relatório de leitura apresentado ao Seminário Cristão Evangélico do Norte como requisito para obtenção de nota na disciplina {disciplina} do {curso}."
        perguntas = [
            ("1. Qual o tema da obra?", dados.get("q1_tema", "{{RESPOSTA_1_TEMA_DA_OBRA}}")),
            ("2. Qual a ideia central e o propósito da obra?", dados.get("q2_ideia_central", "{{RESPOSTA_2_IDEIA_CENTRAL_E_PROPOSITO}}")),
            ("3. Quais os principais argumentos apresentados pelo autor?", dados.get("q3_argumentos", "{{RESPOSTA_3_PRINCIPAIS_ARGUMENTOS}}")),
            ("4. Quais os pontos de problematização e elementos contrastantes presentes da obra?", dados.get("q4_problematizacao", "{{RESPOSTA_4_PROBLEMATIZACAO_E_CONTRASTES}}")),
            ("5. Faça uma síntese do livro.", dados.get("q5_sintese", "{{RESPOSTA_5_SINTESE_DO_LIVRO}}")),
            ("6. Apresente sua avaliação pessoal sobre a obra e diga de que forma este conteúdo poderá ser útil em seu ministério.", dados.get("q6_avaliacao_ministerio", "{{RESPOSTA_6_AVALIACAO_E_UTILIDADE_MINISTERIAL}}")),
            ("7. Escreva a conclusão da leitura.", dados.get("q7_conclusao", "{{RESPOSTA_7_CONCLUSAO}}")),
        ]
        qs = "\n".join(f'<h2 class="sec-titulo">{_esc(q)}</h2><p class="recuo">{_esc(a)}</p>' for q, a in perguntas)
        corpo = f'<p class="ref">{_esc(referencia)}</p>{qs}'
    elif tipo == "paper":
        nat = f"Paper apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
        secoes = dados.get("secoes", [
            {"titulo": "1 INTRODUÇÃO", "paragrafos": ["{{TEXTO_INTRODUCAO}}"]},
            {"titulo": "2 DESENVOLVIMENTO TEÓRICO", "paragrafos": ["{{TEXTO_DESENVOLVIMENTO}}"]},
            {"titulo": "3 CONCLUSÃO", "paragrafos": ["{{TEXTO_CONCLUSAO}}"]},
        ])
        refs = dados.get("referencias", ["{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}"])
        corpo = f'<p><strong>SINOPSE:</strong> {_esc(dados.get("sinopse", "{{TEXTO_DA_SINOPSE}}"))}</p>{_render_secoes_html(secoes)}{_render_refs_html(refs)}'
    elif tipo == "exegese_nt":
        nat = f"Trabalho exegético apresentado para aprovação na disciplina {dados.get('disciplina', 'Exegese do Novo Testamento')} do Seminário Cristão Evangélico do Norte, no {curso}."
        passagem = dados.get("passagem", "{{LIVRO CAPITULO:VERSICULOS}}")
        secoes = dados.get("secoes", [
            {"titulo": f"PASSO 1 — DELIMITAÇÃO DA PERÍCOPE E CONTEXTO HISTÓRICO-LITERÁRIO ({passagem})", "paragrafos": ["{{PASSO_1_CONTEUDO}}"]},
            {"titulo": "PASSO 2 — CRÍTICA TEXTUAL E TRADUÇÃO LITERAL", "paragrafos": ["{{PASSO_2_CONTEUDO}}"]},
            {"titulo": "PASSO 3 — ANÁLISE GRAMATICAL, SINTÁTICA E LÉXICA", "paragrafos": ["{{PASSO_3_CONTEUDO}}"]},
            {"titulo": "PASSO 4 — ANÁLISE ESTRUTURAL E LITERÁRIA", "paragrafos": ["{{PASSO_4_CONTEUDO}}"]},
            {"titulo": "PASSO 5 — CONTEXTO TEOLÓGICO E BÍBLICO", "paragrafos": ["{{PASSO_5_CONTEUDO}}"]},
            {"titulo": "PASSO 6 — IDEIA CENTRAL EXEGÉTICA (SÍNTESE)", "paragrafos": ["{{PASSO_6_CONTEUDO}}"]},
            {"titulo": "PASSO 7 — APLICAÇÃO CONTEMPORÂNEA E HOMILÉTICA", "paragrafos": ["{{PASSO_7_CONTEUDO}}"]},
        ])
        refs = dados.get("referencias", ["{{SOBRENOME, Nome. Comentário exegético. Cidade: Editora, Ano.}}"])
        corpo = f'<p class="bold">Texto Bíblico Base: {_esc(passagem)}</p>{_render_secoes_html(secoes)}{_render_refs_html(refs)}'
    elif tipo.startswith("sermao_"):
        genero_map = {"sermao_narrativa": "Narrativa", "sermao_evangelho": "Evangelho", "sermao_parabola": "Parábola"}
        genero = genero_map[tipo]
        nat = f"Esboço de sermão em {genero.lower()} apresentado para aprovação na disciplina {dados.get('disciplina', 'Prática da Pregação')} do Seminário Cristão Evangélico do Norte, no {curso}."
        corpo = _render_secoes_html(_secoes_sermao(dados, genero))
    elif tipo == "fichamento":
        nat = f"Fichamento de {dados.get('tipo_fichamento', 'Resumo / Transcrição / Comentário')} apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
        fichas = dados.get("fichas", [
            {"cabecalho": "Ficha 1 — {{TEMA_OU_CAPITULO}}", "conteudo": "{{CONTEUDO_DA_FICHA_1}}"},
            {"cabecalho": "Ficha 2 — {{TEMA_OU_CAPITULO}}", "conteudo": "{{CONTEUDO_DA_FICHA_2}}"},
        ])
        fs = "\n".join(f'<h2 class="sec-titulo">{_esc(f["cabecalho"])}</h2><p class="recuo">{_esc(f["conteudo"])}</p>' for f in fichas)
        corpo = f'<p class="ref">{_esc(referencia)}</p>{fs}'
    elif tipo == "resumo":
        nat = f"Resumo apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
        corpo = f'<p class="ref">{_esc(referencia)}</p><p>{_esc(dados.get("texto_resumo", "{{TEXTO_DO_RESUMO_EM_PARAGRAFO_UNICO_150_A_500_PALAVRAS}}"))}</p><p>Palavras-chave: {_esc(dados.get("palavras_chave", "{{palavra1; palavra2; palavra3}}"))}.</p>'
    elif tipo == "resenha":
        nat = f"Resenha apresentada para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
        pars = "\n".join(f'<p class="recuo">{_esc(p)}</p>' for p in dados.get("paragrafos", [
            "{{PARAGRAFO_1_APRESENTACAO_DA_OBRA_E_AUTOR}}",
            "{{PARAGRAFO_2_SINTESE_DAS_IDEIAS_PRINCIPAIS}}",
            "{{PARAGRAFO_3_ANALISE_CRITICA_E_RECOMENDACAO}}",
        ]))
        corpo = f'<p class="ref">{_esc(referencia)}</p>{pars}'
    elif tipo == "artigo":
        nat = f"Artigo científico apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."
        secoes = dados.get("secoes", [
            {"titulo": "1 INTRODUÇÃO", "paragrafos": ["{{TEXTO_INTRODUCAO}}"]},
            {"titulo": "2 REFERENCIAL TEÓRICO", "paragrafos": ["{{TEXTO_REFERENCIAL}}"]},
            {"titulo": "3 MÉTODO", "paragrafos": ["{{TEXTO_METODO}}"]},
            {"titulo": "4 RESULTADOS E DISCUSSÃO", "paragrafos": ["{{TEXTO_RESULTADOS}}"]},
            {"titulo": "5 CONSIDERAÇÕES FINAIS", "paragrafos": ["{{TEXTO_CONSIDERACOES_FINAIS}}"]},
        ])
        refs = dados.get("referencias", ["{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}"])
        corpo = f'<h2 class="sec-titulo">RESUMO</h2><p>{_esc(dados.get("resumo", "{{RESUMO_DO_ARTIGO}}"))}</p><p>Palavras-chave: {_esc(dados.get("palavras_chave", "{{palavra1; palavra2; palavra3}}"))}.</p>{_render_secoes_html(secoes)}{_render_refs_html(refs)}'
    else:
        raise ValueError(f"Tipo inválido: {tipo}")

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<style>
  @page {{
    size: A4;
    margin: 2.54cm 2.54cm 2.54cm 2.54cm;
  }}
  body {{
    font-family: "Times New Roman", "Liberation Serif", Times, serif;
    font-size: 12pt;
    line-height: 1.5;
    color: #000;
    margin: 0;
  }}
  .page-break {{ page-break-after: always; }}
  .page-break-before {{ page-break-before: always; }}
  .center {{ text-align: center; }}
  .right {{ text-align: right; }}
  .bold {{ font-weight: bold; }}
  .italic {{ font-style: italic; }}
  .mt-2 {{ margin-top: 24pt; }}
  .mt-3 {{ margin-top: 36pt; }}
  .capa, .folha-rosto {{
    position: relative;
    height: 24.4cm;
  }}
  .titulo-meio {{ margin-top: 7.5cm; }}
  .titulo-folha {{ margin-top: 8cm; }}
  .rodape {{
    position: absolute;
    bottom: 0;
    left: 0;
    right: 0;
  }}
  .natureza {{
    margin-top: 2cm;
    margin-left: 8.0cm;
    font-size: 10pt;
    line-height: 1.15;
    text-align: justify;
  }}
  .natureza p {{ margin: 0 0 6pt 0; }}
  p {{ margin: 0 0 8pt 0; text-align: justify; }}
  p.recuo {{ text-indent: 1.5cm; }}
  p.ref {{ text-align: left; line-height: 1.15; margin-bottom: 12pt; }}
  h2.sec-titulo {{ font-size: 12pt; font-weight: bold; text-align: left; margin: 12pt 0 6pt 0; }}
  table {{ width: 100%; border-collapse: collapse; margin-top: 12pt; }}
  th, td {{ border: 1px solid #000; padding: 6pt; text-align: center; vertical-align: middle; font-size: 11pt; }}
  th {{ font-weight: bold; }}
  th.ref-row {{ background-color: #f1f3f4; font-weight: normal; text-align: left; }}
</style>
</head>
<body>
{_render_capa_e_folha_html(dados, nat)}
<main>
{corpo}
</main>
</body>
</html>"""


# ==============================================================================
# REGISTRO DOS 12 MODELOS + DADOS ESSENCIAIS
# ==============================================================================


CONSTRUTORES = {
    "declaracao_atual": ("01_Declaracao_Leitura_Atual", construir_declaracao_atual),
    "declaracao_antiga": ("02_Declaracao_Leitura_Antiga", construir_declaracao_antiga),
    "relatorio_leitura": ("03_Relatorio_de_Leitura", construir_relatorio_leitura),
    "paper": ("04_Paper", construir_paper),
    "exegese_nt": ("05_Exegese_Novo_Testamento_7_Passos", construir_exegese_nt),
    "sermao_narrativa": ("06_Sermao_Narrativa", lambda d, p: construir_sermao(d, p, "Narrativa")),
    "sermao_evangelho": ("07_Sermao_Evangelho", lambda d, p: construir_sermao(d, p, "Evangelho")),
    "sermao_parabola": ("08_Sermao_Parabola", lambda d, p: construir_sermao(d, p, "Parábola")),
    "fichamento": ("09_Fichamento", construir_fichamento),
    "resumo": ("10_Resumo", construir_resumo),
    "resenha": ("11_Resenha_Recensao", construir_resenha),
    "artigo": ("12_Artigo_Cientifico", construir_artigo),
}


DADOS_ESSENCIAIS_TEMPLATE: dict[str, dict[str, Any]] = {
    "declaracao_atual": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "", "professor": "",
        "titulo": "", "subtitulo": "", "referencia_abnt": "",
        "cronograma": [{"secao": "Capítulo 1", "titulo": "", "datas": "", "horas": ""}],
    },
    "declaracao_antiga": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "", "professor": "",
        "titulo": "", "subtitulo": "", "data_assinatura": "___ / ___ / 2026",
    },
    "relatorio_leitura": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "", "professor": "",
        "titulo": "", "subtitulo": "", "referencia_abnt": "",
        "q1_tema": "", "q2_ideia_central": "", "q3_argumentos": "",
        "q4_problematizacao": "", "q5_sintese": "", "q6_avaliacao_ministerio": "", "q7_conclusao": "",
    },
    "paper": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "", "professor": "",
        "titulo": "", "subtitulo": "", "sinopse": "",
        "secoes": [{"titulo": "1 INTRODUÇÃO", "paragrafos": [""]}, {"titulo": "2 DESENVOLVIMENTO", "paragrafos": [""]}, {"titulo": "3 CONCLUSÃO", "paragrafos": [""]}],
        "referencias": [""],
    },
    "exegese_nt": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "Exegese do Novo Testamento", "professor": "",
        "titulo": "", "passagem": "",
        "secoes": [{"titulo": f"PASSO {i}", "paragrafos": [""]} for i in range(1, 8)],
        "referencias": [""],
    },
    "sermao_narrativa": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "Prática da Pregação", "professor": "",
        "titulo": "", "passagem": "", "secoes": [{"titulo": "1 DADOS GERAIS", "paragrafos": [""]}],
    },
    "sermao_evangelho": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "Prática da Pregação", "professor": "",
        "titulo": "", "passagem": "", "secoes": [{"titulo": "1 DADOS GERAIS", "paragrafos": [""]}],
    },
    "sermao_parabola": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "Prática da Pregação", "professor": "",
        "titulo": "", "passagem": "", "secoes": [{"titulo": "1 DADOS GERAIS", "paragrafos": [""]}],
    },
    "fichamento": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "", "professor": "",
        "titulo": "", "tipo_fichamento": "Resumo", "referencia_abnt": "",
        "fichas": [{"cabecalho": "Ficha 1", "conteudo": ""}],
    },
    "resumo": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "", "professor": "",
        "titulo": "", "referencia_abnt": "", "texto_resumo": "", "palavras_chave": "",
    },
    "resenha": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "", "professor": "",
        "titulo": "", "referencia_abnt": "", "paragrafos": ["", "", ""],
    },
    "artigo": {
        "aluno": "", "curso": "", "ano": "2026", "disciplina": "", "professor": "",
        "titulo": "", "resumo": "", "palavras_chave": "",
        "secoes": [{"titulo": "1 INTRODUÇÃO", "paragrafos": [""]}], "referencias": [""],
    },
}


def _gerar_pdf_reportlab(tipo: str, dados: dict[str, Any], caminho_pdf: Path) -> None:
    """Fallback multiplataforma (Windows/Linux/macOS) usando ReportLab em A4."""
    doc_pdf = SimpleDocTemplate(
        str(caminho_pdf),
        pagesize=A4,
        topMargin=2.54 * cm,
        bottomMargin=2.54 * cm,
        leftMargin=2.54 * cm,
        rightMargin=2.54 * cm,
    )
    st_center = ParagraphStyle("center", fontName="Times-Roman", fontSize=12, leading=18, alignment=TA_CENTER)
    st_center_b = ParagraphStyle("center_b", parent=st_center, fontName="Times-Bold")
    st_body = ParagraphStyle("body", fontName="Times-Roman", fontSize=12, leading=18, alignment=TA_JUSTIFY, firstLineIndent=1.5 * cm, spaceAfter=6)
    st_left = ParagraphStyle("left", fontName="Times-Roman", fontSize=12, leading=16, alignment=TA_LEFT, spaceAfter=6)
    st_sec = ParagraphStyle("sec", fontName="Times-Bold", fontSize=12, leading=18, alignment=TA_LEFT, spaceBefore=8, spaceAfter=4)
    st_nat = ParagraphStyle("nat", fontName="Times-Roman", fontSize=10, leading=13, alignment=TA_JUSTIFY, leftIndent=8.0 * cm, spaceAfter=4)
    st_right = ParagraphStyle("right", fontName="Times-Roman", fontSize=12, leading=18, alignment=TA_RIGHT, spaceBefore=12)

    inst = _esc(dados.get("instituicao", INSTITUICAO_PADRAO).upper())
    aluno = _esc(dados.get("aluno", "{{NOME_DO_ALUNO}}"))
    titulo = _esc(_titulo_completo(dados))
    cidade = _esc(dados.get("cidade", CIDADE_PADRAO))
    ano = _esc(dados.get("ano", "{{ANO}}"))
    professor = _esc(dados.get("professor", "{{NOME_DO_PROFESSOR}}"))
    disciplina = dados.get("disciplina", "{{DISCIPLINA}}")
    curso = dados.get("curso", "{{CURSO}}")
    referencia = _esc(dados.get("referencia_abnt", "{{SOBRENOME, Nome. Título da obra. Cidade: Editora, Ano.}}"))

    story: list[Any] = [
        Paragraph(inst, st_center_b),
        Spacer(1, 1.2 * cm),
        Paragraph(aluno, st_center),
        Spacer(1, 6.5 * cm),
        Paragraph(titulo, st_center_b),
        Spacer(1, 8.5 * cm),
        Paragraph(f"{cidade}<br/>{ano}", st_center),
        PageBreak(),
        Paragraph(aluno, st_center),
        Spacer(1, 7.5 * cm),
        Paragraph(titulo, st_center_b),
        Spacer(1, 1.5 * cm),
        Paragraph(_esc(f"Trabalho apresentado para aprovação na disciplina {disciplina} do Seminário Cristão Evangélico do Norte, no {curso}."), st_nat),
        Paragraph(f"Professor(a): {professor}", st_nat),
        Spacer(1, 6.5 * cm),
        Paragraph(f"{cidade}<br/>{ano}", st_center),
        PageBreak(),
    ]

    if tipo == "declaracao_atual":
        story.append(Paragraph(_esc(f"Declaro para fins de aprovação na disciplina {disciplina} do {curso} do Seminário Cristão Evangélico do Norte, ministrada pelo/a professor/a {dados.get('professor', '{{PROFESSOR}}')}, que li a obra {_titulo_completo(dados)}. Minha rotina de leitura seguiu o fluxo registrado no cronograma abaixo:"), st_body))
        story.append(Paragraph(referencia, st_left))
        rows = [["Seção", "Título", "Datas", "Horas"]]
        for i in _cronograma_padrao(dados):
            rows.append([i.get("secao", ""), i.get("titulo", ""), i.get("datas", ""), i.get("horas", "")])
        t = Table(rows, colWidths=[3 * cm, 5 * cm, 4.5 * cm, 3.5 * cm])
        t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
        story.append(t)
    elif tipo == "declaracao_antiga":
        story.append(Paragraph("DECLARAÇÃO DE LEITURA", st_center_b))
        story.append(Spacer(1, 0.8 * cm))
        story.append(Paragraph(_esc(f'Eu, {dados.get("aluno", "_________________________")}, declaro, para fins de crédito, que li o livro: "{_titulo_completo(dados)}" como leitura exigida para obtenção de nota da disciplina {disciplina} do {curso} do SCEN.'), st_body))
        story.append(Paragraph("9 Não mintais uns aos outros, pois que já vos despistes do velho homem com os seus feitos 10 e vos vestistes do novo, que se renova para o conhecimento, segundo a imagem daquele que o criou; (Colossenses 3.9-10).", st_left))
        story.append(Paragraph(_esc(f"{dados.get('cidade', CIDADE_PADRAO)}, {dados.get('data_assinatura', '___ / ___ / _____')}."), st_right))
    else:
        doc_tmp = _configurar_documento()
        _, construtor = CONSTRUTORES[tipo]
        construtor(doc_tmp, dados)
        for p in doc_tmp.paragraphs[12:]:
            if p.text.strip():
                is_bold = any(r.bold for r in p.runs)
                story.append(Paragraph(_esc(p.text), st_sec if is_bold else st_body))

    doc_pdf.build(story)


def gerar_documento(tipo: str, dados: dict[str, Any], destino_base: Path) -> tuple[Path, Path]:
    """Gera .docx e .pdf para o tipo especificado e retorna (caminho_docx, caminho_pdf)."""
    if tipo not in CONSTRUTORES:
        raise ValueError(f"Tipo inválido '{tipo}'. Opções: {', '.join(CONSTRUTORES)}")
    destino_base.parent.mkdir(parents=True, exist_ok=True)
    caminho_docx = destino_base.with_suffix(".docx")
    caminho_pdf = destino_base.with_suffix(".pdf")

    doc = _configurar_documento()
    _, construtor = CONSTRUTORES[tipo]
    construtor(doc, dados)
    doc.save(str(caminho_docx))

    if weasyprint is not None:
        try:
            html_str = _construir_html_pdf(tipo, dados)
            weasyprint.HTML(string=html_str).write_pdf(str(caminho_pdf))
            return caminho_docx, caminho_pdf
        except OSError:
            pass
    _gerar_pdf_reportlab(tipo, dados, caminho_pdf)
    return caminho_docx, caminho_pdf


def inicializar_tudo() -> None:
    """Popula Modelos/ (DOCX + PDF) e Informacoes_Essenciais/ (subpastas com DOCX, PDF e JSON)."""
    MODELOS_DIR.mkdir(parents=True, exist_ok=True)
    INFO_ESSENCIAIS_DIR.mkdir(parents=True, exist_ok=True)
    GERADOS_DIR.mkdir(parents=True, exist_ok=True)

    for tipo, (nome_base, _) in CONSTRUTORES.items():
        gerar_documento(tipo, {}, MODELOS_DIR / nome_base)

        subpasta = INFO_ESSENCIAIS_DIR / nome_base
        subpasta.mkdir(parents=True, exist_ok=True)
        gerar_documento(tipo, {}, subpasta / nome_base)
        (subpasta / "dados_essenciais.json").write_text(
            json.dumps(DADOS_ESSENCIAIS_TEMPLATE[tipo], ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Gerador de trabalhos acadêmicos SCEN (DOCX e PDF).")
    parser.add_argument("--init-templates", action="store_true", help="Cria modelos DOCX/PDF e pasta Informacoes_Essenciais/.")
    parser.add_argument("--tipo", choices=list(CONSTRUTORES.keys()), help="Tipo de trabalho a gerar.")
    parser.add_argument("--dados", type=Path, help="Arquivo JSON com os dados do trabalho.")
    parser.add_argument("--saida", type=Path, help="Caminho base (sem extensão) de saída.")
    args = parser.parse_args()

    if args.init_templates:
        inicializar_tudo()
        print(f"Modelos DOCX/PDF e pasta '{INFO_ESSENCIAIS_DIR.name}/' gerados com sucesso.")
        return

    if not args.tipo or not args.dados:
        parser.error("Informe --init-templates ou (--tipo e --dados).")

    dados = json.loads(args.dados.read_text(encoding="utf-8"))
    slug = dados.get("titulo", "trabalho").strip().replace(" ", "_") or "trabalho"
    saida_base = args.saida or (GERADOS_DIR / f"{args.tipo}_{slug}")
    docx_p, pdf_p = gerar_documento(args.tipo, dados, saida_base)
    print(f"Gerados:\n - {docx_p}\n - {pdf_p}")


if __name__ == "__main__":
    main()
