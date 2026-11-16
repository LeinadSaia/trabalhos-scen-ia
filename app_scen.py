#!/usr/bin/env python3
"""Interface gráfica PyQt6 para preenchimento e geração de trabalhos do SCEN com IA.

Compatível com Windows, Linux e macOS.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ai_assistant import (
    chamar_ia,
    extrair_json_resposta,
    extrair_texto_arquivo,
    listar_modelos_disponiveis,
    montar_prompt_ia,
)
from gerar_trabalho import BASE_DIR, CIDADE_PADRAO, GERADOS_DIR, gerar_documento


CONFIG_PATH = BASE_DIR / ".scen_config.json"
RASCUNHOS_DIR = BASE_DIR / "Rascunhos"

ATALHOS_DISCIPLINAS: list[tuple[str, str, dict[str, Any]]] = [
    ("Exegese do NT — Passos 1 a 7", "exegese_nt", {"disciplina": "Exegese do Novo Testamento"}),
    ("Teologia da Família — Paper", "paper", {"disciplina": "Teologia da Família"}),
    ("Teologia da Família — Família do Pastor (Declaração)", "declaracao_atual", {"disciplina": "Teologia da Família", "titulo": "Família do Pastor"}),
    ("Apologética — Por que você acredita? (Relatório)", "relatorio_leitura", {"disciplina": "Apologética", "titulo": "Por que você acredita?"}),
    ("Prática da Pregação — Sermão em Narrativa", "sermao_narrativa", {"disciplina": "Prática da Pregação"}),
    ("Prática da Pregação — Sermão em Evangelho", "sermao_evangelho", {"disciplina": "Prática da Pregação"}),
    ("Prática da Pregação — Sermão em Parábola", "sermao_parabola", {"disciplina": "Prática da Pregação"}),
    ("Soteriologia — Salvos pela Graça (Declaração)", "declaracao_antiga", {"disciplina": "Soteriologia", "titulo": "Salvos pela Graça", "subtitulo": "a doutrina bíblica da salvação"}),
]

MODELOS_LISTA: list[tuple[str, str]] = [
    ("01. Declaração de Leitura Atual (Cronograma)", "declaracao_atual"),
    ("02. Declaração de Leitura Antiga (Simples)", "declaracao_antiga"),
    ("03. Relatório de Leitura (7 Questões)", "relatorio_leitura"),
    ("04. Paper", "paper"),
    ("05. Exegese do Novo Testamento (7 Passos)", "exegese_nt"),
    ("06. Sermão em Narrativa", "sermao_narrativa"),
    ("07. Sermão em Evangelho", "sermao_evangelho"),
    ("08. Sermão em Parábola", "sermao_parabola"),
    ("09. Fichamento", "fichamento"),
    ("10. Resumo", "resumo"),
    ("11. Resenha / Recensão", "resenha"),
    ("12. Artigo Científico", "artigo"),
]

INSTRUCOES_POR_TIPO: dict[str, tuple[str, str]] = {
    "declaracao_atual": (
        "Declaração de Leitura Atual (com Cronograma)",
        "• O que preencher: Cabeçalho + Referência ABNT da obra + Tabela de Cronograma (Seção, Título, Datas e Horas).\n"
        "• Dica: Anexe o sumário/PDF do livro ou cole os capítulos no painel de IA para preencher a tabela automaticamente.",
    ),
    "declaracao_antiga": (
        "Declaração de Leitura Antiga (Simples — Colossenses 3.9-10)",
        "• O que preencher: Cabeçalho (com nome do aluno e obra) + Data da assinatura.\n"
        "• Dica: Ideal para disciplinas que pedem apenas a declaração simples assinada.",
    ),
    "relatorio_leitura": (
        "Relatório de Leitura (7 Questões Oficiais SCEN)",
        "• O que preencher: Cabeçalho + Referência ABNT + Respostas das 7 perguntas (tema, ideia central, argumentos, problematização, síntese, avaliação/ministério, conclusão).\n"
        "• Dica: Anexe o livro/resumo no painel de IA e clique em 'Completar com IA' para gerar as 7 respostas.",
    ),
    "paper": (
        "Paper Acadêmico (SCEN 2025)",
        "• O que preencher: Cabeçalho + Sinopse + Seções numeradas (use '## Título') + Referências ABNT.\n"
        "• Dica: Use os comandos rápidos no painel da direita para ajustar o tom teológico e ministerial.",
    ),
    "exegese_nt": (
        "Exegese do Novo Testamento (Passos 1 a 7)",
        "• O que preencher: Cabeçalho + Passagem Bíblica Base + Conteúdo dos 7 Passos exegéticos + Referências.\n"
        "• Dica: Informe a passagem (ex.: Rm 8:1-11) e anexe suas notas de aula para a IA estruturar os 7 passos.",
    ),
    "sermao_narrativa": (
        "Sermão em Narrativa Bíblica",
        "• O que preencher: Cabeçalho + Texto Base + Tema, Contexto Narrativo, Introdução, Desenvolvimento, Aplicação e Conclusão.\n"
        "• Dica: Cole o texto bíblico e esboço no painel de contexto para gerar a estrutura homilética.",
    ),
    "sermao_evangelho": (
        "Sermão em Evangelho",
        "• O que preencher: Cabeçalho + Texto Base + Tema, Contexto do Ministério de Jesus, Introdução, Exposição, Foco Cristológico e Conclusão.",
    ),
    "sermao_parabola": (
        "Sermão em Parábola",
        "• O que preencher: Cabeçalho + Texto Base + Ocasião da Parábola, Ponto Central, Contrastes, Aplicação ao Reino e Conclusão.",
    ),
    "fichamento": (
        "Fichamento (Resumo / Transcrição / Comentário)",
        "• O que preencher: Cabeçalho + Tipo de Fichamento + Referência ABNT + Fichas (use '## Título da Ficha').",
    ),
    "resumo": (
        "Resumo Acadêmico (150 a 500 palavras)",
        "• O que preencher: Cabeçalho + Referência ABNT + Texto em parágrafo único (150–500 palavras) + Palavras-chave.",
    ),
    "resenha": (
        "Resenha / Recensão Crítica",
        "• O que preencher: Cabeçalho + Referência ABNT + Parágrafos de apresentação, síntese e avaliação crítica.",
    ),
    "artigo": (
        "Artigo Científico (SCEN 2025)",
        "• O que preencher: Cabeçalho + Resumo + Palavras-chave + Seções (Introdução, Referencial, Método, Resultados, Considerações Finais) + Referências.",
    ),
}

TEMA_OPOSTOS = """
    QMainWindow { background-color: #0f172a; }
    QWidget#barraTopo, QWidget#rodape {
        background-color: #0f172a;
        color: #f8fafc;
    }
    QWidget#painelTrabalhos {
        background-color: #0f172a;
        color: #f8fafc;
    }
    QWidget#painelTrabalhos QGroupBox {
        background-color: #1e293b;
        color: #f8fafc;
        border: 1px solid #334155;
        border-radius: 8px;
        margin-top: 12px;
        padding-top: 14px;
        font-weight: bold;
    }
    QWidget#painelTrabalhos QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 6px;
        color: #38bdf8;
    }
    QWidget#painelTrabalhos QListWidget {
        background-color: #0f172a;
        color: #f8fafc;
        border: 1px solid #334155;
        border-radius: 6px;
        padding: 4px;
    }
    QWidget#painelTrabalhos QListWidget::item {
        padding: 7px 8px;
        border-radius: 4px;
    }
    QWidget#painelTrabalhos QListWidget::item:selected {
        background-color: #2563eb;
        color: #ffffff;
        font-weight: bold;
    }
    QWidget#painelTrabalhos QListWidget::item:hover {
        background-color: #1e293b;
    }
    QScrollArea, QWidget#conteudoForm, QWidget#painelIA {
        background-color: #f1f5f9;
        color: #0f172a;
    }
    QFrame#bannerInstrucoes {
        background-color: #0f172a;
        color: #f8fafc;
        border-left: 5px solid #38bdf8;
        border-radius: 8px;
        padding: 10px;
    }
    QFrame#bannerInstrucoes QLabel { color: #f8fafc; }
    QGroupBox {
        background-color: #ffffff;
        color: #0f172a;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        margin-top: 12px;
        padding-top: 14px;
        font-weight: bold;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 6px;
        color: #1e3a8a;
    }
    QLabel { color: #0f172a; font-weight: 500; }
    QWidget#barraTopo QLabel, QWidget#rodape QLabel { color: #f8fafc; }
    QLineEdit, QPlainTextEdit, QComboBox, QTableWidget {
        background-color: #ffffff;
        color: #0f172a;
        border: 2px solid #cbd5e1;
        border-radius: 6px;
        padding: 6px;
        selection-background-color: #2563eb;
    }
    QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus {
        border: 2px solid #2563eb;
    }
    QHeaderView::section {
        background-color: #0f172a;
        color: #f8fafc;
        padding: 6px;
        font-weight: bold;
        border: 1px solid #334155;
    }
    QListWidget {
        background-color: #ffffff;
        color: #0f172a;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
    }
    QCheckBox { color: #0f172a; font-weight: 600; }
    QTabWidget::pane { border: 1px solid #334155; background-color: #f1f5f9; }
    QTabBar::tab {
        background-color: #1e293b;
        color: #cbd5e1;
        padding: 9px 18px;
        margin-right: 4px;
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
        font-weight: bold;
    }
    QTabBar::tab:selected { background-color: #2563eb; color: #ffffff; }
    QPushButton {
        background-color: #e2e8f0;
        color: #0f172a;
        border: 1px solid #94a3b8;
        border-radius: 6px;
        padding: 8px 14px;
        font-weight: bold;
    }
    QPushButton:hover { background-color: #cbd5e1; }
    QPushButton:checked { background-color: #2563eb; color: #ffffff; border: 1px solid #38bdf8; }
    QPushButton#btnPrimario {
        background-color: #2563eb;
        color: #ffffff;
        border: 1px solid #1d4ed8;
    }
    QPushButton#btnPrimario:hover { background-color: #1d4ed8; }
    QPushButton#btnIA {
        background-color: #0d9488;
        color: #ffffff;
        border: 1px solid #0f766e;
    }
    QPushButton#btnIA:hover { background-color: #0f766e; }
    QPushButton#btnTopo {
        background-color: #1e293b;
        color: #f8fafc;
        border: 1px solid #475569;
    }
    QPushButton#btnTopo:hover { background-color: #334155; }
"""

TEMA_ESCURO_TOTAL = """
    QMainWindow, QWidget { background-color: #0b1120; color: #f8fafc; }
    QFrame#bannerInstrucoes {
        background-color: #1e293b;
        border-left: 5px solid #38bdf8;
        border-radius: 8px;
        padding: 10px;
    }
    QGroupBox {
        background-color: #111827;
        color: #f8fafc;
        border: 1px solid #334155;
        border-radius: 8px;
        margin-top: 12px;
        padding-top: 14px;
        font-weight: bold;
    }
    QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 6px; color: #38bdf8; }
    QLineEdit, QPlainTextEdit, QComboBox, QTableWidget, QListWidget {
        background-color: #1e293b;
        color: #f8fafc;
        border: 1px solid #475569;
        border-radius: 6px;
        padding: 6px;
    }
    QListWidget::item:selected { background-color: #2563eb; color: #ffffff; }
    QHeaderView::section { background-color: #0f172a; color: #38bdf8; padding: 6px; font-weight: bold; }
    QPushButton {
        background-color: #1e293b;
        color: #f8fafc;
        border: 1px solid #475569;
        border-radius: 6px;
        padding: 8px 14px;
        font-weight: bold;
    }
    QPushButton:hover { background-color: #334155; }
    QPushButton:checked { background-color: #2563eb; color: #ffffff; }
    QPushButton#btnPrimario { background-color: #2563eb; color: #ffffff; border: none; }
    QPushButton#btnIA { background-color: #0d9488; color: #ffffff; border: none; }
"""


MODELOS_POR_PROVEDOR: dict[str, list[str]] = {
    "Gemini": [
        "gemini-4-carbon",
        "gemini-4-argon",
        "gemini-2.5-pro",
        "gemini-3.5-flash",
        "gemini-2.5-flash",
    ],
    "OpenAI": [
        "gpt-4.1",
        "o3",
        "gpt-4o",
        "o4-mini",
        "gpt-4o-mini",
    ],
    "Ollama": [
        "llama3.3",
        "qwen2.5:14b",
        "llama3.1",
        "mistral",
    ],
}


class AIWorker(QThread):
    """Thread em segundo plano para chamadas à API de IA sem travar a interface."""

    concluido = pyqtSignal(str)
    erro = pyqtSignal(str)

    def __init__(self, provedor: str, prompt: str, api_key: str, modelo: str, base_url: str) -> None:
        super().__init__()
        self.provedor = provedor
        self.prompt = prompt
        self.api_key = api_key
        self.modelo = modelo
        self.base_url = base_url

    def run(self) -> None:
        try:
            resposta = chamar_ia(
                self.provedor,
                self.prompt,
                api_key=self.api_key,
                modelo=self.modelo,
                base_url=self.base_url,
            )
            self.concluido.emit(resposta)
        except Exception as exc:
            self.erro.emit(str(exc))


class ScenApp(QMainWindow):
    """Janela principal do Gerador de Trabalhos SCEN."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("SCEN — Gerador de Trabalhos Acadêmicos (DOCX + PDF + IA)")
        self.resize(1460, 900)

        self.tipo_atual = "relatorio_leitura"
        self.tema_escuro = False
        self.arquivos_contexto: list[tuple[str, str]] = []
        self.campos_dinamicos: dict[str, Any] = {}
        self.worker: AIWorker | None = None

        self._construir_ui()
        self.setStyleSheet(TEMA_OPOSTOS)
        self._carregar_config()
        self._selecionar_tipo("relatorio_leitura", {"disciplina": "Apologética", "titulo": "Por que você acredita?"})

    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout_raiz = QVBoxLayout(central)
        layout_raiz.setContentsMargins(10, 10, 10, 10)
        layout_raiz.setSpacing(8)

        # Barra superior
        barra_widget = QWidget()
        barra_widget.setObjectName("barraTopo")
        barra_topo = QHBoxLayout(barra_widget)
        barra_topo.setContentsMargins(10, 6, 10, 6)

        lbl_titulo = QLabel("SCEN — Trabalhos Acadêmicos (DOCX + PDF + IA)")
        lbl_titulo.setStyleSheet("font-size: 16px; font-weight: bold; color: #38bdf8;")
        barra_topo.addWidget(lbl_titulo)
        barra_topo.addStretch()

        self.btn_tema = QPushButton("Tema: Opostos (Claro/Escuro)")
        self.btn_tema.setObjectName("btnTopo")
        self.btn_tema.clicked.connect(self._alternar_tema)

        self.btn_modo_colunas = QPushButton("Visão Completa (3 Colunas)")
        self.btn_modo_colunas.setCheckable(True)
        self.btn_modo_colunas.setChecked(True)
        self.btn_modo_colunas.clicked.connect(lambda: self._alternar_modo("colunas"))

        self.btn_modo_abas = QPushButton("Modo Guiado (Abas)")
        self.btn_modo_abas.setCheckable(True)
        self.btn_modo_abas.clicked.connect(lambda: self._alternar_modo("abas"))

        barra_topo.addWidget(self.btn_tema)
        barra_topo.addWidget(self.btn_modo_colunas)
        barra_topo.addWidget(self.btn_modo_abas)
        layout_raiz.addWidget(barra_widget)

        # Os 3 painéis principais
        self.painel_trabalhos = self._criar_painel_trabalhos()
        self.painel_formulario = self._criar_painel_formulario()
        self.painel_ia = self._criar_painel_ia()

        self.container_conteudo = QVBoxLayout()
        layout_raiz.addLayout(self.container_conteudo, stretch=1)

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.tabs = QTabWidget()
        self._montar_modo_colunas()

        # Rodapé
        rodape_widget = QWidget()
        rodape_widget.setObjectName("rodape")
        rodape = QHBoxLayout(rodape_widget)
        rodape.setContentsMargins(10, 6, 10, 6)

        self.lbl_status = QLabel("Pronto. Escolha o trabalho à esquerda e siga as instruções no topo.")
        self.lbl_status.setStyleSheet("color: #38bdf8; font-weight: bold;")
        rodape.addWidget(self.lbl_status, stretch=1)

        btn_salvar_rasc = QPushButton("Salvar Rascunho")
        btn_salvar_rasc.clicked.connect(self._salvar_rascunho)
        btn_carregar_rasc = QPushButton("Carregar Rascunho")
        btn_carregar_rasc.clicked.connect(self._carregar_rascunho)
        btn_abrir_pasta = QPushButton("Abrir Pasta Gerados")
        btn_abrir_pasta.clicked.connect(self._abrir_pasta_gerados)
        btn_gerar = QPushButton("Gerar DOCX + PDF")
        btn_gerar.setObjectName("btnPrimario")
        btn_gerar.clicked.connect(self._gerar_arquivos)

        rodape.addWidget(btn_salvar_rasc)
        rodape.addWidget(btn_carregar_rasc)
        rodape.addWidget(btn_abrir_pasta)
        rodape.addWidget(btn_gerar)
        layout_raiz.addWidget(rodape_widget)

    def _alternar_tema(self) -> None:
        self.tema_escuro = not self.tema_escuro
        if self.tema_escuro:
            self.setStyleSheet(TEMA_ESCURO_TOTAL)
            self.btn_tema.setText("Tema: Escuro Total")
        else:
            self.setStyleSheet(TEMA_OPOSTOS)
            self.btn_tema.setText("Tema: Opostos (Claro/Escuro)")

    def _montar_modo_colunas(self) -> None:
        while self.container_conteudo.count():
            self.container_conteudo.takeAt(0)
        self.tabs.hide()
        self.splitter.show()
        self.splitter.addWidget(self.painel_trabalhos)
        self.splitter.addWidget(self.painel_formulario)
        self.splitter.addWidget(self.painel_ia)
        self.splitter.setSizes([280, 700, 460])
        self.container_conteudo.addWidget(self.splitter)

    def _montar_modo_abas(self) -> None:
        while self.container_conteudo.count():
            self.container_conteudo.takeAt(0)
        self.splitter.hide()
        self.tabs.clear()
        self.tabs.addTab(self.painel_trabalhos, "1. Escolher Trabalho")
        self.tabs.addTab(self.painel_formulario, "2. O que Preencher & Formulário")
        self.tabs.addTab(self.painel_ia, "3. Contexto, Comandos & IA")
        self.tabs.show()
        self.container_conteudo.addWidget(self.tabs)

    def _alternar_modo(self, modo: str) -> None:
        if modo == "colunas":
            self.btn_modo_colunas.setChecked(True)
            self.btn_modo_abas.setChecked(False)
            self._montar_modo_colunas()
        else:
            self.btn_modo_colunas.setChecked(False)
            self.btn_modo_abas.setChecked(True)
            self._montar_modo_abas()

    def _criar_painel_trabalhos(self) -> QWidget:
        w = QWidget()
        w.setObjectName("painelTrabalhos")
        lay = QVBoxLayout(w)
        lay.setContentsMargins(6, 6, 6, 6)

        grp_atalhos = QGroupBox("Trabalhos da Lista (SCEN)")
        l1 = QVBoxLayout(grp_atalhos)
        self.lista_atalhos = QListWidget()
        for nome, tipo, pre in ATALHOS_DISCIPLINAS:
            item = QListWidgetItem(nome)
            item.setData(Qt.ItemDataRole.UserRole, (tipo, pre))
            self.lista_atalhos.addItem(item)
        self.lista_atalhos.itemClicked.connect(self._on_atalho_clicado)
        l1.addWidget(self.lista_atalhos)
        lay.addWidget(grp_atalhos)

        grp_todos = QGroupBox("Todos os 12 Modelos")
        l2 = QVBoxLayout(grp_todos)
        self.lista_modelos = QListWidget()
        for nome, tipo in MODELOS_LISTA:
            item = QListWidgetItem(nome)
            item.setData(Qt.ItemDataRole.UserRole, tipo)
            self.lista_modelos.addItem(item)
        self.lista_modelos.itemClicked.connect(self._on_modelo_clicado)
        l2.addWidget(self.lista_modelos)
        lay.addWidget(grp_todos)
        return w

    def _criar_painel_formulario(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        cont = QWidget()
        cont.setObjectName("conteudoForm")
        self.lay_form_root = QVBoxLayout(cont)
        self.lay_form_root.setContentsMargins(10, 8, 10, 10)

        # Banner de Instruções e O que Preencher NO TOPO
        self.banner_instrucoes = QFrame()
        self.banner_instrucoes.setObjectName("bannerInstrucoes")
        lay_ban = QVBoxLayout(self.banner_instrucoes)
        self.lbl_instr_titulo = QLabel("O que preencher neste trabalho")
        self.lbl_instr_titulo.setStyleSheet("font-size: 14px; font-weight: bold; color: #38bdf8;")
        self.lbl_instr_corpo = QLabel("")
        self.lbl_instr_corpo.setWordWrap(True)
        self.lbl_instr_passo = QLabel(
            "Passo a passo: (1) Preencha o Cabeçalho abaixo → (2) Use Contexto + IA à direita ou preencha manualmente → (3) Clique em 'Gerar DOCX + PDF'."
        )
        self.lbl_instr_passo.setWordWrap(True)
        self.lbl_instr_passo.setStyleSheet("color: #cbd5e1; font-size: 11px; margin-top: 4px;")
        lay_ban.addWidget(self.lbl_instr_titulo)
        lay_ban.addWidget(self.lbl_instr_corpo)
        lay_ban.addWidget(self.lbl_instr_passo)
        self.lay_form_root.addWidget(self.banner_instrucoes)

        grp_cab = QGroupBox("1. Cabeçalho Obrigatório (Capa e Folha de Rosto)")
        form_cab = QFormLayout(grp_cab)
        self.ed_aluno = QLineEdit()
        self.ed_aluno.setPlaceholderText("Nome completo do(a) aluno(a)")
        self.ed_curso = QLineEdit()
        self.ed_curso.setPlaceholderText("Ex.: Curso Avançado em Teologia")
        self.ed_ano = QLineEdit("2026")
        self.ed_disciplina = QLineEdit()
        self.ed_disciplina.setPlaceholderText("Ex.: Apologética, Soteriologia...")
        self.ed_professor = QLineEdit()
        self.ed_professor.setPlaceholderText("Nome completo do(a) professor(a)")
        self.ed_titulo = QLineEdit()
        self.ed_titulo.setPlaceholderText("Título do trabalho ou nome do livro lido")
        self.ed_subtitulo = QLineEdit()
        self.ed_subtitulo.setPlaceholderText("Opcional")
        self.ed_cidade = QLineEdit(CIDADE_PADRAO)

        form_cab.addRow("Aluno(a) *:", self.ed_aluno)
        form_cab.addRow("Curso *:", self.ed_curso)
        form_cab.addRow("Ano *:", self.ed_ano)
        form_cab.addRow("Disciplina *:", self.ed_disciplina)
        form_cab.addRow("Professor(a) *:", self.ed_professor)
        form_cab.addRow("Título / Obra *:", self.ed_titulo)
        form_cab.addRow("Subtítulo:", self.ed_subtitulo)
        form_cab.addRow("Cidade:", self.ed_cidade)
        self.lay_form_root.addWidget(grp_cab)

        self.grp_corpo = QGroupBox("2. Conteúdo Específico do Trabalho")
        self.lay_corpo = QVBoxLayout(self.grp_corpo)
        self.lay_form_root.addWidget(self.grp_corpo)
        self.lay_form_root.addStretch()

        scroll.setWidget(cont)
        return scroll

    def _criar_painel_ia(self) -> QWidget:
        w = QWidget()
        w.setObjectName("painelIA")
        lay = QVBoxLayout(w)
        lay.setContentsMargins(6, 6, 6, 6)

        grp_cfg = QGroupBox("Configuração de IA (Opcional — ou use Copiar Prompt)")
        fc = QFormLayout(grp_cfg)
        self.cb_provedor = QComboBox()
        self.cb_provedor.addItems(["Gemini", "OpenAI", "Ollama"])
        self.cb_modelo_ia = QComboBox()
        self.cb_modelo_ia.setEditable(True)
        self.cb_modelo_ia.addItems(MODELOS_POR_PROVEDOR["Gemini"])
        btn_buscar_mod = QPushButton("Buscar da API")
        btn_buscar_mod.setToolTip("Consulta os modelos disponíveis na sua conta/API ou Ollama local")
        btn_buscar_mod.clicked.connect(self._buscar_modelos_ia)
        h_mod = QHBoxLayout()
        h_mod.addWidget(self.cb_modelo_ia, stretch=1)
        h_mod.addWidget(btn_buscar_mod)

        self.ed_api_key = QLineEdit()
        self.ed_api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.ed_api_key.setPlaceholderText("Chave API (não precisa para Ollama)")
        self.ed_base_url = QLineEdit()
        self.ed_base_url.setPlaceholderText("Opcional (Ollama/OpenAI compatível)")
        self.cb_provedor.currentTextChanged.connect(self._on_provedor_mudou)

        fc.addRow("Provedor:", self.cb_provedor)
        fc.addRow("Modelo:", h_mod)
        fc.addRow("API Key:", self.ed_api_key)
        fc.addRow("Base URL:", self.ed_base_url)
        lay.addWidget(grp_cfg)

        grp_ctx = QGroupBox("Contexto, Anexos e Comandos para IA")
        lctx = QVBoxLayout(grp_ctx)

        h_anexos = QHBoxLayout()
        btn_anexar = QPushButton("+ Anexar PDF / DOCX / TXT")
        btn_anexar.clicked.connect(self._anexar_arquivos)
        btn_limpar_anexos = QPushButton("Limpar")
        btn_limpar_anexos.clicked.connect(self._limpar_anexos)
        h_anexos.addWidget(btn_anexar)
        h_anexos.addWidget(btn_limpar_anexos)
        lctx.addLayout(h_anexos)

        self.lista_anexos = QListWidget()
        self.lista_anexos.setMaximumHeight(75)
        lctx.addWidget(self.lista_anexos)

        lctx.addWidget(QLabel("Notas / Texto de Contexto:"))
        self.ed_contexto = QPlainTextEdit()
        self.ed_contexto.setPlaceholderText("Cole aqui resumos, trechos do livro, esboço ou anotações de aula...")
        lctx.addWidget(self.ed_contexto, stretch=2)

        lctx.addWidget(QLabel("Comando / Instrução para a IA:"))
        self.ed_comando = QPlainTextEdit()
        self.ed_comando.setMaximumHeight(80)
        self.ed_comando.setPlaceholderText("Ex.: Preencha as 7 perguntas em tom acadêmico teológico e aplique ao ministério infantil.")
        lctx.addWidget(self.ed_comando)

        h_chips = QHBoxLayout()
        for rotulo, cmd in [
            ("Acadêmico", "Escreva em português acadêmico claro e bem fundamentado."),
            ("Aplicação Ministerial", "Dê ênfase prática à aplicação pastoral e de EBD."),
            ("Conciso", "Seja objetivo e direto nos parágrafos."),
        ]:
            b = QPushButton(rotulo)
            b.clicked.connect(lambda _=False, c=cmd: self._adicionar_comando(c))
            h_chips.addWidget(b)
        lctx.addLayout(h_chips)

        self.chk_preservar = QCheckBox("Manter campos já preenchidos no formulário")
        self.chk_preservar.setChecked(True)
        lctx.addWidget(self.chk_preservar)

        self.btn_completar_ia = QPushButton("Completar com IA")
        self.btn_completar_ia.setObjectName("btnIA")
        self.btn_completar_ia.clicked.connect(self._executar_ia)
        lctx.addWidget(self.btn_completar_ia)

        h_manual = QHBoxLayout()
        btn_copiar_prompt = QPushButton("Copiar Prompt")
        btn_copiar_prompt.clicked.connect(self._copiar_prompt)
        btn_colar_json = QPushButton("Colar JSON da IA")
        btn_colar_json.clicked.connect(self._colar_json_ia)
        h_manual.addWidget(btn_copiar_prompt)
        h_manual.addWidget(btn_colar_json)
        lctx.addLayout(h_manual)

        lay.addWidget(grp_ctx, stretch=1)
        return w

    def _on_provedor_mudou(self, texto: str) -> None:
        self.cb_modelo_ia.clear()
        self.cb_modelo_ia.addItems(MODELOS_POR_PROVEDOR.get(texto, []))

    def _buscar_modelos_ia(self) -> None:
        prov = self.cb_provedor.currentText()
        try:
            modelos = listar_modelos_disponiveis(
                prov,
                api_key=self.ed_api_key.text().strip(),
                base_url=self.ed_base_url.text().strip(),
            )
            if not modelos:
                QMessageBox.information(self, "Modelos", "Nenhum modelo retornado pelo provedor.")
                return
            atual = self.cb_modelo_ia.currentText()
            self.cb_modelo_ia.clear()
            self.cb_modelo_ia.addItems(modelos)
            if atual in modelos:
                self.cb_modelo_ia.setCurrentText(atual)
            self.lbl_status.setText(f"{len(modelos)} modelos carregados de {prov}!")
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao buscar modelos", str(exc))

    def _on_atalho_clicado(self, item: QListWidgetItem) -> None:
        tipo, pre = item.data(Qt.ItemDataRole.UserRole)
        self._selecionar_tipo(tipo, pre)

    def _on_modelo_clicado(self, item: QListWidgetItem) -> None:
        tipo = item.data(Qt.ItemDataRole.UserRole)
        self._selecionar_tipo(tipo)

    def _limpar_layout_corpo(self) -> None:
        while self.lay_corpo.count():
            child = self.lay_corpo.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
            elif child.layout():
                self._limpar_sublayout(child.layout())

    def _limpar_sublayout(self, layout: Any) -> None:
        while layout.count():
            c = layout.takeAt(0)
            if c.widget():
                c.widget().deleteLater()
            elif c.layout():
                self._limpar_sublayout(c.layout())

    def _selecionar_tipo(self, tipo: str, prefill: dict[str, Any] | None = None) -> None:
        self.tipo_atual = tipo
        self.campos_dinamicos.clear()
        self._limpar_layout_corpo()

        titulo_instr, corpo_instr = INSTRUCOES_POR_TIPO.get(tipo, (tipo, ""))
        self.lbl_instr_titulo.setText(f"O que preencher — {titulo_instr}")
        self.lbl_instr_corpo.setText(corpo_instr)

        if prefill:
            if "disciplina" in prefill:
                self.ed_disciplina.setText(prefill["disciplina"])
            if "titulo" in prefill:
                self.ed_titulo.setText(prefill["titulo"])
            if "subtitulo" in prefill:
                self.ed_subtitulo.setText(prefill["subtitulo"])

        self.grp_corpo.setTitle(f"2. Conteúdo Específico — {titulo_instr}")

        def add_line(chave: str, rotulo: str, placeholder: str = "") -> QLineEdit:
            self.lay_corpo.addWidget(QLabel(rotulo))
            ed = QLineEdit()
            ed.setPlaceholderText(placeholder)
            self.lay_corpo.addWidget(ed)
            self.campos_dinamicos[chave] = ed
            return ed

        def add_text(chave: str, rotulo: str, altura: int = 85) -> QPlainTextEdit:
            self.lay_corpo.addWidget(QLabel(rotulo))
            ed = QPlainTextEdit()
            ed.setMinimumHeight(altura)
            self.lay_corpo.addWidget(ed)
            self.campos_dinamicos[chave] = ed
            return ed

        if tipo == "declaracao_atual":
            add_line("referencia_abnt", "Referência ABNT da obra *:", "SOBRENOME, Nome. Título. Cidade: Editora, Ano.")
            self.lay_corpo.addWidget(QLabel("Cronograma de Leitura *:"))
            tab = QTableWidget(3, 4)
            tab.setHorizontalHeaderLabels(["Seção", "Título", "Datas", "Horas"])
            tab.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            tab.setMinimumHeight(200)
            self.lay_corpo.addWidget(tab)
            self.campos_dinamicos["cronograma"] = tab
            hb = QHBoxLayout()
            b_add = QPushButton("+ Linha")
            b_add.clicked.connect(lambda: tab.insertRow(tab.rowCount()))
            b_rem = QPushButton("- Remover Linha")
            b_rem.clicked.connect(lambda: tab.removeRow(tab.currentRow()) if tab.currentRow() >= 0 else None)
            hb.addWidget(b_add)
            hb.addWidget(b_rem)
            self.lay_corpo.addLayout(hb)

        elif tipo == "declaracao_antiga":
            add_line("data_assinatura", "Data da assinatura *:", "16 / 11 / 2026")

        elif tipo == "relatorio_leitura":
            add_line("referencia_abnt", "Referência ABNT da obra *:")
            perguntas = [
                ("q1_tema", "1. Qual o tema da obra?"),
                ("q2_ideia_central", "2. Qual a ideia central e o propósito da obra?"),
                ("q3_argumentos", "3. Quais os principais argumentos apresentados pelo autor?"),
                ("q4_problematizacao", "4. Quais os pontos de problematização e elementos contrastantes?"),
                ("q5_sintese", "5. Faça uma síntese do livro."),
                ("q6_avaliacao_ministerio", "6. Avaliação pessoal e utilidade no ministério:"),
                ("q7_conclusao", "7. Conclusão da leitura:"),
            ]
            for ch, rot in perguntas:
                add_text(ch, rot, 75)

        elif tipo in ("paper", "artigo"):
            if tipo == "paper":
                add_text("sinopse", "Sinopse *:", 75)
            else:
                add_text("resumo", "Resumo *:", 85)
                add_line("palavras_chave", "Palavras-chave (separadas por ;) *:")
            add_text(
                "secoes_texto",
                "Seções (use '## Título da Seção' em uma linha e os parágrafos abaixo):",
                240,
            ).setPlainText("## 1 INTRODUÇÃO\n\n## 2 DESENVOLVIMENTO\n\n## 3 CONCLUSÃO\n")
            add_text("referencias_texto", "Referências ABNT (uma por linha):", 90)

        elif tipo == "exegese_nt":
            add_line("passagem", "Passagem Bíblica Base *:", "Ex.: Rm 8:1-11")
            for i, nome_passo in enumerate(
                [
                    "Passo 1 — Delimitação da Perícope e Contexto",
                    "Passo 2 — Crítica Textual e Tradução Literal",
                    "Passo 3 — Análise Gramatical, Sintática e Léxica",
                    "Passo 4 — Análise EstrUTural e Literária".replace("EstrUTural", "Estrutural"),
                    "Passo 5 — Contexto Teológico e Bíblico",
                    "Passo 6 — Ideia Central Exegética",
                    "Passo 7 — Aplicação Contemporânea e Homilética",
                ],
                start=1,
            ):
                add_text(f"passo_{i}", nome_passo, 75)
            add_text("referencias_texto", "Referências ABNT (uma por linha):", 80)

        elif tipo.startswith("sermao_"):
            add_line("passagem", "Texto Bíblico Base *:", "Ex.: Lc 15:11-32")
            for ch, rot in [
                ("sec_1", "1. Dados Gerais / Tema / Ideia Central"),
                ("sec_2", "2. Contexto Bíblico e Histórico"),
                ("sec_3", "3. Introdução"),
                ("sec_4", "4. Desenvolvimento / Exposição"),
                ("sec_5", "5. Aplicação Prática"),
                ("sec_6", "6. Conclusão"),
            ]:
                add_text(ch, rot, 75)

        elif tipo == "fichamento":
            add_line("tipo_fichamento", "Tipo de Fichamento *:", "Resumo / Transcrição / Comentário")
            add_line("referencia_abnt", "Referência ABNT da obra *:")
            add_text(
                "fichas_texto",
                "Fichas (use '## Cabeçalho da Ficha' e o conteúdo abaixo):",
                220,
            ).setPlainText("## Ficha 1 — Capítulo 1\n")

        elif tipo == "resumo":
            add_line("referencia_abnt", "Referência ABNT da obra *:")
            lbl_cont = QLabel("Palavras: 0 (recomendado SCEN: 150 a 500)")
            ed_res = add_text("texto_resumo", "Texto do Resumo (parágrafo único) *:", 180)
            ed_res.textChanged.connect(
                lambda: lbl_cont.setText(
                    f"Palavras: {len(ed_res.toPlainText().split())} (recomendado SCEN: 150 a 500)"
                )
            )
            self.lay_corpo.addWidget(lbl_cont)
            add_line("palavras_chave", "Palavras-chave (separadas por ;) *:")

        elif tipo == "resenha":
            add_line("referencia_abnt", "Referência ABNT da obra *:")
            add_text("paragrafos_texto", "Parágrafos da Resenha (separe por linha em branco):", 240)

        self.lbl_status.setText(f"Modelo ativo: {titulo_instr}")

    def _coletar_dados(self) -> dict[str, Any]:
        dados: dict[str, Any] = {
            "aluno": self.ed_aluno.text().strip(),
            "curso": self.ed_curso.text().strip(),
            "ano": self.ed_ano.text().strip() or "2026",
            "disciplina": self.ed_disciplina.text().strip(),
            "professor": self.ed_professor.text().strip(),
            "titulo": self.ed_titulo.text().strip(),
            "subtitulo": self.ed_subtitulo.text().strip(),
            "cidade": self.ed_cidade.text().strip() or CIDADE_PADRAO,
        }

        def val(ch: str) -> str:
            w = self.campos_dinamicos.get(ch)
            if isinstance(w, QLineEdit):
                return w.text().strip()
            if isinstance(w, QPlainTextEdit):
                return w.toPlainText().strip()
            return ""

        t = self.tipo_atual
        if t == "declaracao_atual":
            dados["referencia_abnt"] = val("referencia_abnt")
            tab: QTableWidget = self.campos_dinamicos["cronograma"]
            cron = []
            for r in range(tab.rowCount()):
                linha = {
                    "secao": tab.item(r, 0).text().strip() if tab.item(r, 0) else "",
                    "titulo": tab.item(r, 1).text().strip() if tab.item(r, 1) else "",
                    "datas": tab.item(r, 2).text().strip() if tab.item(r, 2) else "",
                    "horas": tab.item(r, 3).text().strip() if tab.item(r, 3) else "",
                }
                if any(linha.values()):
                    cron.append(linha)
            dados["cronograma"] = cron
        elif t == "declaracao_antiga":
            dados["data_assinatura"] = val("data_assinatura")
        elif t == "relatorio_leitura":
            dados["referencia_abnt"] = val("referencia_abnt")
            for k in ["q1_tema", "q2_ideia_central", "q3_argumentos", "q4_problematizacao", "q5_sintese", "q6_avaliacao_ministerio", "q7_conclusao"]:
                dados[k] = val(k)
        elif t in ("paper", "artigo"):
            if t == "paper":
                dados["sinopse"] = val("sinopse")
            else:
                dados["resumo"] = val("resumo")
                dados["palavras_chave"] = val("palavras_chave")
            dados["secoes"] = self._parse_secoes_markdown(val("secoes_texto"))
            dados["referencias"] = [r.strip() for r in val("referencias_texto").splitlines() if r.strip()]
        elif t == "exegese_nt":
            dados["passagem"] = val("passagem")
            titulos = [
                "PASSO 1 — DELIMITAÇÃO DA PERÍCOPE E CONTEXTO HISTÓRICO-LITERÁRIO",
                "PASSO 2 — CRÍTICA TEXTUAL E TRADUÇÃO LITERAL",
                "PASSO 3 — ANÁLISE GRAMATICAL, SINTÁTICA E LÉXICA",
                "PASSO 4 — ANÁLISE ESTRUTURAL E LITERÁRIA",
                "PASSO 5 — CONTEXTO TEOLÓGICO E BÍBLICO",
                "PASSO 6 — IDEIA CENTRAL EXEGÉTICA (SÍNTESE)",
                "PASSO 7 — APLICAÇÃO CONTEMPORÂNEA E HOMILÉTICA",
            ]
            dados["secoes"] = [
                {"titulo": titulos[i - 1], "paragrafos": [p for p in val(f"passo_{i}").split("\n\n") if p.strip()]}
                for i in range(1, 8)
            ]
            dados["referencias"] = [r.strip() for r in val("referencias_texto").splitlines() if r.strip()]
        elif t.startswith("sermao_"):
            dados["passagem"] = val("passagem")
            titulos = [
                "1 DADOS GERAIS DO SERMÃO",
                "2 CONTEXTO BÍBLICO E HISTÓRICO",
                "3 INTRODUÇÃO",
                "4 DESENVOLVIMENTO",
                "5 APLICAÇÃO PRÁTICA",
                "6 CONCLUSÃO",
            ]
            dados["secoes"] = [
                {"titulo": titulos[i - 1], "paragrafos": [p for p in val(f"sec_{i}").split("\n\n") if p.strip()]}
                for i in range(1, 7)
            ]
        elif t == "fichamento":
            dados["tipo_fichamento"] = val("tipo_fichamento")
            dados["referencia_abnt"] = val("referencia_abnt")
            secoes = self._parse_secoes_markdown(val("fichas_texto"))
            dados["fichas"] = [{"cabecalho": s["titulo"], "conteudo": "\n\n".join(s["paragrafos"])} for s in secoes]
        elif t == "resumo":
            dados["referencia_abnt"] = val("referencia_abnt")
            dados["texto_resumo"] = val("texto_resumo")
            dados["palavras_chave"] = val("palavras_chave")
        elif t == "resenha":
            dados["referencia_abnt"] = val("referencia_abnt")
            dados["paragrafos"] = [p.strip() for p in val("paragrafos_texto").split("\n\n") if p.strip()]

        return dados

    def _parse_secoes_markdown(self, texto: str) -> list[dict[str, Any]]:
        secoes: list[dict[str, Any]] = []
        atual: dict[str, Any] | None = None
        for linha in texto.splitlines():
            if linha.strip().startswith("## "):
                if atual:
                    secoes.append(atual)
                atual = {"titulo": linha.strip()[3:].strip(), "paragrafos": []}
            elif linha.strip():
                if not atual:
                    atual = {"titulo": "1 SEÇÃO", "paragrafos": []}
                atual["paragrafos"].append(linha.strip())
        if atual:
            secoes.append(atual)
        return secoes

    def _aplicar_dados(self, dados: dict[str, Any], apenas_vazios: bool = False) -> None:
        def set_line(ed: QLineEdit, val: Any) -> None:
            if val is None or (apenas_vazios and ed.text().strip()):
                return
            ed.setText(str(val))

        def set_text(ed: QPlainTextEdit, val: Any) -> None:
            if val is None or (apenas_vazios and ed.toPlainText().strip()):
                return
            ed.setPlainText(str(val))

        set_line(self.ed_aluno, dados.get("aluno"))
        set_line(self.ed_curso, dados.get("curso"))
        set_line(self.ed_ano, dados.get("ano"))
        set_line(self.ed_disciplina, dados.get("disciplina"))
        set_line(self.ed_professor, dados.get("professor"))
        set_line(self.ed_titulo, dados.get("titulo"))
        set_line(self.ed_subtitulo, dados.get("subtitulo"))
        set_line(self.ed_cidade, dados.get("cidade"))

        t = self.tipo_atual
        if t == "declaracao_atual":
            set_line(self.campos_dinamicos["referencia_abnt"], dados.get("referencia_abnt"))
            if "cronograma" in dados and isinstance(dados["cronograma"], list):
                tab: QTableWidget = self.campos_dinamicos["cronograma"]
                tab.setRowCount(0)
                for row_data in dados["cronograma"]:
                    r = tab.rowCount()
                    tab.insertRow(r)
                    for idx, k in enumerate(["secao", "titulo", "datas", "horas"]):
                        tab.setItem(r, idx, QTableWidgetItem(str(row_data.get(k, ""))))
        elif t == "declaracao_antiga":
            set_line(self.campos_dinamicos["data_assinatura"], dados.get("data_assinatura"))
        elif t == "relatorio_leitura":
            set_line(self.campos_dinamicos["referencia_abnt"], dados.get("referencia_abnt"))
            for k in ["q1_tema", "q2_ideia_central", "q3_argumentos", "q4_problematizacao", "q5_sintese", "q6_avaliacao_ministerio", "q7_conclusao"]:
                set_text(self.campos_dinamicos[k], dados.get(k))
        elif t in ("paper", "artigo"):
            if t == "paper":
                set_text(self.campos_dinamicos["sinopse"], dados.get("sinopse"))
            else:
                set_text(self.campos_dinamicos["resumo"], dados.get("resumo"))
                set_line(self.campos_dinamicos["palavras_chave"], dados.get("palavras_chave"))
            if "secoes" in dados:
                bloco = "\n\n".join(f"## {s['titulo']}\n" + "\n\n".join(s.get("paragrafos", [])) for s in dados["secoes"])
                set_text(self.campos_dinamicos["secoes_texto"], bloco)
            if "referencias" in dados:
                set_text(self.campos_dinamicos["referencias_texto"], "\n".join(dados["referencias"]))
        elif t == "exegese_nt":
            set_line(self.campos_dinamicos["passagem"], dados.get("passagem"))
            if "secoes" in dados:
                for i, sec in enumerate(dados["secoes"][:7], start=1):
                    set_text(self.campos_dinamicos[f"passo_{i}"], "\n\n".join(sec.get("paragrafos", [])))
            if "referencias" in dados:
                set_text(self.campos_dinamicos["referencias_texto"], "\n".join(dados["referencias"]))
        elif t.startswith("sermao_"):
            set_line(self.campos_dinamicos["passagem"], dados.get("passagem"))
            if "secoes" in dados:
                for i, sec in enumerate(dados["secoes"][:6], start=1):
                    set_text(self.campos_dinamicos[f"sec_{i}"], "\n\n".join(sec.get("paragrafos", [])))
        elif t == "fichamento":
            set_line(self.campos_dinamicos["tipo_fichamento"], dados.get("tipo_fichamento"))
            set_line(self.campos_dinamicos["referencia_abnt"], dados.get("referencia_abnt"))
            if "fichas" in dados:
                bloco = "\n\n".join(f"## {f['cabecalho']}\n{f['conteudo']}" for f in dados["fichas"])
                set_text(self.campos_dinamicos["fichas_texto"], bloco)
        elif t == "resumo":
            set_line(self.campos_dinamicos["referencia_abnt"], dados.get("referencia_abnt"))
            set_text(self.campos_dinamicos["texto_resumo"], dados.get("texto_resumo"))
            set_line(self.campos_dinamicos["palavras_chave"], dados.get("palavras_chave"))
        elif t == "resenha":
            set_line(self.campos_dinamicos["referencia_abnt"], dados.get("referencia_abnt"))
            if "paragrafos" in dados:
                set_text(self.campos_dinamicos["paragrafos_texto"], "\n\n".join(dados["paragrafos"]))

    def _anexar_arquivos(self) -> None:
        caminhos, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecionar arquivos de contexto",
            str(BASE_DIR),
            "Documentos (*.pdf *.docx *.txt *.md)",
        )
        for cam in caminhos:
            p = Path(cam)
            try:
                txt = extrair_texto_arquivo(p)
                self.arquivos_contexto.append((p.name, txt))
                self.lista_anexos.addItem(f"{p.name} ({len(txt)} chars)")
            except Exception as exc:
                QMessageBox.warning(self, "Erro ao ler arquivo", f"{p.name}: {exc}")

    def _limpar_anexos(self) -> None:
        self.arquivos_contexto.clear()
        self.lista_anexos.clear()

    def _adicionar_comando(self, cmd: str) -> None:
        atual = self.ed_comando.toPlainText().strip()
        self.ed_comando.setPlainText(f"{atual}\n{cmd}".strip() if atual else cmd)

    def _montar_prompt(self) -> str:
        return montar_prompt_ia(
            self.tipo_atual,
            self._coletar_dados(),
            self.ed_contexto.toPlainText().strip(),
            self.arquivos_contexto,
            self.ed_comando.toPlainText().strip(),
        )

    def _copiar_prompt(self) -> None:
        QApplication.clipboard().setText(self._montar_prompt())
        self.lbl_status.setText("Prompt copiado para a área de transferência!")

    def _colar_json_ia(self) -> None:
        texto, ok = QInputDialog.getMultiLineText(self, "Colar JSON da IA", "Cole aqui o JSON retornado pela IA:")
        if not ok or not texto.strip():
            return
        try:
            dados = extrair_json_resposta(texto)
            self._aplicar_dados(dados, apenas_vazios=self.chk_preservar.isChecked())
            self.lbl_status.setText("Dados da IA importados no formulário!")
        except Exception as exc:
            QMessageBox.critical(self, "JSON inválido", str(exc))

    def _executar_ia(self) -> None:
        self._salvar_config()
        self.btn_completar_ia.setEnabled(False)
        self.lbl_status.setText("Consultando IA...")
        self.worker = AIWorker(
            self.cb_provedor.currentText(),
            self._montar_prompt(),
            self.ed_api_key.text().strip(),
            self.cb_modelo_ia.currentText().strip(),
            self.ed_base_url.text().strip(),
        )
        self.worker.concluido.connect(self._on_ia_ok)
        self.worker.erro.connect(self._on_ia_erro)
        self.worker.start()

    def _on_ia_ok(self, resposta: str) -> None:
        self.btn_completar_ia.setEnabled(True)
        try:
            dados = extrair_json_resposta(resposta)
            self._aplicar_dados(dados, apenas_vazios=self.chk_preservar.isChecked())
            self.lbl_status.setText("Formulário completado com IA!")
        except Exception as exc:
            QMessageBox.warning(self, "Resposta não-JSON", f"Erro ao interpretar JSON: {exc}\n\n{resposta[:500]}")
            self.lbl_status.setText("Erro ao interpretar resposta da IA.")

    def _on_ia_erro(self, erro: str) -> None:
        self.btn_completar_ia.setEnabled(True)
        self.lbl_status.setText("Erro na IA.")
        QMessageBox.critical(self, "Erro na chamada de IA", erro)

    def _gerar_arquivos(self) -> None:
        dados = self._coletar_dados()
        slug = (dados.get("titulo") or self.tipo_atual).strip().replace(" ", "_")
        slug = "".join(c for c in slug if c.isalnum() or c in "_-") or self.tipo_atual
        destino_base = GERADOS_DIR / f"{self.tipo_atual}_{slug}"
        try:
            docx_p, pdf_p = gerar_documento(self.tipo_atual, dados, destino_base)
            self.lbl_status.setText(f"Gerados: {docx_p.name} e {pdf_p.name}")
            QMessageBox.information(
                self,
                "Trabalho Gerado!",
                f"Arquivos criados em '{GERADOS_DIR.name}/':\n\n- {docx_p.name}\n- {pdf_p.name}",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Erro ao gerar", str(exc))

    def _salvar_rascunho(self) -> None:
        RASCUNHOS_DIR.mkdir(parents=True, exist_ok=True)
        dados = self._coletar_dados()
        dados["_tipo"] = self.tipo_atual
        nome = (dados.get("titulo") or self.tipo_atual).replace(" ", "_")
        caminho, _ = QFileDialog.getSaveFileName(self, "Salvar Rascunho", str(RASCUNHOS_DIR / f"{nome}.json"), "JSON (*.json)")
        if caminho:
            Path(caminho).write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
            self.lbl_status.setText(f"Rascunho salvo: {Path(caminho).name}")

    def _carregar_rascunho(self) -> None:
        RASCUNHOS_DIR.mkdir(parents=True, exist_ok=True)
        caminho, _ = QFileDialog.getOpenFileName(self, "Carregar Rascunho", str(RASCUNHOS_DIR), "JSON (*.json)")
        if caminho:
            dados = json.loads(Path(caminho).read_text(encoding="utf-8"))
            if "_tipo" in dados:
                self._selecionar_tipo(dados["_tipo"])
            self._aplicar_dados(dados, apenas_vazios=False)
            self.lbl_status.setText(f"Rascunho carregado: {Path(caminho).name}")

    def _abrir_pasta_gerados(self) -> None:
        GERADOS_DIR.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(str(GERADOS_DIR))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(GERADOS_DIR)])
        else:
            subprocess.Popen(["xdg-open", str(GERADOS_DIR)])

    def _salvar_config(self) -> None:
        cfg = {
            "provedor": self.cb_provedor.currentText(),
            "modelo": self.cb_modelo_ia.currentText().strip(),
            "api_key": self.ed_api_key.text().strip(),
            "base_url": self.ed_base_url.text().strip(),
        }
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")

    def _carregar_config(self) -> None:
        if CONFIG_PATH.exists():
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            self.cb_provedor.setCurrentText(cfg.get("provedor", "Gemini"))
            self.cb_modelo_ia.setCurrentText(cfg.get("modelo", "gemini-4-carbon"))
            self.ed_api_key.setText(cfg.get("api_key", ""))
            self.ed_base_url.setText(cfg.get("base_url", ""))


def main() -> None:
    app = QApplication(sys.argv)
    win = ScenApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
