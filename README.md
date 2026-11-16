# SCEN — Gerador de Trabalhos Acadêmicos (DOCX + PDF + IA)

Sistema completo e padronizado segundo o manual **SCEN 2025 (Trabalhos Acadêmicos)** para gerar trabalhos em **`.docx`** e **`.pdf`** (A4, margens 3 cm/2 cm, Arial 12/10, Capa, Folha de Rosto e Referências ABNT), com **interface gráfica moderna (PyQt6)** compatível com **Windows, Linux e macOS** e preenchimento assistido por **IA**.

---

## Como Executar

### Windows
1. Instale o **Python 3.10+** (marque *"Add Python to PATH"* durante a instalação).
2. Dê um duplo clique em **`Iniciar_SCEN.bat`** (ele cria o ambiente virtual, instala as dependências e abre o aplicativo automaticamente).

### Linux / macOS
```bash
pip install -r requirements.txt
python3 app_scen.py
```
*(Ou execute `./iniciar_scen.sh`.)*

---

## Como Usar a Interface

1. **Escolha o modo de visualização** no topo:
   - **Visão Completa (3 Colunas)**: Trabalhos à esquerda, Formulário no centro e Contexto/IA à direita.
   - **Modo Guiado (Abas)**: Passo 1 (Trabalho) → Passo 2 (O que Preencher & Formulário) → Passo 3 (Contexto, Comandos & IA).
   - **Tema**: Alterne entre **Opostos (Claro/Escuro)** e **Escuro Total**.
2. **Leia o painel "O que preencher" no topo** do formulário para ver os campos exatos exigidos pelo modelo selecionado.
3. **Complete manualmente ou com IA**:
   - Anexe arquivos **`.pdf`**, **`.docx`**, **`.txt`** ou **`.md`** ou cole suas anotações.
   - Digite um comando (ou use os atalhos *Acadêmico*, *Aplicação Ministerial*, *Conciso*) e clique em **Completar com IA** (suporta **Gemini**, **OpenAI** ou **Ollama local**).
   - Sem API Key? Clique em **Copiar Prompt**, cole em qualquer chat de IA e depois clique em **Colar JSON da IA**.
4. Clique em **Gerar DOCX + PDF** para salvar os dois arquivos prontos em `Trabalhos Gerados/`.

---

## Estrutura do Projeto

- `app_scen.py`: Interface gráfica PyQt6 (Windows/Linux/macOS).
- `ai_assistant.py`: Extração de contexto (PDF/DOCX/TXT) e integração com Gemini, OpenAI e Ollama.
- `gerar_trabalho.py`: Motor de geração de `.docx` (`python-docx`) e `.pdf` (`weasyprint` / fallback `reportlab` para Windows).
- `Informacoes_Essenciais/`: Subpastas para cada um dos 12 trabalhos contendo `.docx`, `.pdf`, `dados_essenciais.json` e o guia [00_GUIA_INFORMACOES_ESSENCIAIS.md](Informacoes_Essenciais/00_GUIA_INFORMACOES_ESSENCIAIS.md).
- `Modelos/`: Os 12 modelos prontos em `.docx` e `.pdf`.
- `Iniciar_SCEN.bat` / `iniciar_scen.sh`: Atalhos de execução rápida.

---

## Mapa dos Trabalhos da Lista (`Trabalhos a serem feitos.jpeg`)

| Disciplina | Trabalho | Modelo (`Modelos/`) |
| :--- | :--- | :--- |
| **Exegese do Novo Testamento** | Passos 1 a 7 | `05_Exegese_Novo_Testamento_7_Passos` |
| **Teologia da Família** | Paper | `04_Paper` |
| **Teologia da Família** | Família do Pastor (Declaração) | `01_Declaracao_Leitura_Atual` / `02_Declaracao_Leitura_Antiga` |
| **Apologética** | Por que você acredita? (Relatório) | `03_Relatorio_de_Leitura` |
| **Prática da Pregação** | Sermão em Narrativa / Evangelho / Parábola | `06_Sermao_Narrativa`, `07_Sermao_Evangelho`, `08_Sermao_Parabola` |
| **Soteriologia** | Salvos pela Graça (Declaração) | `01_Declaracao_Leitura_Atual` / `02_Declaracao_Leitura_Antiga` |
