# Guia de Informações Essenciais por Trabalho — SCEN

Cada subpasta abaixo contém:
- `*.docx` (versão editável Word)
- `*.pdf` (versão PDF A4 formatada no padrão SCEN)
- `dados_essenciais.json` (campos exatos necessários para gerar o trabalho pronto)

## Informações Comuns a Todos os Trabalhos
- `aluno`: Nome completo do(a) aluno(a)
- `curso`: Ex.: `Curso Avançado em Teologia` ou `Mestrado em Divindade`
- `ano`: Ano de entrega (ex.: `2026`)
- `disciplina`: Nome da disciplina
- `professor`: Nome completo do(a) professor(a)
- `titulo` (e `subtitulo`, se houver)

*(Padrões automáticos: Instituição = `SEMINÁRIO CRISTÃO EVANGÉLICO DO NORTE`, Cidade = `São Luís`.)*

---

## Informações Específicas por Tipo de Trabalho

1. **Declaração de Leitura Atual** (`01_Declaracao_Leitura_Atual/`)
   - `referencia_abnt` do livro lido
   - `cronograma`: lista de seções/capítulos com `secao`, `titulo`, `datas` e `horas`
2. **Declaração de Leitura Antiga** (`02_Declaracao_Leitura_Antiga/`)
   - `data_assinatura` (ex.: `16 / 11 / 2026`)
3. **Relatório de Leitura** (`03_Relatorio_de_Leitura/`)
   - `referencia_abnt` da obra
   - Respostas das 7 questões oficiais (`q1_tema`, `q2_ideia_central`, `q3_argumentos`, `q4_problematizacao`, `q5_sintese`, `q6_avaliacao_ministerio`, `q7_conclusao`)
4. **Paper** (`04_Paper/`)
   - `sinopse`
   - `secoes` (Introdução, Desenvolvimento, Conclusão)
   - `referencias` (lista ABNT)
5. **Exegese do Novo Testamento — 7 Passos** (`05_Exegese_Novo_Testamento_7_Passos/`)
   - `passagem` bíblica (ex.: `Rm 8:1-11`)
   - Conteúdo dos 7 Passos (delimitação/contexto, crítica textual/tradução, análise gramatical/léxica, estrutura literária, contexto teológico, ideia central, aplicação)
   - `referencias` (lista ABNT)
6. **Sermão em Narrativa / Evangelho / Parábola** (`06_Sermao_Narrativa/`, `07_Sermao_Evangelho/`, `08_Sermao_Parabola/`)
   - `passagem` bíblica
   - Tema / Ideia Central + pontos do esboço (Contexto, Introdução, Desenvolvimento, Aplicação, Conclusão)
7. **Fichamento** (`09_Fichamento/`)
   - `tipo_fichamento` (Resumo, Transcrição ou Comentário)
   - `referencia_abnt` + lista de `fichas` (`cabecalho`, `conteudo`)
8. **Resumo** (`10_Resumo/`)
   - `referencia_abnt`, `texto_resumo` (parágrafo único, 150–500 palavras), `palavras_chave`
9. **Resenha / Recensão** (`11_Resenha_Recensao/`)
   - `referencia_abnt` + `paragrafos` (apresentação, síntese e avaliação crítica)
10. **Artigo Científico** (`12_Artigo_Cientifico/`)
    - `resumo`, `palavras_chave`, `secoes` (Introdução, Referencial Teórico, Método, Resultados e Discussão, Considerações Finais), `referencias`
