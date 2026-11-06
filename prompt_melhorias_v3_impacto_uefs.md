# PROMPT v3 (MELHORIAS): ImpactoUEFS

## 0. Modo de operação (leia primeiro, economiza tokens)

Você é um Desenvolvedor Python Sênior (Streamlit, Pandas, ReportLab). O projeto **já existe e funciona** na pasta atual (`app.py`, `engine.py`, `data_manager.py`, `relatorios.py`, `test_engine.py`, `requirements.txt`, `modelos/`, `logos/`). Esta tarefa é uma **alteração incremental**. Todas as decisões abaixo estão FECHADAS.

1. **NÃO faça perguntas.** Dúvida → aplique a seção 9 (Decisões padrão).
2. **NÃO recrie o projeto nem reescreva arquivos inteiros.** Use `Edit` em trechos. Só use `Write` para arquivos novos.
3. **NÃO explore**: sem `ls`/`find` do projeto, sem ler `modelos/*.xlsx`. Leia cada arquivo `.py` **uma única vez** (`Read`), na ordem da seção 8, e edite.
4. **NÃO repita nem resuma esta especificação**. Sem explicações fora de docstrings curtas.
5. Código novo: Python 3.10+, type hints, PEP 8 (≤ 100 colunas), textos de UI/PDF/erros em português do Brasil, sem `print`, sem variáveis globais mutáveis.
6. Ao terminar, execute **uma vez**: `pip install -r requirements.txt` e `python -m pytest -q`. Corrija somente falhas. Não rode o Streamlit interativamente; valide a UI com `python -c "import app"`.
7. Execute as **verificações finais da seção 7** (greps) e corrija qualquer sobra.
8. Resposta final: **no máximo 6 linhas** (como rodar, resultado dos testes, decisões que você teve de tomar fora da seção 9).

## 1. Resumo das mudanças

| # | Mudança | Arquivos |
|---|---|---|
| M1 | Corrigir o cargo **Técnico de Nível Superior (Jornalista)**: CH única passa a **25h**; caso de referência com valores reais | `engine.py`, `data_manager.py`, `app.py`, `test_engine.py` |
| M2 | **Temporário:** remover "Contrato inicial (meses)" e "Prorrogação (meses)"; assume-se **prorrogação automática** por todo o período projetado | `engine.py`, `app.py`, `relatorios.py`, `test_engine.py` |
| M3 | **Resumo por ano** (e detalhamento mensal) passam a exibir o **detalhamento da RBM**: Vencimento Básico + Só gratificação + Incentivo = RBM | `engine.py`, `relatorios.py`, `app.py` |
| M4 | Renomear **GSTU → "Só gratificação"** em toda a interface, relatórios, planilha e código | todos |
| M5 | Incluir os **3 logos** no cabeçalho do **PDF** | `relatorios.py`, `requirements.txt`, `logos/` |
| M6 | **Remover a Formação Exigida** (campo, validação, exibição, PDF, Excel) | `engine.py`, `app.py`, `relatorios.py`, `test_engine.py` |
| M7 | Ajustes derivados: PDF em paisagem, nota metodológica, template da planilha | `relatorios.py`, `data_manager.py` |

O campo opcional **Observações** (`observacao`) **permanece** como está.

## 2. M1: Jornalista (CH 25h)

Caso de referência (deve funcionar e virar teste):

| Campo | Valor |
|---|---|
| Categoria / Vínculo | Técnicos / Temporário |
| Cargo | Técnico de Nível Superior (Jornalista) |
| Carga Horária | **25h** |
| Nível/Classe | Classe I |
| Incentivo/Adicional | GRATIFICAÇÃO POR COMPETÊNCIA - GPC (Lei 14.963/2025) |
| Vencimento_Basico | 1788.90 |
| Só gratificação (planilha: `Valor_Gratificacao`) | 1129.65 |

Alterações:
* `engine.py`: crie a constante `CH_JORNALISTA = "25h"` e use-a em `CargoDef(... CARGO_JORNALISTA ... chs=(CH_JORNALISTA,))`. Mantenha `"Classe I"` e o texto do adicional `GRATIFICAÇÃO POR COMPETÊNCIA - GPC (Lei 14.963/2025)`. O cálculo é o mesmo dos demais técnicos: `RBM = vencimento + Só gratificação` (a GPC é o valor da coluna de gratificação).
* `app.py` (caption do Jornalista): trocar o texto fixo "35h" por `f"Cargo com base legal própria: {engine.LEI_JORNALISTA} (carga horária única de {cargo.chs[0]})."`.
* `data_manager.py`: nenhuma lista fixa de CHs; tudo deriva do `CATALOGO`. Linha da planilha do Jornalista com CH `35h` deve ser **rejeitada** com mensagem específica: `Linha N: CH '35h' inválida para 'Técnico de Nível Superior (Jornalista)'; permitida(s): 25h.` (se a mensagem atual já for equivalente, mantenha; senão ajuste). `normalizar_ch` já aceita `25H → 25h` (não alterar).
* **Valores esperados do caso de referência** (ingresso 01/2027, 1 vaga), por mês: `vencimento_basico=1788.90`, `gratificacao=1129.65`, `incentivo=0.00`, `rbm=2918.55`, `decimo_terceiro=243.21`, `ferias=81.07`, `inss=648.57`, `custo_total=3891.40`. Total do período: 36 meses → `total_por_vaga = 140090.40`. Resumo Ano 1 (2027): Vencimento 21466.80, Só gratificação 13555.80, RBM 35022.60, Custo por vaga 46696.80.

## 3. M2: Temporário com prorrogação automática

* Remova de `Simulacao` os campos `meses_contrato_inicial` e `meses_prorrogacao`.
* Remova de `_validar_simulacao` todo o bloco de validação do contrato temporário.
* Em `projetar`: remova `limite_contrato`, o ramo `ativo=False`, `_mes_zerado` (função inteira) e o contador `meses_ativos`. **Todos os meses do período são cobrados** para qualquer vínculo (Efetivo ou Temporário). `custo_medio_mensal = total_por_vaga / len(meses)` (arredondado com `arredondar`).
* Remova a chave `ativo` de `CHAVES_MES` e dos dicts de mês; remova `meses_contrato_inicial` e `meses_prorrogacao` dos `metadados`. **Mantenha** `extrapolado`.
* `app.py`: remova o bloco inteiro dos dois `number_input` ("Contrato inicial (meses)" e "Prorrogação (meses)"), os `_padrao("meses_contrato_inicial"...)`/`_padrao("meses_prorrogacao"...)`, as variáveis `meses_inicial`/`meses_prorrogacao` e os dois argumentos na criação de `Simulacao`. Para cargos temporários, exiba no lugar: `st.caption("Contrato temporário: considera-se a prorrogação automática durante todo o período projetado (Ano 1 parcial + Anos 2 e 3).")`.
* `relatorios.py`: remova `OBS_ENCERRADO`, a coluna `Ativo` de `tabela_mensal`, o sufixo `(enc.)` no detalhamento do PDF e a explicação de "(enc.)" na legenda; em `_parametros` remova as linhas "Contrato inicial (meses)" e "Prorrogação (meses)" e, para vínculo Temporário, acrescente `("Regime do contrato", "Temporário, com prorrogação automática durante todo o período projetado", False)`.
* `data_manager.py`: sem alteração.

## 4. M3: Detalhamento da RBM

**Invariante:** em todo mês e em toda soma anual, `Vencimento Básico + Só gratificação + Incentivo == RBM` (os componentes já são arredondados antes de somar em `_componentes_rbm`; não reintroduza arredondamento extra).

`engine.py`:
* Dicts de mês: renomeie a chave `gstu` → `gratificacao` (ordem de `CHAVES_MES`: `mes, ano, mes_calendario, competencia, ano_projecao, vencimento_basico, gratificacao, incentivo, rbm, decimo_terceiro, ferias, inss, custo_total, extrapolado`).
* `COLUNAS_RESUMO` passa a ser exatamente (nesta ordem), usando `ROTULO_GRATIFICACAO` (seção 5) no 4º item:

```
("Ano", "Meses", "Vencimento Básico", "Só gratificação", "Incentivo", "RBM",
 "13º Salário", "Adicional de Férias", "INSS Patronal", "Custo por Vaga", "Custo Total (vagas)")
```
* `resumo_anual`: some também `vencimento_basico`, `gratificacao` e `incentivo` por ano (e na linha `TOTAL`). Colunas monetárias = `COLUNAS_RESUMO[2:]` (já é o que `relatorios.COLUNAS_MONETARIAS_RESUMO` usa; não alterar essa expressão). Exiba sempre as 3 colunas de composição, mesmo quando zeradas (R$ 0,00), para manter layout uniforme.

`relatorios.py`:
* `tabela_mensal`: colunas `Mês, Competência, Ano, Vencimento Básico, Só gratificação, Incentivo, RBM, 13º Salário, Adicional de Férias, INSS Patronal, Custo Total (por vaga), Observação Automática` (sem `Ativo`). Atualize `COLUNAS_MONETARIAS_MENSAL`.
* PDF, tabela de detalhamento mensal: colunas `Competência, Vencimento Básico, Só gratificação, Incentivo, RBM, 13º, Férias, INSS, Custo total por vaga`.
* `app.py`: o `st.dataframe` do resumo já usa `COLUNAS_MONETARIAS_RESUMO`; apenas confira que as 11 colunas aparecem.

## 5. M4: "GSTU" → "Só gratificação"

* `engine.py`: constante `ROTULO_GRATIFICACAO = "Só gratificação"`. **Use-a** em todo texto exibido (UI, PDF, Excel, instruções), em vez de literais.
* Catálogo: `adicional="GSTU"` (Analista, Técnico Universitário, Técnico de Nível Superior, Técnico de Nível Médio) → `adicional=ROTULO_GRATIFICACAO`. Motorista: `adicional="Percentual de incentivo (Percentual_Motorista). Não recebe gratificação"`. Jornalista: mantém o texto da GPC.
* Identificadores: `gstu` → `gratificacao` (variáveis, chaves de dict, docstrings). Docstring/comentário "(não recebe GSTU)" → "(não recebe gratificação)".
* **Coluna da planilha:** `Valor_GSTU` → **`Valor_Gratificacao`**. `COLUNAS_BASE` passa a ser: `Cargo, Vinculo, CH, Vencimento_Basico, Valor_Gratificacao, Percentual_Motorista, Data_Inicio_Vigencia, Data_Fim_Vigencia`.
* **Compatibilidade:** em `data_manager._mapear_colunas`, trate o cabeçalho antigo `Valor_GSTU` (comparando por `_chave_coluna("Valor_GSTU")`) como alias de `Valor_Gratificacao` e emita aviso: `Coluna 'Valor_GSTU' foi renomeada para 'Valor_Gratificacao'; atualize sua planilha.` Se as duas existirem, reaproveite o erro de coluna duplicada já existente.
* Mensagens/validações em `data_manager.py`: trocar `Valor_GSTU` por `Valor_Gratificacao` e "GSTU" por "Só gratificação" (inclui o aviso do Motorista: `Valor_Gratificacao informado para Motorista foi zerado (Motorista não recebe Só gratificação).`).
* Aba `Instruções` do template (`_linhas_instrucoes`): substitua o item do GSTU por `Valor_Gratificacao (Só gratificação): informe o valor da gratificação do cargo (GSTU para técnicos, GPC para o Jornalista); vazio/0 para docentes e para Motorista.` e gere a linha de CHs a partir do catálogo: `"Vinculo: 'Efetivo' ou 'Temporário'. CH válidas: " + ", ".join(CHs únicas do CATALOGO)` (deve listar 25h).
* `gerar_template()` deve **sempre sobrescrever** `modelos/template_tabelas_salariais.xlsx` (hoje só grava se ausente), pois o arquivo atual tem o cabeçalho antigo. `app.py`: `_previa_formatada` usa `("Vencimento_Basico", "Valor_Gratificacao")`; caption do Motorista → "O Motorista não recebe Só gratificação; aplica-se apenas o percentual de incentivo."
* `relatorios.py`: nota metodológica e rótulos usam `ROTULO_GRATIFICACAO`.

## 6. M5, M6 e M7

### M6: remover Formação Exigida
* `engine.py`: remova o campo `formacao_exigida` de `Simulacao`, a validação (mínimo de 3 caracteres) e a chave dos `metadados`.
* `app.py`: remova o `st.text_input("Formação Exigida *")`, a validação no botão (bloco `if len((formacao or "").strip()) < 3`, mantendo apenas o ramo que cria a `Simulacao` e chama `_calcular`), o argumento `formacao_exigida=` e o `st.info(f"**FORMAÇÃO EXIGIDA:** ...")`. O subheader "Detalhes da vaga" permanece (contém Observações, mês, ano e vagas).
* `relatorios.py`: remova `_tabela_formacao`, o estilo `"formacao"` e a linha `("Formação Exigida", ...)` de `_parametros`. O PDF não menciona mais formação.

### M5: logos no PDF
Arquivos já existem em `logos/`. **Renomeie** (via `mv`, uma única vez, sem sobrescrever outros):
* `logos/Logo_uefs.png` → `logos/logo_uefs.png` (brasão + "UEFS", 1056×534, **fundo cinza** RGB(230,229,230), sem transparência)
* `logos/logo_UEFS_50_anos.png` → `logos/logo_uefs_50_anos.png` (selo dos 50 anos, 528×473, transparente)
* `logos/Captura de tela de 2026-07-07 15-46-29.png` → `logos/logo_pgdp.png` (Pró-Reitoria de Gestão e Desenvolvimento de Pessoas, 737×743, fundo branco)

Implementação em `relatorios.py`:
* `_PASTA_LOGOS = Path(__file__).resolve().parent / "logos"`; constantes com os 3 nomes.
* `_carregar_logo(nome: str, altura: float) -> Image | None` (`reportlab.platypus.Image`), decorada/cacheada com `functools.lru_cache` sobre os bytes processados. Se o arquivo não existir ou falhar a leitura → devolve `None` (o PDF é gerado sem aquele logo; **nunca** levante exceção por causa de logo).
* Largura = `altura × (largura_px / altura_px)` (preserve a proporção). Altura padrão: **1,7 cm**.
* **Fundo cinza do `logo_uefs.png`:** com Pillow (`PIL.Image`, `PIL.ImageDraw.floodfill`), converta para RGB e preencha de branco a partir dos quatro cantos `(0,0)`, `(w-1,0)`, `(0,h-1)`, `(w-1,h-1)` com `value=(255,255,255)` e `thresh=14`; salve em `BytesIO` (PNG) e use no `Image`. Aplique o tratamento só a esse logo (os outros dois entram como estão; o selo 50 anos tem canal alfa, e o `Image` do ReportLab usa `mask="auto"`).
* **Cabeçalho** (substitui os dois `Paragraph` iniciais do `historia`): `Table` de 1 linha × 3 colunas, larguras `[6 cm, restante, 7 cm]`, sem grade, `VALIGN=MIDDLE`, `LINEBELOW` de 1 pt na cor `_AZUL_ESCURO` e `BOTTOMPADDING` 6:
  1. Esquerda: `logo_uefs`.
  2. Centro (Paragraphs centralizados): `Universidade Estadual de Feira de Santana – UEFS` (Helvetica-Bold 12), `Pró-Reitoria de Gestão e Desenvolvimento de Pessoas – PGDP` (Helvetica 9), `Relatório de Impacto Orçamentário – Projeção de 3 anos (Ano 1 parcial + Anos 2 e 3)` (Helvetica-Bold 11, cor `_AZUL_ESCURO`).
  3. Direita: tabela interna 1×2 alinhada à direita com `logo_pgdp` e `logo_uefs_50_anos`, espaçamento de 0,3 cm. Se algum logo for `None`, omita a célula correspondente.
* Excel e CSV **não** recebem logos.

### M7: ajustes derivados
* **PDF em paisagem:** `pagesize=landscape(A4)`; `_LARGURA_UTIL = landscape(A4)[0] - 4 * cm`; margens 2 cm. No `CanvasNumerado._desenhar_rodape` troque `largura, _ = A4` por `largura, _ = self._pagesize`.
* **Tabela "Resumo por ano" (11 colunas):** larguras: Ano 2,4 cm; Meses 1,2 cm; as 9 monetárias dividem o restante igualmente; fonte do corpo 7,5 pt; cabeçalho com `Paragraph` (quebra de linha) e `repeatRows=1`; mantenha zebra e linha TOTAL em negrito. Ajuste a fórmula de `larguras` em `_tabela_resumo` (hoje usa `2.6 cm` e `1.1 cm`).
* **Detalhamento mensal (9 colunas):** primeira coluna 2,8 cm, demais iguais; fonte 7,5 pt.
* Mantenha `PageBreak` antes do detalhamento mensal.
* **Nota metodológica:** `RBM = vencimento básico + Só gratificação + incentivo (percentual de titulação para docentes; percentual de incentivo para o Motorista), conforme o cargo. 13º salário = RBM / 12; adicional de férias = (RBM / 3) / 12; INSS patronal = 20% sobre (RBM + 13º + férias). Cada componente é arredondado ao centavo no mês e os totais somam os valores arredondados. Aplica-se a tabela salarial vigente no dia 1 de cada mês. Contratos temporários consideram prorrogação automática durante todo o período projetado. Benefícios indenizatórios (auxílio-alimentação, transporte etc.) estão excluídos.`
* `requirements.txt`: acrescente `pillow>=10.0` (mantenha as demais linhas).
* Docstrings/legendas que citavam "contrato encerrado" ou "(enc.)" devem ser removidas.

## 7. Testes e verificações

Edite `test_engine.py` (os helpers de DataFrame devem usar a coluna `Valor_Gratificacao`; ajuste as chamadas a `Simulacao` removendo `formacao_exigida`, `meses_contrato_inicial` e `meses_prorrogacao`).

**Remover:** `test_temporario_sem_contrato_inicial_gera_erro` e `test_formacao_vazia_gera_erro`.
**Substituir** `test_temporario_encerra_apos_contrato_e_prorrogacao` por `test_temporario_prorrogacao_automatica`: Técnico de Nível Médio Temporário 40h, RBM 5000, ingresso 01/2027 → 36 meses, todos com `custo_total=6666.67` e sem chave `ativo`; com ingresso 07/2027 → 30 meses, todos com custo.
**Renomear** `test_motorista_sem_gstu_*` para `test_motorista_sem_gratificacao_*` (mesmos valores: vencimento 2000, gratificação 300 ignorada, percentual 10 → `gratificacao=0.00`, `rbm=2200.00`, `custo_total=2933.33`; com 20% → `rbm=2400.00`, `custo_total=3200.00`).
**Novos testes:**
1. `test_jornalista_25h_caso_de_referencia`: valores da seção 2 (mês e totais); Resumo Ano 2027: Vencimento 21466.80, Só gratificação 13555.80, Incentivo 0.00, RBM 35022.60, Custo por Vaga 46696.80; TOTAL Custo por Vaga 140090.40.
2. `test_jornalista_ch_35h_rejeitada`: `projetar` com `ch="35h"` levanta `ValueError`; linha de planilha com CH `35h` gera erro de validação.
3. `test_composicao_rbm_fecha`: para Docente, Técnico, Motorista e Jornalista, em cada mês e em cada linha do `resumo_anual` (inclusive TOTAL): `round(vencimento + gratificacao + incentivo, 2) == rbm`.
4. `test_resumo_colunas`: `list(resumo_anual(res).columns) == list(COLUNAS_RESUMO)` e a lista tem as 11 colunas da seção 4.
5. `test_alias_valor_gstu_na_planilha`: planilha com cabeçalho `Valor_GSTU` é aceita, normalizada para `Valor_Gratificacao`, e `avisos` contém a mensagem de renomeação.
6. `test_template_novo_cabecalho`: `gerar_template()` produz aba `Tabelas` com cabeçalho `COLUNAS_BASE` (contendo `Valor_Gratificacao`, sem `Valor_GSTU`) e a aba `Instruções` cita `25h`.
7. `test_pdf_com_logos`: `gerar_pdf` retorna bytes iniciados por `%PDF`, com `pdf.count(b"/Subtype /Image") >= 3`, e o texto `Formação Exigida` **não** aparece nos parâmetros. Em outro caso, monkeypatch de `_PASTA_LOGOS` para pasta vazia → PDF ainda é gerado.
8. Ajuste `test_relatorios_gerados` (Excel: abas `Parâmetros`, `Resumo Anual`, `Detalhamento Mensal`; sem coluna `Ativo`; `Resumo Anual` com as 11 colunas).

**Verificações finais (grep; devem retornar vazio, exceto onde indicado):**
* `grep -rniE "gstu" --include=*.py .` → só pode restar o alias legado em `data_manager.py` (e seu teste).
* `grep -rniE "formacao|formação" --include=*.py .` → vazio.
* `grep -rnE "meses_contrato|meses_prorrogacao|Prorrogação \(meses\)|Contrato inicial|\"ativo\"|\['ativo'\]|OBS_ENCERRADO|enc\." --include=*.py .` → vazio.
* `python -m pytest -q` → todos passam.

## 8. Ordem de execução

`engine.py` → `data_manager.py` → `relatorios.py` → `app.py` → `test_engine.py` → `requirements.txt` → renomear logos → (apagar `modelos/template_tabelas_salariais.xlsx` se existir; `gerar_template()` o recria na primeira chamada, incluindo a dos testes) → pytest → greps.

## 9. Decisões padrão (não pergunte)

* **CH do Jornalista = 25h** (valor informado pelo solicitante). O texto antigo "35h" é descartado. Para mudar no futuro, altere só `CH_JORNALISTA`.
* **Temporário:** prorrogação automática significa cobrar **todos** os meses do período projetado (25 a 36); não há limite de contrato nem campo para isso.
* **Formação Exigida:** removida por completo (campo, validação e relatórios). Observações continuam opcionais.
* **Rótulo exato:** `Só gratificação` (S maiúsculo, resto minúsculo) em toda a interface e relatórios. Na planilha: `Valor_Gratificacao`.
* **Colunas de composição da RBM** aparecem sempre, mesmo zeradas.
* **Logos** só no PDF, em `logos/` com os nomes novos; faltou arquivo → gera sem ele, sem erro.
* **PDF em paisagem** para caber as 11 colunas do resumo.
* Valores monetários sempre `R$ 1.234,56`; arredondamento `ROUND_HALF_UP` por componente no mês (inalterado).
* Qualquer ponto não coberto: escolha a opção mais conservadora (explícita, com erro claro) e cite em uma linha na resposta final.

## 10. Critérios de aceitação

1. `python -m pytest -q` passa e as verificações da seção 7 retornam o esperado.
2. Caso de referência do Jornalista (25h) calcula `custo_total=3891.40` por mês e `140090.40` em 36 meses.
3. Cargo Temporário não mostra campos de contrato e cobra todos os meses do período.
4. "Resumo por ano" (tela, PDF e Excel) mostra Vencimento Básico, Só gratificação, Incentivo e RBM, com soma exata.
5. Nenhum texto "GSTU" visível ao usuário; planilha modelo nova com `Valor_Gratificacao`; planilha antiga com `Valor_GSTU` ainda é aceita com aviso.
6. PDF em paisagem com os três logos no cabeçalho (UEFS à esquerda; PGDP e 50 anos à direita) e sem "Formação Exigida".
