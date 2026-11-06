# PROMPT: Sistema de Impacto Orçamentário da UEFS

## 0. Modo de operação (leia primeiro, economiza tokens)

Você é um Desenvolvedor Python Sênior (Streamlit, Pandas, Engenharia de Software). Todas as decisões de projeto abaixo já estão FECHADAS.

1. **NÃO faça perguntas.** Se algo parecer omisso, aplique a regra da seção 12 (Decisões padrão) e siga.
2. **NÃO repita nem resuma esta especificação.** NÃO explique o código fora de docstrings curtas e comentários necessários.
3. **NÃO liste, leia ou explore diretórios** antes de criar arquivos. O projeto começa vazio.
4. Escreva cada arquivo **uma única vez**, completo, sem abreviações, sem `# ...restante do código...`, sem placeholders.
5. Ordem de criação: `requirements.txt` → `engine.py` → `data_manager.py` → `relatorios.py` → `app.py` → `test_engine.py`.
6. Após criar tudo, execute **uma vez**: `pip install -r requirements.txt` e `python -m pytest -q test_engine.py`. Corrija **somente** falhas. Não rode o Streamlit de forma interativa; valide apenas com `python -c "import app"` (ou `streamlit run` por 5 s se for trivial).
7. Resposta final: **no máximo 6 linhas** (como rodar + decisões da seção 12 que tiveram de ser tomadas, se houver). Sem recapitulação de código.
8. Idioma: textos de UI, mensagens de erro, PDF e docstrings em **português do Brasil**. Identificadores de domínio em português sem acentos (`vencimento_basico`); termos técnicos genéricos podem ficar em inglês.
9. Qualidade: Python 3.10+, PEP 8 (linhas ≤ 100), type hints completos, `from __future__ import annotations`, `dataclasses`, sem variáveis globais mutáveis, sem `print` de depuração, sem `eval/exec`.

## 1. Objetivo

Sistema web (Streamlit monolítico) que calcula e projeta o **impacto orçamentário** da contratação de pessoal da Universidade Estadual de Feira de Santana (UEFS) num período de **3 anos calendário** (ano de ingresso, parcial, do mês de ingresso até 31/dez + os dois anos seguintes completos; de 25 a 36 meses no total), para Docentes e Técnicos, Efetivos e Temporários (REDA). Benefícios indenizatórios (auxílio-alimentação, transporte etc.) estão **excluídos**.

## 2. Stack e arquivos

* Python 3.10+, Streamlit, Pandas, openpyxl (xlsx), ReportLab (PDF). **Proibido** FPDF e SQLite: o estado vive em `st.session_state`.
* `requirements.txt` (exatamente): `streamlit>=1.33`, `pandas>=2.0`, `openpyxl>=3.1`, `reportlab>=4.0`, `pytest>=7.0`.

```
/impacto_orcamentario_uefs
├── app.py              # UI Streamlit: navegação, formulário, resultados, upload/download
├── engine.py           # Catálogo de cargos, dataclasses, motor de projeção mensal (sem Streamlit/IO)
├── data_manager.py     # Planilha modelo, leitura, normalização e validação da base salarial
├── relatorios.py       # Geração de PDF (ReportLab) e Excel/CSV de resultados (sem Streamlit)
├── test_engine.py      # Testes pytest do motor (casos da seção 10)
├── requirements.txt
└── /modelos
    └── template_tabelas_salariais.xlsx   # gerado por data_manager.gerar_template() se não existir
```

`relatorios.py` e `test_engine.py` são os únicos arquivos adicionais permitidos (separam PDF/Excel da UI). Dependências entre módulos (sem ciclos): `engine` ← `data_manager` ← `app`; `engine` ← `relatorios` ← `app`.

## 3. Catálogo de cargos (constante em `engine.py`)

Defina `CATALOGO: dict` (ou lista de `@dataclass(frozen=True) CargoDef`) com exatamente as linhas abaixo. A UI e a validação da planilha derivam **somente** dele.

| Categoria | Vínculo | Cargo (string exata) | Classe/Nível exibido | CHs permitidas | Incentivo/Adicional |
|---|---|---|---|---|---|
| Docentes | Efetivo | Professor Auxiliar | Nível A | 40h | Especialização, 20% |
| Docentes | Efetivo | Professor Assistente | Nível A | 40h | Mestrado, 40% |
| Docentes | Temporário | Professor Substituto (Auxiliar) | Nível A | 40h | Especialização, 20% |
| Docentes | Temporário | Professor Visitante (Adjunto) | Nível A | DE | Doutorado, 60% |
| Técnicos | Efetivo | Analista Universitário | Nível Superior, Ref. S, Grau 1 | 30h | GSTU |
| Técnicos | Efetivo | Técnico Universitário | Nível Médio, Ref. 1, Grau 1 | 30h | GSTU |
| Técnicos | Temporário | Técnico de Nível Superior | Ref. S, Grau 1 | 30h, 40h | GSTU |
| Técnicos | Temporário | Técnico de Nível Superior (Jornalista) | Ref. S, Grau 1 | 35h | GSTU (base legal própria: Lei 14.963/2025) |
| Técnicos | Temporário | Técnico de Nível Médio | Ref. 1, Grau 1 | 30h, 40h | GSTU |
| Técnicos | Temporário | Auxiliar Administrativo (Motorista) | Classe 1 | 30h, 40h | Percentual de incentivo (`Percentual_Motorista`). **NÃO recebe GSTU** |

Notas:
* "Jornalista" é um **cargo separado** (linha própria na tabela e na planilha), com CH única 35h. Na UI ele aparece como opção do cargo "Técnico de Nível Superior (Jornalista)" dentro de Técnicos/Temporário, e exibe um `st.caption` citando a Lei 14.963/2025.
* Valores de CH são strings exatas: `"30h"`, `"35h"`, `"40h"`, `"DE"`.
* Percentuais de incentivo docente ficam em um dicionário `INCENTIVO_DOCENTE = {"Especialização": 0.20, "Mestrado": 0.40, "Doutorado": 0.60}` para fácil alteração. Para cargos sem incentivo, o incentivo é `None`.
* Constantes: `ANOS_PROJECAO = 3`, `ALIQ_INSS_PATRONAL = Decimal("0.20")`. Não existe constante de 36 meses: o total de meses é calculado (seção 4.3).

## 4. Regras de cálculo (núcleo, `engine.py`)

Use **`decimal.Decimal`** internamente. Converta entradas `float` com `Decimal(str(x))`.

### 4.1 RBM do mês (usando os valores da linha de vigência do mês)

* **Docentes:** `RBM = vencimento_basico + vencimento_basico × percentual_incentivo` (o incentivo incide sobre o vencimento básico; `Valor_GSTU` é ignorado/zero).
* **Técnicos (não motorista):** `RBM = vencimento_basico + valor_gstu`.
* **Motorista (não recebe GSTU):** `RBM = vencimento_basico + vencimento_basico × (percentual_motorista / 100)`. O `Valor_GSTU` é sempre ignorado (tratado como 0) para este cargo. O percentual de incentivo vem da planilha (linha vigente, em pontos percentuais: `20` = 20%) e pode ser **sobrescrito** pelo usuário no formulário (se preenchido, o valor do formulário prevalece em todos os meses). O vencimento é o do Auxiliar Administrativo, cadastrado na linha do cargo Motorista.

### 4.2 Verbas mensais

```
decimo_terceiro = RBM / 12
ferias          = (RBM / 3) / 12
base_inss       = RBM + decimo_terceiro + ferias
inss            = base_inss × 0.20
custo_mensal    = RBM + decimo_terceiro + ferias + inss      # = RBM × 4/3
```

Cada componente é arredondado para 2 casas com `ROUND_HALF_UP` **no mês**; totais anuais/gerais são a soma dos valores já arredondados. `custo_mensal` = soma dos componentes arredondados.

### 4.3 Linha do tempo

* Entrada `mes_ingresso` (1–12) e `ano_ingresso`. O mês de ingresso conta como **mês cheio** (o dia é ignorado).
* **Período de cálculo = exatamente 3 anos calendário. NÃO existe Ano 4.**
  * **Ano 1:** do mês de ingresso até dezembro do mesmo ano (`13 − mes_ingresso` meses).
  * **Ano 2:** janeiro a dezembro do ano seguinte (12 meses). **Ano 3:** janeiro a dezembro do ano subsequente (12 meses).
  * `total_meses = (13 − mes_ingresso) + 24` (ingresso em janeiro → 36; em julho → 30; em dezembro → 25).
  * Itere `i = 1..total_meses`; mês calendário = ingresso + (i−1) meses. Campo `ano` = ano calendário do mês; `ano_projecao` = 1 + (ano − ano_ingresso) ∈ {1, 2, 3}. O 13º e as férias continuam provisionados mês a mês (1/12 e 1/36 da RBM), sem fechamento anual adicional.
  * Relatórios agrupam por `ano` com a rotulagem "Ano 1 (AAAA)", "Ano 2 (AAAA)", "Ano 3 (AAAA)".
* **Contratos temporários:** inputs `meses_contrato_inicial` (inteiro 1–36, padrão 12) e `meses_prorrogacao` (inteiro 0–36, padrão 0). Se `i > meses_contrato_inicial + meses_prorrogacao`, todos os valores do mês são **0,00** e `ativo=False`. Efetivos: sempre `ativo=True`.
* **Quantidade de vagas:** input `quantidade_vagas` (inteiro ≥ 1, padrão 1). O motor calcula por vaga; os totais multiplicam por `quantidade_vagas` (exiba ambos: "por vaga" e "total").

### 4.4 Vigência salarial

Para cada mês `i`, defina `ref = primeiro dia do mês calendário`. Selecione a linha da tabela com `Data_Inicio_Vigencia <= ref <= Data_Fim_Vigencia` (`Data_Fim` vazio/NaT = vigência aberta). Regra de reajuste no meio do mês: vale a linha vigente no **dia 1** do mês (documente em docstring; a validação emite *warning* se `Data_Inicio` não for dia 1).
* A tabela é filtrada por `(Cargo, Vinculo, CH)` **antes** do loop.
* Se nenhuma linha cobrir `ref`:
  * `manter_ultimo_valor=False` (padrão): levante `VigenciaNaoEncontradaError` com mensagem listando o primeiro mês descoberto (`MM/AAAA`). A UI exibe via `st.error`.
  * `manter_ultimo_valor=True` (checkbox "Manter último valor vigente para meses sem tabela"): use a última linha cuja vigência terminou antes de `ref` e marque `extrapolado=True` no mês. Se não houver nem isso, levante o erro.
* Se a tabela filtrada estiver vazia: levante `BaseSalarialVaziaError` ("Importe a base salarial na página 2").

### 4.5 Interface do motor

```python
@dataclass(frozen=True)
class Simulacao:
    categoria: str; vinculo: str; cargo: str; ch: str
    formacao_exigida: str
    observacao: str = "" # Campo opcional para detalhamentos do usuário
    mes_ingresso: int; ano_ingresso: int
    quantidade_vagas: int = 1
    meses_contrato_inicial: int | None = None
    meses_prorrogacao: int | None = None
    percentual_motorista: float | None = None   # sobrescreve a planilha, em %
    manter_ultimo_valor: bool = False

@dataclass
class ResultadoSimulacao:
    meses: list[dict]        # total_meses dicts (25 a 36) por vaga
    metadados: dict          # parâmetros + totais, ver abaixo

def projetar(sim: Simulacao, tabela: pd.DataFrame) -> ResultadoSimulacao: ...
def resumo_anual(res: ResultadoSimulacao) -> pd.DataFrame: ...
```

Cada dict de `meses` tem **exatamente** as chaves: `mes` (1..total_meses), `ano`, `mes_calendario` (1..12), `competencia` ("MM/AAAA"), `ano_projecao`, `vencimento_basico`, `gstu`, `incentivo`, `rbm`, `decimo_terceiro`, `ferias`, `inss`, `custo_total`, `ativo`, `extrapolado` (valores `float` com 2 casas).

`metadados`: `total_meses, categoria, vinculo, cargo, ch, nivel_classe_descricao, incentivo_descricao, formacao_exigida, observacao, mes_ingresso, ano_ingresso, quantidade_vagas, meses_contrato_inicial, meses_prorrogacao, percentual_motorista_usado, total_por_vaga, total_geral, custo_medio_mensal (média dos meses ativos), gerado_em (datetime ISO)`.

`resumo_anual` retorna DataFrame com colunas: `Ano`, `Meses`, `RBM`, `13º Salário`, `Adicional de Férias`, `INSS Patronal`, `Custo por Vaga`, `Custo Total (vagas)`, mais uma linha final `TOTAL`.

Validações no início de `projetar` (levantam `ValueError` com mensagem em PT-BR): `formacao_exigida` vazia (após `strip`, mínimo 3 caracteres); combinação categoria/vínculo/cargo/CH fora do catálogo; mês fora de 1–12; ano fora de 2000–2100; temporário sem `meses_contrato_inicial`.

## 5. `data_manager.py`

Colunas obrigatórias da planilha (nomes exatos, nesta ordem): `Cargo, Vinculo, CH, Vencimento_Basico, Valor_GSTU, Percentual_Motorista, Data_Inicio_Vigencia, Data_Fim_Vigencia`.

Funções (tipadas):
* `gerar_template() -> bytes` — xlsx com **apenas cabeçalho** (aba `Tabelas`) + aba `Instruções` com: lista de cargos/vínculos/CHs válidos (derivada do `CATALOGO`), formato de datas (`dd/mm/aaaa`), decimais, "Data_Fim vazia = vigência aberta", "Percentual_Motorista em pontos percentuais (20 = 20%), só para Motorista", "Valor_GSTU vazio/0 para docentes e para Motorista (Motorista não recebe GSTU)". Também grava em `modelos/template_tabelas_salariais.xlsx` se ausente. Ofereça também `gerar_template_csv() -> bytes` (sep `;`, `utf-8-sig`).
* `carregar_arquivo(arquivo) -> pd.DataFrame` — aceita `.xlsx` e `.csv`. CSV: detecte separador (`;` ou `,`) e codificação (tente `utf-8-sig`, depois `latin-1`). Leia tudo como `str` e normalize depois.
* `validar_e_normalizar(df) -> ResultadoValidacao` (`dataclass`: `ok: bool`, `df: pd.DataFrame | None`, `erros: list[str]`, `avisos: list[str]`). Regras:
  1. Colunas obrigatórias presentes (ignore diferenças de caixa/espaços no cabeçalho); colunas extras geram *aviso* e são descartadas.
  2. Strip em textos; `Cargo`/`Vinculo`/`CH` devem existir no `CATALOGO` na combinação exata (Vinculo ∈ {"Efetivo","Temporário"}, aceitar "Temporario" sem acento normalizando). CH `"DE"` aceito em maiúsculas; `"40H"`→`"40h"`.
  3. Números: aceitar `1.234,56`, `1234,56` e `1234.56`. `Vencimento_Basico > 0`; `Valor_GSTU ≥ 0` (vazio → 0); `Percentual_Motorista` entre 0 e 100 (vazio → 0); `Percentual_Motorista > 0` para cargo ≠ Motorista gera *aviso* e é zerado; `Valor_GSTU > 0` para cargo Motorista gera *aviso* e é zerado; Motorista com `Percentual_Motorista` vazio/0 gera *aviso* ("Motorista sem percentual de incentivo").
  4. Datas: aceitar `dd/mm/aaaa`, `aaaa-mm-dd` e serial de Excel/`datetime`. `Data_Inicio_Vigencia` obrigatória; `Data_Fim_Vigencia` opcional; se informada, `>= Data_Inicio`.
  5. **Sobreposição:** por `(Cargo, Vinculo, CH)`, ordene por início e acuse erro quando `inicio_atual <= fim_anterior` (fim vazio = `pd.Timestamp.max`). Mensagem: `Linha N: vigência sobrepõe a linha M (Cargo/Vínculo/CH)`. Numere linhas como no Excel (`índice + 2`).
  6. Todos os erros são **coletados** (não pare no primeiro) e limitados a 50 mensagens ("... e mais K erros").
  7. Retorne `df` com tipos finais: `Vencimento_Basico`, `Valor_GSTU`, `Percentual_Motorista` como `float`; datas como `datetime64[ns]` (fim vazio → `NaT`).
* `filtrar_tabela(df, cargo, vinculo, ch) -> pd.DataFrame` — filtra e ordena por `Data_Inicio_Vigencia`.

## 6. `app.py` — Interface

Config: `st.set_page_config(page_title="Impacto Orçamentário – UEFS", layout="wide")`. Navegação: `st.sidebar.radio` com "Simulação de Impacto" e "Importação de Base Salarial" (monolito; **sem** pasta `pages/`). Estado em `st.session_state`: `base_salarial` (DataFrame validado ou `None`), `resultado` (último `ResultadoSimulacao`), `nome_arquivo_base`.

### Página 2 — Importação de Base Salarial
1. `st.download_button` "Baixar planilha modelo (.xlsx)" e outro "(.csv)".
2. `st.file_uploader` (`type=["xlsx","csv"]`). Ao enviar: carregar → validar. Se `erros`: `st.error` + lista (`st.expander` com todos); **não** altere a base atual. Se ok: salve em `session_state`, `st.success` com nº de linhas e nº de cargos, mostre `avisos` em `st.warning` e `st.dataframe` de pré-visualização formatada. Novo upload válido **substitui** a base inteira (informe isso na tela).
3. Se já existir base, mostre resumo ("N linhas carregadas de <arquivo>") e botão "Limpar base".

### Página 1 — Simulação de Impacto
* Se não houver base: `st.warning` com instrução e `st.stop()` do formulário.
* **Não use `st.form`** (campos dependentes precisam atualizar na hora). Fluxo de widgets com `key`: Categoria → Vínculo (opções filtradas do catálogo) → Cargo (filtrado) → CH (filtrada; se só houver uma, `selectbox` com 1 opção).
* Campos informativos (somente leitura, via `st.text_input(disabled=True)` ou `st.caption`): Nível/Classe e Incentivo/Adicional, derivados do catálogo.
* **Formação Exigida:** `st.text_input` obrigatório com placeholder "Ex.: Doutorado em Letras". Se vazio ao calcular → `st.error` e não calcula.
* **Observações:** `st.text_area` opcional para o usuário adicionar qualquer informação relevante, que será impressa no relatório.
* Mês de ingresso (`selectbox` com nomes de meses em PT-BR) e ano de ingresso (`number_input` 2000–2100). Padrão: mês seguinte ao atual.
* `quantidade_vagas` (`number_input`, min 1, padrão 1).
* Visíveis **apenas** se aplicável: `Percentual de Incentivo do Motorista (%)` (somente cargo Motorista, que não recebe GSTU; padrão = valor da planilha, mostrado como ajuda); `meses_contrato_inicial` e `meses_prorrogacao` (somente Vínculo Temporário).
* Checkbox "Manter último valor vigente para meses sem tabela salarial" (padrão desmarcado).
* Botão "Calcular impacto" (`type="primary"`). Capture `ValueError`/`VigenciaNaoEncontradaError`/`BaseSalarialVaziaError` e exiba com `st.error`; nunca mostre traceback.
* **Resultado** (renderizado a partir de `session_state["resultado"]`, para que downloads não o apaguem):
  * Cards `st.metric` em colunas: Impacto Total (período projetado: Ano 1 parcial + Anos 2 e 3), Custo por vaga, Custo médio mensal, e um card por ano calendário (Ano 1..Ano N).
  * `st.dataframe` do `resumo_anual` (valores em formato `R$ 1.234,56` via função `formatar_brl`; mantenha numérico para exportação).
  * `st.expander` "Detalhamento mensal (período projetado)" com o DataFrame mensal; meses inativos e extrapolados sinalizados em coluna `Observação Automática` ("Contrato encerrado" / "Valor mantido da última vigência").
  * Botões de download: **PDF**, **Excel** (.xlsx) e **CSV**.
* Formatação monetária pt-BR: `formatar_brl(v) -> "R$ 1.234,56"` (separador de milhar `.`, decimal `,`), sem depender de `locale`.

## 7. `relatorios.py`

* `gerar_pdf(res: ResultadoSimulacao, resumo: pd.DataFrame) -> bytes` com ReportLab (`SimpleDocTemplate`, A4, margens 2 cm, fonte Helvetica). Estrutura, nesta ordem:
  1. Cabeçalho: "Universidade Estadual de Feira de Santana – UEFS" e título "Relatório de Impacto Orçamentário – Projeção de 3 anos (Ano 1 parcial + Anos 2 e 3)". Sem logotipos ou imagens.
  2. **Caixa destacada (tabela com fundo cinza-claro e borda) logo abaixo do título: "FORMAÇÃO EXIGIDA: <texto>"** em negrito, corpo maior.
  3. Tabela de parâmetros: Categoria, Vínculo, Cargo, Nível/Classe, Carga Horária, Incentivo/Adicional, Quantidade de vagas, Ingresso (MM/AAAA), contrato inicial e prorrogação (se temporário), percentual de motorista (se aplicável), **Observações do Usuário** (exibir se houver, garantindo o escape do texto), data/hora de geração.
  4. Tabela "Resumo por ano" (colunas do `resumo_anual`, com linha TOTAL em negrito).
  5. `PageBreak` opcional → Tabela "Detalhamento mensal" (competência, RBM, 13º, Férias, INSS, Custo total por vaga; fonte 7–8 pt; `repeatRows=1`; cabeçalho azul-escuro com texto branco, linhas zebradas).
  6. Nota metodológica curta: fórmulas (13º = RBM/12; férias = RBM/3/12; INSS patronal 20% sobre RBM+13º+férias; benefícios indenizatórios excluídos).
  7. Rodapé com "Página X de Y" (use `canvasmaker` customizado) e data de geração.
  Valores em `R$ 1.234,56`. Escape texto livre com `xml.sax.saxutils.escape` antes de usar em `Paragraph`.
* `gerar_excel(res, resumo) -> bytes` (openpyxl via `pd.ExcelWriter`): abas `Parâmetros` (incluindo Formação e Observações), `Resumo Anual`, `Detalhamento Mensal`; valores numéricos com `number_format` `'R$ #,##0.00'`, cabeçalho em negrito, larguras ajustadas, painel congelado.
* `gerar_csv(res) -> bytes` do detalhamento mensal: `sep=";"`, `decimal=","`, `encoding="utf-8-sig"`, `index=False`.

## 8. Tratamento de erros e robustez

* Exceções de domínio próprias em `engine.py`: `VigenciaNaoEncontradaError`, `BaseSalarialVaziaError` (herdam de `ValueError`).
* Nunca deixe o app quebrar com traceback ao usuário; capture na UI.
* Funções puras e sem efeitos colaterais no `engine` e `relatorios`; nenhuma chamada Streamlit fora de `app.py`.
* Arquivos enviados: não confie na extensão; erros de leitura viram `st.error("Não foi possível ler o arquivo: <motivo curto>")`.

## 9. Critérios de aceitação

1. `python -m pytest -q` passa com todos os testes da seção 10.
2. `streamlit run app.py` abre sem erro; baixar o modelo, preencher e subir uma planilha válida habilita a simulação.
3. Planilha com vigências sobrepostas é recusada com mensagem indicando as linhas.
4. Formação Exigida vazia impede o cálculo; preenchida, aparece em destaque no topo do PDF. O campo Observações é opcional e, se preenchido, é exibido nos relatórios PDF e Excel.
5. Campos de Motorista e Temporário só aparecem quando aplicáveis.
6. Valores do motor conferem com os exemplos da seção 10 (centavo a centavo).

## 10. Testes obrigatórios (`test_engine.py`)

Use dados **fictícios** construídos no próprio teste (DataFrame com as colunas da seção 5). Verifique:

1. **Efetivo, sem variação:** Técnico Universitário, 30h, `Vencimento=5000`, `GSTU=0`, ingresso 01/2027 → cada mês: `rbm=5000.00`, `decimo_terceiro=416.67`, `ferias=138.89`, `inss=1111.11`, `custo_total=6666.67`; 36 meses; `total_por_vaga` = soma dos mensais arredondados; anos calendário 2027, 2028, 2029 com 12 meses cada.
2. **Docente:** Professor Auxiliar, `Vencimento=4000` → `rbm=4800.00`, `custo_total=6400.00`.
3. **Motorista (sem GSTU):** `Vencimento=2000`, `GSTU=300` (deve ser ignorado), `Percentual=10` → `gstu=0.00`, `rbm=2200.00`, `custo_total=2933.33`. Com sobrescrita do formulário de 20% → `rbm=2400.00`, `custo_total=3200.00`.
4. **Temporário:** `meses_contrato_inicial=12`, `meses_prorrogacao=12`, base com RBM 5000 → meses 1–24 com `custo_total=6666.67`; meses 25–36 com tudo `0.00` e `ativo=False`.
5. **Vigência múltipla:** ingresso 01/2027; linha A `4000` de 01/01/2027 a 30/04/2027 e linha B `4400` de 01/05/2027 em diante → meses 1–4 com vencimento 4000, mês 5 em diante 4400.
6. **Ano 1 proporcional, sem Ano 4:** ingresso 10/2027 → 2027 com 3 meses, 2028 com 12, 2029 com 12; `total_meses=27`; nenhum mês em 2030. Ingresso 07/2027 → `total_meses=30` (6+12+12). Ingresso 01/2027 → 36. Ingresso 12/2027 → 25.
7. **Erros:** tabela sem cobertura → `VigenciaNaoEncontradaError`; com `manter_ultimo_valor=True` → não levanta e marca `extrapolado=True`; `formacao_exigida=""` → `ValueError`.
8. **Validação de planilha** (`data_manager`): sobreposição detectada; cargo fora do catálogo; `Data_Fim < Data_Inicio`; número `1.234,56` aceito.
9. **Quantidade de vagas = 3:** `total_geral == 3 × total_por_vaga`.

## 11. Arquivos de exemplo

Não crie planilhas de exemplo com valores salariais reais. O template distribuído é **vazio** (só cabeçalho + instruções). Os valores dos testes são fictícios e nunca devem ir para a UI.

## 12. Decisões padrão (aplique se surgir dúvida; não pergunte)

* Dúvida de formato/arredondamento → siga a seção 4.2.
* Porcentagens na planilha: `Percentual_Motorista` em pontos percentuais; incentivos docentes em constante do código.
* Docentes e Motorista não usam GSTU; Técnicos (exceto Motorista) não usam percentual de incentivo; Motorista usa apenas o percentual de incentivo.
* O período projetado nunca passa de 31/dez do terceiro ano calendário (sem Ano 4).
* Planilha sem linha para a combinação escolhida → erro claro, não valor padrão.
* Persistência: apenas sessão (sem banco). Reiniciar o app exige novo upload.
* Sem autenticação, sem i18n, sem tema customizado, sem logotipos.
* Ordem alfabética nas opções onde não houver ordem definida no catálogo; ordem do catálogo nas demais.
* Nomes de arquivos de exportação: `impacto_<cargo_slug>_<AAAAMMDD>.pdf|xlsx|csv`.
* Qualquer ponto realmente não coberto: escolha a opção mais conservadora (mais explícita, com erro claro) e cite-a em uma linha na resposta final.

## 13. Entrega

Crie os arquivos listados na seção 2 com código completo. Resposta final em até 6 linhas: comando de instalação, comando de execução, resultado dos testes e decisões tomadas fora da seção 12.