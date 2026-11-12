# ImpactoUEFS

**Sistema de Simulação e Projeção de Impacto Orçamentário de Pessoal**  
Universidade Estadual de Feira de Santana – UEFS  
Pró-Reitoria de Gestão e Desenvolvimento de Pessoas – PGDP

---

## 📌 Visão Geral

O **ImpactoUEFS** é uma aplicação web desenvolvida em **Python** e **Streamlit** para projetar com precisão o impacto orçamentário decorrente da contratação de servidores públicos (docentes e técnicos administrativos, efetivos e temporários/REDA) na UEFS.

O sistema gera projeções de 3 anos calendário (Ano 1 proporcional a partir do mês de ingresso até dezembro + Anos 2 e 3 completos), aplicando com rigor a legislação estadual e federal de remunerações e encargos previdenciários patronais.

---

## 🚀 Principais Funcionalidades

### 1. Simulação de Impacto Orçamentário
* **Catálogo Oficial:** Docentes (Professor Auxiliar, Assistente, Substituto, Visitante), Técnicos Administrativos (Analista Universitário, Técnico Universitário, Técnico de Nível Superior e Médio), Motorista e Jornalista (Lei 14.963/2025).
* **Alíquotas Patronais Diferenciadas:**
  * **Efetivos (28% – RPPS / FUNPREV e BAPREV):** Conforme art. 68 da Lei Estadual nº 11.357/2009 c/c Lei Estadual nº 14.520/2022.
  * **Temporários / REDA (20% – RGPS / INSS):** Conforme art. 40, § 13 da CF/88 c/c art. 22, I da Lei Federal nº 8.212/1991.
  * **Incidência sobre Férias:** Aplicação do Tema 985 do STF (RE 1.072.485) e art. 92 da Lei Estadual nº 6.677/1994.
* **Componentes Remuneratórios:** Vencimento básico, Gratificação (GSTU/GPC), Incentivo Docente por Titulação (Especialização 20%, Mestrado 40%, Doutorado 60%) e Incentivo do Motorista.

### 2. Base Salarial com Persistência SQLite (`db_salarios.sqlite`)
* **Persistência Nativa:** Base histórica armazenada em banco de dados SQLite local integrado via Pandas.
* **Encadeamento Temporal Automático:** Ao cadastrar um novo reajuste salarial, a vigência anterior do cargo é automaticamente encerrada no dia anterior ao início da nova tabela (`Data_Inicio_Vigencia - 1 dia`), evitando sobreposições.
* **Três Modos de Gestão da Base:**
  1. **Importação em Lote:** Upload de planilhas (`.xlsx` ou `.csv`) com validação detalhada e detecção de conflitos.
  2. **Histórico e Edição Direta:** Tabela dinâmica interativa (`st.data_editor`) para editar células, corrigir valores e auditar a linha do tempo de cada cargo.
  3. **Cadastro Manual de Vigência:** Formulário guiado pelo catálogo oficial para inclusão individual de reajustes anuais.

### 3. Emissão de Relatórios Oficiais
* **PDF Sintético Oficial (1 página A4):** Demonstrativo executivo formatado rigorosamente para página única com o brasão da UEFS, logomarcas da PGDP e UEFS 50 Anos, resumo trienal e quadro financeiro de acréscimo anual.
* **PDF Detalhado:** Relatório técnico com memória de cálculo mensal, notas metodológicas e embasamento jurídico.
* **Planilhas:** Exportação completa em **Excel (`.xlsx`)** e **CSV**.

---

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** Python 3.10+
* **Interface:** Streamlit
* **Processamento de Dados:** Pandas
* **Banco de Dados:** SQLite3 (nativo)
* **Geração de PDF:** ReportLab
* **Planilhas:** OpenPyXL
* **Testes Automatizados:** Pytest (40+ testes unitários)

---

## 📦 Como Instalar e Rodar Localmente ou no Servidor

### 1. Clonar o Repositório
```bash
git clone https://github.com/manoel-roberto/ImpactoUEFS.git
cd ImpactoUEFS
```

### 2. Criar e Ativar o Ambiente Virtual
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Instalar as Dependências
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Executar a Aplicação
```bash
streamlit run app.py
```
A aplicação estará disponível em `http://localhost:8501`.

---

## 🐳 Deploy e Execução com Docker

A aplicação está containerizada e pronta para deploy em ambientes de desenvolvimento, homologação ou produção utilizando **Docker** e **Docker Compose**.

### 1. Pré-requisitos
* [Docker](https://docs.docker.com/get-docker/) instalado.
* [Docker Compose](https://docs.docker.com/compose/) instalado (ou Docker CLI com suporte a Compose v2).

### 2. Executando com Docker Compose (Recomendado)

O arquivo `docker-compose.yml` está configurado para expor a porta `8501`, realizar a montagem de volume (garantindo a persistência do banco `db_salarios.sqlite`, pastas de logos e exportações) e reiniciar o container automaticamente:

1. **Construir e iniciar o container em segundo plano:**
   ```bash
   docker compose up -d --build
   ```

2. **Acompanhar os logs da aplicação:**
   ```bash
   docker compose logs -f
   ```

3. **Verificar o status do container e healthcheck:**
   ```bash
   docker compose ps
   ```

4. **Parar a aplicação:**
   ```bash
   docker compose down
   ```

### 3. Executando com Docker CLI (Sem Compose)

Se optar por executar diretamente via Docker:

1. **Construir a imagem:**
   ```bash
   docker build -t impactouefs:latest .
   ```

2. **Executar o container com volume mapeado:**
   ```bash
   docker run -d \
     --name impactouefs_app \
     -p 8501:8501 \
     -v "$(pwd)":/app \
     --restart always \
     impactouefs:latest
   ```

A aplicação estará disponível em `http://localhost:8501` (ou no IP do servidor de deploy).

---

## 🧪 Executando os Testes Automatizados

O projeto possui suíte de testes com cobertura para o motor de projeção, regras previdenciárias, validação de planilhas, persistência SQLite e relatórios PDF:

```bash
pytest
```

---

## 📂 Estrutura do Projeto

```text
ImpactoUEFS/
├── app.py                 # Interface Streamlit (Simulação e Gestão da Base)
├── data_manager.py        # Gestão da Base Salarial, Leitura, Validação e SQLite
├── engine.py              # Motor de cálculo e projeção orçamentária
├── relatorios.py          # Geradores de PDF (Sintético e Detalhado), Excel e CSV
├── test_engine.py         # Testes automatizados (pytest)
├── requirements.txt       # Dependências do projeto
├── Dockerfile             # Definição da imagem de containerização
├── docker-compose.yml     # Configuração e orquestração Docker Compose
├── .dockerignore          # Arquivos excluídos do contexto de build Docker
├── db_salarios.sqlite     # Banco de dados SQLite local
├── logos/                 # Logomarcas oficiais (UEFS, PGDP, UEFS 50 Anos)
└── modelos/               # Templates modelo (.xlsx e .csv)
```

---

## 🏛️ Institucional

**Universidade Estadual de Feira de Santana (UEFS)**  
**Pró-Reitoria de Gestão e Desenvolvimento de Pessoas (PGDP)**  
Feira de Santana – Bahia – Brasil

