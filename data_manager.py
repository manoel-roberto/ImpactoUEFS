"""Gestão da base salarial: planilha modelo, leitura, validação e persistência SQLite."""
from __future__ import annotations

import io
import math
import re
import sqlite3
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

import engine
from engine import CARGO_MOTORISTA, CATALOGO, COLUNAS_BASE, VINCULO_EFETIVO, VINCULO_TEMPORARIO

COLUNAS_OBRIGATORIAS: tuple[str, ...] = COLUNAS_BASE
COLUNAS_CHAVE_PRIMARIA: tuple[str, ...] = (
    "Cargo",
    "Vinculo",
    "CH",
    "Data_Inicio_Vigencia",
)
COLUNAS_CHAVE_CARGO: tuple[str, ...] = ("Cargo", "Vinculo", "CH")

LIMITE_MENSAGENS = 50
ABA_TABELAS = "Tabelas"
ABA_INSTRUCOES = "Instruções"
DIRETORIO_MODELOS = Path(__file__).resolve().parent / "modelos"
CAMINHO_TEMPLATE = DIRETORIO_MODELOS / "template_tabelas_salariais.xlsx"
CAMINHO_DB = Path(__file__).resolve().parent / "db_salarios.sqlite"

COLUNA_GRATIFICACAO = "Valor_Gratificacao"
COLUNA_GRATIFICACAO_LEGADA = "Valor_GSTU"
AVISO_COLUNA_LEGADA = (
    f"Coluna '{COLUNA_GRATIFICACAO_LEGADA}' foi renomeada para '{COLUNA_GRATIFICACAO}'; "
    "atualize sua planilha."
)

_FORMATOS_DATA: tuple[str, ...] = (
    "%d/%m/%Y",
    "%Y-%m-%d",
    "%Y-%m-%d %H:%M:%S",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%Y-%m-%dT%H:%M:%S",
)
_EXCEL_EPOCA = pd.Timestamp("1899-12-30")

MAPA_SQL_PARA_DF = {
    "cargo": "Cargo",
    "vinculo": "Vinculo",
    "ch": "CH",
    "vencimento_basico": "Vencimento_Basico",
    "gratificacao": COLUNA_GRATIFICACAO,
    "percentual_motorista": "Percentual_Motorista",
    "data_inicio_vigencia": "Data_Inicio_Vigencia",
    "data_fim_vigencia": "Data_Fim_Vigencia",
}

MAPA_DF_PARA_SQL = {
    "Cargo": "cargo",
    "Vinculo": "vinculo",
    "CH": "ch",
    "Vencimento_Basico": "vencimento_basico",
    COLUNA_GRATIFICACAO: "gratificacao",
    "Percentual_Motorista": "percentual_motorista",
    "Data_Inicio_Vigencia": "data_inicio_vigencia",
    "Data_Fim_Vigencia": "data_fim_vigencia",
}


class ArquivoInvalidoError(ValueError):
    """O arquivo enviado não pôde ser lido como planilha .xlsx ou .csv."""


@dataclass
class ResultadoValidacao:
    """Resultado da validação/normalização da base salarial."""

    ok: bool
    df: pd.DataFrame | None
    erros: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Banco de Dados SQLite (Persistência)
# ---------------------------------------------------------------------------
def obter_caminho_db(caminho: Path | str | None = None) -> Path:
    """Retorna o caminho do banco SQLite configurado ou padrão."""
    if caminho is None:
        return CAMINHO_DB
    return Path(caminho).resolve()


def obter_conexao(caminho: Path | str | None = None) -> sqlite3.Connection:
    """Abre conexão com o SQLite e garante que a tabela historico_salarial exista."""
    caminho_banco = obter_caminho_db(caminho)
    caminho_banco.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(caminho_banco))
    conn.row_factory = sqlite3.Row
    inicializar_banco(conn)
    return conn


def inicializar_banco(
    conn: sqlite3.Connection | None = None, caminho: Path | str | None = None
) -> sqlite3.Connection:
    """Cria a tabela historico_salarial com chave primária composta se não existir."""
    fechar_depois = False
    if conn is None:
        conn = obter_conexao(caminho)
        fechar_depois = True

    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS historico_salarial (
            cargo TEXT NOT NULL,
            vinculo TEXT NOT NULL,
            ch TEXT NOT NULL,
            vencimento_basico REAL NOT NULL,
            gratificacao REAL NOT NULL DEFAULT 0.0,
            percentual_motorista REAL NOT NULL DEFAULT 0.0,
            data_inicio_vigencia DATE NOT NULL,
            data_fim_vigencia DATE,
            PRIMARY KEY (cargo, vinculo, ch, data_inicio_vigencia)
        );
        """
    )
    conn.commit()

    if fechar_depois:
        conn.close()
    return conn


def aplicar_encadeamento_sqlite(
    conn: sqlite3.Connection | None = None, caminho: Path | str | None = None
) -> None:
    """Executa o encadeamento temporal automático em historico_salarial no SQLite.

    Agrupa os registros por (cargo, vinculo, ch) e ordena por data_inicio_vigencia.
    Atualiza a data_fim_vigencia do registro anterior para (data_inicio_vigencia - 1 dia)
    do registro subsequente.
    """
    fechar_depois = False
    if conn is None:
        conn = obter_conexao(caminho)
        fechar_depois = True

    df = pd.read_sql_query(
        "SELECT * FROM historico_salarial ORDER BY cargo, vinculo, ch, data_inicio_vigencia ASC",
        conn,
    )
    if df.empty:
        if fechar_depois:
            conn.close()
        return

    df["data_inicio_vigencia"] = pd.to_datetime(df["data_inicio_vigencia"])
    df["data_fim_vigencia"] = pd.to_datetime(df["data_fim_vigencia"])

    alteracoes: list[tuple[str, str, str, str, str]] = []
    for _, grupo_indices in df.groupby(["cargo", "vinculo", "ch"]).groups.items():
        indices = list(grupo_indices)
        for i in range(len(indices) - 1):
            idx_atual = indices[i]
            idx_proximo = indices[i + 1]
            inicio_atual = df.loc[idx_atual, "data_inicio_vigencia"]
            inicio_proximo = df.loc[idx_proximo, "data_inicio_vigencia"]

            if inicio_proximo > inicio_atual:
                fim_esperado = (inicio_proximo - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
                fim_atual = df.loc[idx_atual, "data_fim_vigencia"]
                fim_atual_str = (
                    fim_atual.strftime("%Y-%m-%d") if pd.notna(fim_atual) else None
                )

                if fim_atual_str is None or pd.to_datetime(fim_atual) >= inicio_proximo:
                    alteracoes.append(
                        (
                            fim_esperado,
                            str(df.loc[idx_atual, "cargo"]),
                            str(df.loc[idx_atual, "vinculo"]),
                            str(df.loc[idx_atual, "ch"]),
                            inicio_atual.strftime("%Y-%m-%d"),
                        )
                    )

    if alteracoes:
        cursor = conn.cursor()
        cursor.executemany(
            """
            UPDATE historico_salarial
            SET data_fim_vigencia = ?
            WHERE cargo = ? AND vinculo = ? AND ch = ? AND data_inicio_vigencia = ?
            """,
            alteracoes,
        )
        conn.commit()

    if fechar_depois:
        conn.close()


def salvar_base_sqlite(
    df: pd.DataFrame, caminho: Path | str | None = None
) -> tuple[int, int]:
    """Insere registros validados no SQLite com sobrescrita e encadeamento temporal automático.

    Retorna (quantidade_de_registros_enviados, total_de_registros_no_banco).
    """
    if df is None or df.empty:
        return 0, 0

    conn = obter_conexao(caminho)
    df_sql = df.copy()

    df_sql["Data_Inicio_Vigencia"] = pd.to_datetime(df_sql["Data_Inicio_Vigencia"]).dt.strftime("%Y-%m-%d")
    if "Data_Fim_Vigencia" in df_sql.columns:
        df_sql["Data_Fim_Vigencia"] = df_sql["Data_Fim_Vigencia"].apply(
            lambda d: pd.Timestamp(d).strftime("%Y-%m-%d") if pd.notna(d) else None
        )
    else:
        df_sql["Data_Fim_Vigencia"] = None

    registros: list[tuple[str, str, str, float, float, float, str, str | None]] = []
    for _, row in df_sql.iterrows():
        registros.append(
            (
                str(row["Cargo"]).strip(),
                str(row["Vinculo"]).strip(),
                str(row["CH"]).strip(),
                float(row["Vencimento_Basico"]),
                float(row.get(COLUNA_GRATIFICACAO, 0.0) or 0.0),
                float(row.get("Percentual_Motorista", 0.0) or 0.0),
                str(row["Data_Inicio_Vigencia"]),
                row["Data_Fim_Vigencia"],
            )
        )

    cursor = conn.cursor()
    cursor.executemany(
        """
        INSERT INTO historico_salarial (
            cargo, vinculo, ch, vencimento_basico, gratificacao,
            percentual_motorista, data_inicio_vigencia, data_fim_vigencia
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(cargo, vinculo, ch, data_inicio_vigencia) DO UPDATE SET
            vencimento_basico = excluded.vencimento_basico,
            gratificacao = excluded.gratificacao,
            percentual_motorista = excluded.percentual_motorista,
            data_fim_vigencia = excluded.data_fim_vigencia
        """,
        registros,
    )
    conn.commit()

    # Aplica o encadeamento temporal automático
    aplicar_encadeamento_sqlite(conn)

    cursor.execute("SELECT COUNT(*) FROM historico_salarial")
    total = int(cursor.fetchone()[0])
    conn.close()
    return len(registros), total


def salvar_registro_individual(
    cargo: str,
    vinculo: str,
    ch: str,
    vencimento_basico: float,
    gratificacao: float = 0.0,
    percentual_motorista: float = 0.0,
    data_inicio_vigencia: Any = None,
    data_fim_vigencia: Any = None,
    caminho: Path | str | None = None,
) -> ResultadoValidacao:
    """Valida e salva um registro salarial individual no SQLite, aplicando encadeamento temporal."""
    linha_df = pd.DataFrame(
        [
            {
                "Cargo": str(cargo).strip(),
                "Vinculo": str(vinculo).strip(),
                "CH": str(ch).strip(),
                "Vencimento_Basico": str(vencimento_basico),
                COLUNA_GRATIFICACAO: str(gratificacao),
                "Percentual_Motorista": str(percentual_motorista),
                "Data_Inicio_Vigencia": str(data_inicio_vigencia)
                if data_inicio_vigencia is not None
                else "",
                "Data_Fim_Vigencia": str(data_fim_vigencia)
                if data_fim_vigencia is not None
                else "",
            }
        ]
    )

    resultado = validar_e_normalizar(linha_df)
    if not resultado.ok or resultado.df is None:
        return resultado

    salvar_base_sqlite(resultado.df, caminho=caminho)
    return resultado


def excluir_registro_sqlite(
    cargo: str,
    vinculo: str,
    ch: str,
    data_inicio_vigencia: Any,
    caminho: Path | str | None = None,
) -> None:
    """Exclui um registro específico e reencadeia as vigências restantes no SQLite."""
    conn = obter_conexao(caminho)
    dt_str = pd.Timestamp(data_inicio_vigencia).strftime("%Y-%m-%d")
    cursor = conn.cursor()
    cursor.execute(
        """
        DELETE FROM historico_salarial
        WHERE cargo = ? AND vinculo = ? AND ch = ? AND data_inicio_vigencia = ?
        """,
        (str(cargo).strip(), str(vinculo).strip(), str(ch).strip(), dt_str),
    )
    conn.commit()

    # Reencadeia as vigências restantes do cargo
    aplicar_encadeamento_sqlite(conn)
    conn.close()


def atualizar_base_completa_sqlite(
    df: pd.DataFrame, caminho: Path | str | None = None
) -> ResultadoValidacao:
    """Valida um DataFrame completo (ex.: editado via st.data_editor) e substitui o SQLite."""
    resultado = validar_e_normalizar(df)
    if not resultado.ok or resultado.df is None:
        return resultado

    conn = obter_conexao(caminho)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historico_salarial")
    conn.commit()
    conn.close()

    salvar_base_sqlite(resultado.df, caminho=caminho)
    return resultado


def _converter_sql_para_dataframe(df_sql: pd.DataFrame) -> pd.DataFrame:
    """Converte o retorno de consulta SQL para DataFrame compatível com engine.py."""
    if df_sql.empty:
        return pd.DataFrame(columns=list(COLUNAS_OBRIGATORIAS))
    df = df_sql.rename(columns=MAPA_SQL_PARA_DF)
    for col in list(COLUNAS_OBRIGATORIAS):
        if col not in df.columns:
            df[col] = None
    df = df[list(COLUNAS_OBRIGATORIAS)].copy()
    for col in ("Vencimento_Basico", COLUNA_GRATIFICACAO, "Percentual_Motorista"):
        df[col] = df[col].astype(float)
    for col in ("Data_Inicio_Vigencia", "Data_Fim_Vigencia"):
        df[col] = pd.to_datetime(df[col], errors="coerce").astype("datetime64[ns]")
    return df


def obter_tabela_vigente(
    data_referencia: pd.Timestamp | datetime | str | None = None,
    caminho: Path | str | None = None,
) -> pd.DataFrame:
    """Retorna um DataFrame apenas com as linhas ativas na data passada (para o engine.py)."""
    conn = obter_conexao(caminho)
    if data_referencia is None:
        query = "SELECT * FROM historico_salarial ORDER BY cargo, vinculo, ch, data_inicio_vigencia ASC"
        df_sql = pd.read_sql_query(query, conn)
    else:
        ref_str = pd.Timestamp(data_referencia).strftime("%Y-%m-%d")
        query = """
            SELECT * FROM historico_salarial
            WHERE data_inicio_vigencia <= ?
              AND (data_fim_vigencia IS NULL OR data_fim_vigencia >= ?)
            ORDER BY cargo, vinculo, ch, data_inicio_vigencia ASC
        """
        df_sql = pd.read_sql_query(query, conn, params=(ref_str, ref_str))
    conn.close()
    return _converter_sql_para_dataframe(df_sql)


def obter_historico_cargo(
    cargo: str, caminho: Path | str | None = None
) -> pd.DataFrame:
    """Retorna o histórico completo de um cargo (para a interface e auditoria)."""
    conn = obter_conexao(caminho)
    query = """
        SELECT * FROM historico_salarial
        WHERE cargo = ?
        ORDER BY data_inicio_vigencia ASC, vinculo, ch
    """
    df_sql = pd.read_sql_query(query, conn, params=(cargo.strip(),))
    conn.close()
    return _converter_sql_para_dataframe(df_sql)


def listar_cargos_cadastrados(caminho: Path | str | None = None) -> list[str]:
    """Retorna a lista de cargos únicos existentes no SQLite."""
    conn = obter_conexao(caminho)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT cargo FROM historico_salarial ORDER BY cargo ASC")
    cargos = [str(linha[0]) for linha in cursor.fetchall()]
    conn.close()
    return cargos


def obter_base_completa(caminho: Path | str | None = None) -> pd.DataFrame:
    """Retorna a base salarial consolidada completa do SQLite."""
    conn = obter_conexao(caminho)
    df_sql = pd.read_sql_query(
        "SELECT * FROM historico_salarial ORDER BY cargo, vinculo, ch, data_inicio_vigencia ASC",
        conn,
    )
    conn.close()
    return _converter_sql_para_dataframe(df_sql)


def limpar_banco_sqlite(caminho: Path | str | None = None) -> None:
    """Exclui todos os registros da tabela historico_salarial."""
    conn = obter_conexao(caminho)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historico_salarial")
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# Planilha modelo
# ---------------------------------------------------------------------------
def _linhas_instrucoes() -> list[list[str]]:
    """Linhas da aba de instruções, derivadas do catálogo."""
    linhas: list[list[str]] = [
        ["Instruções de preenchimento – Base Salarial (ImpactoUEFS)"],
        [],
        ["Regras gerais"],
        ["Preencha a aba 'Tabelas' mantendo exatamente o cabeçalho e a ordem das colunas."],
        ["Colunas: " + ", ".join(COLUNAS_OBRIGATORIAS)],
        ["Datas no formato dd/mm/aaaa (também aceitos aaaa-mm-dd e datas do Excel)."],
        ["Decimais: aceitos 1.234,56 / 1234,56 / 1234.56."],
        ["Data_Fim vazia = vigência aberta."],
        ["A vigência deve começar no dia 1 do mês (o cálculo usa a linha vigente no dia 1)."],
        ["Não pode haver sobreposição de vigências para o mesmo Cargo/Vinculo/CH."],
        ["Percentual_Motorista em pontos percentuais (20 = 20%), só para Motorista."],
        [
            f"{COLUNA_GRATIFICACAO} ({engine.ROTULO_GRATIFICACAO}): informe o valor da "
            "gratificação do cargo (GSTU para técnicos, GPC para o Jornalista); vazio/0 para "
            "docentes e para Motorista."
        ],
        ["Para Motorista, informe o vencimento do Técnico Administrativo na linha do cargo."],
        ["Incentivos docentes são definidos no sistema e não devem ser informados aqui."],
        [
            "Vinculo: 'Efetivo' ou 'Temporário'. CH válidas: "
            + ", ".join(engine._unicos([ch for c in CATALOGO for ch in c.chs]))
        ],
        [],
        ["Combinações válidas (Cargo / Vinculo / CH)"],
        ["Categoria", "Vinculo", "Cargo", "CHs válidas", "Nível/Classe", "Incentivo/Adicional"],
    ]
    for definicao in CATALOGO:
        linhas.append(
            [
                definicao.categoria,
                definicao.vinculo,
                definicao.cargo,
                ", ".join(definicao.chs),
                definicao.nivel_classe,
                definicao.incentivo_descricao,
            ]
        )
    return linhas


def gerar_template() -> bytes:
    """Gera a planilha modelo (.xlsx) com apenas o cabeçalho e a aba de instruções.

    Também grava (sempre sobrescrevendo) ``modelos/template_tabelas_salariais.xlsx``.
    """
    livro = Workbook()
    aba = livro.active
    aba.title = ABA_TABELAS
    aba.append(list(COLUNAS_OBRIGATORIAS))
    preenchimento = PatternFill("solid", fgColor="1F3864")
    for indice, celula in enumerate(aba[1], start=1):
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = preenchimento
        celula.alignment = Alignment(horizontal="center")
        letra = celula.column_letter
        aba.column_dimensions[letra].width = max(16, len(COLUNAS_OBRIGATORIAS[indice - 1]) + 4)
    aba.column_dimensions["A"].width = 40
    aba.freeze_panes = "A2"

    instrucoes = livro.create_sheet(ABA_INSTRUCOES)
    for linha in _linhas_instrucoes():
        instrucoes.append(linha)
    for numero in (1, 3, 17, 18):
        for celula in instrucoes[numero]:
            celula.font = Font(bold=True)
    for letra, largura in zip("ABCDEF", (45, 14, 42, 14, 32, 60)):
        instrucoes.column_dimensions[letra].width = largura

    buffer = io.BytesIO()
    livro.save(buffer)
    dados = buffer.getvalue()
    try:
        DIRETORIO_MODELOS.mkdir(parents=True, exist_ok=True)
        CAMINHO_TEMPLATE.write_bytes(dados)
    except OSError:
        pass
    return dados


def gerar_template_csv() -> bytes:
    """Gera o modelo em CSV (apenas cabeçalho, separador ';', UTF-8 com BOM)."""
    return (";".join(COLUNAS_OBRIGATORIAS) + "\r\n").encode("utf-8-sig")


# ---------------------------------------------------------------------------
# Leitura de arquivos
# ---------------------------------------------------------------------------
def _obter_bytes(arquivo: Any) -> bytes:
    """Extrai o conteúdo binário de bytes, caminho ou objeto de arquivo."""
    if isinstance(arquivo, (bytes, bytearray)):
        return bytes(arquivo)
    if isinstance(arquivo, (str, Path)):
        try:
            return Path(arquivo).read_bytes()
        except OSError as exc:
            raise ArquivoInvalidoError(f"arquivo inacessível ({exc.strerror})") from exc
    if hasattr(arquivo, "getvalue"):
        conteudo = arquivo.getvalue()
    elif hasattr(arquivo, "read"):
        conteudo = arquivo.read()
    else:
        raise ArquivoInvalidoError("tipo de entrada não suportado")
    if isinstance(conteudo, str):
        return conteudo.encode("utf-8")
    return bytes(conteudo)


def _ler_xlsx(dados: bytes) -> pd.DataFrame:
    """Lê a aba 'Tabelas' (ou a primeira aba) de um .xlsx como texto."""
    try:
        with pd.ExcelFile(io.BytesIO(dados), engine="openpyxl") as planilha:
            nomes = planilha.sheet_names
            aba = ABA_TABELAS if ABA_TABELAS in nomes else nomes[0]
            return pd.read_excel(planilha, sheet_name=aba, dtype=str)
    except Exception as exc:  # noqa: BLE001 - qualquer falha de leitura vira erro amigável
        raise ArquivoInvalidoError(
            f"planilha Excel inválida ou corrompida ({exc.__class__.__name__})"
        ) from exc


def _ler_csv(dados: bytes) -> pd.DataFrame:
    """Lê um CSV detectando codificação (utf-8-sig/latin-1) e separador (';' ou ',')."""
    texto: str | None = None
    for codificacao in ("utf-8-sig", "latin-1"):
        try:
            texto = dados.decode(codificacao)
            break
        except UnicodeDecodeError:
            continue
    if texto is None:
        raise ArquivoInvalidoError("codificação de texto não reconhecida")
    primeira = next((linha for linha in texto.splitlines() if linha.strip()), "")
    if not primeira:
        raise ArquivoInvalidoError("arquivo vazio")
    separador = ";" if primeira.count(";") >= primeira.count(",") and ";" in primeira else ","
    try:
        return pd.read_csv(
            io.StringIO(texto),
            sep=separador,
            dtype=str,
            keep_default_na=False,
            skip_blank_lines=False,
        )
    except (pd.errors.ParserError, pd.errors.EmptyDataError, ValueError) as exc:
        raise ArquivoInvalidoError(f"CSV malformado ({exc.__class__.__name__})") from exc


def carregar_arquivo(arquivo: Any) -> pd.DataFrame:
    """Carrega .xlsx ou .csv (detectando pelo conteúdo, não pela extensão) como texto."""
    dados = _obter_bytes(arquivo)
    if not dados:
        raise ArquivoInvalidoError("arquivo vazio")
    if dados[:4] == b"PK\x03\x04":
        return _ler_xlsx(dados)
    if dados[:4] == b"\xd0\xcf\x11\xe0":
        raise ArquivoInvalidoError("formato .xls antigo não suportado; salve como .xlsx")
    if b"\x00" in dados[:4096]:
        raise ArquivoInvalidoError("conteúdo binário não reconhecido como .xlsx ou .csv")
    return _ler_csv(dados)


# ---------------------------------------------------------------------------
# Normalização e Validação
# ---------------------------------------------------------------------------
def _sem_acentos(texto: str) -> str:
    """Remove acentos de um texto."""
    decomposto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in decomposto if not unicodedata.combining(c))


def _chave_coluna(nome: Any) -> str:
    """Chave de comparação de cabeçalho (sem caixa, espaços, sublinhados e acentos)."""
    return re.sub(r"[\s_]+", "", _sem_acentos(str(nome)).lower())


def _vazio(valor: Any) -> bool:
    """Indica se o valor é nulo ou texto em branco."""
    if valor is None:
        return True
    if isinstance(valor, str):
        return valor.strip() == ""
    try:
        return bool(pd.isna(valor))
    except (TypeError, ValueError):
        return False


def _texto(valor: Any) -> str:
    """Converte para texto sem espaços nas pontas; nulos viram string vazia."""
    return "" if _vazio(valor) else str(valor).strip()


def normalizar_vinculo(valor: str) -> str:
    """Normaliza o vínculo (aceita 'Temporario' sem acento e variações de caixa)."""
    chave = _sem_acentos(valor).strip().lower()
    if chave == "efetivo":
        return VINCULO_EFETIVO
    if chave == "temporario":
        return VINCULO_TEMPORARIO
    return valor.strip()


def normalizar_ch(valor: str) -> str:
    """Normaliza a carga horária ('40H' → '40h', 'de' → 'DE')."""
    compacto = valor.strip().upper().replace(" ", "")
    if compacto == "DE":
        return "DE"
    correspondencia = re.fullmatch(r"(\d+)H", compacto)
    if correspondencia:
        return f"{int(correspondencia.group(1))}h"
    return valor.strip()


def converter_numero(valor: Any) -> float | None:
    """Converte '1.234,56', '1234,56' ou '1234.56' em float; vazio retorna ``None``."""
    if _vazio(valor):
        return None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        numero = float(valor)
    else:
        texto = str(valor).strip()
        for remover in ("R$", "%", " ", "\u00a0"):
            texto = texto.replace(remover, "")
        if not texto:
            return None
        if "," in texto and "." in texto:
            if texto.rfind(",") > texto.rfind("."):
                texto = texto.replace(".", "").replace(",", ".")
            else:
                texto = texto.replace(",", "")
        elif "," in texto:
            if texto.count(",") > 1:
                raise ValueError(texto)
            texto = texto.replace(",", ".")
        elif texto.count(".") > 1:
            texto = texto.replace(".", "")
        if not re.fullmatch(r"[+-]?\d+(\.\d+)?", texto):
            raise ValueError(texto)
        numero = float(texto)
    if not math.isfinite(numero):
        raise ValueError(str(valor))
    return numero


def converter_data(valor: Any) -> pd.Timestamp | None:
    """Converte 'dd/mm/aaaa', 'aaaa-mm-dd', serial do Excel ou datetime; vazio → ``None``."""
    if _vazio(valor):
        return None
    if isinstance(valor, (datetime, pd.Timestamp)):
        return pd.Timestamp(valor).normalize()
    texto = str(valor).strip()
    for formato in _FORMATOS_DATA:
        try:
            return pd.Timestamp(datetime.strptime(texto, formato)).normalize()
        except ValueError:
            continue
    if re.fullmatch(r"\d+(\.0+)?", texto):
        serial = int(float(texto))
        if 1 <= serial <= 2958465:
            return _EXCEL_EPOCA + pd.Timedelta(days=serial)
    raise ValueError(texto)


def _limitar(mensagens: list[str]) -> list[str]:
    """Limita a lista a 50 mensagens, acrescentando o resumo do excedente."""
    if len(mensagens) <= LIMITE_MENSAGENS:
        return mensagens
    excedente = len(mensagens) - LIMITE_MENSAGENS
    return mensagens[:LIMITE_MENSAGENS] + [f"... e mais {excedente} erros"]


def _mapear_colunas(
    df: pd.DataFrame, erros: list[str], avisos: list[str]
) -> pd.DataFrame | None:
    """Renomeia cabeçalhos para os nomes oficiais e descarta colunas extras."""
    oficiais = {_chave_coluna(c): c for c in COLUNAS_OBRIGATORIAS}
    chave_legada = _chave_coluna(COLUNA_GRATIFICACAO_LEGADA)
    renomear: dict[Any, str] = {}
    extras: list[str] = []
    for coluna in df.columns:
        chave = _chave_coluna(coluna)
        if chave == chave_legada:
            chave = _chave_coluna(COLUNA_GRATIFICACAO)
            if COLUNA_GRATIFICACAO not in renomear.values():
                avisos.append(AVISO_COLUNA_LEGADA)
        if chave in oficiais:
            if oficiais[chave] in renomear.values():
                erros.append(f"Coluna duplicada no cabeçalho: {oficiais[chave]}.")
                continue
            renomear[coluna] = oficiais[chave]
        elif not df[coluna].map(_vazio).all():
            extras.append(str(coluna))
    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in renomear.values()]
    if faltando:
        erros.append(f"Colunas obrigatórias ausentes: {', '.join(faltando)}.")
    if erros:
        return None
    if extras:
        avisos.append(f"Colunas extras ignoradas: {', '.join(extras)}.")
    resultado = df[list(renomear.keys())].rename(columns=renomear)
    return resultado[list(COLUNAS_OBRIGATORIAS)]


def _validar_linha(
    linha: pd.Series, numero: int, erros: list[str], avisos: list[str]
) -> dict[str, Any] | None:
    """Valida e normaliza uma linha; retorna o registro ou ``None`` se houver erro."""
    erros_linha: list[str] = []
    cargo = _texto(linha["Cargo"])
    vinculo = normalizar_vinculo(_texto(linha["Vinculo"]))
    ch = normalizar_ch(_texto(linha["CH"]))

    if not cargo:
        erros_linha.append(f"Linha {numero}: Cargo não informado.")
    elif cargo not in engine.nomes_cargos():
        erros_linha.append(f"Linha {numero}: cargo '{cargo}' fora do catálogo.")
    else:
        definicao = engine.buscar_por_cargo_vinculo(cargo, vinculo)
        if definicao is None:
            erros_linha.append(
                f"Linha {numero}: vínculo '{vinculo}' inválido para o cargo '{cargo}'."
            )
        elif ch not in definicao.chs:
            erros_linha.append(
                f"Linha {numero}: CH '{ch}' inválida para '{cargo}'; "
                f"permitida(s): {', '.join(definicao.chs)}."
            )

    valores: dict[str, float | None] = {}
    for coluna in ("Vencimento_Basico", COLUNA_GRATIFICACAO, "Percentual_Motorista"):
        try:
            valores[coluna] = converter_numero(linha[coluna])
        except ValueError:
            valores[coluna] = None
            erros_linha.append(
                f"Linha {numero}: valor numérico inválido em {coluna} "
                f"('{_texto(linha[coluna])}')."
            )
    vencimento = valores["Vencimento_Basico"]
    gratificacao = valores[COLUNA_GRATIFICACAO] or 0.0
    percentual = valores["Percentual_Motorista"] or 0.0
    if vencimento is None:
        if not _texto(linha["Vencimento_Basico"]):
            erros_linha.append(f"Linha {numero}: Vencimento_Basico não informado.")
    elif vencimento <= 0:
        erros_linha.append(f"Linha {numero}: Vencimento_Basico deve ser maior que zero.")
    if gratificacao < 0:
        erros_linha.append(f"Linha {numero}: {COLUNA_GRATIFICACAO} não pode ser negativo.")
    if not 0 <= percentual <= 100:
        erros_linha.append(f"Linha {numero}: Percentual_Motorista deve estar entre 0 e 100.")

    if cargo == CARGO_MOTORISTA:
        if gratificacao > 0:
            avisos.append(
                f"Linha {numero}: {COLUNA_GRATIFICACAO} informado para Motorista foi zerado "
                f"(Motorista não recebe {engine.ROTULO_GRATIFICACAO})."
            )
            gratificacao = 0.0
        if percentual == 0:
            avisos.append(f"Linha {numero}: Motorista sem percentual de incentivo.")
    elif percentual > 0:
        avisos.append(
            f"Linha {numero}: Percentual_Motorista informado para '{cargo}' foi zerado "
            "(aplica-se apenas ao Motorista)."
        )
        percentual = 0.0

    inicio: pd.Timestamp | None = None
    fim: pd.Timestamp | None = None
    try:
        inicio = converter_data(linha["Data_Inicio_Vigencia"])
        if inicio is None:
            erros_linha.append(f"Linha {numero}: Data_Inicio_Vigencia é obrigatória.")
    except ValueError:
        erros_linha.append(
            f"Linha {numero}: Data_Inicio_Vigencia inválida "
            f"('{_texto(linha['Data_Inicio_Vigencia'])}'); use dd/mm/aaaa."
        )
    try:
        fim = converter_data(linha["Data_Fim_Vigencia"])
    except ValueError:
        erros_linha.append(
            f"Linha {numero}: Data_Fim_Vigencia inválida "
            f"('{_texto(linha['Data_Fim_Vigencia'])}'); use dd/mm/aaaa."
        )
    if inicio is not None and fim is not None and fim < inicio:
        erros_linha.append(
            f"Linha {numero}: Data_Fim_Vigencia anterior à Data_Inicio_Vigencia."
        )
    if inicio is not None and inicio.day != 1:
        avisos.append(
            f"Linha {numero}: Data_Inicio_Vigencia não é dia 1; o cálculo aplica a linha "
            "vigente no dia 1 de cada mês."
        )

    if erros_linha:
        erros.extend(erros_linha)
        return None
    return {
        "_linha": numero,
        "Cargo": cargo,
        "Vinculo": vinculo,
        "CH": ch,
        "Vencimento_Basico": float(vencimento or 0.0),
        COLUNA_GRATIFICACAO: float(gratificacao),
        "Percentual_Motorista": float(percentual),
        "Data_Inicio_Vigencia": inicio,
        "Data_Fim_Vigencia": fim,
    }


def _verificar_sobreposicao(registros: list[dict[str, Any]], erros: list[str]) -> None:
    """Acusa vigências sobrepostas por (Cargo, Vinculo, CH) dentro do mesmo lote enviado."""
    grupos: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for registro in registros:
        chave = (registro["Cargo"], registro["Vinculo"], registro["CH"])
        grupos.setdefault(chave, []).append(registro)
    for (cargo, vinculo, ch), itens in grupos.items():
        ordenados = sorted(itens, key=lambda r: (r["Data_Inicio_Vigencia"], r["_linha"]))
        fim_maximo: pd.Timestamp | None = None
        linha_maxima = 0
        for registro in ordenados:
            inicio = registro["Data_Inicio_Vigencia"]
            if fim_maximo is not None and inicio <= fim_maximo:
                erros.append(
                    f"Linha {registro['_linha']}: vigência sobrepõe a linha {linha_maxima} "
                    f"({cargo}/{vinculo}/{ch})"
                )
            fim = registro["Data_Fim_Vigencia"]
            fim_efetivo = fim if fim is not None else pd.Timestamp.max
            if fim_maximo is None or fim_efetivo > fim_maximo:
                fim_maximo = fim_efetivo
                linha_maxima = registro["_linha"]


def validar_e_normalizar(df: pd.DataFrame) -> ResultadoValidacao:
    """Valida a base salarial, coletando todos os erros e avisos, e normaliza os tipos."""
    erros: list[str] = []
    avisos: list[str] = []
    if df is None or len(df.columns) == 0:
        return ResultadoValidacao(False, None, ["A planilha está vazia."], [])

    mapeado = _mapear_colunas(df, erros, avisos)
    if mapeado is None:
        return ResultadoValidacao(False, None, _limitar(erros), _limitar(avisos))

    registros: list[dict[str, Any]] = []
    linhas_com_dados = 0
    for posicao in range(len(mapeado)):
        linha = mapeado.iloc[posicao]
        if all(_vazio(v) for v in linha.tolist()):
            continue
        linhas_com_dados += 1
        registro = _validar_linha(linha, posicao + 2, erros, avisos)
        if registro is not None:
            registros.append(registro)

    if linhas_com_dados == 0:
        erros.append("A planilha não contém linhas de dados.")
    _verificar_sobreposicao(registros, erros)

    if erros:
        return ResultadoValidacao(False, None, _limitar(erros), _limitar(avisos))

    final = pd.DataFrame(registros).drop(columns=["_linha"])
    final = final[list(COLUNAS_OBRIGATORIAS)].reset_index(drop=True)
    for coluna in ("Vencimento_Basico", COLUNA_GRATIFICACAO, "Percentual_Motorista"):
        final[coluna] = final[coluna].astype(float)
    for coluna in ("Data_Inicio_Vigencia", "Data_Fim_Vigencia"):
        final[coluna] = pd.to_datetime(final[coluna], errors="coerce").astype("datetime64[ns]")
    return ResultadoValidacao(True, final, [], _limitar(avisos))


# ---------------------------------------------------------------------------
# Funções Auxiliares de Memória (compatibilidade com rotinas de teste)
# ---------------------------------------------------------------------------
def encadear_vigencias(df: pd.DataFrame) -> pd.DataFrame:
    """Garante o encadeamento temporal contínuo das vigências para cada (Cargo, Vinculo, CH)."""
    if df is None or df.empty:
        return df.copy() if df is not None else pd.DataFrame(columns=list(COLUNAS_OBRIGATORIAS))

    resultado = df.copy()
    resultado["Data_Inicio_Vigencia"] = pd.to_datetime(resultado["Data_Inicio_Vigencia"]).dt.normalize()
    resultado["Data_Fim_Vigencia"] = pd.to_datetime(resultado["Data_Fim_Vigencia"]).dt.normalize()

    resultado = resultado.sort_values(
        by=["Cargo", "Vinculo", "CH", "Data_Inicio_Vigencia"],
        ascending=[True, True, True, True],
    ).reset_index(drop=True)

    for _, grupo_indices in resultado.groupby(["Cargo", "Vinculo", "CH"]).groups.items():
        indices = list(grupo_indices)
        for i in range(len(indices) - 1):
            idx_atual = indices[i]
            idx_proximo = indices[i + 1]
            inicio_atual = resultado.loc[idx_atual, "Data_Inicio_Vigencia"]
            inicio_proximo = resultado.loc[idx_proximo, "Data_Inicio_Vigencia"]

            if inicio_proximo > inicio_atual:
                fim_encadeado = inicio_proximo - pd.Timedelta(days=1)
                fim_atual = resultado.loc[idx_atual, "Data_Fim_Vigencia"]
                if pd.isna(fim_atual) or fim_atual >= inicio_proximo:
                    resultado.loc[idx_atual, "Data_Fim_Vigencia"] = fim_encadeado

    return resultado


def identificar_conflitos(
    base_existente: pd.DataFrame | None, nova_base: pd.DataFrame
) -> pd.DataFrame:
    """Identifica linhas da nova_base com mesma chave primária da base existente."""
    if (
        base_existente is None
        or base_existente.empty
        or nova_base is None
        or nova_base.empty
    ):
        return pd.DataFrame(columns=list(COLUNAS_OBRIGATORIAS))

    base_norm = base_existente.copy()
    nova_norm = nova_base.copy()
    base_norm["Data_Inicio_Vigencia"] = pd.to_datetime(
        base_norm["Data_Inicio_Vigencia"]
    ).dt.normalize()
    nova_norm["Data_Inicio_Vigencia"] = pd.to_datetime(
        nova_norm["Data_Inicio_Vigencia"]
    ).dt.normalize()

    chaves_existentes = set(
        zip(
            base_norm["Cargo"],
            base_norm["Vinculo"],
            base_norm["CH"],
            base_norm["Data_Inicio_Vigencia"],
        )
    )

    mascara_conflito = nova_norm.apply(
        lambda r: (r["Cargo"], r["Vinculo"], r["CH"], r["Data_Inicio_Vigencia"])
        in chaves_existentes,
        axis=1,
    )
    return nova_norm[mascara_conflito].copy().reset_index(drop=True)


def detalhar_conflitos(
    base_existente: pd.DataFrame | None, nova_base: pd.DataFrame
) -> pd.DataFrame:
    """Monta tabela comparativa detalhada entre os registros existentes e novos em conflito."""
    conflitos = identificar_conflitos(base_existente, nova_base)
    if conflitos.empty or base_existente is None:
        return pd.DataFrame()

    base_norm = base_existente.copy()
    base_norm["Data_Inicio_Vigencia"] = pd.to_datetime(
        base_norm["Data_Inicio_Vigencia"]
    ).dt.normalize()
    nova_norm = conflitos.copy()
    nova_norm["Data_Inicio_Vigencia"] = pd.to_datetime(
        nova_norm["Data_Inicio_Vigencia"]
    ).dt.normalize()

    colunas_chave = list(COLUNAS_CHAVE_PRIMARIA)
    return pd.merge(
        base_norm,
        nova_norm,
        on=colunas_chave,
        suffixes=("_Atual", "_Novo"),
        how="inner",
    )


def mesclar_bases(
    base_existente: pd.DataFrame | None,
    nova_base: pd.DataFrame,
    sobrescrever_conflitos: bool = True,
) -> pd.DataFrame:
    """Mescla a nova_base à base_existente em memória e aplica encadeamento temporal."""
    if nova_base is None or nova_base.empty:
        return (
            base_existente.copy()
            if base_existente is not None
            else pd.DataFrame(columns=list(COLUNAS_OBRIGATORIAS))
        )

    nova_limpa = nova_base.copy()
    nova_limpa["Data_Inicio_Vigencia"] = pd.to_datetime(
        nova_limpa["Data_Inicio_Vigencia"]
    ).dt.normalize()
    if "Data_Fim_Vigencia" in nova_limpa.columns:
        nova_limpa["Data_Fim_Vigencia"] = pd.to_datetime(
            nova_limpa["Data_Fim_Vigencia"]
        ).dt.normalize()

    if base_existente is None or base_existente.empty:
        consolidada = nova_limpa
    else:
        base_limpa = base_existente.copy()
        base_limpa["Data_Inicio_Vigencia"] = pd.to_datetime(
            base_limpa["Data_Inicio_Vigencia"]
        ).dt.normalize()
        if "Data_Fim_Vigencia" in base_limpa.columns:
            base_limpa["Data_Fim_Vigencia"] = pd.to_datetime(
                base_limpa["Data_Fim_Vigencia"]
            ).dt.normalize()

        chaves_novas = set(
            zip(
                nova_limpa["Cargo"],
                nova_limpa["Vinculo"],
                nova_limpa["CH"],
                nova_limpa["Data_Inicio_Vigencia"],
            )
        )

        if sobrescrever_conflitos:
            base_restante = base_limpa[
                ~base_limpa.apply(
                    lambda r: (r["Cargo"], r["Vinculo"], r["CH"], r["Data_Inicio_Vigencia"])
                    in chaves_novas,
                    axis=1,
                )
            ]
            consolidada = pd.concat([base_restante, nova_limpa], ignore_index=True)
        else:
            chaves_existentes = set(
                zip(
                    base_limpa["Cargo"],
                    base_limpa["Vinculo"],
                    base_limpa["CH"],
                    base_limpa["Data_Inicio_Vigencia"],
                )
            )
            novas_filtradas = nova_limpa[
                ~nova_limpa.apply(
                    lambda r: (r["Cargo"], r["Vinculo"], r["CH"], r["Data_Inicio_Vigencia"])
                    in chaves_existentes,
                    axis=1,
                )
            ]
            consolidada = pd.concat([base_limpa, novas_filtradas], ignore_index=True)

    consolidada = encadear_vigencias(consolidada)
    consolidada = consolidada.sort_values(
        by=["Cargo", "Vinculo", "CH", "Data_Inicio_Vigencia"],
        ascending=[True, True, True, True],
    ).reset_index(drop=True)

    for col in ("Vencimento_Basico", COLUNA_GRATIFICACAO, "Percentual_Motorista"):
        if col in consolidada.columns:
            consolidada[col] = consolidada[col].astype(float)
    for col in ("Data_Inicio_Vigencia", "Data_Fim_Vigencia"):
        if col in consolidada.columns:
            consolidada[col] = pd.to_datetime(
                consolidada[col], errors="coerce"
            ).astype("datetime64[ns]")

    return consolidada[list(COLUNAS_OBRIGATORIAS)]


def filtrar_tabela(df: pd.DataFrame, cargo: str, vinculo: str, ch: str) -> pd.DataFrame:
    """Filtra a base por (Cargo, Vinculo, CH) e ordena por Data_Inicio_Vigencia."""
    return engine.filtrar_tabela(df, cargo, vinculo, ch)
