"""Motor de projeção do impacto orçamentário de pessoal da UEFS.

Contém o catálogo de cargos, as dataclasses de entrada e saída e o cálculo mensal.
Este módulo não depende de Streamlit e não realiza operações de entrada/saída.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from types import MappingProxyType
from typing import Any, Mapping

import pandas as pd

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
ANOS_PROJECAO: int = 3
ALIQ_PATRONAL_EFETIVO: Decimal = Decimal("0.28")
ALIQ_PATRONAL_TEMPORARIO: Decimal = Decimal("0.20")
ALIQ_INSS_PATRONAL: Decimal = ALIQ_PATRONAL_TEMPORARIO  # mantida para compatibilidade

INCENTIVO_DOCENTE: Mapping[str, float] = MappingProxyType(
    {"Especialização": 0.20, "Mestrado": 0.40, "Doutorado": 0.60}
)

CATEGORIA_DOCENTES = "Docentes"
CATEGORIA_TECNICOS = "Técnicos"
VINCULO_EFETIVO = "Efetivo"
VINCULO_TEMPORARIO = "Temporário"


def aliquota_patronal_vinculo(vinculo: str) -> Decimal:
    """Retorna a alíquota patronal: Efetivo (28% - RPPS) ou Temporário (20% - RGPS/INSS)."""
    if str(vinculo).strip() == VINCULO_EFETIVO:
        return ALIQ_PATRONAL_EFETIVO
    return ALIQ_PATRONAL_TEMPORARIO
CARGO_MOTORISTA = "Técnico Administrativo (Motorista)"
CARGO_JORNALISTA = "Técnico de Nível Superior (Jornalista)"
LEI_JORNALISTA = "Lei 14.963/2025"
CH_JORNALISTA = "25h"
CHS_JORNALISTA: tuple[str, ...] = ("25h", "35h")
ROTULO_GRATIFICACAO = "Gratificação"

SETORES: dict[str, str] = {
    "GRH": "Gerência de Recursos Humanos",
    "NSA": "Núcleo de Seleção e Alocação",
    "SPAG": "Subgerência de Pagamentos",
    "SAAP": "Subgerência de Admissão e Acompanhamento de Pessoas",
}
SIGLAS_SETORES: tuple[str, ...] = ("GRH", "NSA", "SPAG", "SAAP")
SETOR_PADRAO: str = "GRH"

COLUNAS_BASE: tuple[str, ...] = (
    "Cargo",
    "Vinculo",
    "CH",
    "Vencimento_Basico",
    "Valor_Gratificacao",
    "Percentual_Motorista",
    "Data_Inicio_Vigencia",
    "Data_Fim_Vigencia",
)

COLUNAS_RESUMO: tuple[str, ...] = (
    "Ano",
    "Meses",
    "Vencimento Básico",
    ROTULO_GRATIFICACAO,
    "Incentivo",
    "RBM",
    "13º Salário",
    "Adicional de Férias",
    "INSS Patronal",
    "Custo por Vaga",
    "Quantidade de vagas",
    "Custo Total (vagas)",
)

CHAVES_MES: tuple[str, ...] = (
    "mes",
    "ano",
    "mes_calendario",
    "competencia",
    "ano_projecao",
    "vencimento_basico",
    "gratificacao",
    "incentivo",
    "rbm",
    "decimo_terceiro",
    "ferias",
    "inss",
    "custo_total",
    "extrapolado",
)

_CENTAVO = Decimal("0.01")
_ZERO = Decimal("0")
_CEM = Decimal("100")


# ---------------------------------------------------------------------------
# Exceções de domínio
# ---------------------------------------------------------------------------
class VigenciaNaoEncontradaError(ValueError):
    """Nenhuma linha da tabela salarial cobre uma competência do período projetado."""


class BaseSalarialVaziaError(ValueError):
    """Não há linhas na base salarial para a combinação cargo/vínculo/CH escolhida."""


# ---------------------------------------------------------------------------
# Catálogo de cargos
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class CargoDef:
    """Definição de um cargo do catálogo oficial do sistema."""

    categoria: str
    vinculo: str
    cargo: str
    nivel_classe: str
    chs: tuple[str, ...]
    incentivo: str | None = None
    adicional: str = ""

    @property
    def is_docente(self) -> bool:
        """Indica se o cargo pertence à categoria Docentes."""
        return self.categoria == CATEGORIA_DOCENTES

    @property
    def is_motorista(self) -> bool:
        """Indica se o cargo é o de Motorista (não recebe gratificação)."""
        return self.cargo == CARGO_MOTORISTA

    @property
    def is_efetivo(self) -> bool:
        """Indica se o vínculo é efetivo (RPPS)."""
        return self.vinculo == VINCULO_EFETIVO

    @property
    def is_temporario(self) -> bool:
        """Indica se o vínculo é temporário (REDA)."""
        return self.vinculo == VINCULO_TEMPORARIO

    @property
    def is_jornalista(self) -> bool:
        """Indica se o cargo é o de Jornalista (base legal própria)."""
        return self.cargo == CARGO_JORNALISTA

    @property
    def percentual_incentivo(self) -> float | None:
        """Percentual de incentivo docente (fração) ou ``None`` se não houver."""
        if self.incentivo is None:
            return None
        return INCENTIVO_DOCENTE[self.incentivo]

    @property
    def incentivo_descricao(self) -> str:
        """Descrição textual do incentivo/adicional do cargo."""
        if self.incentivo is not None:
            percentual = Decimal(str(INCENTIVO_DOCENTE[self.incentivo])) * _CEM
            return f"{self.incentivo}, {percentual.quantize(Decimal('1'))}%"
        return self.adicional


CATALOGO: tuple[CargoDef, ...] = (
    CargoDef(
        CATEGORIA_DOCENTES, VINCULO_EFETIVO, "Professor Auxiliar", "Nível A", ("40h",),
        incentivo="Especialização",
    ),
    CargoDef(
        CATEGORIA_DOCENTES, VINCULO_EFETIVO, "Professor Assistente", "Nível A", ("40h",),
        incentivo="Mestrado",
    ),
    CargoDef(
        CATEGORIA_DOCENTES, VINCULO_TEMPORARIO, "Professor Substituto (Auxiliar)", "Nível A",
        ("40h",), incentivo="Especialização",
    ),
    CargoDef(
        CATEGORIA_DOCENTES, VINCULO_TEMPORARIO, "Professor Visitante (Adjunto)", "Nível A",
        ("DE",), incentivo="Doutorado",
    ),
    CargoDef(
        CATEGORIA_TECNICOS, VINCULO_EFETIVO, "Analista Universitário",
        "Nível Superior, Ref. S, Grau 1", ("30h",), adicional=ROTULO_GRATIFICACAO,
    ),
    CargoDef(
        CATEGORIA_TECNICOS, VINCULO_EFETIVO, "Técnico Universitário",
        "Nível Médio, Ref. 1, Grau 1", ("30h",), adicional=ROTULO_GRATIFICACAO,
    ),
    CargoDef(
        CATEGORIA_TECNICOS, VINCULO_TEMPORARIO, "Técnico de Nível Superior", "Ref. S, Grau 1",
        ("30h", "40h"), adicional=ROTULO_GRATIFICACAO,
    ),
    CargoDef(
        CATEGORIA_TECNICOS, VINCULO_TEMPORARIO, CARGO_JORNALISTA, "Classe I",
        chs=CHS_JORNALISTA,
        adicional=f"GRATIFICAÇÃO POR COMPETÊNCIA - GPC ({LEI_JORNALISTA})",
    ),
    CargoDef(
        CATEGORIA_TECNICOS, VINCULO_TEMPORARIO, "Técnico de Nível Médio", "Ref. 1, Grau 1",
        ("30h", "40h"), adicional=ROTULO_GRATIFICACAO,
    ),
    CargoDef(
        CATEGORIA_TECNICOS, VINCULO_TEMPORARIO, CARGO_MOTORISTA, "Classe 1", ("30h", "40h"),
        adicional="Percentual de incentivo (Percentual_Motorista). Não recebe gratificação",
    ),
)


def _unicos(valores: list[str]) -> list[str]:
    """Remove duplicados preservando a ordem de aparição."""
    vistos: list[str] = []
    for valor in valores:
        if valor not in vistos:
            vistos.append(valor)
    return vistos


def listar_categorias() -> list[str]:
    """Categorias na ordem do catálogo."""
    return _unicos([c.categoria for c in CATALOGO])


def listar_vinculos(categoria: str) -> list[str]:
    """Vínculos disponíveis para a categoria, na ordem do catálogo."""
    return _unicos([c.vinculo for c in CATALOGO if c.categoria == categoria])


def listar_cargos(categoria: str, vinculo: str) -> list[str]:
    """Cargos disponíveis para categoria e vínculo, na ordem do catálogo."""
    return _unicos(
        [c.cargo for c in CATALOGO if c.categoria == categoria and c.vinculo == vinculo]
    )


def listar_chs(categoria: str, vinculo: str, cargo: str) -> list[str]:
    """Cargas horárias permitidas para o cargo."""
    definicao = obter_cargo(categoria, vinculo, cargo)
    return list(definicao.chs) if definicao else []


def obter_cargo(categoria: str, vinculo: str, cargo: str) -> CargoDef | None:
    """Retorna a definição do cargo para a combinação informada, se existir."""
    for definicao in CATALOGO:
        if (
            definicao.categoria == categoria
            and definicao.vinculo == vinculo
            and definicao.cargo == cargo
        ):
            return definicao
    return None


def buscar_por_cargo_vinculo(cargo: str, vinculo: str) -> CargoDef | None:
    """Retorna a definição do cargo a partir do nome e do vínculo (usado na planilha)."""
    for definicao in CATALOGO:
        if definicao.cargo == cargo and definicao.vinculo == vinculo:
            return definicao
    return None


def nomes_cargos() -> list[str]:
    """Todos os nomes de cargos do catálogo."""
    return _unicos([c.cargo for c in CATALOGO])


def combinacao_valida(categoria: str, vinculo: str, cargo: str, ch: str) -> bool:
    """Indica se a combinação categoria/vínculo/cargo/CH existe no catálogo."""
    definicao = obter_cargo(categoria, vinculo, cargo)
    return definicao is not None and ch in definicao.chs


# ---------------------------------------------------------------------------
# Dataclasses de entrada e saída
# ---------------------------------------------------------------------------
@dataclass(frozen=True, kw_only=True)
class Simulacao:
    """Parâmetros de uma simulação de impacto orçamentário (argumentos nomeados)."""

    categoria: str
    vinculo: str
    cargo: str
    ch: str
    observacao: str = ""  # Campo opcional para detalhamentos do usuário
    setor: str = SETOR_PADRAO  # Setor responsável: GRH, NSA, SPAG, SAAP
    mes_ingresso: int
    ano_ingresso: int
    quantidade_vagas: int = 1
    percentual_motorista: float | None = None  # sobrescreve a planilha, em %
    manter_ultimo_valor: bool = False


@dataclass
class ResultadoSimulacao:
    """Resultado da projeção: meses por vaga e metadados com parâmetros e totais."""

    meses: list[dict[str, Any]]
    metadados: dict[str, Any]

    @property
    def is_docente(self) -> bool:
        """Indica se a simulação se refere a cargo de Docente."""
        return (
            self.metadados.get("categoria") == CATEGORIA_DOCENTES
            or str(self.metadados.get("cargo", "")).startswith("Professor")
        )


# ---------------------------------------------------------------------------
# Utilitários numéricos e de datas
# ---------------------------------------------------------------------------
def arredondar(valor: Decimal) -> Decimal:
    """Arredonda para 2 casas decimais com ROUND_HALF_UP."""
    return valor.quantize(_CENTAVO, rounding=ROUND_HALF_UP)


def para_decimal(valor: Any) -> Decimal:
    """Converte um valor numérico em ``Decimal`` via ``str``; vazio/NaN vira zero."""
    if valor is None:
        return _ZERO
    if isinstance(valor, Decimal):
        return valor
    try:
        if pd.isna(valor):
            return _ZERO
    except (TypeError, ValueError):
        pass
    try:
        return Decimal(str(valor))
    except InvalidOperation as exc:
        raise ValueError(f"Valor numérico inválido na base salarial: {valor!r}") from exc


def competencia(mes: int, ano: int) -> str:
    """Formata a competência como ``MM/AAAA``."""
    return f"{mes:02d}/{ano:04d}"


def rotulo_ano(ano_projecao: int, ano: int) -> str:
    """Rótulo do ano de projeção, por exemplo ``Ano 1 (2027)``."""
    return f"Ano {ano_projecao} ({ano})"


def calcular_total_meses(mes_ingresso: int) -> int:
    """Meses do período: Ano 1 parcial (13 − mês) + anos seguintes completos."""
    return (13 - mes_ingresso) + 12 * (ANOS_PROJECAO - 1)


def filtrar_tabela(
    tabela: pd.DataFrame | None, cargo: str, vinculo: str, ch: str
) -> pd.DataFrame:
    """Filtra a base salarial por (Cargo, Vinculo, CH) e ordena pelo início da vigência."""
    if tabela is None or tabela.empty:
        return pd.DataFrame(columns=list(COLUNAS_BASE))
    faltando = [col for col in COLUNAS_BASE if col not in tabela.columns]
    if faltando:
        raise ValueError(f"Base salarial sem as colunas obrigatórias: {', '.join(faltando)}.")
    mascara = (
        (tabela["Cargo"].astype(str).str.strip() == cargo)
        & (tabela["Vinculo"].astype(str).str.strip() == vinculo)
        & (tabela["CH"].astype(str).str.strip() == ch)
    )
    filtrada = tabela.loc[mascara, list(COLUNAS_BASE)].copy()
    for coluna in ("Data_Inicio_Vigencia", "Data_Fim_Vigencia"):
        filtrada[coluna] = pd.to_datetime(filtrada[coluna], errors="coerce").dt.normalize()
    filtrada = filtrada.dropna(subset=["Data_Inicio_Vigencia"])
    return filtrada.sort_values("Data_Inicio_Vigencia", kind="stable").reset_index(drop=True)


def selecionar_vigencia(
    filtrada: pd.DataFrame, referencia: pd.Timestamp, manter_ultimo_valor: bool
) -> tuple[pd.Series | None, bool]:
    """Seleciona a linha vigente no dia 1 do mês de referência.

    Regra de reajuste no meio do mês: vale a linha vigente no **dia 1** do mês.
    Retorna ``(linha, extrapolado)``. Se nenhuma linha cobrir a referência e
    ``manter_ultimo_valor`` for verdadeiro, usa a última linha cuja vigência terminou
    antes da referência (``extrapolado=True``). Sem linha aplicável, retorna ``(None, False)``.
    """
    if filtrada.empty:
        return None, False
    inicio = filtrada["Data_Inicio_Vigencia"]
    fim = filtrada["Data_Fim_Vigencia"]
    cobre = (inicio <= referencia) & (fim.isna() | (fim >= referencia))
    if cobre.any():
        return filtrada.loc[cobre].iloc[-1], False
    if manter_ultimo_valor:
        anteriores = filtrada.loc[fim.notna() & (fim < referencia)]
        if not anteriores.empty:
            ordenadas = anteriores.sort_values("Data_Fim_Vigencia", kind="stable")
            return ordenadas.iloc[-1], True
    return None, False


# ---------------------------------------------------------------------------
# Validação e cálculo
# ---------------------------------------------------------------------------
def _validar_simulacao(sim: Simulacao) -> CargoDef:
    """Valida os parâmetros da simulação e retorna a definição do cargo."""
    definicao = obter_cargo(sim.categoria, sim.vinculo, sim.cargo)
    if definicao is None or sim.ch not in definicao.chs:
        raise ValueError(
            "Combinação fora do catálogo: "
            f"{sim.categoria} / {sim.vinculo} / {sim.cargo} / {sim.ch}."
        )
    if not isinstance(sim.mes_ingresso, int) or not 1 <= sim.mes_ingresso <= 12:
        raise ValueError("Mês de ingresso deve estar entre 1 e 12.")
    if not isinstance(sim.ano_ingresso, int) or not 2000 <= sim.ano_ingresso <= 2100:
        raise ValueError("Ano de ingresso deve estar entre 2000 e 2100.")
    if not isinstance(sim.quantidade_vagas, int) or sim.quantidade_vagas < 1:
        raise ValueError("A quantidade de vagas deve ser um inteiro maior ou igual a 1.")
    if sim.percentual_motorista is not None and not 0 <= sim.percentual_motorista <= 100:
        raise ValueError("O percentual de incentivo do motorista deve estar entre 0 e 100.")
    return definicao


def _componentes_rbm(
    definicao: CargoDef, linha: pd.Series, percentual_motorista: Decimal | None
) -> tuple[Decimal, Decimal, Decimal, Decimal]:
    """Retorna (vencimento, gratificacao, incentivo, percentual_motorista_aplicado)."""
    vencimento = para_decimal(linha["Vencimento_Basico"])
    gratificacao = _ZERO
    incentivo = _ZERO
    perc_aplicado = _ZERO
    if definicao.is_docente:
        percentual = Decimal(str(definicao.percentual_incentivo or 0.0))
        incentivo = vencimento * percentual
    elif definicao.is_motorista:
        if percentual_motorista is not None:
            perc_aplicado = percentual_motorista
        else:
            perc_aplicado = para_decimal(linha["Percentual_Motorista"])
        incentivo = vencimento * (perc_aplicado / _CEM)
    else:
        gratificacao = para_decimal(linha["Valor_Gratificacao"])
    return (
        arredondar(vencimento),
        arredondar(gratificacao),
        arredondar(incentivo),
        perc_aplicado,
    )


def projetar(sim: Simulacao, tabela: pd.DataFrame) -> ResultadoSimulacao:
    """Projeta mês a mês o custo por vaga no período de 3 anos calendário.

    O Ano 1 vai do mês de ingresso (contado como mês cheio) até dezembro; os Anos 2 e 3
    são completos. Para cada mês usa-se a linha da tabela vigente no dia 1 do mês.
    Cada componente é arredondado no mês (ROUND_HALF_UP); totais somam valores arredondados.
    Todos os meses são cobrados para qualquer vínculo (temporários: prorrogação automática).
    """
    definicao = _validar_simulacao(sim)
    filtrada = filtrar_tabela(tabela, sim.cargo, sim.vinculo, sim.ch)
    if filtrada.empty:
        raise BaseSalarialVaziaError(
            f"Não há linhas na base salarial para {sim.cargo} / {sim.vinculo} / {sim.ch}. "
            "Importe a base salarial na página 2 (Importação de Base Salarial)."
        )

    total_meses = calcular_total_meses(sim.mes_ingresso)
    perc_form: Decimal | None = None
    if definicao.is_motorista and sim.percentual_motorista is not None:
        perc_form = Decimal(str(sim.percentual_motorista))

    meses: list[dict[str, Any]] = []
    total_por_vaga = _ZERO
    perc_usado: Decimal | None = perc_form

    aliq_patronal = aliquota_patronal_vinculo(sim.vinculo)

    for indice in range(1, total_meses + 1):
        deslocamento = (sim.mes_ingresso - 1) + (indice - 1)
        ano = sim.ano_ingresso + deslocamento // 12
        mes_cal = deslocamento % 12 + 1
        ano_projecao = 1 + (ano - sim.ano_ingresso)
        referencia = pd.Timestamp(year=ano, month=mes_cal, day=1)
        linha, extrapolado = selecionar_vigencia(filtrada, referencia, sim.manter_ultimo_valor)
        if linha is None:
            raise VigenciaNaoEncontradaError(
                "Não há tabela salarial vigente para "
                f"{sim.cargo} / {sim.vinculo} / {sim.ch} a partir da competência "
                f"{competencia(mes_cal, ano)} (primeiro mês descoberto). Atualize a base "
                "salarial ou marque 'Manter último valor vigente para meses sem tabela'."
            )

        vencimento, gratificacao, incentivo, perc_aplicado = _componentes_rbm(
            definicao, linha, perc_form
        )
        if definicao.is_motorista and perc_usado is None:
            perc_usado = perc_aplicado
        rbm_exata = vencimento + gratificacao + incentivo
        rbm = arredondar(rbm_exata)
        decimo_exato = rbm_exata / Decimal(12)
        ferias_exata = (
            _ZERO if ano_projecao == 1 else (rbm_exata / Decimal(3)) / Decimal(12)
        )
        inss_exato = (rbm_exata + decimo_exato) * aliq_patronal
        decimo = arredondar(decimo_exato)
        ferias = arredondar(ferias_exata)
        inss = arredondar(inss_exato)
        custo = rbm + decimo + ferias + inss

        total_por_vaga += custo
        meses.append(
            {
                "mes": indice,
                "ano": ano,
                "mes_calendario": mes_cal,
                "competencia": competencia(mes_cal, ano),
                "ano_projecao": ano_projecao,
                "vencimento_basico": float(vencimento),
                "gratificacao": float(gratificacao),
                "incentivo": float(incentivo),
                "rbm": float(rbm),
                "decimo_terceiro": float(decimo),
                "ferias": float(ferias),
                "inss": float(inss),
                "custo_total": float(custo),
                "extrapolado": bool(extrapolado),
            }
        )

    total_geral = total_por_vaga * Decimal(sim.quantidade_vagas)
    custo_medio = arredondar(total_por_vaga / Decimal(len(meses))) if meses else _ZERO
    metadados: dict[str, Any] = {
        "total_meses": total_meses,
        "categoria": sim.categoria,
        "vinculo": sim.vinculo,
        "cargo": sim.cargo,
        "ch": sim.ch,
        "nivel_classe_descricao": definicao.nivel_classe,
        "incentivo_descricao": definicao.incentivo_descricao,
        "observacao": (sim.observacao or "").strip(),
        "setor": getattr(sim, "setor", SETOR_PADRAO),
        "setor_nome": SETORES.get(getattr(sim, "setor", SETOR_PADRAO), ""),
        "setor_descricao": f"{getattr(sim, 'setor', SETOR_PADRAO)} - {SETORES.get(getattr(sim, 'setor', SETOR_PADRAO), '')}",
        "mes_ingresso": sim.mes_ingresso,
        "ano_ingresso": sim.ano_ingresso,
        "quantidade_vagas": sim.quantidade_vagas,
        "percentual_motorista_usado": (
            float(perc_usado) if definicao.is_motorista and perc_usado is not None else None
        ),
        "total_por_vaga": float(arredondar(total_por_vaga)),
        "total_geral": float(arredondar(total_geral)),
        "custo_medio_mensal": float(custo_medio),
        "aliquota_patronal": float(aliq_patronal),
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
    }
    return ResultadoSimulacao(meses=meses, metadados=metadados)


def resumo_anual(res: ResultadoSimulacao) -> pd.DataFrame:
    """Agrupa os meses por ano calendário e acrescenta a linha ``TOTAL``."""
    vagas = Decimal(int(res.metadados.get("quantidade_vagas", 1)))
    campos = (
        "vencimento_basico",
        "gratificacao",
        "incentivo",
        "rbm",
        "decimo_terceiro",
        "ferias",
        "inss",
        "custo_total",
    )
    grupos: dict[int, dict[str, Any]] = {}
    for mes in res.meses:
        grupo = grupos.setdefault(
            mes["ano"],
            {"ano_projecao": mes["ano_projecao"], "meses": 0, **{c: _ZERO for c in campos}},
        )
        grupo["meses"] += 1
        for campo in campos:
            grupo[campo] += Decimal(str(mes[campo]))

    linhas: list[dict[str, Any]] = []
    totais: dict[str, Any] = {"meses": 0, **{c: _ZERO for c in campos}}
    for ano, grupo in grupos.items():
        linhas.append(
            {
                "Ano": rotulo_ano(grupo["ano_projecao"], ano),
                "Meses": grupo["meses"],
                "Vencimento Básico": float(grupo["vencimento_basico"]),
                ROTULO_GRATIFICACAO: float(grupo["gratificacao"]),
                "Incentivo": float(grupo["incentivo"]),
                "RBM": float(grupo["rbm"]),
                "13º Salário": float(grupo["decimo_terceiro"]),
                "Adicional de Férias": float(grupo["ferias"]),
                "INSS Patronal": float(grupo["inss"]),
                "Custo por Vaga": float(grupo["custo_total"]),
                "Quantidade de vagas": int(vagas),
                "Custo Total (vagas)": float(grupo["custo_total"] * vagas),
            }
        )
        totais["meses"] += grupo["meses"]
        for campo in campos:
            totais[campo] += grupo[campo]
    linhas.append(
        {
            "Ano": "TOTAL",
            "Meses": totais["meses"],
            "Vencimento Básico": float(totais["vencimento_basico"]),
            ROTULO_GRATIFICACAO: float(totais["gratificacao"]),
            "Incentivo": float(totais["incentivo"]),
            "RBM": float(totais["rbm"]),
            "13º Salário": float(totais["decimo_terceiro"]),
            "Adicional de Férias": float(totais["ferias"]),
            "INSS Patronal": float(totais["inss"]),
            "Custo por Vaga": float(totais["custo_total"]),
            "Quantidade de vagas": int(vagas),
            "Custo Total (vagas)": float(totais["custo_total"] * vagas),
        }
    )
    df = pd.DataFrame(linhas, columns=list(COLUNAS_RESUMO))
    if res.is_docente:
        df = df.drop(columns=[ROTULO_GRATIFICACAO])
    else:
        df = df.drop(columns=["Incentivo"])
    return df
