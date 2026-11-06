"""Geração dos relatórios de resultado: PDF (ReportLab), Excel e CSV. Sem Streamlit."""
from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal
from functools import lru_cache
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image as PILImage
from PIL import ImageDraw
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from engine import (
    COLUNAS_RESUMO,
    ROTULO_GRATIFICACAO,
    VINCULO_TEMPORARIO,
    ResultadoSimulacao,
    competencia,
    rotulo_ano,
)

FORMATO_MOEDA_EXCEL = "R$ #,##0.00"
COLUNAS_MONETARIAS_RESUMO: tuple[str, ...] = tuple(
    c for c in COLUNAS_RESUMO[2:] if c != "Quantidade de vagas"
)
COLUNAS_MONETARIAS_MENSAL: tuple[str, ...] = (
    "Vencimento Básico",
    ROTULO_GRATIFICACAO,
    "Incentivo",
    "RBM",
    "13º Salário",
    "Adicional de Férias",
    "INSS Patronal",
    "Custo Total (por vaga)",
)
OBS_EXTRAPOLADO = "Valor mantido da última vigência"

_AZUL_ESCURO = colors.HexColor("#1F3864")
_CINZA_CLARO = colors.HexColor("#E7E6E6")
_ZEBRA = colors.HexColor("#EEF2F7")
_PAGINA = landscape(A4)
_LARGURA_UTIL = _PAGINA[0] - 4 * cm

_PASTA_LOGOS = Path(__file__).resolve().parent / "logos"
LOGO_UEFS = "logo_uefs.png"
LOGO_UEFS_50_ANOS = "logo_uefs_50_anos.png"
LOGO_PGDP = "logo_pgdp.png"
_ALTURA_LOGO = 1.7 * cm


# ---------------------------------------------------------------------------
# Formatação
# ---------------------------------------------------------------------------
def formatar_brl(valor: Any) -> str:
    """Formata número como ``R$ 1.234,56`` sem depender de ``locale``."""
    if valor is None:
        return "R$ 0,00"
    try:
        if pd.isna(valor):
            return "R$ 0,00"
    except (TypeError, ValueError):
        pass
    decimal = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    sinal = "-" if decimal < 0 else ""
    inteiro, centavos = f"{abs(decimal):.2f}".split(".")
    milhares = f"{int(inteiro):,}".replace(",", ".")
    return f"{sinal}R$ {milhares},{centavos}"


def formatar_percentual(valor: float | None) -> str:
    """Formata percentual em pontos (``10`` → ``10,00%``)."""
    if valor is None:
        return "—"
    decimal = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{decimal:.2f}".replace(".", ",") + "%"


def formatar_colunas_brl(df: pd.DataFrame, colunas: tuple[str, ...]) -> pd.DataFrame:
    """Cópia do DataFrame com as colunas monetárias convertidas em texto ``R$``."""
    copia = df.copy()
    for coluna in colunas:
        if coluna in copia.columns:
            copia[coluna] = copia[coluna].map(formatar_brl)
    return copia


def _data_geracao(res: ResultadoSimulacao) -> datetime:
    """Data/hora de geração a partir dos metadados."""
    try:
        return datetime.fromisoformat(str(res.metadados.get("gerado_em")))
    except ValueError:
        return datetime.now()


def nome_arquivo(res: ResultadoSimulacao, extensao: str) -> str:
    """Nome de exportação ``impacto_<cargo_slug>_<AAAAMMDD>.<extensao>``."""
    cargo = unicodedata.normalize("NFKD", str(res.metadados.get("cargo", "cargo")))
    cargo = "".join(c for c in cargo if not unicodedata.combining(c)).lower()
    slug = re.sub(r"[^a-z0-9]+", "_", cargo).strip("_") or "cargo"
    return f"impacto_{slug}_{_data_geracao(res):%Y%m%d}.{extensao}"


def periodo_projetado(res: ResultadoSimulacao) -> str:
    """Descrição do período projetado, ex.: ``01/2027 a 12/2029 (36 meses)``."""
    if not res.meses:
        return "—"
    inicio = res.meses[0]["competencia"]
    fim = res.meses[-1]["competencia"]
    return f"{inicio} a {fim} ({res.metadados.get('total_meses', len(res.meses))} meses)"


def tabela_mensal(res: ResultadoSimulacao) -> pd.DataFrame:
    """Detalhamento mensal por vaga com a coluna ``Observação Automática``."""
    linhas: list[dict[str, Any]] = []
    for mes in res.meses:
        observacao = OBS_EXTRAPOLADO if mes["extrapolado"] else ""
        item: dict[str, Any] = {
            "Mês": mes["mes"],
            "Competência": mes["competencia"],
            "Ano": rotulo_ano(mes["ano_projecao"], mes["ano"]),
            "Vencimento Básico": mes["vencimento_basico"],
        }
        if not res.is_docente:
            item[ROTULO_GRATIFICACAO] = mes["gratificacao"]
        item.update(
            {
                "Incentivo": mes["incentivo"],
                "RBM": mes["rbm"],
                "13º Salário": mes["decimo_terceiro"],
                "Adicional de Férias": mes["ferias"],
                "INSS Patronal": mes["inss"],
                "Custo Total (por vaga)": mes["custo_total"],
                "Observação Automática": observacao,
            }
        )
        linhas.append(item)
    return pd.DataFrame(linhas)


def _parametros(res: ResultadoSimulacao, incluir_totais: bool) -> list[tuple[str, Any, bool]]:
    """Lista de parâmetros ``(rótulo, valor, monetário)`` para os relatórios."""
    md = res.metadados
    itens: list[tuple[str, Any, bool]] = [
        ("Categoria", md["categoria"], False),
        ("Vínculo", md["vinculo"], False),
        ("Cargo", md["cargo"], False),
        ("Nível/Classe", md["nivel_classe_descricao"], False),
        ("Carga Horária", md["ch"], False),
        ("Incentivo/Adicional", md["incentivo_descricao"], False),
        ("Quantidade de vagas", md["quantidade_vagas"], False),
        ("Ingresso (MM/AAAA)", competencia(md["mes_ingresso"], md["ano_ingresso"]), False),
        ("Período projetado", periodo_projetado(res), False),
    ]
    if md.get("percentual_motorista_usado") is not None:
        itens.append(
            (
                "Percentual de incentivo do Motorista",
                formatar_percentual(md["percentual_motorista_usado"]),
                False,
            )
        )
    if incluir_totais:
        itens.append(("Custo total por vaga", md["total_por_vaga"], True))
        itens.append(("Impacto total (todas as vagas)", md["total_geral"], True))
        itens.append(("Custo médio mensal por vaga", md["custo_medio_mensal"], True))
    if md.get("observacao"):
        itens.append(("Observações do Usuário", md["observacao"], False))
    itens.append(("Gerado em", f"{_data_geracao(res):%d/%m/%Y %H:%M}", False))
    return itens


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
def _criar_canvas(rodape: str) -> type[rl_canvas.Canvas]:
    """Cria a classe de canvas que escreve 'Página X de Y' e a data de geração."""

    class CanvasNumerado(rl_canvas.Canvas):
        """Canvas que adia a renderização das páginas para conhecer o total."""

        def __init__(self, *args: Any, **kwargs: Any) -> None:
            super().__init__(*args, **kwargs)
            self._estados_paginas: list[dict[str, Any]] = []

        def showPage(self) -> None:  # noqa: N802 - API do ReportLab
            self._estados_paginas.append(dict(self.__dict__))
            self._startPage()

        def save(self) -> None:
            total = len(self._estados_paginas)
            for estado in self._estados_paginas:
                self.__dict__.update(estado)
                self._desenhar_rodape(total)
                super().showPage()
            super().save()

        def _desenhar_rodape(self, total: int) -> None:
            largura, _ = self._pagesize
            self.saveState()
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.grey)
            self.setStrokeColor(colors.lightgrey)
            self.line(2 * cm, 1.6 * cm, largura - 2 * cm, 1.6 * cm)
            self.drawString(2 * cm, 1.2 * cm, rodape)
            self.drawRightString(
                largura - 2 * cm, 1.2 * cm, f"Página {self._pageNumber} de {total}"
            )
            self.restoreState()

    return CanvasNumerado


def _estilos() -> dict[str, ParagraphStyle]:
    """Estilos de parágrafo baseados em Helvetica."""
    return {
        "instituicao": ParagraphStyle(
            "instituicao", fontName="Helvetica-Bold", fontSize=12, leading=15,
            alignment=TA_CENTER,
        ),
        "pro_reitoria": ParagraphStyle(
            "pro_reitoria", fontName="Helvetica", fontSize=9, leading=12,
            alignment=TA_CENTER,
        ),
        "titulo": ParagraphStyle(
            "titulo", fontName="Helvetica-Bold", fontSize=13, leading=16,
            alignment=TA_CENTER, spaceBefore=3, textColor=_AZUL_ESCURO,
        ),
        "secao": ParagraphStyle(
            "secao", fontName="Helvetica-Bold", fontSize=11, leading=14,
            spaceBefore=12, spaceAfter=6, textColor=_AZUL_ESCURO,
        ),
        "normal": ParagraphStyle("normal", fontName="Helvetica", fontSize=9, leading=12),
        "rotulo": ParagraphStyle(
            "rotulo", fontName="Helvetica-Bold", fontSize=9, leading=12,
        ),
        "cabecalho": ParagraphStyle(
            "cabecalho", fontName="Helvetica-Bold", fontSize=7, leading=9,
            alignment=TA_CENTER, textColor=colors.white,
        ),
        "nota": ParagraphStyle("nota", fontName="Helvetica", fontSize=8, leading=11),
    }


def _texto_paragrafo(valor: Any) -> str:
    """Escapa texto livre para uso em ``Paragraph`` preservando quebras de linha."""
    return escape(str(valor)).replace("\r\n", "\n").replace("\n", "<br/>")


@lru_cache(maxsize=16)
def _bytes_logo(caminho: str, remover_fundo: bool) -> tuple[bytes, int, int] | None:
    """Lê e processa o logo, devolvendo ``(png, largura_px, altura_px)`` ou ``None``.

    Com ``remover_fundo``, preenche de branco o fundo a partir dos quatro cantos.
    """
    try:
        with PILImage.open(caminho) as original:
            original.load()
            largura, altura = original.size
            if remover_fundo:
                imagem = original.convert("RGB")
                for ponto in ((0, 0), (largura - 1, 0), (0, altura - 1),
                              (largura - 1, altura - 1)):
                    ImageDraw.floodfill(imagem, ponto, value=(255, 255, 255), thresh=14)
            else:
                imagem = original.copy()
        saida = io.BytesIO()
        imagem.save(saida, format="PNG")
        return saida.getvalue(), largura, altura
    except (OSError, ValueError):
        return None


def _carregar_logo(nome: str, altura: float = _ALTURA_LOGO) -> Image | None:
    """Cria o ``Image`` do logo preservando a proporção; ``None`` se indisponível."""
    caminho = _PASTA_LOGOS / nome
    if not caminho.is_file():
        return None
    dados = _bytes_logo(str(caminho), nome == LOGO_UEFS)
    if dados is None:
        return None
    png, largura_px, altura_px = dados
    if altura_px <= 0:
        return None
    try:
        return Image(
            io.BytesIO(png),
            width=altura * (largura_px / altura_px),
            height=altura,
            mask="auto",
        )
    except Exception:  # noqa: BLE001 - logo nunca impede a geração do PDF
        return None


def _cabecalho(estilos: dict[str, ParagraphStyle]) -> Table:
    """Cabeçalho do PDF: logo UEFS, textos institucionais e logos PGDP/50 anos."""
    textos = [
        Paragraph("Universidade Estadual de Feira de Santana – UEFS", estilos["instituicao"]),
        Paragraph(
            "Pró-Reitoria de Gestão e Desenvolvimento de Pessoas – PGDP",
            estilos["pro_reitoria"],
        ),
        Paragraph(
            "Relatório de Impacto Orçamentário",
            estilos["titulo"],
        ),
    ]
    logo_esquerda: Any = _carregar_logo(LOGO_UEFS) or ""
    logos_direita = [
        logo for logo in (_carregar_logo(LOGO_PGDP), _carregar_logo(LOGO_UEFS_50_ANOS))
        if logo is not None
    ]
    direita: Any = ""
    if logos_direita:
        celulas: list[Any] = []
        larguras: list[float] = []
        for indice, logo in enumerate(logos_direita):
            if indice:
                celulas.append("")
                larguras.append(0.3 * cm)
            celulas.append(logo)
            larguras.append(logo.drawWidth)
        direita = Table([celulas], colWidths=larguras, hAlign="RIGHT")
        direita.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ]
            )
        )
    larguras_externas = [6 * cm, _LARGURA_UTIL - 13 * cm, 7 * cm]
    tabela = Table([[logo_esquerda, textos, direita]], colWidths=larguras_externas)
    tabela.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, 0), "LEFT"),
                ("ALIGN", (2, 0), (2, 0), "RIGHT"),
                ("LEFTPADDING", (0, 0), (0, 0), 0),
                ("RIGHTPADDING", (2, 0), (2, 0), 0),
                ("LINEBELOW", (0, 0), (-1, 0), 1, _AZUL_ESCURO),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    return tabela


def _tabela_parametros(res: ResultadoSimulacao, estilos: dict[str, ParagraphStyle]) -> Table:
    """Tabela de parâmetros da simulação."""
    dados = []
    for rotulo, valor, monetario in _parametros(res, incluir_totais=False):
        texto = formatar_brl(valor) if monetario else _texto_paragrafo(valor)
        dados.append([Paragraph(escape(rotulo), estilos["rotulo"]),
                      Paragraph(texto, estilos["normal"])])
    tabela = Table(dados, colWidths=[5.2 * cm, _LARGURA_UTIL - 5.2 * cm])
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2F2F2")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return tabela


def _tabela_resumo(resumo: pd.DataFrame, estilos: dict[str, ParagraphStyle]) -> Table:
    """Tabela 'Resumo por ano' com a linha TOTAL em negrito."""
    cabecalho = [Paragraph(escape(str(c)), estilos["cabecalho"]) for c in resumo.columns]
    dados: list[list[Any]] = [cabecalho]
    for _, linha in resumo.iterrows():
        celulas: list[Any] = []
        for coluna in resumo.columns:
            valor = linha[coluna]
            if coluna in COLUNAS_MONETARIAS_RESUMO:
                celulas.append(formatar_brl(valor))
            else:
                celulas.append(str(valor))
        dados.append(celulas)

    colunas_fixas = 0.0
    qtd_monetarias = 0
    for col in resumo.columns:
        if col == "Ano":
            colunas_fixas += 2.2 * cm
        elif col == "Meses":
            colunas_fixas += 1.1 * cm
        elif col == "Quantidade de vagas":
            colunas_fixas += 2.0 * cm
        else:
            qtd_monetarias += 1
    largura_monetaria = (_LARGURA_UTIL - colunas_fixas) / max(qtd_monetarias, 1)
    larguras: list[float] = []
    for col in resumo.columns:
        if col == "Ano":
            larguras.append(2.2 * cm)
        elif col == "Meses":
            larguras.append(1.1 * cm)
        elif col == "Quantidade de vagas":
            larguras.append(2.0 * cm)
        else:
            larguras.append(largura_monetaria)

    tabela = Table(dados, colWidths=larguras, repeatRows=1)
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), _AZUL_ESCURO),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 7.5),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("ALIGN", (1, 1), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, _ZEBRA]),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, -1), (-1, -1), _CINZA_CLARO),
    ]
    if "Quantidade de vagas" in resumo.columns:
        idx_vagas = list(resumo.columns).index("Quantidade de vagas")
        estilo.append(("ALIGN", (idx_vagas, 1), (idx_vagas, -1), "CENTER"))
    tabela.setStyle(TableStyle(estilo))
    return tabela


def _tabela_detalhamento(res: ResultadoSimulacao, estilos: dict[str, ParagraphStyle]) -> Table:
    """Tabela 'Detalhamento mensal' por vaga, zebrada e com cabeçalho repetido."""
    colunas: list[str] = [
        "Competência",
        "Vencimento Básico",
    ]
    if not res.is_docente:
        colunas.append(ROTULO_GRATIFICACAO)
    colunas.extend(
        [
            "Incentivo",
            "RBM",
            "13º",
            "Férias",
            "INSS",
            "Custo total por vaga",
        ]
    )
    dados: list[list[Any]] = [[Paragraph(escape(c), estilos["cabecalho"]) for c in colunas]]
    for mes in res.meses:
        rotulo = mes["competencia"]
        if mes["extrapolado"]:
            rotulo += " *"
        linha = [
            rotulo,
            formatar_brl(mes["vencimento_basico"]),
        ]
        if not res.is_docente:
            linha.append(formatar_brl(mes["gratificacao"]))
        linha.extend(
            [
                formatar_brl(mes["incentivo"]),
                formatar_brl(mes["rbm"]),
                formatar_brl(mes["decimo_terceiro"]),
                formatar_brl(mes["ferias"]),
                formatar_brl(mes["inss"]),
                formatar_brl(mes["custo_total"]),
            ]
        )
        dados.append(linha)
    primeira = 2.8 * cm
    restante = (_LARGURA_UTIL - primeira) / (len(colunas) - 1)
    tabela = Table(dados, colWidths=[primeira] + [restante] * (len(colunas) - 1), repeatRows=1)
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), _AZUL_ESCURO),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 1), (-1, -1), 7.5),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 1), (0, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, _ZEBRA]),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    return tabela


_LARGURA_UTIL_SINTETICO = A4[0] - 3.0 * cm  # 18.0 cm (margens de 1.5 cm)


def _descricao_cargo_sintetico(md: dict[str, Any]) -> str:
    """Formata a descrição do cargo para o cabeçalho do relatório sintético."""
    qtd = f"{int(md.get('quantidade_vagas', 1)):02d}"
    cargo = str(md.get("cargo", "")).strip()
    nivel = str(md.get("nivel_classe_descricao", "")).strip()
    ch = str(md.get("ch", "")).strip()
    cat = str(md.get("categoria", "")).strip()

    if cargo.lower().startswith("professor "):
        classe = cargo[10:].strip().upper()
        return f"{qtd} Professor classe {classe}, {nivel}, regime de {ch}:"
    return f"{qtd} {cat} classe {cargo.upper()}, {nivel}, regime de {ch}:"


def _tabela_logos_sintetico() -> Table | None:
    """Carrega as logomarcas e informações institucionais no topo para o relatório sintético."""
    logo_esq = _carregar_logo(LOGO_UEFS)
    logos_dir = [
        logo for logo in (_carregar_logo(LOGO_PGDP), _carregar_logo(LOGO_UEFS_50_ANOS))
        if logo is not None
    ]

    estilo_uefs = ParagraphStyle(
        "sintetico_inst",
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=13,
        alignment=TA_CENTER,
    )
    estilo_pgdp = ParagraphStyle(
        "sintetico_pro",
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        alignment=TA_CENTER,
    )
    celula_centro = [
        Paragraph("Universidade Estadual de Feira de Santana – UEFS", estilo_uefs),
        Paragraph(
            "Pró-Reitoria de Gestão e Desenvolvimento de Pessoas – PGDP", estilo_pgdp
        ),
    ]

    celula_esq: Any = logo_esq or ""
    celula_dir: Any = ""
    if logos_dir:
        celulas = []
        larguras = []
        for indice, logo in enumerate(logos_dir):
            if indice:
                celulas.append("")
                larguras.append(0.3 * cm)
            celulas.append(logo)
            larguras.append(logo.drawWidth)
        celula_dir = Table([celulas], colWidths=larguras, hAlign="RIGHT")
        celula_dir.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ]
            )
        )

    larg_esq = 4.2 * cm
    larg_dir = 4.2 * cm
    larg_centro = _LARGURA_UTIL_SINTETICO - (larg_esq + larg_dir)
    tabela = Table(
        [[celula_esq, celula_centro, celula_dir]],
        colWidths=[larg_esq, larg_centro, larg_dir],
    )
    tabela.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, 0), "LEFT"),
                ("ALIGN", (1, 0), (1, 0), "CENTER"),
                ("ALIGN", (2, 0), (2, 0), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return tabela


def _tabela_remuneracao_base_sintetico(res: ResultadoSimulacao) -> Table:
    """Tabela 1: Remuneração Base (Topo) no primeiro mês de ingresso."""
    md = res.metadados
    mes1 = res.meses[0] if res.meses else {}
    cargo_nome = str(md.get("cargo", ""))
    vencimento = formatar_brl(mes1.get("vencimento_basico", 0))
    incentivos = formatar_brl(mes1.get("gratificacao", 0) + mes1.get("incentivo", 0))
    total = formatar_brl(mes1.get("rbm", 0))

    dados = [
        ["CARGO", "REMUNERAÇÃO", "INCENTIVOS", "TOTAL"],
        [cargo_nome, vencimento, incentivos, total],
    ]
    tabela = Table(dados, colWidths=[6.0 * cm, 4.0 * cm, 4.0 * cm, 4.0 * cm])
    tabela.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (0, 1), (0, 1), "LEFT"),
                ("ALIGN", (1, 1), (-1, 1), "RIGHT"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return tabela


def _tabelas_matriz_impacto_sintetico(res: ResultadoSimulacao) -> tuple[Table, Paragraph | None, Table]:
    """Tabela 2 (Matriz Principal), Nota do Ano 1 e Tabela 3 (Quadro Resumo)."""
    md = res.metadados
    ano_ingresso = int(md["ano_ingresso"])
    ano1, ano2, ano3 = ano_ingresso, ano_ingresso + 1, ano_ingresso + 2
    vagas = int(md.get("quantidade_vagas", 1))
    qtd_str = f"{vagas:02d}"
    cargo = str(md.get("cargo", ""))
    ch = str(md.get("ch", ""))
    eh_ano1_parcial = int(md.get("mes_ingresso", 1)) > 1

    col_ano1 = f"{ano1}*" if eh_ano1_parcial else str(ano1)
    cabecalho2 = ["Acréscimo", col_ano1, str(ano2), str(ano3)]

    totais_por_ano: dict[int, dict[str, float]] = {}
    for k in (1, 2, 3):
        meses_k = [m for m in res.meses if m["ano_projecao"] == k]
        rbm = sum(m["rbm"] for m in meses_k) * vagas
        decimo = sum(m["decimo_terceiro"] for m in meses_k) * vagas
        ferias = sum(m["ferias"] for m in meses_k) * vagas
        inss = sum(m["inss"] for m in meses_k) * vagas
        sub_prov = rbm + decimo + ferias
        sub_enc = inss
        total = sub_prov + sub_enc
        totais_por_ano[k] = {
            "rbm": rbm,
            "decimo": decimo,
            "ferias": ferias,
            "subtotal_proventos": sub_prov,
            "inss": inss,
            "subtotal_encargos": sub_enc,
            "total_geral": total,
        }

    aliq_pct = int(round(float(md.get("aliquota_patronal", 0.20)) * 100))
    rotulo_encargo = f"Encargo Patronal ({aliq_pct}%) - Anual"
    rotulo_linha8 = f"Total Geral Contratação {qtd_str} {cargo} {ch}"

    dados2 = [
        cabecalho2,
        ["Anual", formatar_brl(totais_por_ano[1]["rbm"]), formatar_brl(totais_por_ano[2]["rbm"]), formatar_brl(totais_por_ano[3]["rbm"])],
        ["13° Salário Variável", formatar_brl(totais_por_ano[1]["decimo"]), formatar_brl(totais_por_ano[2]["decimo"]), formatar_brl(totais_por_ano[3]["decimo"])],
        ["Férias (1/3)", formatar_brl(totais_por_ano[1]["ferias"]), formatar_brl(totais_por_ano[2]["ferias"]), formatar_brl(totais_por_ano[3]["ferias"])],
        ["SUBTOTAL DE PROVENTOS (ANUAL)", formatar_brl(totais_por_ano[1]["subtotal_proventos"]), formatar_brl(totais_por_ano[2]["subtotal_proventos"]), formatar_brl(totais_por_ano[3]["subtotal_proventos"])],
        [rotulo_encargo, formatar_brl(totais_por_ano[1]["inss"]), formatar_brl(totais_por_ano[2]["inss"]), formatar_brl(totais_por_ano[3]["inss"])],
        ["SUBTOTAL DE ENCARGOS (ANUAL)", formatar_brl(totais_por_ano[1]["subtotal_encargos"]), formatar_brl(totais_por_ano[2]["subtotal_encargos"]), formatar_brl(totais_por_ano[3]["subtotal_encargos"])],
        ["TOTAL ANUAL GERAL", formatar_brl(totais_por_ano[1]["total_geral"]), formatar_brl(totais_por_ano[2]["total_geral"]), formatar_brl(totais_por_ano[3]["total_geral"])],
        [rotulo_linha8, formatar_brl(totais_por_ano[1]["total_geral"]), formatar_brl(totais_por_ano[2]["total_geral"]), formatar_brl(totais_por_ano[3]["total_geral"])],
    ]

    tabela2 = Table(dados2, colWidths=[7.5 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm])
    tabela2.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (0, 1), (0, -1), "LEFT"),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                # Linha 4: SUBTOTAL DE PROVENTOS (ANUAL)
                ("FONTNAME", (0, 4), (-1, 4), "Helvetica-Bold"),
                ("BACKGROUND", (0, 4), (-1, 4), colors.HexColor("#E7E6E6")),
                # Linha 6: SUBTOTAL DE ENCARGOS (ANUAL)
                ("FONTNAME", (0, 6), (-1, 6), "Helvetica-Bold"),
                ("BACKGROUND", (0, 6), (-1, 6), colors.HexColor("#E7E6E6")),
                # Linha 7: TOTAL ANUAL GERAL
                ("FONTNAME", (0, 7), (-1, 7), "Helvetica-Bold"),
                ("BACKGROUND", (0, 7), (-1, 7), colors.HexColor("#E7E6E6")),
                # Linha 8: Total Geral Contratação...
                ("FONTNAME", (0, 8), (-1, 8), "Helvetica-Bold"),
            ]
        )
    )

    nota_ano1 = None
    if eh_ano1_parcial:
        meses_nomes = (
            "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
        )
        nome_mes = meses_nomes[int(md["mes_ingresso"]) - 1]
        nota_ano1 = Paragraph(
            f"<i>* contratações em 01 de {nome_mes} de {ano1}</i>",
            ParagraphStyle(
                "nota_ano1_sintetico",
                fontName="Helvetica",
                fontSize=8,
                leading=10,
                textColor=colors.HexColor("#333333"),
            ),
        )

    dados3 = [
        ["TOTAIS", col_ano1, str(ano2), str(ano3)],
        ["Total Geral da Contratação", formatar_brl(totais_por_ano[1]["total_geral"]), formatar_brl(totais_por_ano[2]["total_geral"]), formatar_brl(totais_por_ano[3]["total_geral"])],
    ]
    tabela3 = Table(dados3, colWidths=[7.5 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm])
    tabela3.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (0, 1), (0, 1), "LEFT"),
                ("ALIGN", (1, 1), (-1, 1), "RIGHT"),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )

    return tabela2, nota_ano1, tabela3


def _gerar_pdf_sintetico(res: ResultadoSimulacao, resumo: pd.DataFrame) -> bytes:
    """Gera o Relatório Sintético no layout oficial de demonstrativo da UEFS."""
    md = res.metadados
    gerado = _data_geracao(res)
    ano_ingresso = int(md["ano_ingresso"])
    ano1, ano2, ano3 = ano_ingresso, ano_ingresso + 1, ano_ingresso + 2

    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
        title="Impacto Orçamentário – Demonstrativo Sintético",
        author="Universidade Estadual de Feira de Santana – UEFS",
    )

    estilo_titulo = ParagraphStyle(
        "sintetico_titulo",
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        alignment=TA_CENTER,
        textColor=colors.black,
    )
    estilo_subtitulo = ParagraphStyle(
        "sintetico_subtitulo",
        fontName="Helvetica-Oblique",
        fontSize=8.5,
        leading=11,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#333333"),
    )
    estilo_desc_cargo = ParagraphStyle(
        "sintetico_desc_cargo",
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=13,
        textColor=colors.black,
    )

    historia: list[Any] = []

    logos = _tabela_logos_sintetico()
    if logos is not None:
        historia.append(logos)
        historia.append(Spacer(1, 4))

    historia.append(
        Paragraph(
            f"IMPACTO ORÇAMENTÁRIO DE GASTOS DE PESSOAL PARA SELEÇÃO EM {ano1}, {ano2} E {ano3}",
            estilo_titulo,
        )
    )
    historia.append(Spacer(1, 2))
    historia.append(
        Paragraph("* Valores de referência baseados na planilha importada.", estilo_subtitulo)
    )
    historia.append(Spacer(1, 8))

    historia.append(Paragraph(_descricao_cargo_sintetico(md), estilo_desc_cargo))
    historia.append(Spacer(1, 4))

    historia.append(_tabela_remuneracao_base_sintetico(res))
    historia.append(Spacer(1, 10))

    tabela2, nota_ano1, tabela3 = _tabelas_matriz_impacto_sintetico(res)
    historia.append(tabela2)
    if nota_ano1 is not None:
        historia.append(Spacer(1, 3))
        historia.append(nota_ano1)
    historia.append(Spacer(1, 10))

    historia.append(tabela3)

    rodape = f"ImpactoUEFS – Relatório Sintético – gerado em {gerado:%d/%m/%Y %H:%M}"
    documento.build(historia, canvasmaker=_criar_canvas(rodape))
    return buffer.getvalue()


def _gerar_pdf_detalhado(res: ResultadoSimulacao, resumo: pd.DataFrame) -> bytes:
    """Gera o relatório PDF detalhado (A4 paisagem, margens de 2 cm, Helvetica)."""
    estilos = _estilos()
    gerado = _data_geracao(res)
    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=_PAGINA,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        title="Relatório de Impacto Orçamentário – UEFS",
        author="Universidade Estadual de Feira de Santana – UEFS",
    )
    historia: list[Any] = [
        _cabecalho(estilos),
        Paragraph("Parâmetros da simulação", estilos["secao"]),
        _tabela_parametros(res, estilos),
        Paragraph("Resumo por ano", estilos["secao"]),
        _tabela_resumo(resumo, estilos),
        Spacer(1, 6),
        Paragraph(
            f"Impacto total ({res.metadados['quantidade_vagas']} vaga(s)): "
            f"<b>{formatar_brl(res.metadados['total_geral'])}</b> &nbsp;|&nbsp; "
            f"Custo por vaga: <b>{formatar_brl(res.metadados['total_por_vaga'])}</b> "
            f"&nbsp;|&nbsp; Custo médio mensal por vaga: "
            f"<b>{formatar_brl(res.metadados['custo_medio_mensal'])}</b>",
            estilos["normal"],
        ),
        PageBreak(),
        Paragraph("Detalhamento mensal (valores por vaga)", estilos["secao"]),
        _tabela_detalhamento(res, estilos),
        Spacer(1, 4),
        Paragraph(
            "* = valor mantido da última vigência salarial.",
            estilos["nota"],
        ),
        Paragraph("Nota metodológica e embasamento legal", estilos["secao"]),
        Paragraph(
            f"<b>Composição e Fórmulas:</b> RBM = vencimento básico + {ROTULO_GRATIFICACAO} + incentivo "
            "(titulação para docentes; incentivo para Motorista). 13º salário = RBM / 12; adicional de férias = (RBM / 3) / 12. "
            "Benefícios indenizatórios (auxílio-alimentação, transporte etc.) estão excluídos.<br/>"
            "<b>Encargo Patronal sobre Remuneração e Férias Gozadas (STF Tema 985 / RE 1.072.485):</b><br/>"
            "• <b>Servidor Efetivo (28% - RPPS):</b> Lei Estadual nº 11.357/2009 (art. 68), alterada pela Lei nº 14.520/2022. "
            "Natureza remuneratória das férias gozadas (art. 92 da Lei Estadual nº 6.677/1994) e incidência sobre o terço constitucional.<br/>"
            "• <b>Servidor Temporário / REDA (20% - RGPS/INSS):</b> Art. 40, § 13 da CF/88 c/c Lei Federal nº 8.212/1991 (art. 22, I) "
            "e Decreto nº 3.048/1999 (art. 214, § 4º), com incidência sobre os dias de descanso usufruídos e o terço constitucional.",
            estilos["nota"],
        ),
    ]
    rodape = f"ImpactoUEFS – gerado em {gerado:%d/%m/%Y %H:%M}"
    documento.build(historia, canvasmaker=_criar_canvas(rodape))
    return buffer.getvalue()


def gerar_pdf(
    res: ResultadoSimulacao,
    resumo: pd.DataFrame,
    tipo: str = "sintetico",
) -> bytes:
    """Gera o relatório PDF. tipo='sintetico' (padrão oficial UEFS) ou tipo='detalhado'."""
    if tipo == "detalhado":
        return _gerar_pdf_detalhado(res, resumo)
    return _gerar_pdf_sintetico(res, resumo)


# ---------------------------------------------------------------------------
# Excel e CSV
# ---------------------------------------------------------------------------
def _formatar_aba(
    aba: Worksheet,
    colunas_moeda: set[int],
    congelar: str,
    linhas_moeda: set[int] | None = None,
    negrito_ultima: bool = False,
) -> None:
    """Aplica cabeçalho em negrito, formato monetário, larguras e painel congelado."""
    preenchimento = PatternFill("solid", fgColor="D9E1F2")
    for celula in aba[1]:
        celula.font = Font(bold=True)
        celula.fill = preenchimento
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for linha in aba.iter_rows(min_row=2):
        for celula in linha:
            moeda = celula.column in colunas_moeda
            if linhas_moeda is not None and celula.row in linhas_moeda and celula.column == 2:
                moeda = True
            if moeda and isinstance(celula.value, (int, float)):
                celula.number_format = FORMATO_MOEDA_EXCEL
            if isinstance(celula.value, str) and "\n" in celula.value:
                celula.alignment = Alignment(wrap_text=True, vertical="top")
    if negrito_ultima and aba.max_row > 1:
        for celula in aba[aba.max_row]:
            celula.font = Font(bold=True)
    for indice, coluna in enumerate(aba.iter_cols(), start=1):
        maior = 0
        for celula in coluna:
            if celula.value is None:
                continue
            if isinstance(celula.value, (int, float)) and celula.number_format != "General":
                comprimento = len(formatar_brl(celula.value))
            else:
                comprimento = max(len(parte) for parte in str(celula.value).split("\n"))
            maior = max(maior, comprimento)
        aba.column_dimensions[get_column_letter(indice)].width = min(max(maior + 2, 10), 70)
    aba.freeze_panes = congelar


def gerar_excel(res: ResultadoSimulacao, resumo: pd.DataFrame) -> bytes:
    """Gera o .xlsx com as abas Parâmetros, Resumo Anual e Detalhamento Mensal."""
    parametros = _parametros(res, incluir_totais=True)
    df_parametros = pd.DataFrame(
        {
            "Parâmetro": [p[0] for p in parametros],
            "Valor": pd.Series([p[1] for p in parametros], dtype=object),
        }
    )
    linhas_moeda = {indice + 2 for indice, p in enumerate(parametros) if p[2]}
    mensal = tabela_mensal(res)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        df_parametros.to_excel(escritor, sheet_name="Parâmetros", index=False)
        resumo.to_excel(escritor, sheet_name="Resumo Anual", index=False)
        mensal.to_excel(escritor, sheet_name="Detalhamento Mensal", index=False)

        _formatar_aba(escritor.sheets["Parâmetros"], set(), "A2", linhas_moeda=linhas_moeda)
        moeda_resumo = {
            i for i, c in enumerate(resumo.columns, start=1) if c in COLUNAS_MONETARIAS_RESUMO
        }
        _formatar_aba(escritor.sheets["Resumo Anual"], moeda_resumo, "B2", negrito_ultima=True)
        moeda_mensal = {
            i for i, c in enumerate(mensal.columns, start=1) if c in COLUNAS_MONETARIAS_MENSAL
        }
        _formatar_aba(escritor.sheets["Detalhamento Mensal"], moeda_mensal, "C2")
    return buffer.getvalue()


def gerar_csv(res: ResultadoSimulacao) -> bytes:
    """Gera o CSV do detalhamento mensal (``;``, decimal ``,``, UTF-8 com BOM)."""
    buffer = io.BytesIO()
    tabela_mensal(res).to_csv(
        buffer,
        sep=";",
        decimal=",",
        float_format="%.2f",
        encoding="utf-8-sig",
        index=False,
    )
    return buffer.getvalue()
