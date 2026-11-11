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
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
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
MESES_EXTENSO: tuple[str, ...] = (
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
)

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
        ("Cargo Professor" if res.is_docente else "Cargo", md["cargo"], False),
        ("Nível/Classe", md["nivel_classe_descricao"], False),
        ("Carga Horária", md["ch"], False),
        ("Incentivo" if res.is_docente else "Gratificação", md["incentivo_descricao"], False),
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
    if md.get("setor_descricao"):
        itens.append(("Setor responsável", md["setor_descricao"], False))
    elif md.get("setor"):
        itens.append(("Setor responsável", md["setor"], False))
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
            self.setFont("Helvetica", 7.5)
            self.setFillColor(colors.grey)
            self.setStrokeColor(colors.lightgrey)
            self.line(1.5 * cm, 1.5 * cm, largura - 1.5 * cm, 1.5 * cm)
            self.drawString(1.5 * cm, 1.1 * cm, rodape)
            self.drawRightString(
                largura - 1.5 * cm, 1.1 * cm, f"Página {self._pageNumber} de {total}"
            )
            self.restoreState()

    return CanvasNumerado


def _formatar_rodape(res: ResultadoSimulacao, prefixo: str = "") -> str:
    """Gera o texto do rodapé incluindo setor responsável e data/hora."""
    md = res.metadados
    gerado = _data_geracao(res)
    setor_sigla = str(md.get("setor") or engine.SETOR_PADRAO)
    setor_nome = md.get("setor_nome") or engine.SETORES.get(setor_sigla, "")
    setor_info = f"{setor_sigla} - {setor_nome}" if setor_nome else setor_sigla
    partes: list[str] = []
    if prefixo:
        partes.append(prefixo)
    partes.append(f"Setor: {setor_info}")
    partes.append(f"gerado em {gerado:%d/%m/%Y %H:%M}")
    return " – ".join(partes)


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


def _tabela_parametros_sintetico(res: ResultadoSimulacao) -> Table:
    """Tabela de Parâmetros de Cálculo no relatório sintético."""
    md = res.metadados
    categoria = str(md.get("categoria", ""))
    vinculo = str(md.get("vinculo", ""))
    cargo = str(md.get("cargo", ""))
    nivel = str(md.get("nivel_classe_descricao", ""))
    ch = str(md.get("ch", ""))
    vagas = int(md.get("quantidade_vagas", 1))
    vagas_str = f"{vagas:02d} {'vaga' if vagas == 1 else 'vagas'}"

    mes_ingresso = int(md.get("mes_ingresso", 1))
    ano_ingresso = int(md.get("ano_ingresso", 2027))
    meses_nomes = (
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
    )
    nome_mes = meses_nomes[mes_ingresso - 1]
    texto_contratacao = f"* Início em 01 de {nome_mes} de {ano_ingresso}"

    rotulo_cargo = "Cargo Professor:" if res.is_docente else "Cargo:"

    estilo_titulo = ParagraphStyle(
        "sint_param_titulo",
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        alignment=TA_LEFT,
        textColor=colors.black,
    )
    estilo_lbl = ParagraphStyle(
        "sint_param_lbl",
        fontName="Helvetica-Bold",
        fontSize=7,
        leading=8.5,
        textColor=colors.HexColor("#222222"),
    )
    estilo_val = ParagraphStyle(
        "sint_param_val",
        fontName="Helvetica",
        fontSize=7,
        leading=8.5,
        textColor=colors.black,
    )
    estilo_nota = ParagraphStyle(
        "sint_param_nota",
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#1B365D"),
    )

    dados = [
        [Paragraph("<b>Parâmetros de cálculo</b>", estilo_titulo), "", "", ""],
        [
            Paragraph("<b>Categoria:</b>", estilo_lbl), Paragraph(escape(categoria), estilo_val),
            Paragraph("<b>Vínculo:</b>", estilo_lbl), Paragraph(escape(vinculo), estilo_val),
        ],
        [
            Paragraph(f"<b>{rotulo_cargo}</b>", estilo_lbl), Paragraph(escape(cargo), estilo_val),
            Paragraph("<b>Nível/Classe:</b>", estilo_lbl), Paragraph(escape(nivel), estilo_val),
        ],
        [
            Paragraph("<b>Carga Horária:</b>", estilo_lbl), Paragraph(escape(ch), estilo_val),
            Paragraph("<b>Quantidade de vaga(s):</b>", estilo_lbl), Paragraph(escape(vagas_str), estilo_val),
        ],
        [
            Paragraph(f"<b>{texto_contratacao}</b>", estilo_nota), "", "", "",
        ],
    ]

    tabela = Table(dados, colWidths=[3.2 * cm, 5.8 * cm, 3.5 * cm, 5.5 * cm])
    tabela.setStyle(
        TableStyle(
            [
                ("SPAN", (0, 0), (3, 0)),
                ("BACKGROUND", (0, 0), (3, 0), colors.HexColor("#E7E6E6")),
                ("BACKGROUND", (0, 1), (0, 3), colors.HexColor("#F7F7F7")),
                ("BACKGROUND", (2, 1), (2, 3), colors.HexColor("#F7F7F7")),
                ("SPAN", (0, 4), (3, 4)),
                ("BACKGROUND", (0, 4), (3, 4), colors.HexColor("#EBF3FA")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B0B0B0")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
                ("TOPPADDING", (0, 4), (3, 4), 3),
                ("BOTTOMPADDING", (0, 4), (3, 4), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
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
    adicional_valor = formatar_brl(mes1.get("gratificacao", 0) + mes1.get("incentivo", 0))
    total = formatar_brl(mes1.get("rbm", 0))

    col_composicao = "INCENTIVOS" if res.is_docente else "GRATIFICAÇÃO"

    dados = [
        ["CARGO", "REMUNERAÇÃO", col_composicao, "TOTAL"],
        [cargo_nome, vencimento, adicional_valor, total],
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


def _tabelas_matriz_impacto_sintetico(res: ResultadoSimulacao) -> Table:
    """Tabela 2 (Matriz Principal de Impacto) do relatório sintético."""
    md = res.metadados
    ano_ingresso = int(md["ano_ingresso"])
    ano1, ano2, ano3 = ano_ingresso, ano_ingresso + 1, ano_ingresso + 2
    vagas = int(md.get("quantidade_vagas", 1))
    qtd_str = f"{vagas:02d}"
    cargo = str(md.get("cargo", ""))
    ch = str(md.get("ch", ""))
    eh_ano1_parcial = int(md.get("mes_ingresso", 1)) > 1

    col_ano1 = f"{ano1}*" if eh_ano1_parcial else str(ano1)
    cabecalho2 = ["Composição", col_ano1, str(ano2), str(ano3), "Total (3 Anos)"]

    aliq = float(md.get("aliquota_patronal", 0.20))
    aliq_pct = int(round(aliq * 100))

    # Cálculos unitários (por 1 única vaga) para cada ano de projeção
    tot_u: dict[int, dict[str, float]] = {}
    for k in (1, 2, 3):
        meses_k = [m for m in res.meses if m["ano_projecao"] == k]
        num_m = len(meses_k) if meses_k else 1

        rbm_u = sum(m["rbm"] for m in meses_k)
        decimo_u = sum(m["decimo_terceiro"] for m in meses_k)
        ferias_u = sum(m["ferias"] for m in meses_k)
        sub_prov_u = rbm_u + decimo_u + ferias_u

        enc_rbm_m = (rbm_u / num_m) * aliq
        enc_decimo_m = (decimo_u / num_m) * aliq
        enc_ferias_m = 0.0
        enc_total_m = enc_rbm_m + enc_decimo_m

        enc_total_anual = sum(m["inss"] for m in meses_k)
        sub_enc_u = enc_total_anual

        custo_servidor = sub_prov_u + sub_enc_u

        tot_u[k] = {
            "rbm": rbm_u,
            "decimo": decimo_u,
            "ferias": ferias_u,
            "sub_prov": sub_prov_u,
            "enc_rbm_m": enc_rbm_m,
            "enc_decimo_m": enc_decimo_m,
            "enc_ferias_m": enc_ferias_m,
            "enc_total_m": enc_total_m,
            "enc_total_anual": enc_total_anual,
            "sub_enc": sub_enc_u,
            "custo_servidor": custo_servidor,
        }

    # Total dos 3 anos (Triênio) para 1 vaga
    trienio_u = {
        "rbm": sum(tot_u[k]["rbm"] for k in (1, 2, 3)),
        "decimo": sum(tot_u[k]["decimo"] for k in (1, 2, 3)),
        "ferias": sum(tot_u[k]["ferias"] for k in (1, 2, 3)),
        "sub_prov": sum(tot_u[k]["sub_prov"] for k in (1, 2, 3)),
        "enc_total_anual": sum(tot_u[k]["enc_total_anual"] for k in (1, 2, 3)),
        "sub_enc": sum(tot_u[k]["sub_enc"] for k in (1, 2, 3)),
        "custo_servidor": sum(tot_u[k]["custo_servidor"] for k in (1, 2, 3)),
    }

    # Valores totais multiplicados pela quantidade de vagas
    tot_v = {k: tot_u[k]["custo_servidor"] * vagas for k in (1, 2, 3)}
    trienio_v = trienio_u["custo_servidor"] * vagas

    rotulo_vagas = "vaga" if vagas == 1 else "vagas"
    qtd_fmt = f"{vagas:02d}"

    estilo_linha_rbm = ParagraphStyle(
        "sint_linha_rbm",
        fontName="Helvetica",
        fontSize=8,
        leading=9.5,
        alignment=TA_LEFT,
        textColor=colors.black,
    )
    celula_rbm = Paragraph(
        "Anual (Rendimento Básico Mensal - RBM - 1 vaga)",
        estilo_linha_rbm,
    )

    dados2 = [
        cabecalho2,
        [celula_rbm, formatar_brl(tot_u[1]["rbm"]), formatar_brl(tot_u[2]["rbm"]), formatar_brl(tot_u[3]["rbm"]), formatar_brl(trienio_u["rbm"])],
        ["13° Salário Variável", formatar_brl(tot_u[1]["decimo"]), formatar_brl(tot_u[2]["decimo"]), formatar_brl(tot_u[3]["decimo"]), formatar_brl(trienio_u["decimo"])],
        ["Férias (1/3)", formatar_brl(tot_u[1]["ferias"]), formatar_brl(tot_u[2]["ferias"]), formatar_brl(tot_u[3]["ferias"]), formatar_brl(trienio_u["ferias"])],
        ["SUBTOTAL DE PROVENTOS (1 VAGA)", formatar_brl(tot_u[1]["sub_prov"]), formatar_brl(tot_u[2]["sub_prov"]), formatar_brl(tot_u[3]["sub_prov"]), formatar_brl(trienio_u["sub_prov"])],
        [f"Encargo Patronal ({aliq_pct}%) s/ RBM (Mensal)", formatar_brl(tot_u[1]["enc_rbm_m"]), formatar_brl(tot_u[2]["enc_rbm_m"]), formatar_brl(tot_u[3]["enc_rbm_m"]), "-"],
        [f"Encargo Patronal ({aliq_pct}%) s/ 13º Salário (Mensal)", formatar_brl(tot_u[1]["enc_decimo_m"]), formatar_brl(tot_u[2]["enc_decimo_m"]), formatar_brl(tot_u[3]["enc_decimo_m"]), "-"],
        [f"Encargo Patronal ({aliq_pct}%) - Total Mensal", formatar_brl(tot_u[1]["enc_total_m"]), formatar_brl(tot_u[2]["enc_total_m"]), formatar_brl(tot_u[3]["enc_total_m"]), "-"],
        [f"Encargo Patronal ({aliq_pct}%) - Total Anual", formatar_brl(tot_u[1]["enc_total_anual"]), formatar_brl(tot_u[2]["enc_total_anual"]), formatar_brl(tot_u[3]["enc_total_anual"]), formatar_brl(trienio_u["enc_total_anual"])],
        ["SUBTOTAL DE ENCARGOS (1 VAGA)", formatar_brl(tot_u[1]["sub_enc"]), formatar_brl(tot_u[2]["sub_enc"]), formatar_brl(tot_u[3]["sub_enc"]), formatar_brl(trienio_u["sub_enc"])],
        ["CUSTO TOTAL DE 1 VAGA (ANUAL)", formatar_brl(tot_u[1]["custo_servidor"]), formatar_brl(tot_u[2]["custo_servidor"]), formatar_brl(tot_u[3]["custo_servidor"]), formatar_brl(trienio_u["custo_servidor"])],
        ["Quantidade de Vaga(s)", f"× {vagas}", f"× {vagas}", f"× {vagas}", f"× {vagas}"],
        [f"TOTAL GERAL ({qtd_fmt} {rotulo_vagas.upper()})", formatar_brl(tot_v[1]), formatar_brl(tot_v[2]), formatar_brl(tot_v[3]), formatar_brl(trienio_v)],
    ]

    tabela2 = Table(dados2, colWidths=[6.8 * cm, 2.8 * cm, 2.8 * cm, 2.8 * cm, 2.8 * cm])
    tabela2.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (0, 1), (0, -1), "LEFT"),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("ALIGN", (1, 11), (-1, 11), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                # Linha 4: SUBTOTAL DE PROVENTOS (1 VAGA)
                ("FONTNAME", (0, 4), (-1, 4), "Helvetica-Bold"),
                ("BACKGROUND", (0, 4), (-1, 4), colors.HexColor("#E7E6E6")),
                # Linha 7: Encargo Patronal - Total Mensal
                ("FONTNAME", (0, 7), (-1, 7), "Helvetica-Bold"),
                # Linha 9: SUBTOTAL DE ENCARGOS (1 VAGA)
                ("FONTNAME", (0, 9), (-1, 9), "Helvetica-Bold"),
                ("BACKGROUND", (0, 9), (-1, 9), colors.HexColor("#E7E6E6")),
                # Linha 10: CUSTO TOTAL DE 1 VAGA (ANUAL)
                ("FONTNAME", (0, 10), (-1, 10), "Helvetica-Bold"),
                ("BACKGROUND", (0, 10), (-1, 10), colors.HexColor("#E7E6E6")),
                # Linha 11: Quantidade de Vagas
                ("FONTNAME", (0, 11), (-1, 11), "Helvetica-Bold"),
                # Linha 12: TOTAL GERAL
                ("FONTNAME", (0, 12), (-1, 12), "Helvetica-Bold"),
            ]
        )
    )

    return tabela2


def _bloco_especificacao_patronal_sintetico(res: ResultadoSimulacao) -> Table:
    """Quadro explicativo com a especificação legal e base de cálculo do encargo patronal."""
    md = res.metadados
    aliq_pct = int(round(float(md.get("aliquota_patronal", 0.20)) * 100))
    vinculo = str(md.get("vinculo", ""))
    eh_efetivo = vinculo == "Efetivo"

    if eh_efetivo:
        regime = "RPPS / Estado da Bahia (Lei Estadual nº 11.357/2009 e Lei nº 14.520/2022)"
    else:
        regime = "RGPS / INSS (Art. 40, § 13 da CF/88 c/c Lei Federal nº 8.212/1991, art. 22, I)"

    estilo_bloco = ParagraphStyle(
        "sintetico_bloco_patronal",
        fontName="Helvetica",
        fontSize=7.8,
        leading=10.5,
        textColor=colors.HexColor("#222222"),
    )

    texto = (
        f"<b>Especificação da Base de Incidência dos Encargos Patronais ({aliq_pct}% – {regime}):</b><br/>"
        "O encargo patronal incide mensal e anualmente sobre as seguintes parcelas remuneratórias:<br/>"
        "• <b>Rendimento Básico Mensal (RBM):</b> Vencimento básico somado às gratificações e incentivos do cargo.<br/>"
        "• <b>13º Salário:</b> Gratificação natalina proporcional aos meses de efetivo exercício no ano.<br/>"
        "• <b>1/3 de Férias:</b> Não há incidência de encargo patronal sobre o adicional de férias.<br/>"
        f"<i>Fórmula de Incidência: Encargo Patronal = {aliq_pct}% × (RBM + 13º Salário).</i>"
    )

    p = Paragraph(texto, estilo_bloco)
    tabela = Table([[p]], colWidths=[18.0 * cm])
    tabela.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B0B0B0")),
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F9F9F9")),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    return tabela


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
            f"IMPACTO ORÇAMENTÁRIO {ano1}, {ano2} E {ano3}",
            estilo_titulo,
        )
    )
    historia.append(Spacer(1, 4))
    historia.append(_tabela_parametros_sintetico(res))
    historia.append(Spacer(1, 8))

    historia.append(_tabela_remuneracao_base_sintetico(res))
    historia.append(Spacer(1, 10))

    tabela2 = _tabelas_matriz_impacto_sintetico(res)
    historia.append(tabela2)

    rodape = _formatar_rodape(res)
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
        Paragraph("Parâmetros de cálculo", estilos["secao"]),
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
            "(titulação para docentes; incentivo para Motorista). 13º salário = RBM / 12; adicional de férias = (RBM / 3) / 12 (a partir do 2º ano). "
            "Benefícios indenizatórios (auxílio-alimentação, transporte etc.) estão excluídos.<br/>"
            "<b>Base de Incidência dos Encargos Patronais:</b> O encargo patronal incide mensal e anualmente sobre as "
            "parcelas remuneratórias: <b>Rendimento Básico Mensal (RBM) + 13º Salário</b> "
            "(Fórmula: Encargo = Alíquota × [RBM + 13º Salário]). Não há incidência de encargo patronal sobre o adicional de férias.<br/>"
            "• <b>Servidor Efetivo (28% - RPPS):</b> Lei Estadual nº 11.357/2009 (art. 68), alterada pela Lei nº 14.520/2022.<br/>"
            "• <b>Servidor Temporário / REDA (20% - RGPS/INSS):</b> Art. 40, § 13 da CF/88 c/c Lei Federal nº 8.212/1991 (art. 22, I).",
            estilos["nota"],
        ),
    ]
    rodape = _formatar_rodape(res)
    documento.build(historia, canvasmaker=_criar_canvas(rodape))
    return buffer.getvalue()


def _gerar_pdf_memoria_calculo(res: ResultadoSimulacao, resumo: pd.DataFrame) -> bytes:
    """Gera o relatório oficial de Memória de Cálculo discriminando como cada valor é apurado."""
    md = res.metadados
    gerado = _data_geracao(res)
    ano_ingresso = int(md["ano_ingresso"])
    mes_ingresso = int(md["mes_ingresso"])
    ano1, ano2, ano3 = ano_ingresso, ano_ingresso + 1, ano_ingresso + 2
    vagas = int(md.get("quantidade_vagas", 1))
    aliq = float(md.get("aliquota_patronal", 0.20))
    aliq_pct = int(round(aliq * 100))
    vinculo = str(md.get("vinculo", ""))
    eh_efetivo = (vinculo == "Efetivo")
    regime_desc = "RPPS / Estado da Bahia (28%)" if eh_efetivo else "RGPS / INSS (20%)"

    mes1 = res.meses[0] if res.meses else {}
    venc_m = mes1.get("vencimento_basico", 0.0)
    grat_m = mes1.get("gratificacao", 0.0)
    inc_m = mes1.get("incentivo", 0.0)
    rbm_m = mes1.get("rbm", 0.0)
    decimo_m = mes1.get("decimo_terceiro", 0.0)
    ferias_m = mes1.get("ferias", 0.0)
    sub_prov_m = rbm_m + decimo_m + ferias_m

    enc_rbm_m = rbm_m * aliq
    enc_decimo_m = decimo_m * aliq
    enc_ferias_m = 0.0
    enc_total_m = mes1.get("inss", 0.0)
    custo_total_m = mes1.get("custo_total", 0.0)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title="Memória de Cálculo – Impacto Orçamentário UEFS",
        author="Universidade Estadual de Feira de Santana – UEFS",
    )

    st_titulo = ParagraphStyle("mem_titulo", fontName="Helvetica-Bold", fontSize=11, leading=13, alignment=TA_CENTER)
    st_sub = ParagraphStyle("mem_sub", fontName="Helvetica-Oblique", fontSize=7.5, leading=9.5, alignment=TA_CENTER, textColor=colors.HexColor("#333333"))
    st_sec = ParagraphStyle("mem_sec", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=colors.HexColor("#1B365D"))

    st_th = ParagraphStyle("mem_th", fontName="Helvetica-Bold", fontSize=7, leading=8.5, alignment=TA_CENTER)
    st_td = ParagraphStyle("mem_td", fontName="Helvetica", fontSize=7, leading=8.5, alignment=TA_LEFT)
    st_td_b = ParagraphStyle("mem_td_b", fontName="Helvetica-Bold", fontSize=7, leading=8.5, alignment=TA_LEFT)
    st_td_r = ParagraphStyle("mem_td_r", fontName="Helvetica", fontSize=7, leading=8.5, alignment=TA_RIGHT)
    st_td_rb = ParagraphStyle("mem_td_rb", fontName="Helvetica-Bold", fontSize=7, leading=8.5, alignment=TA_RIGHT)
    st_td_c = ParagraphStyle("mem_td_c", fontName="Helvetica", fontSize=7, leading=8.5, alignment=TA_CENTER)
    st_p = ParagraphStyle("mem_p", fontName="Helvetica", fontSize=7.5, leading=9.5)
    st_pb = ParagraphStyle("mem_pb", fontName="Helvetica-Bold", fontSize=7.5, leading=9.5)

    historia: list[Any] = []

    logos = _tabela_logos_sintetico()
    if logos is not None:
        historia.append(logos)
        historia.append(Spacer(1, 3))

    historia.append(Paragraph("MEMÓRIA DE CÁLCULO DISCRIMINADA DO IMPACTO ORÇAMENTÁRIO", st_titulo))
    historia.append(Spacer(1, 1))
    historia.append(Paragraph(f"Detalhamento Metodológico, Fórmulas e Composição de Custos – Projeção Trienal {ano1}, {ano2} e {ano3}", st_sub))
    historia.append(Spacer(1, 5))

    # 1. Parâmetros
    historia.append(Paragraph("1. PARÂMETROS DE CÁLCULO", st_sec))
    historia.append(Spacer(1, 2))
    eh_parcial = mes_ingresso > 1
    mes_nome = MESES_EXTENSO[mes_ingresso - 1]
    cargo_txt = escape(str(md.get("cargo", "")))
    categ_txt = escape(str(md.get("categoria", "")))
    ch_txt = escape(str(md.get("ch", "")))
    nivel_txt = escape(str(md.get("nivel_classe", "")))
    inc_txt = escape(str(md.get("incentivo_descricao", "—")))
    vagas_txt = f"{vagas:02d} {'vaga' if vagas == 1 else 'vagas'}"

    rotulo_inc_mem = "Incentivo:" if res.is_docente else "Gratificação:"
    dados_param = [
        [Paragraph("<b>Cargo:</b>", st_p), Paragraph(f"{cargo_txt} ({categ_txt} - {vinculo})", st_p),
         Paragraph("<b>Carga Horária / Nível:</b>", st_p), Paragraph(f"{ch_txt} | Nível: {nivel_txt}", st_p)],
        [Paragraph("<b>Ingresso:</b>", st_p), Paragraph(f"01 de {mes_nome} de {ano1} ({13 - mes_ingresso} meses no Ano 1)", st_p),
         Paragraph("<b>Vagas Solicitadas:</b>", st_p), Paragraph(vagas_txt, st_pb)],
        [Paragraph("<b>Previdência / Alíquota:</b>", st_p), Paragraph(f"{regime_desc}", st_p),
         Paragraph(f"<b>{rotulo_inc_mem}</b>", st_p), Paragraph(f"{inc_txt}", st_p)],
    ]
    tab_param = Table(dados_param, colWidths=[3.2 * cm, 5.8 * cm, 3.8 * cm, 5.2 * cm])
    tab_param.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B0B0B0")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2F2F2")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#F2F2F2")),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
    ]))
    historia.append(tab_param)
    historia.append(Spacer(1, 5))

    # 2. Memória Mensal Proventos
    historia.append(Paragraph("2. MEMÓRIA DE CÁLCULO MENSAL DE 1 SERVIDOR (PROVENTOS)", st_sec))
    historia.append(Spacer(1, 2))
    rubrica_comp = "Incentivo" if res.is_docente else "Gratificação"
    regra_comp = "Percentual referente à titulação" if res.is_docente else "Percentual inerente ao cargo/função"
    formula_rbm = "Vencimento + Incentivo" if res.is_docente else "Vencimento + Gratificação"

    dados_prov = [
        [Paragraph("Rubrica / Parcela", st_th), Paragraph("Regra de Apuração / Base de Cálculo", st_th), Paragraph("Fórmula Matemática", st_th), Paragraph("Valor Mensal (1 Serv.)", st_th)],
        [Paragraph("Vencimento Básico", st_td), Paragraph("Tabela Salarial oficial vigente na data de ingresso", st_td), Paragraph("Vencimento fixado em lei", st_td), Paragraph(formatar_brl(venc_m), st_td_r)],
        [Paragraph(rubrica_comp, st_td), Paragraph(regra_comp, st_td), Paragraph("Aplicado sobre o vencimento básico", st_td), Paragraph(formatar_brl(grat_m + inc_m), st_td_r)],
        [Paragraph("Rendimento Básico Mensal (RBM)", st_td_b), Paragraph("Base remuneratória mensal total do servidor", st_td), Paragraph(formula_rbm, st_td), Paragraph(formatar_brl(rbm_m), st_td_rb)],
        [Paragraph("13º Salário Proporcional", st_td), Paragraph("Provisão mensal de gratificação natalina (1/12)", st_td), Paragraph("RBM ÷ 12", st_td), Paragraph(formatar_brl(decimo_m), st_td_r)],
        [Paragraph("Adicional de Férias (1/3)", st_td), Paragraph("Provisão mensal do terço constitucional (1/36)", st_td), Paragraph("(RBM ÷ 3) ÷ 12 = RBM ÷ 36", st_td), Paragraph(formatar_brl(ferias_m), st_td_r)],
        [Paragraph("SUBTOTAL DE PROVENTOS (MENSAL)", st_td_b), Paragraph("Total de direitos remuneratórios mensais por vaga", st_td_b), Paragraph("RBM + 13º Salário + 1/3 de Férias", st_td_b), Paragraph(formatar_brl(sub_prov_m), st_td_rb)],
    ]
    tab_prov = Table(dados_prov, colWidths=[4.8 * cm, 5.8 * cm, 4.4 * cm, 3.0 * cm])
    tab_prov.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("BACKGROUND", (0, 3), (-1, 3), colors.HexColor("#F7F7F7")),
        ("BACKGROUND", (0, 6), (-1, 6), colors.HexColor("#E7E6E6")),
    ]))
    historia.append(tab_prov)
    historia.append(Spacer(1, 5))

    # 3. Memória Mensal Encargos
    historia.append(Paragraph(f"3. MEMÓRIA DE CÁLCULO DOS ENCARGOS PATRONAIS ({aliq_pct}% – {regime_desc})", st_sec))
    historia.append(Spacer(1, 2))
    dados_enc = [
        [Paragraph("Parcela Incidente", st_th), Paragraph("Base de Cálculo", st_th), Paragraph("Alíq.", st_th), Paragraph("Encargo (1 Serv.)", st_th), Paragraph("Fundamentação Legal", st_th)],
        [Paragraph("Sobre Rendimento (RBM)", st_td), Paragraph(formatar_brl(rbm_m), st_td_r), Paragraph(f"{aliq_pct}%", st_td_c), Paragraph(formatar_brl(enc_rbm_m), st_td_r), Paragraph("CF/88 art. 40 §13 / Lei 8.212/91 ou Lei 11.357/09", st_td)],
        [Paragraph("Sobre 13º Salário", st_td), Paragraph(formatar_brl(decimo_m), st_td_r), Paragraph(f"{aliq_pct}%", st_td_c), Paragraph(formatar_brl(enc_decimo_m), st_td_r), Paragraph("Incidência previdenciária s/ gratificação natalina", st_td)],
        [Paragraph("Sobre 1/3 de Férias", st_td), Paragraph(formatar_brl(ferias_m), st_td_r), Paragraph("0%", st_td_c), Paragraph(formatar_brl(0.0), st_td_r), Paragraph("Não incide encargo patronal s/ 1/3 de férias", st_td)],
        [Paragraph("SUBTOTAL ENCARGOS PATRONAIS", st_td_b), Paragraph(formatar_brl(rbm_m + decimo_m), st_td_rb), Paragraph(f"{aliq_pct}%", st_td_c), Paragraph(formatar_brl(enc_total_m), st_td_rb), Paragraph(f"Soma das cotas patronais ({aliq_pct}% s/ RBM + 13º)", st_td_b)],
        [Paragraph("CUSTO MENSAL DE 1 SERVIDOR", st_td_b), Paragraph("—", st_td_c), Paragraph("—", st_td_c), Paragraph(formatar_brl(custo_total_m), st_td_rb), Paragraph("Proventos Mensais + Encargos Patronais Mensais", st_td_b)],
    ]
    tab_enc = Table(dados_enc, colWidths=[4.2 * cm, 2.5 * cm, 1.3 * cm, 2.8 * cm, 7.2 * cm])
    tab_enc.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("BACKGROUND", (0, 4), (-1, 4), colors.HexColor("#E7E6E6")),
        ("BACKGROUND", (0, 5), (-1, 5), colors.HexColor("#D9E1F2")),
    ]))
    historia.append(tab_enc)
    historia.append(Spacer(1, 5))

    # 4. Projeção Trienal
    historia.append(Paragraph("4. PROJEÇÃO TRIENAL E MULTIPLICAÇÃO PELA QUANTIDADE DE VAGAS", st_sec))
    historia.append(Spacer(1, 2))

    tot_u: dict[int, dict[str, Any]] = {}
    for k in (1, 2, 3):
        meses_k = [m for m in res.meses if m["ano_projecao"] == k]
        num_m = len(meses_k) if meses_k else 1
        prov_k = sum(m["rbm"] + m["decimo_terceiro"] + m["ferias"] for m in meses_k)
        enc_k = sum(m["inss"] for m in meses_k)
        tot_u[k] = {
            "meses": num_m,
            "prov": prov_k,
            "enc": enc_k,
            "custo": prov_k + enc_k,
        }

    tri_meses = sum(tot_u[k]["meses"] for k in (1, 2, 3))
    tri_prov = sum(tot_u[k]["prov"] for k in (1, 2, 3))
    tri_enc = sum(tot_u[k]["enc"] for k in (1, 2, 3))
    tri_custo = sum(tot_u[k]["custo"] for k in (1, 2, 3))

    col_ano1 = f"{ano1}*" if eh_parcial else str(ano1)
    dados_tri = [
        [Paragraph("Etapa da Memória de Cálculo", st_th), Paragraph(col_ano1, st_th), Paragraph(str(ano2), st_th), Paragraph(str(ano3), st_th), Paragraph("Total (3 Anos)", st_th)],
        [Paragraph("Meses de efetivo exercício no ano", st_td), Paragraph(f"{tot_u[1]['meses']} meses", st_td_c), Paragraph(f"{tot_u[2]['meses']} meses", st_td_c), Paragraph(f"{tot_u[3]['meses']} meses", st_td_c), Paragraph(f"{tri_meses} meses", st_td_c)],
        [Paragraph("(A) Subtotal de Proventos Anuais (1 Servidor)", st_td), Paragraph(formatar_brl(tot_u[1]["prov"]), st_td_r), Paragraph(formatar_brl(tot_u[2]["prov"]), st_td_r), Paragraph(formatar_brl(tot_u[3]["prov"]), st_td_r), Paragraph(formatar_brl(tri_prov), st_td_r)],
        [Paragraph("(B) Subtotal de Encargos Patronais Anuais (1 Servidor)", st_td), Paragraph(formatar_brl(tot_u[1]["enc"]), st_td_r), Paragraph(formatar_brl(tot_u[2]["enc"]), st_td_r), Paragraph(formatar_brl(tot_u[3]["enc"]), st_td_r), Paragraph(formatar_brl(tri_enc), st_td_r)],
        [Paragraph("(C) Custo Total Anual por Servidor (A + B)", st_td_b), Paragraph(formatar_brl(tot_u[1]["custo"]), st_td_rb), Paragraph(formatar_brl(tot_u[2]["custo"]), st_td_rb), Paragraph(formatar_brl(tot_u[3]["custo"]), st_td_rb), Paragraph(formatar_brl(tri_custo), st_td_rb)],
        [Paragraph("(D) Multiplicador de Vagas Solicitadas", st_td_b), Paragraph(f"× {vagas}", st_td_c), Paragraph(f"× {vagas}", st_td_c), Paragraph(f"× {vagas}", st_td_c), Paragraph(f"× {vagas}", st_td_c)],
        [Paragraph("(E) IMPACTO ORÇAMENTÁRIO TOTAL (C × D)", st_td_b), Paragraph(formatar_brl(tot_u[1]["custo"] * vagas), st_td_rb), Paragraph(formatar_brl(tot_u[2]["custo"] * vagas), st_td_rb), Paragraph(formatar_brl(tot_u[3]["custo"] * vagas), st_td_rb), Paragraph(formatar_brl(tri_custo * vagas), st_td_rb)],
    ]
    tab_tri = Table(dados_tri, colWidths=[6.8 * cm, 2.8 * cm, 2.8 * cm, 2.8 * cm, 2.8 * cm])
    tab_tri.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("BACKGROUND", (0, 4), (-1, 4), colors.HexColor("#E7E6E6")),
        ("BACKGROUND", (0, 6), (-1, 6), colors.HexColor("#D9E1F2")),
    ]))
    historia.append(tab_tri)
    historia.append(Spacer(1, 5))

    # 5. Base Legal
    historia.append(Paragraph("5. EMBASAMENTO LEGAL E NOTAS METODOLÓGICAS", st_sec))
    historia.append(Spacer(1, 2))
    texto_legal = (
        "<b>Regime Previdenciário e Cota Patronal:</b> Conforme art. 40, § 13 da CF/88 c/c Lei Federal nº 8.212/1991 (art. 22, I) "
        "para servidores temporários (RGPS/INSS – 20%) ou Lei Estadual nº 11.357/2009 c/c Lei Estadual nº 14.520/2022 para efetivos (RPPS – 28%).<br/>"
        "<b>Não Incidência Previdenciária sobre 1/3 de Férias:</b> Não há incidência de encargo patronal sobre o adicional de férias.<br/>"
        "<b>Exclusão de Verbas Indenizatórias:</b> Benefícios indenizatórios (auxílio-transporte, alimentação etc.) não integram o custo de pessoal apurado."
    )
    tab_legal = Table([[Paragraph(texto_legal, ParagraphStyle("leg", fontName="Helvetica", fontSize=6.8, leading=8.8, textColor=colors.HexColor("#333333")))]], colWidths=[18.0 * cm])
    tab_legal.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B0B0B0")),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F9F9F9")),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    historia.append(tab_legal)

    rodape = _formatar_rodape(res)
    doc.build(historia, canvasmaker=_criar_canvas(rodape))
    return buffer.getvalue()


def gerar_pdf(
    res: ResultadoSimulacao,
    resumo: pd.DataFrame,
    tipo: str = "sintetico",
) -> bytes:
    """Gera o relatório PDF. tipo='sintetico' (padrão oficial UEFS), tipo='memoria_calculo' ou tipo='detalhado'."""
    if tipo == "detalhado":
        return _gerar_pdf_detalhado(res, resumo)
    if tipo == "memoria_calculo":
        return _gerar_pdf_memoria_calculo(res, resumo)
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
