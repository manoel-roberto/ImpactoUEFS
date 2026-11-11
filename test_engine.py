"""Testes do motor de projeção, da validação da base salarial e dos relatórios."""
from __future__ import annotations

import io
from decimal import Decimal
from typing import Any

import pandas as pd
import pytest
from openpyxl import load_workbook

import data_manager
import engine
import relatorios
from engine import (
    BaseSalarialVaziaError,
    ResultadoSimulacao,
    Simulacao,
    VigenciaNaoEncontradaError,
    projetar,
    resumo_anual,
)

COLUNAS = list(engine.COLUNAS_BASE)
TEC_UNIV = ("Técnico Universitário", "Efetivo", "30h")
MOTORISTA = (engine.CARGO_MOTORISTA, "Temporário", "40h")
TEC_MEDIO_TEMP_40 = ("Técnico de Nível Médio", "Temporário", "40h")
PROF_AUX = ("Professor Auxiliar", "Efetivo", "40h")
JORNALISTA = (engine.CARGO_JORNALISTA, "Temporário", "25h")


def _linha(
    combinacao: tuple[str, str, str],
    vencimento: float,
    gratificacao: float = 0.0,
    percentual: float = 0.0,
    inicio: str = "2020-01-01",
    fim: str | None = None,
) -> dict[str, Any]:
    cargo, vinculo, ch = combinacao
    return {
        "Cargo": cargo,
        "Vinculo": vinculo,
        "CH": ch,
        "Vencimento_Basico": vencimento,
        "Valor_Gratificacao": gratificacao,
        "Percentual_Motorista": percentual,
        "Data_Inicio_Vigencia": inicio,
        "Data_Fim_Vigencia": fim,
    }


def _tabela(*linhas: dict[str, Any]) -> pd.DataFrame:
    df = pd.DataFrame(list(linhas), columns=COLUNAS)
    df["Data_Inicio_Vigencia"] = pd.to_datetime(df["Data_Inicio_Vigencia"])
    df["Data_Fim_Vigencia"] = pd.to_datetime(df["Data_Fim_Vigencia"])
    return df


def _sim(**kwargs: Any) -> Simulacao:
    parametros: dict[str, Any] = {
        "categoria": "Técnicos",
        "vinculo": "Efetivo",
        "cargo": "Técnico Universitário",
        "ch": "30h",
        "mes_ingresso": 1,
        "ano_ingresso": 2027,
    }
    parametros.update(kwargs)
    return Simulacao(**parametros)


def _sim_motorista(**kwargs: Any) -> Simulacao:
    return _sim(
        vinculo="Temporário",
        cargo=engine.CARGO_MOTORISTA,
        ch="40h",
        **kwargs,
    )


def _sim_jornalista(**kwargs: Any) -> Simulacao:
    parametros: dict[str, Any] = {
        "vinculo": "Temporário",
        "cargo": engine.CARGO_JORNALISTA,
        "ch": "25h",
    }
    parametros.update(kwargs)
    return _sim(**parametros)


def _df_texto(linhas: list[list[str]]) -> pd.DataFrame:
    return pd.DataFrame(linhas, columns=COLUNAS, dtype=str)


# 1. Efetivo, sem variação ---------------------------------------------------
def test_efetivo_sem_variacao() -> None:
    res = projetar(_sim(), _tabela(_linha(TEC_UNIV, 5000)))
    assert len(res.meses) == 36
    assert res.metadados["total_meses"] == 36
    assert set(res.meses[0].keys()) == set(engine.CHAVES_MES)
    for mes in res.meses:
        assert mes["rbm"] == 5000.00
        assert mes["decimo_terceiro"] == 416.67
        assert mes["inss"] == 1516.67  # 28% sobre (5000 + 416.666...) sem incidir sobre férias
        if mes["ano_projecao"] == 1:
            assert mes["ferias"] == 0.00
            assert mes["custo_total"] == 6933.34
        else:
            assert mes["ferias"] == 138.89
            assert mes["custo_total"] == 7072.23
        assert mes["extrapolado"] is False
    soma = round(sum(m["custo_total"] for m in res.meses), 2)
    assert res.metadados["total_por_vaga"] == soma == 252933.60
    contagem = pd.Series([m["ano"] for m in res.meses]).value_counts().to_dict()
    assert contagem == {2027: 12, 2028: 12, 2029: 12}
    resumo = resumo_anual(res)
    assert list(resumo.columns) == [c for c in engine.COLUNAS_RESUMO if c != "Incentivo"]
    assert list(resumo["Ano"]) == ["Ano 1 (2027)", "Ano 2 (2028)", "Ano 3 (2029)", "TOTAL"]
    assert resumo.iloc[-1]["Custo por Vaga"] == pytest.approx(252933.60)
    assert res.metadados["custo_medio_mensal"] == 7025.93


# 2. Docente ------------------------------------------------------------------
def test_docente_incentivo_sobre_vencimento() -> None:
    sim = _sim(categoria="Docentes", cargo="Professor Auxiliar", ch="40h")
    res = projetar(sim, _tabela(_linha(PROF_AUX, 4000, gratificacao=999)))
    primeiro = res.meses[0]
    assert primeiro["rbm"] == 4800.00
    assert primeiro["incentivo"] == 800.00
    assert primeiro["gratificacao"] == 0.00
    assert primeiro["ferias"] == 0.00
    assert primeiro["inss"] == 1456.00  # 28% sobre (4800 + 400)
    assert primeiro["custo_total"] == 6656.00
    assert res.metadados["incentivo_descricao"] == "Especialização, 20%"


# 3. Motorista ----------------------------------------------------------------
def test_motorista_sem_gratificacao_com_percentual_da_planilha() -> None:
    tabela = _tabela(_linha(MOTORISTA, 2000, gratificacao=300, percentual=10))
    res = projetar(_sim_motorista(), tabela)
    primeiro = res.meses[0]
    assert primeiro["gratificacao"] == 0.00
    assert primeiro["rbm"] == 2200.00
    assert primeiro["custo_total"] == 2860.00
    assert res.metadados["percentual_motorista_usado"] == 10.0


def test_motorista_sem_gratificacao_percentual_sobrescrito_no_formulario() -> None:
    tabela = _tabela(_linha(MOTORISTA, 2000, gratificacao=300, percentual=10))
    res = projetar(_sim_motorista(percentual_motorista=20.0), tabela)
    for mes in res.meses:
        assert mes["gratificacao"] == 0.00
        assert mes["rbm"] == 2400.00
        if mes["ano_projecao"] == 1:
            assert mes["custo_total"] == 3120.00
        else:
            assert mes["custo_total"] == 3186.67
    assert res.metadados["percentual_motorista_usado"] == 20.0


# 4. Temporário ---------------------------------------------------------------
def test_temporario_prorrogacao_automatica() -> None:
    tabela = _tabela(_linha(TEC_MEDIO_TEMP_40, 4500, gratificacao=500))
    sim = _sim(vinculo="Temporário", cargo="Técnico de Nível Médio", ch="40h")
    res = projetar(sim, tabela)
    assert len(res.meses) == 36
    for mes in res.meses:
        assert set(mes) == set(engine.CHAVES_MES)
        assert not any(chave.startswith("ativ") for chave in mes)
        assert mes["rbm"] == 5000.00
        if mes["ano_projecao"] == 1:
            assert mes["ferias"] == 0.00
            assert mes["custo_total"] == 6500.00
        else:
            assert mes["ferias"] == 138.89
            assert mes["custo_total"] == 6638.89
    assert res.metadados["custo_medio_mensal"] == 6592.59
    assert res.metadados["total_por_vaga"] == 237333.36
    res_julho = projetar(
        _sim(vinculo="Temporário", cargo="Técnico de Nível Médio", ch="40h", mes_ingresso=7),
        tabela,
    )
    assert len(res_julho.meses) == 30
    assert all(m["custo_total"] == 6500.00 for m in res_julho.meses if m["ano_projecao"] == 1)
    assert all(m["custo_total"] == 6638.89 for m in res_julho.meses if m["ano_projecao"] > 1)


# Jornalista ------------------------------------------------------------------
def test_jornalista_25h_caso_de_referencia() -> None:
    tabela = _tabela(_linha(JORNALISTA, 1788.90, gratificacao=1129.65))
    res = projetar(_sim_jornalista(), tabela)
    assert len(res.meses) == 36
    for mes in res.meses:
        assert mes["vencimento_basico"] == 1788.90
        assert mes["gratificacao"] == 1129.65
        assert mes["incentivo"] == 0.00
        assert mes["rbm"] == 2918.55
        assert mes["decimo_terceiro"] == 243.21
        assert mes["inss"] == 632.35
        if mes["ano_projecao"] == 1:
            assert mes["ferias"] == 0.00
            assert mes["custo_total"] == 3794.11
        else:
            assert mes["ferias"] == 81.07
            assert mes["custo_total"] == 3875.18
    assert res.metadados["total_por_vaga"] == 138533.64
    assert res.metadados["nivel_classe_descricao"] == "Classe I"
    resumo = resumo_anual(res)
    ano1 = resumo.iloc[0]
    assert ano1["Ano"] == "Ano 1 (2027)"
    assert ano1["Vencimento Básico"] == pytest.approx(21466.80)
    assert ano1[engine.ROTULO_GRATIFICACAO] == pytest.approx(13555.80)
    assert "Incentivo" not in ano1
    assert ano1["RBM"] == pytest.approx(35022.60)
    assert ano1["Custo por Vaga"] == pytest.approx(45529.32)
    assert resumo.iloc[-1]["Custo por Vaga"] == pytest.approx(138533.64)


def test_jornalista_25h_e_35h_aceitas() -> None:
    tabela = _tabela(
        _linha(JORNALISTA, 1788.90, gratificacao=1129.65),
        _linha((engine.CARGO_JORNALISTA, "Temporário", "35h"), 1788.90, gratificacao=2021.50),
    )
    res_25 = projetar(_sim_jornalista(ch="25h"), tabela)
    assert res_25.meses[0]["rbm"] == 2918.55

    res_35 = projetar(_sim_jornalista(ch="35h"), tabela)
    assert res_35.meses[0]["rbm"] == 3810.40

    with pytest.raises(ValueError):
        projetar(_sim_jornalista(ch="30h"), tabela)

    df_invalido = _df_texto(
        [[engine.CARGO_JORNALISTA, "Temporário", "30h", "1788,90", "1129,65", "",
          "01/01/2027", ""]]
    )
    resultado = data_manager.validar_e_normalizar(df_invalido)
    assert not resultado.ok
    assert (
        "Linha 2: CH '30h' inválida para 'Técnico de Nível Superior (Jornalista)'; "
        "permitida(s): 25h, 35h."
    ) in resultado.erros

    df_ok = _df_texto(
        [
            [engine.CARGO_JORNALISTA, "Temporário", "25H", "1788,90", "1129,65", "",
             "01/01/2027", ""],
            [engine.CARGO_JORNALISTA, "Temporário", "35h", "1788,90", "2021,50", "",
             "01/01/2027", ""],
            ["Técnico Administrativo (Motorista)", "Temporário", "40h", "1610,28", "0",
             "96,96%", "01/04/2026", ""],
        ]
    )
    res_ok = data_manager.validar_e_normalizar(df_ok)
    assert res_ok.ok, res_ok.erros
    assert res_ok.df is not None
    assert res_ok.df.loc[2, "Percentual_Motorista"] == pytest.approx(96.96)


# Composição da RBM e resumo -------------------------------------------------
@pytest.mark.parametrize(
    ("sim", "linha"),
    [
        (
            _sim(categoria="Docentes", cargo="Professor Auxiliar", ch="40h"),
            _linha(PROF_AUX, 4123.45),
        ),
        (_sim(), _linha(TEC_UNIV, 3456.78, gratificacao=987.65)),
        (_sim_motorista(), _linha(MOTORISTA, 2111.11, percentual=17.5)),
        (_sim_jornalista(mes_ingresso=5), _linha(JORNALISTA, 1788.90, gratificacao=1129.65)),
    ],
)
def test_composicao_rbm_fecha(sim: Simulacao, linha: dict[str, Any]) -> None:
    res = projetar(sim, _tabela(linha))
    for mes in res.meses:
        soma = round(mes["vencimento_basico"] + mes["gratificacao"] + mes["incentivo"], 2)
        assert soma == mes["rbm"]
    resumo = resumo_anual(res)
    for _, item in resumo.iterrows():
        componentes = item["Vencimento Básico"]
        if engine.ROTULO_GRATIFICACAO in item:
            componentes += item[engine.ROTULO_GRATIFICACAO]
        if "Incentivo" in item:
            componentes += item["Incentivo"]
        if not sim.cargo.endswith("(Motorista)"):
            assert round(componentes, 2) == round(item["RBM"], 2)


def test_resumo_colunas() -> None:
    res_docente = projetar(
        _sim(categoria="Docentes", cargo="Professor Auxiliar", ch="40h"),
        _tabela(_linha(PROF_AUX, 4000)),
    )
    resumo_doc = resumo_anual(res_docente)
    esperadas_doc = [
        "Ano", "Meses", "Vencimento Básico", "Incentivo", "RBM",
        "13º Salário", "Adicional de Férias", "INSS Patronal", "Custo por Vaga",
        "Quantidade de vagas", "Custo Total (vagas)",
    ]
    assert list(resumo_doc.columns) == esperadas_doc
    assert len(esperadas_doc) == 11

    res_tec = projetar(_sim(), _tabela(_linha(TEC_UNIV, 5000)))
    resumo_tec = resumo_anual(res_tec)
    esperadas_tec = [
        "Ano", "Meses", "Vencimento Básico", "Gratificação", "RBM",
        "13º Salário", "Adicional de Férias", "INSS Patronal", "Custo por Vaga",
        "Quantidade de vagas", "Custo Total (vagas)",
    ]
    assert list(resumo_tec.columns) == esperadas_tec
    assert len(esperadas_tec) == 11


# 5. Vigência múltipla --------------------------------------------------------
def test_vigencia_multipla() -> None:
    tabela = _tabela(
        _linha(TEC_UNIV, 4000, inicio="2027-01-01", fim="2027-04-30"),
        _linha(TEC_UNIV, 4400, inicio="2027-05-01"),
    )
    res = projetar(_sim(), tabela)
    assert [m["vencimento_basico"] for m in res.meses[:4]] == [4000.00] * 4
    assert all(m["vencimento_basico"] == 4400.00 for m in res.meses[4:])


# 6. Ano 1 proporcional, sem Ano 4 ------------------------------------------
@pytest.mark.parametrize(
    ("mes", "esperado", "por_ano"),
    [
        (10, 27, {2027: 3, 2028: 12, 2029: 12}),
        (7, 30, {2027: 6, 2028: 12, 2029: 12}),
        (1, 36, {2027: 12, 2028: 12, 2029: 12}),
        (12, 25, {2027: 1, 2028: 12, 2029: 12}),
    ],
)
def test_ano1_proporcional_sem_ano4(mes: int, esperado: int, por_ano: dict[int, int]) -> None:
    res = projetar(_sim(mes_ingresso=mes), _tabela(_linha(TEC_UNIV, 5000)))
    assert res.metadados["total_meses"] == esperado
    assert len(res.meses) == esperado
    contagem = pd.Series([m["ano"] for m in res.meses]).value_counts().to_dict()
    assert contagem == por_ano
    assert all(m["ano"] != 2030 for m in res.meses)
    assert {m["ano_projecao"] for m in res.meses} == {1, 2, 3}
    assert res.meses[0]["competencia"] == f"{mes:02d}/2027"
    assert res.meses[-1]["competencia"] == "12/2029"


# 7. Erros --------------------------------------------------------------------
def test_vigencia_nao_encontrada() -> None:
    tabela = _tabela(_linha(TEC_UNIV, 5000, inicio="2027-01-01", fim="2027-06-30"))
    with pytest.raises(VigenciaNaoEncontradaError, match="07/2027"):
        projetar(_sim(), tabela)


def test_manter_ultimo_valor_marca_extrapolado() -> None:
    tabela = _tabela(_linha(TEC_UNIV, 5000, inicio="2027-01-01", fim="2027-06-30"))
    res = projetar(_sim(manter_ultimo_valor=True), tabela)
    assert all(not m["extrapolado"] for m in res.meses[:6])
    assert all(m["extrapolado"] for m in res.meses[6:])
    assert all(m["rbm"] == 5000.00 for m in res.meses)


def test_base_vazia_e_combinacao_invalida() -> None:
    with pytest.raises(BaseSalarialVaziaError, match="Importe a base salarial na página 2"):
        projetar(_sim(), _tabela(_linha(PROF_AUX, 4000)))
    with pytest.raises(ValueError):
        projetar(_sim(ch="40h"), _tabela(_linha(TEC_UNIV, 5000)))
    with pytest.raises(ValueError):
        projetar(_sim(mes_ingresso=13), _tabela(_linha(TEC_UNIV, 5000)))
    with pytest.raises(ValueError):
        projetar(_sim(ano_ingresso=1999), _tabela(_linha(TEC_UNIV, 5000)))


# 8. Validação de planilha ---------------------------------------------------
def test_validacao_sobreposicao() -> None:
    df = _df_texto(
        [
            ["Técnico Universitário", "Efetivo", "30h", "5000", "", "", "01/01/2027",
             "30/06/2027"],
            ["Técnico Universitário", "Efetivo", "30h", "5200", "", "", "01/06/2027", ""],
        ]
    )
    resultado = data_manager.validar_e_normalizar(df)
    assert not resultado.ok
    assert resultado.df is None
    assert any("Linha 3: vigência sobrepõe a linha 2" in e for e in resultado.erros)


def test_validacao_cargo_fora_do_catalogo() -> None:
    df = _df_texto([["Reitor", "Efetivo", "40h", "9000", "", "", "01/01/2027", ""]])
    resultado = data_manager.validar_e_normalizar(df)
    assert not resultado.ok
    assert any("fora do catálogo" in e for e in resultado.erros)


def test_validacao_data_fim_anterior_ao_inicio() -> None:
    df = _df_texto(
        [["Técnico Universitário", "Efetivo", "30h", "5000", "", "", "01/05/2027",
          "01/01/2027"]]
    )
    resultado = data_manager.validar_e_normalizar(df)
    assert not resultado.ok
    assert any("anterior" in e for e in resultado.erros)


def test_validacao_numero_brasileiro_e_normalizacoes() -> None:
    df = _df_texto(
        [
            ["Técnico Universitário", "Efetivo", "30h", "1.234,56", "100,50", "", "01/01/2027",
             "31/12/2027"],
            ["Técnico Universitário", "Efetivo", "30h", "1300.10", "", "", "2028-01-01", ""],
            ["Técnico de Nível Médio", "Temporario", "40H", "1234,56", "0", "", "01/01/2027", ""],
            [engine.CARGO_MOTORISTA, "Temporário", "30h", "2000", "300", "15", "01/01/2027", ""],
        ]
    )
    resultado = data_manager.validar_e_normalizar(df)
    assert resultado.ok, resultado.erros
    base = resultado.df
    assert base is not None
    assert base.loc[0, "Vencimento_Basico"] == pytest.approx(1234.56)
    assert base.loc[0, "Valor_Gratificacao"] == pytest.approx(100.50)
    assert base.loc[1, "Vencimento_Basico"] == pytest.approx(1300.10)
    assert base.loc[2, "Vinculo"] == "Temporário"
    assert base.loc[2, "CH"] == "40h"
    assert base.loc[3, "Valor_Gratificacao"] == 0.0
    assert base.loc[3, "Percentual_Motorista"] == 15.0
    assert any("Motorista" in a for a in resultado.avisos)
    assert pd.isna(base.loc[1, "Data_Fim_Vigencia"])
    assert str(base["Data_Inicio_Vigencia"].dtype) == "datetime64[ns]"
    assert base["Vencimento_Basico"].dtype == float


def test_validacao_colunas_ausentes_e_extras() -> None:
    df = pd.DataFrame({"Cargo": ["Técnico Universitário"], "Outra": ["x"]})
    resultado = data_manager.validar_e_normalizar(df)
    assert not resultado.ok
    assert any("ausentes" in e for e in resultado.erros)


def test_template_e_leitura_csv() -> None:
    conteudo = data_manager.gerar_template()
    assert conteudo[:2] == b"PK"
    lido = data_manager.carregar_arquivo(conteudo)
    assert list(lido.columns) == COLUNAS
    csv = (
        ";".join(COLUNAS) + "\n"
        "Técnico Universitário;Efetivo;30h;5.000,00;;;01/01/2027;\n"
    ).encode("latin-1")
    df = data_manager.carregar_arquivo(csv)
    resultado = data_manager.validar_e_normalizar(df)
    assert resultado.ok, resultado.erros
    assert resultado.df is not None
    assert resultado.df.loc[0, "Vencimento_Basico"] == 5000.0


def test_alias_valor_gstu_na_planilha() -> None:
    df = _df_texto(
        [["Técnico Universitário", "Efetivo", "30h", "5000", "250,75", "", "01/01/2027", ""]]
    ).rename(columns={"Valor_Gratificacao": "Valor_GSTU"})
    resultado = data_manager.validar_e_normalizar(df)
    assert resultado.ok, resultado.erros
    assert resultado.df is not None
    assert list(resultado.df.columns) == COLUNAS
    assert resultado.df.loc[0, "Valor_Gratificacao"] == pytest.approx(250.75)
    assert (
        "Coluna 'Valor_GSTU' foi renomeada para 'Valor_Gratificacao'; atualize sua planilha."
        in resultado.avisos
    )
    duplicada = df.assign(Valor_Gratificacao="1")
    resultado_dup = data_manager.validar_e_normalizar(duplicada)
    assert not resultado_dup.ok
    assert any("duplicada" in e for e in resultado_dup.erros)


def test_template_novo_cabecalho() -> None:
    livro = load_workbook(io.BytesIO(data_manager.gerar_template()))
    cabecalho = [c.value for c in livro[data_manager.ABA_TABELAS][1]]
    assert cabecalho == COLUNAS
    assert "Valor_Gratificacao" in cabecalho
    assert "Valor_GSTU" not in cabecalho
    textos = " ".join(
        str(c.value)
        for linha in livro[data_manager.ABA_INSTRUCOES].iter_rows()
        for c in linha
        if c.value is not None
    )
    assert "25h" in textos
    assert "Gratificação" in textos
    assert data_manager.CAMINHO_TEMPLATE.exists()


# 9. Quantidade de vagas -----------------------------------------------------
def test_quantidade_vagas() -> None:
    res = projetar(_sim(quantidade_vagas=3), _tabela(_linha(TEC_UNIV, 5000)))
    assert res.metadados["total_geral"] == pytest.approx(3 * res.metadados["total_por_vaga"])
    resumo = resumo_anual(res)
    assert resumo.iloc[-1]["Custo Total (vagas)"] == pytest.approx(res.metadados["total_geral"])


# Relatórios -----------------------------------------------------------------
def _resultado_exemplo() -> ResultadoSimulacao:
    sim = _sim(observacao="Vaga <urgente> & prioritária\nSegunda linha")
    return projetar(sim, _tabela(_linha(TEC_UNIV, 5000)))


def test_formatar_brl() -> None:
    assert relatorios.formatar_brl(1234.56) == "R$ 1.234,56"
    assert relatorios.formatar_brl(0) == "R$ 0,00"
    assert relatorios.formatar_brl(1234567.891) == "R$ 1.234.567,89"


def test_relatorios_gerados() -> None:
    res = _resultado_exemplo()
    resumo = resumo_anual(res)
    assert relatorios.gerar_pdf(res, resumo)[:4] == b"%PDF"
    excel = relatorios.gerar_excel(res, resumo)
    assert excel[:2] == b"PK"
    abas = pd.read_excel(io.BytesIO(excel), sheet_name=None)
    assert list(abas.keys()) == ["Parâmetros", "Resumo Anual", "Detalhamento Mensal"]
    assert "Incentivo" not in abas["Resumo Anual"].columns
    assert "Quantidade de vagas" in abas["Resumo Anual"].columns
    assert len(abas["Resumo Anual"].columns) == 11
    assert "Ativo" not in abas["Detalhamento Mensal"].columns
    assert engine.ROTULO_GRATIFICACAO in abas["Detalhamento Mensal"].columns

    # Para docente, Incentivo deve estar presente no Resumo Anual (11 colunas) e Gratificação ausente
    res_doc = projetar(
        _sim(categoria="Docentes", cargo="Professor Auxiliar", ch="40h"),
        _tabela(_linha(PROF_AUX, 4000)),
    )
    excel_doc = relatorios.gerar_excel(res_doc, resumo_anual(res_doc))
    abas_doc = pd.read_excel(io.BytesIO(excel_doc), sheet_name=None)
    assert "Incentivo" in abas_doc["Resumo Anual"].columns
    assert engine.ROTULO_GRATIFICACAO not in abas_doc["Resumo Anual"].columns
    assert engine.ROTULO_GRATIFICACAO not in abas_doc["Detalhamento Mensal"].columns
    assert len(abas_doc["Resumo Anual"].columns) == 11

    csv = relatorios.gerar_csv(res).decode("utf-8-sig")
    assert "Observação Automática" in csv
    assert "7072,23" in csv
    assert relatorios.nome_arquivo(res, "pdf").startswith("impacto_tecnico_universitario_")


def test_pdf_com_logos(monkeypatch: pytest.MonkeyPatch, tmp_path: Any) -> None:
    tabela = _tabela(_linha(JORNALISTA, 1788.90, gratificacao=1129.65))
    res = projetar(_sim_jornalista(), tabela)
    resumo = resumo_anual(res)
    pdf = relatorios.gerar_pdf(res, resumo)
    assert pdf[:4] == b"%PDF"
    assert pdf.count(b"/Subtype /Image") >= 3
    rotulos = [p[0] for p in relatorios._parametros(res, incluir_totais=False)]
    rotulos += [p[0] for p in relatorios._parametros(res, incluir_totais=True)]
    assert all("Exigida" not in r for r in rotulos)
    assert "Regime do contrato" not in rotulos
    monkeypatch.setattr(relatorios, "_PASTA_LOGOS", tmp_path)
    pdf_sem_logos = relatorios.gerar_pdf(res, resumo)
    assert pdf_sem_logos[:4] == b"%PDF"
    assert pdf_sem_logos.count(b"/Subtype /Image") == 0


def test_base_inicial_csv_valida() -> None:
    from pathlib import Path
    arquivo = Path("Base inicial 04.11.2026 - Página2.csv")
    if arquivo.exists():
        df = data_manager.carregar_arquivo(arquivo)
        res = data_manager.validar_e_normalizar(df)
        assert res.ok, res.erros
        assert res.df is not None
        assert len(res.df) == 17
        assert "Técnico Administrativo (Motorista)" in res.df["Cargo"].values
        jornalistas = res.df[res.df["Cargo"] == engine.CARGO_JORNALISTA]
        assert set(jornalistas["CH"].values) == {"25h", "35h"}


def test_melhorias_solicitadas_usuario() -> None:
    # 1. Quando tiver Docentes não precisa exibir Gratificação (resumo e detalhamento)
    sim_doc = _sim(
        categoria="Docentes",
        vinculo="Temporário",
        cargo="Professor Substituto (Auxiliar)",
        ch="40h",
        quantidade_vagas=4,
    )
    tabela_doc = _tabela(
        _linha(("Professor Substituto (Auxiliar)", "Temporário", "40h"), 4000)
    )
    res_doc = projetar(sim_doc, tabela_doc)

    resumo_doc = resumo_anual(res_doc)
    assert engine.ROTULO_GRATIFICACAO not in resumo_doc.columns
    assert "Incentivo" in resumo_doc.columns

    mensal_doc = relatorios.tabela_mensal(res_doc)
    assert engine.ROTULO_GRATIFICACAO not in mensal_doc.columns
    assert "Incentivo" in mensal_doc.columns

    # 2. Na coluna de resumo deve incluir na 2ª coluna da direita para esquerda a quantidade de vagas total
    # "Custo por Vaga" * "Quantidade de vagas" = "Custo Total (vagas)"
    cols = list(resumo_doc.columns)
    assert cols[-1] == "Custo Total (vagas)"
    assert cols[-2] == "Quantidade de vagas"
    assert cols[-3] == "Custo por Vaga"

    for _, linha in resumo_doc.iterrows():
        assert linha["Quantidade de vagas"] == 4
        esperado = round(linha["Custo por Vaga"] * linha["Quantidade de vagas"], 2)
        assert round(linha["Custo Total (vagas)"], 2) == pytest.approx(esperado, abs=0.02)

    # 3. Quadro de "Parâmetros da simulação" não precisa exibir "Regime do contrato" (mesmo em Temporário)
    params = relatorios._parametros(res_doc, incluir_totais=False)
    rotulos_params = [p[0] for p in params]
    assert "Regime do contrato" not in rotulos_params

    # Verifica também que PDF e Excel são gerados sem erros
    pdf = relatorios.gerar_pdf(res_doc, resumo_doc)
    assert pdf[:4] == b"%PDF"
    # Cabeçalho: título presente e subtítulo da projeção removido
    assert b"Relat\xc3\xb3rio de Impacto Or\xc3\xa7ament\xc3\xa1rio" in pdf or b"Impacto" in pdf
    assert b"Proje\xc3\xa7\xc3\xa3o de 3 anos (Ano 1" not in pdf
    excel = relatorios.gerar_excel(res_doc, resumo_doc)
    assert excel[:2] == b"PK"


def test_aliquotas_patronais_efetivo_e_temporario() -> None:
    # Efetivo: 28% (RPPS - Lei 11.357/2009 e Lei 14.520/2022)
    assert engine.aliquota_patronal_vinculo("Efetivo") == Decimal("0.28")
    assert engine.ALIQ_PATRONAL_EFETIVO == Decimal("0.28")

    # Temporário: 20% (RGPS/INSS)
    assert engine.aliquota_patronal_vinculo("Temporário") == Decimal("0.20")
    assert engine.ALIQ_PATRONAL_TEMPORARIO == Decimal("0.20")

    # Teste de cálculo: vencimento 3000
    # RBM = 3000, 13º = 250, férias = 0 no ano 1 (e não incide encargo patronal sobre férias)
    # Base encargo = 3250
    tabela_efetivo = _tabela(_linha(TEC_UNIV, 3000))
    res_efetivo = projetar(_sim(), tabela_efetivo)
    primeiro_efetivo = res_efetivo.meses[0]
    # Encargo 28% de 3250 = 910.00
    assert primeiro_efetivo["inss"] == 910.00
    assert primeiro_efetivo["custo_total"] == 4160.00
    assert res_efetivo.metadados["aliquota_patronal"] == 0.28

    tabela_temp = _tabela(_linha(("Técnico de Nível Médio", "Temporário", "40h"), 3000))
    sim_temp = _sim(vinculo="Temporário", cargo="Técnico de Nível Médio", ch="40h")
    res_temp = projetar(sim_temp, tabela_temp)
    primeiro_temp = res_temp.meses[0]
    # Encargo 20% de 3250 = 650.00
    assert primeiro_temp["inss"] == 650.00
    assert primeiro_temp["custo_total"] == 3900.00
    assert res_temp.metadados["aliquota_patronal"] == 0.20


def test_relatorio_pdf_sintetico_layout_uefs() -> None:
    # Simulação Docente (Professor Auxiliar, Efetivo 28%)
    sim = _sim(
        categoria="Docentes",
        vinculo="Efetivo",
        cargo="Professor Auxiliar",
        ch="40h",
        quantidade_vagas=2,
        mes_ingresso=5,
        ano_ingresso=2027,
    )
    tabela = _tabela(_linha(PROF_AUX, 4000))
    res = projetar(sim, tabela)
    resumo = resumo_anual(res)

    # 1. Geração padrão (tipo="sintetico")
    pdf_sintetico = relatorios.gerar_pdf(res, resumo, tipo="sintetico")
    assert pdf_sintetico[:4] == b"%PDF"

    # Verifica logos e informações institucionais no topo do cabeçalho
    tabela_topo = relatorios._tabela_logos_sintetico()
    assert tabela_topo is not None
    textos_centro = [p.text for p in tabela_topo._cellvalues[0][1]]
    assert "Universidade Estadual de Feira de Santana – UEFS" in textos_centro[0]
    assert "Pró-Reitoria de Gestão e Desenvolvimento de Pessoas – PGDP" in textos_centro[1]

    assert pdf_sintetico.count(b"/Subtype /Image") >= 3

    # Verifica metadados no PDF sintético
    assert b"Demonstrativo Sint" in pdf_sintetico or b"Impacto" in pdf_sintetico

    # Verifica linhas da matriz sintética (Tabela 2)
    tabela2 = relatorios._tabelas_matriz_impacto_sintetico(res)
    assert tabela2._cellvalues[0][0] == "Composição"
    assert tabela2._cellvalues[0][4] == "Total (3 Anos)"
    assert "Rendimento Básico Mensal" in tabela2._cellvalues[1][0].text
    assert "1 vaga" in tabela2._cellvalues[1][0].text
    assert "SUBTOTAL DE PROVENTOS (1 VAGA)" in tabela2._cellvalues[4][0]
    assert "13º Salário" in tabela2._cellvalues[6][0]
    assert "Total Mensal" in tabela2._cellvalues[7][0]
    assert "Total Anual" in tabela2._cellvalues[8][0]
    assert tabela2._cellvalues[9][0] == "SUBTOTAL DE ENCARGOS (1 VAGA)"
    assert "CUSTO TOTAL DE 1 VAGA" in tabela2._cellvalues[10][0]
    assert "Quantidade de Vaga(s)" in tabela2._cellvalues[11][0]
    assert "TOTAL GERAL (02 VAGAS)" in tabela2._cellvalues[12][0]

    # Verifica que itens removidos não constam no PDF sintético
    import shutil
    import subprocess
    if shutil.which("pdftotext"):
        proc = subprocess.run(["pdftotext", "-", "-"], input=pdf_sintetico, capture_output=True, check=True)
        txt = proc.stdout.decode("utf-8")
        assert "regime de 40h:" not in txt
        assert "s/ Férias (1/3)" not in txt
        assert "Custo Unitário (1 Servidor)" not in txt
        assert "Especificação da Base de Incidência" not in txt

    # Verifica parâmetros de cálculo no relatório sintético
    tabela_param = relatorios._tabela_parametros_sintetico(res)
    assert tabela_param is not None
    textos_param = [c.text for linha in tabela_param._cellvalues for c in linha if hasattr(c, "text")]
    assert any("Parâmetros de cálculo" in t or "Parâmetros de calculo" in t for t in textos_param)
    assert any("Categoria:" in t for t in textos_param)
    assert any("Vínculo:" in t for t in textos_param)
    assert any("Cargo Professor:" in t for t in textos_param)
    assert any("Nível/Classe:" in t for t in textos_param)
    assert any("Carga Horária:" in t for t in textos_param)
    assert any("Quantidade de vaga(s):" in t for t in textos_param)

    # 2. Geração da Memória de Cálculo (tipo="memoria_calculo")
    pdf_memoria = relatorios.gerar_pdf(res, resumo, tipo="memoria_calculo")
    assert pdf_memoria[:4] == b"%PDF"
    assert pdf_memoria.count(b"/Subtype /Image") >= 3
    assert len(pdf_memoria) > 0
    import shutil
    import subprocess
    if shutil.which("pdftotext"):
        proc = subprocess.run(["pdftotext", "-", "-"], input=pdf_memoria, capture_output=True, check=True)
        txt = proc.stdout.decode("utf-8")
        assert "MEMÓRIA DE CÁLCULO" in txt
        assert "Não incide encargo patronal" in txt
        assert "RBM ÷ 12" in txt

    # 3. Geração detalhada (tipo="detalhado")
    pdf_detalhado = relatorios.gerar_pdf(res, resumo, tipo="detalhado")
    assert pdf_detalhado[:4] == b"%PDF"
    assert pdf_detalhado.count(b"/Subtype /Image") >= 3
    assert len(pdf_detalhado) > 0

    # 4. Teste para Temporário (alíquota 20%)
    sim_temp = _sim(
        categoria="Técnicos",
        vinculo="Temporário",
        cargo="Técnico de Nível Médio",
        ch="40h",
        quantidade_vagas=1,
    )
    tabela_temp = _tabela(_linha(("Técnico de Nível Médio", "Temporário", "40h"), 3000))
    res_temp = projetar(sim_temp, tabela_temp)
    resumo_temp = resumo_anual(res_temp)
    pdf_temp = relatorios.gerar_pdf(res_temp, resumo_temp, tipo="sintetico")
    assert pdf_temp[:4] == b"%PDF"
    pdf_memoria_temp = relatorios.gerar_pdf(res_temp, resumo_temp, tipo="memoria_calculo")
    assert pdf_memoria_temp[:4] == b"%PDF"


def test_historico_encadeamento_temporal_automatico() -> None:
    # Dois registros sucessivos sem data fim preenchida
    base_inicial = _tabela(
        _linha(TEC_UNIV, 4000, inicio="2024-01-01", fim=None),
        _linha(TEC_UNIV, 4500, inicio="2026-01-01", fim=None),
        _linha(PROF_AUX, 6000, inicio="2025-01-01", fim=None),
    )
    encadeada = data_manager.encadear_vigencias(base_inicial)

    # Verifica Técnico Universitário
    tecs = encadeada[encadeada["Cargo"] == "Técnico Universitário"].sort_values("Data_Inicio_Vigencia").reset_index(drop=True)
    assert len(tecs) == 2
    # Registro anterior (2024) fechado em 2025-12-31 (1 dia antes de 2026-01-01)
    assert tecs.loc[0, "Data_Fim_Vigencia"] == pd.Timestamp("2025-12-31")
    # Registro mais recente mantém vigência aberta
    assert pd.isna(tecs.loc[1, "Data_Fim_Vigencia"])

    # Verifica Professor Auxiliar (registro único permanece aberto)
    profs = encadeada[encadeada["Cargo"] == "Professor Auxiliar"].reset_index(drop=True)
    assert len(profs) == 1
    assert pd.isna(profs.loc[0, "Data_Fim_Vigencia"])


def test_historico_identificacao_de_conflitos() -> None:
    base_existente = _tabela(
        _linha(TEC_UNIV, 4000, inicio="2024-01-01", fim="2025-12-31"),
        _linha(TEC_UNIV, 4500, inicio="2026-01-01", fim=None),
    )
    # Planilha nova com 1 registro conflitante (2026-01-01) e 1 registro novo (2027-01-01)
    nova_planilha = _tabela(
        _linha(TEC_UNIV, 4800, inicio="2026-01-01", fim=None),  # conflito de chave primária
        _linha(PROF_AUX, 7000, inicio="2027-01-01", fim=None),  # cargo novo sem conflito
    )

    conflitos = data_manager.identificar_conflitos(base_existente, nova_planilha)
    assert len(conflitos) == 1
    assert conflitos.iloc[0]["Cargo"] == "Técnico Universitário"
    assert conflitos.iloc[0]["Vencimento_Basico"] == 4800.0

    comp = data_manager.detalhar_conflitos(base_existente, nova_planilha)
    assert len(comp) == 1
    assert comp.iloc[0]["Vencimento_Basico_Atual"] == 4500.0
    assert comp.iloc[0]["Vencimento_Basico_Novo"] == 4800.0


def test_historico_merge_com_sobrescrita_e_motor() -> None:
    base_existente = _tabela(
        _linha(TEC_UNIV, 4000, inicio="2024-01-01", fim=None),
    )
    # Nova planilha adiciona salário novo para 2026-01-01
    nova_planilha = _tabela(
        _linha(TEC_UNIV, 5000, inicio="2026-01-01", fim=None),
    )

    base_consolidada = data_manager.mesclar_bases(base_existente, nova_planilha, sobrescrever_conflitos=True)
    assert len(base_consolidada) == 2

    # Verifica que 2024 foi fechado automaticamente
    r1 = base_consolidada[base_consolidada["Data_Inicio_Vigencia"] == pd.Timestamp("2024-01-01")].iloc[0]
    assert r1["Data_Fim_Vigencia"] == pd.Timestamp("2025-12-31")
    assert r1["Vencimento_Basico"] == 4000.0

    # Agora simula sobrescrita para corrigir o salário de 2026 para 5200.0
    correcao_planilha = _tabela(
        _linha(TEC_UNIV, 5200, inicio="2026-01-01", fim=None),
    )
    base_atualizada = data_manager.mesclar_bases(base_consolidada, correcao_planilha, sobrescrever_conflitos=True)
    assert len(base_atualizada) == 2
    r2 = base_atualizada[base_atualizada["Data_Inicio_Vigencia"] == pd.Timestamp("2026-01-01")].iloc[0]
    assert r2["Vencimento_Basico"] == 5200.0

    # Motor executando na base histórica consolidada
    # Ingresso em 2025: usa a vigência de 4000 em 2025 e de 5200 a partir de 2026
    sim = _sim(
        categoria="Técnicos",
        cargo="Técnico Universitário",
        vinculo="Efetivo",
        ch="30h",
        mes_ingresso=1,
        ano_ingresso=2025,
        quantidade_vagas=1,
    )
    resultado = projetar(sim, base_atualizada)
    # Ano 2025 (meses 1 a 12): vencimento 4000
    for m in resultado.meses[:12]:
        assert m["vencimento_basico"] == 4000.0
    # Ano 2026 (meses 13 a 24): vencimento 5200
    for m in resultado.meses[12:24]:
        assert m["vencimento_basico"] == 5200.0


def test_sqlite_persistencia_merge_e_encadeamento(tmp_path: Any) -> None:
    db_path = tmp_path / "test_db.sqlite"

    # 1. Inserção inicial de 2 vigências sucessivas
    df_inicial = _tabela(
        _linha(TEC_UNIV, 4000, inicio="2024-01-01", fim=None),
        _linha(TEC_UNIV, 4500, inicio="2026-01-01", fim=None),
        _linha(PROF_AUX, 6000, inicio="2025-01-01", fim=None),
    )
    qtd_inserida, total_banco = data_manager.salvar_base_sqlite(df_inicial, caminho=db_path)
    assert qtd_inserida == 3
    assert total_banco == 3

    # 2. Verifica encadeamento automático do SQLite
    hist_tec = data_manager.obter_historico_cargo("Técnico Universitário", caminho=db_path)
    assert len(hist_tec) == 2
    # Registro de 2024 fechou automaticamente em 2025-12-31 (1 dia antes de 2026-01-01)
    assert hist_tec.loc[0, "Data_Fim_Vigencia"] == pd.Timestamp("2025-12-31")
    assert pd.isna(hist_tec.loc[1, "Data_Fim_Vigencia"])

    # 3. Sobrescrita de registro com mesma chave primária
    df_correcao = _tabela(
        _linha(TEC_UNIV, 4800, inicio="2026-01-01", fim=None),
    )
    data_manager.salvar_base_sqlite(df_correcao, caminho=db_path)
    hist_tec_atualizado = data_manager.obter_historico_cargo("Técnico Universitário", caminho=db_path)
    assert len(hist_tec_atualizado) == 2
    # Vencimento atualizado para 4800
    assert hist_tec_atualizado.loc[1, "Vencimento_Basico"] == 4800.0

    # 4. Leitura da tabela vigente em uma data específica
    vigente_2024 = data_manager.obter_tabela_vigente("2024-06-01", caminho=db_path)
    linha_tec_2024 = vigente_2024[vigente_2024["Cargo"] == "Técnico Universitário"].iloc[0]
    assert linha_tec_2024["Vencimento_Basico"] == 4000.0

    vigente_2026 = data_manager.obter_tabela_vigente("2026-06-01", caminho=db_path)
    linha_tec_2026 = vigente_2026[vigente_2026["Cargo"] == "Técnico Universitário"].iloc[0]
    assert linha_tec_2026["Vencimento_Basico"] == 4800.0

    # 5. Listar cargos cadastrados
    cargos = data_manager.listar_cargos_cadastrados(caminho=db_path)
    assert "Professor Auxiliar" in cargos
    assert "Técnico Universitário" in cargos

    # 6. Limpeza do banco
    data_manager.limpar_banco_sqlite(caminho=db_path)
    assert len(data_manager.obter_base_completa(caminho=db_path)) == 0


def test_sqlite_cadastro_manual_e_edicao_completa(tmp_path: Any) -> None:
    db_path = tmp_path / "test_manual_db.sqlite"

    # 1. Cadastro manual de vigência para Professor Auxiliar
    res1 = data_manager.salvar_registro_individual(
        cargo="Professor Auxiliar",
        vinculo="Efetivo",
        ch="40h",
        vencimento_basico=5000.0,
        gratificacao=0.0,
        data_inicio_vigencia="2025-01-01",
        caminho=db_path,
    )
    assert res1.ok

    # 2. Cadastro manual de reajuste para 2026-01-01 (deve encadear a primeira para 2025-12-31)
    res2 = data_manager.salvar_registro_individual(
        cargo="Professor Auxiliar",
        vinculo="Efetivo",
        ch="40h",
        vencimento_basico=5500.0,
        data_inicio_vigencia="2026-01-01",
        caminho=db_path,
    )
    assert res2.ok

    hist = data_manager.obter_historico_cargo("Professor Auxiliar", caminho=db_path)
    assert len(hist) == 2
    assert hist.loc[0, "Data_Fim_Vigencia"] == pd.Timestamp("2025-12-31")
    assert pd.isna(hist.loc[1, "Data_Fim_Vigencia"])

    # 3. Exclusão de registro específico e reencadeamento
    data_manager.excluir_registro_sqlite(
        cargo="Professor Auxiliar",
        vinculo="Efetivo",
        ch="40h",
        data_inicio_vigencia="2026-01-01",
        caminho=db_path,
    )
    hist_pos_delete = data_manager.obter_historico_cargo("Professor Auxiliar", caminho=db_path)
    assert len(hist_pos_delete) == 1
    # Ao deletar o posterior, o registro anterior volta a ser o último e deve reabrir a vigência
    assert hist_pos_delete.loc[0, "Vencimento_Basico"] == 5000.0


def test_novas_regras_ferias_encargos_e_parametros_sintetico() -> None:
    # 1. Regra de férias (0 no ano 1, calculadas a partir do ano 2)
    # 2. Regra de encargos patronais (não incide sobre 1/3 de férias)
    sim = _sim(
        categoria="Docentes",
        vinculo="Efetivo",
        cargo="Professor Assistente",
        ch="40h",
        quantidade_vagas=3,
        mes_ingresso=1,
        ano_ingresso=2027,
    )
    tabela = _tabela(_linha(("Professor Assistente", "Efetivo", "40h"), 6000))
    res = projetar(sim, tabela)
    assert len(res.meses) == 36

    for mes in res.meses:
        # Base RBM = 6000 + 2400 (Mestrado 40% para Assistente) = 8400.00
        assert mes["rbm"] == 8400.00
        assert mes["decimo_terceiro"] == 700.00
        # Encargo patronal: 28% sobre (8400 + 700) = 2548.00 (SEM 1/3 de férias)
        assert mes["inss"] == 2548.00

        if mes["ano_projecao"] == 1:
            assert mes["ferias"] == 0.00
            assert mes["custo_total"] == 8400.00 + 700.00 + 0.00 + 2548.00 == 11648.00
        else:
            # A partir do ano 2: Férias = 8400 / 36 = 233.33
            assert mes["ferias"] == 233.33
            assert mes["custo_total"] == 8400.00 + 700.00 + 233.33 + 2548.00 == 11881.33

    # 3. Relatório Sintético com Parâmetros da Simulação
    resumo = resumo_anual(res)
    pdf = relatorios.gerar_pdf(res, resumo, tipo="sintetico")
    assert pdf[:4] == b"%PDF"

    tab_param = relatorios._tabela_parametros_sintetico(res)
    assert tab_param is not None
    textos = [c.text for linha in tab_param._cellvalues for c in linha if hasattr(c, "text")]
    assert any("Parâmetros de cálculo" in t or "Parâmetros de calculo" in t for t in textos)
    assert any("Categoria:" in t for t in textos)
    assert any("Vínculo:" in t for t in textos)
    assert any("Cargo Professor:" in t for t in textos)
    assert any("Nível/Classe:" in t for t in textos)
    assert any("Carga Horária:" in t for t in textos)
    assert any("Quantidade de vaga(s):" in t for t in textos)
    assert any("* Início em 01 de Janeiro de 2027" in t or "* Inicio em 01 de Janeiro de 2027" in t for t in textos)

    # Teste com ingresso em Julho de 2027
    sim_julho = _sim(
        categoria="Docentes",
        vinculo="Efetivo",
        cargo="Professor Assistente",
        ch="40h",
        mes_ingresso=7,
        ano_ingresso=2027,
    )
    res_julho = projetar(sim_julho, tabela)
    tab_param_julho = relatorios._tabela_parametros_sintetico(res_julho)
    textos_julho = [c.text for linha in tab_param_julho._cellvalues for c in linha if hasattr(c, "text")]
    assert any("* Início em 01 de Julho de 2027" in t or "* Inicio em 01 de Julho de 2027" in t for t in textos_julho)

    # Verifica que para Docentes a coluna de composição da remuneração é "INCENTIVOS"
    tab1_doc = relatorios._tabela_remuneracao_base_sintetico(res)
    assert tab1_doc._cellvalues[0][2] == "INCENTIVOS"

    # Verifica que para Técnicos a coluna de composição da remuneração é "GRATIFICAÇÃO"
    sim_tec = _sim(categoria="Técnicos", cargo="Técnico Universitário", ch="30h")
    tabela_tec = _tabela(_linha(TEC_UNIV, 5000))
    res_tec = projetar(sim_tec, tabela_tec)
    tab1_tec = relatorios._tabela_remuneracao_base_sintetico(res_tec)
    assert tab1_tec._cellvalues[0][2] == "GRATIFICAÇÃO"


def test_setores_opcoes_e_rodape() -> None:
    """Verifica as 4 siglas de setor (GRH, NSA, SPAG, SAAP) e sua exibição no rodapé dos relatórios."""
    import subprocess
    import shutil

    assert set(engine.SIGLAS_SETORES) == {"GRH", "NSA", "SPAG", "SAAP"}
    assert "Gerência de Recursos Humanos" in engine.SETORES["GRH"]
    assert "Núcleo de Seleção e Alocação" in engine.SETORES["NSA"]
    assert "Subgerência de Pagamentos" in engine.SETORES["SPAG"]
    assert "Subgerência de Admissão e Acompanhamento de Pessoas" in engine.SETORES["SAAP"]

    tabela = _tabela(_linha(("Professor Auxiliar", "Efetivo", "40h"), 5000))

    for sigla in engine.SIGLAS_SETORES:
        sim = _sim(
            categoria="Docentes",
            vinculo="Efetivo",
            cargo="Professor Auxiliar",
            ch="40h",
            setor=sigla,
        )
        res = projetar(sim, tabela)
        assert res.metadados["setor"] == sigla
        assert engine.SETORES[sigla] in res.metadados["setor_descricao"]

        resumo = resumo_anual(res)
        pdf_sintetico = relatorios.gerar_pdf(res, resumo, tipo="sintetico")
        assert pdf_sintetico[:4] == b"%PDF"

        if shutil.which("pdftotext"):
            proc = subprocess.run(["pdftotext", "-", "-"], input=pdf_sintetico, capture_output=True, check=True)
            txt = proc.stdout.decode("utf-8")
            assert sigla in txt
            assert engine.SETORES[sigla] in txt

        # Verifica nos parâmetros
        params = relatorios._parametros(res, incluir_totais=False)
        rotulos = [p[0] for p in params]
        assert "Setor responsável" in rotulos









