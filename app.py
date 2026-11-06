"""Interface Streamlit do ImpactoUEFS – Sistema de Impacto Orçamentário da UEFS."""
from __future__ import annotations

import importlib
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

import data_manager
import engine
import relatorios

importlib.reload(data_manager)
importlib.reload(engine)
importlib.reload(relatorios)

from engine import (
    BaseSalarialVaziaError,
    ResultadoSimulacao,
    Simulacao,
    VigenciaNaoEncontradaError,
)
from relatorios import formatar_brl

PAGINA_SIMULACAO = "Simulação de Impacto"
PAGINA_IMPORTACAO = "Importação de Base Salarial"
MESES_PT: tuple[str, ...] = (
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
)
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------------------------------------------------------------------------
# Estado e utilitários
# ---------------------------------------------------------------------------
def _obter_base_ativa() -> pd.DataFrame | None:
    """Retorna a base salarial consolidada (da sessão ou carregando do SQLite)."""
    base: pd.DataFrame | None = st.session_state.get("base_salarial")
    if base is None or base.empty:
        try:
            base_db = data_manager.obter_base_completa()
            if not base_db.empty:
                st.session_state["base_salarial"] = base_db
                return base_db
        except Exception:
            pass
        return None
    return base


def _inicializar_estado() -> None:
    """Cria as chaves de sessão usadas pelo aplicativo."""
    padroes: dict[str, Any] = {
        "base_salarial": None,
        "resultado": None,
        "nome_arquivo_base": None,
        "validacao": None,
        "ultimo_upload": None,
        "versao_uploader": 0,
    }
    for chave, valor in padroes.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor


def _padrao(chave: str, valor: Any) -> None:
    """Define o valor inicial de um widget antes de criá-lo."""
    if chave not in st.session_state:
        st.session_state[chave] = valor


def _selectbox_dependente(rotulo: str, opcoes: list[str], chave: str) -> str:
    """Selectbox cujas opções dependem de outro campo; reseta valor inválido."""
    if st.session_state.get(chave) not in opcoes:
        st.session_state[chave] = opcoes[0]
    return str(st.selectbox(rotulo, opcoes, key=chave))


def _proximo_mes(hoje: date) -> tuple[int, int]:
    """Mês e ano seguintes à data informada."""
    if hoje.month == 12:
        return 1, hoje.year + 1
    return hoje.month + 1, hoje.year


def _limpar_resultado() -> None:
    """Descarta o último resultado calculado."""
    st.session_state["resultado"] = None


# ---------------------------------------------------------------------------
# Página 1 – Simulação
# ---------------------------------------------------------------------------
def _percentual_planilha(
    base: pd.DataFrame, cargo: engine.CargoDef, ch: str, mes: int, ano: int
) -> float | None:
    """Percentual de incentivo do Motorista vigente na planilha no mês de ingresso."""
    filtrada = data_manager.filtrar_tabela(base, cargo.cargo, cargo.vinculo, ch)
    if filtrada.empty:
        return None
    referencia = pd.Timestamp(year=ano, month=mes, day=1)
    linha, _ = engine.selecionar_vigencia(filtrada, referencia, True)
    if linha is None:
        linha = filtrada.iloc[0]
    return float(engine.para_decimal(linha["Percentual_Motorista"]))


def _calcular(sim: Simulacao, base: pd.DataFrame) -> None:
    """Executa a projeção e guarda o resultado na sessão, exibindo erros amigáveis."""
    try:
        st.session_state["resultado"] = engine.projetar(sim, base)
    except VigenciaNaoEncontradaError as exc:
        _limpar_resultado()
        st.error(str(exc))
    except BaseSalarialVaziaError as exc:
        _limpar_resultado()
        st.error(str(exc))
    except ValueError as exc:
        _limpar_resultado()
        st.error(str(exc))
    except Exception:  # noqa: BLE001 - nunca exibir traceback ao usuário
        _limpar_resultado()
        st.error("Erro inesperado ao calcular o impacto. Verifique a base e os parâmetros.")


def pagina_simulacao() -> None:
    """Formulário de simulação e exibição do resultado."""
    st.title("Simulação de Impacto Orçamentário")
    st.caption(
        "Projeção de 3 anos calendário: Ano 1 (do mês de ingresso a dezembro) + Anos 2 e 3. "
        "Benefícios indenizatórios não são considerados."
    )
    base = _obter_base_ativa()
    if base is None:
        st.warning(
            "Nenhuma base salarial carregada. Acesse **Importação de Base Salarial** no menu "
            "lateral, baixe a planilha modelo, preencha-a e envie o arquivo."
        )
        st.stop()

    st.subheader("Cargo")
    categoria_previa = st.session_state.get("categoria", engine.listar_categorias()[0])
    vinculo_previo = st.session_state.get("vinculo", engine.listar_vinculos(categoria_previa)[0])
    col1, col2 = st.columns(2)
    with col1:
        categoria = _selectbox_dependente("Categoria", engine.listar_categorias(), "categoria")
        cargo_nome = _selectbox_dependente(
            "Cargo", engine.listar_cargos(categoria, vinculo_previo), "cargo"
        )
    with col2:
        vinculo = _selectbox_dependente("Vínculo", engine.listar_vinculos(categoria), "vinculo")
        ch = _selectbox_dependente(
            "Carga Horária", engine.listar_chs(categoria, vinculo, cargo_nome), "ch"
        )
    cargo = engine.obter_cargo(categoria, vinculo, cargo_nome)
    if cargo is None:
        st.error("Combinação de cargo inválida.")
        st.stop()

    col3, col4 = st.columns(2)
    col3.text_input("Nível/Classe", value=cargo.nivel_classe, disabled=True)
    col4.text_input("Incentivo/Adicional", value=cargo.incentivo_descricao, disabled=True)
    if cargo.is_jornalista:
        if len(cargo.chs) == 1:
            chs_texto = f"carga horária única de {cargo.chs[0]}"
        else:
            chs_texto = f"cargas horárias: {', '.join(cargo.chs)}"
        st.caption(f"Cargo com base legal própria: {engine.LEI_JORNALISTA} ({chs_texto}).")
    if cargo.is_motorista:
        st.caption(
            f"O Motorista não recebe {engine.ROTULO_GRATIFICACAO}; aplica-se apenas o "
            "percentual de incentivo."
        )

    st.subheader("Detalhes da vaga")
    observacao = st.text_area(
        "Observações (opcional)",
        key="observacao",
        help="Informações adicionais que serão impressas nos relatórios.",
    )

    mes_padrao, ano_padrao = _proximo_mes(date.today())
    _padrao("mes_ingresso", mes_padrao)
    _padrao("ano_ingresso", ano_padrao)
    _padrao("quantidade_vagas", 1)
    col5, col6, col7 = st.columns(3)
    mes_ingresso = int(
        col5.selectbox(
            "Mês de ingresso",
            list(range(1, 13)),
            format_func=lambda m: MESES_PT[m - 1],
            key="mes_ingresso",
        )
    )
    ano_ingresso = int(
        col6.number_input(
            "Ano de ingresso", min_value=2000, max_value=2100, step=1, key="ano_ingresso"
        )
    )
    quantidade = int(
        col7.number_input("Quantidade de vagas", min_value=1, step=1, key="quantidade_vagas")
    )

    percentual_motorista: float | None = None
    if cargo.is_motorista:
        valor_planilha = _percentual_planilha(base, cargo, ch, mes_ingresso, ano_ingresso)
        texto_planilha = relatorios.formatar_percentual(valor_planilha)
        entrada = st.number_input(
            "Percentual de Incentivo do Motorista (%)",
            min_value=0.0,
            max_value=100.0,
            value=None,
            step=0.5,
            format="%.2f",
            key="percentual_motorista",
            placeholder=f"Planilha: {texto_planilha}",
            help=(
                f"Valor da planilha no mês de ingresso: {texto_planilha}. Deixe vazio para usar "
                "o percentual da planilha vigente em cada mês; se preenchido, este valor "
                "prevalece em todos os meses."
            ),
        )
        percentual_motorista = None if entrada is None else float(entrada)

    if cargo.is_temporario:
        st.caption(
            "Contrato temporário: considera-se a prorrogação automática durante todo o "
            "período projetado (Ano 1 parcial + Anos 2 e 3)."
        )

    manter = st.checkbox(
        "Manter último valor vigente para meses sem tabela salarial", key="manter_ultimo_valor"
    )

    if st.button("Calcular impacto", type="primary", key="btn_calcular"):
        simulacao = Simulacao(
            categoria=categoria,
            vinculo=vinculo,
            cargo=cargo_nome,
            ch=ch,
            observacao=observacao or "",
            mes_ingresso=mes_ingresso,
            ano_ingresso=ano_ingresso,
            quantidade_vagas=quantidade,
            percentual_motorista=percentual_motorista,
            manter_ultimo_valor=bool(manter),
        )
        _calcular(simulacao, base)

    resultado: ResultadoSimulacao | None = st.session_state["resultado"]
    if resultado is not None:
        _renderizar_resultado(resultado)


def _renderizar_resultado(res: ResultadoSimulacao) -> None:
    """Exibe métricas, resumo anual, detalhamento mensal e downloads."""
    md = res.metadados
    resumo = engine.resumo_anual(res)
    st.divider()
    st.subheader("Resultado da projeção")
    st.caption(
        f"{md['cargo']} · {md['vinculo']} · {md['ch']} · ingresso "
        f"{engine.competencia(md['mes_ingresso'], md['ano_ingresso'])} · "
        f"{relatorios.periodo_projetado(res)} · {md['quantidade_vagas']} vaga(s)"
    )

    col1, col2, col3 = st.columns(3)
    col1.metric(
        "Impacto Total (Ano 1 parcial + Anos 2 e 3)",
        formatar_brl(md["total_geral"]),
        help="Soma do período projetado multiplicada pela quantidade de vagas.",
    )
    col2.metric("Custo por vaga (período)", formatar_brl(md["total_por_vaga"]))
    col3.metric("Custo médio mensal (por vaga)", formatar_brl(md["custo_medio_mensal"]))

    anos = resumo[resumo["Ano"] != "TOTAL"]
    colunas_ano = st.columns(max(len(anos), 1))
    for coluna, (_, linha) in zip(colunas_ano, anos.iterrows()):
        coluna.metric(
            str(linha["Ano"]),
            formatar_brl(linha["Custo Total (vagas)"]),
            help=f"{int(linha['Meses'])} meses; por vaga: {formatar_brl(linha['Custo por Vaga'])}",
        )

    st.markdown("**Resumo por ano**")
    st.dataframe(
        relatorios.formatar_colunas_brl(resumo, relatorios.COLUNAS_MONETARIAS_RESUMO),
        hide_index=True,
    )

    with st.expander("Detalhamento mensal (período projetado)"):
        mensal = relatorios.tabela_mensal(res)
        st.dataframe(
            relatorios.formatar_colunas_brl(mensal, relatorios.COLUNAS_MONETARIAS_MENSAL),
            hide_index=True,
        )

    with st.expander("Embasamento legal dos encargos patronais (Efetivo 28% vs. Temporário 20%)"):
        st.markdown(
            "**1. Servidores Efetivos (28% – RPPS / Estado da Bahia):**\n"
            "- **Base e Alíquota:** Art. 68 da Lei Estadual nº 11.357/2009 c/c Lei Estadual nº 14.520/2022 "
            "(cota patronal de 28% devida pelo Estado aos fundos FUNPREV/BAPREV).\n"
            "- **Férias e Terço (1/3) Constitucional:** Natureza remuneratória habitual (art. 92 da Lei Estadual nº 6.677/1994) "
            "e repercussão geral pacificada pelo STF no **Tema 985 (RE 1.072.485)**.\n\n"
            "**2. Servidores Temporários / REDA (20% – RGPS / INSS):**\n"
            "- **Vínculo e Alíquota:** Art. 40, § 13 da CF/88 (vinculação obrigatória ao RGPS/INSS) c/c art. 22, I da "
            "Lei Federal nº 8.212/1991 (alíquota patronal padrão de 20%).\n"
            "- **Férias e Terço (1/3) Constitucional:** Art. 214, § 4º do Decreto Federal nº 3.048/1999 (férias gozadas integram o "
            "salário-de-contribuição) e jurisprudência vinculante do STF no **Tema 985**."
        )

    if md.get("observacao"):
        st.caption(f"Observações do usuário: {md['observacao']}")

    st.markdown("**Exportar resultado**")
    try:
        pdf_sintetico = relatorios.gerar_pdf(res, resumo, tipo="sintetico")
        pdf_detalhado = relatorios.gerar_pdf(res, resumo, tipo="detalhado")
        excel = relatorios.gerar_excel(res, resumo)
        csv = relatorios.gerar_csv(res)
    except Exception:  # noqa: BLE001 - nunca exibir traceback ao usuário
        st.error("Não foi possível gerar os arquivos de exportação.")
        return
    col_pdf1, col_pdf2, col_xlsx, col_csv = st.columns(4)
    col_pdf1.download_button(
        "Baixar PDF (Sintético)", pdf_sintetico,
        file_name=relatorios.nome_arquivo(res, "pdf"),
        mime="application/pdf", key="download_pdf_sintetico",
        help="Relatório sintético no modelo oficial de demonstrativo da UEFS.",
    )
    col_pdf2.download_button(
        "Baixar PDF (Detalhado)", pdf_detalhado,
        file_name=relatorios.nome_arquivo(res, "pdf").replace(".pdf", "_detalhado.pdf"),
        mime="application/pdf", key="download_pdf_detalhado",
        help="Relatório com detalhamento mensal e metodologia completa.",
    )
    col_xlsx.download_button(
        "Baixar Excel (.xlsx)", excel, file_name=relatorios.nome_arquivo(res, "xlsx"),
        mime=MIME_XLSX, key="download_xlsx",
    )
    col_csv.download_button(
        "Baixar CSV", csv, file_name=relatorios.nome_arquivo(res, "csv"),
        mime="text/csv", key="download_csv",
    )


# ---------------------------------------------------------------------------
# Página 2 – Importação e Gestão de Base Salarial com SQLite
# ---------------------------------------------------------------------------
def _identificar_upload(arquivo: Any) -> str:
    """Identificador estável do arquivo enviado."""
    identificador = getattr(arquivo, "file_id", None)
    if identificador:
        return str(identificador)
    return f"{getattr(arquivo, 'name', '')}-{getattr(arquivo, 'size', '')}"


def _previa_formatada(df: pd.DataFrame) -> pd.DataFrame:
    """Pré-visualização da base com valores em R$ e datas dd/mm/aaaa."""
    if df is None or df.empty:
        return pd.DataFrame()
    previa = df.copy()
    legada = data_manager.COLUNA_GRATIFICACAO_LEGADA
    if legada in previa.columns and data_manager.COLUNA_GRATIFICACAO not in previa.columns:
        previa = previa.rename(columns={legada: data_manager.COLUNA_GRATIFICACAO})
    for coluna in previa.columns:
        if (
            "Vencimento_Basico" in coluna
            or "Valor_Gratificacao" in coluna
            or "Gratificacao" in coluna
            or coluna == "gratificacao"
            or coluna == "vencimento_basico"
        ):
            previa[coluna] = previa[coluna].map(formatar_brl)
        elif "Percentual_Motorista" in coluna or coluna == "percentual_motorista":
            previa[coluna] = previa[coluna].map(relatorios.formatar_percentual)
        elif "Data_" in coluna or "Vigencia" in coluna or "data_" in coluna:
            previa[coluna] = previa[coluna].map(
                lambda d: "" if pd.isna(d) else pd.Timestamp(d).strftime("%d/%m/%Y")
            )
    return previa


def _processar_e_salvar_upload(arquivo: Any) -> None:
    """Valida a planilha enviada e grava os registros no SQLite com encadeamento."""
    nome = str(getattr(arquivo, "name", "arquivo"))
    try:
        df = data_manager.carregar_arquivo(arquivo)
    except ValueError as exc:
        st.session_state["validacao"] = {
            "nome": nome,
            "ok": False,
            "leitura": str(exc),
            "erros": [],
            "avisos": [],
        }
        return
    except Exception:  # noqa: BLE001
        st.session_state["validacao"] = {
            "nome": nome,
            "ok": False,
            "leitura": "formato não reconhecido",
            "erros": [],
            "avisos": [],
        }
        return

    resultado = data_manager.validar_e_normalizar(df)
    if not resultado.ok or resultado.df is None:
        st.session_state["validacao"] = {
            "nome": nome,
            "ok": False,
            "leitura": None,
            "erros": resultado.erros,
            "avisos": resultado.avisos,
        }
        return

    # Gravar no SQLite
    qtd_inserida, total_banco = data_manager.salvar_base_sqlite(resultado.df)
    base_atualizada = data_manager.obter_base_completa()
    st.session_state["base_salarial"] = base_atualizada
    st.session_state["nome_arquivo_base"] = nome
    st.session_state["validacao"] = {
        "nome": nome,
        "ok": True,
        "leitura": None,
        "erros": [],
        "avisos": resultado.avisos,
        "qtd_importada": qtd_inserida,
        "total_banco": total_banco,
        "df_importado": resultado.df,
    }
    _limpar_resultado()


def _exibir_mensagens_validacao() -> None:
    """Mostra o retorno da validação da importação."""
    validacao: dict[str, Any] | None = st.session_state.get("validacao")
    if validacao is None:
        return
    if validacao.get("leitura"):
        st.error(f"Não foi possível ler o arquivo: {validacao['leitura']}")
        return
    if not validacao.get("ok"):
        erros: list[str] = validacao.get("erros", [])
        st.error(
            f"O arquivo '{validacao['nome']}' foi recusado ({len(erros)} problema(s)). "
            "O banco SQLite não foi alterado."
        )
        for mensagem in erros[:5]:
            st.markdown(f"- {mensagem}")
        if len(erros) > 5:
            with st.expander("Ver todos os erros"):
                for mensagem in erros:
                    st.markdown(f"- {mensagem}")
    else:
        st.success(
            f"Arquivo **'{validacao['nome']}'** importado com sucesso! "
            f"Foram gravados/atualizados **{validacao['qtd_importada']}** registro(s). "
            f"Total acumulado no banco SQLite: **{validacao['total_banco']}** vigência(s)."
        )
    if validacao.get("avisos"):
        st.warning("\n".join(f"- {aviso}" for aviso in validacao["avisos"]))


def pagina_importacao() -> None:
    """Página de gestão da base salarial: importação, edição em grade e cadastro manual com SQLite."""
    st.title("Base Salarial (Persistência SQLite)")
    st.markdown(
        "O sistema armazena o histórico salarial no banco local **SQLite** (`db_salarios.sqlite`). "
        "Você pode **importar planilhas**, **editar registros existentes** ou **cadastrar manualmente** novas vigências."
    )

    tab_importar, tab_edicao, tab_manual = st.tabs(
        ["Importar Planilha", "Histórico e Edição", "Cadastro Manual de Vigência"]
    )

    # -----------------------------------------------------------------------
    # Aba 1: Importar Planilha
    # -----------------------------------------------------------------------
    with tab_importar:
        st.subheader("Importação em Lote via Planilha")
        st.caption(
            "Baixe a planilha modelo, preencha as tabelas salariais e envie o arquivo. "
            "Os dados serão validados, gravados no SQLite e as vigências anteriores serão "
            "automaticamente encerradas no dia anterior ao novo reajuste."
        )

        col1, col2 = st.columns(2)
        try:
            modelo_xlsx = data_manager.gerar_template()
            modelo_csv = data_manager.gerar_template_csv()
        except Exception:  # noqa: BLE001
            st.error("Não foi possível gerar a planilha modelo.")
        else:
            col1.download_button(
                "Baixar planilha modelo (.xlsx)", modelo_xlsx,
                file_name="template_tabelas_salariais.xlsx", mime=MIME_XLSX,
                key="download_modelo_xlsx",
            )
            col2.download_button(
                "Baixar planilha modelo (.csv)", modelo_csv,
                file_name="template_tabelas_salariais.csv", mime="text/csv",
                key="download_modelo_csv",
            )

        arquivo = st.file_uploader(
            "Selecione a planilha com a tabela salarial (.xlsx ou .csv):",
            type=["xlsx", "csv"],
            key=f"uploader_{st.session_state['versao_uploader']}",
        )

        if arquivo is not None:
            nome = str(getattr(arquivo, "name", "planilha"))
            try:
                df = data_manager.carregar_arquivo(arquivo)
                resultado = data_manager.validar_e_normalizar(df)
            except ValueError as exc:
                st.error(f"Erro ao ler o arquivo '{nome}': {exc}")
                resultado = None
            except Exception:
                st.error(f"Formato não reconhecido no arquivo '{nome}'.")
                resultado = None

            if resultado is not None:
                if not resultado.ok or resultado.df is None:
                    st.error(
                        f"O arquivo '{nome}' contém {len(resultado.erros)} problema(s) e não pode ser importado:"
                    )
                    for erro in resultado.erros[:5]:
                        st.markdown(f"- {erro}")
                    if len(resultado.erros) > 5:
                        with st.expander("Ver todos os erros"):
                            for erro in resultado.erros:
                                st.markdown(f"- {erro}")
                else:
                    df_validado = resultado.df
                    base_atual = data_manager.obter_base_completa()
                    conflitos = data_manager.identificar_conflitos(base_atual, df_validado)

                    st.info(
                        f"Planilha pronta para importação: **{nome}** ({len(df_validado)} registro(s) válido(s))."
                    )

                    if not conflitos.empty:
                        st.warning(
                            f"⚠️ **Atenção:** Foram encontrados **{len(conflitos)} registro(s)** com a mesma chave "
                            "primária (Cargo, Vínculo, CH e Data de Início) já cadastrados no SQLite. "
                            "Ao clicar em importar, esses registros serão atualizados com os novos valores."
                        )
                        rotulo_botao = "Importar e Atualizar Base no SQLite"
                    else:
                        rotulo_botao = "Importar Base Salarial para o SQLite"

                    with st.expander("Ver prévia dos dados da planilha selecionada", expanded=True):
                        st.dataframe(_previa_formatada(df_validado), hide_index=True)

                    if st.button(rotulo_botao, type="primary", key="btn_executar_importacao"):
                        qtd_inserida, total_banco = data_manager.salvar_base_sqlite(df_validado)
                        st.session_state["base_salarial"] = data_manager.obter_base_completa()
                        st.session_state["nome_arquivo_base"] = nome
                        st.session_state["validacao"] = {
                            "nome": nome,
                            "ok": True,
                            "qtd_importada": qtd_inserida,
                            "total_banco": total_banco,
                            "avisos": resultado.avisos,
                        }
                        st.session_state["versao_uploader"] += 1
                        _limpar_resultado()
                        st.success(
                            f"✅ **Importação concluída com sucesso!** "
                            f"Foram gravados/atualizados **{qtd_inserida}** registro(s). "
                            f"Total no banco SQLite: **{total_banco}** vigência(s)."
                        )
                        st.rerun()

        _exibir_mensagens_validacao()

        # Visualização da Base Atual no SQLite
        base_atual = data_manager.obter_base_completa()
        if not base_atual.empty:
            st.divider()
            st.subheader("Base Salarial Consolidada Atual (SQLite)")
            st.caption(f"Total de {len(base_atual)} vigência(s) cadastradas no banco local.")
            st.dataframe(_previa_formatada(base_atual), hide_index=True)

            col_reset, _ = st.columns([1, 3])
            with col_reset:
                if st.button("Limpar Banco SQLite", key="btn_limpar_banco"):
                    data_manager.limpar_banco_sqlite()
                    st.session_state["base_salarial"] = None
                    st.session_state["nome_arquivo_base"] = None
                    st.session_state["validacao"] = None
                    st.session_state["ultimo_upload"] = None
                    st.session_state["versao_uploader"] += 1
                    _limpar_resultado()
                    st.info("Banco SQLite limpo com sucesso.")
                    st.rerun()

    # -----------------------------------------------------------------------
    # Aba 2: Histórico e Edição
    # -----------------------------------------------------------------------
    with tab_edicao:
        st.subheader("Histórico e Edição Direta de Registros")
        st.caption(
            "Edite qualquer campo diretamente na tabela abaixo (valores, datas, cargas horárias) "
            "e clique em **Salvar Alterações**. O sistema validará os dados e reencadeará as vigências."
        )

        base_edicao = data_manager.obter_base_completa()
        if base_edicao.empty:
            st.info(
                "Nenhum registro cadastrado no banco SQLite. "
                "Utilize a **Aba 1 (Importar Planilha)** ou a **Aba 3 (Cadastro Manual)**."
            )
        else:
            cargos_opcoes = ["Todos os Cargos"] + data_manager.listar_cargos_cadastrados()
            filtro_cargo = st.selectbox(
                "Filtrar tabela para edição:",
                cargos_opcoes,
                key="sb_filtro_edicao",
            )

            if filtro_cargo == "Todos os Cargos":
                df_para_editar = base_edicao.copy()
            else:
                df_para_editar = base_edicao[base_edicao["Cargo"] == filtro_cargo].copy()

            # Prepara colunas de datas para formato compatível com date_input/date editor
            df_para_editar["Data_Inicio_Vigencia"] = pd.to_datetime(
                df_para_editar["Data_Inicio_Vigencia"]
            ).dt.date
            df_para_editar["Data_Fim_Vigencia"] = pd.to_datetime(
                df_para_editar["Data_Fim_Vigencia"]
            ).dt.date

            df_editado = st.data_editor(
                df_para_editar,
                use_container_width=True,
                num_rows="dynamic",
                key="editor_grade_salarial",
                column_config={
                    "Cargo": st.column_config.SelectboxColumn(
                        "Cargo",
                        options=engine.nomes_cargos(),
                        required=True,
                        width="medium",
                    ),
                    "Vinculo": st.column_config.SelectboxColumn(
                        "Vínculo",
                        options=[engine.VINCULO_EFETIVO, engine.VINCULO_TEMPORARIO],
                        required=True,
                        width="small",
                    ),
                    "CH": st.column_config.TextColumn(
                        "Carga Horária",
                        required=True,
                        width="small",
                    ),
                    "Vencimento_Basico": st.column_config.NumberColumn(
                        "Vencimento Básico (R$)",
                        format="R$ %.2f",
                        min_value=0.01,
                        required=True,
                        width="medium",
                    ),
                    "Valor_Gratificacao": st.column_config.NumberColumn(
                        "Gratificação (R$)",
                        format="R$ %.2f",
                        min_value=0.0,
                        width="medium",
                    ),
                    "Percentual_Motorista": st.column_config.NumberColumn(
                        "Percentual Motorista (%)",
                        format="%.2f %%",
                        min_value=0.0,
                        max_value=100.0,
                        width="small",
                    ),
                    "Data_Inicio_Vigencia": st.column_config.DateColumn(
                        "Início Vigência",
                        format="DD/MM/YYYY",
                        required=True,
                        width="medium",
                    ),
                    "Data_Fim_Vigencia": st.column_config.DateColumn(
                        "Fim Vigência (opcional)",
                        format="DD/MM/YYYY",
                        width="medium",
                    ),
                },
            )

            col_salv, col_del = st.columns([2, 1])
            with col_salv:
                if st.button("Salvar Alterações no SQLite", type="primary", key="btn_salvar_grade"):
                    if filtro_cargo != "Todos os Cargos":
                        outros = base_edicao[base_edicao["Cargo"] != filtro_cargo].copy()
                        df_completo = pd.concat([outros, df_editado], ignore_index=True)
                    else:
                        df_completo = df_editado

                    resultado_up = data_manager.atualizar_base_completa_sqlite(df_completo)
                    if not resultado_up.ok:
                        st.error("Não foi possível salvar as alterações devido aos seguintes erros:")
                        for err in resultado_up.erros:
                            st.markdown(f"- {err}")
                    else:
                        st.session_state["base_salarial"] = data_manager.obter_base_completa()
                        _limpar_resultado()
                        st.success("✅ Alterações salvas com sucesso no banco SQLite!")
                        st.rerun()

    # -----------------------------------------------------------------------
    # Aba 3: Cadastro Manual de Vigência
    # -----------------------------------------------------------------------
    with tab_manual:
        st.subheader("Cadastro Manual de Registro / Vigência")
        st.caption(
            "Cadastre individualmente um novo salário ou reajuste anual para qualquer cargo do catálogo."
        )

        with st.form("form_cadastro_manual"):
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                cat_manual = st.selectbox(
                    "Categoria",
                    engine.listar_categorias(),
                    key="man_categoria",
                )
                vinculo_manual = st.selectbox(
                    "Vínculo",
                    engine.listar_vinculos(cat_manual),
                    key="man_vinculo",
                )
            with col_m2:
                cargo_manual = st.selectbox(
                    "Cargo",
                    engine.listar_cargos(cat_manual, vinculo_manual),
                    key="man_cargo",
                )
                ch_opcoes = engine.listar_chs(cat_manual, vinculo_manual, cargo_manual)
                ch_manual = st.selectbox(
                    "Carga Horária",
                    ch_opcoes,
                    key="man_ch",
                )

            cargo_def = engine.obter_cargo(cat_manual, vinculo_manual, cargo_manual)
            if cargo_def is not None:
                st.caption(
                    f"Catálogo oficial: Nível/Classe **{cargo_def.nivel_classe}** · "
                    f"Incentivo/Adicional: **{cargo_def.incentivo_descricao}**"
                )

            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                venc_manual = st.number_input(
                    "Vencimento Básico (R$)",
                    min_value=0.01,
                    value=4000.0,
                    step=100.0,
                    format="%.2f",
                    key="man_vencimento",
                )

            eh_motorista = cargo_def.is_motorista if cargo_def else False
            eh_docente = cargo_def.is_docente if cargo_def else False

            with col_v2:
                if eh_docente or eh_motorista:
                    grat_manual = 0.0
                    motivo = "Docente (recebe incentivo)" if eh_docente else "Motorista (não recebe gratificação)"
                    st.text_input(
                        "Gratificação (R$)",
                        value="R$ 0,00",
                        disabled=True,
                        help=f"Desabilitado: {motivo}",
                    )
                else:
                    grat_manual = float(
                        st.number_input(
                            "Gratificação (R$)",
                            min_value=0.0,
                            value=0.0,
                            step=50.0,
                            format="%.2f",
                            key="man_gratificacao",
                        )
                    )

            with col_v3:
                if eh_motorista:
                    perc_mot_manual = float(
                        st.number_input(
                            "Percentual Motorista (%)",
                            min_value=0.0,
                            max_value=100.0,
                            value=20.0,
                            step=1.0,
                            format="%.2f",
                            key="man_perc_mot",
                        )
                    )
                else:
                    perc_mot_manual = 0.0
                    st.text_input(
                        "Percentual Motorista (%)",
                        value="0,00%",
                        disabled=True,
                        help="Aplica-se apenas ao cargo de Motorista",
                    )

            col_d1, col_d2 = st.columns(2)
            with col_d1:
                hoje = date.today()
                data_ini_default = date(hoje.year, hoje.month, 1)
                data_inicio_manual = st.date_input(
                    "Data de Início da Vigência",
                    value=data_ini_default,
                    key="man_data_inicio",
                    help="Recomenda-se dia 1 do mês de início do reajuste.",
                )

            with col_d2:
                vig_aberta = st.checkbox(
                    "Vigência Aberta (atual, sem data de término)",
                    value=True,
                    key="man_vig_aberta",
                )
                if not vig_aberta:
                    data_fim_manual = st.date_input(
                        "Data de Fim da Vigência",
                        value=date(hoje.year + 1, 12, 31),
                        key="man_data_fim",
                    )
                else:
                    data_fim_manual = None

            btn_submeter_manual = st.form_submit_button(
                "Cadastrar Registro no SQLite",
                type="primary",
            )

        if btn_submeter_manual:
            res_manual = data_manager.salvar_registro_individual(
                cargo=cargo_manual,
                vinculo=vinculo_manual,
                ch=ch_manual,
                vencimento_basico=float(venc_manual),
                gratificacao=float(grat_manual),
                percentual_motorista=float(perc_mot_manual),
                data_inicio_vigencia=data_inicio_manual.strftime("%Y-%m-%d"),
                data_fim_vigencia=data_fim_manual.strftime("%Y-%m-%d") if data_fim_manual else None,
            )
            if not res_manual.ok:
                st.error("Não foi possível cadastrar o registro:")
                for erro_man in res_manual.erros:
                    st.markdown(f"- {erro_man}")
            else:
                st.session_state["base_salarial"] = data_manager.obter_base_completa()
                _limpar_resultado()
                st.success(
                    f"✅ Registro cadastrado com sucesso para **{cargo_manual} ({vinculo_manual}, {ch_manual})** "
                    f"a partir de {data_inicio_manual.strftime('%d/%m/%Y')}!"
                )
                st.rerun()


# ---------------------------------------------------------------------------
# Aplicação
# ---------------------------------------------------------------------------
def main() -> None:
    """Ponto de entrada do aplicativo Streamlit."""
    st.set_page_config(
        page_title="Impacto Orçamentário – UEFS",
        layout="wide",
        page_icon="📊",
    )
    _inicializar_estado()

    # CSS customizado para a barra lateral
    st.sidebar.markdown(
        """
        <style>
            [data-testid="stSidebar"] {
                background-color: #F8FAFC;
                border-right: 1px solid #E2E8F0;
                padding-top: 1.5rem;
            }
            [data-testid="stSidebar"] [data-testid="stRadio"] > label {
                font-weight: 700;
                color: #1F3864;
                font-size: 0.95rem;
                letter-spacing: -0.2px;
                margin-bottom: 0.5rem;
            }
            [data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] > label {
                padding: 0.4rem 0.6rem;
                border-radius: 6px;
                transition: background-color 0.15s ease-in-out;
            }
            [data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] > label:hover {
                background-color: #EEF2F7;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Identidade Institucional no topo da barra lateral
    caminho_logo = Path(__file__).resolve().parent / "logos" / "logo_uefs.png"
    if caminho_logo.is_file():
        st.sidebar.image(str(caminho_logo), use_container_width=True)

    st.sidebar.markdown(
        """
        <div style="text-align: center; margin-bottom: 22px; padding: 0 4px;">
            <h2 style="color: #1F3864; margin-bottom: 3px; font-weight: 800; font-size: 1.35rem; letter-spacing: -0.5px;">
                ImpactoUEFS
            </h2>
            <div style="color: #334155; font-size: 0.82rem; font-weight: 600; line-height: 1.3;">
                Universidade Estadual de Feira de Santana
            </div>
            <div style="color: #64748B; font-size: 0.75rem; margin-top: 2px; line-height: 1.2;">
                Pró-Reitoria de Gestão e Desenvolvimento de Pessoas (PGDP)
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Menu de Navegação Moderno com Ícones
    opcoes_menu = {
        "📊 Simulação de Impacto": PAGINA_SIMULACAO,
        "🗄️ Base Salarial (SQLite)": PAGINA_IMPORTACAO,
    }

    escolha_atual = st.session_state.get("pagina_menu", "📊 Simulação de Impacto")
    if escolha_atual not in opcoes_menu:
        escolha_atual = "📊 Simulação de Impacto"

    item_selecionado = st.sidebar.radio(
        "Navegação",
        list(opcoes_menu.keys()),
        index=list(opcoes_menu.keys()).index(escolha_atual),
        key="pagina_radio",
    )
    st.session_state["pagina_menu"] = item_selecionado
    pagina = opcoes_menu[item_selecionado]

    # Rodapé Institucional Discreto
    st.sidebar.markdown(
        """
        <div style="margin-top: 60px; padding: 14px 10px; background: #FFFFFF; border-radius: 8px; border: 1px solid #E2E8F0; font-size: 0.75rem; color: #64748B; text-align: center; line-height: 1.4;">
            <strong style="color: #1F3864;">ImpactoUEFS</strong><br/>
            Demonstrativo de Impacto Orçamentário<br/>
            <span style="color: #94A3B8;">PGDP · UEFS</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if pagina == PAGINA_SIMULACAO:
        pagina_simulacao()
    else:
        pagina_importacao()


if __name__ == "__main__":
    main()
