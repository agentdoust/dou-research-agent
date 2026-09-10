# -*- coding: utf-8 -*-
"""Interface Streamlit para o Agente de Pesquisa no Diário Oficial da União (DOU).

Reutiliza as funções de D:\\opencode_desktop\\Agents\\DOU_Research_Agent\\agent.py.
Fluxo:
  1. Escolher a data e clicar em "Carregar índices da edição" (baixa PDFs em cache
     e lê o sumário das Seções 1 e 2).
  2. Selecionar os índices (entradas do sumário) a pesquisar.
  3. Executar a pesquisa — a busca fica limitada aos índices selecionados.

Executar:  streamlit run app.py
"""

import io
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

AGENT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(AGENT_DIR))

import agent as ag  # noqa: E402

st.set_page_config(
    page_title="DOU Agente de Pesquisa",
    page_icon="📰",
    layout="wide",
)

MIN_DATA = date(2010, 1, 1)


def capturar_saida(fn):
    """Executa fn capturando o que ela imprimir em stdout e retorna (resultado, log)."""
    buf = io.StringIO()
    old = sys.stdout
    sys.stdout = buf
    try:
        resultado = fn()
    finally:
        sys.stdout = old
    return resultado, buf.getvalue()


def descarregar_indices(sessao, secao: int, data: date):
    """Baixa o PDF da seção (cache) e lê os índices do sumário."""
    cfg = ag.SECOES[secao]
    link = ag.obter_link_edicao(sessao, secao, data)
    if not link:
        return []
    arquivo = link.split("/")[-1].split("?")[0]
    destino = ag.CACHE_DIR / data.strftime("%Y_%m_%d") / arquivo
    ag.baixar_pdf(link, destino)
    return ag.ler_indices(ag.ler_paginas(destino))


def main():
    st.title("📰 DOU Agente de Pesquisa")
    st.caption(
        "Pesquisador isento e técnico — pesquisa o Diário Oficial da União na fonte oficial "
        "(pesquisa.in.gov.br), filtrando o Ministério da Justiça e Segurança Pública por padrão."
    )

    with st.sidebar:
        st.header("Parâmetros")
        hoje = date.today()
        data = st.date_input(
            "Data da edição (qualquer data)",
            value=hoje,
            min_value=MIN_DATA,
        )
        forcar = st.toggle("Forçar novo download (ignorar cache)", value=False)
        carregar = st.button(
            "1️⃣ Carregar índices da edição",
            type="primary",
            use_container_width=True,
        )

        if carregar:
            if forcar:
                cache = ag.CACHE_DIR / data.strftime("%Y_%m_%d")
                if cache.exists():
                    for f in cache.iterdir():
                        f.unlink()
            status = st.status("Baixando edição e lendo o sumário...", expanded=True)
            try:
                sessao = ag.criar_sessao()
                inds = {
                    1: descarregar_indices(sessao, 1, data),
                    2: descarregar_indices(sessao, 2, data),
                }
            except Exception as e:  # noqa: BLE001
                status.update(label="Falha ao carregar índices", state="error")
                with status:
                    st.text(f"{type(e).__name__}: {e}")
            else:
                status.update(label="Índices carregados", state="complete")
                with status:
                    st.text(f"Seção 1: {len(inds[1])} índices")
                    st.text(f"Seção 2: {len(inds[2])} índices")
                    if not inds[1] and not inds[2]:
                        st.text(
                            "Edição não encontrada para esta data "
                            "(pode não ter sido publicada ou ser fim de semana/feriado)."
                        )
                st.session_state["indices"] = inds
                st.session_state["indices_data"] = data

        inds = st.session_state.get("indices")
        dados_carregados = (
            inds is not None
            and st.session_state.get("indices_data") == data
            and (inds.get(1) or inds.get(2))
        )

        opcoes = []
        if dados_carregados:
            opcoes = [
                "Seção 1 · (edição inteira)",
                *[f"Seção 1 · {t['nome']}" for t in inds.get(1, [])],
                "Seção 2 · (edição inteira)",
                *[f"Seção 2 · {t['nome']}" for t in inds.get(2, [])],
            ]
            mjsp_1 = next((f"Seção 1 · {t['nome']}" for t in inds.get(1, []) if t["nome"] == ag.MINISTERIO_ALVO), "")
            mjsp_2 = next((f"Seção 2 · {t['nome']}" for t in inds.get(2, []) if t["nome"] == ag.MINISTERIO_ALVO), "")
            default = [o for o in (mjsp_1, mjsp_2) if o]

        sel = st.multiselect(
            "Índices da edição a pesquisar",
            opcoes,
            default=default if dados_carregados else (),
            disabled=not dados_carregados,
            placeholder="Carregue os índices primeiro",
        )
        st.caption(
            "A pesquisa fica limitada aos índices selecionados. "
            "Selecione '(edição inteira)' para varrer a seção completa."
        )
        st.divider()
        executar = st.button(
            "2️⃣ Executar pesquisa",
            type="primary",
            use_container_width=True,
            disabled=not dados_carregados or not sel,
        )

    if not dados_carregados:
        st.info("Escolha a data ao lado e clique em **1️⃣ Carregar índices da edição**.")
        return

    if executar:
        # agrupa seleção por seção: 1 -> [nomes] ou [None]; 2 -> [nomes] ou [None]
        selecao = {1: [], 2: []}
        for o in sel:
            sec, nome = o.split(" · ", 1)
            s = 1 if "Seção 1" in sec else 2
            if nome.startswith("("):
                selecao[s].append(None)
            else:
                selecao[s].append(nome)
        secoes_a_processar = {s for s, v in selecao.items() if v}

        if forcar:
            cache = ag.CACHE_DIR / data.strftime("%Y_%m_%d")
            if cache.exists():
                for f in cache.iterdir():
                    f.unlink()

        status = st.status("Pesquisando no DOU...", expanded=True)

        def rodar():
            sessao = ag.criar_sessao()
            resultados = {}
            for s in sorted(secoes_a_processar):
                inteira = any(n is None for n in selecao[s])
                nomes = [n for n in selecao[s] if n is not None]
                res = ag.processar_secao(sessao, s, data, indices=None if inteira else nomes)
                resultados.update(res)
            caminho = ag.gerar_docx(data, resultados)
            return resultados, caminho

        try:
            (resultados, caminho), log = capturar_saida(rodar)
        except Exception as e:  # noqa: BLE001
            status.update(label="Falha na execução", state="error")
            st.error(f"{type(e).__name__}: {e}")
            return

        status.update(label="Concluído", state="complete")
        with status:
            for linha in log.splitlines():
                st.text(linha)

        st.session_state["resultados"] = resultados
        st.session_state["caminho"] = str(caminho)
        st.session_state["data"] = data

    resultados = st.session_state.get("resultados")
    if resultados is None:
        st.info("Selecione os índices desejados e clique em **2️⃣ Executar pesquisa**.")
        return

    data = st.session_state.get("data")
    caminho = st.session_state.get("caminho")

    col_tit, col_dl = st.columns([4, 1])
    col_tit.subheader(f"Resultados — edição {data.strftime('%d/%m/%Y')}")
    with col_dl:
        if caminho and Path(caminho).exists():
            with open(caminho, "rb") as fh:
                st.download_button(
                    "⬇️ Baixar relatório Word",
                    data=fh.read(),
                    file_name=Path(caminho).name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

    s1 = resultados.get("secao1", [])
    s2 = resultados.get("secao2", [])

    tab1, tab2 = st.tabs(
        [
            f"Seção 1 — Atos normativos ({len(s1)})",
            f"Seção 2 — PF: nomeações, exonerações, designações e dispensas ({len(s2)})",
        ]
    )

    with tab1:
        if not s1:
            st.info("Nenhum ato normativo de segurança pública identificado.")
        else:
            for k, a in enumerate(s1, 1):
                with st.expander(f"{k}. {a['titulo']}"):
                    corpo = ag.resumir_corpo(a["corpo"])
                    st.write(corpo)

    with tab2:
        if not s2:
            st.info(
                "Nenhuma ocorrência de pessoal da Polícia Federal identificada "
                "(nomeações, exonerações, designações e dispensas)."
            )
        else:
            df = pd.DataFrame([
                {
                    "Nº": i.get("num", ""),
                    "Ato": i.get("ato", ""),
                    "Ação": i.get("verbo", ""),
                    "Nome": i.get("nome", ""),
                    "Detalhes": i.get("resto", ""),
                }
                for i in s2
            ])
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Nº": st.column_config.TextColumn(width="small"),
                    "Ato": st.column_config.TextColumn(width="medium"),
                    "Ação": st.column_config.TextColumn(width="small"),
                    "Nome": st.column_config.TextColumn(width="medium"),
                    "Detalhes": st.column_config.TextColumn(width="large"),
                },
            )

    st.caption(
        f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} — "
        f"fonte: {ag.PESQUISA_URL}"
    )


if __name__ == "__main__":
    main()