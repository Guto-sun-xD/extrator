"""Extrator PDF Web — Interface Streamlit sobre extrator.py"""

import os
import tempfile
import time
from datetime import date, timedelta
from pathlib import Path
from typing import List

import google.generativeai as genai
import pandas as pd
import plotly.express as px
import streamlit as st

from extrator import (
    MODELO_GEMINI,
    TIPOS_AMIGAVEIS,
    DocumentoExtraido,
    ResultadoProcessamento,
    _fmt_brl,
    _fmt_data,
    extrair_dados_pdf,
    gerar_planilha,
)

# ── Configuração da página ────────────────────────────────────────────────────

st.set_page_config(
    page_title="Extrator de PDFs Financeiros",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS global ────────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .block-container { padding-top: 2rem; }
    .stMetric { background: #F0F4F8; border-radius: 8px; padding: 0.5rem 1rem; }
    .hero-title { color: #1F3864; font-size: 2.2rem; font-weight: 700; margin-bottom: 0; }
    .hero-sub { color: #5A6B80; font-size: 1.05rem; margin-top: 0.2rem; }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────

_DEFAULTS = {
    "pagina": "upload",
    "arquivos": [],
    "resultados": [],
    "documentos": [],
    "excel_bytes": None,
    "excel_nome": None,
}

for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ── Helpers ───────────────────────────────────────────────────────────────────


def _get_gemini_model() -> genai.GenerativeModel:
    """Lê a API key do st.secrets (Streamlit Cloud) ou da variável de ambiente (local)."""
    api_key = ""
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except (KeyError, FileNotFoundError):
        api_key = os.getenv("GEMINI_API_KEY", "")

    if not api_key:
        st.error(
            "**GEMINI_API_KEY não configurada.**\n\n"
            "- **Local:** crie `.streamlit/secrets.toml` com `GEMINI_API_KEY = \"sua_chave\"`\n"
            "- **Streamlit Cloud:** configure em *App settings → Secrets*"
        )
        st.stop()

    genai.configure(api_key=api_key)
    return genai.GenerativeModel(
        model_name=MODELO_GEMINI,
        generation_config=genai.GenerationConfig(
            response_mime_type="application/json",
            temperature=0.1,
        ),
    )


def _gerar_excel_bytes(docs: List[DocumentoExtraido]):
    """Gera o Excel via extrator.py e devolve os bytes para download."""
    caminho = gerar_planilha(docs)
    return caminho.read_bytes(), caminho.name


def _doc_para_dict(doc: DocumentoExtraido) -> dict:
    return {
        "Tipo": TIPOS_AMIGAVEIS.get(doc.tipo, doc.tipo),
        "Fornecedor": doc.fornecedor,
        "Valor (R$)": doc.valor_total,
        "Emissão": _fmt_data(doc.data_emissao),
        "Vencimento": _fmt_data(doc.data_vencimento),
        "Categoria": doc.categoria_sugerida.capitalize(),
        "Arquivo": doc.arquivo_original or "",
        "Cód. Barras": doc.codigo_barras or "",
        "Nº Documento": doc.numero_documento or "",
        "Tipo Serviço": doc.tipo_servico or "",
        "Mês Ref.": doc.mes_referencia or "",
        "Nº NF": doc.numero_nf or "",
        "CNPJ Emissor": doc.cnpj_emissor or "",
    }


def _reset_estado():
    st.session_state.pagina = "upload"
    st.session_state.arquivos = []
    st.session_state.resultados = []
    st.session_state.documentos = []
    st.session_state.excel_bytes = None
    st.session_state.excel_nome = None


# ── Página 1: Upload ──────────────────────────────────────────────────────────


def page_upload():
    _, col_hero, _ = st.columns([1, 4, 1])
    with col_hero:
        st.markdown(
            '<p class="hero-title">📄 Extrator de PDFs Financeiros</p>'
            '<p class="hero-sub">Digitalização inteligente com Gemini 2.5 Flash · '
            'Boletos · Contas · Notas Fiscais · Comprovantes</p>',
            unsafe_allow_html=True,
        )

    st.divider()

    col_upload, col_info = st.columns([3, 2], gap="large")

    with col_upload:
        st.subheader("Selecione os PDFs")
        uploaded = st.file_uploader(
            "Upload de PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            label_visibility="collapsed",
        )

        if uploaded:
            st.success(f"**{len(uploaded)} arquivo(s) pronto(s)**")
            with st.expander("Arquivos selecionados", expanded=True):
                for f in uploaded:
                    kb = len(f.getvalue()) / 1024
                    tamanho = f"{kb:.1f} KB" if kb < 1024 else f"{kb/1024:.1f} MB"
                    st.write(f"📄 **{f.name}** — {tamanho}")

            st.markdown("<br>", unsafe_allow_html=True)

            if st.button("🚀 Processar PDFs", type="primary", use_container_width=True):
                # Lê os bytes agora — UploadedFile não sobrevive ao rerun
                st.session_state.arquivos = [
                    {"nome": f.name, "bytes": f.getvalue()} for f in uploaded
                ]
                st.session_state.pagina = "processando"
                st.rerun()
        else:
            st.info("Arraste PDFs aqui ou clique para selecionar. Múltiplos arquivos são aceitos.")

    with col_info:
        st.subheader("Tipos reconhecidos")
        for icon, tipo, desc in [
            ("🏦", "Boleto Bancário", "Boletos de cobrança bancária"),
            ("💡", "Conta de Consumo", "Luz, água, gás, internet, telefone"),
            ("🧾", "Nota Fiscal", "NF-e, NFS-e, NFCe"),
            ("✅", "Comprovante de Pagamento", "PIX, TED, transferência"),
            ("📋", "Recibo", "Recibos de quitação"),
            ("📁", "Outros", "Qualquer outro documento financeiro"),
        ]:
            st.markdown(f"{icon} **{tipo}** — {desc}")

        st.divider()
        st.caption("Os PDFs são enviados ao **Google Files API** e deletados imediatamente após a extração.")

    st.markdown("---")
    st.caption("Feito com [Claude Code](https://claude.ai/code)")


# ── Página 2: Processando ─────────────────────────────────────────────────────


def page_processando():
    st.markdown(
        '<div style="text-align:center;padding:0.5rem 0 1.5rem">'
        '<h2 style="color:#1F3864">⚙️ Processando PDFs...</h2>'
        "</div>",
        unsafe_allow_html=True,
    )

    arquivos = st.session_state.arquivos
    total = len(arquivos)

    barra = st.progress(0, text="Iniciando...")
    card_atual = st.empty()

    st.subheader("Log de processamento")
    placeholder_log = st.empty()

    modelo = _get_gemini_model()

    resultados: List[ResultadoProcessamento] = []
    documentos: List[DocumentoExtraido] = []
    log: List[str] = []

    for i, arq in enumerate(arquivos):
        nome = arq["nome"]
        barra.progress(i / total, text=f"Arquivo {i + 1} de {total}: **{nome}**")
        card_atual.info(f"**Processando agora:** 📄 {nome}")

        # Salva em arquivo temporário para o extrator.py
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(arq["bytes"])
            tmp_path = Path(tmp.name)

        try:
            doc = extrair_dados_pdf(tmp_path, modelo)
            doc.arquivo_original = nome
            documentos.append(doc)
            resultados.append(
                ResultadoProcessamento(pdf_path=nome, sucesso=True, documento=doc)
            )
            log.append(
                f"✅ **{nome}** — "
                f"{TIPOS_AMIGAVEIS.get(doc.tipo, doc.tipo)} · {_fmt_brl(doc.valor_total)}"
            )
        except Exception as exc:
            msg = str(exc)
            resultados.append(
                ResultadoProcessamento(pdf_path=nome, sucesso=False, erro=msg)
            )
            log.append(f"❌ **{nome}** — {msg[:120]}")
        finally:
            try:
                tmp_path.unlink()
            except Exception:
                pass

        placeholder_log.markdown("\n\n".join(log))

    barra.progress(1.0, text="Concluído!")
    card_atual.success("**Todos os arquivos foram processados!**")

    if documentos:
        with st.spinner("Gerando planilha Excel..."):
            try:
                excel_bytes, excel_nome = _gerar_excel_bytes(documentos)
                st.session_state.excel_bytes = excel_bytes
                st.session_state.excel_nome = excel_nome
            except Exception as exc:
                st.warning(f"Não foi possível gerar a planilha: {exc}")

    st.session_state.resultados = resultados
    st.session_state.documentos = documentos
    st.session_state.pagina = "resultados"

    time.sleep(1.2)
    st.rerun()


# ── Página 3: Resultados ──────────────────────────────────────────────────────


def page_resultados():
    documentos: List[DocumentoExtraido] = st.session_state.documentos
    resultados: List[ResultadoProcessamento] = st.session_state.resultados

    col_titulo, col_botao = st.columns([5, 1])
    with col_titulo:
        st.markdown("## 📊 Resultados da Extração")
    with col_botao:
        if st.button("🔄 Processar mais", use_container_width=True):
            _reset_estado()
            st.rerun()

    st.divider()

    # ── Métricas ──
    sucessos = [r for r in resultados if r.sucesso]
    falhas = [r for r in resultados if not r.sucesso]
    total_valor = sum(d.valor_total for d in documentos)
    hoje = date.today()
    proximos_30 = [
        d for d in documentos
        if d.data_vencimento and hoje <= d.data_vencimento <= hoje + timedelta(days=30)
    ]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total processados", len(resultados))
    c2.metric("Extraídos com sucesso", len(sucessos), delta=f"-{len(falhas)} erro(s)" if falhas else None, delta_color="inverse")
    c3.metric("Valor total extraído", _fmt_brl(total_valor))
    c4.metric("Vencem em 30 dias", len(proximos_30))

    # Estilização de métricas via streamlit-extras (opcional)
    try:
        from streamlit_extras.metric_cards import style_metric_cards
        style_metric_cards(border_left_color="#1F3864", box_shadow=True)
    except Exception:
        pass

    if falhas:
        with st.expander(f"⚠️ {len(falhas)} arquivo(s) com erro — clique para ver detalhes"):
            for r in falhas:
                st.error(f"**{Path(r.pdf_path).name}**: {r.erro}")

    if not documentos:
        st.warning("Nenhum documento foi extraído com sucesso.")
        return

    # ── Botão de download ──
    if st.session_state.excel_bytes:
        st.download_button(
            label="⬇️  Baixar planilha Excel",
            data=st.session_state.excel_bytes,
            file_name=st.session_state.excel_nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    st.markdown("---")

    # ── Tabela com filtros ──
    st.subheader("Documentos extraídos")

    df = pd.DataFrame([_doc_para_dict(d) for d in documentos])

    cf1, cf2, cf3 = st.columns([2, 2, 3])
    with cf1:
        tipos_opts = ["Todos"] + sorted(df["Tipo"].unique().tolist())
        filtro_tipo = st.selectbox("Tipo de documento", tipos_opts)
    with cf2:
        cats_opts = ["Todas"] + sorted(df["Categoria"].unique().tolist())
        filtro_cat = st.selectbox("Categoria", cats_opts)
    with cf3:
        busca = st.text_input("Buscar por fornecedor", placeholder="Digite para filtrar...")

    df_vis = df.copy()
    if filtro_tipo != "Todos":
        df_vis = df_vis[df_vis["Tipo"] == filtro_tipo]
    if filtro_cat != "Todas":
        df_vis = df_vis[df_vis["Categoria"] == filtro_cat]
    if busca:
        df_vis = df_vis[df_vis["Fornecedor"].str.contains(busca, case=False, na=False)]

    colunas_principais = ["Tipo", "Fornecedor", "Valor (R$)", "Emissão", "Vencimento", "Categoria", "Arquivo"]
    st.dataframe(
        df_vis[colunas_principais],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Valor (R$)": st.column_config.NumberColumn("Valor (R$)", format="R$ %.2f"),
        },
    )
    st.caption(f"{len(df_vis)} de {len(df)} documento(s) exibido(s)")

    with st.expander("Ver todos os campos extraídos"):
        st.dataframe(df_vis, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── Gráficos ──
    st.subheader("Gráficos")
    col_pizza, col_linha = st.columns(2, gap="large")

    with col_pizza:
        cats_valor = df.groupby("Categoria")["Valor (R$)"].sum().reset_index()
        fig_pizza = px.pie(
            cats_valor,
            names="Categoria",
            values="Valor (R$)",
            title="Gastos por Categoria",
            color_discrete_sequence=px.colors.qualitative.Set3,
            hole=0.3,
        )
        fig_pizza.update_traces(textposition="inside", textinfo="percent+label")
        fig_pizza.update_layout(showlegend=False, height=380, margin=dict(t=50, b=10))
        st.plotly_chart(fig_pizza, use_container_width=True)

    with col_linha:
        df_com_data = df[df["Emissão"] != ""].copy()
        if not df_com_data.empty:
            df_com_data["_data"] = pd.to_datetime(
                df_com_data["Emissão"], format="%d/%m/%Y", errors="coerce"
            )
            df_com_data = df_com_data.dropna(subset=["_data"]).sort_values("_data")
            df_por_data = df_com_data.groupby("_data")["Valor (R$)"].sum().reset_index()
            df_por_data.columns = ["Data", "Valor (R$)"]
            fig_linha = px.line(
                df_por_data,
                x="Data",
                y="Valor (R$)",
                title="Gastos por Data de Emissão",
                markers=True,
                color_discrete_sequence=["#1F3864"],
            )
            fig_linha.update_layout(height=380, margin=dict(t=50, b=10))
            st.plotly_chart(fig_linha, use_container_width=True)
        else:
            st.info("Sem datas de emissão disponíveis para o gráfico de linha.")

    # ── Próximos vencimentos ──
    if proximos_30:
        st.markdown("---")
        st.subheader("⏰ Próximos vencimentos — 30 dias")
        proximos_ord = sorted(proximos_30, key=lambda d: d.data_vencimento)  # type: ignore[arg-type]
        df_prox = pd.DataFrame([_doc_para_dict(d) for d in proximos_ord])
        st.dataframe(
            df_prox[["Vencimento", "Fornecedor", "Tipo", "Valor (R$)"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Valor (R$)": st.column_config.NumberColumn("Valor (R$)", format="R$ %.2f"),
            },
        )

    st.markdown("---")
    st.caption("Feito com [Claude Code](https://claude.ai/code)")


# ── Roteador ──────────────────────────────────────────────────────────────────

_PAGINAS = {
    "upload": page_upload,
    "processando": page_processando,
    "resultados": page_resultados,
}

_PAGINAS.get(st.session_state.pagina, page_upload)()
