import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import date
from utils import (
    ler_arquivo,
    validar_planilha_leads,
    validar_planilha_vendas,
    cruzar_leads_vendas,
    gerar_relatorio_excel,
    gerar_relatorio_pdf,
    render_tabela_html,
)

st.set_page_config(
    page_title="Grupo Digital - Relatório de Performance",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "pagina" not in st.session_state:
    st.session_state.pagina = "upload"
if "tema" not in st.session_state:
    st.session_state.tema = "light"


def get_theme():
    if st.session_state.tema == "dark":
        return dict(
            verde="#4ECB8E", verde_escuro="#3FB57E",
            texto="#E8E8E8", texto_sec="#A0A0A0",
            fundo="#1A1D23", card_bg="#252930",
            borda="#3A3F47", fundo_sec="#2A2D35",
        )
    return dict(
        verde="#3FB57E", verde_escuro="#2A8B5F",
        texto="#2D3436", texto_sec="#95A5A6",
        fundo="#FFFFFF", card_bg="#FFFFFF",
        borda="#E8ECEF", fundo_sec="#F8F9FA",
    )


t = get_theme()
GREEN_SCALE = ["#D5F5E3", "#82E0AA", "#3FB57E", "#2A8B5F", "#1E6B47"]
BRAND_COLORS = ["#3FB57E", "#2A8B5F", "#1E6B47", "#82E0AA", "#D5F5E3", "#48C9B0", "#1ABC9C"]

dark_css = ""
if st.session_state.tema == "dark":
    dark_css = f"""
    .stApp, [data-testid="stAppViewContainer"], .main .block-container {{
        background-color: {t['fundo']} !important;
    }}
    section[data-testid="stSidebar"] {{
        background-color: {t['card_bg']} !important;
        border-right: 1px solid {t['borda']} !important;
    }}
    section[data-testid="stSidebar"] * {{ color: {t['texto']}; }}
    h1, h2, h3, h4, h5 {{ color: {t['texto']} !important; }}
    .stCaption, .stMarkdown p {{ color: {t['texto']} !important; }}
    [data-testid="stExpander"] {{ background-color: {t['card_bg']}; border-color: {t['borda']}; }}
    .stTabs [data-baseweb="tab-panel"] {{ background-color: {t['fundo']}; }}
    .stRadio label, .stCheckbox label {{ color: {t['texto']} !important; }}
    [data-testid="stHeader"] {{ background-color: {t['fundo']} !important; }}
    """

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
header[data-testid="stHeader"] {{ background: transparent; }}

h1 {{ color: {t['texto']}; font-weight: 800; letter-spacing: -0.5px; }}
h2 {{ color: {t['texto']}; font-weight: 700; letter-spacing: -0.3px; }}
h3 {{ color: {t['texto']}; font-weight: 600; }}

.gd-metric-card {{
    background: {t['card_bg']}; border: 1px solid {t['borda']}; border-radius: 12px;
    padding: 2rem 1rem; text-align: center; min-height: 150px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.06);
    transition: box-shadow 0.2s ease, transform 0.2s ease;
    display: flex; flex-direction: column; justify-content: center;
}}
.gd-metric-card:hover {{ box-shadow: 0 4px 12px rgba(0,0,0,0.12); transform: translateY(-2px); }}
.gd-metric-icon {{ font-size: 2rem; margin-bottom: 0.5rem; }}
.gd-metric-value {{
    font-size: 1.5rem; font-weight: 800; color: {t['verde']};
    margin: 0.3rem 0; line-height: 1.2; white-space: nowrap; overflow: visible;
}}
.gd-metric-value-lg {{ font-size: 2rem; }}
.gd-metric-label {{
    font-size: 0.72rem; font-weight: 600; color: {t['texto_sec']};
    text-transform: uppercase; letter-spacing: 0.6px; margin-top: 0.3rem;
}}
.gd-metric-hint {{
    font-size: 0.6rem; color: {t['texto_sec']}; margin-top: 0.2rem; opacity: 0.7;
}}

.gd-section {{
    background: {t['card_bg']}; border: 1px solid {t['borda']}; border-radius: 12px;
    padding: 2rem; box-shadow: 0 1px 3px rgba(0,0,0,0.04); margin-top: 0.5rem;
}}
.gd-section-title {{
    font-size: 1.3rem; font-weight: 700; color: {t['verde']};
    margin-bottom: 1.2rem; padding-bottom: 0.75rem;
    border-bottom: 2px solid {t['verde']};
}}

.gd-rank-card {{
    background: {t['card_bg']}; border: 1px solid {t['borda']}; border-radius: 12px;
    padding: 1.2rem 1.5rem; margin-bottom: 0.75rem;
    display: flex; align-items: center; gap: 1.2rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04); transition: box-shadow 0.2s ease;
}}
.gd-rank-card:hover {{ box-shadow: 0 4px 12px rgba(0,0,0,0.12); }}
.gd-rank-pos {{ font-size: 1.6rem; font-weight: 800; color: {t['verde']}; min-width: 45px; text-align: center; }}
.gd-rank-info {{ flex: 1; }}
.gd-rank-name {{ font-size: 1rem; font-weight: 700; color: {t['texto']}; margin-bottom: 0.2rem; }}
.gd-rank-stats {{ font-size: 0.8rem; color: {t['texto_sec']}; display: flex; gap: 1.2rem; flex-wrap: wrap; }}
.gd-rank-stats span {{ white-space: nowrap; }}
.gd-rank-highlight {{ font-size: 1.1rem; font-weight: 800; color: {t['verde']}; text-align: right; min-width: 70px; white-space: nowrap; }}

.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{
    background-color: {t['verde']}; border: none; border-radius: 8px; font-weight: 600;
    font-family: 'Inter', sans-serif; padding: 0.6rem 1.5rem; transition: background-color 0.2s ease;
}}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {{
    background-color: {t['verde_escuro']}; border: none;
}}
.stDataFrame {{ border-radius: 12px; overflow: hidden; }}
.gd-footer {{ text-align: center; padding: 2.5rem 0 1rem 0; color: {t['texto_sec']}; font-size: 0.8rem; border-top: 1px solid {t['borda']}; margin-top: 3rem; }}
.block-container {{ padding-top: 2.5rem; padding-bottom: 1rem; }}
.gd-divider {{ border: none; border-top: 1px solid {t['borda']}; margin: 1.8rem 0; }}
.gd-spacer {{ height: 2.5rem; }}
.gd-spacer-lg {{ height: 3rem; }}

.gd-nav a {{
    display: block; padding: 0.4rem 0.6rem; color: {t['texto']}; text-decoration: none;
    font-size: 0.85rem; border-radius: 6px; transition: background 0.15s;
}}
.gd-nav a:hover {{ background: {t['fundo_sec']}; }}

{dark_css}
</style>
""", unsafe_allow_html=True)


def fmt_brl(v):
    if pd.isna(v) or v == 0:
        return "R$ 0,00"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_brl_short(v):
    if pd.isna(v) or v == 0:
        return "R$ 0"
    if abs(v) >= 1_000_000:
        return f"R$ {v/1_000_000:,.2f} mi".replace(",", "X").replace(".", ",").replace("X", ".")
    if abs(v) >= 10_000:
        return f"R$ {v/1_000:,.1f} mil".replace(",", "X").replace(".", ",").replace("X", ".")
    return fmt_brl(v)

def fmt_num(v):
    return f"{int(v):,}".replace(",", ".")

def plotly_base():
    return dict(
        font=dict(family="Inter, sans-serif", color=t["texto"]),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=100, t=10, b=40),
        xaxis=dict(gridcolor=t["borda"], gridwidth=1),
        yaxis=dict(gridcolor=t["borda"], gridwidth=1),
        showlegend=False, coloraxis_showscale=False,
    )

SP_LG = '<div class="gd-spacer-lg"></div>'


# ══════════════════════════════════════════════════════════════════════════════
#  PÁGINA DE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.pagina == "upload":
    st.markdown("""<style>
    section[data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    </style>""", unsafe_allow_html=True)

    _, toggle_col = st.columns([6, 1])
    with toggle_col:
        dark = st.toggle("🌙", value=st.session_state.tema == "dark", key="dark_upload")
        if dark != (st.session_state.tema == "dark"):
            st.session_state.tema = "dark" if dark else "light"
            st.rerun()

    _, col_center, _ = st.columns([1, 2, 1])
    with col_center:
        st.markdown(f"""
        <div style="text-align: center; padding: 2rem 0 2rem;">
            <div style="font-size: 3rem; margin-bottom: 0.5rem;">📊</div>
            <div style="font-size: 2.2rem; font-weight: 800; color: {t['texto']}; margin-bottom: 0.2rem;">
                Grupo Digital
            </div>
            <div style="font-size: 0.9rem; color: {t['texto_sec']}; font-weight: 600; letter-spacing: 1.5px; text-transform: uppercase;">
                Relatório de Performance
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<hr class="gd-divider">', unsafe_allow_html=True)

        st.markdown("##### 📁 Bases de Leads")
        arquivos_leads = st.file_uploader(
            "Carregue uma ou mais planilhas de leads",
            type=["csv", "xlsx", "xls"], accept_multiple_files=True, key="leads_uploader",
        )
        st.markdown('<div style="height: 1rem;"></div>', unsafe_allow_html=True)

        st.markdown("##### 💰 Planilha de Vendas")
        arquivo_vendas = st.file_uploader(
            "Carregue a planilha de vendas",
            type=["csv", "xlsx", "xls"], key="vendas_uploader",
        )
        st.markdown('<div style="height: 1.5rem;"></div>', unsafe_allow_html=True)

        pode_gerar = len(arquivos_leads) > 0 and arquivo_vendas is not None
        gerar = st.button("Gerar Relatório", type="primary", use_container_width=True, disabled=not pode_gerar)
        if not pode_gerar:
            st.caption("Adicione ao menos uma base de leads e a planilha de vendas.")

        if gerar and pode_gerar:
            with st.spinner("Processando arquivos..."):
                erros = []
                leads_dfs = []
                for arq in arquivos_leads:
                    try:
                        df = ler_arquivo(arq)
                        valido, msg, warns = validar_planilha_leads(df)
                        if valido:
                            leads_dfs.append(df)
                            for w in warns:
                                st.warning(f"⚠️ {arq.name}: {w}")
                        else:
                            erros.append(f"{arq.name}: {msg}")
                    except Exception as e:
                        erros.append(f"{arq.name}: {e}")

                vendas_raw = None
                try:
                    vendas_raw = ler_arquivo(arquivo_vendas)
                    valido, msg, warns = validar_planilha_vendas(vendas_raw)
                    if not valido:
                        erros.append(f"Vendas: {msg}")
                        vendas_raw = None
                    else:
                        for w in warns:
                            st.warning(f"⚠️ Vendas: {w}")
                except Exception as e:
                    erros.append(f"Vendas: {e}")

                if erros:
                    for e in erros:
                        st.error(f"❌ {e}")
                elif not leads_dfs or vendas_raw is None:
                    st.error("❌ Nenhum arquivo válido para processar.")
                else:
                    leads_consolidado = pd.concat(leads_dfs, ignore_index=True)
                    leads_result, vendas_proc = cruzar_leads_vendas(leads_consolidado, vendas_raw)
                    st.session_state.leads_result = leads_result
                    st.session_state.vendas = vendas_proc
                    st.session_state.pagina = "relatorio"
                    st.rerun()

    st.markdown(f'<div class="gd-footer">Grupo Digital - Soluções Financeiras © {date.today().year}</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  PÁGINA DO RELATÓRIO
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.pagina == "relatorio":

    # ── Sidebar: tema + navegação ──────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"""
        <div style="text-align: center; padding: 0.8rem 0 1rem; border-bottom: 1px solid {t['borda']};">
            <div style="font-size: 1.1rem; font-weight: 800; color: {t['texto']};">Grupo Digital</div>
            <div style="font-size: 0.75rem; color: {t['texto_sec']};">Relatório de Performance</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f'<div style="height: 0.8rem;"></div>', unsafe_allow_html=True)
        dark = st.toggle("🌙 Modo Escuro", value=st.session_state.tema == "dark", key="dark_report")
        if dark != (st.session_state.tema == "dark"):
            st.session_state.tema = "dark" if dark else "light"
            st.rerun()

        st.markdown(f"""
        <div style="margin-top: 1.5rem;" class="gd-nav">
            <div style="font-size: 0.7rem; font-weight: 700; color: {t['texto_sec']}; text-transform: uppercase;
                        letter-spacing: 1px; margin-bottom: 0.6rem; padding-left: 0.6rem;">📑 Navegação</div>
            <a href="#visao-geral" target="_self">📊 Visão Geral</a>
            <a href="#bancos" target="_self">🏦 Bancos</a>
            <a href="#produtos" target="_self">📦 Produtos</a>
            <a href="#banco-produto" target="_self">🔀 Banco × Produto</a>
            <a href="#equipes" target="_self">👥 Equipes</a>
            <a href="#vendedores" target="_self">💼 Vendedores</a>
            <a href="#temporal" target="_self">📅 Temporal</a>
            <a href="#perfil" target="_self">🎂 Perfil Etário</a>
        </div>
        """, unsafe_allow_html=True)

    # ── Dados ──────────────────────────────────────────────────────────────
    leads_result = st.session_state.get("leads_result")
    vendas = st.session_state.get("vendas")

    if leads_result is None or vendas is None:
        st.warning("Nenhum dado processado. Volte à página de upload.")
        if st.button("← Voltar", key="btn_voltar_empty"):
            st.session_state.pagina = "upload"
            st.rerun()
        st.stop()

    if st.button("← Voltar", key="btn_voltar"):
        for k in ["leads_result", "vendas"]:
            st.session_state.pop(k, None)
        st.session_state.pagina = "upload"
        st.rerun()

    hoje = date.today().strftime("%d/%m/%Y")
    total_leads = len(leads_result)
    total_convertidos = int(leads_result["_convertido"].sum())
    total_vendas = len(vendas)
    taxa_conversao = (total_convertidos / total_leads * 100) if total_leads > 0 else 0
    valor_total = vendas["_valor"].sum()
    ticket_medio = (valor_total / total_vendas) if total_vendas > 0 else 0
    bancos_ativos = vendas["_banco"].nunique()
    vendedores_ativos = vendas.loc[vendas["_vendedor"] != "Não informado", "_vendedor"].nunique()
    data_min = vendas["_data_cadastro"].min()
    data_max = vendas["_data_cadastro"].max()
    periodo_txt = f"{data_min.strftime('%d/%m/%Y')} a {data_max.strftime('%d/%m/%Y')}" if pd.notna(data_min) and pd.notna(data_max) else "Período não disponível"

    # ═══════════════════════════════════════════════════════════════════════
    # 1. CABEÇALHO
    # ═══════════════════════════════════════════════════════════════════════
    st.markdown(f'<div id="visao-geral"></div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div style="margin-bottom: 2.5rem;">
        <h1 style="margin-bottom: 0.2rem; font-size: 2rem;">Relatório de Performance</h1>
        <p style="color: {t['texto_sec']}; font-size: 0.9rem; margin: 0;">
            {periodo_txt} &nbsp;·&nbsp; Gerado em {hoje}
        </p>
    </div>
    """, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # 2. CARDS DE MÉTRICAS
    # ═══════════════════════════════════════════════════════════════════════
    row1 = st.columns(4, gap="large")
    cards_r1 = [
        ("👥", fmt_num(total_leads), "Total de Leads", True, ""),
        ("✅", fmt_num(total_vendas), "Total de Vendas", True, ""),
        ("📈", f"{taxa_conversao:.1f}%", "Taxa de Conversão", True, "Convertidos / Leads × 100"),
        ("💰", fmt_brl_short(valor_total), "Valor Total Liberado", False, ""),
    ]
    for col, (icon, val, label, is_lg, hint) in zip(row1, cards_r1):
        cls = "gd-metric-value gd-metric-value-lg" if is_lg else "gd-metric-value"
        hint_html = f'<div class="gd-metric-hint">{hint}</div>' if hint else ""
        with col:
            st.markdown(f"""
            <div class="gd-metric-card" title="{hint}">
                <div class="gd-metric-icon">{icon}</div>
                <div class="{cls}">{val}</div>
                <div class="gd-metric-label">{label}</div>
                {hint_html}
            </div>""", unsafe_allow_html=True)

    st.markdown('<div style="height:1.2rem"></div>', unsafe_allow_html=True)

    row2 = st.columns(3, gap="large")
    cards_r2 = [
        ("🎟️", fmt_brl_short(ticket_medio), "Ticket Médio", "Valor Total / Vendas"),
        ("🏦", fmt_num(bancos_ativos), "Bancos Ativos", ""),
        ("👤", fmt_num(vendedores_ativos), "Vendedores Ativos", ""),
    ]
    for col, (icon, val, label, hint) in zip(row2, cards_r2):
        hint_html = f'<div class="gd-metric-hint">{hint}</div>' if hint else ""
        with col:
            st.markdown(f"""
            <div class="gd-metric-card" title="{hint}">
                <div class="gd-metric-icon">{icon}</div>
                <div class="gd-metric-value">{val}</div>
                <div class="gd-metric-label">{label}</div>
                {hint_html}
            </div>""", unsafe_allow_html=True)

    st.markdown(SP_LG, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # 3. PERFORMANCE POR BANCO
    # ═══════════════════════════════════════════════════════════════════════
    por_banco = vendas.groupby("_banco").agg(
        Vendas=("_valor", "count"), Valor_Total=("_valor", "sum"),
    ).reset_index().rename(columns={"_banco": "Banco"})
    por_banco["Ticket Médio"] = por_banco["Valor_Total"] / por_banco["Vendas"]
    por_banco["% do Total"] = (por_banco["Vendas"] / total_vendas * 100).round(1)
    por_banco = por_banco.sort_values("Valor_Total", ascending=False).reset_index(drop=True)

    st.markdown(f'<div id="bancos"></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="gd-section"><div class="gd-section-title">🏦 Performance por Banco</div>', unsafe_allow_html=True)
    display_banco = por_banco.rename(columns={"Valor_Total": "Valor Total"})
    st.markdown(render_tabela_html(
        display_banco, tema=st.session_state.tema, destacar_top3=True,
        colunas_moeda=["Valor Total", "Ticket Médio"], colunas_percentual=["% do Total"],
    ), unsafe_allow_html=True)

    st.markdown(f'<div style="height:1.2rem"></div>', unsafe_allow_html=True)
    fig_pie = px.pie(por_banco, values="Vendas", names="Banco", color_discrete_sequence=BRAND_COLORS, hole=0.4)
    fig_pie.update_traces(textinfo="percent+label", textfont=dict(size=12, family="Inter", color=t["texto"]))
    layout = plotly_base()
    layout.update({"height": 420, "margin": dict(l=10, r=10, t=10, b=10)})
    fig_pie.update_layout(**layout)
    st.plotly_chart(fig_pie, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown(SP_LG, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # 4. PERFORMANCE POR PRODUTO
    # ═══════════════════════════════════════════════════════════════════════
    por_produto = vendas.groupby("_produto").agg(
        Vendas=("_valor", "count"), Valor_Total=("_valor", "sum"),
    ).reset_index().rename(columns={"_produto": "Produto"})
    por_produto["Ticket Médio"] = por_produto["Valor_Total"] / por_produto["Vendas"]
    por_produto = por_produto.sort_values("Valor_Total", ascending=False).reset_index(drop=True)

    st.markdown(f'<div id="produtos"></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="gd-section"><div class="gd-section-title">📦 Performance por Produto</div>', unsafe_allow_html=True)
    prod_cols = st.columns(min(len(por_produto), 4), gap="large")
    for i, (_, row) in enumerate(por_produto.iterrows()):
        if i >= len(prod_cols):
            break
        with prod_cols[i]:
            st.markdown(f"""
            <div class="gd-metric-card">
                <div class="gd-metric-label">{row['Produto']}</div>
                <div class="gd-metric-value" style="font-size:1.5rem;">{fmt_num(row['Vendas'])}</div>
                <div class="gd-metric-label">vendas · {fmt_brl_short(row['Valor_Total'])}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown('<div style="height:1.5rem"></div>', unsafe_allow_html=True)
    fig_prod = px.bar(por_produto, x="Produto", y="Valor_Total", text="Valor_Total",
                      color="Produto", color_discrete_sequence=BRAND_COLORS)
    fig_prod.update_traces(texttemplate="R$ %{text:,.0f}", textposition="outside",
                           textfont=dict(size=13, family="Inter", color=t["texto"]))
    layout = plotly_base()
    layout.update({"xaxis_title": "", "yaxis_title": "", "height": 380})
    fig_prod.update_layout(**layout)
    st.plotly_chart(fig_prod, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown(SP_LG, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # 5. BANCO × PRODUTO
    # ═══════════════════════════════════════════════════════════════════════
    st.markdown(f'<div id="banco-produto"></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="gd-section"><div class="gd-section-title">🔀 Banco × Produto</div>', unsafe_allow_html=True)
    bp_vendas = vendas.groupby(["_banco", "_produto"]).agg(Vendas=("_valor", "count"), Valor=("_valor", "sum")).reset_index()
    pivot_vendas = bp_vendas.pivot_table(index="_banco", columns="_produto", values="Vendas", fill_value=0, aggfunc="sum")
    pivot_valor = bp_vendas.pivot_table(index="_banco", columns="_produto", values="Valor", fill_value=0, aggfunc="sum")
    hm_scale = [[0, t["fundo_sec"]], [0.5, "#82E0AA"], [1, "#1E6B47"]]

    tab_qty, tab_val = st.tabs(["Quantidade de Vendas", "Valor Total"])
    with tab_qty:
        fig_hm = go.Figure(data=go.Heatmap(
            z=pivot_vendas.values, x=pivot_vendas.columns.tolist(), y=pivot_vendas.index.tolist(),
            colorscale=hm_scale, text=pivot_vendas.values, texttemplate="%{text}",
            textfont=dict(size=14, family="Inter", color=t["texto"]),
            hovertemplate="Banco: %{y}<br>Produto: %{x}<br>Vendas: %{z}<extra></extra>",
        ))
        layout = plotly_base()
        layout.update({"height": max(280, len(pivot_vendas) * 55 + 100), "margin": dict(l=10, r=10, t=10, b=40)})
        fig_hm.update_layout(**layout)
        st.plotly_chart(fig_hm, use_container_width=True)

    with tab_val:
        fig_hm2 = go.Figure(data=go.Heatmap(
            z=pivot_valor.values, x=pivot_valor.columns.tolist(), y=pivot_valor.index.tolist(),
            colorscale=hm_scale, text=[[f"R$ {v:,.0f}" for v in row] for row in pivot_valor.values],
            texttemplate="%{text}", textfont=dict(size=13, family="Inter", color=t["texto"]),
            hovertemplate="Banco: %{y}<br>Produto: %{x}<br>Valor: %{text}<extra></extra>",
        ))
        layout = plotly_base()
        layout.update({"height": max(280, len(pivot_valor) * 55 + 100), "margin": dict(l=10, r=10, t=10, b=40)})
        fig_hm2.update_layout(**layout)
        st.plotly_chart(fig_hm2, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown(SP_LG, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # 6. PERFORMANCE POR EQUIPE
    # ═══════════════════════════════════════════════════════════════════════
    por_equipe = vendas.groupby("_equipe").agg(Vendas=("_valor", "count"), Valor_Total=("_valor", "sum")).reset_index().rename(columns={"_equipe": "Equipe"})
    por_equipe["Ticket Médio"] = por_equipe["Valor_Total"] / por_equipe["Vendas"]
    por_equipe = por_equipe.sort_values("Valor_Total", ascending=False).reset_index(drop=True)

    st.markdown(f'<div id="equipes"></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="gd-section"><div class="gd-section-title">👥 Performance por Equipe</div>', unsafe_allow_html=True)
    for i, row in por_equipe.iterrows():
        pos = ["🥇", "🥈", "🥉"][i] if i < 3 else f"{i+1}º"
        st.markdown(f"""
        <div class="gd-rank-card">
            <div class="gd-rank-pos">{pos}</div>
            <div class="gd-rank-info">
                <div class="gd-rank-name">{row['Equipe']}</div>
                <div class="gd-rank-stats"><span>{int(row['Vendas'])} vendas</span><span>Ticket: {fmt_brl_short(row['Ticket Médio'])}</span></div>
            </div>
            <div class="gd-rank-highlight">{fmt_brl_short(row['Valor_Total'])}</div>
        </div>""", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown(SP_LG, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # 7. PERFORMANCE POR VENDEDOR
    # ═══════════════════════════════════════════════════════════════════════
    por_vendedor = vendas.groupby("_vendedor").agg(Vendas=("_valor", "count"), Valor_Total=("_valor", "sum")).reset_index().rename(columns={"_vendedor": "Vendedor"})
    por_vendedor["Ticket Médio"] = por_vendedor["Valor_Total"] / por_vendedor["Vendas"]
    por_vendedor = por_vendedor.sort_values("Valor_Total", ascending=False).reset_index(drop=True)

    st.markdown(f'<div id="vendedores"></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="gd-section"><div class="gd-section-title">💼 Performance por Vendedor</div>', unsafe_allow_html=True)
    top15 = por_vendedor.head(15)
    for i, row in top15.iterrows():
        pos = ["🥇", "🥈", "🥉"][i] if i < 3 else f"{i+1}º"
        st.markdown(f"""
        <div class="gd-rank-card">
            <div class="gd-rank-pos">{pos}</div>
            <div class="gd-rank-info">
                <div class="gd-rank-name">{row['Vendedor']}</div>
                <div class="gd-rank-stats"><span>{int(row['Vendas'])} vendas</span><span>Ticket: {fmt_brl_short(row['Ticket Médio'])}</span></div>
            </div>
            <div class="gd-rank-highlight">{fmt_brl_short(row['Valor_Total'])}</div>
        </div>""", unsafe_allow_html=True)

    with st.expander("Ver tabela completa de vendedores"):
        display_vend = por_vendedor.rename(columns={"Valor_Total": "Valor Total"})
        st.markdown(render_tabela_html(
            display_vend, tema=st.session_state.tema, destacar_top3=True,
            colunas_moeda=["Valor Total", "Ticket Médio"],
        ), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown(SP_LG, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # 8. ANÁLISE TEMPORAL
    # ═══════════════════════════════════════════════════════════════════════
    vendas_com_data = vendas[vendas["_data_cadastro"].notna()].copy()
    fig_temporal = None

    st.markdown(f'<div id="temporal"></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="gd-section"><div class="gd-section-title">📅 Análise Temporal</div>', unsafe_allow_html=True)

    if vendas_com_data.empty:
        st.info("Sem dados de data de cadastro disponíveis para análise temporal.")
    else:
        agregacao = st.radio("Agregação", ["Diário", "Semanal", "Mensal"], horizontal=True)
        if agregacao == "Diário":
            vendas_com_data["_periodo"] = vendas_com_data["_data_cadastro"].dt.date
        elif agregacao == "Semanal":
            vendas_com_data["_periodo"] = vendas_com_data["_data_cadastro"].dt.to_period("W").apply(lambda x: x.start_time.date())
        else:
            vendas_com_data["_periodo"] = vendas_com_data["_data_cadastro"].dt.to_period("M").apply(lambda x: x.start_time.date())

        temporal = vendas_com_data.groupby("_periodo").agg(
            Vendas=("_valor", "count"), Valor=("_valor", "sum"),
        ).reset_index().rename(columns={"_periodo": "Período"}).sort_values("Período")

        fig_temporal = go.Figure()
        fig_temporal.add_trace(go.Bar(x=temporal["Período"], y=temporal["Valor"], name="Valor",
                               marker_color=t["verde"], opacity=0.7, yaxis="y"))
        fig_temporal.add_trace(go.Scatter(x=temporal["Período"], y=temporal["Vendas"], name="Vendas",
                                   line=dict(color=t["verde_escuro"], width=2.5), mode="lines+markers",
                                   marker=dict(size=6), yaxis="y2"))
        layout = plotly_base()
        layout.update({
            "showlegend": True,
            "yaxis": dict(title="Valor (R$)", gridcolor=t["borda"], side="left"),
            "yaxis2": dict(title="Qtd Vendas", overlaying="y", side="right", gridcolor="rgba(0,0,0,0)"),
            "xaxis": dict(gridcolor=t["borda"]),
            "legend": dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                           font=dict(color=t["texto"])),
            "height": 400, "margin": dict(l=10, r=10, t=30, b=40),
        })
        fig_temporal.update_layout(**layout)
        st.plotly_chart(fig_temporal, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown(SP_LG, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # 9. PERFIL DOS CONVERTIDOS
    # ═══════════════════════════════════════════════════════════════════════
    vendas_com_idade = vendas[vendas["_idade"].notna()].copy()
    fig_faixa = None

    st.markdown(f'<div id="perfil"></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="gd-section"><div class="gd-section-title">🎂 Perfil dos Convertidos (Faixa Etária)</div>', unsafe_allow_html=True)

    if vendas_com_idade.empty:
        st.info("Sem dados de data de nascimento disponíveis para análise de faixa etária.")
    else:
        idade_media = vendas_com_idade["_idade"].mean()
        faixa_order = ["<25", "25-34", "35-44", "45-54", "55-64", "65+"]
        faixas = vendas_com_idade.groupby("_faixa_etaria").agg(
            Quantidade=("_valor", "count")
        ).reindex(faixa_order).fillna(0).astype(int).reset_index().rename(columns={"_faixa_etaria": "Faixa Etária"})

        ci1, ci2 = st.columns([1, 3], gap="large")
        with ci1:
            st.markdown(f"""
            <div class="gd-metric-card">
                <div class="gd-metric-icon">🎂</div>
                <div class="gd-metric-value gd-metric-value-lg">{idade_media:.0f}</div>
                <div class="gd-metric-label">Idade Média</div>
            </div>""", unsafe_allow_html=True)

        with ci2:
            fig_faixa = px.bar(faixas, x="Faixa Etária", y="Quantidade", text="Quantidade",
                            color="Quantidade", color_continuous_scale=GREEN_SCALE)
            fig_faixa.update_traces(textposition="outside",
                                 textfont=dict(size=13, family="Inter", color=t["texto"]))
            layout = plotly_base()
            layout.update({"xaxis_title": "", "yaxis_title": "", "height": 320})
            fig_faixa.update_layout(**layout)
            st.plotly_chart(fig_faixa, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown(SP_LG, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # 10. EXPORTAÇÃO
    # ═══════════════════════════════════════════════════════════════════════
    resumo_geral = {
        "Total de Leads": total_leads, "Leads Convertidos": total_convertidos,
        "Total de Vendas": total_vendas, "Taxa de Conversão (%)": round(taxa_conversao, 2),
        "Valor Total Liberado": round(valor_total, 2), "Ticket Médio": round(ticket_medio, 2),
        "Bancos Ativos": bancos_ativos, "Vendedores Ativos": vendedores_ativos,
        "Período": periodo_txt, "Data do Relatório": hoje,
    }

    leads_conv = leads_result[leads_result["_convertido"]].copy()
    internals_to_keep = {"_nome", "_cpf", "_telefone", "_match_tipo"}
    renamed_targets = {"_nome": "Nome", "_cpf": "CPF", "_telefone": "Telefone"}
    originals_to_drop = set()
    for internal, display in renamed_targets.items():
        if internal in leads_conv.columns:
            originals_to_drop.update(c for c in leads_conv.columns if c == display or c.lower() == display.lower())
    cols_export = [c for c in leads_conv.columns if c not in originals_to_drop and (not c.startswith("_") or c in internals_to_keep)]
    leads_conv_export = leads_conv[cols_export].copy()
    renomear = {"_nome": "Nome", "_cpf": "CPF", "_telefone": "Telefone", "_match_tipo": "Tipo de Match"}
    leads_conv_export.rename(columns={k: v for k, v in renomear.items() if k in leads_conv_export.columns}, inplace=True)
    leads_conv_export = leads_conv_export.loc[:, ~leads_conv_export.columns.duplicated()]

    temporal_excel = pd.DataFrame()
    if not vendas_com_data.empty:
        vendas_com_data["_periodo_d"] = vendas_com_data["_data_cadastro"].dt.date
        temporal_excel = vendas_com_data.groupby("_periodo_d").agg(
            Vendas=("_valor", "count"), Valor=("_valor", "sum"),
        ).reset_index().rename(columns={"_periodo_d": "Data"}).sort_values("Data")

    bp_excel = bp_vendas.rename(columns={"_banco": "Banco", "_produto": "Produto"})

    # ── Gerar chart images para PDF ────────────────────────────────────────
    chart_images = {}
    try:
        chart_images["pie_banco"] = fig_pie.to_image(format="png", width=600, height=400, scale=2)
    except Exception:
        pass
    try:
        chart_images["heatmap"] = fig_hm.to_image(format="png", width=700, height=400, scale=2)
    except Exception:
        pass
    if fig_temporal is not None:
        try:
            chart_images["temporal"] = fig_temporal.to_image(format="png", width=800, height=400, scale=2)
        except Exception:
            pass
    if fig_faixa is not None:
        try:
            chart_images["faixa_etaria"] = fig_faixa.to_image(format="png", width=600, height=350, scale=2)
        except Exception:
            pass

    # ── Botões de download ─────────────────────────────────────────────────
    dl_excel, dl_pdf = st.columns(2, gap="large")
    with dl_excel:
        try:
            with st.spinner("Gerando Excel..."):
                excel_bytes = gerar_relatorio_excel(
                    leads_df=leads_result, vendas_df=vendas, resumo_geral=resumo_geral,
                    por_banco=por_banco, por_produto=por_produto, banco_produto=bp_excel,
                    por_equipe=por_equipe, por_vendedor=por_vendedor,
                    temporal=temporal_excel, leads_convertidos=leads_conv_export,
                )
            st.download_button(
                label="📊 Baixar Excel",
                data=excel_bytes,
                file_name=f"Relatorio_Performance_Grupo_Digital_{date.today().strftime('%d-%m-%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erro ao gerar Excel: {e}")

    with dl_pdf:
        try:
            with st.spinner("Gerando PDF..."):
                pdf_bytes = gerar_relatorio_pdf(
                    resumo_geral=resumo_geral, por_banco=por_banco, por_produto=por_produto,
                    por_equipe=por_equipe, por_vendedor=por_vendedor,
                    periodo_txt=periodo_txt, tema=st.session_state.tema,
                    chart_images=chart_images,
                )
            st.download_button(
                label="📄 Baixar PDF",
                data=pdf_bytes,
                file_name=f"Relatorio_Performance_Grupo_Digital_{date.today().strftime('%d-%m-%Y')}.pdf",
                mime="application/pdf",
                type="secondary", use_container_width=True,
            )
        except ImportError as e:
            st.warning(str(e))
        except Exception as e:
            st.error(f"Erro ao gerar PDF: {e}")

    st.markdown(f'<div class="gd-footer">Grupo Digital - Soluções Financeiras © {date.today().year}</div>', unsafe_allow_html=True)
