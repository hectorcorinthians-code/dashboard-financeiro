
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from utils.data_manager import (
    build_dashboard_metrics,
    get_sheet_dataframe,
    get_sheet_summary,
    list_sheets,
    save_sheet_dataframe,
    workbook_path,
)

st.set_page_config(page_title="Baumevie Dashboard", page_icon="📊", layout="wide")

APP_STATE_FILE = Path(__file__).resolve().parent / "data" / "app_state.json"


def apply_custom_style() -> None:
    st.markdown(
        """
        <style>
        .stApp { background: linear-gradient(180deg, #fffaf8 0%, #fff5f1 100%); }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #fff7f3 0%, #ffe9dc 100%);
            border-right: 1px solid rgba(224, 115, 82, 0.15);
        }
        .hero-card {
            background: linear-gradient(135deg, #111827 0%, #7c2d12 100%);
            color: white; padding: 28px; border-radius: 24px;
            box-shadow: 0 18px 40px rgba(124,45,18,0.16); margin-bottom: 18px;
        }
        .hero-card h1 { margin: 0; font-size: 2rem; line-height: 1.1; }
        .hero-card p { margin: 8px 0 0 0; color: rgba(255,255,255,0.88); font-size: 0.98rem; }
        .section-title { font-size: 1.15rem; font-weight: 700; margin-top: 8px; margin-bottom: 8px; color: #7c2d12; }
        div[data-testid="metric-container"] {
            background: rgba(255,255,255,0.92);
            border: 1px solid rgba(124,45,18,0.10);
            padding: 14px 16px; border-radius: 18px;
            box-shadow: 0 8px 20px rgba(15, 23, 42, 0.04);
        }
        .card {
            background: rgba(255,255,255,0.92);
            border: 1px solid rgba(124,45,18,0.10);
            padding: 16px; border-radius: 18px;
            box-shadow: 0 8px 20px rgba(15, 23, 42, 0.04);
        }
        .nav-note {
            font-size: 0.92rem;
            color: #6b7280;
            margin-top: -6px;
            margin-bottom: 12px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_sheet(sheet_name: str) -> pd.DataFrame:
    return get_sheet_dataframe(sheet_name)


@st.cache_data(show_spinner=False)
def load_metrics() -> dict:
    return build_dashboard_metrics()


def load_operational_summary() -> dict:
    default = {
        "titulo": "Resumo operacional",
        "status": "Operação em andamento",
        "responsavel": "",
        "observacoes": "",
    }
    if not APP_STATE_FILE.exists():
        APP_STATE_FILE.write_text(json.dumps(default, ensure_ascii=False, indent=2), encoding="utf-8")
        return default
    try:
        return json.loads(APP_STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return default


def save_operational_summary(data: dict) -> None:
    APP_STATE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def brl(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def clear_cache() -> None:
    load_sheet.clear()
    load_metrics.clear()


def filter_sales_by_period(sales_by_day: pd.DataFrame, period_label: str) -> pd.DataFrame:
    if sales_by_day.empty:
        return sales_by_day
    sales = sales_by_day.copy().sort_values("Data")
    if period_label == "Últimos 7 dias":
        return sales.tail(7)
    if period_label == "Últimos 30 dias":
        return sales.tail(30)
    return sales


def render_overview(metrics: dict, op_summary: dict, chart_period: str) -> None:
    st.markdown('<div class="section-title">Resumo geral do negócio</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Faturamento total", brl(metrics["faturamento_total"]))
    c2.metric("Lucro bruto", brl(metrics["lucro_total"]))
    c3.metric("Itens vendidos", f"{int(metrics['itens_vendidos'])}")
    c4.metric("Ticket médio", brl(metrics["ticket_medio"]))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Valor em estoque", brl(metrics["valor_estoque"]))
    c6.metric("Custo do estoque", brl(metrics["custo_estoque"]))
    c7.metric("Recebimentos", brl(metrics["total_recebido"]))
    c8.metric("Saldo em conta", brl(metrics["saldo_conta"]))

    sales_by_day = filter_sales_by_period(metrics["sales_by_day"], chart_period)

    row_top_left, row_top_right = st.columns((1.7, 1))
    with row_top_left:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Faturamento por dia")
        if not sales_by_day.empty:
            fig = px.area(sales_by_day, x="Data", y="Faturamento", markers=True)
            fig.update_layout(margin=dict(l=0, r=0, t=20, b=0), height=320)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Ainda não há vendas com data válida para montar a série diária.")
        st.markdown("</div>", unsafe_allow_html=True)

    with row_top_right:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader(op_summary.get("titulo", "Resumo operacional"))
        st.markdown(
            f"""
            **Status:** {op_summary.get('status', '-')}  
            **Responsável:** {op_summary.get('responsavel', '-') or '-'}  
            **Canal destaque:** {metrics['canal_destaque'] or '-'}  
            **Total de entradas:** {int(metrics['total_entradas'])}  
            **Produtos para repor:** {len(metrics['produtos_repor'])}
            """
        )
        if op_summary.get("observacoes"):
            st.caption(op_summary["observacoes"])
        if not metrics["produtos_repor"].empty:
            preview = metrics["produtos_repor"][
                [c for c in ["Produto", "Estoque Atual Calc", "Estoque Mínimo"] if c in metrics["produtos_repor"].columns]
            ].head(8)
            st.dataframe(preview, use_container_width=True, hide_index=True)
        else:
            st.success("Nenhum produto abaixo do estoque mínimo.")
        st.markdown("</div>", unsafe_allow_html=True)

    row_mid_1, row_mid_2 = st.columns(2)
    with row_mid_1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Faturamento por canal")
        by_channel = metrics["sales_by_channel"]
        if not by_channel.empty:
            fig = px.bar(by_channel, x="Canal de Venda", y="Valor Total (R$)", text_auto=".2s")
            fig.update_layout(margin=dict(l=0, r=0, t=20, b=0), height=340)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem dados de canal para exibir.")
        st.markdown("</div>", unsafe_allow_html=True)

    with row_mid_2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Lucro por canal")
        lucro_por_canal = metrics["profit_by_channel"]
        if not lucro_por_canal.empty:
            fig = px.bar(lucro_por_canal, x="Canal de Venda", y="Lucro Bruto (R$)", text_auto=".2s")
            fig.update_layout(margin=dict(l=0, r=0, t=20, b=0), height=340)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem dados de lucro por canal para exibir.")
        st.markdown("</div>", unsafe_allow_html=True)

    row_bottom_1, row_bottom_2 = st.columns(2)
    with row_bottom_1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Top produtos por faturamento")
        top_products = metrics["sales_by_product"].head(10)
        if not top_products.empty:
            fig = px.bar(top_products, x="Produto", y="Valor Total (R$)", hover_data=["Quantidade"])
            fig.update_layout(margin=dict(l=0, r=0, t=20, b=0), height=360, xaxis_title="Produto", yaxis_title="Faturamento")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem produtos vendidos para exibir ranking.")
        st.markdown("</div>", unsafe_allow_html=True)

    with row_bottom_2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Produtos com menor estoque")
        low_stock = metrics["produtos_repor"]
        if not low_stock.empty:
            plot_df = low_stock.copy().head(12)
            fig = px.bar(plot_df, x="Produto", y="Estoque Atual Calc", hover_data=["Estoque Mínimo"])
            fig.update_layout(margin=dict(l=0, r=0, t=20, b=0), height=360, xaxis_title="Produto", yaxis_title="Estoque atual")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.success("Nenhum produto abaixo do estoque mínimo.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Estoque por categoria")
    inv = metrics["inventory_by_category"]
    if not inv.empty:
        fig = px.pie(inv, names="Categoria", values="Valor em Estoque varejo(R$)", hole=0.48)
        fig.update_layout(margin=dict(l=0, r=0, t=20, b=0), height=360)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Sem dados de estoque por categoria.")
    st.markdown("</div>", unsafe_allow_html=True)


def render_editor(selected_sheet: str) -> None:
    st.markdown(f'<div class="section-title">Aba selecionada: {selected_sheet}</div>', unsafe_allow_html=True)
    if selected_sheet == "Estoque inicio Loja":
        st.info("A aba 'Estoque inicio Loja' foi normalizada para abrir corretamente, com cabeçalhos únicos e editáveis.")

    sheet_df = load_sheet(selected_sheet).copy()
    summary = get_sheet_summary(selected_sheet)

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Linhas", summary["rows"])
    s2.metric("Colunas", summary["columns"])
    s3.metric("Colunas numéricas", summary["numeric_columns"])
    s4.metric("Soma numérica", brl(summary["numeric_sum"]))

    if not summary["top_numeric_sums"].empty:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.caption("Principais somas da aba atual")
        st.dataframe(summary["top_numeric_sums"], use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    edited_df = st.data_editor(
        sheet_df,
        key=f"editor_{selected_sheet}",
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        height=520,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    a1, a2 = st.columns(2)
    with a1:
        if st.button(f"💾 Salvar alterações em {selected_sheet}", type="primary", use_container_width=True, key=f"save_{selected_sheet}"):
            try:
                save_sheet_dataframe(selected_sheet, edited_df)
                clear_cache()
                st.success("Alterações salvas com sucesso no arquivo Excel.")
                st.rerun()
            except Exception as exc:
                st.error(f"Não foi possível salvar as alterações: {exc}")

    with a2:
        csv_data = edited_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Baixar aba atual em CSV",
            data=csv_data,
            file_name=f"{selected_sheet.lower().replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True,
            key=f"download_{selected_sheet}",
        )


def render_operational_summary(op_summary: dict) -> None:
    st.markdown('<div class="section-title">Editar resumo operacional</div>', unsafe_allow_html=True)
    st.caption("Essas informações editam o cartão de resumo operacional sem mexer na lógica principal da planilha.")

    with st.form("operational_summary_form"):
        titulo = st.text_input("Título do resumo", value=op_summary.get("titulo", "Resumo operacional"))
        status = st.text_input("Status da operação", value=op_summary.get("status", "Operação em andamento"))
        responsavel = st.text_input("Responsável", value=op_summary.get("responsavel", ""))
        observacoes = st.text_area("Observações", value=op_summary.get("observacoes", ""), height=140)

        submitted = st.form_submit_button("Salvar resumo operacional", use_container_width=True)
        if submitted:
            save_operational_summary(
                {
                    "titulo": titulo,
                    "status": status,
                    "responsavel": responsavel,
                    "observacoes": observacoes,
                }
            )
            st.success("Resumo operacional atualizado com sucesso.")
            st.rerun()

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Pré-visualização")
    st.markdown(
        f"""
        **Título:** {op_summary.get('titulo', 'Resumo operacional')}  
        **Status:** {op_summary.get('status', '-')}  
        **Responsável:** {op_summary.get('responsavel', '-') or '-'}  
        **Observações:** {op_summary.get('observacoes', '-') or '-'}
        """
    )
    st.markdown("</div>", unsafe_allow_html=True)


apply_custom_style()

all_sheets = list_sheets()

if "selected_sheet" not in st.session_state:
    st.session_state.selected_sheet = all_sheets[0]
if "current_screen" not in st.session_state:
    st.session_state.current_screen = "📈 Visão geral"

with st.sidebar:
    st.markdown("## Configurações")
    st.write(f"**Arquivo de dados:** `{workbook_path().name}`")

    chosen_screen = st.radio(
        "Ir para",
        ["📈 Visão geral", "📝 Editor da aba", "⚙️ Resumo operacional"],
        index=["📈 Visão geral", "📝 Editor da aba", "⚙️ Resumo operacional"].index(st.session_state.current_screen),
        key="sidebar_screen_select",
    )
    if chosen_screen != st.session_state.current_screen:
        st.session_state.current_screen = chosen_screen
        st.rerun()

    current_index = all_sheets.index(st.session_state.selected_sheet)
    chosen_sheet = st.selectbox(
        "Escolha uma aba da planilha",
        all_sheets,
        index=current_index,
        key="sidebar_sheet_select"
    )
    if chosen_sheet != st.session_state.selected_sheet:
        st.session_state.selected_sheet = chosen_sheet
        clear_cache()
        st.rerun()

    st.caption(f"Tela atual: **{st.session_state.current_screen}**")
    st.caption(f"Aba atual: **{st.session_state.selected_sheet}**")

    chart_period = st.radio(
        "Período dos gráficos",
        ["Tudo", "Últimos 30 dias", "Últimos 7 dias"],
        index=0,
    )

    if st.button("🔄 Recarregar dashboard", use_container_width=True):
        clear_cache()
        st.rerun()

    st.info("Troquei as tabs por uma navegação lateral fixa para eliminar o problema de clique que não mudava de tela.")

selected_sheet = st.session_state.selected_sheet
current_screen = st.session_state.current_screen
metrics = load_metrics()
op_summary = load_operational_summary()

st.markdown(
    """
    <div class="hero-card">
        <h1>Baumevie Dashboard</h1>
        <p>Painel visual com navegação estável, edição das abas da planilha e resumo operacional separado.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(f'<div class="nav-note">Tela atual: <strong>{current_screen}</strong></div>', unsafe_allow_html=True)

if current_screen == "📈 Visão geral":
    render_overview(metrics, op_summary, chart_period)
elif current_screen == "📝 Editor da aba":
    render_editor(selected_sheet)
else:
    render_operational_summary(op_summary)
