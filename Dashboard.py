# Dashboard.py
from __future__ import annotations

import sys
from pathlib import Path
from io import BytesIO
from collections import Counter

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Dashboard Plano de Ação",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def resource_path(relative_path: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative_path


def safe_rerun():
    try:
        st.rerun()
    except Exception:
        try:
            st.experimental_rerun()
        except Exception:
            pass


def color_for_status(status: str) -> str:
    s = str(status).lower()
    if "conclu" in s:
        return "#16a34a"  # verde
    if "andam" in s:
        return "#f59e0b"  # laranja
    if "pend" in s:
        return "#ef4444"  # vermelho
    return "#64748b"      # cinza neutro


def compute_risk(row):
    base = 0
    status = str(row.get("Status", "")).lower()
    cost = float(row.get("Qual o Custo?", 0) or 0)

    if "pend" in status:
        base += 50
    if "andam" in status:
        base += 20
    if "conclu" in status:
        base -= 30

    if cost > 10000:
        base += 30
    elif cost > 5000:
        base += 15

    return max(0, min(100, base))


def to_excel_bytes(df_to_export: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_export.to_excel(writer, index=False, sheet_name="Plano_Acao")
    return output.getvalue()


def try_parse_dates(series: pd.Series) -> pd.Series:
    first = pd.to_datetime(series, dayfirst=True, errors="coerce")
    second = pd.to_datetime(series, dayfirst=False, errors="coerce")
    return first.fillna(second)


def kpi_card(title: str, value: str, subtitle: str = "") -> str:
    return f"""
    <div style="
        padding:16px 18px;
        border-radius:18px;
        border:1px solid rgba(148,163,184,0.25);
        background:linear-gradient(180deg, rgba(255,255,255,0.85), rgba(248,250,252,0.95));
        box-shadow:0 2px 10px rgba(15,23,42,0.04);
        height:100%;
    ">
        <div style="font-size:0.86rem;color:#64748b;margin-bottom:6px;">{title}</div>
        <div style="font-size:1.55rem;font-weight:700;color:#0f172a;line-height:1.1;">{value}</div>
        <div style="font-size:0.8rem;color:#94a3b8;margin-top:6px;">{subtitle}</div>
    </div>
    """


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar / Load
# ─────────────────────────────────────────────────────────────────────────────

st.sidebar.title("Configurações")
uploaded_file = st.sidebar.file_uploader("Envie Plano_Acao.xlsx", type=["xlsx", "xls"])


@st.cache_data(ttl=60)
def load_data(file_buffer) -> pd.DataFrame:
    df = pd.read_excel(file_buffer, sheet_name=0, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    expected = [
        "O que?", "Por quê?", "Onde?", "Quem?", "Data Inicial",
        "Data Final", "Como?", "Qual o Custo?", "Status", "Contratações"
    ]
    for col in expected:
        if col not in df.columns:
            df[col] = ""

    df["Qual o Custo?"] = pd.to_numeric(df["Qual o Custo?"], errors="coerce").fillna(0)
    df["Data Inicial"] = try_parse_dates(df["Data Inicial"])
    df["Data Final"] = try_parse_dates(df["Data Final"])

    for c in ["O que?", "Quem?", "Status", "Como?", "Contratações", "Onde?", "Por quê?"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str)

    return df


try:
    if uploaded_file is not None:
        df = load_data(uploaded_file)
    else:
        local_file = resource_path("Plano_Acao.xlsx")
        if not local_file.exists():
            st.error(
                "Arquivo Excel 'Plano_Acao.xlsx' não encontrado. "
                "Faça upload na barra lateral com o nome Plano_Acao.xlsx."
            )
            st.stop()

        with open(local_file, "rb") as f:
            df = load_data(f)

except Exception as e:
    st.error(f"Erro ao carregar o arquivo: {e}")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# Normalização e métricas
# ─────────────────────────────────────────────────────────────────────────────

df["Status"] = df["Status"].fillna("Pendente").astype(str)
df["Quem?"] = df["Quem?"].fillna("").astype(str)
df["O que?"] = df["O que?"].fillna("").astype(str)
df["RiskScore"] = df.apply(compute_risk, axis=1)

total_actions = len(df)
total_cost = float(df["Qual o Custo?"].sum())
completed = int(df["Status"].str.lower().str.contains("conclu", na=False).sum())
in_progress = int(df["Status"].str.lower().str.contains("andam|andamento", na=False).sum())
pendentes_count = int(df["Status"].str.lower().str.contains("pend", na=False).sum())
pct_completed = (completed / total_actions * 100) if total_actions else 0
avg_cost = float(df["Qual o Custo?"].mean()) if total_actions else 0
pending_cost = float(df.loc[df["Status"].str.lower().str.contains("pend", na=False), "Qual o Custo?"].sum())

today = pd.Timestamp.now().normalize()
mask_overdue = (
    df["Data Final"].notna()
    & (df["Data Final"] < today)
    & (~df["Status"].str.lower().str.contains("conclu", na=False))
)
overdue_count = int(mask_overdue.sum())

top_who = df["Quem?"].replace("", "Não informado").value_counts().head(5)


# ─────────────────────────────────────────────────────────────────────────────
# Filtros
# ─────────────────────────────────────────────────────────────────────────────

st.sidebar.header("Filtros rápidos")

status_options = sorted([s for s in df["Status"].dropna().astype(str).unique().tolist() if s.strip()])
status_filter = st.sidebar.multiselect(
    "Status",
    options=status_options,
    default=status_options,
)

who_filter = st.sidebar.text_input("Filtrar por 'Quem' (texto)")
min_cost = st.sidebar.number_input("Custo mínimo (R$)", value=0, step=100)

df_filtered = df.copy()

if status_filter:
    df_filtered = df_filtered[df_filtered["Status"].isin(status_filter)]

if who_filter:
    df_filtered = df_filtered[df_filtered["Quem?"].str.contains(who_filter, case=False, na=False)]

if min_cost:
    df_filtered = df_filtered[df_filtered["Qual o Custo?"] >= min_cost]


# ─────────────────────────────────────────────────────────────────────────────
# Layout principal
# ─────────────────────────────────────────────────────────────────────────────

st.title("Dashboard — Plano de Ação")
st.markdown(
    "Painel executivo para monitoramento e análise das ações planejadas. "
    "***Use os filtros na barra lateral para refinar os dados exibidos.***"
)

c1, c2, c3, c4, c5, c6 = st.columns(6)

cards = [
    ("Total de ações", f"{total_actions}", "Quantidade total de registros"),
    ("Custo total", f"R$ {total_cost:,.2f}", "Soma de todos os custos"),
    ("Concluídas", f"{pct_completed:.1f}% ({completed})", "Percentual e total concluído"),
    ("Em andamento", f"{in_progress}", "Status em andamento"),
    ("Pendentes", f"{pendentes_count}", "Itens pendentes"),
    ("Atrasadas", f"{overdue_count}", "Data final vencida"),
]

for col, (title, value, subtitle) in zip([c1, c2, c3, c4, c5, c6], cards):
    col.markdown(kpi_card(title, value, subtitle), unsafe_allow_html=True)

st.markdown("---")


# ─────────────────────────────────────────────────────────────────────────────
# Gráficos
# ─────────────────────────────────────────────────────────────────────────────

left, right = st.columns([2, 1])

with left:
    st.subheader("Visão por Status")

    status_counts = df_filtered["Status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Contagem"]

    if not status_counts.empty:
        fig_pie = px.pie(
            status_counts,
            names="Status",
            values="Contagem",
            hole=0.45,
        )
        fig_pie.update_traces(
            marker=dict(colors=[color_for_status(s) for s in status_counts["Status"]]),
            textinfo="percent+label",
        )
        fig_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("Sem dados para o gráfico de status.")

    st.subheader("Custo por Status")
    cost_by_status = df_filtered.groupby("Status", dropna=False)["Qual o Custo?"].sum().reset_index()

    if not cost_by_status.empty:
        unique_statuses = cost_by_status["Status"].astype(str).unique().tolist()
        color_map = {s: color_for_status(s) for s in unique_statuses}

        fig_bar = px.bar(
            cost_by_status,
            x="Qual o Custo?",
            y="Status",
            orientation="h",
            color="Status",
            color_discrete_map=color_map,
        )
        fig_bar.update_layout(
            showlegend=False,
            margin=dict(t=6, b=6, l=6, r=6),
            height=320,
            xaxis_title="Custo (R$)",
            yaxis_title="Status",
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("Sem dados para o gráfico de custo.")

    st.subheader("Tendência de custo (mensal)")
    if df_filtered["Data Inicial"].notna().any():
        df_time = df_filtered.copy()
        df_time["Mes"] = df_time["Data Inicial"].dt.to_period("M").dt.to_timestamp()
        trend = df_time.groupby("Mes")["Qual o Custo?"].sum().reset_index()

        if not trend.empty:
            fig_trend = px.line(trend, x="Mes", y="Qual o Custo?", markers=True)
            fig_trend.update_layout(margin=dict(t=10, b=10, l=10, r=10), xaxis_title="Mês", yaxis_title="Custo (R$)")
            st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.info("Sem dados suficientes para a tendência mensal.")
    else:
        st.info("Preencha a coluna 'Data Inicial' para ver a tendência mensal.")

with right:
    st.subheader("Métricas executivas")
    st.markdown(f"- **Custo médio por ação:** R$ {avg_cost:,.2f}")
    st.markdown(f"- **Custo pendente total:** R$ {pending_cost:,.2f}")
    st.markdown("- **Top responsáveis:**")
    for who, cnt in top_who.items():
        st.markdown(f"  - **{who}**: {cnt}")

st.markdown("---")


# ─────────────────────────────────────────────────────────────────────────────
# Cronograma
# ─────────────────────────────────────────────────────────────────────────────

st.subheader("Cronograma de Tarefas")

timeline_df = df_filtered.copy().reset_index(drop=True)
timeline_df["Start"] = pd.to_datetime(timeline_df["Data Inicial"], errors="coerce")
timeline_df["End"] = pd.to_datetime(timeline_df["Data Final"], errors="coerce")
timeline_df["End"] = timeline_df["End"].fillna(timeline_df["Start"] + pd.to_timedelta(7, unit="d"))
timeline_df = timeline_df.dropna(subset=["Start"]).copy()

if timeline_df.empty:
    st.info("Não há datas válidas na coluna 'Data Inicial' para construir o cronograma.")
else:
    timeline_df = timeline_df.sort_values(by="Start").reset_index(drop=True)

    raw_labels = timeline_df.apply(
        lambda r: (str(r["O que?"]).strip() or str(r["Quem?"]).strip() or "Sem título")[:55],
        axis=1
    ).tolist()

    label_count = Counter(raw_labels)
    seen = Counter()
    final_labels = []

    for lbl in raw_labels:
        if label_count[lbl] > 1:
            seen[lbl] += 1
            final_labels.append(f"{lbl} ({seen[lbl]})")
        else:
            final_labels.append(lbl)

    timeline_df["Label"] = final_labels
    color_map = {
        s: color_for_status(s)
        for s in timeline_df["Status"].fillna("Pendente").astype(str).unique().tolist()
    }

    periodo_sel = st.selectbox(
        "Período",
        ["Todos", "Próximos 30 dias", "Próximos 90 dias", "Este ano"],
        key="gantt_periodo",
    )

    today_ts = pd.Timestamp.now().normalize()
    df_gantt = timeline_df.copy()

    if periodo_sel == "Próximos 30 dias":
        df_gantt = df_gantt[df_gantt["Start"] <= today_ts + pd.to_timedelta(30, unit="d")]
    elif periodo_sel == "Próximos 90 dias":
        df_gantt = df_gantt[df_gantt["Start"] <= today_ts + pd.to_timedelta(90, unit="d")]
    elif periodo_sel == "Este ano":
        df_gantt = df_gantt[df_gantt["Start"].dt.year == today_ts.year]

    if df_gantt.empty:
        st.info("Nenhuma tarefa no período selecionado.")
    else:
        fig_tl = px.timeline(
            df_gantt,
            x_start="Start",
            x_end="End",
            y="Label",
            color="Status",
            color_discrete_map=color_map,
            custom_data=["Quem?", "O que?", "Qual o Custo?", "Status", "RiskScore"],
        )
        fig_tl.update_yaxes(autorange="reversed")
        fig_tl.add_vline(
            x=today_ts,
            line_dash="dot",
            line_color="rgba(239,68,68,0.75)",
            line_width=2,
        )
        fig_tl.update_layout(
            margin=dict(t=10, b=10, l=10, r=10),
            xaxis_title="Período",
            yaxis_title="Ações",
        )
        st.plotly_chart(fig_tl, use_container_width=True)

st.markdown("---")


# ─────────────────────────────────────────────────────────────────────────────
# Tabela
# ─────────────────────────────────────────────────────────────────────────────

st.subheader("Tabela de Ações")

default_cols = ["O que?", "Quem?", "Status", "Data Inicial", "Data Final", "Qual o Custo?"]
cols_to_show = [c for c in default_cols if c in df_filtered.columns]

if not cols_to_show:
    cols_to_show = list(df_filtered.columns)

st.dataframe(
    df_filtered[cols_to_show].reset_index(drop=True),
    use_container_width=True,
    height=360,
)


# ─────────────────────────────────────────────────────────────────────────────
# Exportação
# ─────────────────────────────────────────────────────────────────────────────

st.download_button(
    "Exportar filtrados para Excel",
    data=to_excel_bytes(df_filtered),
    file_name="plano_acao_filtrado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if st.button("Atualizar agora"):
    safe_rerun()

st.caption('Caso queira utilizar outra planilha como base, envie um novo arquivo "Plano_Acao.xlsx" no upload.')
