import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import date

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Inadimplência • Ranking & Mapeamento", layout="wide")
DEFAULT_EXCEL_NAME = "Indicador de Inadimplência .xlsx"

# =========================
# FORMATTERS (BR)
# =========================
def brl(x) -> str:
    if pd.isna(x):
        return ""
    try:
        x = float(x)
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return ""

def brl_money(x) -> str:
    v = brl(x)
    return f"R$ {v}" if v != "" else ""

def br_pct(x) -> str:
    if pd.isna(x):
        return ""
    try:
        x = float(x) * 100.0
        return f"{x:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return ""

def br_date(dt) -> str:
    if pd.isna(dt):
        return ""
    try:
        return pd.to_datetime(dt).strftime("%d/%m/%Y")
    except Exception:
        return ""

# =========================
# PARSERS
# =========================
def coerce_money(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace({"nan": "", "None": ""})
    has_comma = s.str.contains(",", na=False)
    s = s.where(~has_comma, s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def coerce_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True)

def normalize_text(s: pd.Series) -> pd.Series:
    return s.astype(str).fillna("").str.strip()

# =========================
# LOAD
# =========================
@st.cache_data(show_spinner=False)
def load_base_from_excel(file_or_path) -> pd.DataFrame:
    df = pd.read_excel(file_or_path, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    # Colunas mínimas (tolerante)
    for col in ["CLIENTE", "EMPRESA", "VENCTO", "DTA.CAD", "V.ORIGI", "DUPLICATA", "HISTORICO", "OPERADOR", "OPERADOR2"]:
        if col not in df.columns:
            df[col] = ""

    df["CLIENTE"] = normalize_text(df["CLIENTE"]).replace("", "SEM CLIENTE")
    df["EMPRESA"] = normalize_text(df["EMPRESA"]).replace("", "SEM LOJA")

    # VENDEDOR: prioridade OPERADOR (se tiver), senão OPERADOR2
    op = normalize_text(df["OPERADOR"])
    op2 = normalize_text(df["OPERADOR2"])
    df["VENDEDOR"] = np.where(op != "", op, op2)
    df["VENDEDOR"] = pd.Series(df["VENDEDOR"]).astype(str).str.strip().replace("", "SEM VENDEDOR")

    # Datas
    df["VENCTO_DT"] = coerce_date(df["VENCTO"])
    df["DTA_CAD_DT"] = coerce_date(df["DTA.CAD"])

    # Valor em aberto
    df["VALOR"] = coerce_money(df["V.ORIGI"])

    # Tempo vs hoje (base VENCTO)
    today = pd.Timestamp(date.today())
    df["DIAS_EM_ABERTO"] = (today - df["VENCTO_DT"]).dt.days
    df.loc[df["VENCTO_DT"].isna(), "DIAS_EM_ABERTO"] = np.nan

    # Ano para filtro
    df["ANO"] = df["VENCTO_DT"].dt.year

    return df

def init_state():
    st.session_state.setdefault("empresa_sel", None)
    st.session_state.setdefault("vendedor_sel", None)
    st.session_state.setdefault("cliente_sel", None)

init_state()

# =========================
# SIDEBAR • BASE
# =========================
st.sidebar.title("Base")
up = st.sidebar.file_uploader("Enviar Excel (.xlsx)", type=["xlsx"])

if up is None:
    try:
        df = load_base_from_excel(DEFAULT_EXCEL_NAME)
        st.sidebar.info(f"Lendo arquivo local: {DEFAULT_EXCEL_NAME}")
    except Exception:
        st.warning(
            "Envie o Excel no menu à esquerda, ou coloque o arquivo "
            f"'{DEFAULT_EXCEL_NAME}' na mesma pasta do INAD.py."
        )
        st.stop()
else:
    df = load_base_from_excel(up)

# =========================
# SIDEBAR • FILTROS
# =========================
st.sidebar.title("Filtros")

lojas = sorted(df["EMPRESA"].unique().tolist())
sel_lojas = st.sidebar.multiselect("Loja (EMPRESA)", lojas, default=lojas)
df_f = df[df["EMPRESA"].isin(sel_lojas)].copy()

anos = sorted([int(a) for a in df_f["ANO"].dropna().unique().tolist()])
if anos:
    sel_anos = st.sidebar.multiselect("Ano (VENCTO)", anos, default=anos)
    df_f = df_f[df_f["ANO"].isin(sel_anos)].copy()
else:
    st.sidebar.warning("Sem datas válidas em VENCTO para filtro de Ano.")

st.sidebar.markdown("---")
min_dt = df_f["VENCTO_DT"].min()
max_dt = df_f["VENCTO_DT"].max()
if pd.isna(min_dt) or pd.isna(max_dt):
    st.sidebar.warning("Sem datas válidas para filtro por período.")
else:
    start_date, end_date = st.sidebar.date_input(
        "Período (VENCTO)",
        value=(min_dt.date(), max_dt.date()),
        min_value=min_dt.date(),
        max_value=max_dt.date()
    )
    df_f = df_f[
        (df_f["VENCTO_DT"].dt.date >= start_date) &
        (df_f["VENCTO_DT"].dt.date <= end_date)
    ].copy()

st.sidebar.markdown("---")
include_credits = st.sidebar.checkbox("Incluir créditos/negativos (VALOR < 0)", value=False)
if not include_credits:
    df_f = df_f[df_f["VALOR"] > 0].copy()

only_overdue = st.sidebar.checkbox("Somente vencidos (hoje ≥ VENCTO)", value=False)
if only_overdue:
    df_f = df_f[df_f["DIAS_EM_ABERTO"].notna() & (df_f["DIAS_EM_ABERTO"] >= 0)].copy()

# =========================
# HEADER / KPIs
# =========================
st.title("Dashboard • Inadimplência (V.ORIGI)")

total_open = float(df_f["VALOR"].sum())
qt_tit = int(len(df_f))
qt_cli = int(df_f["CLIENTE"].nunique())
ticket = (total_open / qt_tit) if qt_tit else 0.0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total em aberto", f"R$ {brl(total_open)}")
c2.metric("Qtd. Títulos", f"{qt_tit:,}".replace(",", "."))
c3.metric("Qtd. Clientes", f"{qt_cli:,}".replace(",", "."))
c4.metric("Ticket médio (título)", f"R$ {brl(ticket)}")

st.caption(f"Data base: **{date.today().strftime('%d/%m/%Y')}** | Valor: **V.ORIGI** | Vendedor: **OPERADOR/OPERADOR2**")

# =========================
# 1) EMPRESAS • TABELA + MAPA CLICÁVEL
# =========================
st.subheader("1) Ranking • Empresas (EMPRESA) + Mapa (clique para abrir clientes)")

emp_agg = (
    df_f.groupby("EMPRESA", dropna=False)
    .agg(
        VALOR_ABERTO=("VALOR", "sum"),
        TITULOS=("VALOR", "size"),
        CLIENTES=("CLIENTE", "nunique"),
    )
    .reset_index()
    .sort_values("VALOR_ABERTO", ascending=False)
)

colA, colB = st.columns([1.2, 1.8])

with colA:
    emp_show = emp_agg.copy()
    emp_show["VALOR_ABERTO"] = emp_show["VALOR_ABERTO"].apply(brl_money)
    st.dataframe(emp_show, use_container_width=True, hide_index=True)

with colB:
    fig_emp = px.treemap(
        emp_agg,
        path=["EMPRESA"],
        values="VALOR_ABERTO",
        hover_data={"TITULOS": True, "CLIENTES": True, "VALOR_ABERTO": ":,.2f"},
        title="Mapa de lojas (tamanho = valor em aberto). Selecione um bloco para filtrar."
    )
    fig_emp.update_layout(height=520, margin=dict(l=10, r=10, t=60, b=10))

    sel = st.plotly_chart(fig_emp, use_container_width=True, on_select="rerun", selection_mode="points")
    if isinstance(sel, dict) and sel.get("selection") and sel["selection"].get("points"):
        label = sel["selection"]["points"][0].get("label")
        if label:
            st.session_state["empresa_sel"] = label
            st.session_state["vendedor_sel"] = None
            st.session_state["cliente_sel"] = None

emp_options = emp_agg["EMPRESA"].tolist()
if not st.session_state["empresa_sel"] or st.session_state["empresa_sel"] not in emp_options:
    st.session_state["empresa_sel"] = emp_options[0] if emp_options else None

empresa_sel = st.selectbox(
    "Empresa selecionada",
    options=emp_options,
    index=emp_options.index(st.session_state["empresa_sel"]) if st.session_state["empresa_sel"] in emp_options else 0
)
st.session_state["empresa_sel"] = empresa_sel

df_emp = df_f[df_f["EMPRESA"] == empresa_sel].copy()

st.markdown("**Clientes com títulos em aberto (empresa selecionada) — maior → menor**")
cli_emp_rank = (
    df_emp.groupby("CLIENTE", dropna=False)
    .agg(VALOR_ABERTO=("VALOR", "sum"))
    .reset_index()
    .sort_values("VALOR_ABERTO", ascending=False)
)

cli_emp_rank_show = cli_emp_rank.copy()
cli_emp_rank_show["VALOR_ABERTO"] = cli_emp_rank_show["VALOR_ABERTO"].apply(brl_money)
st.dataframe(cli_emp_rank_show, use_container_width=True, hide_index=True)

st.markdown("---")

# =========================
# 2) DRILL • CLIENTES NA EMPRESA (COLUNAS NOVAS + ORDENAÇÃO)
# =========================
st.subheader("2) Drill • Clientes na empresa (ordenável por valor e datas)")

cli_emp = (
    df_emp.groupby("CLIENTE", dropna=False)
    .agg(
        VALOR_ABERTO=("VALOR", "sum"),
        DATA_ULTIMA_COMPRA=("DTA_CAD_DT", "max"),
        DATA_ULTIMO_VENCTO=("VENCTO_DT", "max"),
    )
    .reset_index()
)

total_emp = float(cli_emp["VALOR_ABERTO"].sum())
cli_emp["PCT_EMPRESA"] = np.where(total_emp > 0, cli_emp["VALOR_ABERTO"] / total_emp, 0.0)

sort_choice = st.selectbox(
    "Ordenar Drill de Clientes por:",
    [
        "Valor (maior → menor)",
        "Valor (menor → maior)",
        "Última compra (mais novo → mais antigo)",
        "Última compra (mais antigo → mais novo)",
        "Último vencimento (mais novo → mais antigo)",
        "Último vencimento (mais antigo → mais novo)",
    ],
    index=0
)

cli_sort = cli_emp.copy()
if sort_choice == "Valor (maior → menor)":
    cli_sort = cli_sort.sort_values("VALOR_ABERTO", ascending=False)
elif sort_choice == "Valor (menor → maior)":
    cli_sort = cli_sort.sort_values("VALOR_ABERTO", ascending=True)
elif sort_choice == "Última compra (mais novo → mais antigo)":
    cli_sort = cli_sort.sort_values("DATA_ULTIMA_COMPRA", ascending=False)
elif sort_choice == "Última compra (mais antigo → mais novo)":
    cli_sort = cli_sort.sort_values("DATA_ULTIMA_COMPRA", ascending=True)
elif sort_choice == "Último vencimento (mais novo → mais antigo)":
    cli_sort = cli_sort.sort_values("DATA_ULTIMO_VENCTO", ascending=False)
else:
    cli_sort = cli_sort.sort_values("DATA_ULTIMO_VENCTO", ascending=True)

cli_show = cli_sort.copy()
cli_show["VALOR_ABERTO"] = cli_show["VALOR_ABERTO"].apply(brl_money)
cli_show["PCT_EMPRESA"] = cli_show["PCT_EMPRESA"].apply(br_pct)
cli_show["DATA_ULTIMA_COMPRA"] = cli_show["DATA_ULTIMA_COMPRA"].apply(br_date)
cli_show["DATA_ULTIMO_VENCTO"] = cli_show["DATA_ULTIMO_VENCTO"].apply(br_date)

st.dataframe(
    cli_show[["CLIENTE", "VALOR_ABERTO", "PCT_EMPRESA", "DATA_ULTIMA_COMPRA", "DATA_ULTIMO_VENCTO"]],
    use_container_width=True,
    hide_index=True
)

st.markdown("---")

# =========================
# 3) MAPA POR VENDEDOR (CLIQUE) + LISTA DE CLIENTES DO VENDEDOR
# =========================
st.subheader("3) Mapeamento • Vendedores (OPERADOR/OPERADOR2) na empresa (clique para abrir clientes)")

ven_agg = (
    df_emp.groupby("VENDEDOR", dropna=False)
    .agg(
        VALOR_ABERTO=("VALOR", "sum"),
        TITULOS=("VALOR", "size"),
        CLIENTES=("CLIENTE", "nunique"),
    )
    .reset_index()
    .sort_values("VALOR_ABERTO", ascending=False)
)

colC, colD = st.columns([1.2, 1.8])

with colC:
    ven_show = ven_agg.copy()
    ven_show["VALOR_ABERTO"] = ven_show["VALOR_ABERTO"].apply(brl_money)
    st.dataframe(ven_show, use_container_width=True, hide_index=True)

with colD:
    fig_ven = px.treemap(
        ven_agg,
        path=["VENDEDOR"],
        values="VALOR_ABERTO",
        hover_data={"TITULOS": True, "CLIENTES": True, "VALOR_ABERTO": ":,.2f"},
        title="Mapa de vendedores (tamanho = valor). Selecione um bloco para filtrar."
    )
    fig_ven.update_layout(height=520, margin=dict(l=10, r=10, t=60, b=10))

    sel2 = st.plotly_chart(fig_ven, use_container_width=True, on_select="rerun", selection_mode="points")
    if isinstance(sel2, dict) and sel2.get("selection") and sel2["selection"].get("points"):
        label2 = sel2["selection"]["points"][0].get("label")
        if label2:
            st.session_state["vendedor_sel"] = label2
            st.session_state["cliente_sel"] = None

ven_options = ven_agg["VENDEDOR"].tolist()
if st.session_state["vendedor_sel"] and st.session_state["vendedor_sel"] not in ven_options:
    st.session_state["vendedor_sel"] = None

vendedor_sel = st.selectbox(
    "Vendedor selecionado (opcional)",
    options=["(todos)"] + ven_options,
    index=0 if not st.session_state["vendedor_sel"] else (["(todos)"] + ven_options).index(st.session_state["vendedor_sel"])
)

if vendedor_sel != "(todos)":
    st.session_state["vendedor_sel"] = vendedor_sel
    df_v = df_emp[df_emp["VENDEDOR"] == vendedor_sel].copy()
else:
    df_v = df_emp.copy()
    st.session_state["vendedor_sel"] = None

st.markdown("**Clientes com títulos em aberto (vendedor selecionado) — maior → menor**")
cli_v_rank = (
    df_v.groupby("CLIENTE", dropna=False)
    .agg(
        VALOR_ABERTO=("VALOR", "sum"),
        TITULOS=("VALOR", "size"),
        DATA_ULTIMA_COMPRA=("DTA_CAD_DT", "max"),
        DATA_ULTIMO_VENCTO=("VENCTO_DT", "max"),
    )
    .reset_index()
    .sort_values("VALOR_ABERTO", ascending=False)
)

cli_v_show = cli_v_rank.copy()
cli_v_show["VALOR_ABERTO"] = cli_v_show["VALOR_ABERTO"].apply(brl_money)
cli_v_show["DATA_ULTIMA_COMPRA"] = cli_v_show["DATA_ULTIMA_COMPRA"].apply(br_date)
cli_v_show["DATA_ULTIMO_VENCTO"] = cli_v_show["DATA_ULTIMO_VENCTO"].apply(br_date)

st.dataframe(cli_v_show, use_container_width=True, hide_index=True)

st.markdown("---")

# =========================
# 4) DETALHAMENTO DE TÍTULOS DO CLIENTE
# =========================
st.subheader("4) Detalhamento • Títulos em aberto do cliente")

base_cliente = df_v if vendedor_sel != "(todos)" else df_emp

rank_clientes = (
    base_cliente.groupby("CLIENTE")["VALOR"].sum()
    .reset_index()
    .sort_values("VALOR", ascending=False)
)

clientes_opts = rank_clientes["CLIENTE"].tolist()
cliente_sel = st.selectbox("Selecione o cliente", options=clientes_opts if clientes_opts else ["(sem clientes)"])
st.session_state["cliente_sel"] = cliente_sel if cliente_sel != "(sem clientes)" else None

if st.session_state["cliente_sel"]:
    det = base_cliente[base_cliente["CLIENTE"] == st.session_state["cliente_sel"]].copy()

    det_show = det[["DTA_CAD_DT", "DUPLICATA", "VENCTO_DT", "VALOR", "EMPRESA", "VENDEDOR", "HISTORICO"]].copy()
    det_show = det_show.rename(columns={
        "DTA_CAD_DT": "DATA_CADASTRO",
        "VENCTO_DT": "DATA_VENCIMENTO",
        "VALOR": "VALOR_ABERTO",
        "EMPRESA": "EMPRESA",
        "VENDEDOR": "VENDEDOR",
        "HISTORICO": "HISTORICO",
    })

    # Ordenação padrão (maior valor, vencimento mais antigo)
    det_show = det_show.sort_values(["VALOR_ABERTO", "DATA_VENCIMENTO"], ascending=[False, True])

    det_out = det_show.copy()
    det_out["DATA_CADASTRO"] = det_out["DATA_CADASTRO"].apply(br_date)
    det_out["DATA_VENCIMENTO"] = det_out["DATA_VENCIMENTO"].apply(br_date)
    det_out["VALOR_ABERTO"] = det_out["VALOR_ABERTO"].apply(brl_money)

    st.dataframe(det_out, use_container_width=True, hide_index=True)

    st.download_button(
        "Baixar títulos do cliente (CSV)",
        data=det_show.to_csv(index=False, sep=";").encode("utf-8"),
        file_name="titulos_cliente.csv",
        mime="text/csv",
    )

with st.expander("Diagnóstico (base filtrada)"):
    st.write("Linhas (filtradas):", len(df_f))
    st.write("Empresa selecionada:", empresa_sel)
    st.write("Vendedor selecionado:", vendedor_sel)
    st.write("Clientes no recorte:", base_cliente["CLIENTE"].nunique() if len(base_cliente) else 0)
    st.write("Período (VENCTO):", br_date(df_f["VENCTO_DT"].min()), "→", br_date(df_f["VENCTO_DT"].max()))
