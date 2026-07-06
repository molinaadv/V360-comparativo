import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
from datetime import date

st.set_page_config(
    page_title="V360 Relatórios | Analítico e Comparativo",
    page_icon="📊",
    layout="wide",
)

# =========================
# ESTILO V360
# =========================
st.markdown(
    """
    <style>
    .main { background-color: #f5f7fb; }
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
    .v360-title {
        font-size: 34px;
        font-weight: 800;
        color: #0B2E4A;
        margin-bottom: 0px;
    }
    .v360-subtitle {
        font-size: 15px;
        color: #5b6b7c;
        margin-bottom: 20px;
    }
    div[data-testid="stMetric"] {
        background: white;
        padding: 18px;
        border-radius: 18px;
        border: 1px solid #e7edf5;
        box-shadow: 0 8px 22px rgba(15, 45, 75, 0.06);
    }
    div[data-testid="stMetric"] label {
        color: #5b6b7c !important;
        font-size: 14px !important;
    }
    div[data-testid="stMetric"] div {
        color: #0B2E4A !important;
        font-weight: 800 !important;
    }
    .section-card {
        background: white;
        border: 1px solid #e7edf5;
        border-radius: 18px;
        padding: 18px 20px;
        box-shadow: 0 8px 22px rgba(15, 45, 75, 0.06);
        margin-bottom: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="v360-title">V360 Relatórios</div>', unsafe_allow_html=True)
st.markdown('<div class="v360-subtitle">Relatório analítico com filtro por subtipo, data de cadastro, data de cumprimento e comparativo mensal.</div>', unsafe_allow_html=True)

# =========================
# FUNÇÕES
# =========================
COL_DATA_CADASTRO = "Data/hora de cadastro"
COL_DATA_CUMPRIDO = "Data/hora conclusão efetiva"
COL_ESCRITORIO = "Escritório responsável"
COL_SUBTIPO = "Subtipo"
COL_STATUS = "Status"
COL_RESPONSAVEL = "Envolvidos / Nome"
COL_USUARIO_CADASTROU = "Usuário que cadastrou"
COL_CUMPRIDO_POR = "Cumprido por"
COL_AREA = "Vínculos com serviço / Área da ação"
COL_TIPO = "Vínculos com serviço / Tipo"
COL_CLIENTE = "Vínculos com serviço / Cliente principal / Nome/razão social"
COL_ID = "Id"

@st.cache_data(show_spinner=False)
def carregar_planilha(arquivo) -> pd.DataFrame:
    df = pd.read_excel(arquivo, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    for col in [COL_DATA_CADASTRO, COL_DATA_CUMPRIDO]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in [COL_ESCRITORIO, COL_SUBTIPO, COL_STATUS, COL_RESPONSAVEL, COL_USUARIO_CADASTROU, COL_CUMPRIDO_POR, COL_AREA, COL_TIPO, COL_CLIENTE]:
        if col in df.columns:
            df[col] = df[col].fillna("Não informado").astype(str).str.strip()

    if COL_ID in df.columns:
        df = df.drop_duplicates(subset=[COL_ID], keep="first")
    else:
        df = df.drop_duplicates()

    return df


def lista_opcoes(df: pd.DataFrame, coluna: str):
    if coluna not in df.columns:
        return []
    valores = sorted([v for v in df[coluna].dropna().astype(str).unique() if v and v != "nan"])
    return ["Todos"] + valores


def filtrar_multiselect(df: pd.DataFrame, coluna: str, selecionados):
    if coluna not in df.columns or not selecionados or "Todos" in selecionados:
        return df
    return df[df[coluna].isin(selecionados)]


def normalizar_data(d):
    if pd.isna(d):
        return None
    return pd.to_datetime(d).date()


def calcular_variacao(valor_atual, valor_anterior):
    if valor_anterior in [0, None] or pd.isna(valor_anterior):
        return None
    return ((valor_atual - valor_anterior) / valor_anterior) * 100


def formatar_delta(v):
    if v is None or pd.isna(v):
        return None
    sinal = "+" if v >= 0 else ""
    return f"{sinal}{v:.1f}%"

# =========================
# BASE DE DADOS AUTOMÁTICA
# =========================
BASE_DIR = Path(__file__).parent
PASTA_DADOS = BASE_DIR / "dados"

ARQUIVO_CADASTRO = PASTA_DADOS / "DATA CADASTRO.xlsx"
ARQUIVO_CUMPRIDO = PASTA_DADOS / "DATA CUMPRIDO.xlsx"

st.markdown("### 📁 Base de dados")
st.caption("Lendo automaticamente os arquivos da pasta `dados/` no GitHub.")

arquivos_faltando = []
if not ARQUIVO_CADASTRO.exists():
    arquivos_faltando.append(str(ARQUIVO_CADASTRO))
if not ARQUIVO_CUMPRIDO.exists():
    arquivos_faltando.append(str(ARQUIVO_CUMPRIDO))

if arquivos_faltando:
    st.error("Não encontrei os arquivos-base na pasta dados.")
    st.write("Arquivos esperados:")
    st.code("dados/DATA CADASTRO.xlsx\ndados/DATA CUMPRIDO.xlsx")
    st.write("Arquivos faltando:")
    st.code("\n".join(arquivos_faltando))
    st.stop()

try:
    df_cadastro = carregar_planilha(ARQUIVO_CADASTRO)
    df_cadastro["Origem da planilha"] = "Data Cadastro"

    df_cumprido = carregar_planilha(ARQUIVO_CUMPRIDO)
    df_cumprido["Origem da planilha"] = "Data Cumprido"

except Exception as e:
    st.error(f"Erro ao carregar os arquivos da pasta dados: {e}")
    st.info("Confira se os arquivos foram enviados como Excel verdadeiro (.xlsx), e não apenas renomeados. Também confirme se o requirements.txt possui openpyxl.")
    st.write("Arquivos encontrados:")
    try:
        st.code(f"{ARQUIVO_CADASTRO.name} - {ARQUIVO_CADASTRO.stat().st_size:,} bytes\n{ARQUIVO_CUMPRIDO.name} - {ARQUIVO_CUMPRIDO.stat().st_size:,} bytes")
    except Exception:
        pass
    st.stop()

df_base = pd.concat([df_cadastro, df_cumprido], ignore_index=True)

if COL_ID in df_base.columns:
    df_base = df_base.drop_duplicates(subset=[COL_ID], keep="first")
else:
    df_base = df_base.drop_duplicates()

st.success(
    f"Base carregada automaticamente: "
    f"{len(df_cadastro):,} registros de cadastro + "
    f"{len(df_cumprido):,} registros de cumprido.".replace(",", ".")
)

# =========================
# FILTROS
# =========================
st.markdown("### 🔎 Filtros")

colf1, colf2, colf3, colf4 = st.columns([1.2, 1, 1, 1])

with colf1:
    tipo_data = st.radio(
        "Usar qual data no relatório?",
        ["Data de Cadastro", "Data de Cumprimento"],
        horizontal=True,
    )

coluna_data = COL_DATA_CADASTRO if tipo_data == "Data de Cadastro" else COL_DATA_CUMPRIDO

datas_validas = df_base[coluna_data].dropna() if coluna_data in df_base.columns else pd.Series(dtype="datetime64[ns]")
if datas_validas.empty:
    st.warning(f"Não existem datas válidas na coluna: {coluna_data}")
    st.stop()

min_data = datas_validas.min().date()
max_data = datas_validas.max().date()

with colf2:
    data_inicial = st.date_input("Data inicial", value=min_data, min_value=min_data, max_value=max_data)
with colf3:
    data_final = st.date_input("Data final", value=max_data, min_value=min_data, max_value=max_data)
with colf4:
    ano_comparativo = st.selectbox(
        "Ano comparativo",
        sorted(datas_validas.dt.year.dropna().astype(int).unique(), reverse=True),
    )

colf5, colf6, colf7, colf8 = st.columns(4)
with colf5:
    subtipos = st.multiselect("Subtipo", lista_opcoes(df_base, COL_SUBTIPO), default=["Todos"])
with colf6:
    escritorios = st.multiselect("Escritório", lista_opcoes(df_base, COL_ESCRITORIO), default=["Todos"])
with colf7:
    responsaveis = st.multiselect("Responsável / Envolvido", lista_opcoes(df_base, COL_RESPONSAVEL), default=["Todos"])
with colf8:
    status = st.multiselect("Status", lista_opcoes(df_base, COL_STATUS), default=["Todos"])

colf9, colf10, colf11 = st.columns(3)
with colf9:
    tipos_servico = st.multiselect("Tipo da ação / serviço", lista_opcoes(df_base, COL_TIPO), default=["Todos"])
with colf10:
    areas = st.multiselect("Área", lista_opcoes(df_base, COL_AREA), default=["Todos"])
with colf11:
    agrupar_por = st.selectbox(
        "Agrupar tabela por",
        ["Escritório responsável", "Subtipo", "Envolvidos / Nome", "Status", "Vínculos com serviço / Tipo", "Vínculos com serviço / Área da ação", "Cumprido por", "Usuário que cadastrou"],
    )

# =========================
# APLICA FILTROS
# =========================
df = df_base.copy()
df = df[df[coluna_data].notna()]
df = df[(df[coluna_data].dt.date >= data_inicial) & (df[coluna_data].dt.date <= data_final)]
df = filtrar_multiselect(df, COL_SUBTIPO, subtipos)
df = filtrar_multiselect(df, COL_ESCRITORIO, escritorios)
df = filtrar_multiselect(df, COL_RESPONSAVEL, responsaveis)
df = filtrar_multiselect(df, COL_STATUS, status)
df = filtrar_multiselect(df, COL_TIPO, tipos_servico)
df = filtrar_multiselect(df, COL_AREA, areas)

df["Ano"] = df[coluna_data].dt.year
df["Mês número"] = df[coluna_data].dt.month
df["Mês"] = df[coluna_data].dt.strftime("%m/%Y")
df["Mês nome"] = df[coluna_data].dt.strftime("%b")

# =========================
# CARDS
# =========================
st.markdown("### 📌 Resumo executivo")

ultimo_mes = df[df["Ano"] == ano_comparativo].groupby("Mês número").size().sort_index()
valor_atual_mes = int(ultimo_mes.iloc[-1]) if len(ultimo_mes) > 0 else 0
valor_mes_anterior = int(ultimo_mes.iloc[-2]) if len(ultimo_mes) > 1 else None
variacao_mes = calcular_variacao(valor_atual_mes, valor_mes_anterior) if valor_mes_anterior is not None else None

total = len(df)
qtd_cumpridos = int(df[COL_STATUS].str.contains("Cumprido", case=False, na=False).sum()) if COL_STATUS in df.columns else 0
qtd_pendentes = int(df[COL_STATUS].str.contains("Pendente|Não Cumprido|Nao Cumprido", case=False, na=False).sum()) if COL_STATUS in df.columns else 0
qtd_subtipos = df[COL_SUBTIPO].nunique() if COL_SUBTIPO in df.columns else 0
qtd_escritorios = df[COL_ESCRITORIO].nunique() if COL_ESCRITORIO in df.columns else 0

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total filtrado", f"{total:,}".replace(",", "."))
c2.metric("Cumpridos", f"{qtd_cumpridos:,}".replace(",", "."))
c3.metric("Pendentes", f"{qtd_pendentes:,}".replace(",", "."))
c4.metric("Subtipos", f"{qtd_subtipos}")
c5.metric("Último mês do ano", f"{valor_atual_mes:,}".replace(",", "."), delta=formatar_delta(variacao_mes))

if df.empty:
    st.warning("Nenhum registro encontrado com os filtros selecionados.")
    st.stop()

# =========================
# COMPARATIVO MENSAL
# =========================
st.markdown("### 📈 Comparativo mensal")

aba1, aba2, aba3, aba4 = st.tabs([
    "Evolução do ano",
    "Ano x Ano",
    "Comparativo por grupo",
    "Analítico",
])

with aba1:
    df_ano = df[df["Ano"] == ano_comparativo].copy()
    mensal = (
        df_ano.groupby(["Mês número", "Mês"], as_index=False)
        .size()
        .rename(columns={"size": "Quantidade"})
        .sort_values("Mês número")
    )

    colg1, colg2 = st.columns([2, 1])
    with colg1:
        fig = px.line(
            mensal,
            x="Mês",
            y="Quantidade",
            markers=True,
            title=f"Evolução mensal - {tipo_data} - {ano_comparativo}",
        )
        fig.update_layout(height=430, margin=dict(l=20, r=20, t=60, b=20))
        st.plotly_chart(fig, use_container_width=True)
    with colg2:
        st.markdown("#### Meses")
        st.dataframe(mensal[["Mês", "Quantidade"]], use_container_width=True, hide_index=True)

with aba2:
    anos_disponiveis = sorted(df["Ano"].dropna().astype(int).unique())
    anos_sel = st.multiselect("Escolha os anos para comparar", anos_disponiveis, default=anos_disponiveis[-2:] if len(anos_disponiveis) >= 2 else anos_disponiveis)
    df_anos = df[df["Ano"].isin(anos_sel)].copy()
    comp_ano = (
        df_anos.groupby(["Ano", "Mês número"], as_index=False)
        .size()
        .rename(columns={"size": "Quantidade"})
        .sort_values(["Ano", "Mês número"])
    )
    meses_pt = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
    comp_ano["Mês"] = comp_ano["Mês número"].map(meses_pt)

    fig2 = px.line(
        comp_ano,
        x="Mês",
        y="Quantidade",
        color="Ano",
        markers=True,
        title="Comparativo Ano x Ano",
        category_orders={"Mês": list(meses_pt.values())},
    )
    fig2.update_layout(height=460, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(fig2, use_container_width=True)

    tabela_ano = comp_ano.pivot_table(index="Mês", columns="Ano", values="Quantidade", aggfunc="sum", fill_value=0).reindex(list(meses_pt.values())).reset_index()
    st.dataframe(tabela_ano, use_container_width=True, hide_index=True)

with aba3:
    grupo_col = agrupar_por
    top_n = st.slider("Quantidade de grupos no gráfico", min_value=3, max_value=20, value=10)

    top_grupos = df[grupo_col].value_counts().head(top_n).index.tolist()
    df_top = df[df[grupo_col].isin(top_grupos)].copy()
    comp_grupo = (
        df_top.groupby([grupo_col, "Ano", "Mês número"], as_index=False)
        .size()
        .rename(columns={"size": "Quantidade"})
        .sort_values([grupo_col, "Ano", "Mês número"])
    )
    comp_grupo["Mês"] = comp_grupo["Mês número"].map({1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"})
    comp_grupo["Período"] = comp_grupo["Mês"] + "/" + comp_grupo["Ano"].astype(str)

    fig3 = px.bar(
        comp_grupo,
        x="Período",
        y="Quantidade",
        color=grupo_col,
        title=f"Comparativo mensal por {grupo_col}",
        barmode="group",
    )
    fig3.update_layout(height=520, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(fig3, use_container_width=True)

    tabela_grupo = comp_grupo.pivot_table(index=grupo_col, columns="Período", values="Quantidade", aggfunc="sum", fill_value=0)
    tabela_grupo["Total"] = tabela_grupo.sum(axis=1)
    tabela_grupo = tabela_grupo.sort_values("Total", ascending=False).reset_index()
    st.dataframe(tabela_grupo, use_container_width=True, hide_index=True)

with aba4:
    st.markdown("#### Tabela agrupada")
    resumo = df.groupby(agrupar_por, as_index=False).size().rename(columns={"size": "Quantidade"}).sort_values("Quantidade", ascending=False)
    st.dataframe(resumo, use_container_width=True, hide_index=True)

    st.markdown("#### Base detalhada")
    colunas_exibir = [
        COL_DATA_CADASTRO,
        COL_DATA_CUMPRIDO,
        COL_ESCRITORIO,
        COL_SUBTIPO,
        COL_STATUS,
        COL_RESPONSAVEL,
        COL_USUARIO_CADASTROU,
        COL_CUMPRIDO_POR,
        COL_AREA,
        COL_TIPO,
        COL_CLIENTE,
        COL_ID,
        "Origem da planilha",
    ]
    colunas_exibir = [c for c in colunas_exibir if c in df.columns]
    st.dataframe(df[colunas_exibir].sort_values(coluna_data, ascending=False), use_container_width=True, hide_index=True)

    csv = df[colunas_exibir].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Baixar resultado filtrado em CSV",
        data=csv,
        file_name="v360_relatorio_filtrado.csv",
        mime="text/csv",
    )

# =========================
# INSIGHTS AUTOMÁTICOS
# =========================
st.markdown("### 💡 Insights automáticos")

melhor_mes = None
pior_mes = None
if not df.empty:
    mensal_total = df.groupby(["Ano", "Mês número"]).size().reset_index(name="Quantidade")
    if not mensal_total.empty:
        meses_pt = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
        melhor = mensal_total.sort_values("Quantidade", ascending=False).iloc[0]
        pior = mensal_total.sort_values("Quantidade", ascending=True).iloc[0]
        melhor_mes = f"{meses_pt[int(melhor['Mês número'])]}/{int(melhor['Ano'])} ({int(melhor['Quantidade'])})"
        pior_mes = f"{meses_pt[int(pior['Mês número'])]}/{int(pior['Ano'])} ({int(pior['Quantidade'])})"

maior_subtipo = df[COL_SUBTIPO].value_counts().idxmax() if COL_SUBTIPO in df.columns and not df.empty else "-"
maior_escritorio = df[COL_ESCRITORIO].value_counts().idxmax() if COL_ESCRITORIO in df.columns and not df.empty else "-"
maior_responsavel = df[COL_RESPONSAVEL].value_counts().idxmax() if COL_RESPONSAVEL in df.columns and not df.empty else "-"

st.markdown(
    f"""
    <div class="section-card">
    <b>Melhor mês:</b> {melhor_mes or '-'}<br>
    <b>Pior mês:</b> {pior_mes or '-'}<br>
    <b>Subtipo com maior volume:</b> {maior_subtipo}<br>
    <b>Escritório com maior volume:</b> {maior_escritorio}<br>
    <b>Responsável/envolvido com maior volume:</b> {maior_responsavel}
    </div>
    """,
    unsafe_allow_html=True,
)
