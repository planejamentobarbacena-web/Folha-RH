import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Análise de Folha RH",
    layout="wide"
)

st.title("📊 Análise de Folha RH")

# =====================================================
# BOTÃO NOVA CONSULTA
# =====================================================

if st.button("🔄 Nova Consulta"):
    st.session_state.clear()
    st.rerun()

# =====================================================
# UPLOAD ARQUIVOS
# =====================================================

st.subheader("Arquivos da Folha")

arquivos = st.file_uploader(
    "Selecione os arquivos da folha",
    type=["xlsx","xls","csv"],
    accept_multiple_files=True
)

st.subheader("Tabela Referência (Opcional)")

arquivo_referencia = st.file_uploader(
    "Tabela de referência",
    type=["xlsx","xls","csv"]
)

# =====================================================
# FUNÇÃO PARA LER ARQUIVOS
# =====================================================

def ler_arquivo(file):

    if file.name.endswith(".csv"):
        df = pd.read_csv(file, sep=";", encoding="latin1")
    else:
        df = pd.read_excel(file)

    return df


# =====================================================
# PROCESSAMENTO
# =====================================================

if arquivos:

    dfs = []

    for arquivo in arquivos:

        df = ler_arquivo(arquivo)

        df["ESTRUTURA ARQUIVO"] = arquivo.name

        dfs.append(df)

    df = pd.concat(dfs, ignore_index=True)

    # garante tipo numérico
    df["VALOR"] = (
        df["VALOR"]
        .astype(str)
        .str.replace(".","",regex=False)
        .str.replace(",",".",regex=False)
    )

    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0)

    # =====================================================
    # VENCIMENTOS
    # =====================================================

    vencimentos = df[df["Tipo Evento"]=="VENCIMENTO"]

    vencimentos = (
        vencimentos
        .groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
        .sum()
        .reset_index()
        .rename(columns={"VALOR":"VENCIMENTOS"})
    )

    # =====================================================
    # DESCONTOS
    # =====================================================

    descontos = df[df["Tipo Evento"]=="DESCONTO"]

    descontos = (
        descontos
        .groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
        .sum()
        .reset_index()
        .rename(columns={"VALOR":"DESCONTOS"})
    )

    # =====================================================
    # BASE
    # =====================================================

    base = vencimentos.merge(
        descontos,
        on=["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"],
        how="outer"
    ).fillna(0)

    # =====================================================
    # IRRF
    # =====================================================

    irrf = df[
        (df["Tipo Evento"]=="DESCONTO") &
        (df["Evento"].str.contains("I.R.R.F",case=False,na=False))
    ]

    irrf = (
        irrf.groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
        .sum()
        .reset_index()
        .rename(columns={"VALOR":"IRRF"})
    )

    base = base.merge(
        irrf,
        on=["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"],
        how="left"
    ).fillna(0)

    # =====================================================
    # PENSAO
    # =====================================================

    pensao = df[
        (df["Tipo Evento"]=="DESCONTO") &
        (df["Evento"].str.contains("pens",case=False,na=False))
    ]

    pensao = (
        pensao.groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
        .sum()
        .reset_index()
        .rename(columns={"VALOR":"PENSAO"})
    )

    base = base.merge(
        pensao,
        on=["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"],
        how="left"
    ).fillna(0)

    # =====================================================
    # COLUNAS PATRONAIS (AJUSTAREMOS DEPOIS)
    # =====================================================

    base["PATRONAL - INSS"] = 0
    base["PATRONAL - SIMPAS"] = 0

    # =====================================================
    # LIQUIDO
    # =====================================================

    base["LIQUIDO"] = base["VENCIMENTOS"] - base["DESCONTOS"]

    # =====================================================
    # TOTAL
    # =====================================================

    base["TOTAL"] = (
        base["VENCIMENTOS"]
        + base["PATRONAL - INSS"]
        + base["PATRONAL - SIMPAS"]
    )

    resultado = base.copy()

    # =====================================================
    # TABELA REFERÊNCIA
    # =====================================================

    if arquivo_referencia:

        tabela_referencia = ler_arquivo(arquivo_referencia)

        tabela_referencia["ORGANOGRAMA"] = tabela_referencia["ORGANOGRAMA"].astype(str)
        resultado["ORGANOGRAMA"] = resultado["ORGANOGRAMA"].astype(str)

        resultado = resultado.merge(
            tabela_referencia,
            on=["ESTRUTURA ARQUIVO","ORGANOGRAMA"],
            how="left"
        )

    # =====================================================
    # FORMATAÇÃO
    # =====================================================

    colunas_moeda = [
        "VENCIMENTOS",
        "DESCONTOS",
        "LIQUIDO",
        "PATRONAL - INSS",
        "PATRONAL - SIMPAS",
        "PENSAO",
        "IRRF",
        "TOTAL"
    ]

    for col in colunas_moeda:
        resultado[col] = resultado[col].astype(float)

    # =====================================================
    # EXIBIR
    # =====================================================

    st.subheader("Resultado da Análise")

    st.dataframe(resultado, use_container_width=True)

    # =====================================================
    # DOWNLOAD EXCEL
    # =====================================================

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:

        resultado.to_excel(
            writer,
            index=False,
            sheet_name="Analise"
        )

    st.download_button(
        "📥 Baixar Excel",
        buffer.getvalue(),
        file_name="analise_folha.xlsx"
    )
