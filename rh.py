import streamlit as st
import pandas as pd
import io

st.markdown("""
<style>

/* botão */
div.stButton > button {
    background-color: #1f4fbf !important;
    border-radius: 8px !important;
    padding: 14px 40px !important;
    min-width: 260px !important;
    height: 52px !important;
    white-space: nowrap !important;
}

/* TEXTO DO BOTÃO */
div.stButton > button span {
    color: white !important;
    font-size: 40px !important;
    font-weight: 600 !important;
}

/* hover */
div.stButton > button:hover {
    background-color: #163a8a !important;
}

</style>
""", unsafe_allow_html=True)

# controle reset
if "reset" not in st.session_state:
    st.session_state.reset = 0


def render():

    # ============================================
    # BOTÃO NOVA CONSULTA CENTRALIZADO
    # ============================================

    col1, col2, col3 = st.columns([3,2,3])

    with col2:
        if st.button("🔄 Nova Consulta", use_container_width=True):
            st.session_state.reset += 1
            st.rerun()

    st.title("Análise de Folha - RH")
    st.write("Envie um ou mais arquivos da folha para análise.")

    # ============================================
    # UPLOAD TABELA REFERENCIA
    # ============================================

    st.subheader("Tabela Referência (Opcional)")

    tabela_ref_file = st.file_uploader(
        "Envie a TABELA REFERENCIA",
        type=["csv", "xlsx"],
        key=f"tabela_ref_{st.session_state.reset}"
    )

    tabela_referencia = None

    if tabela_ref_file is not None:

        if tabela_ref_file.name.endswith(".csv"):
            tabela_referencia = pd.read_csv(
                tabela_ref_file,
                sep=None,
                engine="python",
                dtype=str
            )
        else:
            tabela_referencia = pd.read_excel(
                tabela_ref_file,
                dtype=str
            )

        tabela_referencia.columns = tabela_referencia.columns.str.strip().str.upper()

        tabela_referencia["ORGANOGRAMA"] = tabela_referencia["ORGANOGRAMA"].astype(str)
        tabela_referencia["ESTRUTURA ARQUIVO"] = tabela_referencia["ESTRUTURA ARQUIVO"].astype(str)

        st.success("Tabela referência carregada")

    # ============================================
    # UPLOAD ARQUIVOS
    # ============================================

    arquivos = st.file_uploader(
        "Envie os arquivos da folha",
        type=["csv","xlsx"],
        accept_multiple_files=True,
        key=f"arquivos_{st.session_state.reset}"
    )

    if not arquivos:
        return

    if st.button("Executar análise"):

        base_final = []

        for arquivo in arquivos:

            nome_servidor = arquivo.name.split(".")[0].upper()

            # leitura

            if arquivo.name.endswith(".csv"):
                df = pd.read_csv(
                    arquivo,
                    sep=";",
                    encoding="utf-8",
                    dtype=str
                )
            else:
                df = pd.read_excel(arquivo, dtype=str)

            df.columns = df.columns.str.strip()

            # eventos

            df = df[df["Tipo Evento"].isin(["VENCIMENTO","DESCONTO"])]

            # estrutura

            df["codigo"] = df["Estrutura organizacional"].str.split(" - ").str[0]

            df["SECRETARIA"] = df["codigo"].str[:2]
            df["ORGANOGRAMA"] = df["codigo"].str[4:8]
            df["FONTE"] = df["codigo"].str[-8:]

            df["ESTRUTURA ARQUIVO"] = df["Estrutura organizacional"].str.split(" - ").str[1]

            df["ORGANOGRAMA"] = df["ORGANOGRAMA"].astype(str)
            df["ESTRUTURA ARQUIVO"] = df["ESTRUTURA ARQUIVO"].astype(str)

            # valores

            df["Valor calculado"] = (
                df["Valor calculado"]
                .str.replace(" P","")
                .str.replace(" D","")
            )

            df["VALOR"] = (
                df["Valor calculado"]
                .str.replace(".","")
                .str.replace(",",".")
                .astype(float)
            )

            # vencimentos

            vencimentos = (
                df[df["Tipo Evento"]=="VENCIMENTO"]
                .groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
                .sum()
                .reset_index()
                .rename(columns={"VALOR":"VENCIMENTOS"})
            )

            # descontos

            descontos = (
                df[df["Tipo Evento"]=="DESCONTO"]
                .groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
                .sum()
                .reset_index()
                .rename(columns={"VALOR":"DESCONTOS"})
            )

            base = vencimentos.merge(
                descontos,
                on=["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"],
                how="outer"
            ).fillna(0)

            # cálculos

            base["LIQUIDO"] = base["VENCIMENTOS"] - base["DESCONTOS"]

            base["PATRONAL - INSS"] = 0
            base["PATRONAL - SIMPAS"] = 0
            base["PENSAO"] = 0
            base["IRRF"] = 0
            base["TOTAL"] = 0

            base["SERVIDOR"] = nome_servidor

            # estrutura atualizada

            if tabela_referencia is not None:

                base = base.merge(
                    tabela_referencia,
                    on=["ESTRUTURA ARQUIVO","ORGANOGRAMA"],
                    how="left"
                )

                if "ESTRUTURA ATUALIZADA" in base.columns:
                    base["ESTRUTURA ATUALIZADA"] = base["ESTRUTURA ATUALIZADA"].fillna(base["ESTRUTURA ARQUIVO"])
                else:
                    base["ESTRUTURA ATUALIZADA"] = base["ESTRUTURA ARQUIVO"]

            else:
                base["ESTRUTURA ATUALIZADA"] = base["ESTRUTURA ARQUIVO"]

            base_final.append(base)

        resultado = pd.concat(base_final)

        # ordem das colunas

        resultado = resultado[
            [
                "SERVIDOR",
                "ESTRUTURA ARQUIVO",
                "ESTRUTURA ATUALIZADA",
                "SECRETARIA",
                "ORGANOGRAMA",
                "FONTE",
                "VENCIMENTOS",
                "DESCONTOS",
                "LIQUIDO",
                "PATRONAL - INSS",
                "PATRONAL - SIMPAS",
                "PENSAO",
                "IRRF",
                "TOTAL"
            ]
        ]

        st.subheader("Resultado")
        st.dataframe(resultado, use_container_width=True)

        # ============================================
        # EXPORTAR EXCEL (IDENTAÇÃO CORRIGIDA)
        # ============================================

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

            resultado.to_excel(writer, index=False, sheet_name="Analise")

            workbook = writer.book
            worksheet = writer.sheets["Analise"]

            formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})

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
                idx = resultado.columns.get_loc(col)
                worksheet.set_column(idx, idx, 18, formato_moeda)

        st.download_button(
            "Baixar planilha",
            data=output.getvalue(),
            file_name="analise_rh.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


render()