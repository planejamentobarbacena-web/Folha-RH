import streamlit as st
import pandas as pd
import io

# controle reset uploads
if "reset" not in st.session_state:
    st.session_state.reset = 0


def render():

    if st.button("🔄 Nova Consulta"):
        st.session_state.reset += 1
        st.rerun()

    st.title("Análise de Folha - RH")
    st.write("Envie os arquivos da folha para análise.")

    # ============================================
    # TABELA REFERÊNCIA
    # ============================================

    st.subheader("Tabela Referência (Opcional)")

    tabela_ref_file = st.file_uploader(
        "Envie a TABELA REFERENCIA",
        type=["csv","xlsx"],
        key=f"tabela_{st.session_state.reset}"
    )

    tabela_referencia = None

    if tabela_ref_file:

        if tabela_ref_file.name.endswith(".csv"):
            tabela_referencia = pd.read_csv(tabela_ref_file, sep=None, engine="python", dtype=str)
        else:
            tabela_referencia = pd.read_excel(tabela_ref_file, dtype=str)

        tabela_referencia.columns = tabela_referencia.columns.str.strip().str.upper()

        tabela_referencia["ORGANOGRAMA"] = tabela_referencia["ORGANOGRAMA"].astype(str)
        tabela_referencia["ESTRUTURA ARQUIVO"] = tabela_referencia["ESTRUTURA ARQUIVO"].astype(str)

        st.success("Tabela referência carregada")

    # ============================================
    # PREVIDENCIA
    # ============================================

    st.subheader("Arquivos de Previdência")

    arquivos_prev = st.file_uploader(
        "Envie os arquivos PREVIDENCIA",
        type=["csv","xlsx"],
        accept_multiple_files=True,
        key=f"prev_{st.session_state.reset}"
    )

    previdencia_total = pd.DataFrame()

    if arquivos_prev:

        lista_prev = []

        for arq in arquivos_prev:

            if arq.name.endswith(".csv"):
                prev = pd.read_csv(arq, sep=";", dtype=str)
            else:
                prev = pd.read_excel(arq, dtype=str)

            # coluna A = organograma completo
            prev["codigo"] = prev.iloc[:,0].astype(str)

            prev["SECRETARIA"] = prev["codigo"].str[:2]
            prev["ORGANOGRAMA"] = prev["codigo"].str[6:10]
            prev["FONTE"] = prev["codigo"].str[-8:]

            # coluna E = Fundo (28%)
            prev["VALOR"] = (
                prev.iloc[:,4]
                .astype(str)
                .str.replace(".","",regex=False)
                .str.replace(",",".",regex=False)
                .astype(float)
            )

            prev = prev.groupby(
                ["SECRETARIA","ORGANOGRAMA","FONTE"]
            )["VALOR"].sum().reset_index()

            nome = arq.name.upper()

            if "EFETIVO" in nome:
                prev["TIPO"] = "INSS"
            else:
                prev["TIPO"] = "SIMPAS"

            lista_prev.append(prev)

        previdencia_total = pd.concat(lista_prev)

    # ============================================
    # ARQUIVOS DA FOLHA
    # ============================================

    arquivos = st.file_uploader(
        "Envie os arquivos da folha",
        type=["csv","xlsx"],
        accept_multiple_files=True,
        key=f"folha_{st.session_state.reset}"
    )

    if not arquivos:
        return

    if st.button("Executar análise"):

        base_final = []

        for arquivo in arquivos:

            nome_servidor = arquivo.name.split(".")[0].upper()

            if arquivo.name.endswith(".csv"):
                df = pd.read_csv(arquivo, sep=";", encoding="utf-8", dtype=str)
            else:
                df = pd.read_excel(arquivo, dtype=str)

            df.columns = df.columns.str.strip()

            df = df[df["Tipo Evento"].isin(["VENCIMENTO","DESCONTO"])]

            df["codigo"] = df["Estrutura organizacional"].str.split(" - ").str[0]

            df["SECRETARIA"] = df["codigo"].str[:2]
            df["ORGANOGRAMA"] = df["codigo"].str[4:8]
            df["FONTE"] = df["codigo"].str[-8:]

            df["ESTRUTURA ARQUIVO"] = df["Estrutura organizacional"].str.split(" - ").str[1]

            df["Valor calculado"] = df["Valor calculado"].str.replace(" P","").str.replace(" D","")

            df["VALOR"] = (
                df["Valor calculado"]
                .str.replace(".","",regex=False)
                .str.replace(",",".",regex=False)
                .astype(float)
            )

            vencimentos = (
                df[df["Tipo Evento"]=="VENCIMENTO"]
                .groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
                .sum().reset_index()
                .rename(columns={"VALOR":"VENCIMENTOS"})
            )

            descontos = (
                df[df["Tipo Evento"]=="DESCONTO"]
                .groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
                .sum().reset_index()
                .rename(columns={"VALOR":"DESCONTOS"})
            )

            base = vencimentos.merge(
                descontos,
                on=["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"],
                how="outer"
            ).fillna(0)

            # =================================
            # IRRF
            # =================================

            irrf = df[
                (df["Tipo Evento"]=="DESCONTO") &
                (df["Evento"].str.contains("I.R.R.F",case=False,na=False))
            ]

            irrf = (
                irrf.groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
                .sum().reset_index()
                .rename(columns={"VALOR":"IRRF"})
            )

            base = base.merge(irrf, how="left").fillna(0)

            # =================================
            # PENSAO
            # =================================

            pensao = df[
                (df["Tipo Evento"]=="DESCONTO") &
                (df["Evento"].str.contains("pens",case=False,na=False))
            ]

            pensao = (
                pensao.groupby(["ESTRUTURA ARQUIVO","SECRETARIA","ORGANOGRAMA","FONTE"])["VALOR"]
                .sum().reset_index()
                .rename(columns={"VALOR":"PENSAO"})
            )

            base = base.merge(pensao, how="left").fillna(0)

            # =================================
            # PREVIDENCIA
            # =================================

            if not previdencia_total.empty:

                prev = previdencia_total.copy()

                base = base.merge(
                    prev,
                    on=["SECRETARIA","ORGANOGRAMA","FONTE"],
                    how="left"
                )

                base["PATRONAL - INSS"] = base.apply(
                    lambda x: x["VALOR"] if x.get("TIPO")=="INSS" else 0,
                    axis=1
                )

                base["PATRONAL - SIMPAS"] = base.apply(
                    lambda x: x["VALOR"] if x.get("TIPO")=="SIMPAS" else 0,
                    axis=1
                )

            else:

                base["PATRONAL - INSS"] = 0
                base["PATRONAL - SIMPAS"] = 0

            base["LIQUIDO"] = base["VENCIMENTOS"] - base["DESCONTOS"]

            base["TOTAL"] = (
                base["LIQUIDO"]
                + base["PATRONAL - INSS"]
                + base["PATRONAL - SIMPAS"]
            )

            base["SERVIDOR"] = nome_servidor

            if tabela_referencia is not None:

                base = base.merge(
                    tabela_referencia,
                    on=["ESTRUTURA ARQUIVO","ORGANOGRAMA"],
                    how="left"
                )

                base["ESTRUTURA ATUALIZADA"] = base["ESTRUTURA ATUALIZADA"].fillna(base["ESTRUTURA ARQUIVO"])

            else:

                base["ESTRUTURA ATUALIZADA"] = base["ESTRUTURA ARQUIVO"]

            base_final.append(base)

        resultado = pd.concat(base_final)

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

        def moeda(v):
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")

        view = resultado.copy()

        for col in [
            "VENCIMENTOS","DESCONTOS","LIQUIDO",
            "PATRONAL - INSS","PATRONAL - SIMPAS",
            "PENSAO","IRRF","TOTAL"
        ]:
            view[col] = view[col].apply(moeda)

        st.subheader("Resultado")
        st.dataframe(view, use_container_width=True)

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

            resultado.to_excel(writer, index=False, sheet_name="RH")

            workbook = writer.book
            formato = workbook.add_format({"num_format":"R$ #,##0.00"})

            worksheet = writer.sheets["RH"]

            for i in range(6,14):
                worksheet.set_column(i,i,18,formato)

        st.download_button(
            "Baixar planilha",
            data=output.getvalue(),
            file_name="analise_rh.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


render()
