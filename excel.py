import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="Gerador de CSV por Aba", layout="centered")
st.title("📊 Extrator de CNPJ e NUM por Grupo")
st.markdown("Faça upload de um arquivo Excel com várias abas, e baixe um `.csv` com `CNPJ` e `NUM` por aba.")

uploaded_file = st.file_uploader("📤 Envie o arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for aba, df in xls.items():
                if df.shape[1] >= 6:
                    cnpj_col = df.iloc[1:, 1]
                    num_col = df.iloc[1:, 5]

                    cnpj_col = cnpj_col.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
                    num_col = num_col.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

                    resultado = pd.DataFrame({
                        "CNPJ": cnpj_col,
                        "NUM": num_col
                    })

                    resultado = resultado.dropna(how="all")
                    resultado = resultado[~resultado["CNPJ"].str.lower().str.contains("cnpj", na=False)]
                    resultado = resultado[~resultado["NUM"].str.lower().str.contains("num|telefone|whats", na=False)]

                    csv_buffer = io.StringIO()
                    resultado.to_csv(csv_buffer, index=False)
                    zip_file.writestr(f"{aba}.csv", csv_buffer.getvalue())

        st.success("✅ Arquivos gerados com sucesso!")

        st.download_button(
            label="📦 Baixar arquivos CSV (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="cnpj_num_por_grupo.zip",
            mime="application/zip"
        )

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
