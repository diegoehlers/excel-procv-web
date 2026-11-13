import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Excel PROCV Web", page_icon="📊", layout="wide")
st.title("📊 Excel PROCV Automático (Web)")
st.markdown("Faça o PROCV escolhendo exatamente onde o resultado deve ser gravado.")

def normalizar(v):
    if pd.isna(v):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        try:
            int(s[:-2])
            s = s[:-2]
        except ValueError:
            pass
    return s

arquivo = st.file_uploader("Selecione a planilha do Excel (.xlsx)", type=["xlsx"])

if arquivo is not None:
    xls = pd.ExcelFile(arquivo)
    abas = xls.sheet_names
    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. Planilha Principal (Onde os dados serão inseridos)")
        # Tenta achar a aba "3021" ou "Preenchimento" automaticamente
        index_principal = 0
        for i, nome in enumerate(abas):
            if "3021" in nome or "Preenchimento" in nome:
                index_principal = i
                break
                
        aba_principal = st.selectbox("Escolha a aba principal", abas, index=index_principal)
        df_principal = pd.read_excel(arquivo, sheet_name=aba_principal)
        
        coluna_pesquisa = st.selectbox("Coluna que contém o código para busca", df_principal.columns)
        
        # AQUI ESTÁ A NOVIDADE: Campo para escolher a coluna de destino
        coluna_destino = st.selectbox("🎯 Coluna de Destino (Onde o resultado será escrito)", df_principal.columns)

    with col2:
        st.subheader("2. Base de Dados (Onde buscar a informação)")
        # Tenta achar a aba "ETAPA 2" automaticamente
        index_base = 1 if len(abas) > 1 else 0
        for i, nome in enumerate(abas):
            if "ETAPA 2" in nome or "Mantenedor" in nome:
                index_base = i
                break
                
        aba_base = st.selectbox("Escolha a aba de referência", abas, index=index_base)
        df_base = pd.read_excel(arquivo, sheet_name=aba_base)
        
        coluna_base_busca = st.selectbox("Coluna de busca na base", df_base.columns)
        coluna_base_retorno = st.selectbox("Coluna com o dado a ser obtido", df_base.columns)

    st.divider()
    st.subheader("3. Executar")
    
    st.info(f"O dado da coluna **'{coluna_base_retorno}'** (da aba Base) será copiado para a coluna **'{coluna_destino}'** (da aba Principal).")

    if st.button("🚀 Executar PROCV e Preparar Download", type="primary"):
        with st.spinner("Processando dados... (isso pode levar alguns segundos)"):
            
            # 1. Prepara a base (apenas remove duplicatas, NÃO remove vazios)
            base_normalizada = df_base.copy()
            base_normalizada["__chave__"] = base_normalizada[coluna_base_busca].apply(normalizar)
            base_normalizada = base_normalizada[base_normalizada["__chave__"] != ""]
            base_normalizada = base_normalizada.drop_duplicates(subset=["__chave__"], keep="first")
            
            # 2. Cria o dicionário de busca
            dicionario_busca = dict(zip(base_normalizada["__chave__"], base_normalizada[coluna_base_retorno]))

            # 3. Aplica na principal
            chaves_principal = df_principal[coluna_pesquisa].apply(normalizar)
            resultado_procv = chaves_principal.map(dicionario_busca)

            # 4. Substitui o conteúdo da coluna de destino escolhida
            df_principal[coluna_destino] = resultado_procv.fillna("")

            preenchidas = (df_principal[coluna_destino] != "").sum()
            total_linhas = len(df_principal)

            if preenchidas == 0:
                st.error(f"⚠️ 0 de {total_linhas} linhas foram preenchidas. Verifique os dados.")
            else:
                st.success(f"✅ PROCV executado! **{preenchidas} de {total_linhas}** linha(s) preenchida(s).")

            st.write("**Prévia das 10 primeiras linhas (veja a coluna de destino preenchida):**")
            st.dataframe(df_principal[[coluna_pesquisa, coluna_destino]].head(10))

            # 5. Exportar mantendo o formato original
            arquivo.seek(0)
            wb = openpyxl.load_workbook(arquivo)
            ws = wb[aba_principal]
            headers = [cell.value for cell in ws[1]]

            # Encontra o índice da coluna de destino
            if coluna_destino in headers:
                target_col_idx = headers.index(coluna_destino) + 1
            else:
                st.error(f"Erro: A coluna {coluna_destino} não foi encontrada no cabeçalho do Excel.")
                st.stop()

            for index, row in df_principal.iterrows():
                excel_row = index + 2
                valor = row[coluna_destino]
                if pd.isna(valor) or valor == "":
                    valor = None
                # Escreve exatamente na coluna que você escolheu como destino
                ws.cell(row=excel_row, column=target_col_idx, value=valor)

            buffer = BytesIO()
            wb.save(buffer)

            st.download_button(
                label="📥 Baixar Planilha com os Resultados",
                data=buffer.getvalue(),
                file_name="planilha_resultado_procv.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )