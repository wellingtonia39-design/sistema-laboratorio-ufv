import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import plotly.express as px
from docx import Document
import io
import zipfile

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🪵")

# --- NOME DA PLANILHA NO GOOGLE ---
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- MAPEAMENTO: COLUNA EXCEL -> TAG NO WORD ---
# Ajuste aqui se o nome no Word for diferente
DE_PARA_WORD = {
    "Código UFV": "«Código_UFV»",
    "Data de entrada": "«Data_de_entrada»",
    "Fim da análise": "«Fim_da_análise»",
    "Nome do Cliente ": "«Nome_do_Cliente_»", # Note o espaço no final se houver no excel
    "Cidade": "«Cidade»",
    "Estado": "«Estado»",
    "E-mail": "«Email»",
    "Indentificação de Amostra do cliente": "«Indentificação_de_Amostra_do_cliente»",
    "Madeira": "«Madeira»",
    "Produto utilizado": "«Produto_utilizado»",
    "Aplicação": "«Aplicação»",
    "Norma ABNT": "«Norma_ABNT»",
    "Retenção": "«Retenção»",
    # Mapeamento dos resultados químicos
    "Retenção Cromo (Kg/m³)": "«Retenção_Cromo_Kgm»",
    "Balanço Cromo %": "«Balanço_Cromo_»",
    "Retenção Cobre (Kg/m³)": "«Retenção_Cobre_Kgm»",
    "Balanço Cobre %": "«Balanço_Cobre_»",
    "Retenção Arsênio (Kg/m³)": "«Retenção_Arsênio_Kgm»",
    "Balanço Arsênio %": "«Balanço_Arsênio_»",
    "Balanço Total": "«Balanço_Total_»",
    # Mapeamento de Penetração
    "Grau penetração": "«Grau_penetração»",
    "Descrição Grau ": "«Descrição_Grau_»",
    "Descrição Penetração ": "«Descrição_Penetração_»"
}

# --- FUNÇÕES AUXILIARES ---
def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sh = client.open(NOME_PLANILHA_GOOGLE)
        return sh
    except Exception as e:
        st.error(f"Erro ao conectar no Google: {e}")
        return None

def carregar_dados(aba_nome):
    sh = conectar_google_sheets()
    if sh:
        try:
            ws = sh.worksheet(aba_nome)
            dados = ws.get_all_records()
            return pd.DataFrame(dados)
        except gspread.exceptions.WorksheetNotFound:
            sh.add_worksheet(title=aba_nome, rows=100, cols=20)
            return pd.DataFrame()
        except Exception as e:
            st.error(f"Erro ao ler aba {aba_nome}: {e}")
            return pd.DataFrame()
    return pd.DataFrame()

def salvar_dados(df, aba_nome):
    sh = conectar_google_sheets()
    if sh:
        try:
            ws = sh.worksheet(aba_nome)
            ws.clear()
            # Remove a coluna temporária de seleção antes de salvar
            if "Selecionar" in df.columns:
                df_salvar = df.drop(columns=["Selecionar"])
            else:
                df_salvar = df
            
            lista_dados = [df_salvar.columns.values.tolist()] + df_salvar.values.tolist()
            ws.update(lista_dados)
            st.toast(f"Dados de {aba_nome} salvos com sucesso!", icon="✅")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

# --- FUNÇÃO GERADORA DE RELATÓRIO WORD ---
def preencher_modelo_word(modelo_upload, dados_linha):
    doc = Document(modelo_upload)
    
    # Função interna para substituir texto em parágrafos
    def substituir_no_paragrafo(paragrafo, de, para):
        if de in paragrafo.text:
            # Substituição simples (pode perder formatação parcial da palavra, mas funciona)
            paragrafo.text = paragrafo.text.replace(de, str(para))

    # Itera sobre todas as chaves do dicionário DE_PARA
    for coluna_excel, tag_word in DE_PARA_WORD.items():
        valor = dados_linha.get(coluna_excel, "")
        
        # 1. Substituir nos parágrafos normais
        for p in doc.paragraphs:
            substituir_no_paragrafo(p, tag_word, valor)
            
        # 2. Substituir dentro de tabelas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        substituir_no_paragrafo(p, tag_word, valor)
    
    # Salva em memória
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- INTERFACE PRINCIPAL ---
st.title("🌲 UFV - Controle de Qualidade")

# Menu Lateral
menu = st.sidebar.radio("Módulo:", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard"])
st.sidebar.divider()

# Upload do Modelo (Fica na barra lateral para não ocupar espaço)
st.sidebar.markdown("### 📄 Modelo de Relatório")
arquivo_modelo = st.sidebar.file_uploader("Carregar .docx", type=["docx"])

# ==================================================
# MÓDULO 1: MADEIRA TRATADA
# ==================================================
if menu == "🪵 Madeira Tratada":
    st.header("Análise de Madeira Tratada")
    
    df_madeira = carregar_dados("Madeira")
    
    if not df_madeira.empty:
        # Adiciona coluna de Checkbox para seleção (se não existir)
        if "Selecionar" not in df_madeira.columns:
            df_madeira.insert(0, "Selecionar", False)

        # --- EDITOR DE TABELA ---
        st.caption("Selecione as amostras na primeira coluna para gerar relatório.")
        
        df_editado = st.data_editor(
            df_madeira,
            num_rows="dynamic",
            use_container_width=True,
            height=400,
            key="editor_madeira",
            column_config={
                "Selecionar": st.column_config.CheckboxColumn(
                    "Gerar Relatório?",
                    help="Marque para baixar o Word desta amostra",
                    default=False,
                )
            }
        )
        
        # --- ÁREA DE AÇÃO ---
        col_btn1, col_btn2 = st.columns([1, 1])
        
        # Botão Salvar
        with col_btn1:
            if st.button("💾 SALVAR DADOS", type="primary"):
                salvar_dados(df_editado, "Madeira")
                st.rerun()

        # Botão Gerar Relatório
        with col_btn2:
            amostras_selecionadas = df_editado[df_editado["Selecionar"] == True]
            
            if not amostras_selecionadas.empty:
                st.markdown(f"**{len(amostras_selecionadas)} amostras selecionadas.**")
                
                if arquivo_modelo:
                    if st.button("📄 GERAR RELATÓRIOS WORD"):
                        with st.spinner("Gerando documentos..."):
                            
                            # Caso 1: Apenas uma amostra
                            if len(amostras_selecionadas) == 1:
                                linha = amostras_selecionadas.iloc[0]
                                bio_word = preencher_modelo_word(arquivo_modelo, linha)
                                nome_arquivo = f"Relatorio_{linha.get('Código UFV', 'amostra')}.docx"
                                
                                st.download_button(
                                    label="⬇️ Baixar DOCX",
                                    data=bio_word,
                                    file_name=nome_arquivo,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                            
                            # Caso 2: Múltiplas amostras (Gera ZIP)
                            else:
                                zip_buffer = io.BytesIO()
                                with zipfile.ZipFile(zip_buffer, "w") as zf:
                                    for idx, linha in amostras_selecionadas.iterrows():
                                        bio_word = preencher_modelo_word(arquivo_modelo, linha)
                                        nome_arquivo = f"Relatorio_{linha.get('Código UFV', f'amostra_{idx}')}.docx"
                                        zf.writestr(nome_arquivo, bio_word.getvalue())
                                
                                zip_buffer.seek(0)
                                st.download_button(
                                    label="⬇️ Baixar Todos (ZIP)",
                                    data=zip_buffer,
                                    file_name="Relatorios_UFV.zip",
                                    mime="application/zip"
                                )
                else:
                    st.warning("⚠️ Por favor, faça upload do arquivo .docx do Modelo na barra lateral esquerda.")
            else:
                st.info("Marque a caixinha 'Gerar Relatório?' nas linhas que deseja imprimir.")

# ==================================================
# MÓDULO 2: SOLUÇÃO (Mantido Simples)
# ==================================================
elif menu == "⚗️ Solução Preservativa":
    st.header("Análise de Solução")
    df_solucao = carregar_dados("Solucao")
    
    if not df_solucao.empty:
        df_editado_sol = st.data_editor(df_solucao, num_rows="dynamic", use_container_width=True)
        if st.button("💾 SALVAR DADOS SOLUÇÃO"):
            salvar_dados(df_editado_sol, "Solucao")
            st.rerun()

# ==================================================
# MÓDULO 3: DASHBOARD (Mantido)
# ==================================================
elif menu == "📊 Dashboard":
    st.header("Dashboard Gerencial")
    df_m = carregar_dados("Madeira")
    if not df_m.empty and 'Nome do Cliente ' in df_m.columns:
        contagem = df_m['Nome do Cliente '].value_counts().reset_index()
        contagem.columns = ['Cliente', 'Quantidade']
        st.plotly_chart(px.bar(contagem, x='Cliente', y='Quantidade', title="Análises por Cliente"))
