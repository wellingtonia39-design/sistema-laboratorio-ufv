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

# --- MAPEAMENTO ATUALIZADO (CORREÇÃO DE NOMES) ---
# Esquerda: Nome exato na Coluna do Excel/Google
# Direita: A Tag que está escrita no arquivo Word
DE_PARA_WORD = {
    "Código UFV": "«Código_UFV»",
    "Data de entrada": "«Data_de_entrada»",
    "Fim da análise": "«Fim_da_análise»",
    "Data de Registro": "«Data_de_Emissão»", # Ajuste se necessário
    "Nome do Cliente ": "«Nome_do_Cliente_»", 
    "Cidade": "«Cidade»",
    "Estado": "«Estado»",
    "E-mail": "«Email»",
    "Indentificação de Amostra do cliente": "«Indentificação_de_Amostra_do_cliente»",
    "Madeira": "«Madeira»",
    "Produto utilizado": "«Produto_utilizado»",
    "Aplicação": "«Aplicação»",
    "Norma ABNT": "«Norma_ABNT»",
    
    # --- DADOS QUÍMICOS (Vão passar pela formatação de vírgula) ---
    "Retenção": "«Retenção»",
    "Retenção Cromo (Kg/m³)": "«Retenção_Cromo_Kgm»",
    "Balanço Cromo (%)": "«Balanço_Cromo_»", # Ajustado conforme seu PDF
    "Retenção Cobre (Kg/m³)": "«Retenção_Cobre_Kgm»",
    "Balanço Cobre (%)": "«Balanço_Cobre_»",
    "Retenção Arsênio (Kg/m³)": "«Retenção_Arsênio_Kgm»",
    "Balanço Arsênio (%)": "«Balanço_Arsênio_»",
    "Soma Concentração (%)": "« Retençãoconcentração »", # Corrigido conforme erro no DOCX
    "Balanço Total (%)": "«Balanço_Total_»",
    
    # --- PENETRAÇÃO ---
    "Grau de penetração": "«Grau_penetração»",
    "Descrição Grau ": "«Descrição_Grau_»",
    "Descrição Penetração ": "«Descrição_Penetração_»",
    
    # --- OBSERVAÇÕES ---
    "Observação: Analista de Controle de Qualidade": "«Observação»" # Nome longo corrigido
}

# Lista de campos que devem ser formatados como número (0,00)
CAMPOS_NUMERICOS = [
    "Retenção", "Retenção Cromo (Kg/m³)", "Balanço Cromo (%)",
    "Retenção Cobre (Kg/m³)", "Balanço Cobre (%)",
    "Retenção Arsênio (Kg/m³)", "Balanço Arsênio (%)",
    "Soma Concentração (%)", "Balanço Total (%)"
]

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
            if "Selecionar" in df.columns:
                df_salvar = df.drop(columns=["Selecionar"])
            else:
                df_salvar = df
            lista_dados = [df_salvar.columns.values.tolist()] + df_salvar.values.tolist()
            ws.update(lista_dados)
            st.toast(f"Dados de {aba_nome} salvos com sucesso!", icon="✅")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

# --- FUNÇÃO DE FORMATAÇÃO BRASILEIRA ---
def formatar_numero_br(valor):
    """Converte 6.5 para '6,50' e mantém texto se não for número"""
    try:
        if isinstance(valor, str):
            valor = valor.replace(",", ".") # Garante que string vira float
        float_val = float(valor)
        # Formata com 2 casas decimais e troca ponto por vírgula
        return "{:,.2f}".format(float_val).replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(valor)

# --- GERADOR WORD ---
def preencher_modelo_word(modelo_upload, dados_linha):
    doc = Document(modelo_upload)
    
    def substituir_no_paragrafo(paragrafo, de, para):
        if de in paragrafo.text:
            # Preserva formatação usando 'runs' se possível, senão substitui direto
            if len(paragrafo.runs) > 0 and de in paragrafo.runs[0].text:
                 paragrafo.runs[0].text = paragrafo.runs[0].text.replace(de, str(para))
            else:
                 paragrafo.text = paragrafo.text.replace(de, str(para))

    for coluna_excel, tag_word in DE_PARA_WORD.items():
        valor_bruto = dados_linha.get(coluna_excel, "")
        
        # Aplica formatação de número se for um campo numérico
        if coluna_excel in CAMPOS_NUMERICOS:
            valor_final = formatar_numero_br(valor_bruto)
        else:
            valor_final = str(valor_bruto)

        # Substituição no documento
        for p in doc.paragraphs:
            substituir_no_paragrafo(p, tag_word, valor_final)
            
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        substituir_no_paragrafo(p, tag_word, valor_final)
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- INTERFACE PRINCIPAL ---
st.title("🌲 UFV - Controle de Qualidade")

menu = st.sidebar.radio("Módulo:", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard"])
st.sidebar.divider()
st.sidebar.markdown("### 📄 Modelo de Relatório")
arquivo_modelo = st.sidebar.file_uploader("Carregar .docx", type=["docx"])

if menu == "🪵 Madeira Tratada":
    st.header("Análise de Madeira Tratada")
    df_madeira = carregar_dados("Madeira")
    
    if not df_madeira.empty:
        if "Selecionar" not in df_madeira.columns:
            df_madeira.insert(0, "Selecionar", False)

        df_editado = st.data_editor(
            df_madeira,
            num_rows="dynamic",
            use_container_width=True,
            height=400,
            column_config={
                "Selecionar": st.column_config.CheckboxColumn("Relatório?", width="small")
            }
        )
        
        c1, c2 = st.columns([1, 1])
        if c1.button("💾 SALVAR DADOS", type="primary"):
            salvar_dados(df_editado, "Madeira")
            st.rerun()

        if c2.button("📄 GERAR RELATÓRIOS"):
            selecionados = df_editado[df_editado["Selecionar"] == True]
            if not selecionados.empty and arquivo_modelo:
                with st.spinner("Formatando e gerando..."):
                    if len(selecionados) == 1:
                        linha = selecionados.iloc[0]
                        bio = preencher_modelo_word(arquivo_modelo, linha)
                        st.download_button("⬇️ Baixar DOCX", bio, f"Relatorio_{linha.get('Código UFV','amostra')}.docx")
                    else:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zf:
                            for idx, linha in selecionados.iterrows():
                                bio = preencher_modelo_word(arquivo_modelo, linha)
                                zf.writestr(f"Relatorio_{linha.get('Código UFV', idx)}.docx", bio.getvalue())
                        zip_buffer.seek(0)
                        st.download_button("⬇️ Baixar ZIP", zip_buffer, "Relatorios_UFV.zip", "application/zip")
            elif not arquivo_modelo:
                st.warning("⚠️ Carregue o modelo .docx na barra lateral!")
            else:
                st.info("Selecione pelo menos uma amostra.")

elif menu == "⚗️ Solução Preservativa":
    st.header("Análise de Solução")
    df_sol = carregar_dados("Solucao")
    if not df_sol.empty:
        df_ed = st.data_editor(df_sol, num_rows="dynamic", use_container_width=True)
        if st.button("💾 SALVAR SOLUÇÃO"):
            salvar_dados(df_ed, "Solucao")
            st.rerun()

elif menu == "📊 Dashboard":
    st.header("Dashboard Gerencial")
    df_m = carregar_dados("Madeira")
    if not df_m.empty and 'Nome do Cliente ' in df_m.columns:
        contagem = df_m['Nome do Cliente '].value_counts().reset_index()
        contagem.columns = ['Cliente', 'Quantidade']
        st.plotly_chart(px.bar(contagem, x='Cliente', y='Quantidade', title="Análises por Cliente"))
