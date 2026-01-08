import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import plotly.express as px
from docx import Document
import io
import zipfile
import os
import subprocess
from datetime import datetime
import shutil

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🌲")

# --- NOME DA PLANILHA NO GOOGLE ---
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- MAPEAMENTO (COLUNA EXCEL -> TAG WORD) ---
DE_PARA_WORD = {
    "Código UFV": "«Código_UFV»",
    "Data de entrada": "«Data_de_entrada»",
    "Fim da análise": "«Fim_da_análise»",
    "Data de Registro": "«Data_de_Emissão»",
    "Nome do Cliente": "«Nome_do_Cliente_»", 
    "Cidade": "«Cidade»",
    "Estado": "«Estado»",
    "E-mail": "«Email»",
    "Indentificação de Amostra do cliente": "«Indentificação_de_Amostra_do_cliente»",
    "Madeira": "«Madeira»",
    "Produto utilizado": "«Produto_utilizado»",
    "Aplicação": "«Aplicação»",
    "Norma ABNT": "«Norma_ABNT»",
    "Retenção": "«Retenção»",
    # Químicos
    "Retenção Cromo (Kg/m³)": "«Retenção_Cromo_Kgm»",
    "Balanço Cromo (%)": "«Balanço_Cromo_»",
    "Retenção Cobre (Kg/m³)": "«Retenção_Cobre_Kgm»",
    "Balanço Cobre (%)": "«Balanço_Cobre_»",
    "Retenção Arsênio (Kg/m³)": "«Retenção_Arsênio_Kgm»",
    "Balanço Arsênio (%)": "«Balanço_Arsênio_»",
    "Soma Concentração (%)": "« Retençãoconcentração »",
    "Balanço Total (%)": "«Balanço_Total_»",
    # Penetração
    "Grau de penetração": "«Grau_penetração»",
    "Descrição Grau": "«Descrição_Grau_»",
    "Descrição Penetração": "«Descrição_Penetração_»",
    "Observação: Analista de Controle de Qualidade": "«Observação»"
}

CAMPOS_NUMERICOS = [
    "Retenção", "Retenção Cromo (Kg/m³)", "Balanço Cromo (%)",
    "Retenção Cobre (Kg/m³)", "Balanço Cobre (%)",
    "Retenção Arsênio (Kg/m³)", "Balanço Arsênio (%)",
    "Soma Concentração (%)", "Balanço Total (%)"
]

CAMPOS_DATA = ["Data de entrada", "Fim da análise", "Data de Registro"]

# --- FUNÇÕES DE ARQUIVO E DIAGNÓSTICO ---
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
            df = pd.DataFrame(dados)
            if not df.empty: df.columns = df.columns.str.strip()
            return df
        except: return pd.DataFrame()
    return pd.DataFrame()

def salvar_dados(df, aba_nome):
    sh = conectar_google_sheets()
    if sh:
        try:
            ws = sh.worksheet(aba_nome)
            ws.clear()
            df_salvar = df.drop(columns=["Selecionar"]) if "Selecionar" in df.columns else df
            lista_dados = [df_salvar.columns.values.tolist()] + df_salvar.values.tolist()
            ws.update(lista_dados)
            st.toast(f"Dados salvos!", icon="✅")
        except Exception as e: st.error(f"Erro ao salvar: {e}")

# --- FORMATAÇÃO BRASILEIRA ---
def formatar_numero_br(valor):
    try:
        if valor == "" or valor is None: return ""
        if isinstance(valor, str): valor = valor.replace(",", ".")
        return "{:,.2f}".format(float(valor)).replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(valor)

def formatar_data_br(valor):
    if not valor: return ""
    valor_str = str(valor).strip().split(" ")[0]
    formatos = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y"]
    for fmt in formatos:
        try: return datetime.strptime(valor_str, fmt).strftime("%d/%m/%Y")
        except: continue
    return valor_str

# --- CONVERSOR PDF ---
def converter_docx_para_pdf(docx_bytes):
    try:
        with open("temp.docx", "wb") as f: f.write(docx_bytes.getvalue())
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', 'temp.docx', '--outdir', '.'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=50)
        if os.path.exists("temp.pdf"):
            with open("temp.pdf", "rb") as f: pdf_bytes = f.read()
            os.remove("temp.docx"); os.remove("temp.pdf")
            return pdf_bytes, None
        return None, "Arquivo PDF não foi gerado pelo LibreOffice."
    except Exception as e: return None, str(e)

def preencher_modelo_word(modelo_upload, dados_linha):
    doc = Document(modelo_upload)
    
    def substituir(paragrafo, de, para):
        if de in paragrafo.text:
            alterado = False
            for run in paragrafo.runs:
                if de in run.text:
                    run.text = run.text.replace(de, str(para))
                    alterado = True
            if not alterado: paragrafo.text = paragrafo.text.replace(de, str(para))

    dados_fmt = {}
    for col, tag in DE_PARA_WORD.items():
        val = dados_linha.get(col, "")
        if col in CAMPOS_NUMERICOS: dados_fmt[tag] = formatar_numero_br(val)
        elif col in CAMPOS_DATA: dados_fmt[tag] = formatar_data_br(val)
        else: dados_fmt[tag] = str(val)

    for tag, val in dados_fmt.items():
        if val is None: val = ""
        for p in doc.paragraphs: substituir(p, tag, val)
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells:
                    for p in c.paragraphs: substituir(p, tag, val)
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- INTERFACE PRINCIPAL COM DIAGNÓSTICO ---
st.title("🌲 UFV - Controle de Qualidade (Modo Diagnóstico)")

# --- BARRA LATERAL DE DIAGNÓSTICO ---
st.sidebar.header("🔧 Diagnóstico do Sistema")
pacotes_existe = os.path.exists("packages.txt")
libreoffice_instalado = shutil.which("libreoffice") is not None

if pacotes_existe:
    st.sidebar.success("✅ Arquivo packages.txt encontrado")
    with open("packages.txt", "r") as f:
        conteudo = f.read()
        st.sidebar.text_area("Conteúdo do packages.txt:", conteudo, height=68)
        if "libreoffice" in conteudo:
            st.sidebar.success("✅ Texto 'libreoffice' detectado")
        else:
            st.sidebar.error("❌ Texto 'libreoffice' NÃO encontrado dentro do arquivo!")
else:
    st.sidebar.error("❌ Arquivo packages.txt NÃO existe no GitHub!")

st.sidebar.divider()

if libreoffice_instalado:
    st.sidebar.success("✅ LibreOffice INSTALADO e PRONTO!")
else:
    st.sidebar.error("❌ LibreOffice NÃO está instalado no sistema.")
    st.sidebar.info("Solução: Corrija o packages.txt e clique em 'Reboot App'.")

st.sidebar.divider()

# --- MENU E APLICAÇÃO ---
menu = st.sidebar.radio("Módulo:", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard"])
arquivo_modelo = st.sidebar.file_uploader("Carregar Modelo (.docx)", type=["docx"])

if menu == "🪵 Madeira Tratada":
    st.header("Análise de Madeira Tratada")
    df_madeira = carregar_dados("Madeira")
    
    if not df_madeira.empty:
        if "Selecionar" not in df_madeira.columns: df_madeira.insert(0, "Selecionar", False)
        
        df_editado = st.data_editor(df_madeira, num_rows="dynamic", use_container_width=True, column_config={"Selecionar": st.column_config.CheckboxColumn("Gerar?", width="small")})
        
        c1, c2, c3 = st.columns(3)
        if c1.button("💾 SALVAR", type="primary", use_container_width=True):
            salvar_dados(df_editado, "Madeira"); st.rerun()
        
        selecionados = df_editado[df_editado["Selecionar"] == True]
        
        if c2.button("📄 BAIXAR WORD", use_container_width=True):
            if not selecionados.empty and arquivo_modelo:
                bio = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                st.download_button("⬇️ DOCX", bio, "Relatorio.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else: st.warning("Selecione 1 amostra e carregue o modelo.")

        # BOTÃO PDF
        if c3.button("📄 BAIXAR PDF", use_container_width=True):
            if not libreoffice_instalado:
                st.error("🚫 IMPOSSÍVEL GERAR PDF: O LibreOffice não foi encontrado. Veja o diagnóstico na esquerda.")
            elif selecionados.empty:
                st.warning("Selecione uma amostra na tabela.")
            elif not arquivo_modelo:
                st.warning("Carregue o modelo .docx na barra lateral.")
            else:
                with st.spinner("Gerando PDF..."):
                    bio = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                    pdf_bytes, erro = converter_docx_para_pdf(bio)
                    if pdf_bytes:
                        st.download_button("⬇️ PDF", pdf_bytes, "Relatorio.pdf", "application/pdf")
                    else:
                        st.error(f"Erro técnico: {erro}")

elif menu == "⚗️ Solução Preservativa":
    st.header("Solução Preservativa"); df = carregar_dados("Solucao")
    if not df.empty:
        df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        if st.button("Salvar"): salvar_dados(df_ed, "Solucao"); st.rerun()

elif menu == "📊 Dashboard":
    st.header("Dashboard"); df = carregar_dados("Madeira")
    if not df.empty and 'Nome do Cliente' in df.columns:
        st.plotly_chart(px.bar(df['Nome do Cliente'].value_counts().reset_index(), x='Nome do Cliente', y='count'))
