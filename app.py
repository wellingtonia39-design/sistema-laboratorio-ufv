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

# Campos que devem ser tratados como número
CAMPOS_NUMERICOS = [
    "Retenção", "Retenção Cromo (Kg/m³)", "Balanço Cromo (%)",
    "Retenção Cobre (Kg/m³)", "Balanço Cobre (%)",
    "Retenção Arsênio (Kg/m³)", "Balanço Arsênio (%)",
    "Soma Concentração (%)", "Balanço Total (%)"
]

# Campos de Data
CAMPOS_DATA = ["Data de entrada", "Fim da análise", "Data de Registro"]

# --- CONEXÃO GOOGLE ---
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
        st.error(f"Erro Google: {e}")
        return None

def carregar_dados(aba_nome):
    sh = conectar_google_sheets()
    if sh:
        try:
            ws = sh.worksheet(aba_nome)
            df = pd.DataFrame(ws.get_all_records())
            # Remove espaços extras nos nomes das colunas
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
            ws.update([df_salvar.columns.values.tolist()] + df_salvar.values.tolist())
            st.toast("Salvo!", icon="✅")
        except Exception as e: st.error(f"Erro: {e}")

# --- FORMATAÇÃO INTELIGENTE ---
def formatar_numero_br(valor):
    try:
        if valor == "" or valor is None: return ""
        # Converte string para float
        if isinstance(valor, str): valor = valor.replace(",", ".")
        f_val = float(valor)
        
        # Correção Automática de Escala (Opcional - Ativar se necessário)
        # Se o valor for muito alto (ex: 368 onde deveria ser 3.68), divide por 100?
        # Por segurança, o sistema exibe o que está na tabela. 
        # Se aparecer 368,00, edite na tabela para 3.68
        
        return "{:,.2f}".format(f_val).replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(valor)

def formatar_data_br(valor):
    if not valor: return ""
    valor_str = str(valor).strip().split(" ")[0] # Tira hora
    # Lista de formatos (Incluindo o Americano Mês/Dia/Ano que apareceu no seu erro)
    formatos = [
        "%Y-%m-%d", # 2025-12-19
        "%m/%d/%Y", # 12/19/2025 (Americano)
        "%d/%m/%Y", # 19/12/2025 (BR)
        "%Y/%m/%d",
        "%d-%m-%Y"
    ]
    for fmt in formatos:
        try:
            d = datetime.strptime(valor_str, fmt)
            return d.strftime("%d/%m/%Y") # Força saída BR
        except: continue
    return valor_str

# --- CONVERSOR PDF (MODO DEBUG) ---
def converter_docx_para_pdf(docx_bytes):
    try:
        with open("temp.docx", "wb") as f: f.write(docx_bytes.getvalue())
        
        # Tenta rodar o LibreOffice e captura o erro se falhar
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', 'temp.docx', '--outdir', '.'],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
        )
        
        if os.path.exists("temp.pdf"):
            with open("temp.pdf", "rb") as f: pdf_bytes = f.read()
            os.remove("temp.docx"); os.remove("temp.pdf")
            return pdf_bytes, None
        else:
            # Retorna o erro exato do sistema
            erro_msg = result.stderr.decode()
            return None, f"LibreOffice falhou. Log: {erro_msg}"
    except Exception as e: return None, str(e)

# --- PREENCHIMENTO WORD ---
def preencher_modelo_word(modelo_upload, dados_linha):
    doc = Document(modelo_upload)
    
    def substituir(paragrafo, de, para):
        if de in paragrafo.text:
            # Tenta substituir mantendo estilo (dentro dos runs)
            for run in paragrafo.runs:
                if de in run.text:
                    run.text = run.text.replace(de, str(para))
                    return # Sucesso
            # Se falhar (tag quebrada), substitui no parágrafo (pode perder negrito)
            paragrafo.text = paragrafo.text.replace(de, str(para))

    # Prepara dados
    dados_fmt = {}
    for col, tag in DE_PARA_WORD.items():
        val = dados_linha.get(col, "")
        if col in CAMPOS_NUMERICOS: dados_fmt[tag] = formatar_numero_br(val)
        elif col in CAMPOS_DATA: dados_fmt[tag] = formatar_data_br(val)
        else: dados_fmt[tag] = str(val)

    # Aplica
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

# --- INTERFACE ---
st.title("🌲 UFV - Controle de Qualidade V8")

menu = st.sidebar.radio("Módulo:", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard"])
arquivo_modelo = st.sidebar.file_uploader("Carregar Modelo (.docx)", type=["docx"])

# DIAGNÓSTICO RÁPIDO
if shutil.which("libreoffice"):
    st.sidebar.success("✅ LibreOffice OK")
else:
    st.sidebar.warning("⚠️ LibreOffice NÃO encontrado no PATH.")

if menu == "🪵 Madeira Tratada":
    st.header("Madeira Tratada")
    df = carregar_dados("Madeira")
    
    if not df.empty:
        if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
        
        st.info("💡 Verifique os valores na tabela. O relatório imprime exatamente o que está aqui.")
        df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True, 
                             column_config={"Selecionar": st.column_config.CheckboxColumn("Gerar?", width="small")})
        
        c1, c2, c3 = st.columns(3)
        if c1.button("💾 SALVAR", type="primary", use_container_width=True):
            salvar_dados(df_ed, "Madeira"); st.rerun()
            
        sel = df_ed[df_ed["Selecionar"] == True]
        
        if c2.button("📄 BAIXAR WORD", use_container_width=True):
            if not sel.empty and arquivo_modelo:
                bio = preencher_modelo_word(arquivo_modelo, sel.iloc[0])
                st.download_button("⬇️ DOCX", bio, "Relatorio.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        # BOTÃO PDF FORÇADO
        if c3.button("📄 BAIXAR PDF", use_container_width=True):
            if sel.empty: st.warning("Selecione uma amostra.")
            elif not arquivo_modelo: st.error("Falta modelo.")
            else:
                with st.spinner("Gerando PDF..."):
                    bio = preencher_modelo_word(arquivo_modelo, sel.iloc[0])
                    pdf_bytes, erro = converter_docx_para_pdf(bio)
                    
                    if pdf_bytes:
                        st.download_button("⬇️ PDF PRONTO", pdf_bytes, "Relatorio.pdf", "application/pdf")
                    else:
                        st.error("Erro na conversão!")
                        st.code(erro) # Mostra o erro técnico na tela para sabermos o que houve

elif menu == "⚗️ Solução Preservativa":
    st.header("Solução"); df = carregar_dados("Solucao")
    if not df.empty:
        df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        if st.button("Salvar"): salvar_dados(df_ed, "Solucao"); st.rerun()
