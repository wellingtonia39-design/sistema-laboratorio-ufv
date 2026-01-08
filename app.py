import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docx import Document
import io
import os
import subprocess
import shutil
from datetime import datetime

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🌲")
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- DIAGNÓSTICO AVANÇADO ---
st.sidebar.title("🔧 Diagnóstico Técnico")
# Tenta achar o executável com vários nomes possíveis
lo_bin = shutil.which("libreoffice") or shutil.which("soffice")

if lo_bin:
    st.sidebar.success(f"✅ Conversor Ativo!\nBinário: {lo_bin}")
else:
    st.sidebar.error("❌ Conversor NÃO encontrado.")
    st.sidebar.info("O arquivo packages.txt no GitHub deve conter: libreoffice, default-jre")

# --- MAPEAMENTO ---
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
    "Retenção Cromo (Kg/m³)": "«Retenção_Cromo_Kgm»",
    "Balanço Cromo (%)": "«Balanço_Cromo_»",
    "Retenção Cobre (Kg/m³)": "«Retenção_Cobre_Kgm»",
    "Balanço Cobre (%)": "«Balanço_Cobre_»",
    "Retenção Arsênio (Kg/m³)": "«Retenção_Arsênio_Kgm»",
    "Balanço Arsênio (%)": "«Balanço_Arsênio_»",
    "Soma Concentração (%)": "« Retençãoconcentração »",
    "Balanço Total (%)": "«Balanço_Total_»",
    "Grau de penetração": "«Grau_penetração»",
    "Descrição Grau": "«Descrição_Grau_»",
    "Descrição Penetração": "«Descrição_Penetração_»",
    "Observação: Analista de Controle de Qualidade": "«Observação»"
}

CAMPOS_NUMERICOS = ["Retenção", "Retenção Cromo (Kg/m³)", "Balanço Cromo (%)", "Retenção Cobre (Kg/m³)", "Balanço Cobre (%)", "Retenção Arsênio (Kg/m³)", "Balanço Arsênio (%)", "Soma Concentração (%)", "Balanço Total (%)"]
CAMPOS_DATA = ["Data de entrada", "Fim da análise", "Data de Registro"]

# --- FUNÇÕES ---
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
        except Exception as e: st.error(f"Erro Salvar: {e}")

def formatar_numero_br(valor):
    try:
        if valor == "" or valor is None: return ""
        if isinstance(valor, str): valor = valor.replace(",", ".")
        return "{:,.2f}".format(float(valor)).replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(valor)

def formatar_data_br(valor):
    if not valor: return ""
    v = str(valor).strip().split(" ")[0]
    for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"]:
        try: return datetime.strptime(v, fmt).strftime("%d/%m/%Y")
        except: continue
    return v

def converter_docx_para_pdf(docx_bytes):
    # Tenta encontrar o comando certo
    cmd_exec = shutil.which("libreoffice") or shutil.which("soffice")
    
    if not cmd_exec:
        return None, "O programa LibreOffice não foi encontrado no servidor."

    try:
        # Salva temporário
        with open("temp.docx", "wb") as f: f.write(docx_bytes.getvalue())
        
        # Comando Otimizado para Servidor (headless, sem logo, sem user profile)
        comando = [
            cmd_exec, 
            '--headless', 
            '--convert-to', 'pdf', 
            '--outdir', '.', 
            '--nologo', 
            '--nofirststartwizard',
            'temp.docx'
        ]
        
        # Executa com timeout maior
        processo = subprocess.run(
            comando,
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE, 
            timeout=120 # Aumentei para 2 minutos
        )
        
        if os.path.exists("temp.pdf"):
            with open("temp.pdf", "rb") as f: pdf = f.read()
            # Limpeza
            if os.path.exists("temp.docx"): os.remove("temp.docx")
            if os.path.exists("temp.pdf"): os.remove("temp.pdf")
            return pdf, None
        else:
            return None, f"Falha na conversão.\nSaída: {processo.stdout.decode()}\nErro: {processo.stderr.decode()}"
            
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

# --- INTERFACE ---
st.title("🌲 Sistema UFV")
menu = st.sidebar.radio("Ir para:", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard"])
st.sidebar.markdown("---")
arquivo_modelo = st.sidebar.file_uploader("1. Carregue o Modelo .docx aqui", type=["docx"])

if menu == "🪵 Madeira Tratada":
    st.header("Análise de Madeira Tratada")
    df = carregar_dados("Madeira")
    
    if not df.empty:
        if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
        
        df_ed = st.data_editor(
            df, num_rows="dynamic", use_container_width=True, height=400,
            column_config={"Selecionar": st.column_config.CheckboxColumn("Selecionar?", width="small")}
        )
        
        if st.button("💾 SALVAR DADOS NO GOOGLE SHEETS", type="primary", use_container_width=True):
            salvar_dados(df_ed, "Madeira"); st.rerun()
            
        st.divider()
        st.markdown("### 🖨️ Relatórios")
        c1, c2 = st.columns(2)
        
        # Botão Word
        if c1.button("📝 Gerar DOCX (Garantido)", use_container_width=True):
            sel = df_ed[df_ed["Selecionar"] == True]
            if not sel.empty and arquivo_modelo:
                bio = preencher_modelo_word(arquivo_modelo, sel.iloc[0])
                st.download_button("⬇️ Baixar DOCX", bio, "Relatorio.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else: st.error("Selecione uma linha e carregue o modelo.")

        # Botão PDF
        if c2.button("📄 Gerar PDF (Experimental)", use_container_width=True):
            sel = df_ed[df_ed["Selecionar"] == True]
            if sel.empty: st.error("Selecione uma linha!")
            elif not arquivo_modelo: st.error("Carregue o modelo!")
            else:
                with st.spinner("Gerando PDF (Pode demorar até 1 min)..."):
                    bio = preencher_modelo_word(arquivo_modelo, sel.iloc[0])
                    pdf, erro = converter_docx_para_pdf(bio)
                    
                    if pdf:
                        st.success("Sucesso!")
                        st.download_button("⬇️ Baixar PDF", pdf, "Relatorio.pdf", "application/pdf")
                    else:
                        st.error("Erro na conversão PDF.")
                        st.code(erro) # Mostre esse erro para mim!

elif menu == "⚗️ Solução Preservativa":
    st.info("Mude para Madeira"); df=carregar_dados("Solucao"); st.data_editor(df)