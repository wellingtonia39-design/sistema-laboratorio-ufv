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

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🌲")

# --- NOME DA PLANILHA ---
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- DIAGNÓSTICO LIBREOFFICE (Topo Lateral) ---
st.sidebar.title("🔧 Status PDF")
lo_path = shutil.which("libreoffice")
if lo_path:
    st.sidebar.success("✅ Conversor PDF Ativo")
else:
    st.sidebar.error("❌ Conversor PDF Inativo")
    st.sidebar.info("Crie o arquivo packages.txt com 'libreoffice' no GitHub.")

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
    try:
        with open("temp.docx", "wb") as f: f.write(docx_bytes.getvalue())
        cmd = ['libreoffice', '--headless', '--convert-to', 'pdf', 'temp.docx', '--outdir', '.']
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        if os.path.exists("temp.pdf"):
            with open("temp.pdf", "rb") as f: pdf = f.read()
            os.remove("temp.docx"); os.remove("temp.pdf")
            return pdf, None
        return None, "Erro: Arquivo PDF não gerado."
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
arquivo_modelo = st.sidebar.file_uploader("Carregar Modelo (.docx)", type=["docx"])

# --- ABA MADEIRA ---
if menu == "🪵 Madeira Tratada":
    st.header("Análise de Madeira Tratada")
    df = carregar_dados("Madeira")
    
    if not df.empty:
        if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
        
        # TABELA
        df_ed = st.data_editor(
            df, 
            num_rows="dynamic", 
            use_container_width=True, 
            height=400,
            column_config={"Selecionar": st.column_config.CheckboxColumn("Selecionar?", width="small")}
        )
        
        # BOTÃO SALVAR (Separado para evitar confusão)
        if st.button("💾 SALVAR DADOS NO GOOGLE SHEETS", type="primary", use_container_width=True):
            salvar_dados(df_ed, "Madeira")
            st.rerun()
            
        st.divider()
        st.markdown("### 🖨️ Área de Impressão")
        
        # LAYOUT DE BOTÕES LADO A LADO
        col_docx, col_pdf = st.columns([1, 1])
        
        # 1. BOTÃO DOCX
        with col_docx:
            st.markdown("##### Opção 1: Word")
            if st.button("📝 Gerar Relatórios DOCX", use_container_width=True):
                selecionados = df_ed[df_ed["Selecionar"] == True]
                if selecionados.empty:
                    st.error("⚠️ Selecione pelo menos uma linha na tabela acima.")
                elif not arquivo_modelo:
                    st.error("⚠️ Carregue o modelo .docx na barra lateral.")
                else:
                    if len(selecionados) == 1:
                        bio = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                        st.download_button("⬇️ Baixar DOCX Agora", bio, "Relatorio.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dw_docx")
                    else:
                        st.info("Para baixar vários, use a versão ZIP (não ativa neste botão).")

        # 2. BOTÃO PDF
        with col_pdf:
            st.markdown("##### Opção 2: PDF")
            # Este botão aparece SEMPRE. Não tem IF escondendo ele.
            if st.button("📄 Gerar Relatórios PDF", use_container_width=True):
                selecionados = df_ed[df_ed["Selecionar"] == True]
                
                # Validações
                if selecionados.empty:
                    st.error("⚠️ Selecione uma linha na tabela acima!")
                elif not arquivo_modelo:
                    st.error("⚠️ Carregue o modelo .docx na barra lateral!")
                else:
                    # Processo de Geração
                    with st.spinner("⏳ Convertendo para PDF..."):
                        # Passo 1: Gera Word
                        bio_docx = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                        
                        # Passo 2: Converte
                        pdf_bytes, erro = converter_docx_para_pdf(bio_docx)
                        
                        if pdf_bytes:
                            st.success("PDF Gerado!")
                            st.download_button("⬇️ Baixar PDF Agora", pdf_bytes, "Relatorio.pdf", "application/pdf", key="dw_pdf")
                        else:
                            st.error("❌ Falha na conversão.")
                            st.code(f"Erro técnico: {erro}")
                            if not lo_path:
                                st.warning("Diagnóstico: O servidor não achou o LibreOffice. Verifique packages.txt")

elif menu == "⚗️ Solução Preservativa":
    st.info("Mude para a aba Madeira para ver os relatórios")
    df = carregar_dados("Solucao")
    if not df.empty:
        df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        if st.button("Salvar Solução"): salvar_dados(df_ed, "Solucao"); st.rerun()
