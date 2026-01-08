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

# --- NOME DA PLANILHA NO GOOGLE ---
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- DIAGNÓSTICO DE SISTEMA (MOSTRA NO TOPO) ---
st.sidebar.title("🔧 Status do Servidor")
libreoffice_path = shutil.which("libreoffice")

if libreoffice_path:
    st.sidebar.success(f"✅ LibreOffice Encontrado!\nCaminho: {libreoffice_path}")
else:
    st.sidebar.error("❌ LibreOffice NÃO Encontrado!")
    st.sidebar.warning("O botão PDF vai aparecer, mas vai dar erro ao clicar.")
    st.sidebar.info("Verifique se criou o arquivo 'packages.txt' no GitHub.")

# --- MAPEAMENTO SIMPLIFICADO ---
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

def converter_docx_para_pdf(docx_bytes):
    # Função blindada para tentar converter e mostrar erro se falhar
    try:
        with open("temp.docx", "wb") as f: f.write(docx_bytes.getvalue())
        
        # Tenta converter
        processo = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', 'temp.docx', '--outdir', '.'],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
        )
        
        if os.path.exists("temp.pdf"):
            with open("temp.pdf", "rb") as f: pdf_bytes = f.read()
            os.remove("temp.docx"); os.remove("temp.pdf")
            return pdf_bytes, None
        else:
            return None, f"Erro LibreOffice: {processo.stderr.decode()}"
    except Exception as e: return None, str(e)

def preencher_modelo_word(modelo_upload, dados_linha):
    doc = Document(modelo_upload)
    
    def substituir(paragrafo, de, para):
        if de in paragrafo.text:
            try:
                # Tenta manter estilo
                for run in paragrafo.runs:
                    if de in run.text:
                        run.text = run.text.replace(de, str(para))
                        return
                # Fallback
                paragrafo.text = paragrafo.text.replace(de, str(para))
            except: pass

    # Preenchimento Simples (Sem formatação de virgula por enquanto, foco no PDF)
    dados_simples = {tag: str(dados_linha.get(col, "")) for col, tag in DE_PARA_WORD.items()}

    for tag, val in dados_simples.items():
        for p in doc.paragraphs: substituir(p, tag, val)
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells:
                    for p in c.paragraphs: substituir(p, tag, val)
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- MENU ---
st.title("🌲 Sistema UFV (Debug Mode)")
menu = st.sidebar.radio("Ir para:", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard"])
st.sidebar.markdown("---")
arquivo_modelo = st.sidebar.file_uploader("1. Carregue o Modelo .docx aqui", type=["docx"])

# --- ABA MADEIRA ---
if menu == "🪵 Madeira Tratada":
    st.header("Análise de Madeira Tratada")
    df = carregar_dados("Madeira")
    
    if not df.empty:
        if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
        
        # Tabela
        df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True, 
                             column_config={"Selecionar": st.column_config.CheckboxColumn("Selecionar", width="small")})
        
        # BOTÕES (Agora fora de colunas para garantir visibilidade)
        st.divider()
        st.markdown("### Ações")
        
        col1, col2, col3 = st.columns(3)
        
        if col1.button("💾 1. SALVAR DADOS", type="primary", use_container_width=True):
            salvar_dados(df_ed, "Madeira"); st.rerun()
            
        selecionados = df_ed[df_ed["Selecionar"] == True]
        
        if col2.button("📄 2. Baixar em WORD", use_container_width=True):
            if not selecionados.empty and arquivo_modelo:
                bio = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                st.download_button("⬇️ Download DOCX", bio, "Relatorio.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        
        # --- O BOTÃO DE PDF (SEM TRAVAS) ---
        if col3.button("📄 3. Baixar em PDF", use_container_width=True):
            if selecionados.empty:
                st.error("Selecione uma linha na tabela acima!")
            elif not arquivo_modelo:
                st.error("Carregue o arquivo .docx na barra lateral esquerda!")
            else:
                st.info("Tentando gerar PDF... Aguarde.")
                bio = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                
                # Tenta converter independente de verificação
                pdf_bytes, erro = converter_docx_para_pdf(bio)
                
                if pdf_bytes:
                    st.success("Sucesso!")
                    st.download_button("⬇️ CLIQUE AQUI PARA BAIXAR PDF", pdf_bytes, "Relatorio.pdf", "application/pdf")
                else:
                    st.error("ERRO CRÍTICO NA CONVERSÃO:")
                    st.code(erro) # Mostra o erro técnico
                    if not libreoffice_path:
                        st.warning("Diagnóstico: O servidor não encontrou o LibreOffice. Verifique o packages.txt")

elif menu == "⚗️ Solução Preservativa":
    st.info("Mude para a aba Madeira para testar o PDF")
    df = carregar_dados("Solucao")
    if not df.empty:
        st.data_editor(df, use_container_width=True)
