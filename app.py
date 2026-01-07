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
# Nota: Removi os espaços extras das chaves para bater com a limpeza automática
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

# Campos numéricos para formatar (Vírgula)
CAMPOS_NUMERICOS = [
    "Retenção", "Retenção Cromo (Kg/m³)", "Balanço Cromo (%)",
    "Retenção Cobre (Kg/m³)", "Balanço Cobre (%)",
    "Retenção Arsênio (Kg/m³)", "Balanço Arsênio (%)",
    "Soma Concentração (%)", "Balanço Total (%)"
]

# Campos de Data
CAMPOS_DATA = ["Data de entrada", "Fim da análise", "Data de Registro"]

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
            df = pd.DataFrame(dados)
            # LIMPEZA CRÍTICA: Remove espaços extras dos nomes das colunas (ex: "Cliente " -> "Cliente")
            df.columns = df.columns.str.strip()
            return df
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

# --- FORMATAÇÃO BRASILEIRA ---
def formatar_numero_br(valor):
    try:
        if valor == "" or valor is None: return ""
        # Se vier como string "6,5", troca pra "6.5" pra virar float
        if isinstance(valor, str):
            valor = valor.replace(",", ".")
        float_val = float(valor)
        # Formata BR
        return "{:,.2f}".format(float_val).replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(valor)

def formatar_data_br(valor):
    if not valor: return ""
    valor = str(valor).strip()
    # Tenta vários formatos
    formatos = ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"]
    for fmt in formatos:
        try:
            d = datetime.strptime(valor, fmt)
            return d.strftime("%d/%m/%Y")
        except:
            continue
    return valor

# --- CONVERSOR PDF (Robustez) ---
def verificar_libreoffice():
    """Verifica se o LibreOffice está instalado no sistema"""
    return shutil.which("libreoffice") is not None

def converter_docx_para_pdf(docx_bytes):
    try:
        with open("temp_doc.docx", "wb") as f:
            f.write(docx_bytes.getvalue())
        
        # Comando para converter
        cmd = ['libreoffice', '--headless', '--convert-to', 'pdf', 'temp_doc.docx', '--outdir', '.']
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
        
        if os.path.exists("temp_doc.pdf"):
            with open("temp_doc.pdf", "rb") as f:
                pdf_bytes = f.read()
            os.remove("temp_doc.docx")
            os.remove("temp_doc.pdf")
            return pdf_bytes
        return None
    except Exception as e:
        st.error(f"Erro PDF: {e}")
        return None

# --- PREENCHIMENTO WORD (Estilo Seguro) ---
def preencher_modelo_word(modelo_upload, dados_linha):
    doc = Document(modelo_upload)
    
    def substituir_preservando_estilo(paragrafo, de, para):
        if de in paragrafo.text:
            # Tenta substituir dentro dos 'runs' (pedaços de texto com mesmo estilo)
            # Isso evita perder negrito/italico
            substituiu_no_run = False
            for run in paragrafo.runs:
                if de in run.text:
                    run.text = run.text.replace(de, str(para))
                    substituiu_no_run = True
            
            # Se a tag estava quebrada entre runs (ex: parte normal, parte negrito),
            # substitui o texto inteiro do parágrafo (pode perder formatação mista, mas preenche)
            if not substituiu_no_run:
                paragrafo.text = paragrafo.text.replace(de, str(para))

    # Prepara os dados formatados
    dados_formatados = {}
    for col, tag in DE_PARA_WORD.items():
        valor = dados_linha.get(col, "")
        if col in CAMPOS_NUMERICOS:
            dados_formatados[tag] = formatar_numero_br(valor)
        elif col in CAMPOS_DATA:
            dados_formatados[tag] = formatar_data_br(valor)
        else:
            dados_formatados[tag] = str(valor)

    # Aplica no documento
    for tag, valor in dados_formatados.items():
        for p in doc.paragraphs:
            substituir_preservando_estilo(p, tag, valor)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        substituir_preservando_estilo(p, tag, valor)
    
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

# Verifica LibreOffice para aviso visual
tem_pdf = verificar_libreoffice()
if not tem_pdf:
    st.sidebar.warning("⚠️ Aviso: Conversor PDF não detectado. Adicione 'libreoffice' ao packages.txt e reinicie o app.")

if menu == "🪵 Madeira Tratada":
    st.header("Análise de Madeira Tratada")
    df_madeira = carregar_dados("Madeira")
    
    if not df_madeira.empty:
        if "Selecionar" not in df_madeira.columns:
            df_madeira.insert(0, "Selecionar", False)

        st.info("💡 Edite os valores na tabela abaixo e clique em Salvar.")
        df_editado = st.data_editor(
            df_madeira,
            num_rows="dynamic",
            use_container_width=True,
            height=400,
            column_config={
                "Selecionar": st.column_config.CheckboxColumn("Gerar Doc?", width="small")
            }
        )
        
        # Botões lado a lado
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("💾 SALVAR DADOS", type="primary", use_container_width=True):
                salvar_dados(df_editado, "Madeira")
                st.rerun()
        
        with col2:
            if st.button("📄 BAIXAR WORD", use_container_width=True):
                selecionados = df_editado[df_editado["Selecionar"] == True]
                if not selecionados.empty and arquivo_modelo:
                    if len(selecionados) == 1:
                        bio = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                        st.download_button("⬇️ Download .docx", bio, "Relatorio.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    else:
                        # Lógica ZIP para vários
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zf:
                            for idx, linha in selecionados.iterrows():
                                bio = preencher_modelo_word(arquivo_modelo, linha)
                                zf.writestr(f"Relatorio_{idx}.docx", bio.getvalue())
                        zip_buffer.seek(0)
                        st.download_button("⬇️ Download ZIP", zip_buffer, "Relatorios.zip", "application/zip")
                elif not arquivo_modelo: st.error("Falta o modelo .docx!")
                else: st.warning("Selecione uma amostra.")

        with col3:
            # Botão PDF só funciona se tiver selecionado
            if st.button("📄 BAIXAR PDF", use_container_width=True, disabled=not tem_pdf):
                selecionados = df_editado[df_editado["Selecionar"] == True]
                if not selecionados.empty and arquivo_modelo:
                    with st.spinner("Gerando PDF..."):
                        if len(selecionados) == 1:
                            bio_docx = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                            pdf_bytes = converter_docx_para_pdf(bio_docx)
                            if pdf_bytes:
                                st.download_button("⬇️ Download .pdf", pdf_bytes, "Relatorio.pdf", "application/pdf")
                            else:
                                st.error("Erro ao converter.")
                        else:
                            st.warning("Para PDF, selecione apenas 1 por vez.")
                elif not arquivo_modelo: st.error("Falta o modelo .docx!")
                else: st.warning("Selecione uma amostra.")

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
    if not df_m.empty and 'Nome do Cliente' in df_m.columns:
        contagem = df_m['Nome do Cliente'].value_counts().reset_index()
        contagem.columns = ['Cliente', 'Quantidade']
        st.plotly_chart(px.bar(contagem, x='Cliente', y='Quantidade', title="Análises por Cliente"))