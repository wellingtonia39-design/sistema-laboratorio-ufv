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

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🪵")

# --- NOME DA PLANILHA NO GOOGLE ---
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- MAPEAMENTO (COLUNA EXCEL -> TAG WORD) ---
# Verifique se as tags no seu Word estão EXATAMENTE assim (letras maiúsculas/minúsculas importam)
DE_PARA_WORD = {
    "Código UFV": "«Código_UFV»",
    "Data de entrada": "«Data_de_entrada»",
    "Fim da análise": "«Fim_da_análise»",
    "Data de Registro": "«Data_de_Emissão»",
    "Nome do Cliente ": "«Nome_do_Cliente_»", 
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
    "Soma Concentração (%)": "« Retençãoconcentração »", # Com espaços conforme seu arquivo
    "Balanço Total (%)": "«Balanço_Total_»",
    
    # Penetração
    "Grau de penetração": "«Grau_penetração»",
    "Descrição Grau ": "«Descrição_Grau_»",
    "Descrição Penetração ": "«Descrição_Penetração_»",
    "Observação: Analista de Controle de Qualidade": "«Observação»"
}

# Campos que são Datas
CAMPOS_DATA = ["Data de entrada", "Fim da análise", "Data de Registro"]

# Campos Numéricos (para formatar com vírgula)
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

# --- FORMATAÇÃO BRASILEIRA ---
def formatar_numero_br(valor):
    """Converte ponto para vírgula e garante 2 casas decimais"""
    try:
        if not valor and valor != 0: return ""
        if isinstance(valor, str):
            valor = valor.replace(",", ".")
        float_val = float(valor)
        return "{:,.2f}".format(float_val).replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(valor)

def formatar_data_br(valor):
    """Tenta converter datas diversas para DD/MM/AAAA"""
    if not valor: return ""
    valor = str(valor).strip()
    formatos = ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"]
    for fmt in formatos:
        try:
            data_obj = datetime.strptime(valor, fmt)
            return data_obj.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return valor # Retorna original se falhar

# --- GERADOR PDF (Via LibreOffice) ---
def converter_docx_para_pdf(docx_bytes):
    """Salva o DOCX temporariamente, converte com LibreOffice e retorna bytes do PDF"""
    try:
        # Salva DOCX temporário
        with open("temp_doc.docx", "wb") as f:
            f.write(docx_bytes.getvalue())
        
        # Chama LibreOffice (precisa estar instalado no packages.txt)
        # O comando --headless roda sem interface gráfica (ideal para servidores)
        processo = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', 'temp_doc.docx', '--outdir', '.'],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        
        if os.path.exists("temp_doc.pdf"):
            with open("temp_doc.pdf", "rb") as f:
                pdf_bytes = f.read()
            # Limpeza
            os.remove("temp_doc.docx")
            os.remove("temp_doc.pdf")
            return pdf_bytes
        else:
            return None
    except Exception as e:
        st.error(f"Erro na conversão PDF: {e}")
        return None

# --- PREENCHIMENTO WORD (Melhorado para não quebrar estilo) ---
def preencher_modelo_word(modelo_upload, dados_linha):
    doc = Document(modelo_upload)
    
    # Função que tenta manter o estilo original (negrito, fonte, etc)
    def substituir_com_estilo(paragrafo, de, para):
        if de in paragrafo.text:
            # Tenta substituir mantendo o estilo do primeiro 'run' que contém o texto
            texto_completo = paragrafo.text
            novo_texto = texto_completo.replace(de, str(para))
            
            # Se a substituição for simples, tenta preservar runs (é complexo, então
            # a estratégia mais segura para não desfigurar é limpar e reescrever 
            # com o estilo do primeiro run, ou apenas substituir o texto se for simples)
            
            # Estratégia Segura: Substituição direta no texto do parágrafo
            # (Pode perder negrito parcial se a tag estiver no meio de uma frase formatada,
            # mas evita quebra de tabela)
            for run in paragrafo.runs:
                if de in run.text:
                    run.text = run.text.replace(de, str(para))
                    return # Substituiu no run específico, mantém estilo
            
            # Se a tag estiver dividida entre runs (ex: "«" num run e "Tag»" noutro),
            # a substituição acima falha. O fallback é substituir o texto do parágrafo todo.
            paragrafo.text = novo_texto

    for coluna_excel, tag_word in DE_PARA_WORD.items():
        valor_bruto = dados_linha.get(coluna_excel, "")
        
        # Formatações
        if coluna_excel in CAMPOS_NUMERICOS:
            valor_final = formatar_numero_br(valor_bruto)
        elif coluna_excel in CAMPOS_DATA:
            valor_final = formatar_data_br(valor_bruto)
        else:
            valor_final = str(valor_bruto)

        # Substituição
        for p in doc.paragraphs:
            substituir_com_estilo(p, tag_word, valor_final)
            
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        substituir_com_estilo(p, tag_word, valor_final)
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- INTERFACE ---
st.title("🌲 UFV - Controle de Qualidade V3")

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
            column_config={"Selecionar": st.column_config.CheckboxColumn("Relatório?", width="small")}
        )
        
        c1, c2, c3 = st.columns([1, 1, 1])
        if c1.button("💾 SALVAR DADOS", type="primary"):
            salvar_dados(df_editado, "Madeira")
            st.rerun()

        # Botão Word
        if c2.button("📄 BAIXAR WORD (.docx)"):
            selecionados = df_editado[df_editado["Selecionar"] == True]
            if not selecionados.empty and arquivo_modelo:
                with st.spinner("Gerando Word..."):
                    if len(selecionados) == 1:
                        linha = selecionados.iloc[0]
                        bio = preencher_modelo_word(arquivo_modelo, linha)
                        st.download_button("⬇️ Download DOCX", bio, f"Relatorio_{linha.get('Código UFV','amostra')}.docx", key="dw_word")
                    else:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zf:
                            for idx, linha in selecionados.iterrows():
                                bio = preencher_modelo_word(arquivo_modelo, linha)
                                zf.writestr(f"Relatorio_{linha.get('Código UFV', idx)}.docx", bio.getvalue())
                        zip_buffer.seek(0)
                        st.download_button("⬇️ Download ZIP (Word)", zip_buffer, "Relatorios_UFV.zip", "application/zip", key="dw_zip")
            elif not arquivo_modelo: st.warning("Carregue o modelo!")
            else: st.info("Selecione uma amostra.")

        # Botão PDF
        if c3.button("📄 BAIXAR PDF (.pdf)"):
            selecionados = df_editado[df_editado["Selecionar"] == True]
            if not selecionados.empty and arquivo_modelo:
                with st.spinner("Convertendo para PDF (Isso pode demorar um pouco)..."):
                    # Processo individual para PDF
                    if len(selecionados) == 1:
                        linha = selecionados.iloc[0]
                        bio_docx = preencher_modelo_word(arquivo_modelo, linha)
                        pdf_bytes = converter_docx_para_pdf(bio_docx)
                        
                        if pdf_bytes:
                            st.download_button("⬇️ Download PDF", pdf_bytes, f"Relatorio_{linha.get('Código UFV','amostra')}.pdf", "application/pdf", key="dw_pdf")
                        else:
                            st.error("Falha na conversão PDF. Verifique se o 'libreoffice' está no packages.txt ou tente baixar em Word.")
                    else:
                        st.warning("Para PDF, selecione apenas uma amostra por vez para evitar sobrecarga do servidor.")
            elif not arquivo_modelo: st.warning("Carregue o modelo!")
            else: st.info("Selecione uma amostra.")

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
