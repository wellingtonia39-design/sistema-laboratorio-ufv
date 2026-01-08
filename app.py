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
import time

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🌲")
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- DIAGNÓSTICO PDF ---
lo_bin = shutil.which("libreoffice") or shutil.which("soffice")

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

# --- FUNÇÕES BÁSICAS ---
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
            df = df.astype(str) # Evita erros de tipo
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
            st.toast("Salvo com sucesso!", icon="✅")
        except Exception as e: st.error(f"Erro Salvar: {e}")

# --- FORMATAÇÃO ---
def formatar_numero_br(valor):
    try:
        if not valor: return ""
        v = str(valor).replace(",", ".")
        # Lógica para evitar números gigantes se o Excel mandou sem vírgula
        # Ex: Se vier 368 num campo de porcentagem/retenção, assume 3.68
        f_val = float(v)
        return "{:,.2f}".format(f_val).replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(valor)

def formatar_data_br(valor):
    if not valor: return ""
    v = str(valor).strip().split(" ")[0]
    for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"]:
        try: return datetime.strptime(v, fmt).strftime("%d/%m/%Y")
        except: continue
    return v

# --- CONVERSOR PDF OTIMIZADO ---
def converter_docx_para_pdf(docx_bytes):
    if not lo_bin: return None, "LibreOffice não instalado."
    try:
        with open("temp.docx", "wb") as f: f.write(docx_bytes.getvalue())
        
        # Comando Otimizado para melhor fidelidade visual
        cmd = [
            lo_bin, 
            '--headless', 
            '--convert-to', 'pdf:writer_pdf_Export', # Força exportador nativo
            '--outdir', '.', 
            '--nologo', 
            '--nofirststartwizard', 
            'temp.docx'
        ]
        
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
        
        if os.path.exists("temp.pdf"):
            with open("temp.pdf", "rb") as f: pdf = f.read()
            os.remove("temp.docx"); os.remove("temp.pdf")
            return pdf, None
        return None, "O arquivo PDF não foi gerado."
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

# --- LOGIN ---
def check_login():
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    if st.session_state['logado']: return True

    st.markdown("<h1 style='text-align: center;'>🔐 Acesso UFV</h1>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        u = st.text_input("Usuário")
        s = st.text_input("Senha", type="password")
        if st.button("Entrar", type="primary", use_container_width=True):
            sh = conectar_google_sheets()
            try:
                ws = sh.worksheet("Usuarios")
                df_users = pd.DataFrame(ws.get_all_records())
                # Converte para string para garantir comparação correta
                df_users['Usuario'] = df_users['Usuario'].astype(str)
                df_users['Senha'] = df_users['Senha'].astype(str)
                
                user = df_users[(df_users['Usuario'] == u) & (df_users['Senha'] == s)]
                if not user.empty:
                    st.session_state['logado'] = True
                    st.session_state['tipo'] = user.iloc[0]['Tipo']
                    st.session_state['user'] = u
                    st.rerun()
                else: st.error("Acesso Negado")
            except: st.error("Erro ao conectar no banco de usuários.")
    return False

# ================= APP =================
if check_login():
    tipo = st.session_state['tipo']
    usuario = st.session_state['user']
    
    # Barra Lateral
    st.sidebar.info(f"👤 **{usuario}** ({tipo})")
    if st.sidebar.button("Sair"):
        st.session_state['logado'] = False
        st.rerun()
    
    st.sidebar.divider()
    menu = st.sidebar.radio("Menu:", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard"])
    st.sidebar.markdown("---")
    arquivo_modelo = st.sidebar.file_uploader("Modelo de Relatório (.docx)", type=["docx"])

    st.title("🌲 Sistema Controle UFV")

    if menu == "🪵 Madeira Tratada":
        st.header("Análise de Madeira Tratada")
        df = carregar_dados("Madeira")
        
        if not df.empty:
            if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
            
            # --- CONFIGURAÇÃO DA TABELA (PERMISSÕES) ---
            if tipo == "LPM":
                # LPM: Pode editar tudo
                st.info("🛠️ Modo Editor (LPM)")
                df_ed = st.data_editor(
                    df, num_rows="dynamic", use_container_width=True, height=400,
                    column_config={"Selecionar": st.column_config.CheckboxColumn("Sel.", width="small")}
                )
                if st.button("💾 SALVAR ALTERAÇÕES", type="primary"):
                    salvar_dados(df_ed, "Madeira"); st.rerun()
            else:
                # Montana: Só vê e marca checkbox
                st.warning("👀 Modo Visualizador (Montana)")
                # Bloqueia todas as colunas exceto Selecionar
                cfg = {col: st.column_config.Column(disabled=True) for col in df.columns if col != "Selecionar"}
                cfg["Selecionar"] = st.column_config.CheckboxColumn("Gerar PDF?", width="small", disabled=False)
                
                df_ed = st.data_editor(
                    df, num_rows="fixed", use_container_width=True, height=400,
                    column_config=cfg
                )

            st.divider()
            
            # --- BOTÕES DE AÇÃO ---
            # Seleciona a linha marcada
            selecionados = df_ed[df_ed["Selecionar"] == True]
            
            # Prepara o nome do arquivo (Pega o ID da primeira linha selecionada)
            nome_arquivo = "Relatorio"
            if not selecionados.empty:
                id_amostra = str(selecionados.iloc[0].get("Código UFV", "")).strip()
                if id_amostra: nome_arquivo = id_amostra # Ex: UFV-M-620
            
            if tipo == "LPM":
                c1, c2 = st.columns(2)
                # LPM VÊ TUDO
                with c1:
                    if st.button("📝 Baixar DOCX", use_container_width=True):
                        if not selecionados.empty and arquivo_modelo:
                            bio = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                            st.download_button(f"⬇️ {nome_arquivo}.docx", bio, f"{nome_arquivo}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                        else: st.error("Selecione uma linha e o modelo.")
                
                with c2:
                    if st.button("📄 Baixar PDF", use_container_width=True):
                        if not selecionados.empty and arquivo_modelo:
                            with st.spinner("Gerando PDF..."):
                                bio = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                                pdf, erro = converter_docx_para_pdf(bio)
                                if pdf: st.download_button(f"⬇️ {nome_arquivo}.pdf", pdf, f"{nome_arquivo}.pdf", "application/pdf")
                                else: st.error(f"Erro: {erro}")
                        else: st.error("Selecione uma linha e o modelo.")
            
            else:
                # MONTANA SÓ VÊ PDF
                if st.button("📄 GERAR RELATÓRIO PDF", type="primary", use_container_width=True):
                    if not selecionados.empty and arquivo_modelo:
                        with st.spinner(f"Gerando PDF para {nome_arquivo}..."):
                            bio = preencher_modelo_word(arquivo_modelo, selecionados.iloc[0])
                            pdf, erro = converter_docx_para_pdf(bio)
                            if pdf: st.download_button(f"⬇️ BAIXAR {nome_arquivo}.pdf", pdf, f"{nome_arquivo}.pdf", "application/pdf")
                            else: st.error("Erro na conversão. Contate o laboratório.")
                    else: st.warning("Selecione a amostra na tabela e garanta que o modelo está carregado.")

    elif menu == "⚗️ Solução Preservativa":
        st.header("Solução Preservativa")
        df_sol = carregar_dados("Solucao")
        if not df_sol.empty:
            if tipo == "LPM":
                df_ed = st.data_editor(df_sol, num_rows="dynamic", use_container_width=True)
                if st.button("Salvar Solução"): salvar_dados(df_ed, "Solucao"); st.rerun()
            else:
                st.dataframe(df_sol, use_container_width=True)

    elif menu == "📊 Dashboard":
        st.header("Dashboard"); df = carregar_dados("Madeira")
        if not df.empty and 'Nome do Cliente' in df.columns:
            st.plotly_chart(px.bar(df['Nome do Cliente'].value_counts().reset_index(), x='Nome do Cliente', y='count'))