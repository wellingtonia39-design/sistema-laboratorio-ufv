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

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🌲")
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- DIAGNÓSTICO PDF ---
lo_bin = shutil.which("libreoffice") or shutil.which("soffice")

# --- MAPEAMENTO WORD ---
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
            # Converte tudo para string para evitar erro de edição no Streamlit
            df = df.astype(str)
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
        if not valor: return ""
        v = str(valor).replace(",", ".")
        return "{:,.2f}".format(float(v)).replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(valor)

def formatar_data_br(valor):
    if not valor: return ""
    v = str(valor).strip().split(" ")[0]
    for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"]:
        try: return datetime.strptime(v, fmt).strftime("%d/%m/%Y")
        except: continue
    return v

def converter_docx_para_pdf(docx_bytes):
    if not lo_bin: return None, "LibreOffice não instalado."
    try:
        with open("temp.docx", "wb") as f: f.write(docx_bytes.getvalue())
        cmd = [lo_bin, '--headless', '--convert-to', 'pdf', '--outdir', '.', '--nologo', '--nofirststartwizard', 'temp.docx']
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
        if os.path.exists("temp.pdf"):
            with open("temp.pdf", "rb") as f: pdf = f.read()
            os.remove("temp.docx"); os.remove("temp.pdf")
            return pdf, None
        return None, "Erro na conversão."
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

# --- SISTEMA DE LOGIN ---
def check_login():
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    if 'tipo_usuario' not in st.session_state: st.session_state['tipo_usuario'] = None
    
    if st.session_state['logado']: return True

    st.markdown("<h1 style='text-align: center;'>🔐 Acesso Restrito UFV</h1>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        
        if st.button("Entrar", type="primary", use_container_width=True):
            sh = conectar_google_sheets()
            try:
                # Busca usuários na aba "Usuarios"
                ws = sh.worksheet("Usuarios")
                dados = ws.get_all_records()
                df_users = pd.DataFrame(dados)
                
                # Verifica credenciais
                user_encontrado = df_users[
                    (df_users['Usuario'].astype(str) == usuario) & 
                    (df_users['Senha'].astype(str) == senha)
                ]
                
                if not user_encontrado.empty:
                    st.session_state['logado'] = True
                    st.session_state['tipo_usuario'] = user_encontrado.iloc[0]['Tipo']
                    st.session_state['nome_usuario'] = usuario
                    st.toast(f"Bem-vindo, {usuario}!", icon="👋")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
            except Exception as e:
                st.error(f"Erro ao conectar na base de usuários: {e}")
                st.info("Verifique se a aba 'Usuarios' foi criada na planilha.")
    return False

# ===============================================
# APLICAÇÃO PRINCIPAL
# ===============================================

if check_login():
    # --- BARRA LATERAL ---
    tipo_user = st.session_state['tipo_usuario']
    st.sidebar.markdown(f"👤 **{st.session_state['nome_usuario']}** ({tipo_user})")
    
    if st.sidebar.button("Sair / Logout"):
        st.session_state['logado'] = False
        st.rerun()
        
    st.sidebar.divider()
    menu = st.sidebar.radio("Navegação:", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard"])
    st.sidebar.markdown("---")
    arquivo_modelo = st.sidebar.file_uploader("Modelo de Relatório (.docx)", type=["docx"])

    st.title("🌲 Sistema Controle UFV")

    # --- ABA MADEIRA ---
    if menu == "🪵 Madeira Tratada":
        st.header("Análise de Madeira Tratada")
        df = carregar_dados("Madeira")
        
        if not df.empty:
            if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
            
            # --- LÓGICA DE PERMISSÃO ---
            if tipo_user == "LPM":
                # LPM: Edita tudo
                st.info("🛠️ Modo Editor: Você pode alterar dados e salvar.")
                df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True, height=400,
                                     column_config={"Selecionar": st.column_config.CheckboxColumn("Selecionar?", width="small")})
                
                if st.button("💾 SALVAR ALTERAÇÕES", type="primary"):
                    salvar_dados(df_ed, "Madeira"); st.rerun()
            
            else:
                # Montana: Só seleciona (O resto fica travado)
                st.warning("👀 Modo Visualizador: Edição bloqueada.")
                
                # Configura todas as colunas para disabled=True, menos "Selecionar"
                col_config = {"Selecionar": st.column_config.CheckboxColumn("Selecionar?", width="small", disabled=False)}
                for col in df.columns:
                    if col != "Selecionar":
                        col_config[col] = st.column_config.Column(disabled=True) # Trava coluna
                
                df_ed = st.data_editor(df, num_rows="fixed", use_container_width=True, height=400, column_config=col_config)
                # Sem botão salvar para Montana

            st.divider()
            
            # ÁREA DE RELATÓRIOS (Visível para ambos)
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📝 Gerar DOCX", use_container_width=True):
                    sel = df_ed[df_ed["Selecionar"] == True]
                    if not sel.empty and arquivo_modelo:
                        bio = preencher_modelo_word(arquivo_modelo, sel.iloc[0])
                        st.download_button("⬇️ Baixar DOCX", bio, "Relatorio.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    else: st.error("Selecione e carregue o modelo.")

            with col2:
                if st.button("📄 Gerar PDF", use_container_width=True):
                    sel = df_ed[df_ed["Selecionar"] == True]
                    if not sel.empty and arquivo_modelo:
                        with st.spinner("Gerando PDF..."):
                            bio = preencher_modelo_word(arquivo_modelo, sel.iloc[0])
                            pdf, erro = converter_docx_para_pdf(bio)
                            if pdf: st.download_button("⬇️ Baixar PDF", pdf, "Relatorio.pdf", "application/pdf")
                            else: st.error(f"Erro: {erro}")
                    else: st.error("Selecione e carregue o modelo.")

    # --- ABA SOLUÇÃO ---
    elif menu == "⚗️ Solução Preservativa":
        st.header("Solução Preservativa")
        df_sol = carregar_dados("Solucao")
        if not df_sol.empty:
            if tipo_user == "LPM":
                df_sol_ed = st.data_editor(df_sol, num_rows="dynamic", use_container_width=True)
                if st.button("Salvar Solução"): salvar_dados(df_sol_ed, "Solucao"); st.rerun()
            else:
                st.dataframe(df_sol, use_container_width=True) # Apenas visualização para Montana

    # --- ABA DASHBOARD ---
    elif menu == "📊 Dashboard":
        st.header("Dashboard Gerencial")
        df_m = carregar_dados("Madeira")
        if not df_m.empty and 'Nome do Cliente' in df_m.columns:
            import plotly.express as px
            st.plotly_chart(px.bar(df_m['Nome do Cliente'].value_counts().reset_index(), x='Nome do Cliente', y='count'))