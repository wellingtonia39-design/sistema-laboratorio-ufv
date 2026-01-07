import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import plotly.express as px

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🪵")

# --- NOME DA PLANILHA NO GOOGLE ---
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- CONEXÃO COM GOOGLE SHEETS ---
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

# --- FUNÇÕES DE CARREGAR/SALVAR ---
def carregar_dados(aba_nome):
    sh = conectar_google_sheets()
    if sh:
        try:
            ws = sh.worksheet(aba_nome)
            dados = ws.get_all_records()
            return pd.DataFrame(dados)
        except gspread.exceptions.WorksheetNotFound:
            # Se a aba não existir, cria ela vazia
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
            # Prepara o DataFrame para envio (header + valores)
            lista_dados = [df.columns.values.tolist()] + df.values.tolist()
            ws.update(lista_dados)
            st.toast(f"Dados de {aba_nome} salvos com sucesso!", icon="✅")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

# --- INTERFACE PRINCIPAL ---
st.title("🌲 UFV - Controle de Qualidade de Madeira e Soluções")

# Menu Lateral
menu = st.sidebar.radio("Selecione o Módulo:", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard Geral"])
st.sidebar.divider()
st.sidebar.info("Modo de Edição Ativado via Google Sheets")

# ==================================================
# MÓDULO 1: MADEIRA TRATADA
# ==================================================
if menu == "🪵 Madeira Tratada":
    st.header("Análise de Madeira Tratada (NBR 16143)")
    
    # Carrega dados
    df_madeira = carregar_dados("Madeira")
    
    if df_madeira.empty:
        st.warning("A planilha 'Madeira' está vazia ou sem cabeçalho no Google. Adicione a primeira linha lá.")
    else:
        # Métricas Rápidas
        col1, col2, col3 = st.columns(3)
        total_amostras = len(df_madeira)
        # Tenta contar status se a coluna existir
        pendentes = len(df_madeira[df_madeira['Situação'] == 'Recebida']) if 'Situação' in df_madeira.columns else 0
        col1.metric("Total de Amostras", total_amostras)
        col2.metric("Amostras Recebidas/Pendentes", pendentes)
        
        st.divider()
        
        # --- EDITOR DE TABELA ---
        st.subheader("📝 Editar Registros")
        st.caption("Altere os valores diretamente na tabela abaixo e clique em SALVAR.")
        
        # O data_editor permite editar células como no Excel
        df_editado = st.data_editor(
            df_madeira,
            num_rows="dynamic", # Permite adicionar linhas
            use_container_width=True,
            height=500,
            key="editor_madeira"
        )
        
        # Botão de Salvar
        col_save1, col_save2 = st.columns([1, 4])
        if col_save1.button("💾 SALVAR ALTERAÇÕES", type="primary"):
            salvar_dados(df_editado, "Madeira")
            st.rerun()

# ==================================================
# MÓDULO 2: SOLUÇÃO PRESERVATIVA
# ==================================================
elif menu == "⚗️ Solução Preservativa":
    st.header("Análise de Solução Preservativa")
    
    df_solucao = carregar_dados("Solucao")
    
    if df_solucao.empty:
        st.warning("A planilha 'Solucao' está vazia. Adicione o cabeçalho no Google Sheets.")
    else:
        # Métricas
        c1, c2 = st.columns(2)
        c1.metric("Total de Soluções", len(df_solucao))
        
        # Exemplo de verificação de pH se a coluna existir
        if 'pH da solução' in df_solucao.columns:
            media_ph = pd.to_numeric(df_solucao['pH da solução'], errors='coerce').mean()
            c2.metric("pH Médio Global", f"{media_ph:.2f}")

        st.divider()
        
        st.subheader("📝 Editar Registros de Solução")
        df_editado_sol = st.data_editor(
            df_solucao,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_solucao"
        )
        
        if st.button("💾 SALVAR DADOS SOLUÇÃO", type="primary"):
            salvar_dados(df_editado_sol, "Solucao")
            st.rerun()

# ==================================================
# MÓDULO 3: DASHBOARD
# ==================================================
elif menu == "📊 Dashboard Geral":
    st.header("Visão Gerencial do Laboratório")
    
    df_m = carregar_dados("Madeira")
    
    if not df_m.empty and 'Nome do Cliente ' in df_m.columns: # Espaço no nome conforme seu CSV
        st.subheader("Amostras por Cliente")
        # Conta ocorrências por cliente
        contagem = df_m['Nome do Cliente '].value_counts().reset_index()
        contagem.columns = ['Cliente', 'Quantidade']
        
        fig = px.bar(contagem, x='Cliente', y='Quantidade', title="Volume de Análises por Cliente")
        st.plotly_chart(fig, use_container_width=True)
    
    if not df_m.empty and 'Situação' in df_m.columns:
        st.subheader("Status das Análises")
        fig2 = px.pie(df_m, names='Situação', title="Distribuição dos Status")
        st.plotly_chart(fig2, use_container_width=True)
        
    if df_m.empty:
        st.info("Preencha dados na aba Madeira para ver os gráficos.")