import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from fpdf import FPDF
import io
import os
from datetime import datetime

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🌲")
NOME_PLANILHA_GOOGLE = "UFV_Laboratorio_DB"

# --- CLASSE DO PDF ---
class RelatorioPDF(FPDF):
    def header(self):
        # Logos
        if os.path.exists("logo_ufv.png"):
            self.image("logo_ufv.png", 10, 8, 30)
        if os.path.exists("logo_montana.png"):
            self.image("logo_montana.png", 170, 8, 30)
        
        self.set_y(15)
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'Relatório de Ensaio', 0, 1, 'C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

# --- FUNÇÕES AUXILIARES ---
def formatar_numero(valor, is_quimico=False):
    try:
        if not valor and valor != 0: return "-"
        v_str = str(valor).replace(",", ".")
        v_float = float(v_str)
        if is_quimico and v_float > 100: v_float /= 100.0
        return "{:,.2f}".format(v_float).replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(valor)

def formatar_data(valor):
    try:
        if not valor: return "-"
        v_str = str(valor).strip().split(" ")[0]
        for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y"]:
            try: return datetime.strptime(v_str, fmt).strftime("%d/%m/%Y")
            except: continue
        return v_str
    except: return str(valor)

def clean_text(text):
    if pd.isna(text): return ""
    return str(text).encode('latin-1', 'replace').decode('latin-1')

# --- GERADOR PDF (LAYOUT TABULAR) ---
def gerar_pdf_nativo(dados):
    pdf = RelatorioPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Altura padrão das linhas
    H_LINE = 7 

    # --- TABELA DE TOPO (ID e Datas) ---
    pdf.set_y(40)
    pdf.set_font('Arial', 'B', 9)
    
    # Linha 1: ID (Direita)
    id_rel = clean_text(dados.get("Código UFV", ""))
    pdf.set_x(130) # Move para direita
    pdf.cell(30, H_LINE, "Número ID:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(40, H_LINE, id_rel, 1, 1, 'C')
    
    # Linha 2: Datas (Esquerda e Direita)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(30, H_LINE, "Data de Entrada:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    dt_ent = formatar_data(dados.get("Data de entrada", ""))
    pdf.cell(40, H_LINE, dt_ent, 1, 0, 'C')
    
    # Espaço no meio
    pdf.cell(50, H_LINE, "", 0, 0) 
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(30, H_LINE, "Data Emissão:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    # Tenta pegar Data Registro, se não tiver, pega Fim Analise
    dt_emissao = formatar_data(dados.get("Data de Registro", "") if dados.get("Data de Registro") else dados.get("Fim da análise"))
    pdf.cell(40, H_LINE, dt_emissao, 1, 1, 'C')
    
    pdf.ln(5)

    # --- TABELA DADOS DO CLIENTE ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(190, H_LINE, "DADOS DO CLIENTE", 1, 1, 'L', fill=False) # Cabeçalho cinza se quiser: fill=True e set_fill_color
    
    # Linha Cliente
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(30, H_LINE, "Cliente:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(160, H_LINE, clean_text(dados.get("Nome do Cliente", "")), 1, 1, 'L')
    
    # Linha Cidade / Email
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(30, H_LINE, "Cidade/UF:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    cid = clean_text(dados.get("Cidade", ""))
    uf = clean_text(dados.get("Estado", ""))
    pdf.cell(65, H_LINE, f"{cid}/{uf}", 1, 0, 'L')
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(30, H_LINE, "E-mail:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(65, H_LINE, clean_text(dados.get("E-mail", "")), 1, 1, 'L')
    
    pdf.ln(5)

    # --- TABELA IDENTIFICAÇÃO DA AMOSTRA ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(190, H_LINE, clean_text("IDENTIFICAÇÃO DA AMOSTRA"), 1, 1, 'L')
    
    # Linha Amostra
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(30, H_LINE, "Ref. Cliente:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(160, H_LINE, clean_text(dados.get("Indentificação de Amostra do cliente", "")), 1, 1, 'L')
    
    # Linha Madeira / Produto
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(30, H_LINE, "Madeira:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(65, H_LINE, clean_text(dados.get("Madeira", "")), 1, 0, 'L')
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(30, H_LINE, "Produto:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(65, H_LINE, clean_text(dados.get("Produto utilizado", "")), 1, 1, 'L')

    # Linha Aplicação / Norma / Retenção
    # Vamos dividir 190mm por 3 blocos (aprox)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, H_LINE, clean_text("Aplicação:"), 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(43, H_LINE, clean_text(dados.get("Aplicação", "")), 1, 0, 'L')
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(25, H_LINE, "Norma ABNT:", 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(38, H_LINE, clean_text(dados.get("Norma ABNT", "")), 1, 0, 'L')
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(26, H_LINE, clean_text("Retenção Esp.:"), 1, 0, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(38, H_LINE, formatar_numero(dados.get("Retenção", ""), True), 1, 1, 'C')
    
    pdf.ln(5)

    # --- TABELA DE RESULTADOS QUÍMICOS ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(190, H_LINE, clean_text("RESULTADOS DE RETENÇÃO"), 1, 1, 'C')

    # Cabeçalho
    pdf.set_font('Arial', 'B', 8)
    # Linha 1 Headers
    x_i = pdf.get_x()
    y_i = pdf.get_y()
    
    pdf.cell(40, 10, "Ingredientes ativos", 1, 0, 'C')
    pdf.cell(35, 10, clean_text("Distribuição dos\nteores de i.a (kg/m3)"), 1, 0, 'C')
    
    # Bloco Balanceamento
    pdf.cell(75, 5, clean_text("Balanceamento químico"), 1, 0, 'C')
    
    # Metodo (Lado direito)
    pdf.set_xy(x_i + 150, y_i)
    pdf.cell(40, 10, clean_text("Método utilizado\nno ensaio"), 1, 0, 'C')

    # Linha 2 Headers
    pdf.set_xy(x_i + 75, y_i + 5)
    pdf.cell(25, 5, clean_text("Resultados (%)"), 1, 0, 'C')
    pdf.cell(50, 5, clean_text("Padrões (Min - Max)"), 1, 0, 'C')
    
    pdf.set_xy(x_i, y_i + 10) # Próxima linha de dados

    # DADOS
    pdf.set_font('Arial', '', 9)

    def linha_tab(nome, v_kg, v_pct, min_v, max_v, met=""):
        pdf.cell(40, H_LINE, clean_text(nome), 1, 0, 'L')
        pdf.cell(35, H_LINE, v_kg, 1, 0, 'C')
        pdf.cell(25, H_LINE, v_pct, 1, 0, 'C')
        pdf.cell(25, H_LINE, min_v, 1, 0, 'C')
        pdf.cell(25, H_LINE, max_v, 1, 0, 'C')
        pdf.cell(40, H_LINE, clean_text(met), 1, 1, 'C')

    linha_tab("Teor de CrO3 (Cromo)", 
              formatar_numero(dados.get("Retenção Cromo (Kg/m³)", ""), True),
              formatar_numero(dados.get("Balanço Cromo (%)", "")),
              "41,8", "53,2", "Metodo UFV 01")
              
    linha_tab("Teor de CuO (Cobre)", 
              formatar_numero(dados.get("Retenção Cobre (Kg/m³)", ""), True),
              formatar_numero(dados.get("Balanço Cobre (%)", "")),
              "15,2", "22,8", "")
              
    linha_tab("Teor de As2O5 (Arsênio)", 
              formatar_numero(dados.get("Retenção Arsênio (Kg/m³)", ""), True),
              formatar_numero(dados.get("Balanço Arsênio (%)", "")),
              "27,3", "40,7", "")
    
    # Total
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(40, H_LINE, clean_text("RETENÇÃO TOTAL"), 1, 0, 'L')
    pdf.cell(35, H_LINE, formatar_numero(dados.get("Soma Concentração (%)", ""), True), 1, 0, 'C')
    pdf.cell(25, H_LINE, formatar_numero(dados.get("Balanço Total (%)", "")), 1, 0, 'C')
    pdf.cell(90, H_LINE, clean_text("Nota: Resultados restritos as amostras"), 1, 1, 'C')
    
    pdf.ln(5)

    # --- TABELA PENETRAÇÃO ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(190, H_LINE, clean_text("RESULTADOS DE PENETRAÇÃO"), 1, 1, 'C')
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, H_LINE, "Grau", 1, 0, 'C')
    pdf.cell(50, H_LINE, "Tipo", 1, 0, 'L')
    pdf.cell(120, H_LINE, clean_text("Descrição"), 1, 1, 'L')
    
    pdf.set_font('Arial', '', 9)
    # Precisamos calcular a altura para a descrição longa
    desc = clean_text(dados.get("Descrição Penetração", ""))
    
    # Altura dinâmica (se o texto for longo)
    # Truque: Usamos multi_cell apenas na ultima coluna, mas precisamos saber a altura Y
    x_start = pdf.get_x()
    y_start = pdf.get_y()
    
    # Grau
    pdf.cell(20, 12, clean_text(dados.get("Grau de penetração", "")), 1, 0, 'C')
    # Tipo
    pdf.cell(50, 12, clean_text(dados.get("Descrição Grau", "")), 1, 0, 'L')
    
    # Descrição (MultiCell dentro de uma "célula" manual)
    pdf.set_xy(x_start + 70, y_start)
    pdf.multi_cell(120, 6, desc, 1, 'L')
    
    # Desenha a borda da caixa da descrição manualmente para cobrir a altura total (12)
    pdf.rect(x_start + 70, y_start, 120, 12)
    
    pdf.set_y(y_start + 12) # Avança para baixo
    pdf.ln(5)

    # --- OBSERVAÇÕES ---
    obs = clean_text(dados.get("Observação: Analista de Controle de Qualidade", ""))
    if obs:
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(190, H_LINE, "Observacoes:", 1, 1, 'L') # Header Obs
        pdf.set_font('Arial', '', 9)
        pdf.multi_cell(190, 6, obs, 1, 'L') # Conteúdo Obs com borda

    # --- ASSINATURAS ---
    pdf.set_y(-40)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 5, clean_text("Dr. Vinicius Resende de Castro - Supervisor do laboratório"), 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# --- CONEXÃO GOOGLE ---
def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
        client = gspread.authorize(creds)
        return client.open(NOME_PLANILHA_GOOGLE)
    except: return None

def carregar_dados(aba_nome):
    sh = conectar_google_sheets()
    if sh:
        try:
            df = pd.DataFrame(sh.worksheet(aba_nome).get_all_records())
            if not df.empty: df.columns = df.columns.str.strip()
            return df.astype(str)
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
        except: st.error("Erro Salvar")

# --- LOGIN ---
def check_login():
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    if st.session_state['logado']: return True
    st.markdown("<h2 style='text-align: center;'>🔐 Acesso UFV</h2>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        u = st.text_input("Usuário")
        s = st.text_input("Senha", type="password")
        if st.button("Entrar", type="primary"):
            sh = conectar_google_sheets()
            try:
                df = pd.DataFrame(sh.worksheet("Usuarios").get_all_records()).astype(str)
                user = df[(df['Usuario'] == u) & (df['Senha'] == s)]
                if not user.empty:
                    st.session_state.update({'logado': True, 'tipo': user.iloc[0]['Tipo'], 'user': u})
                    st.rerun()
                else: st.error("Acesso Negado")
            except: st.error("Erro Conexão")
    return False

# --- APP ---
if check_login():
    tipo = st.session_state['tipo']
    st.sidebar.info(f"👤 {st.session_state['user']} ({tipo})")
    if st.sidebar.button("Sair"): st.session_state['logado'] = False; st.rerun()

    st.title("🌲 Sistema Controle UFV")
    menu = st.sidebar.radio("Menu", ["🪵 Madeira Tratada", "⚗️ Solução Preservativa", "📊 Dashboard"])

    if menu == "🪵 Madeira Tratada":
        st.subheader("Análise de Madeira")
        df = carregar_dados("Madeira")
        if not df.empty:
            if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
            
            if tipo == "LPM":
                df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_config={"Selecionar": st.column_config.CheckboxColumn("Sel", width="small")})
                if st.button("💾 SALVAR", type="primary"): salvar_dados(df_ed, "Madeira"); st.rerun()
            else:
                cfg = {col: st.column_config.Column(disabled=True) for col in df.columns if col != "Selecionar"}
                cfg["Selecionar"] = st.column_config.CheckboxColumn("PDF", width="small", disabled=False)
                df_ed = st.data_editor(df, num_rows="fixed", use_container_width=True, column_config=cfg)

            sel = df_ed[df_ed["Selecionar"] == True]
            if st.button("📄 GERAR RELATÓRIO PDF", type="primary", use_container_width=True):
                if not sel.empty:
                    with st.spinner("Gerando PDF..."):
                        linha = sel.iloc[0]
                        nome = f"{linha.get('Código UFV', 'Relatorio')}.pdf"
                        try:
                            pdf = gerar_pdf_nativo(linha)
                            st.download_button(f"⬇️ BAIXAR {nome}", pdf, nome, "application/pdf")
                        except Exception as e: st.error(f"Erro: {e}")
                else: st.warning("Selecione uma amostra.")

    elif menu == "⚗️ Solução Preservativa":
        st.subheader("Solução")
        df = carregar_dados("Solucao")
        if not df.empty:
            if tipo == "LPM":
                df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True)
                if st.button("Salvar"): salvar_dados(df_ed, "Solucao"); st.rerun()
            else: st.dataframe(df, use_container_width=True)

    elif menu == "📊 Dashboard":
        df = carregar_dados("Madeira")
        if not df.empty:
            import plotly.express as px
            st.plotly_chart(px.bar(df['Nome do Cliente'].value_counts().reset_index(), x='Nome do Cliente', y='count'))
