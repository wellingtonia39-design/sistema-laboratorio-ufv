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

# --- CLASSE DO PDF (LAYOUT RÉPLICA) ---
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

def gerar_pdf_nativo(dados):
    pdf = RelatorioPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # --- CABEÇALHO DE DADOS (Topo) ---
    pdf.set_font('Arial', 'B', 10)
    
    # Coluna Esquerda (Data Entrada)
    pdf.set_xy(10, 35)
    pdf.cell(40, 5, "Data de Entrada", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(40, 5, formatar_data(dados.get("Data de entrada", "")), 0, 1)

    # Coluna Direita (ID e Emissão)
    pdf.set_xy(140, 35)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(50, 5, "Número ID", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(50, 5, clean_text(dados.get("Código UFV", "")), 0, 1)
    
    pdf.set_xy(140, 48)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(50, 5, "Data de Emissão", 0, 1)
    pdf.set_font('Arial', '', 10)
    dt_emissao = formatar_data(dados.get("Data de Registro", "") if dados.get("Data de Registro") else dados.get("Fim da análise"))
    pdf.cell(50, 5, dt_emissao, 0, 1)

    pdf.ln(10) # Espaço
    pdf.set_y(60) # Força posição Y para começar o corpo

    # --- CLIENTE ---
    # Linha 1: Cliente
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(190, 5, "Cliente", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(190, 6, clean_text(dados.get("Nome do Cliente", "")), 0, 1)
    pdf.ln(2)

    # Linha 2: Cidade/UF e E-mail
    y_start = pdf.get_y()
    
    # Cidade
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(95, 5, "Cidade/UF", 0, 1)
    pdf.set_font('Arial', '', 10)
    cidade = clean_text(dados.get("Cidade", ""))
    estado = clean_text(dados.get("Estado", ""))
    pdf.cell(95, 6, f"{cidade}/{estado}", 0, 0)
    
    # E-mail (Lado direito)
    pdf.set_xy(105, y_start)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(95, 5, "E-mail", 0, 1)
    pdf.set_xy(105, y_start + 5)
    pdf.set_font('Arial', '', 10)
    pdf.cell(95, 6, clean_text(dados.get("E-mail", "")), 0, 1)
    
    pdf.ln(5)

    # --- IDENTIFICAÇÃO DA AMOSTRA ---
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(190, 8, clean_text("Identificação da amostra"), 0, 1)
    
    # Amostra
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(190, 5, "Amostra", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(190, 6, clean_text(dados.get("Indentificação de Amostra do cliente", "")), 0, 1)
    pdf.ln(2)

    # Linha: Madeira | Produto
    y_amostra = pdf.get_y()
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(95, 5, "Madeira", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(95, 6, clean_text(dados.get("Madeira", "")), 0, 0)

    pdf.set_xy(105, y_amostra)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(95, 5, "Produto", 0, 1)
    pdf.set_xy(105, y_amostra + 5)
    pdf.set_font('Arial', '', 10)
    pdf.cell(95, 6, clean_text(dados.get("Produto utilizado", "")), 0, 1)
    pdf.ln(2)

    # Linha: Aplicação | Norma | Retenção
    y_app = pdf.get_y()
    
    # Aplicação
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(63, 5, clean_text("Aplicação"), 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(63, 6, clean_text(dados.get("Aplicação", "")), 0, 0)

    # Norma
    pdf.set_xy(73, y_app)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(63, 5, "Norma ABNT", 0, 1)
    pdf.set_xy(73, y_app + 5)
    pdf.set_font('Arial', '', 10)
    pdf.cell(63, 6, clean_text(dados.get("Norma ABNT", "")), 0, 0)

    # Retenção Esp.
    pdf.set_xy(136, y_app)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(54, 5, clean_text("Retenção"), 0, 1)
    pdf.set_xy(136, y_app + 5)
    pdf.set_font('Arial', '', 10)
    pdf.cell(54, 6, formatar_numero(dados.get("Retenção", ""), True), 0, 1)
    
    pdf.ln(8)

    # --- TABELA DE RETENÇÃO (Cópia Visual) ---
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(190, 8, clean_text("Resultados de Retenção"), 0, 1)

    # Cabeçalho Complexo
    pdf.set_font('Arial', 'B', 8)
    # Linha 1 dos headers
    x_init = pdf.get_x()
    y_init = pdf.get_y()
    
    pdf.cell(40, 10, "Ingredientes ativos", 1, 0, 'C')
    pdf.cell(35, 10, clean_text("Distribuição dos\nteores de i.a (kg/m3)"), 1, 0, 'C') # Quebra linha manual se precisar
    
    # Bloco Balanceamento (Merged)
    pdf.cell(75, 5, clean_text("Balanceamento químico"), 1, 0, 'C')
    
    pdf.set_xy(x_init + 150, y_init)
    pdf.cell(40, 10, clean_text("Método utilizado\nno ensaio"), 1, 0, 'C')

    # Linha 2 dos headers (Sub-colunas)
    pdf.set_xy(x_init + 75, y_init + 5)
    pdf.cell(25, 5, clean_text("Resultados (%)"), 1, 0, 'C')
    pdf.cell(50, 5, clean_text("Padrões (Min - Max)"), 1, 0, 'C')
    
    pdf.set_xy(x_init, y_init + 10) # Volta para o começo da próxima linha

    # DADOS DA TABELA
    pdf.set_font('Arial', '', 9)
    altura_linha = 7

    def linha_tabela(nome, val_kg, val_pct, min_v, max_v, metodo=""):
        pdf.cell(40, altura_linha, clean_text(nome), 1, 0, 'L')
        pdf.cell(35, altura_linha, val_kg, 1, 0, 'C')
        pdf.cell(25, altura_linha, val_pct, 1, 0, 'C')
        pdf.cell(25, altura_linha, min_v, 1, 0, 'C') # Min
        pdf.cell(25, altura_linha, max_v, 1, 0, 'C') # Max
        pdf.cell(40, altura_linha, clean_text(metodo), 1, 1, 'C')

    # Cromo
    linha_tabela("Teor de CrO3 (Cromo)", 
                 formatar_numero(dados.get("Retenção Cromo (Kg/m³)", ""), True),
                 formatar_numero(dados.get("Balanço Cromo (%)", "")),
                 "41,8", "53,2", "Metodo UFV 01")
    # Cobre
    linha_tabela("Teor de CuO (Cobre)", 
                 formatar_numero(dados.get("Retenção Cobre (Kg/m³)", ""), True),
                 formatar_numero(dados.get("Balanço Cobre (%)", "")),
                 "15,2", "22,8", "")
    # Arsênio
    linha_tabela("Teor de As2O5 (Arsênio)", 
                 formatar_numero(dados.get("Retenção Arsênio (Kg/m³)", ""), True),
                 formatar_numero(dados.get("Balanço Arsênio (%)", "")),
                 "27,3", "40,7", "")
    
    # Linha Final (Retenção Total)
    pdf.set_font('Arial', 'B', 9)
    val_tot = formatar_numero(dados.get("Soma Concentração (%)", ""), True)
    bal_tot = formatar_numero(dados.get("Balanço Total (%)", "")) # Geralmente 100 ou próximo
    
    pdf.cell(40, altura_linha, clean_text("Retenção"), 1, 0, 'L')
    pdf.cell(35, altura_linha, val_tot, 1, 0, 'C')
    pdf.cell(25, altura_linha, bal_tot, 1, 0, 'C')
    pdf.cell(90, altura_linha, clean_text("Nota: os resultados restringem-se as amostras"), 1, 1, 'C')

    pdf.ln(5)

    # --- RESULTADOS DE PENETRAÇÃO ---
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(190, 8, clean_text("Resultados de Penetração"), 0, 1)

    # Header Tabela Penetração
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 6, "Grau", 1, 0, 'C')
    pdf.cell(50, 6, "Tipo", 1, 0, 'L')
    pdf.cell(120, 6, clean_text("Descrição da penetração"), 1, 1, 'L')

    # Dados
    pdf.set_font('Arial', '', 9)
    grau = clean_text(dados.get("Grau de penetração", ""))
    tipo_pen = clean_text(dados.get("Descrição Grau", ""))
    desc_pen = clean_text(dados.get("Descrição Penetração", ""))

    pdf.cell(20, 12, grau, 1, 0, 'C')
    pdf.cell(50, 12, tipo_pen, 1, 0, 'L')
    # Multi-cell para descrição longa
    x_desc = pdf.get_x()
    y_desc = pdf.get_y()
    pdf.multi_cell(120, 6, desc_pen, 1, 'L')
    # Desenha borda manual no resto da linha se multi_cell quebrar
    pdf.set_xy(10, y_desc + 12) # Avança para depois da tabela

    pdf.ln(5)

    # --- OBSERVAÇÕES ---
    obs = clean_text(dados.get("Observação: Analista de Controle de Qualidade", ""))
    if obs:
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(190, 6, clean_text("Observação: Analista de Controle de Qualidade"), 0, 1)
        pdf.set_font('Arial', '', 10)
        pdf.multi_cell(190, 5, obs, 0, 'L')
    
    # --- ASSINATURA ---
    # Posiciona no rodapé da página (aprox 4cm do fim)
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
                    with st.spinner("Gerando..."):
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
