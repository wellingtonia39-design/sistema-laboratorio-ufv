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

# --- FUNÇÕES DE TEXTO E DADOS ---
def clean_text(text):
    if pd.isna(text): return ""
    return str(text).encode('latin-1', 'replace').decode('latin-1')

def formatar_numero(valor, is_quimico=False):
    # Se for None ou string vazia, retorna traço. Mas se for 0, formata!
    if valor is None or valor == "":
        return "-"
    
    try:
        v_str = str(valor).replace(",", ".")
        v_float = float(v_str)
        
        # Correção para números químicos (ex: 368 -> 3.68)
        if is_quimico and v_float > 100: 
            v_float /= 100.0
            
        return "{:,.2f}".format(v_float).replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(valor)

def formatar_data(valor):
    try:
        if not valor: return "-"
        v_str = str(valor).strip().split(" ")[0]
        for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y"]:
            try: return datetime.strptime(v_str, fmt).strftime("%d/%m/%Y")
            except: continue
        return v_str
    except: return str(valor)

# Função Inteligente para Buscar Colunas (Tenta variações do nome)
def buscar_valor(dados, lista_nomes):
    for nome in lista_nomes:
        if nome in dados:
            return dados[nome]
    return "" # Não achou nada

# --- CLASSE DO PDF ---
class RelatorioPDF(FPDF):
    def header(self):
        if os.path.exists("logo_ufv.png"):
            self.image("logo_ufv.png", 10, 8, 30)
        if os.path.exists("logo_montana.png"):
            self.image("logo_montana.png", 170, 8, 30)
        
        self.set_y(15)
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, clean_text('Relatório de Ensaio'), 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, clean_text(f'Página {self.page_no()}'), 0, 0, 'C')

    # Função para desenhar campo tipo formulário (Texto em cima, caixa embaixo)
    def campo_form(self, label, valor, x, y, w, h=7, align='L'):
        self.set_xy(x, y)
        self.set_font('Arial', '', 8)
        self.cell(w, 4, clean_text(label), 0, 0, 'L')
        self.set_xy(x, y + 4)
        self.set_font('Arial', '', 10)
        self.cell(w, h, clean_text(valor), 1, 0, align)

# --- GERADOR PDF ---
def gerar_pdf_nativo(dados):
    pdf = RelatorioPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Busca inteligente de dados
    dt_entrada = formatar_data(buscar_valor(dados, ["Data de entrada", "Entrada"]))
    id_rel = clean_text(buscar_valor(dados, ["Código UFV", "Codigo UFV", "ID"]))
    
    raw_emissao = buscar_valor(dados, ["Data de Registro", "Data Emissao", "Fim da análise"])
    dt_emissao = formatar_data(raw_emissao)

    # --- TOPO ---
    y_topo = 35
    pdf.campo_form("Data de Entrada", dt_entrada, x=10, y=y_topo, w=50, align='C')
    pdf.campo_form("Número ID", id_rel, x=140, y=y_topo - 5, w=60, align='C')
    pdf.campo_form("Data de Emissão", dt_emissao, x=140, y=y_topo + 8, w=60, align='C')

    pdf.set_y(y_topo + 25)

    # --- CLIENTE ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, clean_text("DADOS DO CLIENTE"), 0, 1, 'L')
    y_cli = pdf.get_y()
    
    nome_cli = buscar_valor(dados, ["Nome do Cliente", "Cliente"])
    pdf.campo_form("Cliente", nome_cli, x=10, y=y_cli, w=190)
    
    y_cli += 13
    cid = clean_text(buscar_valor(dados, ["Cidade", "Municipio"]))
    uf = clean_text(buscar_valor(dados, ["Estado", "UF"]))
    email = clean_text(buscar_valor(dados, ["E-mail", "Email"]))
    
    pdf.campo_form("Cidade/UF", f"{cid}/{uf}", x=10, y=y_cli, w=90)
    pdf.campo_form("E-mail", email, x=105, y=y_cli, w=95)
    
    pdf.set_y(y_cli + 15)

    # --- AMOSTRA ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, clean_text("IDENTIFICAÇÃO DA AMOSTRA"), 0, 1, 'L')
    y_ams = pdf.get_y()
    
    ref_cli = buscar_valor(dados, ["Indentificação de Amostra do cliente", "Amostra"])
    pdf.campo_form("Ref. Cliente (Amostra)", ref_cli, x=10, y=y_ams, w=190)
    
    y_ams += 13
    madeira = buscar_valor(dados, ["Madeira", "Especie"])
    produto = buscar_valor(dados, ["Produto utilizado", "Produto"])
    pdf.campo_form("Madeira", madeira, x=10, y=y_ams, w=90)
    pdf.campo_form("Produto", produto, x=105, y=y_ams, w=95)
    
    y_ams += 13
    aplicacao = buscar_valor(dados, ["Aplicação", "Aplicacao"])
    norma = buscar_valor(dados, ["Norma ABNT", "Norma"])
    ret_esp = formatar_numero(buscar_valor(dados, ["Retenção", "Retencao"]), True)
    
    pdf.campo_form("Aplicação", aplicacao, x=10, y=y_ams, w=60)
    pdf.campo_form("Norma ABNT", norma, x=75, y=y_ams, w=60)
    pdf.campo_form("Retenção Esp.", ret_esp, x=140, y=y_ams, w=60, align='C')

    pdf.set_y(y_ams + 20)

    # --- QUÍMICA ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(190, 8, clean_text("RESULTADOS DE RETENÇÃO"), 1, 1, 'C')

    # Header Tabela
    pdf.set_font('Arial', 'B', 8)
    x_i = pdf.get_x()
    y_i = pdf.get_y()
    
    pdf.cell(40, 10, clean_text("Ingredientes ativos"), 1, 0, 'C')
    pdf.cell(35, 10, clean_text("Resultado (kg/m3)"), 1, 0, 'C')
    pdf.cell(75, 5, clean_text("Balanceamento químico"), 1, 0, 'C')
    pdf.set_xy(x_i + 150, y_i)
    pdf.cell(40, 10, clean_text("Método"), 1, 0, 'C')
    pdf.set_xy(x_i + 75, y_i + 5)
    pdf.cell(25, 5, clean_text("Resultados (%)"), 1, 0, 'C')
    pdf.cell(50, 5, clean_text("Padrões (Min - Max)"), 1, 0, 'C')
    pdf.set_xy(x_i, y_i + 10)

    # Dados Tabela
    pdf.set_font('Arial', '', 9)
    H_ROW = 7

    def linha_tab(nome, v_kg, v_pct, min_v, max_v, met=""):
        pdf.cell(40, H_ROW, clean_text(nome), 1, 0, 'L')
        pdf.cell(35, H_ROW, v_kg, 1, 0, 'C')
        pdf.cell(25, H_ROW, v_pct, 1, 0, 'C')
        pdf.cell(25, H_ROW, min_v, 1, 0, 'C')
        pdf.cell(25, H_ROW, max_v, 1, 0, 'C')
        pdf.cell(40, H_ROW, clean_text(met), 1, 1, 'C')

    # Cromo
    cr_kg = formatar_numero(buscar_valor(dados, ["Retenção Cromo (Kg/m³)", "Retencao Cromo"]), True)
    cr_bal = formatar_numero(buscar_valor(dados, ["Balanço Cromo (%)", "Balanco Cromo"]))
    linha_tab("Cromo (CrO3)", cr_kg, cr_bal, "41,8", "53,2", "Metodo UFV 01")
              
    # Cobre
    cu_kg = formatar_numero(buscar_valor(dados, ["Retenção Cobre (Kg/m³)", "Retencao Cobre"]), True)
    cu_bal = formatar_numero(buscar_valor(dados, ["Balanço Cobre (%)", "Balanco Cobre"]))
    linha_tab("Cobre (CuO)", cu_kg, cu_bal, "15,2", "22,8", "")
              
    # Arsênio
    as_kg = formatar_numero(buscar_valor(dados, ["Retenção Arsênio (Kg/m³)", "Retencao Arsenio"]), True)
    as_bal = formatar_numero(buscar_valor(dados, ["Balanço Arsênio (%)", "Balanco Arsenio"]))
    linha_tab("Arsenio (As2O5)", as_kg, as_bal, "27,3", "40,7", "")
    
    # Total
    pdf.set_font('Arial', 'B', 9)
    # Tenta várias chaves para achar o Total
    tot_kg = formatar_numero(buscar_valor(dados, ["Soma Concentração (%)", "Soma Concentração", "Retenção Total", "Total"]), True)
    tot_bal = formatar_numero(buscar_valor(dados, ["Balanço Total (%)", "Balanço Total", "Balanco Total"]))
    
    pdf.cell(40, H_ROW, clean_text("RETENÇÃO TOTAL"), 1, 0, 'L')
    pdf.cell(35, H_ROW, tot_kg, 1, 0, 'C')
    pdf.cell(25, H_ROW, tot_bal, 1, 0, 'C')
    pdf.cell(90, H_ROW, clean_text("Nota: Resultados restritos as amostras"), 1, 1, 'C')
    
    pdf.ln(8)

    # --- PENETRAÇÃO ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(190, 6, clean_text("RESULTADOS DE PENETRAÇÃO"), 0, 1, 'C')
    y_pen = pdf.get_y()

    grau = buscar_valor(dados, ["Grau de penetração", "Grau", "grau"])
    tipo_grau = buscar_valor(dados, ["Descrição Grau", "Tipo Grau"])
    
    pdf.campo_form("Grau", grau, x=10, y=y_pen, w=30, align='C')
    pdf.campo_form("Tipo", tipo_grau, x=45, y=y_pen, w=50, align='C')
    
    desc = clean_text(buscar_valor(dados, ["Descrição Penetração", "Descricao Penetracao"]))
    pdf.set_xy(100, y_pen)
    pdf.set_font('Arial', '', 8)
    pdf.cell(90, 4, clean_text("Descrição"), 0, 0, 'L')
    pdf.set_xy(100, y_pen + 4)
    pdf.set_font('Arial', '', 9)
    pdf.multi_cell(100, 6, desc, 1, 'L')

    pdf.set_y(y_pen + 20)

    # --- OBS ---
    obs = clean_text(buscar_valor(dados, ["Observação: Analista de Controle de Qualidade", "Observação", "Obs"]))
    if obs:
        y_obs = pdf.get_y()
        pdf.campo_form("Observações", obs, x=10, y=y_obs, w=190, h=10)

    # Rodapé
    pdf.set_y(-35)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 5, clean_text("Dr. Vinicius Resende de Castro - Supervisor do laboratório"), 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# --- GOOGLE ---
def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
        return gspread.authorize(creds).open(NOME_PLANILHA_GOOGLE)
    except: return None

def carregar_dados(aba_nome):
    sh = conectar_google_sheets()
    if sh:
        try:
            df = pd.DataFrame(sh.worksheet(aba_nome).get_all_records())
            if not df.empty: df.columns = df.columns.str.strip() # Remove espaços extras
            return df.astype(str)
        except: return pd.DataFrame()
    return pd.DataFrame()

def salvar_dados(df, aba_nome):
    sh = conectar_google_sheets()
    if sh:
        try:
            ws = sh.worksheet(aba_nome)
            ws.clear()
            df_s = df.drop(columns=["Selecionar"]) if "Selecionar" in df.columns else df
            ws.update([df_s.columns.values.tolist()] + df_s.values.tolist())
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

    # --- DEBUGGER (USE ISTO SE OS CAMPOS CONTINUAREM VAZIOS) ---
    with st.sidebar.expander("🕵️ Ver Nomes das Colunas"):
        df_debug = carregar_dados("Madeira")
        if not df_debug.empty:
            st.write("Colunas encontradas:")
            st.write(list(df_debug.columns))
            st.write("Primeira linha de dados:")
            st.write(df_debug.iloc[0].to_dict())

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
