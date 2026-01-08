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

# --- CLASSE DO PDF (LAYOUT FIXO) ---
class RelatorioPDF(FPDF):
    def header(self):
        # Logos (Verifica se existem antes de desenhar)
        if os.path.exists("logo_ufv.png"):
            self.image("logo_ufv.png", 10, 8, 33) # x, y, w
        if os.path.exists("logo_montana.png"):
            self.image("logo_montana.png", 160, 8, 40)
        
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'Relatório de Ensaio', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

# --- FUNÇÕES AUXILIARES ---
def formatar_numero(valor, is_quimico=False):
    """Formata para 0,00 e corrige escala se necessário"""
    try:
        if not valor: return "-"
        v_str = str(valor).replace(",", ".")
        v_float = float(v_str)
        
        # Correção automática: Se for químico e > 100 (ex: 368), vira 3.68
        if is_quimico and v_float > 100:
            v_float = v_float / 100.0
            
        return "{:,.2f}".format(v_float).replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(valor)

def formatar_data(valor):
    try:
        if not valor: return "-"
        v_str = str(valor).strip().split(" ")[0] # Remove hora
        for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y"]:
            try: return datetime.strptime(v_str, fmt).strftime("%d/%m/%Y")
            except: continue
        return v_str
    except: return str(valor)

def clean_text(text):
    """Remove caracteres incompatíveis com latin-1"""
    if pd.isna(text): return ""
    return str(text).encode('latin-1', 'replace').decode('latin-1')

# --- GERADOR PDF (A MÁGICA ACONTECE AQUI) ---
def gerar_pdf_nativo(dados):
    pdf = RelatorioPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # --- CABEÇALHO DO RELATÓRIO ---
    pdf.set_font('Arial', 'B', 12)
    pdf.ln(10) # Espaço após logos
    
    # ID DO RELATÓRIO
    id_relatorio = clean_text(dados.get("Código UFV", ""))
    pdf.cell(0, 8, f"Número ID: {id_relatorio}", 0, 1, 'C')
    pdf.ln(5)

    # --- DATAS ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(95, 6, "Data de Entrada:", 0, 0)
    pdf.cell(95, 6, "Data de Emissão (Fim da análise):", 0, 1)
    
    pdf.set_font('Arial', '', 10)
    dt_entrada = formatar_data(dados.get("Data de entrada", ""))
    dt_fim = formatar_data(dados.get("Fim da análise", ""))
    pdf.cell(95, 6, dt_entrada, 0, 0)
    pdf.cell(95, 6, dt_fim, 0, 1)
    pdf.ln(3)

    # --- DADOS DO CLIENTE (Bloco Cinza) ---
    pdf.set_fill_color(240, 240, 240) # Cinza claro
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "DADOS DO CLIENTE", 1, 1, 'L', fill=True)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(25, 6, "Cliente:", 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(165, 6, clean_text(dados.get("Nome do Cliente", "")), 0, 1)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(25, 6, "Cidade/UF:", 0, 0)
    pdf.set_font('Arial', '', 9)
    cidade = clean_text(dados.get("Cidade", ""))
    estado = clean_text(dados.get("Estado", ""))
    pdf.cell(80, 6, f"{cidade}/{estado}", 0, 0)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(15, 6, "E-mail:", 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(70, 6, clean_text(dados.get("E-mail", "")), 0, 1)
    pdf.ln(3)

    # --- DADOS DA AMOSTRA ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "IDENTIFICAÇÃO DA AMOSTRA", 1, 1, 'L', fill=True)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(40, 6, "Ref. Cliente:", 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(150, 6, clean_text(dados.get("Indentificação de Amostra do cliente", "")), 0, 1)
    
    # Linha com 3 colunas
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 6, "Madeira:", 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(40, 6, clean_text(dados.get("Madeira", "")), 0, 0)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 6, "Produto:", 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(40, 6, clean_text(dados.get("Produto utilizado", "")), 0, 0)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 6, "Aplicação:", 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(50, 6, clean_text(dados.get("Aplicação", "")), 0, 1)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(25, 6, "Norma ABNT:", 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(60, 6, clean_text(dados.get("Norma ABNT", "")), 0, 0)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(25, 6, "Retenção Esp.:", 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(60, 6, clean_text(dados.get("Retenção", "")), 0, 1)
    pdf.ln(5)

    # --- RESULTADOS QUÍMICOS (TABELA COMPLEXA) ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "RESULTADOS DE RETENÇÃO", 1, 1, 'C', fill=True)
    
    # Cabeçalho da Tabela
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, 10, "Ingrediente Ativo", 1, 0, 'C')
    pdf.cell(30, 10, "Resultado (kg/m3)", 1, 0, 'C') # Correção: m3 em vez de m³ para evitar erro de encoding simples
    pdf.cell(30, 10, "Balanco Quimico (%)", 1, 0, 'C')
    pdf.cell(50, 5, "Padroes Normativos (%)", 1, 0, 'C')
    pdf.cell(40, 10, "Metodo", 1, 0, 'C') # Sem acento para segurança
    
    # Sub-cabeçalho (Min/Max) - Posição manual
    x_atual = pdf.get_x()
    y_atual = pdf.get_y()
    pdf.set_xy(x_atual - 90, y_atual + 5) # Volta para a coluna Padrões
    pdf.cell(25, 5, "Min", 1, 0, 'C')
    pdf.cell(25, 5, "Max", 1, 0, 'C')
    pdf.set_xy(x_atual, y_atual + 5) # Volta para o final
    pdf.ln(5)

    # Linhas da Tabela
    pdf.set_font('Arial', '', 9)
    
    # Linha Cromo
    val_cr = formatar_numero(dados.get("Retenção Cromo (Kg/m³)", ""), True)
    bal_cr = formatar_numero(dados.get("Balanço Cromo (%)", ""))
    pdf.cell(40, 7, "Cromo (CrO3)", 1, 0)
    pdf.cell(30, 7, val_cr, 1, 0, 'C')
    pdf.cell(30, 7, bal_cr, 1, 0, 'C')
    pdf.cell(25, 7, "41,8", 1, 0, 'C') # Padrão
    pdf.cell(25, 7, "53,2", 1, 0, 'C') # Padrão
    pdf.cell(40, 7, "Metodo UFV 01", 1, 1, 'C')

    # Linha Cobre
    val_cu = formatar_numero(dados.get("Retenção Cobre (Kg/m³)", ""), True)
    bal_cu = formatar_numero(dados.get("Balanço Cobre (%)", ""))
    pdf.cell(40, 7, "Cobre (CuO)", 1, 0)
    pdf.cell(30, 7, val_cu, 1, 0, 'C')
    pdf.cell(30, 7, bal_cu, 1, 0, 'C')
    pdf.cell(25, 7, "15,2", 1, 0, 'C') 
    pdf.cell(25, 7, "22,8", 1, 0, 'C') 
    pdf.cell(40, 7, "-", 1, 1, 'C')

    # Linha Arsênio
    val_as = formatar_numero(dados.get("Retenção Arsênio (Kg/m³)", ""), True)
    bal_as = formatar_numero(dados.get("Balanço Arsênio (%)", ""))
    pdf.cell(40, 7, "Arsenio (As2O5)", 1, 0)
    pdf.cell(30, 7, val_as, 1, 0, 'C')
    pdf.cell(30, 7, bal_as, 1, 0, 'C')
    pdf.cell(25, 7, "27,3", 1, 0, 'C') 
    pdf.cell(25, 7, "40,7", 1, 0, 'C') 
    pdf.cell(40, 7, "-", 1, 1, 'C')

    # Linha Total
    pdf.set_font('Arial', 'B', 9)
    val_tot = formatar_numero(dados.get("Soma Concentração (%)", ""), True) # Aqui usa nome exato da planilha
    bal_tot = formatar_numero(dados.get("Balanço Total (%)", ""))
    pdf.cell(40, 7, "RETENCAO TOTAL", 1, 0)
    pdf.cell(30, 7, val_tot, 1, 0, 'C')
    pdf.cell(30, 7, bal_tot, 1, 0, 'C')
    pdf.cell(90, 7, "Nota: Resultados restritos as amostras.", 1, 1, 'C')
    pdf.ln(5)

    # --- PENETRAÇÃO ---
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "RESULTADOS DE PENETRAÇÃO", 1, 1, 'C', fill=True)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(20, 6, "Grau", 1, 0, 'C')
    pdf.cell(40, 6, "Tipo", 1, 0, 'L')
    pdf.cell(130, 6, "Descricao", 1, 1, 'L')
    
    pdf.set_font('Arial', '', 9)
    grau = clean_text(dados.get("Grau de penetração", ""))
    tipo_pen = clean_text(dados.get("Descrição Grau", "")) # Ajuste conforme nome exato coluna
    desc_pen = clean_text(dados.get("Descrição Penetração", "")) # Ajuste conforme nome exato coluna
    
    pdf.cell(20, 10, grau, 1, 0, 'C')
    pdf.cell(40, 10, tipo_pen, 1, 0, 'L')
    pdf.multi_cell(130, 10, desc_pen, 1, 'L')
    pdf.ln(5)

    # --- OBSERVAÇÕES ---
    obs = clean_text(dados.get("Observação: Analista de Controle de Qualidade", ""))
    if obs:
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(0, 6, "Observacoes:", 0, 1)
        pdf.set_font('Arial', '', 9)
        pdf.multi_cell(0, 5, obs, 1)
        pdf.ln(5)

    # --- ASSINATURAS ---
    pdf.ln(20) # Espaço para assinatura
    y_ass = pdf.get_y()
    
    pdf.line(20, y_ass, 90, y_ass)
    pdf.line(120, y_ass, 190, y_ass)
    
    pdf.set_font('Arial', '', 8)
    pdf.cell(95, 5, "Analista de Controle de Qualidade", 0, 0, 'C')
    pdf.cell(95, 5, "Dr. Vinicius Resende de Castro", 0, 1, 'C')
    pdf.cell(95, 5, "", 0, 0, 'C')
    pdf.cell(95, 5, "Supervisor do Laboratorio", 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# --- CONEXÃO GOOGLE ---
def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
        client = gspread.authorize(creds)
        return client.open(NOME_PLANILHA_GOOGLE)
    except Exception as e:
        st.error(f"Erro Google: {e}")
        return None

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
        except: st.error("Erro ao salvar")

# --- LOGIN ---
def check_login():
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    if st.session_state['logado']: return True
    st.markdown("<h2 style='text-align: center;'>🔐 Acesso UFV</h2>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        u = st.text_input("Usuário")
        s = st.text_input("Senha", type="password")
        if st.button("Entrar", type="primary", use_container_width=True):
            sh = conectar_google_sheets()
            try:
                df = pd.DataFrame(sh.worksheet("Usuarios").get_all_records()).astype(str)
                user = df[(df['Usuario'] == u) & (df['Senha'] == s)]
                if not user.empty:
                    st.session_state.update({'logado': True, 'tipo': user.iloc[0]['Tipo'], 'user': u})
                    st.rerun()
                else: st.error("Acesso Negado")
            except: st.error("Erro Login")
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
            
            # Tabela (Editável ou Travada)
            if tipo == "LPM":
                df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_config={"Selecionar": st.column_config.CheckboxColumn("Sel.", width="small")})
                if st.button("💾 SALVAR DADOS", type="primary"): salvar_dados(df_ed, "Madeira"); st.rerun()
            else:
                cfg = {col: st.column_config.Column(disabled=True) for col in df.columns if col != "Selecionar"}
                cfg["Selecionar"] = st.column_config.CheckboxColumn("PDF", width="small", disabled=False)
                df_ed = st.data_editor(df, num_rows="fixed", use_container_width=True, column_config=cfg)

            # Botão PDF
            selecionados = df_ed[df_ed["Selecionar"] == True]
            if st.button("📄 GERAR PDF (OFICIAL)", type="primary", use_container_width=True):
                if not selecionados.empty:
                    linha = selecionados.iloc[0]
                    nome_arq = f"{linha.get('Código UFV', 'Relatorio')}.pdf"
                    
                    try:
                        pdf_bytes = gerar_pdf_nativo(linha)
                        st.download_button(f"⬇️ BAIXAR {nome_arq}", pdf_bytes, nome_arq, "application/pdf")
                        st.success("PDF Gerado com Sucesso!")
                    except Exception as e:
                        st.error(f"Erro ao desenhar PDF: {e}")
                else:
                    st.warning("Selecione uma amostra.")

    elif menu == "⚗️ Solução Preservativa":
        st.subheader("Solução")
        df = carregar_dados("Solucao")
        if not df.empty:
            if tipo == "LPM":
                df_ed = st.data_editor(df, num_rows="dynamic", use_container_width=True)
                if st.button("Salvar"): salvar_dados(df_ed, "Solucao"); st.rerun()
            else: st.dataframe(df, use_container_width=True)
            
    elif menu == "📊 Dashboard":
        st.subheader("Dashboard")
        df = carregar_dados("Madeira")
        if not df.empty:
            import plotly.express as px
            st.plotly_chart(px.bar(df['Nome do Cliente'].value_counts().reset_index(), x='Nome do Cliente', y='count'))