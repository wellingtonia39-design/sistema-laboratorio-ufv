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

# --- FUNÇÕES DE AJUDA ---
def clean_text(text):
    if pd.isna(text): return ""
    return str(text).encode('latin-1', 'replace').decode('latin-1')

def formatar_numero(valor, is_quimico=False):
    # Se for vazio, None ou traço, retorna 0.00 para cálculo, ou string vazia
    if valor in [None, "", "-"]: return 0.0 if is_quimico else ""
    
    try:
        v_str = str(valor).replace(",", ".")
        v_float = float(v_str)
        
        # Correção de escala (ex: 368 vira 3.68)
        if is_quimico and v_float > 100: v_float /= 100.0
        
        return v_float # Retorna float para poder somar se precisar
    except: return 0.0

def float_para_str(valor):
    # Converte o float de volta para string bonita (3.68 -> 3,68)
    if valor == 0: return "-"
    return "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_data(valor):
    try:
        if not valor: return "-"
        v_str = str(valor).strip().split(" ")[0]
        for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y"]:
            try: return datetime.strptime(v_str, fmt).strftime("%d/%m/%Y")
            except: continue
        return v_str
    except: return str(valor)

# --- BUSCA INTELIGENTE (IGNORA MAIUSCULAS E ESPAÇOS) ---
def buscar_valor(dados, chaves_possiveis):
    # Cria um dicionario com as chaves da planilha todas em minusculo
    dados_lower = {k.strip().lower(): v for k, v in dados.items()}
    
    for chave in chaves_possiveis:
        chave_limpa = chave.strip().lower()
        if chave_limpa in dados_lower:
            return dados_lower[chave_limpa]
    return "" # Não achou

# --- CLASSE PDF ---
class RelatorioPDF(FPDF):
    def header(self):
        if os.path.exists("logo_ufv.png"):
            self.image("logo_ufv.png", 10, 8, 25) # Logo menor
        if os.path.exists("logo_montana.png"):
            self.image("logo_montana.png", 175, 8, 25)
        
        self.set_y(12)
        self.set_font('Arial', 'B', 14) # Fonte Título menor
        self.cell(0, 10, clean_text('Relatório de Ensaio'), 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 7)
        self.cell(0, 10, clean_text(f'Página {self.page_no()}'), 0, 0, 'C')

    # CAMPO FORMULÁRIO (LABEL FORA, CAIXA DENTRO)
    def campo_form(self, label, valor, x, y, w, h=6, align='L', multi=False):
        # Rótulo (Label)
        self.set_xy(x, y)
        self.set_font('Arial', '', 7) # Label bem pequena
        self.cell(w, 3, clean_text(label), 0, 0, 'L')
        
        # Valor (Caixa)
        self.set_xy(x, y + 3.5)
        self.set_font('Arial', '', 9) # Valor tamanho 9 (confortável)
        
        if multi:
            # Para textos longos (Observação, Descrição)
            # Desenha a borda manual
            self.rect(x, y + 3.5, w, h)
            self.multi_cell(w, 4, clean_text(valor), 0, align)
        else:
            # Texto normal
            self.cell(w, h, clean_text(valor), 1, 0, align)

# --- GERADOR ---
def gerar_pdf_nativo(dados):
    pdf = RelatorioPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # --- PREPARAÇÃO DOS DADOS ---
    dt_entrada = formatar_data(buscar_valor(dados, ["Data de entrada", "Entrada"]))
    id_rel = clean_text(buscar_valor(dados, ["Código UFV", "Codigo", "ID"]))
    
    # Data Emissão
    raw_emi = buscar_valor(dados, ["Data de Registro", "Emissao", "Fim da análise"])
    dt_emissao = formatar_data(raw_emi)

    # --- TOPO ---
    y_topo = 30
    pdf.campo_form("Data de Entrada", dt_entrada, 10, y_topo, 40, align='C')
    pdf.campo_form("Número ID", id_rel, 150, y_topo - 5, 50, align='C')
    pdf.campo_form("Data de Emissão", dt_emissao, 150, y_topo + 8, 50, align='C')

    pdf.set_y(y_topo + 20)

    # --- CLIENTE ---
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, clean_text("DADOS DO CLIENTE"), 0, 1, 'L')
    y = pdf.get_y() + 1
    
    nome = buscar_valor(dados, ["Nome do Cliente", "Cliente"])
    pdf.campo_form("Cliente", nome, 10, y, 190)
    
    y += 11
    cid = clean_text(buscar_valor(dados, ["Cidade", "Municipio"]))
    uf = clean_text(buscar_valor(dados, ["Estado", "UF"]))
    email = clean_text(buscar_valor(dados, ["E-mail", "Email"]))
    
    pdf.campo_form("Cidade/UF", f"{cid}/{uf}", 10, y, 90)
    pdf.campo_form("E-mail", email, 105, y, 95)
    
    y += 15

    # --- AMOSTRA ---
    pdf.set_y(y)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, clean_text("IDENTIFICAÇÃO DA AMOSTRA"), 0, 1, 'L')
    y = pdf.get_y() + 1
    
    ref = buscar_valor(dados, ["Indentificação de Amostra do cliente", "Amostra", "Ref"])
    pdf.campo_form("Ref. Cliente (Amostra)", ref, 10, y, 190)
    
    y += 11
    mad = buscar_valor(dados, ["Madeira", "Especie"])
    prod = buscar_valor(dados, ["Produto utilizado", "Produto"])
    pdf.campo_form("Madeira", mad, 10, y, 90)
    pdf.campo_form("Produto", prod, 105, y, 95)
    
    y += 11
    app = buscar_valor(dados, ["Aplicação", "Aplicacao"])
    norma = buscar_valor(dados, ["Norma ABNT", "Norma"])
    # Retenção Específica (Valor alvo)
    ret_esp_raw = buscar_valor(dados, ["Retenção", "Retenção Esp", "Retencao"])
    ret_esp = float_para_str(formatar_numero(ret_esp_raw, True))
    
    pdf.campo_form("Aplicação", app, 10, y, 60)
    pdf.campo_form("Norma ABNT", norma, 75, y, 60)
    pdf.campo_form("Retenção Esp.", ret_esp, 140, y, 60, align='C')

    y += 20

    # --- RESULTADOS QUÍMICOS ---
    pdf.set_y(y)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(190, 6, clean_text("RESULTADOS DE RETENÇÃO"), 1, 1, 'C')
    
    # Cabeçalho Tabela
    pdf.set_font('Arial', 'B', 7) # Fonte reduzida no header
    h_row = 6
    
    x = 10
    curr_y = pdf.get_y()
    
    pdf.cell(40, 10, clean_text("Ingredientes ativos"), 1, 0, 'C')
    pdf.cell(30, 10, clean_text("Resultado (kg/m3)"), 1, 0, 'C')
    
    # Bloco Balanceamento
    pdf.cell(80, 5, clean_text("Balanceamento químico"), 1, 0, 'C')
    
    pdf.set_xy(x + 150, curr_y)
    pdf.cell(40, 10, clean_text("Método"), 1, 0, 'C')
    
    # Sub-bloco
    pdf.set_xy(x + 70, curr_y + 5)
    pdf.cell(30, 5, clean_text("Resultados (%)"), 1, 0, 'C')
    pdf.cell(50, 5, clean_text("Padrões (Min - Max)"), 1, 0, 'C')
    
    pdf.set_xy(x, curr_y + 10)

    # --- VALORES QUÍMICOS (Cálculo Automático) ---
    v_cr = formatar_numero(buscar_valor(dados, ["Retenção Cromo (Kg/m³)", "Retencao Cromo", "Cromo"]), True)
    v_cu = formatar_numero(buscar_valor(dados, ["Retenção Cobre (Kg/m³)", "Retencao Cobre", "Cobre"]), True)
    v_as = formatar_numero(buscar_valor(dados, ["Retenção Arsênio (Kg/m³)", "Retencao Arsenio", "Arsenio"]), True)
    
    # Balanços
    b_cr = buscar_valor(dados, ["Balanço Cromo (%)", "Balanco Cromo"])
    b_cu = buscar_valor(dados, ["Balanço Cobre (%)", "Balanco Cobre"])
    b_as = buscar_valor(dados, ["Balanço Arsênio (%)", "Balanco Arsenio"])

    pdf.set_font('Arial', '', 8) # Fonte valor tabela

    def row_tab(nome, v_val, v_bal, min_v, max_v, met=""):
        pdf.cell(40, h_row, clean_text(nome), 1, 0, 'L')
        pdf.cell(30, h_row, float_para_str(v_val), 1, 0, 'C')
        pdf.cell(30, h_row, clean_text(str(v_bal)), 1, 0, 'C')
        pdf.cell(25, h_row, min_v, 1, 0, 'C')
        pdf.cell(25, h_row, max_v, 1, 0, 'C')
        pdf.cell(40, h_row, clean_text(met), 1, 1, 'C')

    row_tab("Cromo (CrO3)", v_cr, b_cr, "41,8", "53,2", "Metodo UFV 01")
    row_tab("Cobre (CuO)", v_cu, b_cu, "15,2", "22,8", "")
    row_tab("Arsenio (As2O5)", v_as, b_as, "27,3", "40,7", "")
    
    # --- CÁLCULO TOTAL (Se a planilha falhar) ---
    # Tenta ler Total da planilha
    v_total = formatar_numero(buscar_valor(dados, ["Soma Concentração (%)", "Retenção Total", "Total"]), True)
    
    # Se planilha vier zerada, soma nós mesmos
    if v_total == 0:
        v_total = v_cr + v_cu + v_as
        
    b_total = buscar_valor(dados, ["Balanço Total (%)", "Balanco Total"])

    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, h_row, clean_text("RETENÇÃO TOTAL"), 1, 0, 'L')
    pdf.cell(30, h_row, float_para_str(v_total), 1, 0, 'C')
    pdf.cell(30, h_row, clean_text(str(b_total)), 1, 0, 'C')
    pdf.cell(90, h_row, clean_text("Nota: Resultados restritos as amostras"), 1, 1, 'C')
    
    pdf.ln(8)

    # --- PENETRAÇÃO (Correção do GRAU vazio) ---
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(190, 6, clean_text("RESULTADOS DE PENETRAÇÃO"), 0, 1, 'C')
    y = pdf.get_y() + 1

    # Tenta pegar Grau com várias chaves
    grau = buscar_valor(dados, ["Grau de penetração", "Grau", "grau", "Grau Penetracao"])
    tipo_grau = buscar_valor(dados, ["Descrição Grau", "Tipo Grau", "Tipo"])
    
    pdf.campo_form("Grau", grau, 10, y, 30, align='C')
    pdf.campo_form("Tipo", tipo_grau, 45, y, 50, align='C')
    
    desc = buscar_valor(dados, ["Descrição Penetração", "Descricao Penetracao", "Descricao"])
    
    # Descrição usando MultiCell (Com caixa manual)
    pdf.set_xy(100, y)
    pdf.set_font('Arial', '', 7)
    pdf.cell(90, 3, clean_text("Descrição"), 0, 0, 'L')
    
    # Caixa desenhada manualmente para o texto não vazar
    pdf.set_xy(100, y + 3.5)
    pdf.set_font('Arial', '', 8)
    pdf.rect(100, y + 3.5, 100, 12) # Caixa fixa de altura 12
    pdf.multi_cell(100, 4, clean_text(desc), 0, 'L')

    y += 20

    # --- OBSERVAÇÕES (MultiCell para texto longo) ---
    obs = buscar_valor(dados, ["Observação: Analista de Controle de Qualidade", "Observação", "Obs", "Observacoes"])
    
    if obs:
        pdf.set_y(y)
        # Usa multi=True para quebrar texto
        pdf.campo_form("Observações", obs, 10, y, 190, h=15, align='L', multi=True)

    # --- ASSINATURA ---
    pdf.set_y(-35)
    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 5, clean_text("Dr. Vinicius Resende de Castro - Supervisor do laboratório"), 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# --- GOOGLE E DADOS ---
def conectar():
    try:
        s = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        c = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(c, s)
        return gspread.authorize(creds).open(NOME_PLANILHA_GOOGLE)
    except: return None

def carregar(aba):
    sh = conectar()
    if sh:
        try:
            df = pd.DataFrame(sh.worksheet(aba).get_all_records())
            if not df.empty: df.columns = df.columns.str.strip()
            return df
        except: pass
    return pd.DataFrame()

def salvar(df, aba):
    sh = conectar()
    if sh:
        try:
            ws = sh.worksheet(aba)
            ws.clear()
            d = df.drop(columns=["Selecionar"]) if "Selecionar" in df.columns else df
            ws.update([d.columns.values.tolist()] + d.astype(str).values.tolist())
            st.toast("Salvo!")
        except: st.error("Erro salvar")

# --- LOGIN ---
def login():
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    if st.session_state['logado']: return True
    
    c1,c2,c3 = st.columns([1,2,1])
    with c2:
        st.title("🔐 Login UFV")
        u = st.text_input("User")
        s = st.text_input("Pass", type="password")
        if st.button("Entrar", type="primary", use_container_width=True):
            sh = conectar()
            try:
                df = pd.DataFrame(sh.worksheet("Usuarios").get_all_records()).astype(str)
                user = df[(df['Usuario']==u) & (df['Senha']==s)]
                if not user.empty:
                    st.session_state.update({'logado':True, 'tipo':user.iloc[0]['Tipo'], 'user':u})
                    st.rerun()
                else: st.error("Negado")
            except: st.error("Erro Login")
    return False

# --- APP ---
if login():
    tipo = st.session_state['tipo']
    st.sidebar.info(f"👤 {st.session_state['user']} ({tipo})")
    if st.sidebar.button("Sair"): st.session_state['logado']=False; st.rerun()

    # ESPIONAR COLUNAS (Para você ver o nome do Grau se falhar)
    with st.sidebar.expander("🕵️ Espiar Colunas"):
        df_x = carregar("Madeira")
        if not df_x.empty: st.write(list(df_x.columns))

    st.title("🌲 Sistema Controle UFV")
    menu = st.sidebar.radio("Menu", ["Madeira Tratada", "Solução", "Dashboard"])

    if menu == "Madeira Tratada":
        df = carregar("Madeira")
        if not df.empty:
            if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
            
            if tipo == "LPM":
                df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
                if st.button("Salvar"): salvar(df, "Madeira"); st.rerun()
            else:
                cfg = {c: st.column_config.Column(disabled=True) for c in df.columns if c!="Selecionar"}
                cfg["Selecionar"] = st.column_config.CheckboxColumn(disabled=False)
                df = st.data_editor(df, column_config=cfg, use_container_width=True)

            sel = df[df["Selecionar"]==True]
            if st.button("📄 GERAR PDF", type="primary", use_container_width=True):
                if not sel.empty:
                    try:
                        linha = sel.iloc[0].to_dict() # Converte para dicionario para busca
                        pdf = gerar_pdf_nativo(linha)
                        nome = f"{linha.get('Código UFV','Relatorio')}.pdf"
                        st.download_button(f"⬇️ {nome}", pdf, nome, "application/pdf")
                    except Exception as e: st.error(f"Erro: {e}")
                else: st.warning("Selecione um item.")

    elif menu == "Solução":
        df = carregar("Solucao")
        if not df.empty: st.dataframe(df, use_container_width=True)
    
    elif menu == "Dashboard":
        df = carregar("Madeira")
        if not df.empty:
            import plotly.express as px
            st.plotly_chart(px.bar(df['Nome do Cliente'].value_counts(), title="Amostras por Cliente"))
