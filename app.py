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

# --- FUNÇÕES ---
def clean_text(text):
    if pd.isna(text): return ""
    return str(text).encode('latin-1', 'replace').decode('latin-1')

def formatar_numero(valor, is_quimico=False):
    if valor in [None, "", "-"]: return 0.0
    try:
        v_str = str(valor).replace(",", ".")
        v_float = float(v_str)
        
        # Lógica inteligente para porcentagens quebradas
        if is_quimico:
            # Se for > 100 (ex: 458), divide por 10.0 (vira 45.8) ou 100 dependendo da grandeza
            # Assumindo que balanceamento é sempre < 100%
            if v_float > 100: 
                v_float /= 10.0 # Tenta dividir por 10 primeiro (458 -> 45.8)
                if v_float > 100: v_float /= 10.0 # Se ainda for alto, divide de novo
                
        return v_float
    except: return 0.0

def float_para_str(valor):
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

def buscar_valor(dados, chaves_possiveis):
    # Cria mapa de chaves normalizadas
    dados_norm = {k.strip().lower(): v for k, v in dados.items()}
    
    for chave in chaves_possiveis:
        k = chave.strip().lower()
        # Busca exata
        if k in dados_norm:
            val = dados_norm[k]
            if str(val).strip() != "": return val
            
        # Busca parcial (ex: "Grau" acha "Grau de Penetração")
        for coluna_real in dados_norm.keys():
            if k in coluna_real:
                val = dados_norm[coluna_real]
                if str(val).strip() != "": return val
                
    return ""

# --- CLASSE PDF ---
class RelatorioPDF(FPDF):
    def header(self):
        if os.path.exists("logo_ufv.png"): self.image("logo_ufv.png", 10, 8, 25)
        if os.path.exists("logo_montana.png"): self.image("logo_montana.png", 175, 8, 25)
        self.set_y(12)
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, clean_text('Relatório de Ensaio'), 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 6)
        self.cell(0, 10, clean_text(f'Página {self.page_no()}'), 0, 0, 'C')

    def campo_form(self, label, valor, x, y, w, h=6, align='L', multi=False):
        # Label (Fonte tamanho 6)
        self.set_xy(x, y)
        self.set_font('Arial', '', 6)
        self.cell(w, 3, clean_text(label), 0, 0, 'L')
        
        # Valor (Fonte tamanho 8 - mais compacto)
        self.set_xy(x, y + 3)
        self.set_font('Arial', '', 8)
        if multi:
            self.rect(x, y + 3, w, h)
            self.multi_cell(w, 4, clean_text(valor), 0, align)
        else:
            self.cell(w, h, clean_text(valor), 1, 0, align)

# --- GERADOR ---
def gerar_pdf_nativo(dados):
    pdf = RelatorioPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # --- 1. CABEÇALHO ---
    dt_entrada = formatar_data(buscar_valor(dados, ["Data de entrada", "Entrada"]))
    id_rel = clean_text(buscar_valor(dados, ["Código UFV", "ID", "Codigo"]))
    dt_emissao = formatar_data(buscar_valor(dados, ["Data de Registro", "Emissao", "Fim da análise"]))

    y = 30
    pdf.campo_form("Data de Entrada", dt_entrada, 10, y, 40, align='C')
    pdf.campo_form("Número ID", id_rel, 150, y - 5, 50, align='C')
    pdf.campo_form("Data de Emissão", dt_emissao, 150, y + 8, 50, align='C')

    # --- 2. CLIENTE ---
    y += 20
    pdf.set_y(y); pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, clean_text("DADOS DO CLIENTE"), 0, 1, 'L')
    y += 6
    
    cli = buscar_valor(dados, ["Nome do Cliente", "Cliente", "Empresa"])
    pdf.campo_form("Cliente", cli, 10, y, 190)
    
    y += 11
    cid = clean_text(buscar_valor(dados, ["Cidade", "Municipio"]))
    uf = clean_text(buscar_valor(dados, ["Estado", "UF"]))
    email = clean_text(buscar_valor(dados, ["E-mail", "Email"]))
    
    pdf.campo_form("Cidade/UF", f"{cid}/{uf}", 10, y, 90)
    pdf.campo_form("E-mail", email, 105, y, 95)

    # --- 3. AMOSTRA ---
    y += 15
    pdf.set_y(y); pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, clean_text("IDENTIFICAÇÃO DA AMOSTRA"), 0, 1, 'L')
    y += 6
    
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
    ret_esp = float_para_str(formatar_numero(buscar_valor(dados, ["Retenção", "Retenção Esp."]), True))
    
    pdf.campo_form("Aplicação", app, 10, y, 60)
    pdf.campo_form("Norma ABNT", norma, 75, y, 60)
    pdf.campo_form("Retenção Esp.", ret_esp, 140, y, 60, align='C')

    # --- 4. QUÍMICA ---
    y += 20
    pdf.set_y(y); pdf.set_font('Arial', 'B', 9)
    pdf.cell(190, 6, clean_text("RESULTADOS DE RETENÇÃO"), 1, 1, 'C')
    
    # Header
    pdf.set_font('Arial', 'B', 7)
    x = 10; cy = pdf.get_y()
    pdf.cell(40, 10, clean_text("Ingredientes ativos"), 1, 0, 'C')
    pdf.cell(30, 10, clean_text("Resultado (kg/m3)"), 1, 0, 'C')
    pdf.cell(80, 5, clean_text("Balanceamento químico"), 1, 0, 'C')
    pdf.set_xy(x+150, cy)
    pdf.cell(40, 10, clean_text("Método"), 1, 0, 'C')
    
    pdf.set_xy(x+70, cy+5)
    pdf.cell(30, 5, clean_text("Resultados (%)"), 1, 0, 'C')
    pdf.cell(50, 5, clean_text("Padrões (Min - Max)"), 1, 0, 'C')
    pdf.set_xy(x, cy+10)

    # Valores (Busca agressiva)
    pdf.set_font('Arial', '', 8)
    
    kg_cr = formatar_numero(buscar_valor(dados, ["Retenção Cromo (Kg/m³)", "Retenção Cromo", "Cromo", "Teor de CrO3"]), True)
    kg_cu = formatar_numero(buscar_valor(dados, ["Retenção Cobre (Kg/m³)", "Retenção Cobre", "Cobre", "Teor de CuO"]), True)
    kg_as = formatar_numero(buscar_valor(dados, ["Retenção Arsênio (Kg/m³)", "Retenção Arsênio", "Arsenio", "Teor de As2O5"]), True)

    # Porcentagens
    pc_cr = formatar_numero(buscar_valor(dados, ["Balanço Cromo (%)", "Balanço Cromo", "Cromo %", "Bal. Cr"]), True)
    pc_cu = formatar_numero(buscar_valor(dados, ["Balanço Cobre (%)", "Balanço Cobre", "Cobre %", "Bal. Cu"]), True)
    pc_as = formatar_numero(buscar_valor(dados, ["Balanço Arsênio (%)", "Balanço Arsênio", "Arsenio %", "Bal. As"]), True)
    
    # Total
    val_tot = formatar_numero(buscar_valor(dados, ["Soma Concentração (%)", "Retenção Total", "Total"]), True)
    if val_tot == 0: val_tot = kg_cr + kg_cu + kg_as # Soma se vazio

    def row(n, kg, pc, mn, mx, mt=""):
        pdf.cell(40, 6, clean_text(n), 1, 0, 'L')
        pdf.cell(30, 6, float_para_str(kg), 1, 0, 'C')
        pdf.cell(30, 6, float_para_str(pc) if pc > 0 else "-", 1, 0, 'C')
        pdf.cell(25, 6, mn, 1, 0, 'C')
        pdf.cell(25, 6, mx, 1, 0, 'C')
        pdf.cell(40, 6, clean_text(mt), 1, 1, 'C')

    row("Cromo (CrO3)", kg_cr, pc_cr, "41,8", "53,2", "Metodo UFV 01")
    row("Cobre (CuO)", kg_cu, pc_cu, "15,2", "22,8", "")
    row("Arsenio (As2O5)", kg_as, pc_as, "27,3", "40,7", "")

    # Linha Total
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, 6, clean_text("RETENÇÃO TOTAL"), 1, 0, 'L')
    pdf.cell(30, 6, float_para_str(val_tot), 1, 0, 'C')
    # Cálculo % Total (Soma das %)
    tot_pc = pc_cr + pc_cu + pc_as
    pdf.cell(30, 6, float_para_str(tot_pc) if tot_pc > 0 else "-", 1, 0, 'C')
    pdf.cell(90, 6, clean_text("Nota: Resultados restritos as amostras"), 1, 1, 'C')

    # --- 5. PENETRAÇÃO (Correção Grau) ---
    y = pdf.get_y() + 5
    pdf.set_y(y); pdf.set_font('Arial', 'B', 9)
    pdf.cell(190, 6, clean_text("RESULTADOS DE PENETRAÇÃO"), 0, 1, 'C')
    y += 7
    
    # Busca Grau (Tenta várias colunas)
    grau = buscar_valor(dados, ["Grau de penetração", "Grau", "Nota", "G", "Classificação"])
    tipo = buscar_valor(dados, ["Descrição Grau", "Tipo", "Tipo Penetracao"])
    
    pdf.campo_form("Grau", grau, 10, y, 30, align='C')
    pdf.campo_form("Tipo", tipo, 45, y, 50, align='C')
    
    desc = buscar_valor(dados, ["Descrição Penetração", "Descrição", "Obs Penetracao"])
    pdf.set_xy(100, y)
    pdf.set_font('Arial', '', 6)
    pdf.cell(90, 3, clean_text("Descrição"), 0, 0, 'L')
    pdf.set_xy(100, y+3)
    pdf.set_font('Arial', '', 8)
    pdf.rect(100, y+3, 100, 12)
    pdf.multi_cell(100, 4, clean_text(desc), 0, 'L')

    # --- 6. OBSERVAÇÕES ---
    y += 20
    obs = buscar_valor(dados, ["Observação: Analista de Controle de Qualidade", "Observação", "Obs"])
    if obs:
        pdf.set_y(y)
        pdf.campo_form("Observações", obs, 10, y, 190, h=12, multi=True)

    # --- 7. ASSINATURA ---
    pdf.set_y(-35)
    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 5, clean_text("Dr. Vinicius Resende de Castro - Supervisor do laboratório"), 0, 1, 'C')

    return pdf.output(dest='S').encode('latin-1')

# --- CONEXÃO ---
def carregar(aba):
    try:
        s = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        c = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(c, s)
        client = gspread.authorize(creds)
        sh = client.open(NOME_PLANILHA_GOOGLE)
        df = pd.DataFrame(sh.worksheet(aba).get_all_records())
        if not df.empty: df.columns = df.columns.str.strip()
        return df
    except: return pd.DataFrame()

def salvar(df, aba):
    try:
        s = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        c = dict(st.secrets["gcp_service_account"])
        client = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_dict(c, s))
        ws = client.open(NOME_PLANILHA_GOOGLE).worksheet(aba)
        ws.clear()
        d = df.drop(columns=["Selecionar"]) if "Selecionar" in df.columns else df
        ws.update([d.columns.values.tolist()] + d.astype(str).values.tolist())
        st.toast("Salvo!")
    except: st.error("Erro Salvar")

# --- MAIN ---
def main():
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    
    if not st.session_state['logado']:
        c1,c2,c3 = st.columns([1,2,1])
        with c2:
            st.title("🔐 Login")
            u = st.text_input("User"); p = st.text_input("Pass", type="password")
            if st.button("Entrar", type="primary"):
                try:
                    s = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                    c = dict(st.secrets["gcp_service_account"])
                    cl = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_dict(c, s))
                    users = pd.DataFrame(cl.open(NOME_PLANILHA_GOOGLE).worksheet("Usuarios").get_all_records()).astype(str)
                    match = users[(users['Usuario']==u) & (users['Senha']==p)]
                    if not match.empty:
                        st.session_state.update({'logado':True, 'tipo':match.iloc[0]['Tipo'], 'user':u})
                        st.rerun()
                    else: st.error("Errado")
                except: st.error("Erro Conexão")
        return

    # Logado
    tipo = st.session_state['tipo']
    st.sidebar.info(f"👤 {st.session_state['user']} ({tipo})")
    if st.sidebar.button("Sair"): st.session_state['logado']=False; st.rerun()

    # DEBUG: Espiar Colunas
    with st.sidebar.expander("🕵️ Espiar Colunas (Se falhar)"):
        df_x = carregar("Madeira")
        if not df_x.empty: st.write(list(df_x.columns))

    st.title("🌲 Sistema UFV")
    menu = st.sidebar.radio("Ir para", ["Madeira Tratada", "Solução", "Dashboard"])

    if menu == "Madeira Tratada":
        df = carregar("Madeira")
        if not df.empty:
            if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
            
            if tipo == "LPM":
                df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
                if st.button("💾 Salvar"): salvar(df, "Madeira"); st.rerun()
            else:
                cfg = {c: st.column_config.Column(disabled=True) for c in df.columns if c!="Selecionar"}
                cfg["Selecionar"] = st.column_config.CheckboxColumn(disabled=False)
                df = st.data_editor(df, column_config=cfg, use_container_width=True)

            sel = df[df["Selecionar"]==True]
            if st.button("📄 GERAR PDF", type="primary"):
                if not sel.empty:
                    try:
                        linha = sel.iloc[0].to_dict()
                        pdf = gerar_pdf_nativo(linha)
                        nm = f"{linha.get('Código UFV','Relatorio')}.pdf"
                        st.download_button(f"⬇️ {nm}", pdf, nm, "application/pdf")
                    except Exception as e: st.error(f"Erro: {e}")
                else: st.warning("Selecione um item.")

    elif menu == "Solução":
        df = carregar("Solucao")
        if not df.empty: st.dataframe(df)

    elif menu == "Dashboard":
        df = carregar("Madeira")
        if not df.empty:
            import plotly.express as px
            st.plotly_chart(px.bar(df['Nome do Cliente'].value_counts()))

if __name__ == "__main__":
    main()
