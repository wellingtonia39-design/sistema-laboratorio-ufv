import streamlit as st
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from oauth2client.service_account import ServiceAccountCredentials
from fpdf import FPDF
import io
import os
from datetime import datetime

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Sistema Controle UFV", layout="wide", page_icon="🌲")
NOME_ARQUIVO_EXCEL = "Planilha controle UFV.xlsx"

# --- CONEXÃO DRIVE ---
def get_drive_service():
    scope = ["https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
    return build('drive', 'v3', credentials=creds)

def encontrar_id_arquivo(service, nome_arquivo):
    query = f"name = '{nome_arquivo}' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    return items[0]['id'] if items else None

# --- INTELIGÊNCIA DE CORREÇÃO ---
def corrigir_valores_dataframe(df):
    cols_quimicas = [
        'Retenção', 'Retenção Cromo', 'Retenção Cobre', 'Retenção Arsênio',
        'Balanço Cromo', 'Balanço Cobre', 'Balanço Arsênio', 'Soma Concentração'
    ]
    for col in df.columns:
        for alvo in cols_quimicas:
            if alvo.lower() in col.lower():
                df[col] = df[col].apply(lambda x: corrigir_numero_individual(x))
    return df

def corrigir_numero_individual(valor):
    try:
        if pd.isna(valor) or valor == "": return 0.0
        v = float(str(valor).replace(",", "."))
        if v > 1000: v /= 100.0
        if v > 100:  v /= 100.0
        return v
    except: return valor

# --- OPERAÇÕES DE ARQUIVO ---
@st.cache_data(ttl=60)
def carregar_excel_drive(aba_nome):
    try:
        service = get_drive_service()
        file_id = encontrar_id_arquivo(service, NOME_ARQUIVO_EXCEL)
        if not file_id: 
            st.error("Arquivo não encontrado no Drive.")
            return pd.DataFrame()

        request = service.files().get_media(fileId=file_id)
        arquivo_bytes = io.BytesIO(request.execute())
        
        df = pd.read_excel(arquivo_bytes, sheet_name=aba_nome)
        df.columns = df.columns.str.strip()
        df = corrigir_valores_dataframe(df)
        return df
    except Exception as e:
        st.error(f"Erro ao carregar: {e}")
        return pd.DataFrame()

def salvar_excel_drive(df, aba_nome):
    try:
        service = get_drive_service()
        file_id = encontrar_id_arquivo(service, NOME_ARQUIVO_EXCEL)
        if not file_id: return
        
        buffer_novo = io.BytesIO()
        with pd.ExcelWriter(buffer_novo, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=aba_nome, index=False)
            
        buffer_novo.seek(0)
        media = MediaIoBaseUpload(buffer_novo, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        service.files().update(fileId=file_id, media_body=media).execute()
        
        st.toast("Salvo com sucesso!", icon="✅")
        st.cache_data.clear()
    except Exception as e: st.error(f"Erro ao salvar: {e}")

# --- FORMATADORES ---
def clean_text(text): return str(text).encode('latin-1', 'replace').decode('latin-1') if not pd.isna(text) else ""

def formatar_numero_pdf(valor):
    try: return "{:,.2f}".format(float(str(valor).replace(",", "."))).replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(valor)

def formatar_data(valor):
    if not valor: return "-"
    if isinstance(valor, datetime): return valor.strftime("%d/%m/%Y")
    v_str = str(valor).strip().split(" ")[0]
    for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y"]:
        try: return datetime.strptime(v_str, fmt).strftime("%d/%m/%Y")
        except: continue
    return v_str

def buscar_valor(dados, chaves):
    dados_norm = {k.strip().lower(): v for k, v in dados.items()}
    for chave in chaves:
        k = chave.strip().lower()
        if k in dados_norm and str(dados_norm[k]).strip() != "": return dados_norm[k]
        for col in dados_norm:
            if k in col and str(dados_norm[col]).strip() != "": return dados_norm[col]
    return ""

class RelatorioPDF(FPDF):
    def header(self):
        if os.path.exists("logo_ufv.png"): self.image("logo_ufv.png", 10, 8, 25)
        if os.path.exists("logo_montana.png"): self.image("logo_montana.png", 175, 8, 25)
        self.set_y(12); self.set_font('Arial', 'B', 14); self.cell(0, 10, clean_text('Relatório de Ensaio'), 0, 1, 'C')
    def footer(self):
        self.set_y(-15); self.set_font('Arial', 'I', 6); self.cell(0, 10, clean_text(f'Página {self.page_no()}'), 0, 0, 'C')
    
    # CORREÇÃO AQUI: Garante que align seja passado corretamente
    def campo_form(self, label, valor, x, y, w, h=6, align='L', multi=False):
        self.set_xy(x, y); self.set_font('Arial', '', 6); self.cell(w, 3, clean_text(label), 0, 0, 'L')
        self.set_xy(x, y+3); self.set_font('Arial', '', 8)
        if multi: self.rect(x, y+3, w, h); self.multi_cell(w, 4, clean_text(valor), 0, align)
        else: self.cell(w, h, clean_text(valor), 1, 0, align)

def gerar_pdf_nativo(dados):
    pdf = RelatorioPDF(); pdf.add_page(); pdf.set_auto_page_break(auto=True, margin=15)
    
    # 1. Cabeçalho (CORRIGIDO: align='C')
    y = 30
    pdf.campo_form("Data de Entrada", formatar_data(buscar_valor(dados, ["Data de entrada", "Entrada"])), 10, y, 40, align='C')
    pdf.campo_form("Número ID", clean_text(buscar_valor(dados, ["Código UFV", "ID"])), 150, y-5, 50, align='C')
    pdf.campo_form("Data de Emissão", formatar_data(buscar_valor(dados, ["Data de Registro", "Fim da análise"])), 150, y+8, 50, align='C')

    # 2. Cliente
    y += 20; pdf.set_y(y); pdf.set_font('Arial', 'B', 9); pdf.cell(0, 5, clean_text("DADOS DO CLIENTE"), 0, 1, 'L')
    y += 6; pdf.campo_form("Cliente", buscar_valor(dados, ["Nome do Cliente"]), 10, y, 190)
    y += 11; pdf.campo_form("Cidade/UF", f"{buscar_valor(dados,['Cidade'])}/{buscar_valor(dados,['Estado'])}", 10, y, 90)
    pdf.campo_form("E-mail", buscar_valor(dados, ["E-mail"]), 105, y, 95)

    # 3. Amostra (CORRIGIDO: align='C')
    y += 15; pdf.set_y(y); pdf.set_font('Arial', 'B', 9); pdf.cell(0, 5, clean_text("IDENTIFICAÇÃO DA AMOSTRA"), 0, 1, 'L')
    y += 6; pdf.campo_form("Ref. Cliente", buscar_valor(dados, ["Indentificação de Amostra"]), 10, y, 190)
    y += 11; pdf.campo_form("Madeira", buscar_valor(dados, ["Madeira"]), 10, y, 90)
    pdf.campo_form("Produto", buscar_valor(dados, ["Produto"]), 105, y, 95)
    y += 11; pdf.campo_form("Aplicação", buscar_valor(dados, ["Aplicação"]), 10, y, 60)
    pdf.campo_form("Norma ABNT", buscar_valor(dados, ["Norma"]), 75, y, 60)
    pdf.campo_form("Retenção Esp.", formatar_numero_pdf(buscar_valor(dados, ["Retenção"])), 140, y, 60, align='C')

    # 4. Química
    y += 20; pdf.set_y(y); pdf.set_font('Arial', 'B', 9); pdf.cell(190, 6, clean_text("RESULTADOS DE RETENÇÃO"), 1, 1, 'C')
    pdf.set_font('Arial', 'B', 7); x=10; cy=pdf.get_y()
    pdf.cell(40, 10, clean_text("Ingredientes ativos"), 1, 0, 'C'); pdf.cell(30, 10, clean_text("Resultado (kg/m3)"), 1, 0, 'C')
    pdf.cell(80, 5, clean_text("Balanceamento químico"), 1, 0, 'C'); pdf.set_xy(x+150, cy); pdf.cell(40, 10, clean_text("Método"), 1, 0, 'C')
    pdf.set_xy(x+70, cy+5); pdf.cell(30, 5, clean_text("Resultados (%)"), 1, 0, 'C'); pdf.cell(50, 5, clean_text("Padrões"), 1, 0, 'C'); pdf.set_xy(x, cy+10)

    def pegar_num(keys): return formatar_numero_pdf(buscar_valor(dados, keys))
    kg_cr = pegar_num(["Retenção Cromo", "Cromo"]); kg_cu = pegar_num(["Retenção Cobre", "Cobre"]); kg_as = pegar_num(["Retenção Arsênio", "Arsenio"])
    pc_cr = pegar_num(["Balanço Cromo", "Cromo %"]); pc_cu = pegar_num(["Balanço Cobre", "Cobre %"]); pc_as = pegar_num(["Balanço Arsênio", "Arsenio %"])

    pdf.set_font('Arial', '', 8)
    def row(n, k, p, mn, mx, mt=""):
        pdf.cell(40, 6, clean_text(n), 1, 0, 'L'); pdf.cell(30, 6, k, 1, 0, 'C'); pdf.cell(30, 6, p, 1, 0, 'C')
        pdf.cell(25, 6, mn, 1, 0, 'C'); pdf.cell(25, 6, mx, 1, 0, 'C'); pdf.cell(40, 6, clean_text(mt), 1, 1, 'C')
    row("Cromo (CrO3)", kg_cr, pc_cr, "41,8", "53,2", "Metodo UFV 01")
    row("Cobre (CuO)", kg_cu, pc_cu, "15,2", "22,8", "")
    row("Arsenio (As2O5)", kg_as, pc_as, "27,3", "40,7", "")

    try: val_tot = float(kg_cr.replace(",", ".")) + float(kg_cu.replace(",", ".")) + float(kg_as.replace(",", "."))
    except: val_tot = 0
    pdf.set_font('Arial', 'B', 8); pdf.cell(40, 6, clean_text("RETENÇÃO TOTAL"), 1, 0, 'L')
    pdf.cell(30, 6, formatar_numero_pdf(val_tot), 1, 0, 'C'); pdf.cell(30, 6, "-", 1, 0, 'C'); pdf.cell(90, 6, clean_text("Nota: Resultados restritos as amostras"), 1, 1, 'C')

    # 5. Penetração (CORRIGIDO: align='C')
    y = pdf.get_y() + 5; pdf.set_y(y); pdf.set_font('Arial', 'B', 9); pdf.cell(190, 6, clean_text("RESULTADOS DE PENETRAÇÃO"), 0, 1, 'C')
    y += 7; pdf.campo_form("Grau", buscar_valor(dados, ["Grau"]), 10, y, 30, align='C'); pdf.campo_form("Tipo", buscar_valor(dados, ["Tipo"]), 45, y, 50, align='C')
    pdf.set_xy(100, y); pdf.set_font('Arial', '', 6); pdf.cell(90, 3, clean_text("Descrição"), 0, 0, 'L')
    pdf.set_xy(100, y+3); pdf.set_font('Arial', '', 8); pdf.rect(100, y+3, 100, 12)
    pdf.multi_cell(100, 4, clean_text(buscar_valor(dados, ["Descrição Penetração", "Descricao"])), 0, 'L')

    # 6. Observações e Assinatura
    y += 20; obs = buscar_valor(dados, ["Observação", "Obs"])
    if obs: pdf.set_y(y); pdf.campo_form("Observações", obs, 10, y, 190, 12, 'L', True)
    pdf.set_y(-35); pdf.set_font('Arial', '', 9); pdf.cell(0, 5, clean_text("Dr. Vinicius Resende de Castro - Supervisor do laboratório"), 0, 1, 'C')
    return pdf.output(dest='S').encode('latin-1')

# --- MAIN ---
def main():
    if 'logado' not in st.session_state: st.session_state['logado'] = False
    if not st.session_state['logado']:
        c1,c2,c3 = st.columns([1,2,1])
        with c2:
            st.title("🔐 Login UFV")
            u = st.text_input("Usuário"); p = st.text_input("Senha", type="password")
            if st.button("Entrar", type="primary"):
                if u == "admin" and p == "admin": 
                    st.session_state.update({'logado':True, 'tipo':'LPM', 'user':u}); st.rerun()
                elif u == "montana" and p == "montana":
                    st.session_state.update({'logado':True, 'tipo':'Montana', 'user':u}); st.rerun()
                else: st.error("Acesso Negado")
        return

    tipo = st.session_state['tipo']
    st.sidebar.info(f"👤 {st.session_state['user']} ({tipo})")
    if st.sidebar.button("Sair"): st.session_state['logado'] = False; st.rerun()

    st.title("🌲 Sistema Controle UFV")
    menu = st.sidebar.radio("Menu", ["Madeira Tratada", "Solução"])

    if menu == "Madeira Tratada":
        df = carregar_excel_drive("Madeira Tratada")
        if not df.empty:
            if "Selecionar" not in df.columns: df.insert(0, "Selecionar", False)
            if tipo == "LPM":
                df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
                if st.button("💾 SALVAR DADOS", type="primary"): salvar_excel_drive(df, "Madeira Tratada")
            else:
                cfg = {c: st.column_config.Column(disabled=True) for c in df.columns if c!="Selecionar"}
                cfg["Selecionar"] = st.column_config.CheckboxColumn(disabled=False)
                df = st.data_editor(df, column_config=cfg, use_container_width=True)

            sel = df[df["Selecionar"]==True]
            if st.button("📄 GERAR RELATÓRIO PDF", type="primary"):
                if not sel.empty:
                    try:
                        linha = sel.iloc[0].to_dict()
                        pdf = gerar_pdf_nativo(linha)
                        nome = f"{linha.get('Código UFV','Relatorio')}.pdf"
                        st.download_button(f"⬇️ BAIXAR {nome}", pdf, nome, "application/pdf")
                    except Exception as e: st.error(f"Erro no PDF: {e}")
                else: st.warning("Selecione uma amostra.")
    
    elif menu == "Solução":
        df = carregar_excel_drive("Solução Preservativa")
        if not df.empty: st.dataframe(df, use_container_width=True)

if __name__ == "__main__":
    main()
