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

# ✅ IDs CONFIGURADOS (Conta Nova)
ID_ARQUIVO_EXCEL = "1L0qTK6oy2axnCSlLadoyk9q5fExSnA6v"
ID_PASTA_RAIZ = "1nZtJjVZUVx65GtjnmpTn5Hw_eZOXwpIY"

# --- CONEXÃO DRIVE ---
def get_drive_service():
    scope = ["https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
    return build('drive', 'v3', credentials=creds)

# --- GERENCIADOR DE PASTAS ---
def get_or_create_folder(service, folder_name, parent_id):
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and '{parent_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    items = results.get('files', [])
    if items: return items[0]['id']
    else:
        metadata = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
        return service.files().create(body=metadata, fields='id', supportsAllDrives=True).execute().get('id')

def salvar_pdf_organizado(pdf_bytes, nome_arquivo, data_entrada_raw):
    try:
        if not ID_PASTA_RAIZ: st.error("⚠️ ID da pasta não configurado."); return
        service = get_drive_service()
        meses = {1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'}
        data_obj = datetime.now()
        if isinstance(data_entrada_raw, datetime): data_obj = data_entrada_raw
        elif data_entrada_raw and str(data_entrada_raw).strip() not in ["", "-"]:
            try:
                v_str = str(data_entrada_raw).strip().split(" ")[0]
                for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"]:
                    try: data_obj = datetime.strptime(v_str, fmt); break
                    except: continue
            except: pass
        
        ano_str = str(data_obj.year); mes_str = meses[data_obj.month]
        ano_id = get_or_create_folder(service, ano_str, ID_PASTA_RAIZ)
        mes_id = get_or_create_folder(service, mes_str, ano_id)
        
        nome_limpo = nome_arquivo.replace("/", "-").replace("\\", "-")
        media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype='application/pdf', resumable=False)
        metadata = {'name': nome_limpo, 'parents': [mes_id]}
        service.files().create(body=metadata, media_body=media, fields='id', supportsAllDrives=True).execute()
        st.balloons(); st.toast(f"Salvo: {ano_str}/{mes_str}", icon="✅"); st.success(f"Arquivo **{nome_limpo}** salvo em: **{ano_str} > {mes_str}**")
    except Exception as e: st.error(f"Erro ao salvar: {e}")

# --- MATEMÁTICA E DADOS ---
def corrigir_numero_individual(v):
    try:
        if pd.isna(v) or v=="": return 0.0
        val = float(str(v).replace(",", "."))
        if val > 1000: val /= 100.0
        if val > 100: val /= 100.0
        return val
    except: return v

def corrigir_valores_dataframe(df):
    cols = ['Retenção', 'Retenção Cromo', 'Retenção Cobre', 'Retenção Arsênio', 'Balanço Cromo', 'Balanço Cobre', 'Balanço Arsênio', 'Soma Concentração', 'Balanço Total']
    for col in df.columns:
        for alvo in cols:
            if alvo.lower() in col.lower(): df[col] = df[col].apply(corrigir_numero_individual)
    return df

@st.cache_data(ttl=60)
def carregar_excel_drive(aba_nome):
    try:
        service = get_drive_service()
        request = service.files().get_media(fileId=ID_ARQUIVO_EXCEL)
        df = pd.read_excel(io.BytesIO(request.execute()), sheet_name=aba_nome)
        df.columns = df.columns.str.strip()
        return corrigir_valores_dataframe(df)
    except Exception as e: st.error(f"Erro Excel: {e}"); return pd.DataFrame()

def salvar_excel_drive(df, aba_nome):
    try:
        service = get_drive_service()
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer: df.to_excel(writer, sheet_name=aba_nome, index=False)
        buf.seek(0)
        media = MediaIoBaseUpload(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        service.files().update(fileId=ID_ARQUIVO_EXCEL, media_body=media, supportsAllDrives=True).execute()
        st.toast("Salvo!", icon="💾"); st.cache_data.clear()
    except Exception as e: st.error(f"Erro Salvar: {e}")

# --- HELPERS ---
def clean_text(text): return str(text).encode('latin-1', 'replace').decode('latin-1') if not pd.isna(text) else ""
def fmt_num(v): 
    try: return "{:,.2f}".format(float(str(v).replace(",", "."))).replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(v)
def fmt_date(v):
    if not v: return "-"
    if isinstance(v, datetime): return v.strftime("%d/%m/%Y")
    s = str(v).strip().split(" ")[0]
    for f in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"]:
        try: return datetime.strptime(s, f).strftime("%d/%m/%Y")
        except: continue
    return s
def get_val(d, keys):
    dn = {k.strip().lower(): v for k, v in d.items()}
    for k in keys:
        k = k.strip().lower()
        if k in dn and str(dn[k]).strip()!="": return dn[k]
        for c in dn:
            if k in c and str(dn[c]).strip()!="": return dn[c]
    return ""

# --- CLASSE PDF (RENOVADA) ---
class RPDF(FPDF):
    def header(self):
        # Logos com tamanhos ajustados para parecerem iguais
        if os.path.exists("logo_ufv.png"): self.image("logo_ufv.png", 10, 8, 25)
        if os.path.exists("logo_montana.png"): self.image("logo_montana.png", 175, 8, 25) # Mesmo tamanho
        self.set_y(12); self.set_font('Arial','B',14); self.cell(0,10,clean_text('Relatório de Ensaio'),0,1,'C')
    
    def footer(self):
        self.set_y(-15); self.set_font('Arial','I',6); self.cell(0,10,clean_text(f'Página {self.page_no()}'),0,0,'C')
    
    # Novo Field com Negrito no Label
    def field(self, label, valor, x, y, w, h=6, align='L', multi=False):
        self.set_xy(x, y)
        self.set_font('Arial', 'B', 8) # Label agora é Bold e Maior (8)
        self.cell(w, 3, clean_text(label), 0, 0, 'L')
        
        self.set_xy(x, y+3)
        self.set_font('Arial', '', 8) # Valor normal
        if multi: 
            self.rect(x, y+3, w, h)
            self.multi_cell(w, 4, clean_text(valor), 0, align)
        else: 
            self.cell(w, h, clean_text(valor), 1, 0, align)

def gerar_pdf(d):
    pdf = RPDF(); pdf.add_page(); pdf.set_auto_page_break(auto=True, margin=15)
    
    # 1. Cabeçalho
    y = 30
    pdf.field("Data de Entrada", fmt_date(get_val(d, ["Data de entrada", "Entrada"])), 10, y, 40, align='C')
    pdf.field("Número ID", clean_text(get_val(d, ["Código UFV", "ID"])), 150, y-5, 50, align='C')
    pdf.field("Data de Emissão", fmt_date(get_val(d, ["Data de Registro", "Fim da análise"])), 150, y+8, 50, align='C')

    # 2. Cliente
    y += 20; pdf.set_y(y); pdf.set_font('Arial', 'B', 9); pdf.cell(0, 5, clean_text("DADOS DO CLIENTE"), 0, 1, 'L')
    y += 6; pdf.field("Cliente", get_val(d, ["Nome do Cliente"]), 10, y, 190)
    y += 11; pdf.field("Cidade/UF", f"{get_val(d,['Cidade'])}/{get_val(d,['Estado'])}", 10, y, 90)
    pdf.field("E-mail", get_val(d, ["E-mail"]), 105, y, 95)

    # 3. Amostra
    y += 15; pdf.set_y(y); pdf.set_font('Arial', 'B', 9); pdf.cell(0, 5, clean_text("IDENTIFICAÇÃO DA AMOSTRA"), 0, 1, 'L')
    y += 6; pdf.field("Ref. Cliente", get_val(d, ["Indentificação de Amostra"]), 10, y, 190)
    y += 11; pdf.field("Madeira", get_val(d, ["Madeira"]), 10, y, 90)
    pdf.field("Produto", get_val(d, ["Produto"]), 105, y, 95)
    y += 11; pdf.field("Aplicação", get_val(d, ["Aplicação"]), 10, y, 60)
    pdf.field("Norma ABNT", get_val(d, ["Norma"]), 75, y, 60)
    pdf.field("Retenção Esp.", fmt_num(get_val(d, ["Retenção"])), 140, y, 60, align='C')

    # 4. Química (Tabela Complexa)
    y += 20; pdf.set_y(y); pdf.set_font('Arial', 'B', 9); pdf.cell(190, 6, clean_text("RESULTADOS DE RETENÇÃO"), 1, 1, 'C')
    
    # Cabeçalhos da Tabela
    pdf.set_font('Arial', 'B', 7); x=10; cy=pdf.get_y()
    pdf.cell(40, 10, clean_text("Ingredientes ativos"), 1, 0, 'C')
    pdf.cell(30, 10, clean_text("Resultado (kg/m3)"), 1, 0, 'C')
    # Bloco "Balanceamento"
    pdf.cell(80, 5, clean_text("Balanceamento químico"), 1, 0, 'C')
    
    # ** MÉTODO MESCLADO **
    # Salva posição antes de desenhar
    pdf.set_xy(x+150, cy) 
    # Desenha célula gigante (Altura 10 do cabeçalho + 18 das 3 linhas = 28? Não, vamos alinhar só o titulo aqui)
    pdf.cell(40, 10, clean_text("Método"), 1, 0, 'C')
    
    # Sub-cabeçalhos do Balanceamento
    pdf.set_xy(x+70, cy+5)
    pdf.cell(30, 5, clean_text("Resultados (%)"), 1, 0, 'C')
    pdf.cell(50, 5, clean_text("Padrões"), 1, 0, 'C')
    
    # Posiciona para os dados
    pdf.set_xy(x, cy+10)
    y_dados_inicio = cy+10

    # Valores Químicos
    kg_cr=fmt_num(get_val(d,["Retenção Cromo","Cromo"]))
    kg_cu=fmt_num(get_val(d,["Retenção Cobre","Cobre"]))
    kg_as=fmt_num(get_val(d,["Retenção Arsênio","Arsenio"]))
    
    pc_cr=fmt_num(get_val(d,["Balanço Cromo","Cromo %"]))
    pc_cu=fmt_num(get_val(d,["Balanço Cobre","Cobre %"]))
    pc_as=fmt_num(get_val(d,["Balanço Arsênio","Arsenio %"]))

    pdf.set_font('Arial', '', 8)
    
    # Função para desenhar linha (sem a coluna método)
    def row_data(n, k, p, mn, mx):
        pdf.cell(40, 6, clean_text(n), 1, 0, 'L')
        pdf.cell(30, 6, k, 1, 0, 'C')
        pdf.cell(30, 6, p, 1, 0, 'C')
        pdf.cell(25, 6, mn, 1, 0, 'C')
        pdf.cell(25, 6, mx, 1, 0, 'C')
        # Pula a coluna método aqui, vamos desenhar ela separada
        pdf.set_x(pdf.get_x() + 40) 
        pdf.ln(6)

    # Desenha célula gigante do Método (Altura 3 linhas * 6 = 18)
    pdf.set_xy(160, y_dados_inicio) # Coluna Método começa no X=160
    pdf.cell(40, 18, clean_text("Metodo UFV 01"), 1, 0, 'C')
    
    # Volta para desenhar os dados linha a linha
    pdf.set_xy(10, y_dados_inicio)
    
    # Nomes com "falsa" subscrição (usando texto normal pois PDF padrão limita)
    row_data("Teor de CrO3 (Cromo)", kg_cr, pc_cr, "41,8", "53,2")
    row_data("Teor de CuO (Cobre)", kg_cu, pc_cu, "15,2", "22,8")
    row_data("Teor de As2O5 (Arsênio)", kg_as, pc_as, "27,3", "40,7")

    # Linha Total
    try: tot = float(kg_cr.replace(",",".")) + float(kg_cu.replace(",",".")) + float(kg_as.replace(",","."))
    except: tot=0
    
    # Busca o valor da coluna BM (Balanço Total)
    bm_total = fmt_num(get_val(d, ["Balanço Total", "Balanço Total %", "BM"]))
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, 6, clean_text("RETENÇÃO TOTAL"), 1, 0, 'L')
    pdf.cell(30, 6, fmt_num(tot), 1, 0, 'C')
    pdf.cell(30, 6, bm_total, 1, 0, 'C') # Valor da Coluna BM
    pdf.cell(90, 6, clean_text("Nota: Resultados restritos as amostras"), 1, 1, 'C')

    # 5. Penetração
    y = pdf.get_y() + 5; pdf.set_y(y); pdf.set_font('Arial', 'B', 9); pdf.cell(190, 6, clean_text("RESULTADOS DE PENETRAÇÃO"), 0, 1, 'C')
    y += 7
    
    # CORREÇÃO DO TIPO: Busca Descrição do Grau (Coluna AB)
    tipo_correto = get_val(d, ["Descrição do Grau", "Descricao do Grau", "Grau Descricao"])
    
    pdf.field("Grau", get_val(d, ["Grau"]), 10, y, 30, align='C')
    pdf.field("Tipo", tipo_correto, 45, y, 50, align='C') # Agora puxa da coluna certa
    
    pdf.set_xy(100, y); pdf.set_font('Arial', 'B', 8); pdf.cell(90, 3, clean_text("Descrição"), 0, 0, 'L')
    pdf.set_xy(100, y+3); pdf.set_font('Arial', '', 8); pdf.rect(100, y+3, 100, 12)
    pdf.multi_cell(100, 4, clean_text(get_val(d, ["Descrição Penetração", "Descricao"])), 0, 'L')

    # 6. Observações
    y += 20; obs = get_val(d, ["Observação", "Obs"])
    # Observação em Negrito no Label
    if obs: pdf.set_y(y); pdf.field("Observações", obs, 10, y, 190, 12, 'L', True)
    
    pdf.set_y(-35); pdf.set_font('Arial', '', 9); pdf.cell(0, 5, clean_text("Dr. Vinicius Resende de Castro - Supervisor do laboratório"), 0, 1, 'C')
    return pdf.output(dest='S').encode('latin-1')

# --- MAIN ---
def main():
    if 'logado' not in st.session_state: st.session_state['logado']=False
    if not st.session_state['logado']:
        c1,c2,c3=st.columns([1,2,1])
        with c2:
            st.title("🔐 Login"); u=st.text_input("User"); p=st.text_input("Pass",type="password")
            if st.button("Entrar",type="primary"):
                if (u=="admin" and p=="admin") or (u=="montana" and p=="montana"): st.session_state.update({'logado':True,'tipo':u.capitalize(),'user':u}); st.rerun()
                else: st.error("Erro")
        return
    
    st.sidebar.info(f"👤 {st.session_state['user']}"); 
    if st.sidebar.button("Sair"): st.session_state['logado']=False; st.rerun()
    st.title("🌲 Sistema Controle UFV")
    menu=st.sidebar.radio("Menu",["Madeira Tratada","Solução"])
    
    if menu=="Madeira Tratada":
        df=carregar_excel_drive("Madeira Tratada")
        if not df.empty:
            if "Selecionar" not in df.columns: df.insert(0,"Selecionar",False)
            df=st.data_editor(df, num_rows="dynamic", use_container_width=True)
            if st.session_state['tipo']=="Lpm":
                if st.button("💾 SALVAR DADOS NO EXCEL", type="primary"): salvar_excel_drive(df,"Madeira Tratada")
            
            sel=df[df["Selecionar"]==True]
            st.divider(); 
            
            if not sel.empty:
                st.subheader("📄 Gerar Relatório")
                try:
                    l=sel.iloc[0].to_dict()
                    pdf_bytes=gerar_pdf(l)
                    nome_arquivo = f"{l.get('Código UFV','Relatorio')}.pdf"
                    
                    c_down, c_cloud = st.columns(2)
                    
                    with c_down:
                        st.download_button("⬇️ BAIXAR PDF (PC)", pdf_bytes, nome_arquivo, "application/pdf", type="primary")
                    
                    with c_cloud:
                        if st.button("☁️ SALVAR NO DRIVE COMPARTILHADO"):
                            salvar_pdf_organizado(pdf_bytes, nome_arquivo, get_val(l,["Data de entrada"]))
                            
                except Exception as e: st.error(f"Erro na geração: {e}")
            else: st.warning("Selecione um item para gerar PDF.")
    
    elif menu=="Solução":
        df=carregar_excel_drive("Solução Preservativa")
        if not df.empty: st.dataframe(df)

if __name__ == "__main__":
    main()
