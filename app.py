import streamlit as st
import pandas as pd
import requests
import json
import time
import plotly.graph_objects as go # Gráfico manual (robusto)
import plotly.express as px
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Robô Investidor Pro 10.0", layout="wide", page_icon="🎯")

# --- CONSTANTES ---
NOME_PLANILHA_GOOGLE = "carteira_robo_db"

# --- MAPEAMENTO DE SETORES ---
SETORES = {
    "WEGE3": "Indústria", "VALE3": "Mineração", "PSSA3": "Seguros",
    "ITUB4": "Bancos", "ITSA4": "Bancos", "BBAS3": "Bancos",
    "TAEE11": "Elétrica", "CPLE6": "Elétrica", "EGIE3": "Elétrica",
    "IVVB11": "Dólar/Exterior", "BTLG11": "FII Logística",
    "HGLG11": "FII Logística", "KNCR11": "FII Papel",
    "MXRF11": "FII Híbrido", "XPML11": "FII Shopping",
    "PETR4": "Petróleo", "CURY3": "Construção", "CXSE3": "Seguros",
    "DIRR3": "Construção", "POMO4": "Indústria", "RECV3": "Petróleo"
}

# --- ESTRATÉGIAS ---
CARTEIRAS_PRONTAS = {
    "🏆 Carteira Recomendada IA": {
        "WEGE3": 10, "ITUB4": 15, "VALE3": 10, "TAEE11": 10, "PSSA3": 5, 
        "IVVB11": 20, "HGLG11": 10, "KNCR11": 10, "MXRF11": 10
    },
    "Carteira Dividendos (Rico)": {
        "CURY3": 10, "CXSE3": 10, "DIRR3": 10, "ITSA4": 10, 
        "ITUB4": 10, "PETR4": 10, "POMO4": 10, "RECV3": 10, "VALE3": 10
    },
    "Carteira FIIs (Rico)": {
        "XPML11": 10, "RBRR11": 10, "RBRX11": 9, "XPCI11": 9,
        "BTLG11": 6, "LVBI11": 6, "PCIP11": 6, "PVBI11": 6,
        "KNCR11": 5, "BRCO11": 5, "XPLG11": 4, "KNSC11": 1
    }
}

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

# --- GERENCIAMENTO DE ABAS ---
def pegar_aba_carteira(sh):
    try: return sh.get_worksheet(0)
    except: return sh.add_worksheet(title="Carteira", rows=100, cols=10)

def pegar_aba_config(sh):
    try: return sh.worksheet("Config")
    except: 
        ws = sh.add_worksheet(title="Config", rows=5, cols=5)
        ws.update([["Senha", "MetaMensal"], ["123456", 1000.0]])
        return ws

# --- CARREGAR/SALVAR ---
def carregar_carteira():
    sh = conectar_google_sheets()
    if sh:
        try:
            ws = pegar_aba_carteira(sh)
            dados = ws.get_all_records()
            carteira = {}
            for linha in dados:
                t = linha['Ticker']
                if not t: continue
                
                # Tratamento robusto para valores vazios
                qtde = linha.get('Qtd', 0); qtde = 0 if qtde == '' else int(qtde)
                meta = linha.get('Meta', 0); meta = 0 if meta == '' else int(meta)
                
                try: pm = float(str(linha.get('PM', 0)).replace(',', '.'))
                except: pm = 0.0
                
                try: divs = float(str(linha.get('Divs', 0)).replace(',', '.'))
                except: divs = 0.0
                
                # NOVO CAMPO: TETO (SNIPER)
                try: teto = float(str(linha.get('Teto', 0)).replace(',', '.'))
                except: teto = 0.0
                
                carteira[t] = {'qtde': qtde, 'meta_pct': meta, 'pm': pm, 'divs': divs, 'teto': teto}
            return carteira
        except: return {}
    return {}

def salvar_carteira(carteira):
    sh = conectar_google_sheets()
    if sh:
        ws = pegar_aba_carteira(sh)
        # Atualizando cabeçalho com TETO
        linhas = [["Ticker", "Qtd", "Meta", "PM", "Divs", "Teto"]]
        for t, dados in carteira.items():
            linhas.append([
                t, 
                dados['qtde'], 
                dados['meta_pct'], 
                dados.get('pm', 0.0), 
                dados.get('divs', 0.0),
                dados.get('teto', 0.0)
            ])
        ws.clear()
        ws.update(linhas)

def carregar_config():
    padrao = {"senha": "123456", "meta_mensal": 1000.00}
    sh = conectar_google_sheets()
    if sh:
        try:
            ws = pegar_aba_config(sh)
            dados = ws.get_all_records()
            if dados:
                return {
                    "senha": str(dados[0].get('Senha', '123456')),
                    "meta_mensal": float(str(dados[0].get('MetaMensal', 1000)).replace(',', '.'))
                }
        except: pass
    return padrao

def salvar_config(conf):
    sh = conectar_google_sheets()
    if sh:
        ws = pegar_aba_config(sh)
        ws.clear()
        ws.update([["Senha", "MetaMensal"], [conf['senha'], conf['meta_mensal']]])

# --- COTAÇÃO ---
def obter_preco_atual(ticker):
    if not ticker.endswith(".SA"): ticker = f"{ticker}.SA"
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?interval=1d&range=1d"
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        r = requests.get(url, headers=headers, timeout=3)
        if r.status_code == 200:
            return float(r.json()['chart']['result'][0]['meta']['regularMarketPrice'])
    except: return 0.0
    return 0.0

def obter_setor(ticker):
    return SETORES.get(ticker.replace(".SA","").strip(), "Outros")

def obter_link_investidor10(ticker):
    tipo = "fundos-imobiliarios" if "11" in ticker else "acoes"
    return f"https://investidor10.com.br/{tipo}/{ticker.lower()}/"

# --- CÁLCULO ---
def calcular_compras(df, aporte):
    caixa = aporte
    df = df.copy()
    df['comprar_qtd'] = 0
    df['custo_total'] = 0.0
    if df['meta_pct'].sum() == 0: return df, caixa
    while caixa > 0:
        patr_sim = (df['qtde']*df['preco_atual']).sum() + (df['comprar_qtd']*df['preco_atual']).sum() + caixa
        if patr_sim == 0: break
        df['pct_sim'] = ((df['qtde']+df['comprar_qtd'])*df['preco_atual']/patr_sim)*100
        df['gap'] = df['meta_pct'] - df['pct_sim']
        cand = df[(df['preco_atual'] <= caixa) & (df['gap'] > 0)]
        if cand.empty: break
        melhor = cand['gap'].idxmax()
        preco = df.loc[melhor, 'preco_atual']
        df.loc[melhor, 'comprar_qtd'] += 1
        df.loc[melhor, 'custo_total'] += preco
        caixa -= preco
    return df, caixa

# --- LOGIN ---
def check_password():
    if 'config_cache' not in st.session_state: st.session_state['config_cache'] = carregar_config()
    conf = st.session_state['config_cache']

    if 'logado' not in st.session_state: st.session_state['logado'] = False
    if st.session_state['logado']: return True
    
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown("## 🔐 Acesso Sniper (Cloud)")
        senha = st.text_input("Digite sua senha:", type="password")
        if st.button("Entrar", type="primary"):
            if senha == conf['senha']:
                st.session_state['logado'] = True; st.rerun()
            else: st.error("Senha incorreta!")
    return False

# ================= APP START =================
if check_password():
    if 'config_cache' not in st.session_state: st.session_state['config_cache'] = carregar_config()
    conf = st.session_state['config_cache']
    
    with st.sidebar:
        st.title("🎯 Painel Sniper")
        menu = st.radio("Navegação", ["🏠 Minha Carteira", "⚙️ Configurações"])
        st.divider()
        st.success("Google Drive: Conectado ✅")
        if st.button("🔒 Sair"): st.session_state['logado']=False; st.rerun()
        st.divider()
        modo_live = st.toggle("🔄 Modo Live (60s)")

    if 'carteira_cache' not in st.session_state:
        with st.spinner("Calibrando Mira do Sniper..."):
            st.session_state['carteira_cache'] = carregar_carteira()
    carteira_completa = st.session_state['carteira_cache']

    # ================= TELA: CARTEIRA =================
    if menu == "🏠 Minha Carteira":
        st.title("Minha Carteira (Nuvem ☁️)")

        if not carteira_completa: st.warning("Carteira vazia no Google Sheets.")

        # --- FILTROS ---
        st.markdown("### 🔍 Visualização")
        opcoes_filtro = ["Todas"] + list(CARTEIRAS_PRONTAS.keys()) + ["Personalizados"]
        filtro_selecionado = st.multiselect("Filtrar Carteiras:", opcoes_filtro, default=["Todas"])
        
        carteira_exibicao = {}
        if "Todas" in filtro_selecionado or not filtro_selecionado:
            carteira_exibicao = carteira_completa.copy()
        else:
            tickers_permitidos = []
            todos_prontos = []
            for k, v in CARTEIRAS_PRONTAS.items(): todos_prontos.extend(v.keys())
            for selecao in filtro_selecionado:
                if selecao in CARTEIRAS_PRONTAS: tickers_permitidos.extend(CARTEIRAS_PRONTAS[selecao].keys())
                elif selecao == "Personalizados":
                    for t in carteira_completa.keys():
                        if t not in todos_prontos: tickers_permitidos.append(t)
            for t, dados in carteira_completa.items():
                if t in tickers_permitidos: carteira_exibicao[t] = dados

        st.divider()

        # --- APORTE E AÇÃO ---
        c1, c2 = st.columns([1, 2])
        aporte = c1.number_input("💰 Aporte (R$)", value=1000.00, step=100.0)
        c2.write(""); c2.write("")
        executar = c2.button("🚀 Analisar Carteira", type="primary")

        # --- EDIÇÃO ---
        with st.expander(f"📝 Editar Ativos / Configurar Alertas ({len(carteira_exibicao)} visíveis)", expanded=True):
            add = st.text_input("Novo Ticker (ex: BBAS3)")
            if st.button("Adicionar") and add:
                t = add.upper().strip().replace(".SA","")
                if t not in carteira_completa: 
                    carteira_completa[t]={'qtde':0,'meta_pct':10,'pm':0.0,'divs':0.0, 'teto':0.0}
                    salvar_carteira(carteira_completa) 
                    st.session_state['carteira_cache'] = carteira_completa
                    st.rerun()

            st.divider()
            mudou_algo = False
            remover_lista = []
            
            # CABEÇALHO DO EDITOR
            cols_head = st.columns([1, 0.8, 0.8, 1, 1, 1, 0.5])
            cols_head[0].markdown("**Ativo**")
            cols_head[1].markdown("**Qtd**")
            cols_head[2].markdown("**Meta%**")
            cols_head[3].markdown("**PM**")
            cols_head[4].markdown("**Divs**")
            cols_head[5].markdown("**🎯 Teto (Alert)**") # NOVO
            
            for t in list(carteira_exibicao.keys()):
                cols = st.columns([1, 0.8, 0.8, 1, 1, 1, 0.5])
                cols[0].write(f"**{t}**")
                
                nq = cols[1].number_input(f"Q_{t}", value=int(carteira_completa[t]['qtde']), min_value=0, step=1, key=f"q_{t}", label_visibility="collapsed")
                nm = cols[2].number_input(f"M_{t}", value=int(carteira_completa[t]['meta_pct']), min_value=0, step=1, key=f"m_{t}", label_visibility="collapsed")
                np = cols[3].number_input(f"P_{t}", value=float(carteira_completa[t].get('pm',0)), min_value=0.0, step=0.01, format="%.2f", key=f"p_{t}", label_visibility="collapsed")
                nd = cols[4].number_input(f"D_{t}", value=float(carteira_completa[t].get('divs',0)), min_value=0.0, step=0.01, format="%.2f", key=f"d_{t}", label_visibility="collapsed")
                # NOVO CAMPO TETO
                nt = cols[5].number_input(f"T_{t}", value=float(carteira_completa[t].get('teto',0)), min_value=0.0, step=0.01, format="%.2f", key=f"t_{t}", label_visibility="collapsed", help="Se o preço cair abaixo disso, o robô avisa!")

                if cols[6].button("🗑️", key=f"del_{t}"): remover_lista.append(t); mudou_algo=True
                
                # Atualização
                if (nq!=carteira_completa[t]['qtde'] or nm!=carteira_completa[t]['meta_pct'] or 
                    np!=carteira_completa[t].get('pm',0) or nd!=carteira_completa[t].get('divs',0) or 
                    nt!=carteira_completa[t].get('teto',0)):
                    
                    carteira_completa[t].update({'qtde':nq, 'meta_pct':nm, 'pm':np, 'divs':nd, 'teto':nt})
                    mudou_algo=True
            
            if remover_lista:
                for t in remover_lista: del carteira_completa[t]
                salvar_carteira(carteira_completa); st.session_state['carteira_cache'] = carteira_completa; st.rerun()
            if mudou_algo: 
                salvar_carteira(carteira_completa); st.session_state['carteira_cache'] = carteira_completa

        # --- DASHBOARD ---
        if executar or modo_live:
            if carteira_exibicao:
                with st.spinner("Varrendo o Mercado..."):
                    df = pd.DataFrame.from_dict(carteira_exibicao, orient='index')
                    precos = {}
                    for t in df.index: precos[t] = obter_preco_atual(t)
                    df['preco_atual'] = df.index.map(precos)
                    df = df[df['preco_atual'] > 0]

                    if not df.empty:
                        # --- LÓGICA DO SNIPER (ALERTA DE PREÇO) ---
                        oportunidades = []
                        for t in df.index:
                            preco = df.loc[t, 'preco_atual']
                            teto = df.loc[t, 'teto']
                            if teto > 0 and preco <= teto:
                                oportunidades.append(f"{t}: R$ {preco:.2f} (Abaixo de R$ {teto:.2f})")
                        
                        if oportunidades:
                            st.toast(f"🚨 {len(oportunidades)} Oportunidades Detectadas!", icon="🔥")
                            st.warning(f"### 🔥 ALERTA DE COMPRA (PREÇO TETO ATINGIDO):\n" + "\n".join([f"- {op}" for op in oportunidades]))

                        # --- CÁLCULOS NORMAIS ---
                        df['total_atual'] = df['qtde'] * df['preco_atual']
                        df['total_inv'] = df['qtde'] * df['pm']
                        df['lucro_cota'] = df['total_atual'] - df['total_inv']
                        df['lucro_real'] = df['lucro_cota'] + df['divs']
                        df['rentab_pct'] = df.apply(lambda x: (x['lucro_real']/x['total_inv'])*100 if x['total_inv']>0 else 0, axis=1)
                        df['yoc_pct'] = df.apply(lambda x: (x['divs']/x['total_inv'])*100 if x['total_inv']>0 else 0, axis=1)
                        df['setor'] = df.index.map(obter_setor)
                        df['link_analise'] = df.index.map(obter_link_investidor10)

                        df_fim, sobra = calcular_compras(df, aporte)
                        
                        patr = df_fim['total_atual'].sum()
                        lucro = df_fim['lucro_real'].sum()
                        projecao_anual = patr * 0.08 
                        
                        k1, k2, k3, k4 = st.columns(4)
                        k1.metric("Patrimônio", f"R$ {patr:,.2f}")
                        k2.metric("Lucro Real", f"R$ {lucro:,.2f}", delta=f"{(lucro/patr*100) if patr>0 else 0:.1f}%")
                        k3.metric("🔮 Projeção Anual (8%)", f"R$ {projecao_anual:,.2f}", delta="Estimado")
                        k4.metric("Caixa", f"R$ {sobra:,.2f}")

                        st.divider()

                        # --- LISTA DE COMPRAS ---
                        st.subheader("🛒 Ordem de Compra")
                        compra = df_fim[df_fim['comprar_qtd']>0].sort_values('custo_total', ascending=False)
                        if not compra.empty:
                            st.dataframe(compra[['preco_atual','meta_pct','comprar_qtd','custo_total']].style.format({'preco_atual':'R$ {:.2f}','custo_total':'R$ {:.2f}','meta_pct':'{:.0f}%'}), use_container_width=True)
                        else: st.success("Aguarde! Nenhuma compra necessária.")

                        st.divider()
                        st.subheader("🔎 Detalhes Interativos (com Alertas)")
                        
                        cols = ['link_analise', 'qtde','pm','preco_atual', 'teto', 'divs','lucro_real','rentab_pct', 'yoc_pct']
                        df_show = df_fim[cols].sort_values('rentab_pct', ascending=False)
                        
                        st.dataframe(
                            df_show,
                            column_config={
                                "link_analise": st.column_config.LinkColumn("Analisar", display_text="Ver no Inv10"),
                                "pm": st.column_config.NumberColumn("PM", format="R$ %.2f"),
                                "preco_atual": st.column_config.NumberColumn("Preço", format="R$ %.2f"),
                                "teto": st.column_config.NumberColumn("🎯 Teto", format="R$ %.2f"), # NOVO
                                "divs": st.column_config.NumberColumn("Divs", format="R$ %.2f"),
                                "lucro_real": st.column_config.NumberColumn("Lucro", format="R$ %.2f"),
                                "rentab_pct": st.column_config.NumberColumn("% Ret", format="%.1f%%"),
                                "yoc_pct": st.column_config.NumberColumn("% YoC", format="%.1f%%"),
                            },
                            use_container_width=True,
                            hide_index=False
                        )

                        # --- SIMULADOR BOLA DE NEVE (BLINDADO) ---
                        st.divider()
                        with st.expander("🔮 Simulador Bola de Neve (O Futuro)", expanded=False):
                            st.caption("Veja o poder dos juros compostos com seu aporte mensal atual.")
                            col_sim1, col_sim2, col_sim3 = st.columns(3)
                            anos = col_sim1.slider("Anos investindo", 1, 30, 10)
                            taxa_anual = col_sim2.number_input("Taxa Anual Média (%)", value=10.0, step=0.5)
                            aporte_sim = col_sim3.number_input("Aporte Mensal (R$)", value=float(aporte), step=100.0)
                            
                            taxa_mensal = (1 + taxa_anual/100)**(1/12) - 1
                            meses = anos * 12
                            
                            evolucao = []
                            total = patr
                            total_investido = patr
                            
                            for m in range(meses):
                                total = total * (1 + taxa_mensal) + aporte_sim
                                total_investido += aporte_sim
                                if m % 12 == 0:
                                    evolucao.append({"Ano": (m//12)+1, "Total Acumulado": total, "Total Investido": total_investido})
                            
                            evolucao.append({"Ano": anos, "Total Acumulado": total, "Total Investido": total_investido})
                            df_ev = pd.DataFrame(evolucao)
                            
                            st.metric(f"Patrimônio em {anos} anos", f"R$ {total:,.2f}", delta=f"Lucro de R$ {total - total_investido:,.2f}")
                            
                            try:
                                fig_ev = go.Figure()
                                fig_ev.add_trace(go.Scatter(x=df_ev['Ano'], y=df_ev['Total Investido'], fill='tozeroy', mode='lines', name='Saiu do Bolso', line=dict(color='#808080')))
                                fig_ev.add_trace(go.Scatter(x=df_ev['Ano'], y=df_ev['Total Acumulado'], fill='tonexty', mode='lines', name='Com Juros', line=dict(color='#00cc96')))
                                fig_ev.update_layout(title="Curva Exponencial de Riqueza", xaxis_title="Anos", yaxis_title="Patrimônio (R$)")
                                st.plotly_chart(fig_ev, use_container_width=True)
                            except Exception as e:
                                st.warning("Erro visual no gráfico (não afeta os cálculos).")
                                st.dataframe(df_ev)

            else: st.info("Filtro vazio.")

    elif menu == "⚙️ Configurações":
        st.title("Configurações (Nuvem ☁️)")
        
        st.subheader("🎯 Meta Mensal")
        nm = st.number_input("Renda Passiva Desejada (R$)", value=float(conf['meta_mensal']))
        if nm != conf['meta_mensal']:
            conf['meta_mensal'] = nm
            salvar_config(conf) 
            st.session_state['config_cache'] = conf
            st.success("Meta Salva na Nuvem!")

        st.divider()

        st.subheader("🔑 Alterar Senha")
        s1 = st.text_input("Nova Senha", type="password")
        s2 = st.text_input("Confirmar", type="password")
        if st.button("Salvar Senha"):
            if s1 == s2 and len(s1)>3:
                conf['senha'] = s1
                salvar_config(conf)
                st.session_state['config_cache'] = conf
                st.success("Senha atualizada! Faça login novamente.")
                time.sleep(2); st.session_state['logado']=False; st.rerun()
            else: st.error("Senhas diferentes ou muito curta.")
        
        st.divider()
        st.subheader("Importar Modelo")
        mod = st.selectbox("Escolha:", ["..."] + list(CARTEIRAS_PRONTAS.keys()))
        if st.button("Aplicar Modelo"):
            if mod != "...":
                novos = CARTEIRAS_PRONTAS[mod]
                for t, m in novos.items():
                    if t not in carteira_completa: carteira_completa[t] = {'qtde':0, 'meta_pct':m, 'pm':0.0, 'divs':0.0, 'teto':0.0}
                salvar_carteira(carteira_completa)
                st.session_state['carteira_cache'] = carteira_completa
                st.toast("Modelo aplicado!")
                time.sleep(1); st.rerun()

    if modo_live: time.sleep(60); st.rerun()