import streamlit as st
import pandas as pd
from modules.config import PAGE_TITLE, PAGE_ICON, LAYOUT, carregar_config, salvar_config_local, COLS_PADRAO_MADEIRA, COLS_PADRAO_SOLUCAO
from modules.auth import check_auth
from modules.drive_manager import carregar_excel_drive, salvar_excel_drive, salvar_pdf_organizado
from modules.pdf_generator import gerar_pdf
from modules.utils import get_val

# --- CONFIGURAÇÃO INICIAL ---
st.set_page_config(page_title=PAGE_TITLE, layout=LAYOUT, page_icon=PAGE_ICON)

def main():
    # 1. Autenticação
    if not check_auth():
        return

    # 2. Interface Principal
    st.title(f"{PAGE_ICON} {PAGE_TITLE}")
    
    # 3. Menu Lateral
    menu = st.sidebar.radio("Menu", ["Madeira Tratada", "Solução"])
    config = carregar_config()

    if menu == "Madeira Tratada":
        render_madeira_tratada(config)
    elif menu == "Solução":
        render_solucao(config)

def render_madeira_tratada(config):
    df = carregar_excel_drive("Madeira Tratada")
    if df.empty:
        st.warning("Não foi possível carregar os dados de Madeira Tratada.")
        return

    if "Selecionar" not in df.columns: 
        df.insert(0, "Selecionar", False)
    
    # SELETOR DE COLUNAS
    cols_disponiveis = [c for c in df.columns if c not in ["Selecionar", "Código UFV"]]
    padrao = [c for c in COLS_PADRAO_MADEIRA if c in cols_disponiveis]
    escolha_usuario = config.get("Madeira", padrao)
    
    with st.expander("⚙️ Personalizar Colunas (Adicionar/Remover)"):
        cols_visiveis = st.multiselect("Marque as colunas que deseja ver:", cols_disponiveis, default=escolha_usuario)
        if st.button("💾 Salvar Preferência"):
            config["Madeira"] = cols_visiveis
            salvar_config_local(config)
            st.success("Preferência Salva!")
            st.rerun()

    cols_finais = ["Selecionar", "Código UFV"] + cols_visiveis
    cols_finais = [c for c in cols_finais if c in df.columns]

    st.markdown("### 🔎 Buscar/Editar Amostra")
    col_busca, col_info = st.columns([1, 3])
    with col_busca: 
        numero_busca = st.text_input("Digite o número (ex: 620)", placeholder="Busque para Editar...")
    
    column_config_dates = {
        "Data de entrada": st.column_config.DateColumn("Data de entrada", format="DD/MM/YYYY"),
        "Início da análise": st.column_config.DateColumn("Início da análise", format="DD/MM/YYYY"),
        "Fim da análise": st.column_config.DateColumn("Fim da análise", format="DD/MM/YYYY"),
        "Data de Registro": st.column_config.DateColumn("Data de Registro", format="DD/MM/YYYY"),
    }

    # Lógica de Filtro
    termo_busca = None
    if numero_busca:
        termo_busca = numero_busca
        termo_full = f"UFV-M-{numero_busca}"
        # Tenta buscar pelo termo completo ou parcial
        mask = df['Código UFV'].astype(str).str.contains(termo_full, case=False, na=False)
        if not mask.any():
            mask = df['Código UFV'].astype(str).str.contains(numero_busca, case=False, na=False)
        
        df_filtrado = df[mask]
        
        with col_info: 
            st.info(f"Encontrados: {len(df_filtrado)}. Edite e clique em CALCULAR.")
        
        df_view = st.data_editor(
            df_filtrado[cols_finais], 
            num_rows="dynamic", 
            use_container_width=True, 
            key="tabela_filtrada",
            column_config=column_config_dates
        )
        
        if st.session_state.get('user') in ["admin", "Lpm"]: # Permissão
             if st.button("🧮 CALCULAR E SALVAR (Mesclar)", type="primary"):
                df.update(df_view)
                salvar_excel_drive(df, "Madeira Tratada")
                st.success("Atualizado!")
                st.rerun()
    else:
        with col_info: st.info("Mostrando tabela completa.")
        df_view = st.data_editor(
            df[cols_finais], 
            num_rows="dynamic", 
            use_container_width=True, 
            key="tabela_completa",
            column_config=column_config_dates
        )
        
        if st.session_state.get('user') in ["admin", "Lpm"]:
            if st.button("🧮 CALCULAR E SALVAR TUDO", type="primary"): 
                df.update(df_view)
                salvar_excel_drive(df, "Madeira Tratada")
                st.rerun()

    # Seleção para PDF
    st.divider()
    
    # Identificar linhas selecionadas
    # Se houve busca, precisamos mapear a seleção da filtered view para o df original
    sel_codes = []
    
    # Streamlit data_editor keys are tricky. We trust the user edits the viewed dataframe.
    # To find selected rows, we check the 'Selecionar' column in the edited dataframe (df_view)
    if not df_view.empty and "Selecionar" in df_view.columns:
        sel_codes = df_view[df_view['Selecionar'] == True]['Código UFV'].tolist()

    if sel_codes:
        st.subheader("📄 Gerar Relatório")
        # Pegamos os dados DO DATAFRAME COMPLETO (ou atualizado) para garantir que temos todas as colunas para o PDF, 
        # não só as visíveis.
        sel_row = df[df['Código UFV'].isin(sel_codes)]
        
        if not sel_row.empty:
            l = sel_row.iloc[0].to_dict() # Pega o primeiro selecionado
            
            try:
                pdf_bytes = gerar_pdf(l)
                nome_arquivo = f"{l.get('Código UFV', 'Relatorio')}.pdf"
                
                c_down, c_cloud = st.columns(2)
                with c_down: 
                    st.download_button("⬇️ BAIXAR PDF (PC)", pdf_bytes, nome_arquivo, "application/pdf", type="primary")
                with c_cloud:
                    if st.button("☁️ SALVAR NO DRIVE COMPARTILHADO"): 
                        salvar_pdf_organizado(pdf_bytes, nome_arquivo, get_val(l, ["Data de entrada"]))
            except Exception as e:
                st.error(f"Erro na geração do PDF: {e}")
    else:
        if numero_busca and (df_filtrado.empty if 'df_filtrado' in locals() else True):
             st.warning("Nenhum resultado encontrado.")

def render_solucao(config):
    df = carregar_excel_drive("Solução Preservativa")
    if df.empty:
        st.warning("Não foi possível carregar os dados de Solução.")
        return

    cols_disponiveis = [c for c in df.columns if c not in ["Código UFV"]]
    padrao = [c for c in COLS_PADRAO_SOLUCAO if c in cols_disponiveis]
    escolha_usuario = config.get("Solução", padrao)
    
    with st.expander("⚙️ Personalizar Colunas"):
        cols_visiveis = st.multiselect("Marque as colunas que deseja ver:", cols_disponiveis, default=escolha_usuario)
        if st.button("💾 Salvar Preferência Solução"):
            config["Solução"] = cols_visiveis
            salvar_config_local(config)
            st.rerun()
    
    cols_finais = ["Código UFV"] + cols_visiveis
    cols_finais = [c for c in cols_finais if c in df.columns]
    
    column_config_dates = {
            "Data de entrada": st.column_config.DateColumn("Data de entrada", format="DD/MM/YYYY"),
            "Início da análise": st.column_config.DateColumn("Início da análise", format="DD/MM/YYYY"),
            "Fim da análise": st.column_config.DateColumn("Fim da análise", format="DD/MM/YYYY"),
            "Data de Registro": st.column_config.DateColumn("Data de Registro", format="DD/MM/YYYY"),
    }

    st.data_editor(df[cols_finais], use_container_width=True, column_config=column_config_dates)

if __name__ == "__main__":
    main()
