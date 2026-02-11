import pandas as pd
from modules.calculations import aplicar_formulas_excel
from modules.utils import to_float
import sys

# Redirect output
with open('debug_output.txt', 'w', encoding='utf-8') as f:
    # Load Excel
    excel_file = 'Planilha controle UFV.xlsx'
    df_excel = pd.read_excel(excel_file, sheet_name='Madeira Tratada')
    df_excel.columns = df_excel.columns.str.strip()
    df_excel = df_excel.loc[:, ~df_excel.columns.str.contains('^Unnamed')]
    df_valid = df_excel[df_excel['Código UFV'].notna()].copy()

    # Run Python Logic
    df_result = aplicar_formulas_excel(df_valid.copy())

    f.write("\n--- DEBUGGING BALANCE & STATUS ---\n")

    col_inputs = ['Cromo (%)', 'Cobre (%)', ' Arsênio (%)']
    col_bal_cr = 'Balanço Cromo %'
    col_ret_actual = 'Retenção/concentração'
    col_ret_target = 'Retenção'
    col_status = 'Observação'
    col_app = 'Aplicação'

    # Mapping
    excel_cols = {
        'Balanço Cromo %': 'Balanço Cromo %',
        'Retenção': 'Retenção',
        'Observação': 'Observação'
    }
    # Fallback map
    for k, v in excel_cols.items():
        if v not in df_valid.columns:
            for col in df_valid.columns:
                if k in col or k.replace('%', '(%)') in col:
                    excel_cols[k] = col
                    break
    
    for i, row in df_result.iterrows():
        # Helper to clean inputs
        cr = to_float(row.get(col_inputs[0]))
        cu = to_float(row.get(col_inputs[1]))
        as_val = to_float(row.get(col_inputs[2]))
        
        # Get Excel Values safely
        try:
            ex_bal = float(str(df_valid.loc[i, excel_cols['Balanço Cromo %']]).replace(',','.'))
        except: ex_bal = 0.0
        
        py_bal = row.get(col_bal_cr, 0)
        
        # Check Balance Discrepancy
        if abs(py_bal - ex_bal) > 0.1 and (cr+cu+as_val) > 0:
            f.write(f"\n--- ROW {i} (Code: {row.get('Código UFV')}) ---\n")
            f.write(f"BALANCE ERROR:\n")
            f.write(f"  Inputs: Cr={cr}, Cu={cu}, As={as_val}\n")
            soma_raw = cr + cu + as_val
            soma_rounded = round(soma_raw, 2)
            f.write(f"  Sum Input (Raw): {soma_raw}\n")
            f.write(f"  Sum Input (Rounded): {soma_rounded}\n")
            f.write(f"  Calc Py (using rounded sum?): {cr} / {soma_rounded} * 100 = {round(cr/soma_rounded*100, 2) if soma_rounded else 0}\n")
            f.write(f"  Calc Raw (using raw sum?): {cr} / {soma_raw} * 100 = {round(cr/soma_raw*100, 2) if soma_raw else 0}\n")
            f.write(f"  Excel: {ex_bal}\n")
            
        # Check Status Discrepancy
        py_stat = str(row.get(col_status, '')).strip()
        try: ex_stat = str(df_valid.loc[i, excel_cols['Observação']]).strip()
        except: ex_stat = ""
        
        # Clean text for comparison
        py_stat_clean = " ".join(py_stat.split())
        ex_stat_clean = " ".join(ex_stat.split())
        
        if py_stat_clean != ex_stat_clean and ex_stat != "nan" and ex_stat != "":
            f.write(f"\n--- ROW {i} STATUS ERROR ---\n")
            f.write(f"  App: {row.get(col_app)}\n")
            try: target = float(str(df_valid.loc[i, excel_cols['Retenção']]).replace(',','.'))
            except: target = 0.0
            
            # The code calculates target based on REGRAS_RETENCAO
            # Check if code target matches excel target
            f.write(f"  Target (Excel): {target}\n")
            f.write(f"  Actual (Py): {row.get(col_ret_actual)}\n")
            f.write(f"  Status Py: {py_stat[:50]}...\n")
            f.write(f"  Status Ex: {ex_stat[:50]}...\n")

        if i > 20: break
