import pandas as pd
import numpy as np
from modules.calculations import aplicar_formulas_excel
import sys

# Redirect output
with open('verification_report.txt', 'w', encoding='utf-8') as f:
    try:
        excel_file = 'Planilha controle UFV.xlsx'
        df_excel = pd.read_excel(excel_file, sheet_name='Madeira Tratada')
        df_excel.columns = df_excel.columns.str.strip()
        df_excel = df_excel.loc[:, ~df_excel.columns.str.contains('^Unnamed')]

        df_valid = df_excel[df_excel['Código UFV'].notna()].copy()
        
        f.write(f"Analyzed {len(df_valid)} rows.\n")

        # Run Logic
        df_result = aplicar_formulas_excel(df_valid.copy())

        # Map for common variations found in analysis
        col_map = {
            'Soma Concentração (%)': ['Soma Concentração', 'Soma de Concentração'],
            'Balanço Cromo %': ['Balanço Cromo (%)', 'Balanço Cromo'],
            'Balanço Cobre %': ['Balanço Cobre (%)', 'Balanço Cobre'],
            'Balanço Arsênio %': ['Balanço Arsênio (%)', 'Balanço Arsênio'],
            'Retenção Cromo (Kg/m³)': ['Retenção Cromo', 'Retenção Cromo kg/m3'],
            'Retenção Cobre (Kg/m³)': ['Retenção Cobre', 'Retenção Cobre kg/m3'],
            'Retenção Arsênio (Kg/m³)': ['Retenção Arsênio', 'Retenção Arsênio kg/m3'],
            'Retenção/concentração': ['Retenção Total', 'Retenção Total (Kg/m³)']
        }

        compare_cols = [
            'Volume (cm³)', 
            'Densidade (Kg/m³)', 
            'Soma Concentração (%)',
            'Balanço Total (%)',
            'Balanço Cromo %',
            'Balanço Cobre %',
            'Balanço Arsênio %',
            'Retenção Cromo (Kg/m³)',
            'Retenção Cobre (Kg/m³)',
            'Retenção Arsênio (Kg/m³)',
            'Retenção/concentração',
            'Observação'
        ]

        for col in compare_cols:
            target_col_excel = col
            if col not in df_valid.columns:
                # Try alternatives
                found = False
                if col in col_map:
                    for alias in col_map[col]:
                        if alias in df_valid.columns:
                            target_col_excel = alias
                            found = True
                            break
                if not found:
                    f.write(f"SKIP {col}: Not found in Excel\n")
                    continue
            
            # Compare
            diffs = 0
            for i, row in df_result.iterrows():
                try:
                    py_val = row.get(col, 0)
                    ex_val = df_valid.loc[i, target_col_excel]
                    
                    # Convert to float
                    v_py = float(str(py_val).replace(',','.')) if pd.notna(py_val) and py_val != '' else 0.0
                    v_ex = float(str(ex_val).replace(',','.')) if pd.notna(ex_val) and ex_val != '' else 0.0
                    
                    if abs(v_py - v_ex) > 0.05:
                        diffs += 1
                        if diffs <= 3:
                            f.write(f"DIFF [{col}] Row {i} | Excel({target_col_excel}): {v_ex} | Py: {v_py}\n")
                except:
                    # Text comparison for 'Observação'
                    t_py = str(py_val).strip()
                    t_ex = str(ex_val).strip()
                    if t_py != t_ex:
                        diffs += 1
                        if diffs <= 1:
                             f.write(f"DIFF TEXT [{col}] Row {i}\nExcel: {t_ex}\nPy:    {t_py}\n")

            if diffs == 0:
                f.write(f"✅ {col}: MATCH\n")
            else:
                f.write(f"❌ {col}: {diffs} discrepancies\n")

    except Exception as e:
        f.write(f"\nCRITICAL ERROR: {str(e)}\n")
