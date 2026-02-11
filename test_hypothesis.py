import pandas as pd
import sys

# Redirect output
with open('hypothesis_results.txt', 'w', encoding='utf-8') as f:
    # Load Excel
    try:
        df_excel = pd.read_excel('Planilha controle UFV.xlsx', sheet_name='Madeira Tratada')
        df_excel.columns = df_excel.columns.str.strip()
        df_excel = df_excel.loc[:, ~df_excel.columns.str.contains('^Unnamed')]
        df_valid = df_excel[df_excel['Código UFV'].notna()].copy()
        
        f.write(f"Analyzed {len(df_valid)} rows.\n")
        
        exact_matches = 0
        total_checked = 0

        for i, row in df_valid.iterrows():
            try:
                # Try to get balance
                bal_val = row.get('Balanço Cromo %')
                if pd.isna(bal_val): continue
                ex_bal = float(str(bal_val).replace(',','.'))
                
                # Try to get retention
                ret_cr_val = row.get('Retenção Cromo (Kg/m³)')
                if pd.isna(ret_cr_val): ret_cr_val = row.get('Retenção Cromo')
                
                ret_tot_val = row.get('Retenção/concentração')
                if pd.isna(ret_tot_val): ret_tot_val = row.get('Retenção Total')
                
                if pd.isna(ret_cr_val) or pd.isna(ret_tot_val): continue
                
                ret_cr = float(str(ret_cr_val).replace(',','.'))
                ret_total = float(str(ret_tot_val).replace(',','.'))
                
                if ret_total > 0:
                    calc_bal = (ret_cr / ret_total) * 100
                    calc_bal_rounded = round(calc_bal, 1)
                    
                    if abs(calc_bal_rounded - ex_bal) < 0.1:
                        exact_matches += 1
                    else:
                        if total_checked < 5:
                            f.write(f"Row {i} MISMATCH: Ex={ex_bal} | Hyp({ret_cr}/{ret_total})={calc_bal_rounded}\n")
                    
                    total_checked += 1
            except Exception as e:
                continue

        if total_checked > 0:
            f.write(f"\nHypothesis Accuracy: {exact_matches}/{total_checked} ({exact_matches/total_checked*100:.1f}%)\n")
        else:
            f.write("No rows checked.\n")

        f.write("\n--- DEBUG ROW 115 ---\n")
        try:
            # Need to find row 115 by index if it exists in valid df
            # Note: valid df index might not match if rows were dropped.
            # But we kept valid rows with code.
            # Let's search by Code if we can't find index 115
            if 115 in df_valid.index:
                row115 = df_valid.loc[115]
                f.write(f"Code: {row115.get('Código UFV')}\n")
                f.write(f"App: {row115.get('Aplicação')}\n")
                f.write(f"Retenção Target (Ex): {row115.get('Retenção')}\n")
                f.write(f"Retenção Actual (Ex): {row115.get('Retenção/concentração')}\n")
                f.write(f"Status (Ex): {row115.get('Observação')}\n")
            else:
                f.write("Row 115 not found in valid dataframe.\n")
        except Exception as e:
            f.write(f"Error reading Row 115: {e}\n")

    except Exception as e:
        f.write(f"Critical Error: {e}\n")
