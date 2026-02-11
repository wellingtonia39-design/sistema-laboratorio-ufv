import pandas as pd
import sys

# Redirect output to file
with open('excel_analysis.txt', 'w', encoding='utf-8') as f:
    try:
        xl = pd.ExcelFile('Planilha controle UFV.xlsx')
        f.write(f"SHEETS: {xl.sheet_names}\n")
        for sheet in xl.sheet_names:
            f.write(f"\n--- SHEET: {sheet} ---\n")
            df = pd.read_excel(xl, sheet)
            f.write(f"COLUMNS: {df.columns.tolist()}\n")
    except Exception as e:
        f.write(f"ERROR: {str(e)}\n")
