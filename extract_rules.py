import pandas as pd
import sys

# Redirect output
with open('rules.txt', 'w', encoding='utf-8') as f:
    try:
        df = pd.read_excel('Planilha controle UFV.xlsx', sheet_name='Madeira Tratada')
        df.columns = df.columns.str.strip()
        df = df.dropna(subset=['Aplicação', 'Retenção'])
        
        rules = df.groupby('Aplicação')['Retenção'].agg(lambda x: x.mode().iloc[0] if not x.mode().empty else x.mean())
        
        f.write("REGRAS_RETENCAO = {\n")
        for app, ret in rules.items():
            f.write(f'    "{app.strip()}": {float(ret)},\n')
        f.write("}\n")
            
    except Exception as e:
        f.write(f"Error: {str(e)}\n")
