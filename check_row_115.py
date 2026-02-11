import pandas as pd

try:
    df = pd.read_excel('Planilha controle UFV.xlsx', sheet_name='Madeira Tratada')
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    
    # Locate Row 115 (UFV-M-116 based on previous output)
    row = df[df['Código UFV'] == 'UFV-M-116'].iloc[0]
    
    print("--- ROW 115 ANALYSIS ---")
    print(f"Code: {row['Código UFV']}")
    print(f"Balanço Cr: {row['Balanço Cromo %']} (Range: 41.8 - 53.2)")
    print(f"Balanço Cu: {row['Balanço Cobre %']} (Range: 15.2 - 22.8)")
    print(f"Balanço As: {row['Balanço Arsênio %']} (Range: 27.3 - 40.7)")
    print(f"Retenção Total: {row['Retenção/concentração']} (Target: {row['Retenção']})")
    print(f"Status: {row['Observação']}")
    
except Exception as e:
    print(f"Error: {e}")
