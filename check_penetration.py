import pandas as pd

try:
    df = pd.read_excel('Planilha controle UFV.xlsx', sheet_name='Madeira Tratada')
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    
    # Locate Row 115 (UFV-M-116)
    row = df[df['Código UFV'] == 'UFV-M-116'].iloc[0]
    
    print("--- ROW 115 PENETRATION ---")
    print(f"Code: {row['Código UFV']}")
    print(f"Grau penetração: {row['Grau penetração']}")
    print(f"Descrição Grau: {row['Descrição Grau ']}")
    print(f"Observação: {row['Observação']}")

    # Check a row that passed (e.g., Row 7)
    row7 = df.iloc[7]
    print("\n--- ROW 7 PENETRATION ---")
    print(f"Code: {row7['Código UFV']}")
    print(f"Grau penetração: {row7['Grau penetração']}")
    print(f"Observação: {row7['Observação']}")
    
except Exception as e:
    print(f"Error: {e}")
