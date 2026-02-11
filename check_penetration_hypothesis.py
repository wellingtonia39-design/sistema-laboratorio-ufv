import pandas as pd
from modules.utils import to_float

# Load Excel
excel_file = 'Planilha controle UFV.xlsx'
df = pd.read_excel(excel_file, sheet_name='Madeira Tratada')
df.columns = df.columns.str.strip()
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
df_valid = df[df['Código UFV'].notna()].copy()

print("\n--- PENETRATION HYPOTHESIS TEST ---")
# Hypothesis: Approval REQUIRES Grau <= 2 (Full Penetration)

fail_but_passed_retention = 0
fail_due_to_grau = 0
pass_with_bad_grau = 0

for i, row in df_valid.iterrows():
    try:
        status = str(row['Observação']).strip()
        is_approved = "de acordo" in status.lower()
        
        # Get Retention
        ret_actual = to_float(row.get('Retenção/concentração', row.get('Retenção Total')))
        ret_target = to_float(row.get('Retenção'))
        
        # Get Grau
        grau = to_float(row.get('Grau penetração'))
        
        retention_ok = ret_actual >= ret_target
        grau_ok = (grau > 0 and grau <= 2) # Assumption
        
        if is_approved:
            if not grau_ok:
                print(f"Row {i} PASSED but has Bad Grau ({grau}). Code: {row['Código UFV']}")
                pass_with_bad_grau += 1
        else:
            # Failed
            if retention_ok:
                fail_but_passed_retention += 1
                if not grau_ok:
                    fail_due_to_grau += 1
                    if fail_due_to_grau <= 5:
                        print(f"Row {i} FAILED despite Retention OK ({ret_actual}>={ret_target}). Grau={grau} (Bad). Code: {row['Código UFV']}")

    except: continue

print(f"\nStats:")
print(f"Rows Failed despite Retention OK: {fail_but_passed_retention}")
print(f"  Of those, explained by Bad Grau (>2): {fail_due_to_grau}")
print(f"Rows Passed describing Bad Grau: {pass_with_bad_grau}")

if pass_with_bad_grau == 0 and fail_due_to_grau == fail_but_passed_retention:
    print("\nCONCLUSION: Hypothesis STRONG. Approval requires Grau <= 2.")
elif pass_with_bad_grau > 0:
    print("\nCONCLUSION: Hypothesis WEAK. Some passed even with Grau > 2.")
else:
    print("\nCONCLUSION: Hypothesis INCOMPLETE. Some failed despite Good Grau.")
