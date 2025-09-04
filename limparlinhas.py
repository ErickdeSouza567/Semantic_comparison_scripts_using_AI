import pandas as pd

# ===============================
# 1. Carregar o arquivo Excel
# ===============================
df = pd.read_excel(r"C:\Users\erick216008\OneDrive - Sistema Fiep\Área de Trabalho\Cods\frasonsql\cbo2002_lista_separada.xlsx")

# ===============================
# 2. Remover linhas vazias
# ===============================
# Considera vazia se a coluna 'CBO 2002' estiver vazia ou nula
df_limpo = df[df['CBO 2002'].notna() & (df['CBO 2002'] != "")]

# ===============================
# 3. Salvar em um novo Excel
# ===============================
df_limpo.to_excel("cbo2002_lista_limpa.xlsx", index=False)

print("Arquivo processado e limpo com sucesso!")
