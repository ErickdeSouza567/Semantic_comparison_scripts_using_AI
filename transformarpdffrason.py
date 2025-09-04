import pandas as pd
import re

# ===============================
# 1. Carregar o Excel
# ===============================
# header=None indica que não há cabeçalho
df = pd.read_excel(r"C:\cbo2002_lista.xlsx", header=None)

# ===============================
# 2. Função para separar os dados
# ===============================
def separar_linha(linha):
    # Inicializa os campos vazios
    tipo = ""
    cbo = ""
    titulo = ""
    
    texto = str(linha).strip()
    
    # Ignorar linhas que não têm padrão CBO (ex: cabeçalhos)
    if not re.search(r'(Sinônimo|Ocupação)?\d{4,}-\d{2}', texto):
        return pd.Series([cbo, titulo, tipo])
    
    # Extrair tipo (Sinônimo ou Ocupação)
    match_tipo = re.match(r'(Sinônimo|Ocupação)', texto)
    if match_tipo:
        tipo = match_tipo.group(1)
    
    # Extrair código CBO
    match_cbo = re.search(r'(\d{4,}-\d{2})', texto)
    if match_cbo:
        cbo = match_cbo.group(1)
    
    # Extrair título (restante da linha após tipo e código)
    match_titulo = re.search(r'(?:Sinônimo|Ocupação)?\d{4,}-\d{2}\s*(.*)', texto)
    if match_titulo:
        titulo = match_titulo.group(1).strip()
    
    return pd.Series([cbo, titulo, tipo])

# ===============================
# 3. Aplicar a função à primeira coluna
# ===============================
df[['CBO 2002', 'Títulos', 'Tipo']] = df[0].apply(separar_linha)

# ===============================
# 4. Salvar em novo Excel
# ===============================
df.to_excel("cbo2002_lista_separada.xlsx", index=False)

print("Arquivo processado com sucesso!")
