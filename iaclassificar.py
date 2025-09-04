import pandas as pd
import numpy as np
from sentence_transformers import SentenceTransformer, util
import torch

# --- 1. Ler arquivos ---
df_resultado = pd.read_excel(r"C:\Users\erick216008\OneDrive - Sistema Fiep\Área de Trabalho\Cods\frasonsql\resultado_modificado.xlsx")
df_competencias = pd.read_excel(r"C:\Users\erick216008\OneDrive - Sistema Fiep\Área de Trabalho\Cods\frasonsql\competencias.xlsx")

# --- 2. Criar listas de nomes e ids ---
competencias_nomes = df_competencias['a008_nome'].dropna().tolist()
competencias_ids = df_competencias.loc[df_competencias['a008_nome'].notna(), 'a008_id'].tolist()

# --- 3. Carregar modelo de embeddings offline ---
print("Carregando modelo de embeddings offline...")
model = SentenceTransformer("all-mpnet-base-v2")

# --- 4. Gerar embeddings para todas as competências ---
print("Gerando embeddings para competências...")
embeddings_comp = model.encode(
    competencias_nomes,
    convert_to_tensor=True,
    batch_size=64,
    show_progress_bar=True,
    device='cpu',
    num_workers=8
)

# --- 5. Gerar embeddings para todas as vagas ---
print("Gerando embeddings para vagas...")
textos_vagas = (df_resultado['Profissão'] + " " + df_resultado['Classificação']).tolist()
embeddings_vagas = model.encode(
    textos_vagas,
    convert_to_tensor=True,
    batch_size=64,
    show_progress_bar=True,
    device='cpu',
    num_workers=6
)

# --- 6. Calcular todas as similaridades de uma vez ---
print("Calculando similaridades...")
similaridades = util.cos_sim(embeddings_vagas, embeddings_comp)  # shape: (num_vagas, num_competencias)
similaridades_np = similaridades.cpu().numpy()

# --- 7. Selecionar top 10 competências para cada vaga ---
NUM_TOP = 15
top_indices = np.argsort(-similaridades_np, axis=1)[:, :NUM_TOP]

# Montar lista de competências com IDs
competencias_top = [
    [f"{competencias_ids[j]}: {competencias_nomes[j]}" for j in row]
    for row in top_indices
]

df_resultado['Competencias Relevantes'] = ["; ".join(row) for row in competencias_top]

# --- 8. Salvar em Excel ---
df_resultado.to_excel(r"C:\Users\erick216008\Downloads\resultado_com_competencias.xlsx", index=False)

print(f"✅ Processo concluído! As {NUM_TOP} competências mais relevantes (com ID) por vaga foram salvas em 'resultado_com_competencias.xlsx'")
