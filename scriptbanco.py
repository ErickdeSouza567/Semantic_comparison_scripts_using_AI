import pandas as pd
import pyodbc
from datetime import datetime

# ==========================
# CONFIGURAÇÕES
# ==========================
EXCEL_PATH = r"C:\Users\erick216008\OneDrive - Sistema Fiep\Área de Trabalho\Cods\frasonsql\resultado_com_competencias.xlsx"

# String de conexão ao SQL Server (estilo solicitado — autenticação integrada / Trusted Connection)
conn = pyodbc.connect(
    r'DRIVER={SQL Server};'
    r'SERVER=SRVDCV156\SQL2022_HOMOLOG;'   # ajuste para seu servidor\instância
    r'DATABASE=CorporeRM_Diario2;'         # ajuste para seu banco
    r'Trusted_Connection=yes;'
)
cursor = conn.cursor()

# ==========================
# LEITURA DO EXCEL
# ==========================
df = pd.read_excel(EXCEL_PATH)

# ==========================
# FUNÇÃO PARA INSERIR CARGOS
# ==========================
def inserir_cargo(nome_profissao):
    agora = datetime.now()

    cursor.execute("""
        INSERT INTO dbo.t052_cargo (a052_nome, a052_status, created_at, updated_at, a052_data_suplente)
        OUTPUT INSERTED.a052_id
        VALUES (?, ?, ?, ?, ?)
    """, (nome_profissao, 1, agora, agora, agora))

    cargo_id = cursor.fetchone()[0]  # pega o id gerado
    return cargo_id

# ==========================
# FUNÇÃO PARA INSERIR VÍNCULOS NA t053_area_cargo
# ==========================
def inserir_vinculo(area_id, cargo_id, competencia_id):
    agora = datetime.now()

    cursor.execute("""
        INSERT INTO dbo.t053_area_cargo (a050_id, a052_id, a008_id, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?)
    """, (area_id, cargo_id, competencia_id, agora, agora))

# ==========================
# PROCESSAMENTO LINHA A LINHA
# ==========================
for _, row in df.iterrows():
    profissao = row["Profissão"]
    area_id = int(row["Id_Area"])
    competencias_raw = str(row["Competencias Relevantes"])  # já vem como string "id: nome; id: nome"

    # 1. Insere o cargo
    cargo_id = inserir_cargo(profissao)

    # 2. Para cada competência, insere vínculo
    competencias = [c.strip() for c in competencias_raw.split(";") if c.strip()]
    for comp in competencias:
        try:
            comp_id = int(comp.split(":")[0])  # pega só o id antes dos dois pontos
            inserir_vinculo(area_id, cargo_id, comp_id)
        except Exception as e:
            print(f"⚠ Erro ao processar competência '{comp}' para profissão '{profissao}': {e}")

    # Confirma após cada profissão
    conn.commit()
    print(f"✅ Cargo '{profissao}' inserido com ID {cargo_id} e {len(competencias)} competências vinculadas.")

# ==========================
# FINALIZA
# ==========================
cursor.close()
conn.close()
print("🚀 Processo finalizado com sucesso!")
