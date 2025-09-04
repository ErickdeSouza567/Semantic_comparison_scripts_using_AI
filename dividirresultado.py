import pandas as pd
import os

# Define o nome do arquivo de entrada que você enviou.
input_file = r"C:\Users\erick216008\OneDrive - Sistema Fiep\Área de Trabalho\Cods\resultado.xlsx"

# Define o nome para o arquivo de saída.
output_file = r"C:\Users\erick216008\OneDrive - Sistema Fiep\Área de Trabalho\Cods\resultado_modificado.xlsx"

try:
    # Lê o arquivo Excel para um DataFrame do pandas
    df = pd.read_excel(input_file)
    
    # Verifica se a coluna 'Profissão' existe no DataFrame
    if 'Profissão' in df.columns:
        # Divide a coluna 'Profissão' em duas novas colunas, usando '-' como separador
        split_data = df['Profissão'].str.split('-', n=1, expand=True)

        # Primeira parte antes do traço
        df['Profissão'] = split_data[0].str.strip()
        # Segunda parte depois do traço
        df['Classificação'] = split_data[1].str.strip()

        # Salva o DataFrame modificado em um novo arquivo Excel
        df.to_excel(output_file, index=False)

        print(f'Processo concluído com sucesso!')
        print(f'O arquivo modificado foi salvo como: {output_file}')

    else:
        print("Erro: A coluna 'Profissão' não foi encontrada no arquivo Excel.")

except FileNotFoundError:
    print(f"Erro: O arquivo não foi encontrado em '{input_file}'.")
except Exception as e:
    print(f"Ocorreu um erro: {e}")
