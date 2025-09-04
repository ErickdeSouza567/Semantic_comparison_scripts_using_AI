# -*- coding: utf-8 -*-
import time
import re
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor, as_completed

MAX_WORKERS = 3          # Reduzido para não sobrecarregar
TIMEOUT = 30             # Aumenta o tempo de espera
MAX_RETRIES = 2          # Tentativas caso dê erro

# --- FUNÇÃO PARA BUSCAR DESCRIÇÃO DE UM ÚNICO CBO ---
def scrape_cbo(cbo_code):
    cbo_code = re.sub(r'\D', '', str(cbo_code))
    
    for attempt in range(1, MAX_RETRIES + 1):
        chrome_options = Options()
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
        # chrome_options.add_argument('--headless')  # opcional, descomente se quiser headless
        
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        try:
            driver.get("http://www.mtecbo.gov.br/cbosite/pages/pesquisas/BuscaPorCodigo.jsf")
            wait = WebDriverWait(driver, TIMEOUT)

            # Campo CBO
            input_field = wait.until(EC.presence_of_element_located((By.NAME, "formBuscaPorCodigo:j_idt79")))
            input_field.clear()
            for char in cbo_code:
                input_field.send_keys(char)
                time.sleep(0.1)  # slow typing

            # Botão consultar
            consult_btn = wait.until(EC.element_to_be_clickable((By.ID, "formBuscaPorCodigo:btConsultarCodigo")))
            consult_btn.click()

            # Ícone do arquivo
            file_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[img[contains(@src,'arquivo.png')]]")))
            file_icon.click()

            # Obter descrição
            titulo_div = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "titulo_familia")))
            full_text = titulo_div.text
            description = full_text.split("::")[1].strip() if "::" in full_text else full_text.strip()

            return cbo_code, description

        except Exception as e:
            if attempt == MAX_RETRIES:
                return cbo_code, f"Erro: {e.__class__.__name__}"
            time.sleep(2)  # espera antes de tentar novamente
        finally:
            driver.quit()

# --- FUNÇÃO PRINCIPAL ---
def main():
    caminho_planilha = r"C:cbo2002_lista_limpa_filtrado.xlsx"
    if not os.path.exists(caminho_planilha):
        print(f"Arquivo não encontrado: {caminho_planilha}")
        return

    df = pd.read_excel(caminho_planilha, dtype={'CBO 2002': str})
    df['prefixo4'] = df['CBO 2002'].str.replace(r'\D', '', regex=True).str[:4]
    df_vazios = df[df['Descrição'].isna()]
    cbos_para_buscar = df_vazios.drop_duplicates(subset=['prefixo4'])['CBO 2002'].tolist()

    if not cbos_para_buscar:
        print("Nenhuma descrição vazia para processar.")
        return

    print(f"Encontrados {len(df_vazios)} CBOs vazios. Realizando {len(cbos_para_buscar)} buscas únicas...")

    resultados = {}
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(scrape_cbo, cbo): cbo for cbo in cbos_para_buscar}

        for i, future in enumerate(as_completed(futures)):
            cbo_original = futures[future]
            try:
                cbo_processado, descricao = future.result()
                prefixo = re.sub(r'\D', '', cbo_processado)[:4]
                resultados[prefixo] = descricao
                print(f"Progresso: {i+1}/{len(cbos_para_buscar)} | Prefixo {prefixo} -> {descricao}")
            except Exception as e:
                print(f"Erro ao processar CBO {cbo_original}: {e}")

    print("\nAtualizando a planilha...")
    for prefixo, descricao in resultados.items():
        if "Erro" not in descricao:
            df.loc[df['prefixo4'] == prefixo, 'Descrição'] = descricao

    df = df.drop(columns=['prefixo4'])
    df.to_excel(caminho_planilha, index=False)
    print("Planilha atualizada com sucesso!")

if __name__ == "__main__":
    main()
