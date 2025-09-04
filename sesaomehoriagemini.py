# -*- coding: utf-8 -*-
import time
import re
import os
import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor, as_completed

MAX_WORKERS = 5   # aumente para paralelismo
TIMEOUT = 30
MAX_RETRIES = 3

def wait_for_page_to_load(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
    except:
        print("  -> Aviso: Tempo de carregamento excedido.")

def processar_cbo(cbo_code):
    """Processa um único CBO e retorna (prefixo4, descricao ou erro)"""
    cbo_code = ''.join(filter(str.isdigit, str(cbo_code)))
    prefixo4 = cbo_code[:4]

    for attempt in range(1, MAX_RETRIES+1):
        chrome_options = Options()
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        # chrome_options.add_argument("--headless")  # opcional

        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        try:
            url_site = 'http://www.mtecbo.gov.br/cbosite/pages/pesquisas/BuscaPorCodigo.jsf'
            driver.get(url_site)
            wait_for_page_to_load(driver)
            wait = WebDriverWait(driver, TIMEOUT)

            # Preencher o campo com slow typing
            campo_cbo = wait.until(EC.presence_of_element_located((By.NAME, 'formBuscaPorCodigo:j_idt79')))
            campo_cbo.clear()
            for char in cbo_code:
                campo_cbo.send_keys(char)
                time.sleep(0.15)

            # JS para disparar eventos JSF
            driver.execute_script("""
                var campo = arguments[0];
                var valor = arguments[1];
                campo.focus();
                campo.value = valor;
                campo.dispatchEvent(new Event('keydown', { bubbles: true }));
                campo.dispatchEvent(new Event('keyup', { bubbles: true }));
                campo.dispatchEvent(new Event('input', { bubbles: true }));
                campo.dispatchEvent(new Event('change', { bubbles: true }));
                campo.blur();
            """, campo_cbo, cbo_code)
            time.sleep(0.5)

            # Botão Consultar
            botao_consultar = wait.until(EC.element_to_be_clickable((By.ID, 'formBuscaPorCodigo:btConsultarCodigo')))
            botao_consultar.click()
            wait_for_page_to_load(driver)

            # Ícone de detalhes
            icone_detalhes = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[img[contains(@src, '/cbosite/images/arquivo.png')]]")))
            icone_detalhes.click()
            wait_for_page_to_load(driver)

            # Capturar descrição
            div_titulo = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'titulo_familia')))
            texto_completo = div_titulo.text
            descricao = texto_completo.split('::',1)[1].strip() if '::' in texto_completo else texto_completo.strip()

            driver.quit()
            return prefixo4, descricao

        except Exception as e:
            driver.quit()
            if attempt == MAX_RETRIES:
                erro_msg = f"Erro final: {type(e).__name__}"
                return prefixo4, erro_msg
            time.sleep(2)

def preencher_descricoes_cbo():
    caminho_planilha = r'C:\Users\erick216008\OneDrive - Sistema Fiep\Área de Trabalho\Cods\frasonsql\cbo2002_lista_limpa_filtrado.xlsx'
    if not os.path.exists(caminho_planilha):
        print(f"Erro: arquivo não encontrado: {caminho_planilha}")
        return

    df = pd.read_excel(caminho_planilha, dtype={'CBO 2002': str})
    df['prefixo4'] = df['CBO 2002'].str.replace(r'\D','', regex=True).str[:4]
    df_vazios = df[df['Descrição'].isna()]
    cbos_para_buscar = df_vazios.drop_duplicates(subset=['prefixo4'])['CBO 2002'].tolist()

    print(f"Total de CBOs a buscar: {len(cbos_para_buscar)}")

    descricoes_cache = {}
    resultados_logs = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(processar_cbo, cbo): cbo for cbo in cbos_para_buscar}

        for i, future in enumerate(as_completed(futures)):
            cbo_original = futures[future]
            try:
                prefixo, descricao = future.result()
                resultados_logs.append((prefixo, descricao))
                if 'Erro' not in descricao:
                    df.loc[df['CBO 2002'].str[:4] == prefixo, 'Descrição'] = descricao
                    descricoes_cache[prefixo] = descricao
                    print(f"[{i+1}/{len(cbos_para_buscar)}] Prefixo {prefixo} atualizado: {descricao}")
                else:
                    print(f"[{i+1}/{len(cbos_para_buscar)}] Prefixo {prefixo} falhou: {descricao}")
            except Exception as e:
                print(f"[{i+1}/{len(cbos_para_buscar)}] Erro ao processar CBO {cbo_original}: {e}")
                traceback.print_exc()

    # Salvar Excel apenas uma vez
    df = df.drop(columns=['prefixo4'])
    df.to_excel(caminho_planilha, index=False)
    print("\nProcesso concluído! Excel salvo com sucesso.")
    print("\nResumo dos resultados:")
    for log in resultados_logs:
        print(f"Prefixo {log[0]} -> {log[1]}")

if __name__ == '__main__':
    preencher_descricoes_cbo()
