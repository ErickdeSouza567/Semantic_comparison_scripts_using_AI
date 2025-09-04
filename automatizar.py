import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import time
import os
import traceback

def wait_for_page_to_load(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
    except TimeoutException:
        print("  -> Aviso: Tempo de carregamento da página excedido.")

def preencher_descricoes_cbo():
    caminho_planilha = r'C:\cbo2002_lista_limpa_filtrado.xlsx'
    url_site = 'http://www.mtecbo.gov.br/cbosite/pages/pesquisas/BuscaPorCodigo.jsf'

    if not os.path.exists(caminho_planilha):
        print(f"Erro: arquivo não encontrado: {caminho_planilha}")
        return

    print("Iniciando navegador...")
    chrome_options = Options()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')

    servico = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    driver.implicitly_wait(10)

    df = pd.read_excel(caminho_planilha, dtype={'CBO 2002': str})
    total_linhas = len(df)

    # Dicionário para armazenar descrições já encontradas pelos 4 primeiros dígitos
    descricoes_cache = {}

    for index, row in df.iterrows():
        if pd.isna(row['Descrição']):
            cbo_code_raw = str(row['CBO 2002'])
            cbo_code = ''.join(filter(str.isdigit, cbo_code_raw))
            prefixo4 = cbo_code[:4]

            # Se já temos a descrição para esse prefixo, reaplicar direto
            if prefixo4 in descricoes_cache:
                descricao = descricoes_cache[prefixo4]
                print(f"\nAplicando descrição em todos os CBOs com prefixo {prefixo4} já encontrados: '{descricao}'")
                df.loc[df['CBO 2002'].str[:4] == prefixo4, 'Descrição'] = descricao
                df.to_excel(caminho_planilha, index=False)
                continue

            print(f"\nProcessando linha {index + 1}/{total_linhas} | CBO: {cbo_code_raw} -> {cbo_code}")

            max_retries = 3
            for attempt in range(max_retries):
                try:
                    driver.get(url_site)
                    wait_for_page_to_load(driver)

                    wait = WebDriverWait(driver, 20)
                    campo_cbo = wait.until(EC.presence_of_element_located((By.NAME, 'formBuscaPorCodigo:j_idt79')))

                    print("  -> Focando e limpando o campo de CBO...")
                    campo_cbo.click()
                    time.sleep(0.2)
                    campo_cbo.send_keys("\u0001")  # Ctrl+A
                    campo_cbo.send_keys("\b" * 20)

                    print(f"  -> Digitando CBO: {cbo_code}")
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
                    print("  -> Eventos JSF disparados no campo")

                    time.sleep(1.0)

                    botao_consultar = wait.until(EC.element_to_be_clickable((By.ID, 'formBuscaPorCodigo:btConsultarCodigo')))
                    print("  -> Clicando no botão Consultar...")
                    botao_consultar.click()
                    wait_for_page_to_load(driver)

                    icone_detalhes = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//a[img[contains(@src, '/cbosite/images/arquivo.png')]]")
                    ))
                    print("  -> Clicando no ícone de detalhes...")
                    icone_detalhes.click()
                    wait_for_page_to_load(driver)

                    div_titulo = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'titulo_familia')))
                    texto_completo = div_titulo.text

                    if '::' in texto_completo:
                        descricao = texto_completo.split('::', 1)[1].strip()
                        print(f"  -> Descrição encontrada: '{descricao}'")
                        # Atualiza todos os CBOs com mesmo prefixo
                        df.loc[df['CBO 2002'].str[:4] == prefixo4, 'Descrição'] = descricao
                        df.to_excel(caminho_planilha, index=False)
                        # Salva no cache para próximos
                        descricoes_cache[prefixo4] = descricao
                        break
                    else:
                        print(f"  -> Formato inesperado da descrição: {texto_completo}")
                        break

                except Exception as e:
                    print(f"  -> ERRO na tentativa {attempt + 1}/{max_retries} para CBO {cbo_code}: {type(e).__name__}")
                    print("  -> Mensagem detalhada do erro:")
                    print(traceback.format_exc())

                    screenshot_path = f"screenshot_cbo_{cbo_code}_{index}.png"
                    html_path = f"page_cbo_{cbo_code}_{index}.html"
                    try:
                        driver.save_screenshot(screenshot_path)
                        with open(html_path, "w", encoding="utf-8") as f:
                            f.write(driver.page_source)
                        print(f"  -> Screenshot salvo em {screenshot_path}")
                        print(f"  -> HTML salvo em {html_path}")
                    except Exception as inner_e:
                        print(f"  -> Falha ao salvar screenshot/HTML: {inner_e}")

                    if attempt + 1 < max_retries:
                        print("  -> Tentando novamente em 2 segundos...")
                        time.sleep(2)
                    else:
                        print(f"  -> Falha final ao processar CBO {cbo_code}")

            time.sleep(3)

    print("\nProcesso concluído!")
    driver.quit()

if __name__ == '__main__':
    preencher_descricoes_cbo()

# o legal