import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait 
import undetected_chromedriver as uc
import keyboard  # Biblioteca para capturar teclas
import autoit
from login import SENHA, USUARIO

def clicar_links_tabela(driver):
    try:
        # Encontra todas as linhas da tabela
        linhas = driver.find_elements(By.XPATH, '//*[@id="table-search-coupons"]/tbody/tr')
        
        print(f"Número de linhas encontradas: {len(linhas)}")

        # Itera sobre cada linha
        for indice, linha in enumerate(linhas, 1):
            try:
                # Tenta encontrar o link 'a' na coluna específica
                link = linha.find_element(By.XPATH, f'//*[@id="table-search-coupons"]/tbody/tr[{indice}]/td[4]/a')
                
                print(f"Clicando no link da linha {indice}")
                
                # Rola a página até o elemento
                driver.execute_script("arguments[0].scrollIntoView(true);", link)
                
                # Clica no link
                link.click()
                time.sleep(1.5)
                download = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div[2]/div/div/div[3]/button[3]'))
                )
                
                download.click()
                time.sleep(1.5)
                fechar = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'close'))
                )
                fechar.click()
                # autoit.send_key('{ENTER}')
                # time.sleep(2)
                # autoit.send_key('{ESC}')
                # Tempo de espera entre cliques
                time.sleep(1)
                
                # Aqui você pode adicionar lógica para:
                # - Tratar a nova janela/aba
                # - Coletar informações
                # - Voltar para a página anterior
                
                # Exemplo de voltar para a página anterior
                
                
                # Espera a tabela recarregar
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="table-search-coupons"]'))
                )

            except Exception as e:
                print(f"Erro ao processar linha {indice}: {e}")
                continue

    except Exception as e:
        print(f"Erro ao processar tabela: {e}")
def realizar_login(driver):
    print("Realizando LOGIN")
    usuario = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="txtUsuario"]'))
    )
    usuario.send_keys(USUARIO)
    time.sleep(0.5)
    usuario = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="txtSenha"]'))
    )
    usuario.send_keys(SENHA)
    
def aguardar_enter(driver):
    print("Pressione F2 para continuar...")
    keyboard.wait('f2')  # Aguarda o pressionamento da tecla F2
    print("Continuando...")

    # Trocar para a nova aba
    abas = driver.window_handles
    if len(abas) > 1:
        driver.switch_to.window(abas[-1])
        print("Mudou para a nova aba!")
    else:
        print("Nenhuma nova aba encontrada.")

    # Tentar localizar o elemento específico
    try:
        element = WebDriverWait(driver, 200).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="table-search-coupons"]'))
        )
        print('ACHOUUUU')
        
        # Chama a função para clicar nos links da tabela
        clicar_links_tabela(driver)
        
        return element
    except Exception as e:
        print(f"Erro ao encontrar o elemento: {e}")
        return None

# Configuração do Chrome
service = Service(ChromeDriverManager().install())
options = uc.ChromeOptions()
# Adicionando argumentos para permitir pop-ups e evitar bloqueios
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
options.add_argument("--allow-running-insecure-content")
options.add_argument("--ignore-certificate-errors")

# Inicia o navegador
driver = uc.Chrome(service=service, options=options)
driver.implicitly_wait(10)

try:
    # Abra a página desejada
    driver.get("https://servicos.sefaz.ce.gov.br/internet/AcessoSeguro/ServicoSenha/logarusuario/login.asp")
    time.sleep(2)
    realizar_login(driver)
    # Aguarda o F2 e tenta localizar o elemento na nova aba
    elemento = aguardar_enter(driver)

    if elemento:
        # Elemento encontrado (opcional, já que a função clicar_links_tabela já foi chamada)
        elemento_texto = elemento.text
        print(f"Texto do elemento: {elemento_texto}")

finally:
    # Feche o navegador após a execução
    driver.quit()

time.sleep(2)
print("Processo finalizado... ")
time.sleep(10)