import time
# from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait 
import undetected_chromedriver as uc
import autoit 
import pandas as pd
from lerPlanilha import puxarCnpj, nomeEmpresas
# Configura o serviço do ChromeDriver


dados = []
empresas = nomeEmpresas()
pxacnpj = puxarCnpj()
# empresas = ['teste2', 'teste3']
# pxacnpj = ["39.644.188/0001-46","39.644.188/0001-46"]
# Abre uma página da web
service = Service(ChromeDriverManager().install())
        # Configurações do Chrome com o perfil
options = uc.ChromeOptions()


# Inicia o navegador
driver = uc.Chrome(service=service, options=options)
driver.implicitly_wait(10)
autoit.send("{f11}")
for empresa, cnpj in zip(empresas, pxacnpj):
    
    driver.get("https://issadmin.sefin.fortaleza.ce.gov.br/grpfor/pagesPublic/simplesNacional/consultarSituacao.seam")
    cnpjAtual = cnpj
    time.sleep(2)
    
    areatext_cnpj = driver.find_element(By.XPATH, '//*[@id="pesquisaForm:cnpjDec:cnpj"]')
    areatext_cnpj.send_keys(cnpj)
    time.sleep(0.1)
    areacaptcha_Text = driver.find_element(By.XPATH, '//*[@id="pesquisaForm:captchaDecor:inputCaptcha"]')
    autoit.send("{TAB}")
    autoit.send("{TAB}")
    time.sleep(10)
    captcha = True
    iteration_count = 0
    while captcha:
        button_submit = driver.find_element(By.XPATH, '//*[@id="pesquisaForm:botaoPesquisar"]')     
        button_submit.click()
        time.sleep(2)
        if iteration_count != 0:
            autoit.send("{TAB}")
            autoit.send("{TAB}")   
        
        try:  
            iteration_count = iteration_count + 1  
            # Espera até que o elemento captcha seja encontrado
            captcha_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="pesquisaForm:j_id53"]/div[1]/ul'))
            )              
            

            # Se o elemento for encontrado, interrompe o loop
            if captcha_element:
                print("Pagina encontrada, validando...")
                list_items = driver.find_elements(By.TAG_NAME, 'li')
                item = list_items[4]
                icon = item.find_element(By.CSS_SELECTOR, "i")
                # Verifica se a classe do <i> é a desejada
                if "fa-check" in icon.get_attribute("class"):
                    # Toma uma ação se a condição for satisfeita
                    print(item.text, ": Verdadeiro")
                    situacao = "Tem"
                else:
                    print(item.text, ": Falso")
                    situacao = "Nao tem"    
                    
                break
        except Exception as e:
            # Se ocorrer uma exceção (por exemplo, tempo limite), você pode lidar com ela aqui
            print("Pagina não encontrada, tentando novamente em 10 segundos...")
            
    
            
    
    dados.append({"Empresa":empresa, "CNPJ":cnpj, "Contabilista":situacao})        
            
    
    
    
            
    time.sleep(3)


# Cria um DataFrame a partir dos dados acumulados
df = pd.DataFrame(dados)

# Salva o DataFrame em um arquivo Excel
df.to_excel("resultado_empresas.xlsx", index=False)

time.sleep(2)
print("Processo finalizado... Planilha criada.")
time.sleep(10)