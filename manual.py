import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select

import autoit
import keyboard
import re

def processo1_Dte(driver, mes_desejado, ano_desejado ):
    
    time.sleep(1)
    
    driver.get("https://portal-dte.sefaz.ce.gov.br/#/index")
    
    print('Selecione o certificado da empresa desejada e quando a opção SIGET aparecer aperte F2')
    
    keyboard.wait("f2")
    
    print('F2 pressionado, iniciando o processo')
    
    inscricao = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/my-app/div/div/div/app-breadcrumb/div/div[2]/span'))
    )
    stringInscricao = inscricao.text
    
    # Expressão regular para capturar o CGF
    padrao_cgf = r"CGF:\s(\d{2}\.\d{3}\.\d{3}-\d{1})"

    # Procura o padrão na string
    resultado = re.search(padrao_cgf, stringInscricao)
        
    # Verifica se encontrou o CGF
    if resultado:
        cgf = resultado.group(1)  # Captura o valor original
        cgf_formatado = re.sub(r"[^\d]", "", cgf)  # Remove tudo que não é número
        
        
    siget = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/my-app/div/div/div/app-home/section/div/div[2]/div/ul/li[1]'))
    )
    time.sleep(12)        
    siget.click()
    
    
    time.sleep(16)
    autoit.send('{ENTER}')
    time.sleep(5)
    autoit.send('{ENTER}')
    time.sleep(2)
    autoit.send('{ENTER}')
    time.sleep(1)
    
    # Guarda o identificador da janela original
    original_window = driver.current_window_handle

    # Espera até que uma nova janela esteja aberta
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

    # Captura todos os identificadores de janelas abertas
    windows = driver.window_handles

    # Muda para a nova janela
    for window in windows:
        if window != original_window:
            driver.switch_to.window(window)
            break
    
    NfeCfe = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="menu_indicadores_nfce"]'))
    )          
    NfeCfe.click()
    
    time.sleep(3)
    
    #SELECIONAR MES E ANO
    selectMes = driver.find_element(By.XPATH, '//*[@id="mes_select"]')
    optionMes = Select(selectMes)
    optionMes.select_by_index(mes_desejado - 1)
    
    selectAno = driver.find_element(By.XPATH, '//*[@id="ano_select"]')
    optionAno = Select(selectAno)
    optionAno.select_by_value(f'{ano_desejado}')
    
    
    time.sleep(2)
    pesquisar = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos"]/div[1]/div[1]/button'))
    )          
    pesquisar.click()
    
    time.sleep(30)
    #ESPERA O LOADING PARAR
    WebDriverWait(driver, 100).until_not(
    EC.presence_of_element_located((By.CLASS_NAME, 'modal fade in'))
    )
    print('loading sumiu')
    tabelaValor = driver.find_element(By.XPATH, f'//*[@id="tab_emitidos"]/table/tbody[1]/tr/td[3]/div')
    
    if(tabelaValor.text != '0,00'):
        valor = WebDriverWait(driver, 500).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos"]/table/tbody[1]/tr/td[3]/div/a'))
        )          
        valor.click()

        time.sleep(10)
        
        download = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ModalDet"]/div/div/div[2]/div[1]/div/div/button'))
        )

        time.sleep(2)          
        download.click()
        
        time.sleep(8)
        
        csvDownload = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ModalDet"]/div/div/div[2]/div[1]/div/div/ul/li[2]/a'))
        ) 
        csvDownload.click()
        time.sleep(8)
        
        x = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ModalDet"]/div/div/div[1]/button'))
        ) 
        x.click()                                    
        time.sleep(2)
        return ['True', cgf_formatado]
    
    else:
        print('Essa empresa não tem cupons fiscais autorizados até o momento.')
        return ['False', cgf_formatado]
    
    