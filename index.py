import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support.ui import Select
import undetected_chromedriver as uc
import keyboard  # Biblioteca para capturar teclas
import autoit
from selenium.common.exceptions import NoSuchElementException
import os
import glob

from login import SENHA, USUARIO

# Solicita ao usuário para digitar uma string
entradaEscricao = input("Digite as inscrições estaduais separados por vírgula e espaço: ")
entradaDataInicio = input("Com o modelo DD/MM/YYYY digite o periodo de inicio: ")
entradaDataFinal = input("Agora digite o periodo final:")
escricoes =  entradaEscricao.split(", ")
print(escricoes)



def entrarDTE(driver, numeroIncricao):
    # Verifica se o número tem 9 dígitos
    if len(numeroIncricao) != 9:
        raise ValueError("O número deve ter exatamente 9 dígitos")
    
    # Formata o número no padrão desejado
    numero_formatado = f"{numeroIncricao[:2]}.{numeroIncricao[2:5]}.{numeroIncricao[5:8]}-{numeroIncricao[8]}"
    
    driver.get("https://portal-dte.sefaz.ce.gov.br/#/index")

    time.sleep(5)
    inicio = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/my-app/div/div/div/app-index/div/div/div[2]/div[1]/div/div/a[1]/div'))
    )
                
    inicio.click()
    
    time.sleep(3)
     
    selecaoC = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/my-app/div/div/div/app-certificado/div/ul/li/button'))
    )
                
    selecaoC.click()

    time.sleep(3)
    ativo = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/my-app/div/div/div/app-perfil/div/div[1]/table/tbody/tr/td[1]'))
    )        
    ativo.click()

    time.sleep(1)

    botaoEntrar = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/my-app/div/div/div/app-perfil/div/div[2]/button[2]'))
    )          
    botaoEntrar.click()
    time.sleep(1)
    LinhasTabela = driver.find_elements(By.XPATH, '/html/body/my-app/div/div/div/app-procuracao/div/div[2]/table/tbody/tr')
    
    for linha in LinhasTabela:
        cells = linha.find_elements(By.TAG_NAME, 'td')
        inscricao = cells[1].text
        if(numero_formatado == inscricao):
            linha.click()
            
            time.sleep(5)
            
            confirma = driver.find_elements(By.XPATH, '/html/body/my-app/div/div/div/app-procuracao/div/div[3]/button')
            botaoEntrar2 = confirma[1]
            botaoEntrar2.click()
            # for botoes in confirma:
            #     print("oi",botoes.text)
            
            siget = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/my-app/div/div/div/app-home/section/div/div[2]/div/ul/li[1]'))
            )
            time.sleep(12)        
            siget.click()
            
            
            time.sleep(12)
            autoit.send('{ENTER}')
            time.sleep(5)
            autoit.send('{ENTER}')
            time.sleep(2)
            autoit.send('{ENTER}')
            time.sleep(1)
            
            
def EntrarSiget(driver):
    
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
    
    pesquisar = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos"]/div[1]/div[1]/button'))
    )          
    pesquisar.click()
    
    time.sleep(20)
    
    valor = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos"]/table/tbody[1]/tr/td[3]/div/a'))
    )          
    valor.click()
    
    time.sleep(5)
    
    download = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="ModalDet"]/div/div/div[2]/div[1]/div/div/button'))
    )          
    download.click()
    
    csvDownload = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="ModalDet"]/div/div/div[2]/div[1]/div/div/ul/li[2]/a'))
    ) 
    csvDownload.click()
                        
    time.sleep(5000)
                
            
            
            
        
        

    
    print("Pressione F2 para continuar...")
    keyboard.wait('f2')  # Aguarda o pressionamento da tecla F2
    print("Continuando...")
    
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
                download = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div[2]/div/div/div[3]/button[3]'))
                )
                
                download.click()
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
    
    #Preenchendo dados usuario
    usuario = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="txtUsuario"]'))
    )
    usuario.send_keys(USUARIO)
    time.sleep(0.5)
    
    #Preenchendo dados senha
    senha = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="txtSenha"]'))
    )
    senha.send_keys(SENHA)
    time.sleep(0.5)
    
    #Select BOX / seleciona CONTADOR
    select_element = driver.find_element(By.XPATH, '//*[@id="cboTipoUsuario"]')
    select = Select(select_element)
    select.select_by_value("3")

    #Clica no botão entrar
    buttonLogin = driver.find_element(By.XPATH, '//*[@id="btEntrar"]')
    buttonLogin.click()
    
def iniciar_processo(driver, inscricaoEstadual):
        selectMFE = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="listaservico"]/li[11]/a'))
        )
        selectMFE.click()
        
        acessarMFE = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="menu_dir"]/ul/li/a'))
        )
        acessarMFE.click()
        
        # Aguarde a tabela carregar
        WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="form1"]/table'))
        )

        #Lista das linhas
        linhas = driver.find_elements(By.XPATH, '//*[@id="form1"]/table/tbody/tr')
        # print(linhas)
        
        for linha in linhas:
            celula = linha.find_element(By.XPATH, './td[1]')
            celulaNome = linha.find_element(By.XPATH, './td[2]') # Ajustar índice conforme necessário
            textoNome = celulaNome.text
            texto_da_celula = celula.text
            
            if(texto_da_celula == inscricaoEstadual):
                celula.click()
                print(texto_da_celula,'', textoNome)
                

def comeca_consulta(driver):
    # Encontre a quarta <li> dentro da ul com o id 'menulist_root'
    fourth_li = driver.find_element(By.XPATH, '//*[@id="menulist_root"]/li[4]')

    # Agora encontre o link <a> dentro desse quarto <li>
    link = fourth_li.find_element(By.TAG_NAME, 'a')
    
    
    link.click()
    time.sleep(2)
    
    #Preencher periodo inicial.
    dataincio = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="form-start-date-search-coupons"]'))
        )
    dataincio.send_keys(entradaDataInicio)
    
    time.sleep(1)
    
    #Preencher periodo final.
    datafinal = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="form-end-date-search-coupons"]'))
    )
    datafinal.send_keys(entradaDataFinal)
    
    time.sleep(1)
    
    #Clica no botão consultar.
    botaoConsulta = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[3]/form/div[6]/div/div/button[1]'))
    )
    botaoConsulta.click()
    
    time.sleep(2)
    

                    
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
        comeca_consulta(driver)
        
        element = WebDriverWait(driver, 200).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="table-search-coupons"]'))
        )
        
        botao100 = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[3]/div/div[2]/div/div/div/div/button[4]'))
        )
        botao100.click()
        
        print('Clickou no botão 100')
        time.sleep(2)
    except Exception as e:
        print(f"Erro ao encontrar o elemento: {e}")
        return None
        
           
    quadrados = driver.find_elements(By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[3]/div/div[2]/div/div/div/ul/li')
    if(quadrados):
        print(len(quadrados))
    
        cont = len(quadrados) - 2
        
        for i in range(cont):
            quadradoAtual = driver.find_element(By.XPATH, f'//*[@id="conteudo_central"]/div/div/div/div[3]/div/div[2]/div/div/div/ul/li[{i+2}]/a')
            quadradoAtual.click()
            
            time.sleep(2)
            # Chama a função para clicar nos links da tabela
            clicar_links_tabela(driver)
            time.sleep(2)
    
    else:
        clicar_links_tabela(driver)
            
   
        
    return element
        

# Configuração do Chrome
service = Service(ChromeDriverManager().install())
options = uc.ChromeOptions()

# Adiciona o caminho para o perfil do Chrome que contém as extensões instaladas
options.add_argument("--user-data-dir=C:/Users/pedro/AppData/Local/Google/Chrome/User Data")
options.add_argument("--profile-directory=Default")  # Modifique se necessário

# Configurações para evitar bloqueios
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-gpu")
options.add_argument("--allow-running-insecure-content")
options.add_argument("--ignore-certificate-errors")

# Inicia o navegador
driver = uc.Chrome(service=service, options=options)
driver.implicitly_wait(10)


    
try:
    for item in escricoes:
        
        entrarDTE(driver,item)
        EntrarSiget(driver)
        # Abra a página desejada
        driver.get("https://servicos.sefaz.ce.gov.br/internet/AcessoSeguro/ServicoSenha/logarusuario/login.asp")
        time.sleep(2)
        realizar_login(driver)
        # teste = "69077339"
        iniciar_processo(driver, item)
        # Aguarda o F2 e tenta localizar o elemento na nova aba
        elemento = aguardar_enter(driver)
        
        #Saindo do login para depois logar de novo
    time.sleep(5)
    sair = WebDriverWait(driver, 200).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[1]/div[1]/img[1]'))
        )
   
    sair.click()
    time.sleep(30)
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
