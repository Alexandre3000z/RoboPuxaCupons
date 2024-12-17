import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support.ui import Select
import undetected_chromedriver as uc
import keyboard  # Biblioteca para capturar teclas
import autoit
from selenium.common.exceptions import NoSuchElementException
import os
import glob
import pandas as pd
from datetime import datetime, timedelta
import calendar
from login import SENHA, USUARIO
from organizador import analisadorXmls, organizarPastas

# Data atual
data_atual = datetime.now()

# Subtrair um mês
mes_passado = data_atual.replace(day=1) - timedelta(days=1)

# Nome do mês passado
nome_mes_passado = calendar.month_name[mes_passado.month]

# Solicita ao usuário para digitar uma string
entradaEscricao = input("Digite as inscrições estaduais separados por vírgula e espaço: ")
print(f'Serão emitidos todos os cupons de {nome_mes_passado}')
entradaDataInicio = input("Com o modelo DD/MM/YYYY digite o periodo de inicio: ")
entradaDataFinal = input("Agora digite o periodo final:")


escricoes =  entradaEscricao.split(", ")

listaCFEtotal = []

nomedaempresa = ''
# Diretório de downloads do usuário (substitua conforme necessário)
downloads_directory = os.path.join(os.path.expanduser('~'), 'Downloads')
print('Arquivos serão baixados em: ',downloads_directory)

# Opções para o usuário
print("\nEscolha uma das opções abaixo:")
print("1- Tudo")
print("2- Autorizados e Cancelados")
print("3- Cancelados e Cancelamentos")
print("4- Autorizados e Cancelamentos")
print("5- Somente Autorizados")
print("6- Somente Cancelados")
print("7- Somente Cancelamentos")

# Valida a escolha do usuário
while True:
    try:
        opcao = int(input("\nDigite o número da opção desejada (1-7): "))
        if opcao in range(1, 8):
            break
        else:
            print("Opção inválida! Digite um número entre 1 e 7.")
    except ValueError:
        print("Entrada inválida! Digite apenas números entre 1 e 7.")

# Confirmando a escolha do usuário
opcoes_dict = {
    1: "Tudo",
    2: "Autorizados e Cancelados",
    3: "Cancelados e Cancelamentos",
    4: "Autorizados e Cancelamentos",
    5: "Somente Autorizados",
    6: "Somente Cancelados",
    7: "Somente Cancelamentos"
}

print(f"\nVocê selecionou: {opcoes_dict[opcao]}")



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
    time.sleep(15)
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
    
    time.sleep(3)
    #ESPERA O LOADING PARAR
    WebDriverWait(driver, 100).until_not(
    EC.presence_of_element_located((By.CLASS_NAME, 'modal fade in'))
    )
    print('loading sumiu')
    
    valor = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos"]/table/tbody[1]/tr/td[3]/div/a'))
    )          
    valor.click()
    
    time.sleep(5)
    
    download = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="ModalDet"]/div/div/div[2]/div[1]/div/div/button'))
    )

    time.sleep(2)          
    download.click()
    
    csvDownload = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="ModalDet"]/div/div/div[2]/div[1]/div/div/ul/li[2]/a'))
    ) 
    csvDownload.click()
    
    
    #continuar o codigo aqui Alexandre
    
def BaixarOsCancelados(driver):
    x = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="ModalDet"]/div/div/div[1]/button'))
    ) 
    x.click()                                    
    time.sleep(2)
    
    cancelados = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-nfce/div/div/section[2]/div/div/div/ul/li[2]/a'))
    ) 
    cancelados.click()          
    
    pesquisar = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos_outros"]/div[1]/div[1]/button'))
    ) 
    pesquisar.click()   
           
    time.sleep(3)
    #ESPERA O LOADING PARAR
    WebDriverWait(driver, 100).until_not(
    EC.presence_of_element_located((By.CLASS_NAME, 'modal fade in'))
    )
    print('loading sumiu')
    
    mes = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos_outros"]/table/tbody/tr/td[3]/div/a'))
    )
    time.sleep(2)
    mes.click()  
    time.sleep(2)
    download = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="Modal"]/div/div/div[2]/div[1]/div/div/button'))
    ) 
    download.click()  
    time.sleep(2)
    csv = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="Modal"]/div/div/div[2]/div[1]/div/div/ul/li[2]/a'))
    ) 
    csv.click()  
    
def tratarCSV(directory, lista):
    # Obtém todos os arquivos CSV no diretório especificado
    list_of_files = glob.glob(os.path.join(directory, '*.csv'))
    if not list_of_files:
        print("Nenhum arquivo CSV encontrado.")
        
    # Encontra o arquivo com a data de modificação mais recente
    latest_file = max(list_of_files, key=os.path.getmtime)
    df = pd.read_csv(latest_file)
    coluna = 'Unnamed: 0'
    # Verifica se a coluna "Chave de Acesso" existe no DataFrame
    if coluna in df.columns:
        # Armazena os valores da coluna "Chave de Acesso" em uma lista
        chave_de_acesso_list = df[coluna].dropna().tolist()
        chave_de_acesso_list = chave_de_acesso_list[3:-1]#Remover 3 primeiros e ultimo item
        listaCFEtotal.extend(chave_de_acesso_list)
        
        print('Foram encontrados ', len(chave_de_acesso_list), ' Cupons ', lista)
    else:
        print(f"A coluna {coluna} não foi encontrada no CSV.")  
          
            

    
def clicar_links_tabela(driver):
    try:

        linhas = driver.find_elements(By.XPATH, '//*[@id="table-search-coupons"]/tbody/tr')

        # Itera sobre cada linha
        for indice, linha in enumerate(linhas, 1):
            try:
                if len(linhas) > 1:
                    print(f'Baixando {indice} de {len(linhas)}')
                    print(f'Processando: {linha.text}')
                
                # Tenta encontrar o link 'a' na coluna específica
                link = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, f'//*[@id="table-search-coupons"]/tbody/tr[{indice}]/td[4]/a'))
                )
                
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
                
                
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="table-search-coupons"]'))
                )

            except Exception as e:
                print(f"Erro ao processar linha {indice}: {e}")
                continue

    except Exception as e:
        print(f"Erro ao processar tabela: {e}")
        
   #PASSO 1 AMBIENTE SEGURO     
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
    
    
    #PASSO 2 AMBIENTE SEGURO   
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
            inscricaoEstadual = inscricaoEstadual.lstrip('0')

            if(texto_da_celula == inscricaoEstadual):
                celula.click()
                print('Inscrição estadual: ',texto_da_celula,' Empresa: ', textoNome)
                break
            
                

def comeca_consulta(driver, cfe):
    # Encontre a quarta <li> dentro da ul com o id 'menulist_root'
    fourth_li = driver.find_element(By.XPATH, '//*[@id="menulist_root"]/li[4]')

    # Agora encontre o link <a> dentro desse quarto <li>
    link = fourth_li.find_element(By.TAG_NAME, 'a')
    
    
    link.click()
    # time.sleep(0.5)
    
    cfekey = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="cfeKey"]'))
        )
    cfekey.clear()
    time.sleep(0.2)
    cfekey.send_keys(cfe)
    
    time.sleep(0.1)
    
    
    #Clica no botão consultar.
    botaoConsulta = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[3]/form/div[6]/div/div/button[1]'))
    )
    botaoConsulta.click()
    
    # time.sleep(1.5)
    

  #PASSO 3 AMBIENTE SEGURO                     
def iniciarDownloads(driver):
    time.sleep(13)    
    # Captura as abas abertas no navegador
    abas = driver.window_handles

    # Certifica-se de que há pelo menos 3 abas abertas
    if len(abas) >= 3:
        # Troca para a terceira aba (última aberta)
        driver.switch_to.window(abas[2])  # O índice 2 representa a terceira aba
    else:
        print("A terceira aba não foi encontrada.")

    actions = ActionChains(driver)
    
    try:
        # print('esses são os xmls atuais',listaCFEtotal[2:10])
        analisexml = analisadorXmls(listaCFEtotal)
        for index, cupom in enumerate (analisexml, start=1):
            # Lógica de atualização e limpeza de cache
            if index % 50 == 0:  # A cada 50 itens
                print("Limpando o cache e atualizando a página...")
                
                # Atalho para abrir DevTools (Ctrl + Shift + I)
                actions.key_down(Keys.CONTROL).key_down(Keys.SHIFT).send_keys("i").key_up(Keys.SHIFT).key_up(Keys.CONTROL).perform()
                time.sleep(1)
                
                # Atalho para limpar cache (Ctrl + Shift + R para hard refresh no Chrome)
                actions.key_down(Keys.CONTROL).key_down(Keys.SHIFT).send_keys("r").key_up(Keys.SHIFT).key_up(Keys.CONTROL).perform()
                time.sleep(3)  # Aguarda a atualização

                print("Página recarregada com cache limpo.")
            
            # Continue com as operações normais no loop
            time.sleep(8)  # Simulação de tempo entre iterações
            comeca_consulta(driver,cupom)
            
            clicar_links_tabela(driver)            
            # Mostrar o progresso
            print(f"Baixando {index} de {len(analisexml)}")
            
            
            # Aqui você coloca o processamento para cada item
            # Por exemplo:
            print(f"Processando: {cupom}")
            
    except Exception as e:
        print(f"Erro ao encontrar o elemento: {e}")
        return None
def baixarCancelamento(driver):
    print('Iniciando download dos cancelamentos')
    time.sleep(3)
    # Encontre a quarta <li> dentro da ul com o id 'menulist_root'
    fourth_li = driver.find_element(By.XPATH, '//*[@id="menulist_root"]/li[4]')

    # Agora encontre o link <a> dentro desse quarto <li>
    link = fourth_li.find_element(By.TAG_NAME, 'a')
    
    
    link.click()
    
    time.sleep(2)
    
    limpar = WebDriverWait(driver, 200).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[3]/form/div[6]/div/div/button[2]'))
        )
    limpar.click()
    
    tipo = WebDriverWait(driver, 200).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[3]/form/div[3]/div/div/div/div/div[1]/span'))
        )
    tipo.click()
    
    time.sleep(0.5)
    
    autoit.send('{DOWN}')
    
    time.sleep(0.4)
    
    autoit.send('{ENTER}')
    
    time.sleep(0.4)
    inicioperiodo = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="form-start-date-search-coupons"]'))
        )
    time.sleep(0.2)
    inicioperiodo.send_keys(entradaDataInicio)
    
    time.sleep(0.4)
    finalperiodo = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="form-end-date-search-coupons"]'))
        )
    time.sleep(0.2)
    finalperiodo.send_keys(entradaDataFinal)
    
    consulta = WebDriverWait(driver, 200).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[3]/form/div[6]/div/div/button[1]'))
        )
    consulta.click()
    
    time.sleep(5)
    
# Configuração do Chrome
service = Service(ChromeDriverManager().install())
options = uc.ChromeOptions()

# Adiciona o caminho para o perfil do Chrome que contém as extensões instaladas
options.add_argument("--user-data-dir=C:/Users/ADM/AppData/Local/Google/Chrome/User Data")
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
        time.sleep(5)
        if opcao in (1, 2, 4, 5):
            tratarCSV(downloads_directory, 'autorizados')
        BaixarOsCancelados(driver)
        time.sleep(5)
        if opcao in (1, 2, 3, 6):
            tratarCSV(downloads_directory, 'cancelados')

        # Abra a página desejada
        driver.get("https://servicos.sefaz.ce.gov.br/internet/AcessoSeguro/ServicoSenha/logarusuario/login.asp")
        time.sleep(2)
        realizar_login(driver)
        # teste = "69077339"
        iniciar_processo(driver, item)
        # Aguarda o F2 e tenta localizar o elemento na nova aba
        iniciarDownloads(driver)
        if opcao in (1, 3, 4, 7):
            baixarCancelamento(driver)
            
        blocos = driver.find_elements(By.XPATH,'//*[@id="conteudo_central"]/div/div/div/div[3]/div/div[2]/div/div/div/div/button')
        if blocos:
            for index, item in enumerate(blocos):
                if index == len(blocos) - 1:
                    item.click()
                    time.sleep(5)            
            #testarrrrrr
            paginas = driver.find_elements(By.XPATH,'//*[@id="conteudo_central"]/div/div/div/div[3]/div/div[2]/div/div/div/ul/li/a')
            if paginas:
                for num, item in enumerate(paginas):
                    
                    if num != 0 and num != len(paginas) - 1:
                        if num != 1:   
                            item.click()
                            time.sleep(1)
                        clicar_links_tabela(driver)
            else:
                clicar_links_tabela(driver)        
        
                      
            
        #Saindo do login para depois logar de novo
    time.sleep(5)
    organizarPastas()
    sair = WebDriverWait(driver, 200).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[1]/div[1]/img[1]'))
        )
   
    sair.click()
    time.sleep(30)

finally:
 
    # Feche o navegador após a execução
    driver.quit()

time.sleep(2)
print("Processo finalizado... ")
time.sleep(10)
