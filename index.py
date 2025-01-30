import time
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support.ui import Select
import undetected_chromedriver as uc
import keyboard  # Biblioteca para capturar teclas
import autoit
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import os
import glob
import pandas as pd
from datetime import datetime, timedelta
import calendar
from login import SENHA, USUARIO
from organizador import analisadorXmls, organizarPastas
import requests
from T import TOKEN

def verificar_arquivo(url, texto_procurado):
    try:
        response = requests.get(url)
        response.raise_for_status()  # Gera um erro se a requisição falhar
        conteudo = response.text
        return texto_procurado in conteudo
    except Exception as e:
        print(f"Erro ao acessar o arquivo: {e}")
        return False

# Fluxo principal
def validarAcesso():
    
    url = "https://drive.google.com/uc?export=download&id=1k9Y8YgnsE4KPvQ3_62MdELLj6UhVztsH"
    texto_procurado = TOKEN
    
    if verificar_arquivo(url, texto_procurado):
        
        return True
        # Código autorizado
    else:
        # print("Licença não encontrada. Execução não autorizada.")
        return False
validacao = validarAcesso()


def get_previos_month_and_year():
    today = datetime.today()
    if today.month ==  1:
        previous_month = 12
        year = today.year - 1
    else:
        previous_month = today.month - 1
        year = today.year
    return previous_month,year

mesano = get_previos_month_and_year()
mes_passado = mesano[0]
ano = mesano[1]
nome_mes_passado = calendar.month_name[mes_passado]

def gerar_datas_mes(mes: int, ano: int):
        """
        Gera as datas de início e fim do mês fornecido.
        :param mes: Mês desejado (1-12)
        :param ano: Ano desejado (ex: 2024)
        :return: variavel_inicio, variavel_final como strings no formato DD/MM/AAAA
        """
        # Primeiro dia do mês
        primeiro_dia = f"01/{mes:02d}/{ano}"
        
        # Último dia do mês usando o módulo 'calendar'
        ultimo_dia_numero = calendar.monthrange(ano, mes)[1]  # Retorna o último dia do mês
        ultimo_dia = f"{ultimo_dia_numero:02d}/{mes:02d}/{ano}"
        
        return primeiro_dia, ultimo_dia


    
variavel_inicio, variavel_final = gerar_datas_mes(mes_passado, ano)




# Solicita ao usuário para digitar uma string
# entradaEscricao = input("Digite as inscrições estaduais separados por vírgula e espaço: ")
print(f'Serão emitidos todos os cupons de {nome_mes_passado}')

# escricoes =  entradaEscricao.split(", ")

listaCFEtotal = []
listaEscricoesAS = []
listaEscricoesDTE = []
listaTotal = []
#//*[@id="modalMensagem"]/div/div/div[1]/button/span
nomedaempresa = ''
# Diretório de downloads do usuário (substitua conforme necessário)
downloads_directory = os.path.join(os.path.expanduser('~'), 'Downloads')
print('Arquivos serão baixados em: ',downloads_directory)

def pegarForAmbiente(driver):
    driver.get("https://servicos.sefaz.ce.gov.br/internet/AcessoSeguro/ServicoSenha/logarusuario/login.asp")
    time.sleep(2)
    autoit.send('{F11}')
    time.sleep(1)
    realizar_login(driver)
    
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
    print('2 segundos')
    time.sleep(2)
    for linha in linhas:
        inscricao = linha.find_element(By.XPATH, './td[1]')
        celulaNome = linha.find_element(By.XPATH, './td[2]') # Ajustar índice conforme necessário
        textoNome = celulaNome.text
        texto_da_celula = f'0{inscricao.text}'
        listaEscricoesAS.append(texto_da_celula)
    
    del listaEscricoesAS[0]
    print('Foram registradas ',len(listaEscricoesAS), 'empresas no ambiente seguro')    
    
    saida = WebDriverWait(driver,20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="usuarioLog"]/a[2]'))
    )
    saida.click()
    
    time.sleep(5)
    
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
    ativo = WebDriverWait(driver, 300).until(
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
        situacao = cells[5].text
        numero_formatado = inscricao.replace('.', '').replace('-', '')
        if situacao == 'Válida':
            listaEscricoesDTE.append(numero_formatado)
    print('Foram registradas ',len(listaEscricoesDTE), ' empresas na SEFAZ DTE')    
    listafiltrada  = [numero for numero in listaEscricoesAS if numero in listaEscricoesDTE]
    
    
    
    listaTotal.extend(listafiltrada)
    print('Feito a filtragem DTE > AMBIENTE SEGURO: ', len(listafiltrada), ' empresas tem procuração nas duas aplicações.')


def entrarDTE(driver, numeroIncricao):
    # Verifica se o número tem 9 dígitos
    if len(numeroIncricao) != 9:
        raise ValueError("O número deve ter exatamente 9 dígitos")
    
    
   
    # Formata o número no padrão desejado
    numero_formatado = f"{numeroIncricao[:2]}.{numeroIncricao[2:5]}.{numeroIncricao[5:8]}-{numeroIncricao[8]}"
    
    # Obtém todas as guias abertas
    abas = driver.window_handles

    # Fecha todas as guias menos a última
    for aba in abas[:-1]:  # Loop através de todas menos a última
        driver.switch_to.window(aba)
        driver.close()

    # Alterna para a última guia
    driver.switch_to.window(abas[-1])
    
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
    
    try:
        ativo = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/my-app/div/div/div/app-perfil/div/div[1]/table/tbody/tr/td[1]'))
        )        
        ativo.click()
    except:
        print('Carregamento infinito detectado, corrigindo...')
        driver.refresh()
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
    
    try:
        siget = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/my-app/div/div/div/app-home/section/div/div[2]/div/ul/li[1]'))
        )
        time.sleep(12)        
        siget.click()
    except:
        print('A empresa não possui procuração no DTE, fazer procuração e tentar novamente...')
        
        driver.get('https://portal-dte.sefaz.ce.gov.br/#/home')
        
        time.sleep(2)
        
        perfildte = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/my-app/header/div/div/nav/ul/li[3]/a'))
        )       
        perfildte.click()
        time.sleep(3)
        sairPefildte = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/my-app/header/div[2]/div/nav/ul/li[2]/div/button'))
        )       
        sairPefildte.click()
        
        time.sleep(2)
        return None
    
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
    
    time.sleep(8)
    
    #SELECIONAR MES E ANO
    selectMes = driver.find_element(By.XPATH, '//*[@id="mes_select"]')
    optionMes = Select(selectMes)
    optionMes.select_by_index(mes_passado - 1)
    
    time.sleep(1)
    
    selectAno = driver.find_element(By.XPATH, '//*[@id="ano_select"]')
    optionAno = Select(selectAno)
    optionAno.select_by_value(f'{ano}')
    
    
    time.sleep(2)
    
    
    pesquisar = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos"]/div[1]/div[1]/button'))
    )          
    pesquisar.click()
    
    time.sleep(15)
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
        return True
    
    else:
        print('Essa empresa não tem cupons fiscais autorizados até o momento.')
        return False
    
    #continuar o codigo aqui Alexandre

def sairDte(driver):
    perfilA = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="perfil-empresa"]/li[2]/a'))
    ) 
    perfilA.click()
    time.sleep(2)
    
    sairA = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="perfil-empresa"]/li[2]/ul/li[2]/div/a'))
    ) 
    sairA.click()
    time.sleep(2)
    
    driver.get('https://portal-dte.sefaz.ce.gov.br/#/home')
    
    time.sleep(2)
    
    perfildte = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/my-app/header/div/div/nav/ul/li[3]/a'))
    )       
    perfildte.click()
    time.sleep(3)
    sairPefildte = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/my-app/header/div[2]/div/nav/ul/li[2]/div/button'))
    )       
    sairPefildte.click()
    
    time.sleep(2)
    
def BaixarOsCancelados(driver):
    cancelados = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-nfce/div/div/section[2]/div/div/div/ul/li[2]/a'))
    ) 
    cancelados.click()
    
    time.sleep(7)
             
    #SELECIONAR MES E ANO
    selectMes = driver.find_element(By.XPATH, '//*[@id="mes_select"]')
    optionMes = Select(selectMes)
    optionMes.select_by_index(mes_passado - 1)
    
    selectAno = driver.find_element(By.XPATH, '//*[@id="ano_select"]')
    optionAno = Select(selectAno)
    optionAno.select_by_value(f'{ano}')
    
    
    time.sleep(2)
    pesquisar = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos_outros"]/div[1]/div[1]/button'))
    ) 
    pesquisar.click()  
            
    time.sleep(15)
    #ESPERA O LOADING PARAR
    WebDriverWait(driver, 50).until_not(
    EC.presence_of_element_located((By.CLASS_NAME, 'modal fade in'))
    )
    print('loading sumiu')
           
    time.sleep(10)
    tabelaValor = driver.find_element(By.XPATH, f'//*[@id="tab_emitidos_outros"]/table/tfoot/tr/td[3]')
    if(tabelaValor.text != '0,00'):

        mes = WebDriverWait(driver, 5000).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="tab_emitidos_outros"]/table/tbody/tr/td[3]/div/a'))
        ) 
        mes.click()  
        time.sleep(6)
        download = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="Modal"]/div/div/div[2]/div[1]/div/div/button'))
        ) 
        download.click()  
        time.sleep(8)
        csv = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="Modal"]/div/div/div[2]/div[1]/div/div/ul/li[2]/a'))
        ) 
        csv.click()
        time.sleep(8)
        
        x = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="Modal"]/div/div/div[1]/button'))
        ) 
        x.click()
        time.sleep(2)
        
        perfilA = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="perfil-empresa"]/li[2]/a'))
        ) 
        perfilA.click()
        time.sleep(2)
        
        sairA = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="perfil-empresa"]/li[2]/ul/li[2]/div/a'))
        ) 
        sairA.click()
        time.sleep(2)
        
        driver.get('https://portal-dte.sefaz.ce.gov.br/#/home')
        
        time.sleep(2)
        
        perfildte = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/my-app/header/div/div/nav/ul/li[3]/a'))
        )       
        perfildte.click()
        time.sleep(3)
        sairPefildte = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/my-app/header/div[2]/div/nav/ul/li[2]/div/button'))
        )       
        sairPefildte.click()
        
        time.sleep(2)
        
        return True
        
    else:
        
        print('Essa empresa não tem cupons fiscais de cancelamento até o momento')
        perfil = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="perfil-empresa"]/li[2]/a'))
        ) 
        perfil.click()
        
        time.sleep(1)
        
        sair = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="perfil-empresa"]/li[2]/ul/li[2]/div/a'))
        ) 
        sair.click()
        
        time.sleep(3)
        
        driver.get('https://portal-dte.sefaz.ce.gov.br/#/home')
        
        time.sleep(2)
        
        perfil = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/my-app/header/div/div/nav/ul/li[3]/a'))
        )       
        perfil.click()
        
        sairPefil = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/my-app/header/div[2]/div/nav/ul/li[2]/div/button'))
        )       
        sairPefil.click()
        
        time.sleep(2)
        return False
    
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
          
            

    
def clicar_links_tabela(driver, index):
    try:
        if(index % 50 == 0 and index != 0):
            # Limpar o cache usando DevTools Protocol
            driver.execute_cdp_cmd('Network.clearBrowserCache', {})
            driver.execute_cdp_cmd('Network.clearBrowserCookies', {})

            # Atualizar a página
            driver.refresh()
            time.sleep(8)
            
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
        
    # Localização do elemento
    element_locator = (By.CSS_SELECTOR, "div.modal-backdrop.am-fade")

    # Aguarde até que a classe 'ng-hide' seja adicionada ao elemento
    try:
        WebDriverWait(driver, 150).until(
            lambda driver: "ng-hide" in driver.find_element(*element_locator).get_attribute("class")
        )
        print("O elemento adquiriu a classe 'ng-hide'.")
    except Exception as e:
        driver.refresh()
        print("Timeout: o elemento não adquiriu a classe 'ng-hide' dentro do tempo esperado.")
        time.sleep(5)
        
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
def iniciarDownloads(driver, inscricao):
    time.sleep(13)
    
    # Captura as abas abertas no navegador
    abas = driver.window_handles

    # Certifica-se de que há pelo menos 3 abas abertas
    if len(abas) >= 3:
        driver.close()

        driver.switch_to.window(abas[2])  
        
        
    else:
        print("A segunda aba não foi encontrada.")

    try:
        
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH , '//*[@id="conteudo_central"]/div/div[2]/div/div/div[1]/h4')))
            print('ENCONTROU O AVISO')
            
            button100 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH , '//*[@id="conteudo_central"]/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/button[4]')))
            button100.click()
            
            tabelaAvisos = driver.find_elements(By.XPATH, '//*[@id="table-mfes-list"]/tbody/tr/td/div/a')
            quantidade = len(tabelaAvisos)
            print(f'{quantidade} avisos indentificados, fazendo limpeza...')
            
            for index in range(quantidade):
                tabelaAvisosAtual = driver.find_elements(By.XPATH, '//*[@id="table-mfes-list"]/tbody/tr/td/div/a')
                for item in tabelaAvisosAtual:
                    time.sleep(1)
                    item.click()
                    time.sleep(2)
                    driver.find_element(By.XPATH, '//*[@id="mfe-manufacturer-detail"]/div[3]/button').click()
                    time.sleep(2)
            time.sleep(2)
            fechar = driver.find_element(By.XPATH, '//*[@id="conteudo_central"]/div/div[2]/div/div/div[3]/button')
            fechar.click()
            time.sleep(3)
        except:
            print('Nenhum alerta no ambiente seguro') 
            
        analisexml = analisadorXmls(listaCFEtotal)
        for index, cupom in enumerate (analisexml, start=1):
            try:
                if index % 1998 == 0:
                    sairAmbienteSeguro(driver)
                    time.sleep(2)
                    driver.get('https://servicos.sefaz.ce.gov.br/internet/acessoseguro/servicosenha/logarusuario/login.asp')
                    time.sleep(2)
                    realizar_login(driver)
                    
                    iniciar_processo(driver, inscricao)
                    time.sleep(10)
                    print('ERA PARA ELE MUDAR AS ABASSSSS')
                    # Lista todas as janelas/abas abertas
                    window_handles = driver.window_handles

                    # Mantém apenas a última aba aberta
                    for handle in window_handles[:-1]:
                        driver.switch_to.window(handle)
                        driver.close()

                    # Troca para a última aba
                    driver.switch_to.window(window_handles[-1])
                    
                # Lógica de atualização e limpeza de cache
                if index % 50 == 0:  # A cada 50 itens
                    
                    print("Limpando o cache e atualizando a página...")
                    
                    autoit.send("^+r")
                    time.sleep(3)
                    driver.refresh()
                    time.sleep(7)
                    print("Página recarregada com cache limpo.")
                    
                # Continue com as operações normais no loop
                time.sleep(0.1)  # Simulação de tempo entre iterações
                comeca_consulta(driver,cupom)
                
                clicar_links_tabela(driver, index)            
                # Mostrar o progresso
                print(f"Baixando {index} de {len(analisexml)}")
                
                
                # Aqui você coloca o processamento para cada item
                # Por exemplo:
                print(f"Processando: {cupom}")
            except Exception as e:
                print(f"O {index} foi interrompido, tentando corrigir problema.")
                driver.refresh()
                time.sleep(15)  # Simulação de tempo entre iterações
                comeca_consulta(driver,cupom)
                
                clicar_links_tabela(driver)            
                # Mostrar o progresso
                
                
                # Aqui você coloca o processamento para cada item
                # Por exemplo:
                print(f"Processando: {cupom}")
                    
            
    except Exception as e:
        print(f"Erro de carregamento infinito, o ambiente seguro está instável")
         
        return None
    
    if len(analisexml) == 0:
        return False
    else:
        return True
        

def baixarCancelamento(driver, decisao):
    if decisao:
        try:
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
            inicioperiodo.send_keys(variavel_inicio)
            
            time.sleep(0.4)
            finalperiodo = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="form-end-date-search-coupons"]'))
                )
            time.sleep(0.2)
            finalperiodo.send_keys(variavel_final)
            
            consulta = WebDriverWait(driver, 200).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div/div/div[3]/form/div[6]/div/div/button[1]'))
                )
            consulta.click()
            
            time.sleep(5)
            
        except:
            print('Erro ao baixar os cupons de cancelamento, ambiente seguro instável')
    
    else:
        return None        
     
def sairAmbienteSeguro(driver):
    try:
        try:
            # Encontre a quarta <li> dentro da ul com o id 'menulist_root'
            fourth_li = driver.find_element(By.XPATH, '//*[@id="menulist_root"]/li[4]')

            # Agora encontre o link <a> dentro desse quarto <li>
            link = fourth_li.find_element(By.TAG_NAME, 'a')
            
            
            link.click()
            
            time.sleep(3)
            
            #SAINDO DO SISTEMA DE FORMA CORRETA
            
            sair2 = WebDriverWait(driver, 50).until(
                EC.presence_of_all_elements_located((By.XPATH, '//*[@id="menulist_root"]/div[5]/li')) 
            )
            
            sair2Sim = sair2[1].find_element(By.TAG_NAME, 'a')
            sair2Sim.click()
            
            time.sleep(3)
            
            sairConfirma = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo_central"]/div/div[2]/div/div/div[3]/button[1]')) 
            )
            
            sairConfirma.click()
            time.sleep(3)
        except:
            print('Indo para etapa 2 DA SAIDA AMBIENTE SEGURO ')   
            
             
        driver.get('https://www.sefaz.ce.gov.br/ambiente-seguro/')
        
        time.sleep(2)
        
        login = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH , '//*[@id="main"]/section/div/div/ul/li[1]/a'))
        )
        login.click()
        
        time.sleep(3)
        
        sairAS = WebDriverWait(driver,50).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="textoContainer"]/form/table/tbody/tr[2]/td/input'))
        )
        sairAS.click()
        
    except:
        
        print('Ambiente seguro já foi desconectado de forma correta, continuando')    

if validacao == True:   
    try:    
        # Configuração do Chrome
        service = Service(ChromeDriverManager().install())
        options = uc.ChromeOptions()

        
        # Obtém o caminho para a pasta do usuário
        user_data_dir = os.path.join(
            os.path.expanduser("~"), 
            "AppData", "Local", "Google", "Chrome", "User Data"
        )

        # Adiciona o argumento ao Selenium
        options.add_argument(f"--user-data-dir={user_data_dir}")
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
    except:
        print('Erro com o navegador Google Chorme, atualize o navegador e tente novamente...')
        time.sleep(10)

        
    try:
        pegarForAmbiente(driver)
        for indice, item in enumerate(listaTotal):
            print(f'Atualizado {indice + 1} de {len(listaTotal)} empresas...')
            listaCFEtotal.clear()
            inicio = entrarDTE(driver,item)
            if inicio == True:
                if inicio == True:
                    tratarCSV(downloads_directory, 'autorizados')
                    
                metade = BaixarOsCancelados(driver)
                
                if(metade == True):
                    time.sleep(5)
                    tratarCSV(downloads_directory, 'cancelados')
                
                if(inicio == True or metade == True ):
                    # Abra a página desejada
                    driver.get("https://servicos.sefaz.ce.gov.br/internet/AcessoSeguro/ServicoSenha/logarusuario/login.asp")
                    time.sleep(2)
                    realizar_login(driver)
                    # teste = "69077339"
                    iniciar_processo(driver, item)
                    continueC = iniciarDownloads(driver, item)
                    baixarCancelamento(driver, continueC)
                    sairAmbienteSeguro(driver)
                else:
                    print('Empresa não tem cupons, indo para a próxima.')
                    time.sleep(6)
            else:
                print('Empresa não tem cupons, indo para a próxima.')
                sairDte(driver)
            time.sleep(5)        #Saindo do login para depois logar de novo    
            if inicio == True:
                print('Todos os cupons baixados, indo para a próxima empresa.')
        
        
        time.sleep(5)

    finally:
        print('Fechando navegador')
        # Feche o navegador após a execução
        driver.quit()

    time.sleep(2)
    print("Processo finalizado... ")
    time.sleep(10)
else:
    print('A chave de licença do software é inválida, por gentileza, entre em contato com o responsável.')