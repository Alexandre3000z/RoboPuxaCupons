import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from art import text2art, art
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support.ui import Select
import undetected_chromedriver as uc
import autoit
from colorama import init, Fore, Style
import os
import glob
import pandas as pd
import calendar
from login import SENHA, USUARIO
from organizador import analisadorXmls, organizarPastas, apagarCSV
from manual import processo1_Dte
import keyboard


# Efeito de digitação
def type_animation(text, delay=0.004):
    for line in text.split("\n"):
        for char in line:
            print(Fore.YELLOW + char, end="", flush=True)  # Mostra caractere por caractere
            time.sleep(delay)
        print()  # Nova linha
        time.sleep(0.004)  # Pequeno delay entre linhas


init(autoreset=True)


ascii_art = text2art("BUSINESS", font="roman")
# Remove linhas em branco
# Remove linhas em branco internas
linhas = [line for line in ascii_art.split("\n") if line.strip() != ""]# Adiciona um pequeno espaçamento acima e abaixo
espacinho = 1  # Define o número de linhas vazias
art_final = "\n" * espacinho + "\n".join(linhas) + "\n" * espacinho

type_animation('=' * 100)
type_animation(art_final)
type_animation('=' * 100)
print('\n')





# Opções para o usuário
print("Escolha uma das opções abaixo:")
print('\n')
print("1- Processo Automatico (Com procuração)")
print("2- Processo Manual (Sem procuração)")


# Valida a escolha do usuário
while True:
    try:
        opcaoInicial = int(input("\nDigite o número da opção desejada (1-2): "))
        if opcaoInicial in range(1, 3):
            break
        else:
            print("Opção inválida! Digite um número entre 1 e 2.")
    except ValueError:
        print("Entrada inválida! Digite apenas números entre 1 e 2.")



if opcaoInicial == 1:
    entradaEscricao = input("Digite as inscrições estaduais separados por vírgula e espaço: ")
    escricoes =  entradaEscricao.split(", ")




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


mes_desejado = int(input("Informe o mês (1-12): "))
ano_desejado = int(input("Informe o ano (ex: 2024): "))
variavel_inicio, variavel_final = gerar_datas_mes(mes_desejado, ano_desejado)




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
    
   
    time.sleep(5)
    
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
    
def BaixarOsCancelados(driver):
    cancelados = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-nfce/div/div/section[2]/div/div/div/ul/li[2]/a'))
    ) 
    cancelados.click()
    
    time.sleep(7)
             
    #SELECIONAR MES E ANO
    selectMes = driver.find_element(By.XPATH, '//*[@id="mes_select"]')
    optionMes = Select(selectMes)
    optionMes.select_by_index(mes_desejado - 1)
    
    selectAno = driver.find_element(By.XPATH, '//*[@id="ano_select"]')
    optionAno = Select(selectAno)
    optionAno.select_by_value(f'{ano_desejado}')
    
    
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
                driver.refresh()
                time.sleep(7)
                print(f"Erro ao processar linha {indice}: {e}")
                continue

    except Exception as e:
        driver.refresh()
        time.sleep(7)
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
        print("Timeout: o elemento não adquiriu a classe 'ng-hide' dentro do tempo esperado.")
        
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
def sairAmbienteSeguro(driver):
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
                        
           
def iniciarDownloads(driver, inscricao):
    time.sleep(13)    
    # Captura as abas abertas no navegador
    abas = driver.window_handles
    if(opcao == 7):
        janelas = 2
    else:
        janelas = 3    
        
    # Certifica-se de que há pelo menos 3 ou 2 abas abertas
    if len(abas) >= janelas:
        # Troca para a última aberta
        driver.switch_to.window(abas[janelas-1])  # O índice 1 representa a segunda aba
    else:
        print("A terceira aba não foi encontrada.")

    if(opcao != 7):
        try:
            # print('esses são os xmls atuais',listaCFEtotal[2:10])
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
                    
                    clicar_links_tabela(driver)            
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
        
def baixarCancelamento(driver):
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
    
#O CÓDIGO COMEÇA A SER EXECUTADO AQUI    
    
    
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
time.sleep(2)
autoit.send('{F11}')
time.sleep(2)

    
try:
    if(opcaoInicial == 1):
        for item in escricoes:
            if(opcao != 7):
                autorizados = entrarDTE(driver,item)
                time.sleep(5)
                if opcao in (1, 2, 4, 5):
                    if(autorizados == True):
                        tratarCSV(downloads_directory, 'autorizados')
                    time.sleep(1.5)
                    apagarCSV(downloads_directory)
                    
                cancelados = BaixarOsCancelados(driver)
                time.sleep(5)
                if opcao in (1, 2, 3, 6):
                    if(cancelados == True):
                        tratarCSV(downloads_directory, 'cancelados')
                    time.sleep(1.5)
                    apagarCSV(downloads_directory)

            # Abra a página desejada
            driver.get("https://servicos.sefaz.ce.gov.br/internet/AcessoSeguro/ServicoSenha/logarusuario/login.asp")
            
            time.sleep(2)
            realizar_login(driver)
            # teste = "69077339"
            iniciar_processo(driver, item)
            iniciarDownloads(driver, item)
            
            if opcao in (1, 3, 4, 7):
                baixarCancelamento(driver)
                
            blocos = driver.find_elements(By.XPATH,'//*[@id="conteudo_central"]/div/div/div/div[3]/div/div[2]/div/div/div/div/button')
            if blocos:
                for index, item in enumerate(blocos):
                    if index == len(blocos) - 1:
                        item.click()
                        time.sleep(5)            

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
        
    #ENTRANDO NO MODO MANUAL 
    else:
        if(opcao != 7):
            
            autorizados = processo1_Dte(driver)
            inscricao = autorizados[1]
            print(f'Incrição estadual: {inscricao}')
            time.sleep(5)
            if opcao in (1, 2, 4, 5):
                if(autorizados[0] == True):
                    tratarCSV(downloads_directory, 'autorizados')
                time.sleep(1.5)
                apagarCSV(downloads_directory)
                
            cancelados = BaixarOsCancelados(driver)
            time.sleep(5)
            if opcao in (1, 2, 3, 6):
                if(cancelados[0] == True):
                    tratarCSV(downloads_directory, 'cancelados')
                time.sleep(1.5)
                apagarCSV(downloads_directory)

        # Abra a página desejada
        driver.get("https://servicos.sefaz.ce.gov.br/internet/AcessoSeguro/ServicoSenha/logarusuario/login.asp")
        
        time.sleep(2)
        
        print("Entre no Ambiente Seguro com seu Login ou certificado e após isso pressione F2")
        keyboard.wait("f2")
        print("F2 pressionado, continuando o processo...")
        # teste = "69077339"
        iniciar_processo(driver, inscricao)
        iniciarDownloads(driver, inscricao)
        
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
    
    
    time.sleep(10)

finally:
 
    # Feche o navegador após a execução
    driver.quit()

time.sleep(2)
print("Processo finalizado... ")
time.sleep(10)
