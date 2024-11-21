import os
import locale
import shutil
from datetime import datetime

# Define o locale para português do Brasil
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

# Obtém o diretório de Downloads do usuário
downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")

# Obtém o ano e o mês atuais
ano_atual = datetime.now().year
mes_atual = datetime.now().strftime('%B')
mes_atualNumero = datetime.now().month
nome_mes = f"{mes_atualNumero:02d}-{mes_atual.capitalize()}"

# Loop para criar diretórios de 1 a 10
for inscricao in range(1, 11):
    # Define o caminho completo do diretório
    diretorio = os.path.join(downloads_dir, str(inscricao), str(ano_atual), str(nome_mes))
    
    # Verifica se o diretório já existe, se não, cria
    if not os.path.exists(diretorio):
        os.makedirs(diretorio)
        print(f"Diretório criado: {diretorio}")
    else:
        print(f"Diretório já existe: {diretorio}")
    
    # Converte 'inscricao' para string de dois dígitos
    inscricao_str = f"{inscricao:02d}"
    
    # Lista de arquivos XML no diretório de destino
    arquivos_existentes = set(os.listdir(diretorio))
    
    # Percorre todos os arquivos no diretório de Downloads
    for arquivo in os.listdir(downloads_dir):
        # Verifica se o arquivo tem extensão .xml
        if arquivo.lower().endswith('.xml'):
            # Caminho completo do arquivo
            caminho_arquivo = os.path.join(downloads_dir, arquivo)
            
            # Verifica se o quinto e sexto dígito do nome do arquivo são iguais ao 'inscricao_str'
            if len(arquivo) > 5 and arquivo[4:6] == inscricao_str:
                # Verifica se o arquivo já não está presente no diretório de destino
                if arquivo not in arquivos_existentes:
                    # Move o arquivo para o diretório de destino
                    shutil.move(caminho_arquivo, os.path.join(diretorio, arquivo))
                    print(f"Arquivo {arquivo} movido para {diretorio} porque estava faltando")
                else:
                    print(f"Arquivo {arquivo} já está presente no diretório {diretorio}")
