@echo off
echo Verificando se o pip está instalado...

pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo O pip não está instalado ou não está no PATH.
    echo Certifique-se de ter Python e pip instalados no sistema.
    pause
    exit /b 1
)

echo Instalando dependências do arquivo requirements.txt...
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo Ocorreu um erro durante a instalação das dependências.
    pause
    exit /b 1
)

echo Todas as dependências foram instaladas com sucesso!
pause
