@echo off
cd /d "%~dp0"

echo Iniciando Gestor Fiscal (NFSe e Sefaz)...

if exist ".venv\Scripts\activate.bat" (
    echo Ativando ambiente virtual...
    call ".venv\Scripts\activate.bat"
) else (
    echo [AVISO] Ambiente virtual nao encontrado.
    echo O aplicativo tentara rodar com o Python global.
)

echo Verificando dependencias...

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] Falha ao instalar as dependencias do requirements.txt
    echo Certifique-se de ter conexao com a internet e que o Python esta instalado.
    pause
    exit /b 1
)

python "src\app_gui.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ------------------------------------------
    echo Erro ao abrir o aplicativo.
    echo ------------------------------------------
    pause
)
