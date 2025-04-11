@echo off
echo ============================================
echo        SKY LABS PHOTOSHOP - INSTALADOR
echo ============================================
echo.


python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Python nao encontrado! Por favor, instale o Python 3.6 ou superior.
    echo Baixe em: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo Python encontrado. Verificando dependencias...
echo.

echo Verificando pywin32...
python -c "import win32com" >nul 2>&1
if %errorlevel% neq 0 (
    echo Instalando pywin32...
    pip install pywin32
    if %errorlevel% neq 0 (
        echo Tentando modo alternativo...
        python -m pip install pywin32
        if %errorlevel% neq 0 (
            echo Tentando versao especifica...
            python -m pip install pywin32==305
            if %errorlevel% neq 0 (
                echo ERRO: Falha ao instalar pywin32.
                pause
                exit /b 1
            )
        )
    )
    echo Executando pos-instalacao do pywin32...
    python -m pywin32_postinstall -install
) else (
    echo pywin32 ja instalado.
)

echo Verificando pandas...
python -c "import pandas" >nul 2>&1
if %errorlevel% neq 0 (
    echo Instalando pandas...
    pip install pandas
    if %errorlevel% neq 0 (
        python -m pip install pandas
        if %errorlevel% neq 0 (
            echo ERRO: Falha ao instalar pandas.
            pause
            exit /b 1
        )
    )
) else (
    echo pandas ja instalado.
)

echo Verificando openpyxl...
python -c "import openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo Instalando openpyxl...
    pip install openpyxl
    if %errorlevel% neq 0 (
        python -m pip install openpyxl
        if %errorlevel% neq 0 (
            echo ERRO: Falha ao instalar openpyxl.
            pause
            exit /b 1
        )
    )
) else (
    echo openpyxl ja instalado.
)

echo.
echo Todas as dependencias instaladas com sucesso!
echo.

if not exist "lista_nomes.xlsx" (
    echo AVISO: Arquivo "lista_nomes.xlsx" nao encontrado.
    echo Sera necessario especificar o caminho para o arquivo Excel ao executar o programa.
    echo.
)

tasklist | find /i "Photoshop.exe" >nul 2>&1
if %errorlevel% neq 0 (
    echo AVISO: Photoshop nao detectado em execucao.
    echo Por favor, abra o Photoshop e um arquivo PSD antes de continuar.
    echo.
)

echo ============================================
echo              INICIANDO PROGRAMA
echo ============================================
echo.
echo Pressione qualquer tecla para iniciar o aplicativo...
pause >nul

cls
python editar_e_exportar.py
pause 