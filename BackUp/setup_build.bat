@echo off
echo =========================================
echo SETUP - Sistema de Gestao Financeira
echo =========================================
echo.

REM Navegar para o diretório do projeto
cd /d "C:\Users\Obras\Documents\GitHub\Sistema-Financeiro"

echo Diretorio atual: %CD%
echo.

REM Verificar se Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado!
    echo Instale Python antes de continuar.
    pause
    exit /b 1
)

echo ✅ Python encontrado
python --version

REM Verificar se pip está disponível
pip --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: pip nao encontrado!
    pause
    exit /b 1
)

echo ✅ pip encontrado
echo.

REM Instalar/atualizar PyInstaller
echo 📦 Instalando PyInstaller...
pip install --upgrade pyinstaller

REM Instalar dependências específicas se necessário
echo 📦 Verificando dependências...
pip install --upgrade pillow xlwings babel python-dateutil validate-docbr tkcalendar reportlab pandas numpy matplotlib openpyxl python-dotenv

echo.
echo =========================================
echo EXECUTANDO BUILD
echo =========================================

REM Executar o script de build
python build_script.py

echo.
echo =========================================
echo BUILD CONCLUIDO
echo =========================================

REM Verificar se o executável foi criado
if exist "dist\Sistema_Gestao_Financeira.exe" (
    echo ✅ Executavel criado com sucesso!
    echo 📁 Localizacao: %CD%\dist\Sistema_Gestao_Financeira.exe
    
    REM Perguntar se quer executar o teste
    echo.
    set /p test_exe="Deseja testar o executavel agora? (s/n): "
    if /i "%test_exe%"=="s" (
        echo 🚀 Iniciando teste...
        start "" "dist\Sistema_Gestao_Financeira.exe"
    )
) else (
    echo ❌ ERRO: Executavel nao foi criado!
    echo Verifique os erros acima.
)

echo.
pause