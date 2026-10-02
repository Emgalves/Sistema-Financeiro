@echo off
REM Ajuste a linha abaixo para a pasta raiz do projeto (a que contém "src")
cd /d "C:\Users\SEU_USUARIO\Documents\GitHub\Sistema-Financeiro"
if not exist "%APPDATA%\SistemaFinanceiro\gmail" mkdir "%APPDATA%\SistemaFinanceiro\gmail"
python -m src.email_fiscal.coletor >> "%APPDATA%\SistemaFinanceiro\gmail\coletor.log" 2>&1
