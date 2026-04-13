@echo off
set BASE=%~dp0
set LOG=%BASE%log.txt

echo ==== %date% %time% ==== >> "%LOG%"

start "" outlook.exe
timeout /t 20 >nul

python -u "%BASE%automatizacao_pegar_planilhas.py" >> "%LOG%" 2>&1
if errorlevel 1 (echo ERRO NO DOWNLOAD >> "%LOG%" & exit /b 1)

python -u "%BASE%merge_planilhas.py" >> "%LOG%" 2>&1
if errorlevel 1 (echo ERRO NO MERGE >> "%LOG%" & exit /b 1)

python -u "%BASE%upload_app.py" >> "%LOG%" 2>&1
if errorlevel 1 (echo ERRO NO UPLOAD >> "%LOG%" & exit /b 1)
