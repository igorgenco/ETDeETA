@echo off
set LOG=C:\GencoServer\Genco IT\Planilhas para testes\log.txt

echo ==== %date% %time% ==== >> "%LOG%"

start "" outlook.exe
timeout /t 20 >nul

python -u "C:\GencoServer\Genco IT\Planilhas para testes\automatizacao_pegar_planilhas.py" >> "%LOG%" 2>&1
if errorlevel 1 (echo ERRO NO DOWNLOAD >> "%LOG%" & exit /b 1)

python -u "C:\GencoServer\Genco IT\Planilhas para testes\merge_planilhas.py" >> "%LOG%" 2>&1
if errorlevel 1 (echo ERRO NO MERGE >> "%LOG%" & exit /b 1)

python -u "C:\GencoServer\Genco IT\Planilhas para testes\upload_app.py" >> "%LOG%" 2>&1
if errorlevel 1 (echo ERRO NO UPLOAD >> "%LOG%" & exit /b 1)
