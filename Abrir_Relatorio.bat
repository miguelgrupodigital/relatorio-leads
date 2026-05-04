@echo off
title Grupo Digital - Relatório de Performance
echo ================================================
echo   Grupo Digital - Relatório de Performance
echo ================================================
echo.

cd /d "%~dp0"

echo Verificando dependências...
pip install -q -r requirements.txt
echo.

echo Abrindo o navegador e iniciando o app...
echo Feche esta janela para encerrar o servidor.
echo.

python -m streamlit run app.py --server.headless true
pause
