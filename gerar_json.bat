@echo off
chcp 65001 > nul
echo.
echo  =============================================
echo   Agroquima -- Gerador de Painel Orcamentario
echo  =============================================
echo.
echo  Arquivos necessarios nesta pasta:
echo    - Mapa_Despesa.xlsx
echo    - Base_regionais.xlsx
echo    - gerar_json.py
echo.
echo  Gerando aqm_data.json...
echo.

python gerar_json.py

echo.
if %ERRORLEVEL% NEQ 0 (
    echo  ERRO: verifique se o Python esta instalado.
    echo  Baixe em: https://www.python.org/downloads/
    echo  Marque a opcao "Add Python to PATH" durante a instalacao.
    echo.
    echo  Tambem e necessario instalar as dependencias:
    echo    pip install pandas openpyxl
)
pause
