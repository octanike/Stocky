@echo off
title Sistema de Gestion y Stock
echo Verificando e instalando librerias de requirements.txt...
pip install -r requirements.txt

echo.
echo Iniciando la aplicacion...
python app.py

pause