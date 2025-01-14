@echo off
REM Asegúrate de que pip está instalado
python -m ensurepip

REM Actualizar pip a la última versión
python -m pip install --upgrade pip

REM Instalar los módulos necesarios
pip install openpyxl
pip install numpy
pip install pandas
pip install matplotlib

echo Instalación completada.
pause