@echo off
setlocal
cd /d "%~dp0"
echo ==== Creando el ejecutable SistemaAcceso.exe ====
set PY=python
%PY% --version >nul 2>&1
if errorlevel 1 set PY=py
%PY% --version >nul 2>&1
if errorlevel 1 (
  echo No se encontro Python en esta computadora.
  echo Instalalo desde https://www.python.org/downloads/ y marca la casilla "Add Python to PATH".
  pause
  exit /b 1
)
%PY% -m pip install openpyxl pyinstaller
if errorlevel 1 (
  echo No se pudieron instalar las herramientas. Revisa tu conexion a internet.
  pause
  exit /b 1
)
%PY% -m PyInstaller --noconfirm --windowed --name SistemaAcceso login_app.py
if errorlevel 1 (
  echo Ocurrio un error al crear el ejecutable.
  pause
  exit /b 1
)
echo.
echo Listo. La carpeta dist\SistemaAcceso contiene el programa.
echo Comprime esa carpeta en un .zip para entregarla.
explorer dist\SistemaAcceso
pause
