@echo off
setlocal
cd /d "%~dp0"
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
%PY% -c "import openpyxl" >nul 2>&1
if errorlevel 1 %PY% -m pip install openpyxl
%PY% login_app.py
