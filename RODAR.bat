@echo off
title Automacao FGTS Digital
cd /d "%~dp0"

where node >nul 2>&1
if %errorlevel% NEQ 0 (
  echo Node.js nao encontrado. Rode primeiro o INSTALAR.bat
  pause
  exit /b
)

if not exist "node_modules" (
  echo Instalando dependencias (npm install)...
  call npm install
)

node fgts.js
echo.
pause
