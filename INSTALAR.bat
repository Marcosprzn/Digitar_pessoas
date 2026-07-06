@echo off
REM ============================================================
REM  Instalador de dependencias - Automacao FGTS Digital
REM  Da dois cliques neste arquivo.
REM ============================================================
title Instalador FGTS Digital

echo.
echo Iniciando instalador... (pode pedir permissao de Administrador)
echo.

REM Tenta elevar para Administrador (necessario para instalar o navegador)
net session >nul 2>&1
if %errorlevel% NEQ 0 (
  echo Solicitando privilegios de Administrador...
  powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'"
  exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar_dependencias.ps1"

echo.
echo ============================================================
echo  Fim. Leia as mensagens acima e o arquivo de log gerado.
echo ============================================================
pause
