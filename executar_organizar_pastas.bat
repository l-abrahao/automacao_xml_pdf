@echo off
title Criar estrutura de pastas - Clientes
cd /d "%~dp0"

echo.
echo ============================================================
echo   CRIANDO ESTRUTURA DE PASTAS DOS CLIENTES
echo ============================================================
echo.

python organizar_pastas.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro na execucao.
    pause
)
