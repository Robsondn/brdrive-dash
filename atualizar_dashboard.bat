@echo off
chcp 65001 >nul
title BR DRIVE — Atualizando Dashboard

echo.
echo  ╔══════════════════════════════════════╗
echo  ║     BR DRIVE — ATUALIZAR DASHBOARD   ║
echo  ╚══════════════════════════════════════╝
echo.

cd /d "%~dp0"

echo  [1/2] Gerando paginas HTML...
echo.

py dashboard.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [ERRO] Falha ao gerar o dashboard.
    echo  Verifique se o Python esta instalado e o Excel esta fechado.
    echo.
    pause
    exit /b 1
)

echo.
echo  [2/2] Dashboard atualizado com sucesso!
echo.
echo  Pressione qualquer tecla para fechar...
pause >nul
