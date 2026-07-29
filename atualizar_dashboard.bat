@echo off
chcp 65001 >nul
title BR DRIVE — Atualizando Dashboard

echo.
echo  ╔══════════════════════════════════════╗
echo  ║     BR DRIVE — ATUALIZAR DASHBOARD   ║
echo  ╚══════════════════════════════════════╝
echo.

cd /d "%~dp0"

echo  [1/3] Gerando paginas HTML...
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
echo  [2/3] Publicando no Render (GitHub)...
echo.

git add brdrive_output/ dashboard.py
git diff --cached --quiet && (
    echo  Sem alteracoes para publicar. Render ja esta atualizado.
    goto fim
)
git commit -m "Atualizar dashboard %date% %time%"
git push origin main

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [AVISO] Falha ao publicar no GitHub. Dashboard local foi gerado.
    echo.
) else (
    echo.
    echo  [3/3] Publicado! Links:
    echo.
    echo     https://brdrive-dash.onrender.com/home.html
    echo     https://robsondn.github.io/brdrive-dash/home.html
    echo.
)

:fim

echo  Pressione qualquer tecla para fechar...
pause >nul
