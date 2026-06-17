@echo off
chcp 65001 >nul
title BR DRIVE — Servidor Ativo

echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║         BR DRIVE — SERVIDOR DE DASHBOARDS        ║
echo  ╚══════════════════════════════════════════════════╝
echo.
echo  ┌─────────────────────────────────────────────────┐
echo  │  Link fixo para compartilhar (nunca muda):      │
echo  │                                                 │
echo  │    http://%COMPUTERNAME%:8080/home.html         │
echo  │                                                 │
echo  │  Envie esse link para quem quiser visualizar    │
echo  └─────────────────────────────────────────────────┘
echo.
echo  Mantenha esta janela aberta enquanto o servidor estiver rodando.
echo  Para encerrar: feche esta janela ou pressione Ctrl+C
echo.
echo  Aguardando conexoes...
echo.

py "%~dp0servir.py"
