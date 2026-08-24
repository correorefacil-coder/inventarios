@echo off
chcp 65001 > nul
title Mas Campo - Servidor Backend (BD SQLite)
echo.
echo  ================================================
echo   MAS CAMPO - SERVIDOR API REST + BASE DE DATOS
echo  ================================================
echo.
echo  Base de datos: backend\mascampo.db
echo  API:           http://localhost:5000
echo.
echo  Iniciando servidor...
echo  (No cierre esta ventana mientras usa el sistema)
echo.
cd /d "%~dp0"
py backend\api.py
echo.
echo  El servidor se detuvo. Presione una tecla para cerrar.
pause > nul
