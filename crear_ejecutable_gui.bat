#  SCRIPT PARA EJECUTAR AUTO-PY-TO-EXE (INTERFAZ GRÁFICA)
# ========================================================

@echo off
echo.
echo ========================================
echo   AUTO-PY-TO-EXE - INTERFAZ GRÁFICA
echo ========================================
echo.

echo  Iniciando interfaz web de auto-py-to-exe...
echo  Se abrirá en tu navegador por defecto
echo.
echo  CONFIGURACIÓN RECOMENDADA:
echo     Script Location: main.py
echo     One File:  (marcado)
echo     Console Window:  (desmarcado)
echo     Additional Files: 
echo      - Diseno.lbx
echo      - config.key
echo     Hidden Imports:
echo      - win32com.client
echo      - requests
echo      - tkinter
echo.

echo  Ejecutando auto-py-to-exe...
auto-py-to-exe

echo.
echo ========================================
echo   INTERFAZ CERRADA
echo ========================================
pause
