#  SCRIPT DE CONSTRUCCIÓN AUTOMATIZADO PARA IMPRESORA ETIQUETAS SAPAL
# =====================================================================

@echo off
echo.
echo ========================================
echo   CONSTRUYENDO EJECUTABLE SAPAL
echo ========================================
echo.

# Limpiar builds anteriores
echo  Limpiando builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

echo  Archivos anteriores eliminados
echo.

# Verificar archivos necesarios
echo  Verificando archivos necesarios...
if not exist main.py (
    echo  ERROR: main.py no encontrado
    pause
    exit /b 1
)

if not exist Diseno.lbx (
    echo   ADVERTENCIA: Diseno.lbx no encontrado - se creará ejecutable sin plantilla
)

if not exist config.key (
    echo   ADVERTENCIA: config.key no encontrado - se creará ejecutable sin configuración
)

echo  Verificación completada
echo.

# Crear ejecutable con PyInstaller
echo  Creando ejecutable con PyInstaller...
echo.

pyinstaller ^
    --name='Impresora_Etiquetas_SAPAL' ^
    --onefile ^
    --windowed ^
    --clean ^
    --hidden-import=win32com.client ^
    --hidden-import=requests ^
    --hidden-import=base64 ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.messagebox ^
    --hidden-import=tkinter.simpledialog ^
    --add-data='Diseno.lbx;.' ^
    --add-data='config.key;.' ^
    main.py

echo.
if %errorlevel% equ 0 (
    echo  ¡EJECUTABLE CREADO EXITOSAMENTE!
    echo.
    echo  Ubicación: dist\Impresora_Etiquetas_SAPAL.exe
    echo  Tamaño aproximado: 45-55 MB
    echo.
    
    # Verificar tamaño del archivo
    for %%I in (dist\Impresora_Etiquetas_SAPAL.exe) do echo  Tamaño real: %%~zI bytes
    
    echo.
    echo  SIGUIENTE PASO:
    echo    Copia el ejecutable a cualquier PC Windows
    echo    No necesita Python instalado
    echo    Incluye todos los archivos necesarios
    echo.
    
    # Preguntar si quiere probar el ejecutable
    set /p test='¿Quieres probar el ejecutable ahora? (s/n): '
    if /i '%test%'=='s' (
        echo  Ejecutando aplicación...
        start dist\Impresora_Etiquetas_SAPAL.exe
    )
    
) else (
    echo  ERROR: Falló la creación del ejecutable
    echo  Revisa los mensajes de error arriba
)

echo.
echo ========================================
echo   PROCESO COMPLETADO
echo ========================================
pause
