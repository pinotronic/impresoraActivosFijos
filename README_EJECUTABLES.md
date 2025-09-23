#  IMPRESORA ETIQUETAS SAPAL - GUÍA COMPLETA DE EJECUTABLES

##  EJECUTABLE YA CREADO Y LISTO PARA USAR

###  Ubicación del Ejecutable:
`
 dist/Impresora_Etiquetas_SAPAL.exe
 Tamaño: ~23 MB
 Portable - No requiere instalación de Python
`

###  Características del Ejecutable:
-  **Completamente Portable**: No necesita Python instalado en el PC de destino
-  **Auto-contenido**: Incluye todas las dependencias (tkinter, win32com, requests, etc.)
-  **Archivos Integrados**: Incluye automáticamente Diseno.lbx y config.key
-  **Interfaz Gráfica**: Abre sin ventana de consola
-  **Verificación OAuth2**: Sistema completo de validación SAPAL
-  **Multiplataforma**: Funciona en cualquier PC Windows (7, 8, 10, 11)

---

##  MÉTODOS PARA CREAR EJECUTABLES

###  MÉTODO 1: Scripts Automatizados (Recomendado)

#### Para Línea de Comandos:
`atch
# Ejecutar el archivo:
crear_ejecutable.bat
`

#### Para Interfaz Gráfica:
`atch
# Ejecutar el archivo:
crear_ejecutable_gui.bat
`

###  MÉTODO 2: PyInstaller Manual

#### Comando Básico:
`ash
pyinstaller --onefile --windowed main.py
`

#### Comando Completo (Recomendado):
`ash
pyinstaller --name="Impresora_Etiquetas_SAPAL" --onefile --windowed --clean --hidden-import=win32com.client --hidden-import=requests --hidden-import=base64 --hidden-import=tkinter --add-data="Diseno.lbx;." --add-data="config.key;." main.py
`

###  MÉTODO 3: auto-py-to-exe (Interfaz Gráfica)

#### Instalación:
`ash
pip install auto-py-to-exe
`

#### Ejecución:
`ash
auto-py-to-exe
`

#### Configuración en la Interfaz:
1. **Script Location**: main.py
2. **Onefile**:  Marcado
3. **Console Window**:  Desmarcado  
4. **Icon**: (Opcional) Seleccionar archivo .ico
5. **Additional Files**: 
   - Agregar: Diseno.lbx
   - Agregar: config.key
6. **Hidden Imports**: 
   - win32com.client
   - equests
   - 	kinter
7. **Name**: Impresora_Etiquetas_SAPAL

#### Importar Configuración:
`ash
auto-py-to-exe --config config_auto_py_to_exe.json
`

---

##  TODO LIST PARA PERSONALIZACIÓN

###  Completado:
- [x]  Ejecutable básico funcional
- [x]  Incluir archivos de datos (Diseno.lbx, config.key)  
- [x]  Optimizar tamaño (~23 MB)
- [x]  Scripts de construcción automatizados
- [x]  Interfaz gráfica con auto-py-to-exe
- [x]  Configuración JSON preestablecida

###  Mejoras Opcionales:
- [ ]  Agregar icono personalizado (.ico)
- [ ]  Crear instalador con NSIS o Inno Setup
- [ ]  Firmar digitalmente el ejecutable
- [ ]  Optimizar tamaño con UPX (compresión)
- [ ]  Auto-actualizador integrado
- [ ]  Crear paquete ZIP para distribución

---

##  CREAR ICONO PERSONALIZADO

### 1. Conseguir imagen PNG/JPG:
- Tamaño recomendado: 256x256 o 512x512 píxeles
- Formato cuadrado

### 2. Convertir a ICO:
- **Online**: usar convertio.co, icoconvert.com
- **Software**: GIMP, Photoshop, o convertidores gratuitos

### 3. Aplicar icono:
`ash
# Agregar --icon="ruta/al/icono.ico" al comando PyInstaller
pyinstaller --icon="icono.ico" --onefile --windowed main.py
`

---

##  CREAR INSTALADOR (Opcional)

### Con NSIS (Nullsoft Scriptable Install System):

1. **Descargar NSIS**: https://nsis.sourceforge.io/
2. **Crear script .nsi**:
`
sis
!define APP_NAME "Impresora Etiquetas SAPAL"
!define APP_EXE "Impresora_Etiquetas_SAPAL.exe"

Name ""
OutFile "Instalador_.exe"
InstallDir "\"

Section "MainSection"
    SetOutPath 
    File "dist\"
    WriteUninstaller "\Uninstall.exe"
    CreateShortcut "\.lnk" "\"
SectionEnd
`

### Con Inno Setup:
1. **Descargar Inno Setup**: https://jrsoftware.org/isinfo.php  
2. **Usar wizard** para configurar instalador gráficamente

---

##  DISTRIBUCIÓN Y USO

### Para Distribuir:
1. **Ejecutable Único**: Copiar dist/Impresora_Etiquetas_SAPAL.exe
2. **Con Archivos**: Asegurar que Diseno.lbx y config.key estén disponibles
3. **Paquete ZIP**: Crear archivo comprimido con ejecutable + manual

### Para Usuarios Finales:
1.  **No necesita instalación de Python**
2.  **Solo ejecutar el archivo .exe**
3.  **Funciona en cualquier PC Windows**
4.  **Antivirus puede alertar**: Es normal con ejecutables de Python

---

##  SOLUCIÓN DE PROBLEMAS

### Antivirus Detecta Como Malware:
`
 SOLUCIÓN: Agregar excepción en antivirus
 CAUSA: PyInstaller genera ejecutables que algunos antivirus marcan como sospechosos
 ALTERNATIVA: Firmar digitalmente el ejecutable
`

### Ejecutable No Abre:
`
 VERIFICAR:
-  Todos los archivos necesarios incluidos
-  Python 3.8+ usado para compilar
-  Dependencias correctas instaladas
-  Rutas relativas (no absolutas) en el código
`

### Error de Módulos:
`
 SOLUCIÓN: Agregar --hidden-import=nombre_modulo
 EJEMPLO: --hidden-import=win32com.client
`

### Archivo Muy Grande:
`
 OPTIMIZACIONES:
- Use --onefile para un solo archivo
- Considere --exclude-module para quitar módulos no usados
- Use UPX para comprimir: --upx-dir=ruta_upx
`

---

##  COMANDOS DE REFERENCIA RÁPIDA

### Reconstruir Ejecutable:
`ash
crear_ejecutable.bat
`

### Interfaz Gráfica:
`ash  
crear_ejecutable_gui.bat
`

### Manual PyInstaller:
`ash
pyinstaller --name="Impresora_Etiquetas_SAPAL" --onefile --windowed --add-data="Diseno.lbx;." --add-data="config.key;." main.py
`

### Limpiar Builds:
`ash
rmdir /s /q build dist
del *.spec
`

---

##  RESULTADO FINAL

 **Tu aplicación ahora es un ejecutable profesional que:**
-  Se ejecuta en cualquier PC Windows sin dependencias
-  Incluye verificación OAuth2 del API SAPAL  
-  Maneja impresión de etiquetas Brother P-touch
-  Tiene interfaz gráfica profesional con validación en tiempo real
-  Es completamente portable y fácil de distribuir

 **¡Listo para usar en producción!**

