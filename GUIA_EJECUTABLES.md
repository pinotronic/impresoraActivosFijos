#  GUÍA COMPLETA PARA CREAR EJECUTABLES DE PYTHON 
# ================================================

##  EJECUTABLE YA CREADO
Tu ejecutable ya está listo en:
 Ubicación: dist/Impresora_Etiquetas_SAPAL.exe
 Tamaño: ~45-55 MB (incluye Python + todas las dependencias)
 Incluye: Archivo plantilla Diseno.lbx y config.key automáticamente

##  TODO LIST PARA CREAR EJECUTABLES

- [x]  Instalar PyInstaller (ya instalado: versión 6.11.1)
- [x]  Crear ejecutable básico con PyInstaller
- [ ]  Instalar auto-py-to-exe (interfaz gráfica)
- [ ]  Crear script de construcción automatizado
- [ ]  Crear ejecutable con icono personalizado
- [ ]  Crear instalador (opcional)

##  MÉTODO 1: PyInstaller (Línea de Comandos)

### Comando Básico:
pyinstaller --onefile --windowed main.py

### Comando Avanzado (el que usamos):
pyinstaller --name='Impresora_Etiquetas_SAPAL' --onefile --windowed --hidden-import=win32com.client --hidden-import=requests --hidden-import=base64 --hidden-import=tkinter --add-data='Diseno.lbx;.' --add-data='config.key;.' main.py

### Explicación de Parámetros:
 --name: Nombre del ejecutable
 --onefile: Un solo archivo .exe (en lugar de carpeta)
 --windowed: Sin ventana de consola (para aplicaciones GUI)
 --hidden-import: Módulos que PyInstaller no detecta automáticamente
 --add-data: Incluir archivos adicionales ('archivo_origen;carpeta_destino')

##  MÉTODO 2: auto-py-to-exe (Interfaz Gráfica)

### Instalación:
pip install auto-py-to-exe

### Uso:
auto-py-to-exe

### Ventajas:
 Interfaz gráfica fácil de usar
 Vista previa del comando PyInstaller
 Configuración guardable en JSON
 Ideal para principiantes

##  MÉTODO 3: Script de Construcción Automatizado

Crearemos un script que:
 Limpia builds anteriores
 Crea el ejecutable
 Verifica que funciona
 Crea un paquete distribuible

