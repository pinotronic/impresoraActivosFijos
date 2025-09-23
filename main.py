import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import win32com.client
import requests
import base64
import os
import time
import sys

def get_template_path():
    """Obtiene la ruta de la plantilla en la misma carpeta del script"""
    return os.path.join(os.path.dirname(__file__), "Diseno.lbx")

def check_template_exists():
    """Verifica si la plantilla existe"""
    template_path = get_template_path()
    exists = os.path.exists(template_path)
    return exists, template_path

def load_config():
    """Carga la configuración desde config.key"""
    config = {}
    config_path = os.path.join(os.path.dirname(__file__), "config.key")
    
    try:
        with open(config_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if '=' in line and not line.startswith('#'):
                    key, value = line.split('=', 1)
                    config[key.strip()] = value.strip()
        return config
    except Exception as e:
        raise Exception(f"Error leyendo config.key: {e}")

def get_oauth_token():
    """Obtiene token OAuth2 usando client credentials"""
    try:
        config = load_config()
        client_id = config.get('CLIENT_ID')
        client_secret = config.get('CLIENT_SECRET')
        token_url = config.get('TOKEN_URL')
        
        if not all([client_id, client_secret, token_url]):
            raise Exception("Faltan credenciales en config.key")
        
        # Encode credentials para Basic Auth
        authorization = base64.b64encode(
            bytes(f"{client_id}:{client_secret}", "utf-8")
        ).decode("ascii")
        
        headers = {
            "Authorization": f"Basic {authorization}",
            "Content-Type": "application/x-www-form-urlencoded"
        }
        
        body = {
            "grant_type": "client_credentials"
        }
        
        response = requests.post(token_url, data=body, headers=headers, timeout=10)
        
        if response.status_code == 200:
            token_data = response.json()
            return token_data.get('access_token')
        else:
            raise Exception(f"Error obteniendo token: {response.status_code} - {response.text}")
            
    except Exception as e:
        raise Exception(f"Error en autenticación OAuth2: {e}")

def verify_folio(folio_number, access_token):
    """Verifica si un folio existe en el servicio"""
    try:
        config = load_config()
        api_base_url = config.get('API_BASE_URL')
        
        if not api_base_url:
            raise Exception("API_BASE_URL no configurada en config.key")
        
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        # Construir URL para consultar folio específico
        folio_url = f"{api_base_url}/{folio_number}"
        
        response = requests.get(folio_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            # Verificar el contenido de la respuesta
            try:
                data = response.json()
                items = data.get('items', [])
                
                if items and len(items) > 0:
                    # Folio existe y es válido
                    item = items[0]
                    description = item.get('description', 'Sin descripción')
                    return True, f"Folio válido: {description}"
                else:
                    # Folio no encontrado (array vacío)
                    return False, "Folio no encontrado en el sistema"
                    
            except Exception as e:
                return False, f"Error procesando respuesta: {e}"
        elif response.status_code == 404:
            # Folio no encontrado
            return False, "Folio no encontrado"
        else:
            # Otro error
            return False, f"Error verificando folio: {response.status_code}"
            
    except requests.exceptions.Timeout:
        return False, "Timeout en verificación"
    except requests.exceptions.ConnectionError:
        return False, "Error de conexión"
    except Exception as e:
        return False, f"Error: {e}"

def test_api_connection():
    """Prueba la conexión con el API y las credenciales"""
    try:
        config = load_config()
        result = "PRUEBA DE CONEXIÓN AL API SAPAL:\n\n"
        
        # Mostrar configuración (sin revelar secretos completos)
        client_id = config.get('CLIENT_ID', 'NO CONFIGURADO')
        if len(client_id) > 10:
            client_id_display = client_id[:6] + "..." + client_id[-4:]
        else:
            client_id_display = client_id
            
        result += f"Client ID: {client_id_display}\n"
        result += f"Token URL: {config.get('TOKEN_URL', 'NO CONFIGURADO')}\n"
        result += f"API Base URL: {config.get('API_BASE_URL', 'NO CONFIGURADO')}\n\n"
        
        # Probar obtención de token
        result += "PASO 1: Obteniendo token OAuth2...\n"
        try:
            access_token = get_oauth_token()
            result += " Token obtenido exitosamente\n\n"
            
            # Probar verificación con un folio de ejemplo
            result += "PASO 2: Probando verificación de folio (ejemplo: 322719)...\n"
            is_valid, msg = verify_folio(322719, access_token)
            if is_valid:
                result += f" Folio 322719: {msg}\n"
            else:
                result += f" Folio 322719: {msg}\n"
                
            result += "\nSISTEMA LISTO PARA VERIFICAR FOLIOS\n"
            result += "Las etiquetas solo se imprimirán si los folios existen en el sistema SAPAL."
            
        except Exception as e:
            result += f" Error obteniendo token: {e}\n"
            result += "\nREVISAR:\n"
            result += "1. Credenciales en config.key\n"
            result += "2. Conexión a internet\n"
            result += "3. URLs del servicio\n"
        
        return result
        
    except Exception as e:
        return f"Error probando conexión: {e}"

def diagnose_image_problem():
    try:
        # Verificar que la plantilla existe
        exists, template_path = check_template_exists()
        if not exists:
            return f"❌ ERROR: No se encontró el archivo Diseno.lbx en:\n{template_path}\n\nAsegúrate de que el archivo esté en la misma carpeta que este script."
        
        bpacDoc = win32com.client.Dispatch("bpac.Document")
        
        if not bpacDoc.Open(template_path):
            return f"No se pudo abrir la plantilla en: {template_path}"
        
        result = "DIAGNOSTICO DE IMAGEN EN PLANTILLA:\n\n"
        
        image_names = ["Image1", "Image2", "Picture1", "Picture2", "Logo", "LOGO", "Imagen"]
        
        images_found = 0
        problems = []
        
        for img_name in image_names:
            try:
                img_obj = bpacDoc.GetObject(img_name)
                if img_obj is not None and getattr(img_obj, 'Type', -1) == 2:
                    images_found += 1
                    result += f"IMAGEN ENCONTRADA: {img_name}\n"
                    
                    if hasattr(img_obj, 'BackColor'):
                        back_color = getattr(img_obj, 'BackColor', 0)
                        if back_color == 0:
                            problems.append(f"{img_name}: Fondo negro")
                            result += "   PROBLEMA: Fondo negro detectado\n"
                        else:
                            result += f"   Color fondo: {back_color}\n"
                    
                    if hasattr(img_obj, 'Transparent'):
                        transparent = getattr(img_obj, 'Transparent', False)
                        if transparent:
                            problems.append(f"{img_name}: Transparencia activada")
                            result += "   PROBLEMA: Transparencia activada\n"
                        else:
                            result += "   Transparencia: Desactivada\n"
                    
                    result += "\n"
            except:
                continue
        
        if images_found == 0:
            result += "No se detectaron imagenes en la plantilla\n\n"
            result += "POSIBLES SOLUCIONES:\n"
            result += "1. La imagen puede tener un nombre diferente\n"
            result += "2. Usar 'Explorar Objetos' para ver todos los elementos\n"
        else:
            result += f"RESUMEN: {images_found} imagenes encontradas\n"
            result += f"Problemas detectados: {len(problems)}\n\n"
            
            if problems:
                result += "SOLUCIONES PARA IMAGEN NEGRA:\n\n"
                result += "PASO 1 - Abrir Brother P-touch Editor:\n"
                result += "1. Abrir archivo: Diseno.lbx\n"
                result += "2. Hacer clic derecho en la imagen  Propiedades\n\n"
                
                result += "PASO 2 - Configurar imagen:\n"
                result += " Cambiar 'Transparente' a 'No'\n"
                result += " Establecer color de fondo a 'Blanco'\n"
                result += " Si imagen esta 'invertida', cambiar a 'Normal'\n\n"
                
                result += "PASO 3 - Alternativas:\n"
                result += " Editar imagen: PNG transparente  PNG con fondo blanco\n"
                result += " Cambiar formato: PNG  JPG (sin transparencia)\n"
                result += " Verificar que imagen no este en negativo\n"
        
        bpacDoc.Close
        return result
        
    except Exception as e:
        return f"Error en diagnostico: {e}"

def print_single_folio(folio_number, access_token=None):
    """Imprime un folio individual después de verificarlo"""
    try:
        # Verificar folio primero si tenemos token
        if access_token:
            is_valid, verification_msg = verify_folio(folio_number, access_token)
            if not is_valid:
                return False, f"Folio {folio_number} no válido: {verification_msg}"
        
        # Verificar que la plantilla existe
        exists, template_path = check_template_exists()
        if not exists:
            return False, f"❌ No se encontró Diseno.lbx en: {template_path}"
        
        # Proceder con impresión si el folio es válido
        bpacDoc = win32com.client.Dispatch("bpac.Document")
        
        if not bpacDoc.Open(template_path):
            return False, f"No se pudo abrir plantilla para folio {folio_number} en: {template_path}"
        
        folio_str = str(folio_number)
        bpacDoc.GetObject("Barcode1").Text = folio_str
        bpacDoc.GetObject("Text2").Text = folio_str
        
        try:
            bpacDoc.StartPrint("", 0)
            result = bpacDoc.PrintOut(1, 0)
            try:
                bpacDoc.EndPrint()
            except:
                pass
            
            bpacDoc.Close
            return True, f"Folio {folio_number} verificado e impreso"
            
        except Exception as e:
            try:
                bpacDoc.DoPrint(1, 0)
                bpacDoc.Close
                return True, f"Folio {folio_number} verificado e impreso (método alternativo)"
            except:
                bpacDoc.Close
                return False, f"Error imprimiendo folio {folio_number}: {e}"
                
    except Exception as e:
        return False, f"Error general folio {folio_number}: {e}"

def print_folio_range_with_logging(inicio, hasta, log_callback, root_widget=None):
    """Imprime rango de folios con verificación previa y logging"""
    try:
        folio_inicio = int(inicio)
        folio_hasta = int(hasta)
        
        if folio_inicio > folio_hasta:
            return False, "Folio inicio mayor que folio final"
        
        total = folio_hasta - folio_inicio + 1
        log_callback(f"=== INICIANDO VERIFICACIÓN DE FOLIOS ===")
        log_callback(f"Rango: {folio_inicio} - {folio_hasta} (Total: {total} folios)")
        log_callback("")
        
        # Obtener token OAuth2
        try:
            log_callback("🔑 Obteniendo token de autenticación...")
            access_token = get_oauth_token()
            log_callback("✅ Token obtenido exitosamente")
            log_callback("")
        except Exception as e:
            error_msg = f"❌ Error obteniendo token: {e}"
            log_callback(error_msg)
            return False, error_msg
        
        # Verificar folios primero
        log_callback("🔍 VERIFICANDO FOLIOS:")
        log_callback("")
        folios_validos = []
        folios_invalidos = []
        
        for folio in range(folio_inicio, folio_hasta + 1):
            log_callback(f"Verificando folio {folio}...")
            is_valid, msg = verify_folio(folio, access_token)
            
            if is_valid:
                folios_validos.append(folio)
                log_callback(f"  ✅ Folio {folio}: {msg}")
            else:
                folios_invalidos.append((folio, msg))
                log_callback(f"  ❌ Folio {folio}: {msg}")
            
            # Pequeña pausa para que se vea el progreso
            if root_widget:
                root_widget.update_idletasks()
            time.sleep(0.1)
        
        log_callback("")
        log_callback(f"📊 RESUMEN DE VERIFICACIÓN:")
        log_callback(f"  • Folios válidos: {len(folios_validos)}")
        log_callback(f"  • Folios inválidos: {len(folios_invalidos)}")
        log_callback("")
        
        if not folios_validos:
            error_msg = f"❌ Ningún folio válido encontrado en el rango {folio_inicio}-{folio_hasta}"
            log_callback(error_msg)
            return False, error_msg
        
        # Imprimir solo folios válidos
        log_callback("🖨️ IMPRIMIENDO FOLIOS VÁLIDOS:")
        log_callback("")
        exitosos = 0
        fallidos = 0
        detalles = []
        
        for folio in folios_validos:
            log_callback(f"Imprimiendo folio {folio}...")
            success, msg = print_single_folio(folio)  # Sin verificación nuevamente
            detalles.append(f"Folio {folio}: {msg}")
            if success:
                exitosos += 1
                log_callback(f"  ✅ {msg}")
            else:
                fallidos += 1
                log_callback(f"  ❌ {msg}")
        
        # Agregar información de folios inválidos
        for folio, reason in folios_invalidos:
            detalles.append(f"Folio {folio}: OMITIDO - {reason}")
        
        log_callback("")
        log_callback(f"📋 RESULTADO FINAL:")
        log_callback(f"  • Total verificados: {total}")
        log_callback(f"  • Válidos: {len(folios_validos)}, Inválidos: {len(folios_invalidos)}")
        log_callback(f"  • Impresos exitosamente: {exitosos}, Fallos de impresión: {fallidos}")
        log_callback("=== PROCESO COMPLETADO ===")
        
        resumen = f"Total verificados: {total}\n"
        resumen += f"Válidos: {len(folios_validos)}, Inválidos: {len(folios_invalidos)}\n"
        resumen += f"Impresos exitosamente: {exitosos}, Fallos de impresión: {fallidos}\n\n"
        resumen += "\n".join(detalles)
        
        return exitosos > 0, resumen
        
    except ValueError:
        return False, "Valores deben ser números enteros"
    except Exception as e:
        return False, f"Error: {e}"

def create_gui():
    root = tk.Tk()
    root.title("Impresora de Etiquetas b-PAC con Verificación SAPAL")
    root.geometry("700x600")
    
    inicio_var = tk.StringVar()
    hasta_var = tk.StringVar()
    
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    ttk.Label(main_frame, text="Impresora de Etiquetas SAPAL", 
              font=('Arial', 16, 'bold')).pack(pady=(0, 20))
    
    inicio_frame = ttk.Frame(main_frame)
    inicio_frame.pack(fill=tk.X, pady=5)
    ttk.Label(inicio_frame, text="Folio Inicio:", width=12).pack(side=tk.LEFT)
    inicio_entry = ttk.Entry(inicio_frame, textvariable=inicio_var, font=('Arial', 12))
    inicio_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
    
    hasta_frame = ttk.Frame(main_frame)
    hasta_frame.pack(fill=tk.X, pady=5)
    ttk.Label(hasta_frame, text="Folio Hasta:", width=12).pack(side=tk.LEFT)
    hasta_entry = ttk.Entry(hasta_frame, textvariable=hasta_var, font=('Arial', 12))
    hasta_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
    
    # Área de mensajes de validación
    validation_frame = ttk.LabelFrame(main_frame, text="Mensajes de Validación", padding="10")
    validation_frame.pack(fill=tk.BOTH, expand=True, pady=10)
    
    validation_text = tk.Text(validation_frame, wrap=tk.WORD, font=('Consolas', 9), height=10)
    validation_scroll = ttk.Scrollbar(validation_frame, orient=tk.VERTICAL, command=validation_text.yview)
    validation_text.configure(yscrollcommand=validation_scroll.set)
    
    validation_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    validation_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    
    def log_message(message):
        """Agrega un mensaje al área de validación"""
        validation_text.insert(tk.END, f"{message}\n")
        validation_text.see(tk.END)
        root.update_idletasks()
    
    def clear_messages():
        """Limpia el área de mensajes"""
        validation_text.delete(1.0, tk.END)
    
    # Verificar plantilla al inicio
    exists, template_path = check_template_exists()
    if exists:
        log_message("✅ Plantilla encontrada: Diseno.lbx")
        log_message(f"📁 Ruta: {template_path}")
    else:
        log_message("❌ ADVERTENCIA: No se encontró Diseno.lbx")
        log_message(f"📁 Se busca en: {template_path}")
        log_message("⚠️ Copia el archivo Diseno.lbx a la carpeta del programa")
    
    def procesar_impresion():
        inicio = inicio_var.get().strip()
        hasta = hasta_var.get().strip()
        
        if not inicio or not hasta:
            messagebox.showerror("Error", "Ambos campos son obligatorios")
            return
        
        try:
            folio_inicio = int(inicio)
            folio_hasta = int(hasta)
        except ValueError:
            messagebox.showerror("Error", "Deben ser números enteros")
            return
        
        if folio_inicio > folio_hasta:
            messagebox.showerror("Error", "Folio inicio mayor que folio final")
            return
        
        total = folio_hasta - folio_inicio + 1
        if total > 10:
            respuesta = messagebox.askyesno("Confirmar", 
                f"Se verificarán {total} folios y se imprimirán solo los válidos. ¿Continuar?")
            if not respuesta:
                return
        
        # Limpiar mensajes anteriores
        clear_messages()
        
        # Usar la función con logging
        success, msg = print_folio_range_with_logging(inicio, hasta, log_message, root)
        
        if success:
            messagebox.showinfo("Resultado", f"Proceso completado!\n\nRevisa el área de mensajes para detalles completos.")
            inicio_var.set("")
            hasta_var.set("")
            inicio_entry.focus()
        else:
            messagebox.showerror("Error", f"Error en el proceso:\n\n{msg}")
    
    def limpiar_campos():
        inicio_var.set("")
        hasta_var.set("")
        inicio_entry.focus()
    
    def limpiar_mensajes():
        clear_messages()
        log_message("Mensajes limpiados.")
    
    def probar_folio_individual():
        """Prueba un solo folio para ver la respuesta del API"""
        folio_test = tk.simpledialog.askstring("Probar Folio", "Ingrese un folio para probar:")
        
        if folio_test:
            try:
                folio_num = int(folio_test.strip())
                clear_messages()
                log_message(f"=== PRUEBA DE FOLIO INDIVIDUAL ===")
                log_message(f"Probando folio: {folio_num}")
                log_message("")
                
                # Obtener token
                log_message("🔑 Obteniendo token...")
                try:
                    access_token = get_oauth_token()
                    log_message("✅ Token obtenido")
                except Exception as e:
                    log_message(f"❌ Error obteniendo token: {e}")
                    return
                
                # Verificar folio
                log_message(f"🔍 Verificando folio {folio_num}...")
                is_valid, msg = verify_folio(folio_num, access_token)
                
                if is_valid:
                    log_message(f"✅ RESULTADO: {msg}")
                    log_message("   ⚠️ ESTE FOLIO SE IMPRIMIRÍA")
                else:
                    log_message(f"❌ RESULTADO: {msg}")
                    log_message("   ℹ️ Este folio NO se imprimiría")
                
                log_message("=== FIN DE PRUEBA ===")
                
            except ValueError:
                messagebox.showerror("Error", "Debe ingresar un número válido")
    
    def diagnosticar_imagen():
        respuesta = messagebox.askyesno("Diagnostico", 
            "Analizar imagenes en la plantilla para identificar\n"
            "por que salen en negro. ¿Continuar?")
        
        if respuesta:
            resultado = diagnose_image_problem()
            ventana = tk.Toplevel(root)
            ventana.title("Diagnostico de Imagen")
            ventana.geometry("700x500")
            
            frame = ttk.Frame(ventana, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            texto = tk.Text(frame, wrap=tk.WORD, font=('Consolas', 10))
            scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=texto.yview)
            texto.configure(yscrollcommand=scroll.set)
            
            texto.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scroll.pack(side=tk.RIGHT, fill=tk.Y)
            
            texto.insert(tk.END, resultado)
            texto.config(state=tk.DISABLED)
            
            ttk.Button(ventana, text="Cerrar", command=ventana.destroy).pack(pady=10)
    
    def probar_conexion():
        respuesta = messagebox.askyesno("Probar Conexión", 
            "Probar conexión con el API SAPAL y verificar credenciales.\n"
            "¿Continuar?")
        
        if respuesta:
            resultado = test_api_connection()
            ventana = tk.Toplevel(root)
            ventana.title("Prueba de Conexión API SAPAL")
            ventana.geometry("700x500")
            
            frame = ttk.Frame(ventana, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            texto = tk.Text(frame, wrap=tk.WORD, font=('Consolas', 10))
            scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=texto.yview)
            texto.configure(yscrollcommand=scroll.set)
            
            texto.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scroll.pack(side=tk.RIGHT, fill=tk.Y)
            
            texto.insert(tk.END, resultado)
            texto.config(state=tk.DISABLED)
            
            ttk.Button(ventana, text="Cerrar", command=ventana.destroy).pack(pady=10)
    
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(pady=10)
    
    # Primera fila de botones
    button_row1 = ttk.Frame(button_frame)
    button_row1.pack(pady=5)
    
    ttk.Button(button_row1, text="Imprimir", 
               command=procesar_impresion).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_row1, text="Limpiar Campos", 
               command=limpiar_campos).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_row1, text="Probar Folio", 
               command=probar_folio_individual).pack(side=tk.LEFT, padx=5)
    
    # Segunda fila de botones
    button_row2 = ttk.Frame(button_frame)
    button_row2.pack(pady=5)
    
    ttk.Button(button_row2, text="Probar API", 
               command=probar_conexion).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_row2, text="Limpiar Mensajes", 
               command=limpiar_mensajes).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_row2, text="Diagnosticar Imagen", 
               command=diagnosticar_imagen).pack(side=tk.LEFT, padx=5)
    
    info_text = "FUNCIONAMIENTO CON VERIFICACIÓN SAPAL:\n"
    info_text += "• Los folios se verifican contra el API SAPAL antes de imprimir\n"
    info_text += "• Solo se imprimen etiquetas de folios válidos\n"
    info_text += "• Use 'Probar Folio' para verificar un folio específico\n"
    info_text += "• El área de mensajes muestra el proceso completo de verificación"
    
    ttk.Label(main_frame, text=info_text, justify=tk.LEFT, 
              font=('Arial', 9)).pack(pady=10)
    
    inicio_entry.focus()
    root.bind('<Return>', lambda e: procesar_impresion())
    
    return root

if __name__ == "__main__":
    root = create_gui()
    root.mainloop()
