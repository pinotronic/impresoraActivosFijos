import tkinter as tk
from tkinter import ttk, messagebox
import win32com.client
import threading
import os

class LabelPrinterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Impresora de Etiquetas b-PAC")
        self.root.geometry("500x400")
        self.root.resizable(True, True)
        
        # Configurar el tema
        style = ttk.Style()
        style.theme_use('clam')
        
        # Variables
        self.inicio_var = tk.StringVar()
        self.hasta_var = tk.StringVar()
        
        # Ruta de la plantilla (en la misma carpeta que el script)
        self.template_path = os.path.join(os.path.dirname(__file__), "Diseno.lbx")
        
        self.setup_ui()
        
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar peso de las columnas y filas
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="🏷️ Impresora de Etiquetas", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Campo "Inicio"
        ttk.Label(main_frame, text="Inicio:", font=('Arial', 12)).grid(
            row=1, column=0, sticky=tk.W, pady=(0, 10))
        
        self.inicio_entry = ttk.Entry(main_frame, textvariable=self.inicio_var, 
                                     font=('Arial', 12), width=30)
        self.inicio_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Campo "Hasta"
        ttk.Label(main_frame, text="Hasta:", font=('Arial', 12)).grid(
            row=2, column=0, sticky=tk.W, pady=(0, 10))
        
        self.hasta_entry = ttk.Entry(main_frame, textvariable=self.hasta_var, 
                                    font=('Arial', 12), width=30)
        self.hasta_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Frame para botones
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        # Botón procesar
        self.process_button = ttk.Button(button_frame, text="🖨️ Procesar e Imprimir", 
                                        command=self.process_labels, 
                                        style='Accent.TButton')
        self.process_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botón limpiar
        clear_button = ttk.Button(button_frame, text="🗑️ Limpiar", 
                                 command=self.clear_fields)
        clear_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botón probar conexión
        test_button = ttk.Button(button_frame, text="🔍 Probar Plantilla", 
                                command=self.test_template)
        test_button.pack(side=tk.LEFT)
        
        # Área de texto para logs
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log de Actividad", padding="10")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Texto de log con scrollbar
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, 
                               font=('Consolas', 10))
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar peso para que el log se expanda
        main_frame.rowconfigure(4, weight=1)
        
        # Agregar mensaje inicial
        self.log("🚀 Aplicación iniciada. Lista para imprimir etiquetas.")
        self.log(f"📂 Plantilla: {self.template_path}")
        
        # Focus en el primer campo
        self.inicio_entry.focus()
        
    def log(self, message):
        """Agregar mensaje al log"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_fields(self):
        """Limpiar todos los campos"""
        self.inicio_var.set("")
        self.hasta_var.set("")
        self.log("🗑️ Campos limpiados")
        self.inicio_entry.focus()
        
    def test_template(self):
        """Probar conexión con la plantilla"""
        self.log("🔍 Probando conexión con plantilla...")
        
        try:
            bpacDoc = win32com.client.Dispatch("bpac.Document")
            result = bpacDoc.Open(self.template_path)
            
            if result:
                self.log("✅ Plantilla abierta exitosamente")
                
                # Probar objetos disponibles
                objects_found = []
                test_objects = ["Barcode1", "Text1", "Text2"]
                
                for obj_name in test_objects:
                    try:
                        obj = bpacDoc.GetObject(obj_name)
                        if obj is not None:
                            objects_found.append(obj_name)
                            self.log(f"   ✅ {obj_name}: Disponible")
                        else:
                            self.log(f"   ❌ {obj_name}: No disponible")
                    except Exception as e:
                        self.log(f"   ❌ {obj_name}: Error - {e}")
                
                if objects_found:
                    self.log(f"🎉 Objetos encontrados: {', '.join(objects_found)}")
                else:
                    self.log("⚠️ No se encontraron objetos válidos")
                    
                bpacDoc.Close
                
            else:
                self.log("❌ No se pudo abrir la plantilla")
                messagebox.showerror("Error", "No se pudo abrir la plantilla.\nVerifica que el archivo existe y b-PAC esté instalado.")
                
        except Exception as e:
            self.log(f"❌ Error al probar plantilla: {e}")
            messagebox.showerror("Error", f"Error al conectar con b-PAC:\n{e}")
            
    def print_single_label(self, inicio_text, hasta_text, barcode_value):
        """Imprimir una sola etiqueta"""
        try:
            bpacDoc = win32com.client.Dispatch("bpac.Document")
            
            if not bpacDoc.Open(self.template_path):
                return False, "No se pudo abrir la plantilla"
            
            # Establecer valores en los objetos
            try:
                # Código de barras
                barcode_obj = bpacDoc.GetObject("Barcode1")
                if barcode_obj:
                    barcode_obj.Text = barcode_value
                
                # Texto 1 (Inicio)
                text1_obj = bpacDoc.GetObject("Text1")
                if text1_obj:
                    text1_obj.Text = inicio_text
                
                # Texto 2 (Hasta)
                text2_obj = bpacDoc.GetObject("Text2")
                if text2_obj:
                    text2_obj.Text = hasta_text
                    
            except Exception as e:
                bpacDoc.Close
                return False, f"Error al establecer valores: {e}"
            
            # Imprimir
            try:
                bpacDoc.StartPrint("", 0)
                result = bpacDoc.PrintOut(1, 0)
                
                # Intentar EndPrint, pero no fallar si hay error
                try:
                    bpacDoc.EndPrint()
                except:
                    pass
                    
                if result:
                    bpacDoc.Close
                    return True, "Impresión exitosa"
                else:
                    # Intentar método alternativo
                    alt_result = bpacDoc.DoPrint(1, 0)
                    bpacDoc.Close
                    if alt_result:
                        return True, "Impresión exitosa (método alternativo)"
                    else:
                        return False, "Fallo en ambos métodos de impresión"
                        
            except Exception as e:
                bpacDoc.Close
                return False, f"Error en impresión: {e}"
                
        except Exception as e:
            return False, f"Error general: {e}"
            
    def process_labels(self):
        """Procesar e imprimir etiquetas"""
        inicio = self.inicio_var.get().strip()
        hasta = self.hasta_var.get().strip()
        
        # Validaciones
        if not inicio:
            messagebox.showerror("Error", "El campo 'Inicio' es obligatorio")
            self.inicio_entry.focus()
            return
            
        if not hasta:
            messagebox.showerror("Error", "El campo 'Hasta' es obligatorio")
            self.hasta_entry.focus()
            return
        
        # Deshabilitar botón durante procesamiento
        self.process_button.config(state='disabled')
        
        # Ejecutar en hilo separado para no bloquear la UI
        thread = threading.Thread(target=self._process_labels_thread, args=(inicio, hasta))
        thread.daemon = True
        thread.start()
        
    def _process_labels_thread(self, inicio, hasta):
        """Hilo para procesar etiquetas sin bloquear la UI"""
        try:

            self.log(f"🚀 Iniciando procesamiento...")
            self.log(f"   📝 Inicio: {inicio}")
            self.log(f"   📝 Hasta: {hasta}")
            
            # Generar código de barras basado en los valores
            barcode_value = f"{inicio}-{hasta}"
            
            self.log(f"   🏷️ Código de barras: {barcode_value}")
            
            # Imprimir etiqueta
            success, message = self.print_single_label(inicio, hasta, barcode_value)
            
            if success:
                self.log(f"✅ {message}")
                self.root.after(0, lambda: messagebox.showinfo("Éxito", "¡Etiqueta impresa exitosamente!"))
            else:
                self.log(f"❌ Error: {message}")
                self.root.after(0, lambda: messagebox.showerror("Error", f"Error al imprimir:\n{message}"))
                
        except Exception as e:
            self.log(f"❌ Error inesperado: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error inesperado:\n{e}"))
            
        finally:
            # Rehabilitar botón
            self.root.after(0, lambda: self.process_button.config(state='normal'))

def main():
    root = tk.Tk()
    app = LabelPrinterGUI(root)
    
    # Configurar el cierre de la aplicación
    def on_closing():
        if messagebox.askokcancel("Salir", "¿Estás seguro que quieres salir?"):
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # Centrar ventana
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()