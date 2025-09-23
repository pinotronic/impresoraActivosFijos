import win32com.client

def print_label_with_bpac(template_path, barcode_value):
    """
    Función para imprimir etiquetas usando b-PAC
    
    Args:
        template_path (str): Ruta completa al archivo de plantilla .lbx
        barcode_value (str): Valor del código de barras a imprimir
    """
    bpacDoc = None
    try:
        # Crear el objeto de documento b-PAC
        bpacDoc = win32com.client.Dispatch("bpac.Document")
        
        # Abrir la plantilla
        print(f"Abriendo plantilla: {template_path}")
        result = bpacDoc.Open(template_path)
        if not result:
            print("❌ Error: No se pudo abrir la plantilla")
            return False
            
        print("✅ Plantilla abierta exitosamente")
        
        # Establecer el valor del código de barras
        # Sabemos que solo existe 'Barcode1' en la plantilla
        barcode_obj = bpacDoc.GetObject("Barcode1")
        if barcode_obj is not None:
            barcode_obj.Text = barcode_value
            print(f"✅ Código de barras establecido: {barcode_value}")
        else:
            print("❌ Error: No se pudo encontrar el objeto Barcode1")
            return False
        
        # Imprimir el documento
        print("🖨️ Iniciando impresión...")
        bpacDoc.StartPrint("", 0)  # Usar impresora predeterminada
        bpacDoc.PrintOut(1, 0)     # Imprimir 1 copia
        bpacDoc.EndPrint()
        print("✅ Documento enviado a impresión")
        
        return True
        
    except Exception as e:
        print(f"❌ Error durante la impresión: {e}")
        return False
        
    finally:
        # Cerrar el documento
        if bpacDoc is not None:
            try:
                # Usar la propiedad Close en lugar del método
                bpacDoc.Close
                print("📄 Documento cerrado")
            except Exception as e:
                print(f"⚠️ Advertencia al cerrar documento: {e}")

def discover_template_objects(template_path):
    """
    Función para descubrir qué objetos están disponibles en una plantilla
    """
    bpacDoc = None
    try:
        bpacDoc = win32com.client.Dispatch("bpac.Document")
        
        if not bpacDoc.Open(template_path):
            print("❌ No se pudo abrir la plantilla para exploración")
            return []
            
        # Lista ampliada de posibles nombres de objetos
        possible_names = [
            # Códigos de barras
            "Barcode1", "Barcode2", "Barcode3", "BarCode1", "BARCODE1",
            # Textos
            "Text1", "Text2", "Text3", "TEXT1", "TEXT2", "TEXT3",
            "Label1", "Label2", "Label3", "LABEL1", "LABEL2", "LABEL3",
            # Códigos
            "Code", "CODE", "Codigo", "CODIGO", "codigo",
            # Campos genéricos
            "Field1", "Field2", "Field3", "FIELD1", "FIELD2", "FIELD3",
            "Object1", "Object2", "Object3", "OBJECT1", "OBJECT2", "OBJECT3",
            # Otros nombres comunes
            "Title", "TITLE", "Description", "DESCRIPTION", "Name", "NAME"
        ]
        
        found_objects = []
        
        for name in possible_names:
            try:
                obj = bpacDoc.GetObject(name)
                if obj is not None:
                    obj_type = getattr(obj, 'Type', 'Desconocido')
                    found_objects.append((name, obj_type))
                    print(f"✅ Objeto encontrado: '{name}' (Tipo: {obj_type})")
            except:
                pass  # Silenciosamente ignorar errores
                
        return found_objects
        
    except Exception as e:
        print(f"❌ Error durante la exploración: {e}")
        return []
        
    finally:
        if bpacDoc is not None:
            try:
                bpacDoc.Close
            except:
                pass

if __name__ == "__main__":
    # Configuración
    template_path = r"C:\Users\pvargas\OneDrive - sapal365\Documentos\My Labels\Diseno.lbx"
    
    print("🔍 === EXPLORACIÓN DE PLANTILLA ===")
    available_objects = discover_template_objects(template_path)
    
    if available_objects:
        print(f"\n📋 Objetos disponibles en la plantilla:")
        for name, obj_type in available_objects:
            print(f"  - {name} (Tipo: {obj_type})")
    else:
        print("\n❌ No se encontraron objetos en la plantilla")
    
    print("\n🖨️ === IMPRESIÓN DE ETIQUETA ===")
    # Imprimir con el código de barras
    success = print_label_with_bpac(template_path, "123456789")
    
    if success:
        print("\n🎉 ¡Impresión completada exitosamente!")
    else:
        print("\n💥 Hubo problemas durante la impresión")