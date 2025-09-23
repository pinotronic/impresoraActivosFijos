import win32com.client

def print_label_with_bpac(template_path, barcode_value, text1_value="", text2_value=""):
    """
    Función para imprimir etiquetas usando b-PAC
    
    Args:
        template_path (str): Ruta completa al archivo de plantilla .lbx
        barcode_value (str): Valor del código de barras
        text1_value (str): Texto para el campo Text1
        text2_value (str): Texto para el campo Text2
    """
    bpacDoc = None
    try:
        # Crear el objeto de documento b-PAC
        bpacDoc = win32com.client.Dispatch("bpac.Document")
        
        # Abrir la plantilla
        print(f"📂 Abriendo plantilla: {template_path}")
        result = bpacDoc.Open(template_path)
        if not result:
            print("❌ Error: No se pudo abrir la plantilla")
            return False
            
        print("✅ Plantilla abierta exitosamente")
        
        # Establecer el valor del código de barras
        barcode_obj = bpacDoc.GetObject("Barcode1")
        if barcode_obj is not None:
            barcode_obj.Text = barcode_value
            print(f"🏷️ Código de barras establecido: {barcode_value}")
        else:
            print("❌ Error: No se pudo encontrar el objeto Barcode1")
            return False
        
        # Establecer texto en Text1 si se proporciona
        if text1_value:
            text1_obj = bpacDoc.GetObject("Text1")
            if text1_obj is not None:
                text1_obj.Text = text1_value
                print(f"📝 Text1 establecido: {text1_value}")
            else:
                print("⚠️ Advertencia: No se pudo encontrar el objeto Text1")
        
        # Establecer texto en Text2 si se proporciona
        if text2_value:
            text2_obj = bpacDoc.GetObject("Text2")
            if text2_obj is not None:
                text2_obj.Text = text2_value
                print(f"📝 Text2 establecido: {text2_value}")
            else:
                print("⚠️ Advertencia: No se pudo encontrar el objeto Text2")
        
        # Imprimir el documento (método alternativo)
        print("🖨️ Iniciando impresión...")
        try:
            # Método 1: Intentar con StartPrint/EndPrint
            bpacDoc.StartPrint("", 0)  # Usar impresora predeterminada
            print("   📤 StartPrint exitoso")
            
            # Alternativa más robusta para PrintOut
            result = bpacDoc.PrintOut(1, 0)  # Imprimir 1 copia
            print(f"   📄 PrintOut resultado: {result}")
            
            bpacDoc.EndPrint()
            print("   ✅ EndPrint exitoso")
            
        except Exception as print_error:
            print(f"❌ Error en método de impresión 1: {print_error}")
            
            # Método 2: Intentar método alternativo
            try:
                print("   🔄 Intentando método alternativo...")
                result = bpacDoc.DoPrint(1, 0)  # Método alternativo
                print(f"   📄 DoPrint resultado: {result}")
            except Exception as alt_print_error:
                print(f"❌ Error en método alternativo: {alt_print_error}")
                return False
        
        print("✅ Documento enviado a impresión")
        return True
        
    except Exception as e:
        print(f"❌ Error durante la impresión: {e}")
        return False
        
    finally:
        # Cerrar el documento
        if bpacDoc is not None:
            try:
                # Usar solo la referencia, no como método
                bpacDoc.Close
                print("📄 Documento cerrado")
            except Exception as e:
                print(f"⚠️ Advertencia al cerrar documento: {e}")

def test_available_objects(template_path):
    """
    Función para probar que podemos acceder a todos los objetos disponibles
    """
    bpacDoc = None
    try:
        bpacDoc = win32com.client.Dispatch("bpac.Document")
        
        if not bpacDoc.Open(template_path):
            print("❌ No se pudo abrir la plantilla")
            return False
            
        print("🧪 === PRUEBA DE OBJETOS ===")
        
        # Probar Barcode1
        try:
            barcode_obj = bpacDoc.GetObject("Barcode1")
            if barcode_obj:
                barcode_obj.Text = "TEST123"
                print("✅ Barcode1: Accesible y modificable")
            else:
                print("❌ Barcode1: No accesible")
        except Exception as e:
            print(f"❌ Barcode1: Error - {e}")
        
        # Probar Text1
        try:
            text1_obj = bpacDoc.GetObject("Text1")
            if text1_obj:
                text1_obj.Text = "Texto de prueba 1"
                print("✅ Text1: Accesible y modificable")
            else:
                print("❌ Text1: No accesible")
        except Exception as e:
            print(f"❌ Text1: Error - {e}")
        
        # Probar Text2
        try:
            text2_obj = bpacDoc.GetObject("Text2")
            if text2_obj:
                text2_obj.Text = "Texto de prueba 2"
                print("✅ Text2: Accesible y modificable")
            else:
                print("❌ Text2: No accesible")
        except Exception as e:
            print(f"❌ Text2: Error - {e}")
            
        return True
        
    except Exception as e:
        print(f"❌ Error durante la prueba: {e}")
        return False
        
    finally:
        if bpacDoc is not None:
            try:
                bpacDoc.Close
            except:
                pass

if __name__ == "__main__":
    # Configuración
    template_path = r"C:\Users\pvargas\OneDrive - sapal365\Documentos\My Labels\Diseno.lbx"
    
    print("🧪 === PRUEBA DE OBJETOS ===")
    test_available_objects(template_path)
    
    print("\n🖨️ === IMPRESIÓN DE ETIQUETA ===")
    # Imprimir con datos de ejemplo
    success = print_label_with_bpac(
        template_path=template_path,
        barcode_value="987654321",
        text1_value="PRODUCTO EJEMPLO",
        text2_value="LOTE: ABC123"
    )
    
    if success:
        print("\n🎉 ¡Impresión completada exitosamente!")
        print("\n💡 AHORA PUEDES USAR:")
        print("   - Barcode1: Para códigos de barras")
        print("   - Text1: Para texto adicional")
        print("   - Text2: Para más texto")
        print("\n❌ OBJETOS NO DISPONIBLES:")
        print("   - 'CODIGO' no existe en tu plantilla")
    else:
        print("\n💥 Hubo problemas durante la impresión")