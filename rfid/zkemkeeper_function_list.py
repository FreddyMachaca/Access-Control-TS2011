"""
Extractor de funciones del DLL zkemkeeper.dll a través de COM

Este script se conecta al componente COM de zkemkeeper y enumera
todas las funciones disponibles, junto con información sobre los
parámetros y valores de retorno si es posible obtenerlos.
"""

import os
import sys
import time
from datetime import datetime

try:
    import pythoncom
    import win32com.client
    import win32com.client.gencache
    from win32com.client import makepy
except ImportError:
    print("Este script requiere PyWin32. Instálelo con: pip install pywin32")
    sys.exit(1)

def create_zkemkeeper_types():
    """Intenta generar información de tipos para zkemkeeper"""
    try:
        # Forzar generación de archivos de caché de tipos
        win32com.client.gencache.EnsureModule('{00853A19-BD51-419B-9269-2DABE57EB61F}', 0, 1, 0)
        print("Generación de tipos para zkemkeeper completada")
        return True
    except:
        print("No se pudo generar información de tipos para zkemkeeper")
        return False

def extract_zkemkeeper_methods():
    """Extrae todos los métodos disponibles en el objeto COM zkemkeeper"""
    try:
        # Inicializar COM
        pythoncom.CoInitialize()
        
        # Crear una instancia del objeto COM
        zk = win32com.client.Dispatch("zkemkeeper.ZKEM.1")
        
        # Obtener todos los métodos
        method_names = [m for m in dir(zk) if not m.startswith('_')]
        
        # Organizar los métodos por categorías
        categories = {
            "Conexión": [],
            "Control de Acceso": [],
            "Tarjetas": [],
            "Usuarios": [],
            "Registros": [],
            "Sistema": [],
            "Otros": []
        }
        
        # Clasificar los métodos
        for method in method_names:
            if method.startswith(("Connect", "Disconnect", "SetCommPassword", "GetLastError")):
                categories["Conexión"].append(method)
            elif method.startswith(("Get", "Set")) and ("Door" in method or "Lock" in method or "Alarm" in method):
                categories["Control de Acceso"].append(method)
            elif "Card" in method:
                categories["Tarjetas"].append(method)
            elif "User" in method or "Enroll" in method or "Fingerprint" in method:
                categories["Usuarios"].append(method)
            elif "Log" in method or "Record" in method or "Attendance" in method:
                categories["Registros"].append(method)
            elif method.startswith(("Get", "Set")) and ("Time" in method or "Date" in method or "Device" in method):
                categories["Sistema"].append(method)
            else:
                categories["Otros"].append(method)
        
        return categories, method_names
        
    except Exception as e:
        print(f"Error al extraer métodos: {e}")
        return None, []
    finally:
        pythoncom.CoUninitialize()

def try_connect_device():
    """Intenta conectarse a un dispositivo ZKTeco"""
    try:
        pythoncom.CoInitialize()
        zk = win32com.client.Dispatch("zkemkeeper.ZKEM.1")
        
        print("\n==== PRUEBA DE CONEXIÓN ====")
        print("Esta parte intentará conectarse a un dispositivo ZKTeco para")
        print("verificar que la interfaz COM funciona correctamente.")
        
        # Solicitar información de conexión
        ip = input("Ingrese la dirección IP del dispositivo (default: 192.168.0.201): ") or "192.168.0.201"
        port_str = input("Ingrese el puerto del dispositivo (default: 14370): ") or "14370"
        port = int(port_str)
        
        print(f"\nIntentando conectar a {ip}:{port}...")
        
        # Intentar la conexión
        if zk.Connect_Net(ip, port):
            print("✓ Conectado exitosamente al dispositivo")
            device_id = 1
            
            print("\n==== INFORMACIÓN DEL DISPOSITIVO ====")
            # Firmware
            try:
                success, firmware = zk.GetFirmwareVersion(device_id)
                if success:
                    print(f"Firmware: {firmware}")
            except:
                print("No se pudo obtener el firmware")
            
            # Serial
            try:
                success, serial = zk.GetSerialNumber(device_id)
                if success:
                    print(f"Número de Serie: {serial}")
            except:
                print("No se pudo obtener el número de serie")
            
            # Product code
            try:
                success, model = zk.GetProductCode(device_id)
                if success:
                    print(f"Modelo: {model}")
            except:
                print("No se pudo obtener el modelo")
            
            # Vendor
            try:
                success, vendor = zk.GetVendor()
                if success:
                    print(f"Fabricante: {vendor}")
            except:
                print("No se pudo obtener el fabricante")
            
            # Desconectar
            zk.Disconnect()
            print("\n✓ Dispositivo desconectado")
            return True
        else:
            print("✗ No se pudo conectar al dispositivo")
            error = zk.GetLastError()
            print(f"  Código de error: {error}")
            return False
    except Exception as e:
        print(f"Error durante la prueba de conexión: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

def save_methods_to_file():
    """Guarda la lista de métodos en un archivo"""
    categories, all_methods = extract_zkemkeeper_methods()
    
    if not categories:
        print("No se pudo obtener la lista de métodos")
        return
    
    result_file = os.path.join("C:/Users/PC-HP/Desktop/rfid", "zkemkeeper_methods.txt")
    
    with open(result_file, 'w', encoding='utf-8') as f:
        f.write("MÉTODOS DISPONIBLES EN ZKEMKEEPER.DLL\n")
        f.write("=" * 40 + "\n")
        f.write(f"Fecha de generación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        f.write(f"Total de métodos encontrados: {len(all_methods)}\n\n")
        
        f.write("MÉTODOS POR CATEGORÍA:\n")
        f.write("=" * 40 + "\n\n")
        
        for category, methods in categories.items():
            f.write(f"\n{category.upper()} ({len(methods)} métodos):\n")
            f.write("-" * 40 + "\n")
            for method in sorted(methods):
                f.write(f"  - {method}\n")
        
        f.write("\n\nLISTA COMPLETA DE MÉTODOS:\n")
        f.write("=" * 40 + "\n\n")
        
        for i, method in enumerate(sorted(all_methods), 1):
            f.write(f"{i:3}. {method}\n")
    
    print(f"\nLista de métodos guardada en {result_file}")

def main():
    print("EXTRACTOR DE FUNCIONES DE ZKEMKEEPER.DLL")
    print("=" * 40)
    print("Este script extrae los métodos disponibles en zkemkeeper.dll")
    print("mediante la interfaz COM, que es la forma recomendada de")
    print("interactuar con este componente.")
    
    print("\nOpciones:")
    print("1. Mostrar métodos por categoría")
    print("2. Guardar métodos en un archivo")
    print("3. Probar conexión a un dispositivo ZKTeco")
    print("4. Salir")
    
    choice = input("\nSeleccione una opción (1-4): ")
    
    if choice == "1":
        # Crear información de tipos (opcional)
        create_zkemkeeper_types()
        
        # Extraer y mostrar métodos
        categories, all_methods = extract_zkemkeeper_methods()
        if categories:
            print(f"\nSe encontraron {len(all_methods)} métodos en total\n")
            
            for category, methods in categories.items():
                print(f"\n{category.upper()} ({len(methods)} métodos):")
                print("-" * 40)
                for method in sorted(methods):
                    print(f"  - {method}")
    
    elif choice == "2":
        create_zkemkeeper_types()
        save_methods_to_file()
    
    elif choice == "3":
        try_connect_device()
    
    elif choice == "4":
        print("Saliendo...")
    
    else:
        print("Opción inválida")

if __name__ == "__main__":
    main()
