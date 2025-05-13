"""
Script para probar las funciones disponibles en la interfaz zkemkeeper COM.
Utiliza este script para identificar qué funciones son compatibles con tu
versión específica del SDK ZKTeco.
"""

import os
import sys
import time
from datetime import datetime

try:
    import pythoncom
    import win32com.client
except ImportError:
    print("Este script requiere PyWin32. Instálelo con: pip install pywin32")
    sys.exit(1)

def test_zkemkeeper_functions():
    """Prueba diversas funciones de zkemkeeper para determinar compatibilidad"""
    print("=" * 60)
    print("     PRUEBA DE FUNCIONES DISPONIBLES EN ZKEMKEEPER.DLL     ")
    print("=" * 60)
    
    try:
        # Inicialización de COM
        pythoncom.CoInitialize()
        zkem = win32com.client.Dispatch("zkemkeeper.ZKEM.1")
        
        print("Interfaz COM zkemkeeper inicializada correctamente")
        
        # Prueba de conexión
        ip = input("Ingrese la dirección IP del dispositivo (default: 192.168.0.201): ") or "192.168.0.201"
        port = input("Ingrese el puerto del dispositivo (default: 14370): ") or "14370"
        
        print(f"Intentando conectar a {ip}:{port}...")
        
        # Intentar conectar
        if zkem.Connect_Net(ip, int(port)):
            print("✓ Conectado correctamente al dispositivo")
            machine_id = 1  # ID de dispositivo por defecto
            
            # Tabla de funciones a probar
            functions_to_test = [
                # Funciones básicas
                ("GetFirmwareVersion", lambda: zkem.GetFirmwareVersion(machine_id)),
                ("GetDeviceMAC", lambda: zkem.GetDeviceMAC(machine_id)),
                ("GetSerialNumber", lambda: zkem.GetSerialNumber(machine_id)),
                ("GetDeviceIP", lambda: zkem.GetDeviceIP(machine_id)),
                ("GetProductCode", lambda: zkem.GetProductCode(machine_id)),
                ("GetVendor", lambda: zkem.GetVendor()),
                ("GetPlatform", lambda: zkem.GetPlatform(machine_id)),
                ("GetCardFun", lambda: zkem.GetCardFun(machine_id)),
                
                # Funciones para control de acceso
                ("EnableDevice", lambda: zkem.EnableDevice(machine_id, True)),
                ("DisableDevice", lambda: zkem.DisableDevice(machine_id)),
                ("GetLastError", lambda: zkem.GetLastError()),
                
                # Funciones para eventos en tiempo real
                ("RegEvent", lambda: zkem.RegEvent(machine_id, 65535)),  # 65535 = todos los eventos
                ("GetRTLog", lambda: zkem.GetRTLog(machine_id)),
                
                # Funciones específicas para tarjetas
                ("GetStrCardNumber", lambda: zkem.GetStrCardNumber()),
                ("ReadCard", lambda: zkem.ReadCard(machine_id, "")),
                ("GetHIDEventCardNumAsStr", lambda: zkem.GetHIDEventCardNumAsStr()),
                ("PollCard", lambda: zkem.PollCard()),
                
                # Función para leer eventos de log
                ("ReadGeneralLogData", lambda: zkem.ReadGeneralLogData(machine_id)),
            ]
            
            # Probar cada función e informar resultado
            results = []
            print("\n--- Probando funciones del SDK ---")
            for func_name, func_call in functions_to_test:
                try:
                    result = func_call()
                    print(f"✓ {func_name}: Disponible (Resultado: {result})")
                    results.append((func_name, True, str(result)))
                except Exception as e:
                    print(f"✗ {func_name}: No disponible o error ({str(e)})")
                    results.append((func_name, False, str(e)))
            
            # Test específico para lectura de tarjetas
            print("\n--- Prueba especial: Lectura de tarjeta en tiempo real ---")
            print("Presente una tarjeta al lector dentro de los próximos 10 segundos...")
            
            # Habilitar eventos en tiempo real si está disponible
            try:
                zkem.RegEvent(machine_id, 65535)
                card_detected = False
                start_time = time.time()
                
                while time.time() - start_time < 10 and not card_detected:
                    # Intentar varios métodos para detectar tarjetas
                    try:
                        # Método 1: PollCard
                        if hasattr(zkem, 'PollCard') and zkem.PollCard():
                            card_num = zkem.GetStrCardNumber()
                            print(f"✓ Tarjeta detectada (PollCard): {card_num}")
                            card_detected = True
                            break
                    except:
                        pass
                    
                    # Método 2: GetLastEvent
                    try:
                        if hasattr(zkem, 'GetLastEvent') and zkem.GetLastEvent():
                            card_num = zkem.GetStrCardNumber()
                            if card_num:
                                print(f"✓ Tarjeta detectada (GetLastEvent): {card_num}")
                                card_detected = True
                                break
                    except:
                        pass
                    
                    # Pausa breve
                    time.sleep(0.2)
                
                if not card_detected:
                    print("✗ No se detectó ninguna tarjeta en el tiempo asignado")
            except Exception as e:
                print(f"✗ Error al intentar leer tarjetas: {e}")
            
            # Resumen de disponibilidad
            print("\n--- Resumen de Funciones ---")
            available = sum(1 for _, status, _ in results if status)
            print(f"Funciones disponibles: {available} de {len(functions_to_test)}")
            
            # Desconectar
            zkem.Disconnect()
            print("\nDispositivo desconectado")
            
        else:
            print("✗ No se pudo conectar al dispositivo")
            try:
                error_code = zkem.GetLastError()
                print(f"Código de error: {error_code}")
            except:
                pass
        
        print("\n--- Finalizado ---")
    
    except Exception as e:
        print(f"Error general: {e}")
    finally:
        # Liberar recursos COM
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    # Verificar si zkemkeeper.dll existe
    if os.path.exists("zkemkeeper.dll"):
        print("✓ zkemkeeper.dll encontrado")
    else:
        print("✗ zkemkeeper.dll no encontrado - puede que el script no funcione correctamente")
    
    test_zkemkeeper_functions()
