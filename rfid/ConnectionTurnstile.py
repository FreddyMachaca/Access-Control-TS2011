"""
Módulo optimizado para la conexión y lectura de tarjetas RFID de un molinete ZKTeco C2-260
basado en las funciones disponibles detectadas en el dispositivo específico.
"""
import time
import sys
from datetime import datetime

try:
    import pythoncom
    import win32com.client
except ImportError:
    print("Este script requiere PyWin32. Instálelo con: pip install pywin32")
    print("Ejecutando: pip install pywin32")
    import subprocess
    subprocess.call([sys.executable, "-m", "pip", "install", "pywin32"])
    print("Por favor, reinicie el script después de la instalación.")
    sys.exit(1)

class ConnectionTurnstile:
    def __init__(self):
        """Inicializa la conexión con el molinete ZKTeco"""
        self.device_id = 1  # ID de dispositivo por defecto
        self.zkem = None
        self.connected = False
        self.device_info = {}
        
        # Inicializar COM
        try:
            pythoncom.CoInitialize()
            self.zkem = win32com.client.Dispatch("zkemkeeper.ZKEM.1")
            print("Interfaz COM zkemkeeper inicializada correctamente")
        except Exception as e:
            print(f"Error al inicializar zkemkeeper: {e}")
            sys.exit(1)
    
    def connect(self, ip="192.168.0.201", port=14370):
        """Conectar al dispositivo ZKTeco usando los parámetros proporcionados"""
        if not self.zkem:
            print("La interfaz COM no está inicializada.")
            return False
        
        try:
            print(f"Conectando a {ip}:{port}...")
            if self.zkem.Connect_Net(ip, port):
                self.connected = True
                print("✓ Conexión establecida exitosamente")
                
                # Obtener información básica del dispositivo
                self._get_device_info()
                
                # Habilitar el dispositivo para recibir eventos
                if self.zkem.EnableDevice(self.device_id, True):
                    print("✓ Dispositivo habilitado")
                else:
                    print("⚠ No se pudo habilitar el dispositivo")
                
                return True
            else:
                print("✗ Error al conectar al dispositivo")
                error_code = self.zkem.GetLastError()
                print(f"  Código de error: {error_code}")
                return False
        except Exception as e:
            print(f"Error durante la conexión: {e}")
            return False
    
    def disconnect(self):
        """Desconectar del dispositivo"""
        if self.connected and self.zkem:
            try:
                self.zkem.Disconnect()
                print("Dispositivo desconectado")
                self.connected = False
                return True
            except Exception as e:
                print(f"Error al desconectar: {e}")
        return False
    
    def _get_device_info(self):
        """Obtener y almacenar información del dispositivo"""
        if not self.connected:
            return
        
        try:
            # Usar las funciones que sabemos que funcionan según las pruebas
            self.device_info["firmware"] = self.zkem.GetFirmwareVersion(self.device_id)[1]
            self.device_info["mac"] = self.zkem.GetDeviceMAC(self.device_id)[1]
            self.device_info["serial"] = self.zkem.GetSerialNumber(self.device_id)[1]
            self.device_info["ip"] = self.zkem.GetDeviceIP(self.device_id)[1]
            self.device_info["product_code"] = self.zkem.GetProductCode(self.device_id)[1]
            self.device_info["vendor"] = self.zkem.GetVendor()[1]
            self.device_info["platform"] = self.zkem.GetPlatform(self.device_id)[1]
            
            print("\n=== Información del Dispositivo ===")
            print(f"Modelo: {self.device_info.get('product_code', 'Desconocido')}")
            print(f"Firmware: {self.device_info.get('firmware', 'Desconocido')}")
            print(f"Número de Serie: {self.device_info.get('serial', 'Desconocido')}")
            print(f"Dirección MAC: {self.device_info.get('mac', 'Desconocido')}")
            print(f"Dirección IP: {self.device_info.get('ip', 'Desconocido')}")
            print(f"Fabricante: {self.device_info.get('vendor', 'Desconocido')}")
            print(f"Plataforma: {self.device_info.get('platform', 'Desconocido')}")
            print("=" * 40)
            
        except Exception as e:
            print(f"Error al obtener información del dispositivo: {e}")
    
    def read_cards(self, duration=None):
        """
        Lee tarjetas RFID durante un tiempo específico o indefinidamente.
        
        Args:
            duration: Tiempo en segundos para leer tarjetas. None para leer indefinidamente.
        """
        if not self.connected:
            print("No hay conexión activa con el dispositivo.")
            return
        
        try:
            # Registrar para eventos
            if not self.zkem.RegEvent(self.device_id, 65535):  # 65535 para todos los eventos
                print("⚠ No se pudo registrar para eventos")
                return
            
            print("\nIniciando lectura de tarjetas RFID...")
            print("Presente una tarjeta al lector. Presione Ctrl+C para detener.")
            
            # Variables para seguimiento
            start_time = time.time()
            last_card = None
            card_timeout = 2  # segundos para considerar lecturas duplicadas
            last_card_time = 0
            
            try:
                while True:
                    # Verificar si hemos alcanzado la duración especificada
                    if duration and time.time() - start_time > duration:
                        print(f"\nTiempo de lectura ({duration}s) completado.")
                        break
                    
                    # Procesamiento de eventos manual
                    # No usamos GetLastEvent() ni PollCard() porque no funcionan según las pruebas
                    
                    # Intentamos usar GetStrCardNumber directamente
                    result, card_number = self.zkem.GetStrCardNumber()
                    
                    if result and card_number and card_number != "0":
                        # Evitar lecturas duplicadas de la misma tarjeta
                        if card_number != last_card or time.time() - last_card_time > card_timeout:
                            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            print(f"\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                            print(f"📍 Tarjeta RFID detectada")
                            print(f"📌 Número: {card_number}")
                            print(f"🕒 Hora: {timestamp}")
                            print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                            
                            last_card = card_number
                            last_card_time = time.time()
                    
                    # Alternativamente, intentamos usar GetHIDEventCardNumAsStr
                    result, hid_card = self.zkem.GetHIDEventCardNumAsStr()
                    
                    if result and hid_card:
                        if hid_card != last_card or time.time() - last_card_time > card_timeout:
                            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            print(f"\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                            print(f"📍 Tarjeta HID detectada")
                            print(f"📌 Número: {hid_card}")
                            print(f"🕒 Hora: {timestamp}")
                            print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                            
                            last_card = hid_card
                            last_card_time = time.time()
                    
                    # Pausa corta para no consumir CPU
                    time.sleep(0.1)
                    
            except KeyboardInterrupt:
                print("\nLectura de tarjetas detenida por el usuario")
            
        except Exception as e:
            print(f"Error durante la lectura de tarjetas: {e}")
        finally:
            # Intentar cancelar el registro de eventos
            try:
                self.zkem.RegEvent(self.device_id, 0)  # 0 para cancelar todos los eventos
            except:
                pass
    
    def get_users(self):
        """Intenta obtener la lista de usuarios/tarjetas registradas en el dispositivo"""
        if not self.connected:
            print("No hay conexión activa con el dispositivo.")
            return
        
        print("\nObteniendo usuarios registrados...")
        
        try:
            # Limpiar buffer
            self.zkem.EnableDevice(self.device_id, False)
            time.sleep(0.5)
            
            # Obtener todos los usuarios
            if not self.zkem.ReadAllUserID(self.device_id):
                print("⚠ No se pudo leer la información de usuarios")
                self.zkem.EnableDevice(self.device_id, True)
                return
            
            print("\n=== Usuarios/Tarjetas registradas ===")
            count = 0
            
            # Utilizar método alternativo para leer usuarios
            try:
                while True:
                    result, user_id = self.zkem.SSR_GetAllUserInfo(self.device_id)
                    if not result:
                        break
                        
                    # Intentar obtener el número de tarjeta
                    card_result, card_number = self.zkem.GetStrCardNumber()
                    
                    if card_result and card_number and card_number != "0":
                        count += 1
                        print(f"Usuario #{count}: ID={user_id}, Tarjeta={card_number}")
            except:
                if count == 0:
                    print("No se encontraron usuarios con el método principal")
            
            # Si no se encontraron usuarios, intentar con un método alternativo
            if count == 0:
                try:
                    # Iterar por posibles IDs de usuario
                    for i in range(1, 100):
                        user_id = str(i)
                        # Intentar obtener información del usuario
                        if self.zkem.SSR_GetUserInfo(self.device_id, user_id):
                            # Intentar obtener número de tarjeta
                            card_result, card_number = self.zkem.GetStrCardNumber()
                            if card_result and card_number and card_number != "0":
                                count += 1
                                print(f"Usuario #{count}: ID={user_id}, Tarjeta={card_number}")
                except:
                    pass
            
            if count == 0:
                print("No se encontraron usuarios registrados con tarjetas")
            else:
                print(f"Total de usuarios/tarjetas encontrados: {count}")
            
        except Exception as e:
            print(f"Error al obtener usuarios: {e}")
        finally:
            self.zkem.EnableDevice(self.device_id, True)

def main():
    """Función principal para probar la conexión al molinete"""
    print("=" * 60)
    print("     SISTEMA DE CONEXIÓN A MOLINETE ZKTECO C2-260")
    print("=" * 60)
    
    conn = ConnectionTurnstile()
    
    # Solicitar datos de conexión
    ip = input("Ingrese la dirección IP del dispositivo (default: 192.168.0.201): ") or "192.168.0.201"
    port_str = input("Ingrese el puerto del dispositivo (default: 14370): ") or "14370"
    port = int(port_str)
    
    if conn.connect(ip, port):
        try:
            while True:
                print("\n" + "=" * 50)
                print("         MENÚ DE OPERACIONES")
                print("=" * 50)
                print("1. Leer tarjetas RFID")
                print("2. Ver información del dispositivo")
                print("3. Buscar usuarios/tarjetas registradas")
                print("4. Salir")
                print("=" * 50)
                
                option = input("Seleccione una opción: ")
                
                if option == "1":
                    # Preguntar por la duración de la lectura
                    dur_str = input("Ingrese duración de la lectura en segundos (Enter para indefinido): ")
                    duration = int(dur_str) if dur_str else None
                    
                    conn.read_cards(duration)
                    input("\nPresione Enter para continuar...")
                
                elif option == "2":
                    conn._get_device_info()
                    input("\nPresione Enter para continuar...")
                
                elif option == "3":
                    conn.get_users()
                    input("\nPresione Enter para continuar...")
                
                elif option == "4":
                    print("Finalizando programa...")
                    break
                
                else:
                    print("Opción inválida. Por favor seleccione una opción válida.")
            
        except KeyboardInterrupt:
            print("\nPrograma interrumpido por el usuario")
        finally:
            conn.disconnect()
    
    print("Programa finalizado.")

if __name__ == "__main__":
    main()
