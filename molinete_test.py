from ctypes import windll, byref, create_string_buffer, c_int, c_char_p, c_long, c_ulong, c_void_p, c_bool, POINTER
import os
import sys
import time

class ZKTecoDevice:
    def __init__(self):
        self.commpro = None
        self.hcommpro = 0
        self.connected = False
        self.machine_number = 1
        
        # Cargar la librería del SDK
        try:
            self.commpro = windll.LoadLibrary("plcommpro.dll")
            print("Librería SDK cargada correctamente")
        except Exception as e:
            print(f"Error al cargar la librería: {e}")
            sys.exit(1)
    
    def connect(self, ip_address="192.168.0.201", port=14370, timeout=4000, password=""):
        if self.connected:
            print("Ya está conectado al dispositivo")
            return True
            
        try:
            params = f"protocol=TCP,ipaddress={ip_address},port={port},timeout={timeout},passwd={password}"
            constr = create_string_buffer(params.encode())
            self.hcommpro = self.commpro.Connect(constr)
            
            if self.hcommpro != 0:
                self.connected = True
                print(f"Conectado al dispositivo: {ip_address}:{port}")
                return True
            else:
                error_code = self.commpro.PullLastError()
                print(f"Error de conexión. Código: {error_code}. Verifique la IP y que el dispositivo esté encendido.")
                return False
        except Exception as e:
            print(f"Error en la conexión: {e}")
            return False
    
    def disconnect(self):
        if not self.connected:
            print("No hay conexión activa")
            return
            
        try:
            self.commpro.Disconnect(self.hcommpro)
            self.connected = False
            self.hcommpro = 0
            print("Desconectado del dispositivo")
        except Exception as e:
            print(f"Error al desconectar: {e}")
    
    def get_device_info(self):
        if not self.connected:
            print("No hay conexión activa")
            return
        
        try:
            # Obtener parámetros del dispositivo usando GetDeviceParam
            buffer = create_string_buffer(2048)
            items = "DeviceID,Door1SensorType,Door1Drivertime,Door1Intertime,~ZKFPVersion"
            p_items = create_string_buffer(items.encode())
            ret = self.commpro.GetDeviceParam(self.hcommpro, buffer, 2048, p_items)
            
            if ret >= 0:
                print(f"Información del dispositivo: {buffer.value.decode()}")
            else:
                error_code = self.commpro.PullLastError()
                print(f"Error al obtener información del dispositivo. Código: {error_code}")
            
            # Verificar eventos recientes
            rt_log = create_string_buffer(256)
            ret = self.commpro.GetRTLog(self.hcommpro, rt_log, 256)
            
            if ret >= 0:
                if ret == 0:
                    print("No hay eventos recientes en el dispositivo")
                else:
                    print(f"Eventos recientes detectados: {rt_log.value.decode()}")
            else:
                error_code = self.commpro.PullLastError()
                print(f"Error al verificar eventos del dispositivo. Código: {error_code}")
                
        except Exception as e:
            print(f"Error al obtener información del dispositivo: {e}")

    def read_card(self):
        if not self.connected:
            print("No hay conexión activa")
            return None
        
        try:
            print("Acerque la tarjeta RFID al lector...")
            print("Esperando lectura de tarjeta RFID (10 segundos máximo)...")
            
            # Limpiar eventos previos
            self._clear_previous_events()
            
            # Esperar a que se detecte una tarjeta RFID
            timeout = time.time() + 10  # 10 segundos de timeout
            
            while time.time() < timeout:
                try:
                    # Crear buffer para el número de tarjeta HID
                    card_buffer = create_string_buffer(64)
                    
                    # Obtener el número de tarjeta HID del evento más reciente
                    ret = self.commpro.GetHIDEventCardNumAsStr(byref(card_buffer))
                    
                    if ret:  # True indica éxito
                        card_number = card_buffer.value.decode('utf-8', errors='ignore').strip()
                        if card_number and card_number != "0" and len(card_number) > 0:
                            print(f"Tarjeta RFID detectada: {card_number}")
                            return card_number
                    
                    # También verificar eventos en tiempo real como respaldo
                    rt_log = create_string_buffer(256)
                    ret_rt = self.commpro.GetRTLog(self.hcommpro, rt_log, 256)
                    
                    if ret_rt > 0:
                        event_data = rt_log.value.decode('utf-8', errors='ignore').strip()
                        if event_data:
                            print(f"Evento detectado: {event_data}")
                            card_number = self._parse_card_event(event_data)
                            if card_number:
                                print(f"Tarjeta RFID detectada (evento): {card_number}")
                                return card_number
                    
                except Exception as e:
                    # Error en esta iteración, pero continuar intentando
                    pass
                
                time.sleep(0.2)  # Pausa entre lecturas
            
            print("Tiempo de espera agotado (10 segundos). No se detectó ninguna tarjeta RFID.")
            return None
            
        except Exception as e:
            print(f"Error al leer la tarjeta RFID: {e}")
            return None

    def _clear_previous_events(self):
        """Limpia eventos previos del buffer"""
        try:
            # Leer y descartar eventos previos que puedan estar en el buffer
            count = 0
            while count < 10:  # Máximo 10 intentos para evitar bucle infinito
                rt_log = create_string_buffer(256)
                ret = self.commpro.GetRTLog(self.hcommpro, rt_log, 256)
                if ret <= 0:
                    break
                count += 1
            if count > 0:
                print(f"Se limpiaron {count} eventos previos del buffer")
        except Exception as e:
            print(f"Error al limpiar eventos previos: {e}")

    def _parse_card_event(self, event_data):
        """Extrae el número de tarjeta del evento"""
        try:
            if not event_data:
                return None
            
            # Buscar CardNo en el evento
            if "CardNo=" in event_data:
                # Extraer el número después de CardNo=
                parts = event_data.split("CardNo=")
                if len(parts) > 1:
                    card_part = parts[1].split("\t")[0].split(",")[0].strip()
                    if card_part and card_part.isdigit():
                        return card_part
            
            # Intentar otros formatos posibles
            card_indicators = ["Cardno=", "Card=", "CardNumber="]
            for indicator in card_indicators:
                if indicator in event_data:
                    parts = event_data.split(indicator)
                    if len(parts) > 1:
                        card_part = parts[1].split("\t")[0].split(",")[0].strip()
                        if card_part and card_part.replace("-", "").isdigit():
                            return card_part.replace("-", "")
            
            return None
            
        except Exception as e:
            print(f"Error al parsear evento de tarjeta: {e}")
            return None

    def _print_error_description(self, error_code):
        """Imprime la descripción del código de error"""
        error_descriptions = {
            0: "Sin errores - No hay eventos nuevos disponibles (esto es normal)",
            -1: "El comando no se envió correctamente",
            -2: "El comando no tiene respuesta",
            -3: "El buffer no es suficiente",
            -4: "La descompresión falló",
            -5: "La longitud de los datos leídos no es correcta",
            -6: "La longitud de los datos descomprimidos no coincide con la esperada",
            -7: "El comando se repite",
            -8: "La conexión no está autorizada",
            -9: "Error de datos: El resultado CRC es fallido",
            -10: "Error de datos: PullSDK no puede resolver los datos",
            -11: "Error de parámetro de datos",
            -12: "El comando no se ejecutó correctamente",
            -13: "Error de comando: Este comando no está disponible",
            -14: "La contraseña de comunicación no es correcta",
            -15: "Error al escribir el archivo",
            -16: "Error al leer el archivo",
            -17: "El archivo no existe",
            -99: "Error desconocido",
            -100: "La estructura de la tabla no existe",
            -101: "En la estructura de la tabla, el campo Condition no existe",
            -102: "El número total de campos no es consistente",
            -103: "La secuencia de campos no es consistente",
            -104: "Error de datos de eventos en tiempo real",
            -105: "Ocurren errores de datos durante la resolución de datos",
            -106: "Desbordamiento de datos: Los datos entregados son más de 4 MB de longitud",
            -107: "Error al obtener la estructura de la tabla",
            -108: "Opciones inválidas"
        }
        
        description = error_descriptions.get(error_code, f"Código de error desconocido: {error_code}")
        print(f"Descripción del error: {description}")

    def control_device(self, operation_id=1, door_id=1, index=1, state=3):
        """Controla las acciones del dispositivo (abrir puerta, cancelar alarma, etc.)"""
        if not self.connected:
            print("No hay conexión activa")
            return False
        
        try:
            print(f"Controlando dispositivo - Operación: {operation_id}, Puerta: {door_id}, Estado: {state}")
            
            options = create_string_buffer(b"")
            ret = self.commpro.ControlDevice(
                self.hcommpro,
                operation_id,  # Operation ID
                door_id,       # Door number
                index,         # Address type (1=door output, 2=auxiliary output)
                state,         # Opening time in seconds (5 seconds)
                0,             # Reserved parameter
                options        # Options (empty)
            )
            
            if ret >= 0:
                if operation_id == 1:
                    print(f"Puerta {door_id} activada por {state} segundos exitosamente")
                elif operation_id == 2:
                    print("Alarma cancelada exitosamente")
                elif operation_id == 3:
                    print("Dispositivo reiniciado exitosamente")
                elif operation_id == 4:
                    status = "habilitado" if state == 1 else "deshabilitado"
                    print(f"Estado normal abierto {status} exitosamente")
                return True
            else:
                error_code = self.commpro.PullLastError()
                print(f"Error al controlar dispositivo. Código: {error_code}")
                self._print_error_description(error_code)
                return False
                
        except Exception as e:
            print(f"Error al controlar dispositivo: {e}")
            return False

    def test_device_communication(self):
        """Prueba la comunicación con el dispositivo"""
        if not self.connected:
            print("No hay conexión activa")
            return False
        
        try:
            print("Probando comunicación con el dispositivo...")
            
            # Probar obtener información básica del dispositivo
            buffer = create_string_buffer(1024)
            items = "DeviceID,~DeviceName,~SerialNumber,Door1SensorType"
            p_items = create_string_buffer(items.encode())
            ret = self.commpro.GetDeviceParam(self.hcommpro, buffer, 1024, p_items)
            
            if ret >= 0:
                device_info = buffer.value.decode('utf-8', errors='ignore')
                print(f"Comunicación exitosa. Info del dispositivo: {device_info}")
                return True
            else:
                error_code = self.commpro.PullLastError()
                print(f"Error en comunicación. Código: {error_code}")
                self._print_error_description(error_code)
                return False
                
        except Exception as e:
            print(f"Error al probar comunicación: {e}")
            return False

def menu():
    device = ZKTecoDevice()
    
    while True:
        print("\n=== CONTROL DE MOLINETE TS 2011 PRO ===")
        print("1. Conectar al dispositivo")
        print("2. Obtener información del dispositivo")
        print("3. Leer tarjeta RFID")
        print("4. Activar molinete (5 segundos)")
        print("5. Probar comunicación")
        print("6. Desconectar")
        print("0. Salir")
        
        opcion = input("\nSeleccione una opción: ")
        
        if opcion == "1":
            ip = input("IP del dispositivo [192.168.0.201]: ") or "192.168.0.201"
            device.connect(ip_address=ip)
        elif opcion == "2":
            device.get_device_info()
        elif opcion == "3":
            device.read_card()
        elif opcion == "4":
            device.control_device(operation_id=1, door_id=1, index=1, state=5)
        elif opcion == "5":
            device.test_device_communication()
        elif opcion == "6":
            device.disconnect()
        elif opcion == "0":
            if device.connected:
                device.disconnect()
            print("Programa finalizado")
            break
        else:
            print("Opción inválida. Intente nuevamente.")

if __name__ == "__main__":
    menu()
