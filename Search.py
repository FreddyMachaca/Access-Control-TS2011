import ctypes
from ctypes import create_string_buffer

def buscar_dispositivos():
    """
    Busca dispositivos ZKTeco en la red local usando SearchDevice.
    """
    try:
        # Cargar la biblioteca PullSDK
        plcommpro = ctypes.windll.LoadLibrary("C:\Sdk-Zkteco\Plcommpro.dll")

        # Configurar los parámetros para la búsqueda
        comm_type = "UDP"  # Usamos UDP para buscar dispositivos en la red
        address = "255.255.255.255"  # Dirección de broadcast
        buffer_size = 64 * 1024  # Tamaño del buffer (64 KB)
        dev_buf = create_string_buffer(buffer_size)  # Buffer para almacenar los resultados

        # Llamar a la función SearchDevice
        ret = plcommpro.SearchDevice(comm_type.encode('utf-8'), address.encode('utf-8'), dev_buf)

        if ret > 0:
            print(f"Se encontraron {ret} dispositivos:")
            # Decodificar el buffer y mostrar los resultados
            devices = dev_buf.value.decode('utf-8').strip().split("\r\n")
            for i, device in enumerate(devices, 1):
                print(f"Dispositivo {i}: {device}")
        else:
            print("No se encontraron dispositivos en la red.")
    
    except Exception as e:
        print(f"Error durante la búsqueda de dispositivos: {e}")

if __name__ == "__main__":
    buscar_dispositivos()