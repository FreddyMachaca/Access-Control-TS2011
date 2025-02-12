import ctypes
from ctypes import create_string_buffer

plcommpro = ctypes.windll.LoadLibrary("C:\Sdk-Zkteco\Plcommpro.dll")

def conectar_molinete(ip, puerto=4370, timeout=5000):
    """
    Conecta el molinete usando el PullSDK.
    :param ip: Dirección IP del molinete.
    :param puerto: Puerto del molinete (por defecto 4370).
    :param timeout: Tiempo de espera en milisegundos.
    :return: Handle de conexión si es exitosa, None si falla.
    """
    try:
        #parámetros de conexión
        params = f"protocol=TCP,ipaddress={ip},port={puerto},timeout={timeout},passwd="
        constr = create_string_buffer(params.encode('utf-8'))
        
        handle = plcommpro.Connect(constr)
        if handle > 0:
            print("Conexión exitosa con el molinete.")
            return handle
        else:
            error_code = plcommpro.PullLastError()
            print(f"Fallo en la conexión. Código de error: {error_code}")
            return None
    except Exception as e:
        print(f"Error durante la conexión: {e}")
        return None


def leer_id_tarjeta(handle):
    """
    Lee el ID de una tarjeta RFID desde el molinete.
    :param handle: Handle de conexión devuelto por Connect.
    """
    try:
        buffer_size = 256
        rt_log = create_string_buffer(buffer_size)
        
        ret = plcommpro.GetRTLog(handle, rt_log, buffer_size)
        if ret >= 0:
            log_data = rt_log.value.decode('utf-8').strip()
            if log_data:
                fields = log_data.split(',')
                card_id = fields[3]
                print(f"ID de la tarjeta leído: {card_id}")
            else:
                print("No se recibieron datos del molinete.")
        else:
            print(f"Error al leer el registro en tiempo real. Código de error: {ret}")
    except Exception as e:
        print(f"Error al leer el ID de la tarjeta: {e}")


def desconectar_molinete(handle):
    """
    Desconecta el molinete.
    :param handle: Handle de conexión devuelto por Connect.
    """
    try:
        plcommpro.Disconnect(handle)
        print("Desconexión exitosa.")
    except Exception as e:
        print(f"Error durante la desconexión: {e}")


if __name__ == "__main__":
    ip_molinete = "192.168.12.154"
    puerto_molinete = 4370

    handle = conectar_molinete(ip_molinete, puerto_molinete)
    if handle:
        leer_id_tarjeta(handle)

        desconectar_molinete(handle)