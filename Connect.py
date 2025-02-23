import ctypes
from ctypes import create_string_buffer

plcommpro = None

def conectar_molinete(ip, puerto):
    global plcommpro
    try:
        plcommpro = ctypes.windll.LoadLibrary("plcommpro.dll")
        
        params = f"protocol=TCP,ipaddress={ip},port={puerto},timeout=4000,passwd="
        
        constr = create_string_buffer(params.encode('utf-8'))
        
        hcommpro = plcommpro.Connect(constr)
        
        if hcommpro != 0:
            print(f"Conectado al molinete {ip}:{puerto}")
            return hcommpro
        else:
            print("Error al conectar al molinete.")
            return None

    except Exception as e:
        print(f"Error durante la conexión: {e}")
        return None

if __name__ == "__main__":
    ip_molinete = "192.168.0.201"
    puerto_molinete = 14370

    hcommpro = conectar_molinete(ip_molinete, puerto_molinete)

    if hcommpro:
        print("Conexión exitosa.")
    else:
        print("Fallo al conectar.")
