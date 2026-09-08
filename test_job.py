import sys
import os
sys.path.append(r"c:\Users\eliseo_lopezp\proyecto1")

from as400_core.ejecutor_cl import EjecutorCL
from conexion_config.conexion import ConexionAS400

def probar_espera_trabajo():
    cl = EjecutorCL()
    job_name = "WMSI_CTF14"
    print(f"Iniciando prueba de espera para el trabajo: {job_name}")
    try:
        # Llamar al método esperar_trabajo del EjecutorCL
        resultado = cl.esperar_trabajo(job_name)
        print(f"Resultado de la espera: {resultado}")
    except Exception as e:
        print(f"Error durante la espera: {e}")

if __name__ == '__main__':
    probar_espera_trabajo()
