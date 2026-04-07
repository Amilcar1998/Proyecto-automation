"""
Sistema de Automatización AS400
Punto de entrada principal de la aplicación
"""

import os
import sys
import time
from flujo_completo import flujo_completo
from validacion_db.transferencias_wms_infor import *
from Jira_Utilidades.jira_utilidades import JiraClient
from validacion_db.Validacion_PO import *
from validacion_db.PO_Automática import *
from validacion_db.Validacion_PO import procesar_orden_krws
from conexion_config.logging_config import get_logger


logger = get_logger(__file__)


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "AUTO":
        procesar_transferencias_wms_infor()
        main_PO()
        return
        
    """Menú principal de la aplicación"""
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("\n" + "="*60)
        print("🚀 SISTEMA DE AUTOMATIZACIÓN AS400")
        print("="*60)
        print("1. Ejecutar flujo completo de garantías")
        print("2. Procesar transferencias WMS-INFOR")
        print("3. Ejecutar comandos CL en AS400")
        print("4. Jira independiente")
        print("5. validacion PO")
        print("6. PO AUTOMATICA VENDOR")
        print("0. Salir")
        print("-"*60)
        
        opcion = input("Selecciona una opción (1-6): ").strip()
        
        if opcion == "1":
            print("\nIniciando flujo completo...")
            flujo_completo()
            input("\nPresiona Enter para volver al menú...")
        elif opcion == "2":
            print("\nIniciando procesamiento de transferencias WMS-INFOR...")
            procesar_transferencias_wms_infor()
            input("\nPresiona Enter para volver al menú...")
        elif opcion == "3":
            ejecutar_comando_manual()
            input("\nPresiona Enter para volver al menú...")
        elif opcion == "4":
            load = "ASN_04_165569_1"
            ruta = r"C:\Users\eliseo_lopezp\proyecto1\Reportes\Transferencias_WMS_INFOR_20260116\Reporte_Orden_ASN_04_165569_1_20260116.docx"

            J = JiraClient()
            #J.main_jira(load, ruta)


        elif opcion =="5":
            main_PO()
            input("\nPresiona Enter para volver al menú...")

        elif opcion == "6":
            main_vendor()
            input("\nPresiona Enter para volver al menú...")
        elif opcion == "0":
            print(" ¡Hasta luego!")
            break
        else:
            print(" Opción inválida. Intenta de nuevo.")
            time.sleep(1)

def ejecutar_comando_manual():
    """Permite al usuario ejecutar comandos CL manualmente"""
    from scripts_as400.ejecutor_cl import EjecutorCL
    
    print("\n" + "="*60)
    print("EJECUTAR COMANDOS CL EN AS400")
    print("="*60)
    print("Ejemplos de comandos:")
    print("- CALL PGM(RIUNICOM63/SIWINTRACL) PARM('11')")
    print("- DSPLIBL")
    print("- WRKACTJOB")
    print("\nEscribe 'salir' para volver al menú principal")
    print("-"*60)
    
    cl = EjecutorCL()
    
    while True:
        comando = input("\nComando CL > ").strip()
        
        if comando.lower() == 'salir':
            break
        
        if not comando:
            continue
            
        print("\n1. Ejecutar directamente")
        print("2. Enviar como trabajo (SBMJOB)")
        opcion = input("Selecciona una opción (1-2): ").strip()
        
        if opcion == "1":
            resultado = cl.ejecutar(comando)
            if resultado:
                print("✅ Comando ejecutado correctamente")
            else:
                print("❌ Error al ejecutar el comando")
        elif opcion == "2":
            job_name = input("Nombre del trabajo (opcional): ").strip()
            job_name = job_name if job_name else None
            
            resultado = cl.submit_job(comando, job_name=job_name)
            if resultado:
                print(f"✅ Trabajo enviado: {resultado.get('job_name')}")
            else:
                print("❌ Error al enviar el trabajo")
        else:
            print("❌ Opción inválida")

if __name__ == "__main__":
    if not os.path.exists('logs'):
        os.makedirs('logs')
    main()

