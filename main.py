"""
Sistema de Automatización AS400
Punto de entrada principal de la aplicación
"""

import os
import sys
import time
import logging

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass
from flujo_completo import flujo_completo
from wms_infor.transferencias_wms_infor import procesar_transferencias_wms_infor
from jira.jira_utilidades import JiraClient
from ordenes_compra.Validacion_PO import main_PO, procesar_orden_krws
from ordenes_compra.PO_Automática import main_vendor
from conexion_config.logging_config import get_logger


# --- INICIO: MODIFICACIÓN PARA MOSTRAR LOGS EN CONSOLA ---
# Configuración centralizada del logger.
# Esto asegura que todos los logs de la aplicación se muestren en consola.
root_logger = logging.getLogger()
root_logger.setLevel(logging.INFO) # Nivel mínimo para todos los handlers

# Limpiar handlers existentes para evitar duplicados
if root_logger.hasHandlers():
    root_logger.handlers.clear()

# Handler para la consola (stdout)
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
root_logger.addHandler(console_handler)
# --- FIN: MODIFICACIÓN ---

logger = get_logger(__file__)



def obtener_bases_desde_json():
    import json
    config_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'archivos_config'))
    config_json_path = os.path.join(config_dir, 'destinatarios_flujos.json')
    bases = ['14']
    try:
        if os.path.exists(config_json_path):
            with open(config_json_path, 'r', encoding='utf-8') as f:
                config_flujos = json.load(f)
                bases = config_flujos.get('bases_trabajo', ['14'])
    except Exception as e:
        logger.warning(f"No se pudo leer bases_trabajo: {e}")
    return bases

def main():
    if len(sys.argv) > 1 and sys.argv[1] == "AUTO":
        logger.info("--- INICIO MODO AUTOMÁTICO ---")
        bases = obtener_bases_desde_json()

        for base in bases:
            base_str = f"RI{base}DB"
            logger.info(f"=== EJECUTANDO PARA LA BASE: {base_str} ===")
            sys.stdout.flush()
            
            logger.info(f"1. Procesando transferencias WMS-INFOR ({base_str})...")
            sys.stdout.flush()
            resultado_trf = procesar_transferencias_wms_infor(BASE=base_str)
            if resultado_trf is False:
                logger.info("   -> No se encontraron transferencias WMS-INFOR pendientes.")
                sys.stdout.flush()
                
            logger.info(f"2. Procesando validaciones de PO ({base_str})...")
            sys.stdout.flush()
            main_PO(BASE=base_str)
            
            logger.info(f"3. Creando PO Automática (Vendor Storage) ({base_str})...")
            sys.stdout.flush()
            
            from ordenes_compra.PO_Automática import extraer_ktrhp, obtener_fecha_proceso
            
            fecha = obtener_fecha_proceso()
            datos_ktrhp = extraer_ktrhp(base_str, fecha)
            
            # Eliminar spools del usuario ELOPEZ
            try:
                from as400_core.ejecutor_cl import EjecutorCL
                cl_spool = EjecutorCL()
                logger.info("   -> Eliminando spools previos del usuario ELOPEZ...")
                cl_spool.ejecutar("DLTSPLF FILE(*ALL) SELECT(ELOPEZ)")
            except Exception as e:
                logger.warning(f"   -> Advertencia al intentar eliminar spools de ELOPEZ: {e}")
            
            if not datos_ktrhp:
                logger.info(f"   -> No hay datos en KTRHP para {base_str}. Omitiendo creación de POs.")
            else:
                # Generar POS antes de validar
                comando_pos = f"CALL PGM(RIUNICOM63/SICRT207CL) PARM('{base}' '')"
                logger.info(f"   -> Ejecutando via SBMJOB: {comando_pos}")
                try:
                    from as400_core.ejecutor_cl import EjecutorCL
                    from conexion_config.conexion import ConexionAS400
                    import time
                    
                    cl = EjecutorCL()
                    resultado = cl.submit_job(comando_pos, job_name=f"TRANSPO{base}")
                    
                    if resultado:
                        job_name = resultado.get('job_name')
                        cl.esperar_trabajo(job_name)
                        
                except Exception as e:
                    logger.error(f"   -> Error al ejecutar SBMJOB {comando_pos}: {e}")

            # Dejar que realice las validaciones
            try:
                main_vendor(BASE=base_str)
                
                # Ejecutar comando AUTOMIR
                try:
                    from as400_core.ejecutor_cl import EjecutorCL
                    cl_automir = EjecutorCL()
                    cmd_automir = "CALL PGM(ELOPEZ/AUTOMIR)"
                    logger.info(f"   -> Ejecutando AUTOMIR via SBMJOB: {cmd_automir}")
                    cl_automir.submit_job(cmd_automir, job_name="AUTOMIR")
                except Exception as e:
                    logger.error(f"   -> Error al ejecutar SBMJOB AUTOMIR: {e}")
                    
            except Exception as e:
                logger.error(f"Error en PO Automática: {e}")
                
        logger.info("--- EJECUTANDO COMANDOS FINALES CL ---")
        try:
            from as400_core.ejecutor_cl import EjecutorCL
            cl_executor = EjecutorCL()
            
            for base in bases:
                
                comando_chg = f"CHGCURLIB CURLIB(RI{base}DB)"
                logger.info(f"   -> Ejecutando: {comando_chg}")
                cl_executor.ejecutar(comando_chg)

                comando_final = f"CALL PGM(RIUNICOM63/SIINF001) PARM('{base}')"
                logger.info(f"   -> Enviando trabajo para país/base {base}: SBMJOB CMD({comando_final})")
                
                # Ejecutamos el job
                resultado_cl = cl_executor.submit_job(comando_final)
                if resultado_cl:
                    logger.info(f"   -> Trabajo enviado exitosamente para {base}.")
                else:
                    logger.error(f"   -> Error al enviar el trabajo CL para {base}.")
        except Exception as e:
            logger.error(f"Error en comando final CL: {e}")
                
        logger.info("--- FIN MODO AUTOMÁTICO ---")
        sys.stdout.flush()  # Forzar la escritura en consola
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
            bases = obtener_bases_desde_json()
            for base in bases:
                base_str = f"RI{base}DB"
                print(f"\nIniciando procesamiento de transferencias WMS-INFOR para {base_str}...")
                procesar_transferencias_wms_infor(BASE=base_str)
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
            bases = obtener_bases_desde_json()
            for base in bases:
                base_str = f"RI{base}DB"
                print(f"\nIniciando validacion PO para {base_str}...")
                main_PO(BASE=base_str)
            input("\nPresiona Enter para volver al menú...")

        elif opcion == "6":
            bases = obtener_bases_desde_json()
            for base in bases:
                base_str = f"RI{base}DB"
                print(f"\nIniciando PO AUTOMATICA VENDOR para {base_str}...")
                main_vendor(BASE=base_str)
            input("\nPresiona Enter para volver al menú...")
        elif opcion == "0":
            print(" ¡Hasta luego!")
            break
        else:
            print(" Opción inválida. Intenta de nuevo.")
            time.sleep(1)

def ejecutar_comando_manual():
    """Permite al usuario ejecutar comandos CL manualmente"""
    from as400_core.ejecutor_cl import EjecutorCL
    
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
