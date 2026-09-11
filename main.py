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

def ejecutar_automatizacion():
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

        if not datos_ktrhp:
            logger.info(f"   -> No hay datos en KTRHP para {base_str}. Omitiendo creación de POs y conciliación.")
        else:
            # Eliminar spools del usuario ELOPEZ
            try:
                from as400_core.ejecutor_cl import EjecutorCL
                cl_spool = EjecutorCL()
                logger.info("   -> Eliminando spools previos del usuario ELOPEZ...")
                cl_spool.ejecutar("DLTSPLF FILE(*SELECT) SELECT(ELOPEZ)")
            except Exception as e:
                logger.warning(f"   -> Advertencia al intentar eliminar spools de ELOPEZ: {e}")
            
            # Generar POS antes de validar
            comando_pos = f"CALL PGM(RIUNICOM63/SICRT207CL) PARM('{base}' '')"
            logger.info(f"   -> Ejecutando via SBMJOB: {comando_pos}")
            try:
                from as400_core.ejecutor_cl import EjecutorCL
                from conexion_config.conexion import ConexionAS400
                import time
                
                cl = EjecutorCL()
                
                # Cambiar la biblioteca actual antes de ejecutar SBMJOB
                comando_chg = f"CHGCURLIB CURLIB({base_str})"
                logger.info(f"   -> Ejecutando previo al job: {comando_chg}")
                cl.ejecutar(comando_chg)
                
                resultado = cl.submit_job(comando_pos, job_name=f"TRANSPO{base}")
                
                if resultado:
                    job_name = resultado.get('job_name')
                    # Le ponemos un límite de 15 intentos (~4 min) para trabajos rápidos como TRANSPO
                    cl.esperar_trabajo(job_name, limite_sin_ver=15)
                    
            except Exception as e:
                logger.error(f"   -> Error al ejecutar SBMJOB {comando_pos}: {e}")

            # Dejar que realice las validaciones (Conciliación y AUTOMIR)
            try:
                conciliacion_ok = main_vendor(BASE=base_str)
                
                # Ejecutar comando AUTOMIR solo si hubo datos de conciliación
                if conciliacion_ok:
                    try:
                        from as400_core.ejecutor_cl import EjecutorCL
                        cl_automir = EjecutorCL()
                        cmd_automir = "CALL PGM(ELOPEZ/AUTOMIR)"
                        logger.info(f"   -> Ejecutando AUTOMIR via SBMJOB: {cmd_automir}")
                        res_automir = cl_automir.submit_job(cmd_automir, job_name="AUTOMIR")
                        
                        if res_automir:
                            logger.info("   -> AUTOMIR sometido correctamente (no se monitoreará su finalización según configuración).")
                            # cl_automir.esperar_trabajo("AUTOMIR")
                    except Exception as e:
                        logger.error(f"   -> Error al ejecutar SBMJOB AUTOMIR: {e}")
                else:
                    logger.info("   -> Omitiendo ejecución de AUTOMIR ya que no hubo datos de conciliación (KTRHP/KPUHP vacíos).")
                    
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

def actualizar_estado(proceso, estado):
    """
    Lee o actualiza el estado del proceso en la base de datos de forma atómica.
    Si estado es None, devuelve el estado actual ('Y' o 'N').
    Si estado es 'Y' o 'N', lo actualiza.
    Abre y cierra la conexión inmediatamente para evitar timeouts en procesos largos.
    """
    from conexion_config.conexion import ConexionAS400
    db = ConexionAS400()
    conexion = db.conectar()
    if not conexion:
        logger.error("No se pudo conectar a DB2 para gestión de estado.")
        return None

    try:
        with conexion.cursor() as cursor:
            if estado is None:
                cursor.execute(f"SELECT ESTADO_ACTIVO FROM ELOPEZ.PROCESO_ESTADO WHERE PROCESO = '{proceso}'")
                row = cursor.fetchone()
                return row[0].upper() if row else None
            else:
                cursor.execute(f"UPDATE ELOPEZ.PROCESO_ESTADO SET ESTADO_ACTIVO = '{estado}' WHERE PROCESO = '{proceso}'")
                conexion.commit()
                return True
    except Exception as e:
        logger.error(f"Error gestionando estado {proceso}: {e}")
        return False
    finally:
        db.cerrar_conexion()

def main():
    if len(sys.argv) <= 1 or sys.argv[1] != "AUTO":
        logger.warning("El script requiere el argumento 'AUTO' para iniciar. Saliendo...")
        return

    logger.info("Verificando si hay procesos activos en ELOPEZ.PROCESO_ESTADO...")
    
    estado_actual = actualizar_estado('AUTOMATIZACION', None)
    
    if estado_actual == 'Y':
        print("\n[!] ADVERTENCIA: Ya existe un proceso de AUTOMATIZACION en ejecución.")
        logger.warning("Ejecución detenida: La bandera ESTADO_ACTIVO está en 'Y'.")
        return
        
    logger.info("Marcando inicio del proceso (ESTADO_ACTIVO = 'Y')...")
    if not actualizar_estado('AUTOMATIZACION', 'Y'):
        print("No se pudo bloquear el proceso en DB2. Saliendo por seguridad.")
        return

    try:
        ejecutar_automatizacion()
    except Exception as e:
        logger.exception(f"Error general en la automatización: {e}")
    finally:
        logger.info("Liberando bandera de ejecución (ESTADO_ACTIVO = 'N')...")
        actualizar_estado('AUTOMATIZACION', 'N')

if __name__ == "__main__":
    if not os.path.exists('logs'):
        os.makedirs('logs')
    main()
