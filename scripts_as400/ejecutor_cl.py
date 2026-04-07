import os
from conexion_config.logging_config import get_logger
import pyodbc
import re
import time
from conexion_config.conexion import ConexionAS400

class EjecutorCL:
    """Clase para ejecutar comandos CL en AS400 con Reintento Automático"""
    
    _instancia = None  # Variable de clase para singleton
    
    def __new__(cls):
        """Implementa patrón singleton"""
        if cls._instancia is None:
            cls._instancia = super(EjecutorCL, cls).__new__(cls)
            cls._instancia.inicializado = False
        return cls._instancia
    
    def __init__(self):
        """Constructor único"""
        if not self.inicializado:
            self.logger = get_logger(__file__)
            self.db_manager = ConexionAS400() # Instancia del gestor de conexión
            self.inicializado = True

    def _ejecutar_sql_robusto(self, sql, params=None):
        """
        Método centralizado que ejecuta SQL con reintentos infinitos 
        hasta que la conexión se establezca.
        """
        espera_segundos = 5
        
        while True: # BUCLE INFINITO DE REINTENTO
            conexion = None
            cursor = None
            try:
                # 1. Obtener conexión (El Singleton de ConexionAS400 maneja la lógica interna)
                conexion = self.db_manager.conectar()
                
                if not conexion:
                    raise Exception("La conexión retornó None.")

                # 2. Ejecutar
                cursor = conexion.cursor()
                if params:
                    cursor.execute(sql, params)
                else:
                    cursor.execute(sql)
                
                conexion.commit()
                
                # 3. Si es un SELECT o devuelve algo, intentar capturarlo (opcional)
                try:
                    # QCMDEXC generalmente no devuelve filas, pero SBMJOB podría devolver info si se ajusta
                    if cursor.description:
                        return cursor.fetchall()
                except:
                    pass

                return True # Éxito

            except Exception as e:
                error_msg = str(e).lower()
                
                # 4. DETECCIÓN DE ERRORES DE CONEXIÓN
                if any(x in error_msg for x in ['closed connection', 'communication link failure', 'invalid cursor state', 'none', 'connection is busy']):
                    self.logger.warning(f"⚠️ Conexión perdida ({e}). Reintentando en {espera_segundos}s...")
                    
                    # CRÍTICO: Forzar al Singleton a limpiar la conexión vieja
                    try:
                        if hasattr(self.db_manager, 'cerrar_conexion'):
                            self.db_manager.cerrar_conexion()
                    except:
                        pass

                    time.sleep(espera_segundos)
                    continue # Vuelve al inicio del While
                
                else:
                    # Errores de sintaxis o lógica del comando (No se arreglan reconectando)
                    self.logger.error(f"❌ Error fatal ejecutando SQL: {e}")
                    raise e # Romper flujo si el comando está mal escrito

            finally:
                # Limpieza de cursor segura
                try:
                    if cursor: cursor.close()
                except:
                    pass

    def ejecutar(self, comando_cl):
        """
        Ejecuta un comando CL directo (CALL, DLTE, CPYF, etc.)
        """
        self.logger.info(f"Ejecutando CL: {comando_cl}")
        
        # Formato requerido por QCMDEXC: Comando y Longitud(15,5)
        longitud = len(comando_cl)
        
        # Usamos parámetros (?) para evitar problemas de formato de strings
        sql = "CALL QSYS2.QCMDEXC(?, ?)"
        
        # Ejecutamos con el sistema de reintentos
        return self._ejecutar_sql_robusto(sql, (comando_cl, float(longitud)))

    def submit_job(self, comando, job_name=None, job_queue="QBATCH", job_desc=None):
        """
        Envía un trabajo batch (SBMJOB) con reintento automático.
        """
        # 1. Generar nombre de trabajo si no existe
        if not job_name:
            base_name = re.sub(r'[^A-Za-z0-9]', '', comando.split(' ')[0])
            job_name = f"{base_name[:6]}{int(time.time()) % 1000}"
        
        self.logger.info(f"Preparando SBMJOB: {job_name} -> {comando}")

        # 2. Construir el comando SBMJOB completo
        sbmjob_cmd = f"SBMJOB CMD({comando}) JOB({job_name})"
        
        if job_queue:
            sbmjob_cmd += f" JOBQ({job_queue})"
        if job_desc:
            sbmjob_cmd += f" JOBD({job_desc})"

        # 3. Preparar ejecución
        longitud = len(sbmjob_cmd)
        sql = "CALL QSYS2.QCMDEXC(?, ?)"
        
        # 4. Ejecutar con robustez
        exito = self._ejecutar_sql_robusto(sql, (sbmjob_cmd, float(longitud)))
        
        if exito:
            self.logger.info(f"[OK] Trabajo enviado: {job_name}")
            return {
                'job_name': job_name,
                'submitted': True
            }
        return False

    