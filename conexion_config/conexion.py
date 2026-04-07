import pyodbc
import os
from conexion_config.logging_config import get_logger

# Singleton para mantener una única instancia de conexión
_CONEXION_SINGLETON = None

class ConexionAS400:
    """Clase para manejar la conexión a AS400 con patrón Singleton"""
    
    def __init__(self):
        # CORRECCIÓN IMPORTANTE: Asignamos el objeto logger a self.logger
        self.logger = get_logger(__file__)
        self.conexion = None
        self.credenciales = self.leer_credenciales()
        self.driver_funcional = None  # Guarda el driver que funcionó
    
    def leer_credenciales(self):
        """Lee las credenciales de conexión desde un archivo de configuración"""
        cred_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'as400_credentials.txt'))
        
        # Valores por defecto
        credenciales = {
            'DSN': 'AS400',
            'UID': 'ELOPEZ',
            'PWD': 'MAY2024',
            'DATABASE': 'RI11DB',
            'SYSTEM': '192.168.151.64'
        }
        
        # Intentar leer desde archivo
        if os.path.exists(cred_path):
            try:
                with open(cred_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.strip() and not line.strip().startswith('#'):
                            key, value = line.strip().split('=', 1)
                            credenciales[key.strip().upper()] = value.strip()
            except Exception as e:
                self.logger.error(f"Error leyendo credenciales: {e}")
        else:
            # Crear archivo con valores por defecto
            try:
                os.makedirs(os.path.dirname(cred_path), exist_ok=True)
                with open(cred_path, 'w', encoding='utf-8') as f:
                    f.write("# Credenciales para conexión a AS400\n")
                    for key, value in credenciales.items():
                        f.write(f"{key}={value}\n")
                self.logger.info(f"Archivo de credenciales creado en: {cred_path}")
            except Exception as e:
                self.logger.error(f"Error creando archivo de credenciales: {e}")
        
        return credenciales
    
    def conectar(self):
        """Establece conexión con AS400 usando conexión directa ODBC"""
        global _CONEXION_SINGLETON
        
        # 1. Verificar si ya existe una conexión activa (Singleton)
        if _CONEXION_SINGLETON is not None:
            try:
                # Validar que la conexión siga viva
                cursor = _CONEXION_SINGLETON.cursor()
                cursor.execute("SELECT 1 FROM SYSIBM.SYSDUMMY1")
                cursor.fetchone()
                cursor.close()
                
                # Está viva, la reutilizamos
                self.conexion = _CONEXION_SINGLETON
                return self.conexion
            except:
                # Está muerta, limpiamos
                self.logger.warning("Conexión previa perdida. Reconectando...")
                try:
                    _CONEXION_SINGLETON.close()
                except:
                    pass
                _CONEXION_SINGLETON = None
        
        # 2. Intentar establecer nueva conexión
        try:
            # Opción A: Si ya sabemos qué driver funcionó antes, lo usamos
            if self.driver_funcional:
                try:
                    conn_str = self._construir_cadena_conexion(self.driver_funcional)
                    self.conexion = pyodbc.connect(conn_str)
                    _CONEXION_SINGLETON = self.conexion
                    return self.conexion
                except Exception as e:
                    self.logger.warning(f"Fallo al reconectar con driver conocido: {e}")
                    self.driver_funcional = None # Resetear para intentar buscar de nuevo

            # Opción B: Probar con el driver estándar de IBM i (iSeries Access)
            drivers_a_probar = [
                "iSeries Access ODBC Driver",
                "Client Access ODBC Driver (32-bit)",
                "IBM i Access ODBC Driver"
            ]

            for driver in drivers_a_probar:
                try:
                    conn_str = self._construir_cadena_conexion(driver)
                    self.logger.info(f"Intentando conectar con: {driver}...")
                    
                    self.conexion = pyodbc.connect(conn_str)
                    
                    # Si llegamos aquí, funcionó
                    self.driver_funcional = driver
                    _CONEXION_SINGLETON = self.conexion
                    self.logger.info(f"[OK] Conectado exitosamente con {driver}")
                    return self.conexion
                except:
                    continue # Si falla, prueba el siguiente
            
            # Si llegamos aquí, ningún driver funcionó
            self.logger.error("No se pudo conectar a AS400. Verifique drivers ODBC instalados.")
            return None

        except Exception as e:
            self.logger.error(f"Error crítico al conectar a AS400: {e}")
            return None
    
    def _construir_cadena_conexion(self, driver):
        """Helper para construir el string de conexión según el driver"""
        return (
            f"DRIVER={{{driver}}};"
            f"SYSTEM={self.credenciales['SYSTEM']};"
            f"UID={self.credenciales['UID']};"
            f"PWD={self.credenciales['PWD']};"
            f"DBNAME={self.credenciales['DATABASE']};"
            f"CMT=0;DFT=3;LANGUAGEID=ENU;PKG=QGPL/DEFAULT(IBM),2,0,1,0,512;TRANSLATE=1"
        )

    def cerrar_conexion(self):
        """Cierra la conexión con AS400"""
        global _CONEXION_SINGLETON
        
        if self.conexion:
            try:
                self.conexion.close()
                self.conexion = None
                _CONEXION_SINGLETON = None
                return True
            except Exception as e:
                self.logger.error(f"Error al cerrar conexión: {e}")
                return False
        return True