import subprocess
import sys
import os
import shutil
from datetime import datetime
import pandas as pd
import pyodbc
import logging
import traceback

# --- CONFIGURACIÓN DE LOGGING ---
LOG_FILE = f"proceso_insercion_{datetime.now().strftime('%Y%m%d')}.log"

# Configuración base
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Logger por módulo
logger = logging.getLogger(__name__)

# Añadir file handler local si no existe
log_abspath = os.path.abspath(LOG_FILE)
has_file = False

for h in list(logger.handlers):
    if isinstance(h, logging.FileHandler):
        try:
            if os.path.abspath(getattr(h, 'baseFilename', '')) == log_abspath:
                has_file = True
                break
        except Exception:
            continue

if not has_file:
    fh = logging.FileHandler(log_abspath, mode='w', encoding='utf-8')
    fh.setLevel(logging.INFO)
    fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(fmt)
    logger.addHandler(fh)

# --- CONFIGURACIÓN DE CONEXIÓN ODBC ---
DSN_NAME = "RI_TEST"  
ODBC_UID = "ELOPEZ"
ODBC_PWD = "MAY2024"
SCHEMA_NAME = "RI12DB"
TABLE_NAME = "RPLORDEN"

# Se agrega CMTS=0 (Commitment Control None) para evitar error SQL7008 en IBM i
CONNECTION_STRING = f"DSN={DSN_NAME};UID={ODBC_UID};PWD={ODBC_PWD};CMTS=0;"

# --- FUNCIONES DE INSTALACIÓN ---
def install_package(package):
    try:
        if package == 'pyodbc':
            import pyodbc
        else:
            __import__(package)
    except ImportError:
        logger.warning(f"Instalando paquete '{package}'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

required_packages = ['pandas', 'pyodbc', 'openpyxl']
for package in required_packages:
    install_package(package)

# --- FUNCIONES DE BASE DE DATOS ---

def get_odbc_connection():
    """Establece conexión desactivando el control de compromiso para evitar SQL7008."""
    try:
        conn = pyodbc.connect(CONNECTION_STRING)
        # Para tablas sin Journal, autocommit debe ser True
        conn.autocommit = True 
        logger.info(f"Conexión a DB via DSN '{DSN_NAME}' (Sin Commitment Control) establecida.")
        return conn
    except pyodbc.Error as e:
        logger.error(f"[ERROR] Error CRÍTICO al conectar: {e}")
        sys.exit(1)

def execute_insert_db(cursor, insert_query, row_data):
    """Ejecuta la inserción y captura errores específicos de SQL."""
    try:
        cursor.execute(insert_query)
        return True
    except pyodbc.Error as e:
        logger.error(f"[ERROR] Error de SQL en registro ID {row_data.get('ID', 'N/A')}: {e}")
        logger.debug(f"SQL FALLIDO: {insert_query}")
        return False

# --- PROCESO PRINCIPAL ---

def main():
    logger.info("Iniciando proceso de inserción de Órdenes de Compra (PO).")
    
    # Limpieza de carpetas temporales
    for carpeta in ["local", "Importado"]:
        if os.path.exists(carpeta):
            shutil.rmtree(carpeta)
            logger.info(f"Carpeta '{carpeta}' borrada.")

    fecha_PO = datetime.now().strftime('%Y%m%d')

    # Carga de Excel
    try:
        archivo_excel = "Formatos/OC compras.xlsx"
        df = pd.read_excel(archivo_excel)
        logger.info(f"Archivo cargado: {len(df)} filas encontradas.")
    except Exception as e:
        logger.critical(f"[ERROR] No se pudo cargar el archivo: {e}")
        return

    # Conexión
    db_conn = get_odbc_connection()
    db_cursor = db_conn.cursor()

    total_inserts = 0
    total_processed = 0

    # Agrupación y Procesamiento
    try:
        # Agrupamos para mantener la lógica original del script
        grouped = df.groupby(['ID', 'TIPO', 'PAIS', 'PREDISTRIBUIDO'])

        for (id_valor, tipo_valor, pais_valor, predistribuido_valor), group in grouped:
            for row_index, row in group.iterrows():
                total_processed += 1

                # Lógica de País e ID
                tipo_valor_insert = 'I' if str(row['PAIS']).upper() == 'IMPORTADO' else 'L'

                print(f"tipo_valor_insert:{tipo_valor_insert}")

                if(tipo_valor_insert) =='I':
                    proveedor ='010479'
                else:
                    proveedor ='013873'
                
                try:
                    id = str(int(row['ID'])).zfill(6)
                except:
                    id = '000000'

                llave = f"{fecha_PO}-{proveedor}-{id}"

                print(f"proveedor: {llave}")

                fila_contador = total_processed # Contador incremental

                # Construcción del INSERT
                # Nota: Asegúrate de que el orden de las columnas en VALUES coincida con RPLORDEN
                insert_query = (
                    f"INSERT INTO {SCHEMA_NAME}.{TABLE_NAME} "
                    f"VALUES ('02', '02', '999', '{llave}', '{row['TIENDA']}', '{proveedor}', "
                    f"'256', 'P', 0, '{fecha_PO}', '{fecha_PO}', '{fila_contador}', "
                    f"'{row['ITEMS']}', '{row['CANTIDADES']}', '0', "
                    f"'{SCHEMA_NAME}', 'RIBLY', 'PEGASO', '2       ', '42', ' ', '{tipo_valor_insert}')"
                )

                if execute_insert_db(db_cursor, insert_query, row):
                    total_inserts += 1
                    if total_inserts % 10 == 0:
                        logger.info(f"Progreso: {total_inserts} registros insertados...")
                
    except Exception as e:
        logger.error(f"[ERROR] Error durante el procesamiento: {e}")
        logger.error(traceback.format_exc())
    finally:
        db_cursor.close()
        db_conn.close()
        logger.info(f"--- RESUMEN FINAL ---")
        logger.info(f"Filas procesadas: {total_processed}")
        logger.info(f"Inserciones exitosas: {total_inserts}")
        logger.info("Conexión cerrada.")

if __name__ == "__main__":
    main()