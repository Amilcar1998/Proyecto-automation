import paramiko
import xml.etree.ElementTree as ET
import io
import datetime
import random
import os
import sys

# Agregar la raíz del proyecto al sys.path para poder importar conexion_config
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from conexion_config.conexion import ConexionAS400

def get_random_date_current_month():
    """Retorna una fecha aleatoria dentro del mes actual en formato YYYYMMDD"""
    now = datetime.datetime.now()
    year = now.year
    month = now.month
    
    # Determinar el último día del mes
    if month == 12:
        last_day = 31
    else:
        last_day = (datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)).day
        
    random_day = random.randint(1, last_day)
    random_date = datetime.date(year, month, random_day)
    return random_date.strftime("%Y%m%d")

def get_matching_upcs_from_db(upcs_list, conexion):
    """
    Recibe una lista de UPCs y consulta la base de datos para ver si alguno pertenece a OPTICA o AUTOMOTRIZ.
    Retorna True si hay al menos una coincidencia, de lo contrario False.
    """
    if not upcs_list:
        return False
        
    try:
        cursor = conexion.cursor()
        
        # Generar los marcadores de posición para la consulta SQL
        placeholders = ','.join(['?'] * len(upcs_list))
        query = f"SELECT UPC_CORP FROM RI11DB.CONSULRP3 WHERE UPC_CORP IN ({placeholders}) AND DES_DIV IN ('AUTOMOTRIZ', 'OPTICA')"
        
        cursor.execute(query, upcs_list)
        result = cursor.fetchone()
        
        cursor.close()
        
        # Si fetchone() retorna algo, significa que hay al menos un match
        return bool(result)
    except Exception as e:
        print(f"Error consultando la base de datos: {e}")
        return False

def process_and_move_xmls(sftp_host, sftp_port, sftp_user, sftp_pass, sftp_archive_path, sftp_in_path):
    # 1. Establecer conexión a BD
    print("Iniciando conexión a la Base de Datos AS400...")
    db_manager = ConexionAS400()
    conexion_db = db_manager.conectar()
    
    if not conexion_db:
        print("Error crítico: No se pudo conectar a la base de datos AS400. Abortando script.")
        return

    # 2. Conectar al SFTP
    try:
        print(f"Conectando por SFTP a {sftp_host}:{sftp_port}...")
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        
        # Conectarse (sin disabled_algorithms para mantener compatibilidad con la v2.12.0)
        ssh.connect(hostname=sftp_host, port=sftp_port, username=sftp_user, password=sftp_pass)
        
        sftp = ssh.open_sftp()
        print(f"Cambiando al directorio {sftp_archive_path}...")
        sftp.chdir(sftp_archive_path)
        
        # Obtener lista de archivos en el directorio
        files = sftp.listdir()
        xml_files = [f for f in files if f.lower().endswith('.xml')][:10] # Limitado a 10 archivos para pruebas
        print(f"Se procesarán {len(xml_files)} archivos XML (limitado para prueba).")
        
        for xml_file in xml_files:
            print(f"\nAnalizando archivo: {xml_file}")
            
            # Descargar el contenido del archivo en memoria
            r = io.BytesIO()
            try:
                with sftp.open(xml_file, 'r') as remote_file:
                    r.write(remote_file.read())
                r.seek(0)
            except Exception as e:
                print(f"Error al leer {xml_file}: {e}")
                continue
            
            try:
                tree = ET.parse(r)
                root = tree.getroot()
                
                # Extraer los UPCs únicos del archivo
                upcs = set()
                for transaction in root.findall('transactionDetail'):
                    for inventory_detail in transaction.findall('inventoryDetail'):
                        upc = inventory_detail.findtext('upc', '')
                        if upc:
                            # Aseguramos que el UPC se envíe como string limpio
                            upcs.add(str(upc).strip())
                
                # Consultar la BD solo si el archivo contenía UPCs
                is_match = False
                if upcs:
                    is_match = get_matching_upcs_from_db(list(upcs), conexion_db)
                
                if is_match:
                    print(f" -> ¡Match Encontrado! Pertenece a Óptica/Automotriz.")
                    
                    # Alterar las fechas
                    fechas_modificadas = 0
                    for transaction in root.findall('transactionDetail'):
                        t_date = transaction.find('transactionDate')
                        if t_date is not None:
                            t_date.text = get_random_date_current_month()
                            fechas_modificadas += 1
                    
                    print(f" -> Se modificaron {fechas_modificadas} fechas.")
                    
                    # Guardar el XML modificado en memoria
                    modified_xml_io = io.BytesIO()
                    modified_xml_io.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
                    tree.write(modified_xml_io, encoding='utf-8', xml_declaration=False)
                    modified_xml_io.seek(0)
                    
                    # Subir a la carpeta IN (asumiendo que la ruta real usa minúsculas y está separada por país)
                    # Extraer el país del nombre (Ej: InventoryDischarge_SV_... -> SV)
                    pais = xml_file.split('_')[1] if '_' in xml_file else 'SV'
                    in_file_path = f"/integration/uc4/out/InventoryDischarge/{xml_file}"
                    
                    try:
                        with sftp.open(in_file_path, 'w') as out_file:
                            out_file.write(modified_xml_io.read().decode('utf-8'))
                    except IOError:
                        print(f" -> Error IOError subiendo a {in_file_path}. ¿Existe la carpeta destino?")
                        raise
                    
                    # Eliminar el archivo original de la carpeta archive
                    sftp.remove(xml_file)
                    print(f" -> Archivo movido a 'in' y eliminado de 'archive'.")
                else:
                    print(" -> Sin match. Se ignora el archivo.")
                    
            except ET.ParseError:
                print(f"Error: El archivo {xml_file} no es un XML válido o está corrupto.")
                
        # Cerrar conexiones
        sftp.close()
        ssh.close()
        db_manager.cerrar_conexion()
        print("\nProceso completado exitosamente.")
        
    except paramiko.AuthenticationException:
        print("Error de Autenticación: Verifica tu usuario y contraseña SFTP.")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

if __name__ == "__main__":
    # Configuración de credenciales y servidor SFTP
    SFTP_HOST = "UPOSAP01UAT.unicomer.com" 
    SFTP_PORT = 22
    SFTP_USER = "uc4uat"
    SFTP_PASS = "Uc4engin3"
    
    # Directorios Origen y Destino
    SFTP_ARCHIVE_PATH = "/integration/uc4/archive/InventoryDischarge"
    SFTP_IN_PATH = "/integration/uc4/in/InventoryDischarge"
    
    process_and_move_xmls(SFTP_HOST, SFTP_PORT, SFTP_USER, SFTP_PASS, SFTP_ARCHIVE_PATH, SFTP_IN_PATH)
