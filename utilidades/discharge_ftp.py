import ftplib
import os
import time
import datetime
import logging
import sys
import itertools
from datetime import timedelta

# Asegurar que se puede importar conexion_config
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from conexion_config.conexion import ConexionAS400

FTP_HOST = '192.168.151.64'
FTP_USER = 'ELOPEZ'
FTP_PASS = 'MAY2024'

fecha_actual = time.strftime('%Y%m%d')

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
LOGGER = logging.getLogger("FTP_INV_DISCHARGE")


def clean_field(v: str) -> str:
    """
    Normaliza un campo para que:
    - None -> ""
    - "   " -> ""   (evita | | y genera ||)
    - " texto " -> "texto" (trim)
    """
    if v is None:
        return ""
    s = str(v)
    return "" if s.strip() == "" else s.strip()


def normalize_pipe_line(line: str) -> str:
    """
    Normaliza una línea separada por pipes:
    - Quita \n
    - Split por '|'
    - Limpia cada campo (espacios -> vacío)
    - Reconstruye y termina con '\n'
    - Previene que Java ignore campos vacíos al final
    """
    raw = line.rstrip("\n")
    parts = raw.split("|")
    parts = [clean_field(p) for p in parts]
    
    # Si el último elemento es vacío, le ponemos un espacio 
    # para evitar que el split("\\|") de Java reduzca la longitud del arreglo.
    if len(parts) > 0 and parts[-1] == "":
        parts[-1] = " "
        
    return "|".join(parts) + "\n"

def obtener_upcs_de_bd(pais):

    if pais == 'GT':
        npais = 'RI12DB'
    elif pais == 'SV':
        npais = 'RI11DB'
    elif pais == 'NI':
        npais = 'RI14DB'
    elif pais == 'HN':
        npais = 'RI13DB'
    



    try:
        db_manager = ConexionAS400()
        conexion = db_manager.conectar()
        if not conexion:
            return []
            
        cursor = conexion.cursor()
        
        # Ejecutar DELETEs primero
        deletes = [
            f"DELETE FROM {npais}.ORPOSINDH",
            f"DELETE FROM {npais}.ORPOSINDD",
            f"DELETE FROM {npais}.KSWMP",
            f"DELETE FROM {npais}.SIF108P WHERE D88SI = '{fecha_actual}'"
        ]
        
        for d in deletes:
            try:
                cursor.execute(d)
            except Exception as e:
                LOGGER.error(f"Error ejecutando DELETE: {d} -> {e}")
                
        conexion.commit()

        query = f"""
        SELECT C.UPC_CORP 
        FROM {npais}.KSKUP k 
        INNER JOIN RIUSRMOD63.SIF106P sp ON SP.SIFPOVENDO = K.VNÑSK AND sp.SIFPODEPTO = K.DPÑSK
        INNER JOIN RI11DB.CONSULRP3 C ON C.SKU_CORP = K.SKUSK AND C.COD_PROV = SP.SIFPOVENDO
        WHERE k.SRLSK <> 'D' AND K.B34SK > 0 ORDER BY c.INT_UPDATE DESC 
        """
        
        print(query)



        cursor.execute(query)
        results = cursor.fetchall()
        cursor.close()
        db_manager.cerrar_conexion()
        
        upcs = []
        for row in results:
            if row[0]:
                upcs.append(str(row[0]).strip())
        return upcs
    except Exception as e:
        LOGGER.error(f"Error obteniendo UPCs de BD: {e}")
        return []

def modificar_y_mover_ftp(pais, dias_atras=5):
    fecha_desde = (datetime.date.today() - timedelta(days=dias_atras)).strftime('%Y-%m-%d')
    fecha_hasta = time.strftime('%Y-%m-%d')
    
    LOGGER.info(f"INICIO modificar_y_mover_ftp | pais={pais} dias_atras={dias_atras}")
    LOGGER.info(f"Rango fechas | desde={fecha_desde} hasta={fecha_hasta}")

    try:
        CARPETA_BUSQUEDA = f'/integration/uc4/archive/InventoryDischarge/{pais}'
        LOCAL_DOWNLOAD_DIR = f'archivos_FTP/descargas_ftp_{pais}'
        DESTINO_DIR = f'/integration/uc4/in/inventorydischarge/{pais}'

        fecha_fin = datetime.datetime.today()
        fecha_inicio = fecha_fin - timedelta(days=dias_atras)

        LOGGER.info(f"Rango fechas | desde={fecha_inicio.strftime('%Y-%m-%d')} hasta={fecha_fin.strftime('%Y-%m-%d')}")

        if not os.path.exists(LOCAL_DOWNLOAD_DIR):
            os.makedirs(LOCAL_DOWNLOAD_DIR)
            LOGGER.info(f"Carpeta local creada | dir={LOCAL_DOWNLOAD_DIR}")

        with ftplib.FTP(FTP_HOST, timeout=10) as ftp:
            ftp.login(FTP_USER, FTP_PASS)
            LOGGER.info(f"Conectado al FTP | host={FTP_HOST} pais={pais}")

            archivos = ftp.nlst(CARPETA_BUSQUEDA)

            archivos_encontrados = []
            prefijo = f"InventoryDischarge_{pais}_"

            for archivo in archivos:
                nombre = os.path.basename(archivo)

                if not nombre.startswith(prefijo):
                    continue

                resto = nombre[len(prefijo):]

                if len(resto) < 8 or not resto[:8].isdigit():
                    LOGGER.warning(f"No se pudo extraer fecha válida | archivo={nombre}")
                    continue

                fecha_str = resto[:8]

                try:
                    fecha_archivo = datetime.datetime.strptime(fecha_str, "%Y%m%d")
                except Exception:
                    LOGGER.warning(f"Fecha inválida detectada | archivo={nombre} fecha={fecha_str}")
                    continue

                if fecha_inicio.date() <= fecha_archivo.date() <= fecha_fin.date():
                    archivos_encontrados.append(archivo)
                    LOGGER.info(f"Archivo dentro de rango | archivo={nombre} fecha={fecha_str}")

            # Identificar y eliminar duplicados manteniendo el más reciente
            archivos_por_base = {}
            for archivo in archivos_encontrados:
                nombre = os.path.basename(archivo)
                if ".txt" in nombre:
                    base = nombre.split(".txt")[0] + ".txt"
                else:
                    base = nombre
                
                if base not in archivos_por_base:
                    archivos_por_base[base] = []
                archivos_por_base[base].append(archivo)
                
            archivos_finales = []
            for base, lista_archivos in archivos_por_base.items():
                if len(lista_archivos) > 1:
                    lista_archivos.sort()  # El último será el que tiene más timestamps (el más reciente)
                    archivos_finales.append(lista_archivos[-1])
                    
                    # Eliminar los duplicados del FTP
                    for duplicado in lista_archivos[:-1]:
                        LOGGER.warning(f"Eliminando archivo duplicado en archive | archivo={duplicado}")
                        try:
                            ftp.delete(duplicado)
                        except Exception as e:
                            LOGGER.error(f"Error eliminando duplicado {duplicado} | error={e}")
                else:
                    archivos_finales.append(lista_archivos[0])
            
            archivos_encontrados = archivos_finales

            if not archivos_encontrados:
                LOGGER.warning(f"No se encontraron archivos dentro del rango | pais={pais}")
                return

            upcs_bd = obtener_upcs_de_bd(pais)
            if upcs_bd:
                LOGGER.info(f"Se obtuvieron {len(upcs_bd)} UPCs de la BD para reemplazo múltiple.")
                ciclo_upcs = itertools.cycle(upcs_bd)
            else:
                LOGGER.warning("No se pudieron obtener UPCs de la BD. Se mantendrán los originales.")
                ciclo_upcs = None

            LOGGER.info(f"Total archivos a procesar | cantidad={len(archivos_encontrados)}")

            for archivo_encontrado in archivos_encontrados:
                LOGGER.info(f"INICIO archivo | path={archivo_encontrado}")
                local_file_path = os.path.join(LOCAL_DOWNLOAD_DIR, os.path.basename(archivo_encontrado))

                try:
                    # Descargar
                    with open(local_file_path, 'wb') as f:
                        ftp.retrbinary(f'RETR {archivo_encontrado}', f.write)

                    # Leer texto
                    with open(local_file_path, 'r', encoding='utf-8') as f:
                        lineas = f.readlines()

                    anio_mes_actual = fecha_actual[:6]

                    lineas_modificadas = []
                    archivo_invalido = False
                    trx_invalidas = []

                    for num_linea, linea in enumerate(lineas, 1):
                        # Normaliza para eliminar trailing spaces y "campos espacio"
                        # PERO ojo: para HDR/TRX procesamos específico, para lo demás normalizamos genérico.

                        if linea.startswith('HDR|'):
                            raw = linea.rstrip('\n')
                            partes = raw.split('|')
                            partes = [clean_field(p) for p in partes]

                            if len(partes) > 0 and partes[-1] == "":
                                partes[-1] = " "

                            # Validación estricta como en Java (InventoryDischargeReq)
                            if len(partes) != 4:
                                archivo_invalido = True
                                trx_invalidas.append(num_linea)
                                LOGGER.error(f"HDR inválida | linea={num_linea} campos={len(partes)} raw={repr(raw)}")
                                break

                            # Regla original de tu HDR
                            if len(partes) > 3 and len(partes[3]) == 14 and partes[3].isdigit():
                                partes[3] = anio_mes_actual + partes[3][6:]

                            lineas_modificadas.append("|".join(partes) + "\n")
                            continue

                        if linea.startswith('TRX|'):
                            raw = linea.rstrip('\n')
                            partes = raw.split('|')

                            # Limpia campos (convierte " " en "")
                            partes = [clean_field(p) for p in partes]
                            
                            if len(partes) > 0 and partes[-1] == "":
                                partes[-1] = " "

                            # Validación estricta para TRX (17 campos separados por 16 pipes)
                            if len(partes) != 17:
                                archivo_invalido = True
                                trx_invalidas.append(num_linea)
                                LOGGER.error(f"TRX inválida | linea={num_linea} campos={len(partes)} raw={repr(raw)}")
                                break

                            # Regla original: actualizar fechas YYYYMMDD en posiciones 7..15
                            for idx in range(7, 16):
                                if len(partes[idx]) == 8 and partes[idx].isdigit():
                                    partes[idx] = anio_mes_actual + partes[idx][6:]

                            lineas_modificadas.append("|".join(partes) + "\n")
                            continue

                        if linea.startswith('INV|'):
                            raw = linea.rstrip('\n')
                            partes = raw.split('|')
                            partes = [clean_field(p) for p in partes]
                            
                            if len(partes) > 0 and partes[-1] == "":
                                partes[-1] = " "
                                
                            # Reemplazar el UPC ciclando por la lista obtenida
                            if ciclo_upcs and len(partes) > 1:
                                partes[1] = next(ciclo_upcs)
                                
                            lineas_modificadas.append("|".join(partes) + "\n")
                            continue

                        # Otras líneas: normaliza en general para quitar trailing y campos " "
                        lineas_modificadas.append(normalize_pipe_line(linea))

                    if archivo_invalido:
                        LOGGER.error(f"Archivo inválido será eliminado | archivo={archivo_encontrado} trx_invalidas={trx_invalidas}")

                        try:
                            ftp.delete(archivo_encontrado)
                            LOGGER.warning(f"Eliminado del FTP | archivo={archivo_encontrado}")
                        except Exception as e:
                            LOGGER.exception(f"Error en bloque de borrado FTP | error={e}")

                        try:
                            os.remove(local_file_path)
                            LOGGER.warning(f"Eliminado local | archivo={local_file_path}")
                        except Exception as e:
                            LOGGER.exception(f"Error eliminando local | error={e}")

                        continue

                    # Guardar archivo modificado
                    with open(local_file_path, 'w', encoding='utf-8', newline='\n') as f:
                        f.writelines(lineas_modificadas)

                    nombre_archivo_orig = os.path.basename(local_file_path)
                    if nombre_archivo_orig.startswith(prefijo) and len(nombre_archivo_orig[len(prefijo):]) >= 8:
                        fecha_str_orig = nombre_archivo_orig[len(prefijo):len(prefijo)+8]
                        nuevo_nombre_archivo = nombre_archivo_orig.replace(fecha_str_orig, fecha_actual, 1)
                    else:
                        nuevo_nombre_archivo = nombre_archivo_orig

                    destino_final = f'{DESTINO_DIR}/{nuevo_nombre_archivo}'

                    # Subir al destino
                    with open(local_file_path, 'rb') as f:
                        ftp.storbinary('STOR ' + destino_final, f)
                        
                    # Eliminar original del FTP una vez procesado con éxito
                    try:
                        ftp.delete(archivo_encontrado)
                    except Exception as e:
                        LOGGER.error(f"No se pudo eliminar el original en FTP: {e}")

                    LOGGER.info(f"Archivo procesado correctamente | destino={destino_final}")

                except Exception as e:
                    LOGGER.exception(f"Error procesando archivo | error={e}")
                    if os.path.exists(local_file_path):
                        os.remove(local_file_path)
                    continue

    except ftplib.all_errors as e:
        LOGGER.exception(f"Error FTP general | error={e}")

def limpiar_logs_riorpos():
    LOGGER.info("INICIO limpiar_logs_riorpos")
    try:
        CARPETA_LOGS = '/integration/uc4/log/RIORPOS'
        with ftplib.FTP(FTP_HOST, timeout=10) as ftp:
            ftp.login(FTP_USER, FTP_PASS)
            LOGGER.info(f"Conectado al FTP para limpiar logs | host={FTP_HOST}")
            
            archivos = ftp.nlst(CARPETA_LOGS)
            logs_fecha = []
            
            for archivo in archivos:
                nombre = os.path.basename(archivo)
                if nombre == 'RIORPOS.log':
                    continue  # Mantener vivo el principal
                
                if nombre.startswith('RIORPOS.log.'):
                    logs_fecha.append(archivo)
            
            # Ordenar para encontrar el más reciente (YYYY-MM-DD permite orden alfabético correcto)
            logs_fecha.sort()
            
            if logs_fecha:
                mas_reciente = logs_fecha.pop()  # Quitamos el más reciente para no borrarlo
                LOGGER.info(f"Log más reciente preservado: {mas_reciente}")
                
            # Eliminar los restantes
            for log_borrar in logs_fecha:
                try:
                    ftp.delete(log_borrar)
                    LOGGER.info(f"Log antiguo eliminado del FTP: {log_borrar}")
                except Exception as e:
                    LOGGER.error(f"Error eliminando log {log_borrar} | error={e}")
                    
    except ftplib.all_errors as e:
        LOGGER.exception(f"Error FTP limpiando logs | error={e}")


def ejecutar_comando_as400(comando: str):
    LOGGER.info(f"INICIO ejecutar_comando_as400 | comando={comando}")
    try:
        db_manager = ConexionAS400()
        conexion = db_manager.conectar()
        if not conexion:
            LOGGER.error("No se pudo conectar a AS400 para ejecutar el comando.")
            return False
            
        cursor = conexion.cursor()
        # Escapar comillas simples para la llamada a QCMDEXC
        comando_escapado = comando.replace("'", "''")
        query = f"CALL QSYS2.QCMDEXC('{comando_escapado}')"
        
        cursor.execute(query)
        conexion.commit()
        
        cursor.close()
        db_manager.cerrar_conexion()
        LOGGER.info("Comando AS400 ejecutado exitosamente.")
        return True
    except Exception as e:
        LOGGER.error(f"Error ejecutando comando AS400: {e}")
        return False


if __name__ == '__main__':
    for pais in ['SV']:
        modificar_y_mover_ftp(pais, dias_atras=30)
        comando_qsh = f"SBMJOB CMD(QSH CMD('/integration/uc4/component/script/invDischarge.sh /integration/uc4/in/InventoryDischarge/{pais}')) JOB(INV{pais})"
        ejecutar_comando_as400(comando_qsh)
        
    limpiar_logs_riorpos()
    
    # Ejecutar comando en el AS400 utilizando SBMJOB
    