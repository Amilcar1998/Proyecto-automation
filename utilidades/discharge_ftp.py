import ftplib
import os
import time
import datetime
import logging
from datetime import timedelta

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
    - Evita trailing spaces
    """
    raw = line.rstrip("\n")
    parts = raw.split("|")
    parts = [clean_field(p) for p in parts]
    return "|".join(parts).rstrip() + "\n"


def modificar_y_mover_ftp(pais, dias_atras=5):
    LOGGER.info(f"INICIO modificar_y_mover_ftp | pais={pais} dias_atras={dias_atras}")

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

            if not archivos_encontrados:
                LOGGER.warning(f"No se encontraron archivos dentro del rango | pais={pais}")
                return

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

                            # Regla original de tu HDR
                            if len(partes) > 3 and len(partes[3]) == 14 and partes[3].isdigit():
                                partes[3] = anio_mes_actual + partes[3][6:]

                            lineas_modificadas.append("|".join(partes).rstrip() + "\n")
                            continue

                        if linea.startswith('TRX|'):
                            raw = linea.rstrip('\n')
                            partes = raw.split('|')

                            # Limpia campos (convierte " " en "")
                            partes = [clean_field(p) for p in partes]

                            # Si por alguna razón viene con menos/más campos, valida estricto
                            # (No recortamos vacíos finales porque Java suele contar campos)
                            if len(partes) != 16:
                                archivo_invalido = True
                                trx_invalidas.append(num_linea)
                                LOGGER.error(f"TRX inválida | linea={num_linea} campos={len(partes)} raw={repr(raw)}")
                                break

                            # Regla original: actualizar fechas YYYYMMDD en posiciones 7..15
                            for idx in range(7, 16):
                                if len(partes[idx]) == 8 and partes[idx].isdigit():
                                    partes[idx] = anio_mes_actual + partes[idx][6:]

                            lineas_modificadas.append("|".join(partes).rstrip() + "\n")
                            continue

                        # Otras líneas (INV, etc.): normaliza en general para quitar trailing y campos " "
                        lineas_modificadas.append(normalize_pipe_line(linea))

                    if archivo_invalido:
                        LOGGER.error(f"Archivo inválido será eliminado | archivo={archivo_encontrado} trx_invalidas={trx_invalidas}")

                        try:
                            ftp.delete(archivo_encontrado)
                            LOGGER.warning(f"Eliminado del FTP | archivo={archivo_encontrado}")
                        except Exception as e:
                            LOGGER.exception(f"Error eliminando del FTP | error={e}")

                        try:
                            os.remove(local_file_path)
                            LOGGER.warning(f"Eliminado local | archivo={local_file_path}")
                        except Exception as e:
                            LOGGER.exception(f"Error eliminando local | error={e}")

                        continue

                    # Guardar archivo modificado
                    with open(local_file_path, 'w', encoding='utf-8', newline='\n') as f:
                        f.writelines(lineas_modificadas)

                    destino_final = f'{DESTINO_DIR}/{os.path.basename(local_file_path)}'

                    # Subir al destino
                    with open(local_file_path, 'rb') as f:
                        ftp.storbinary('STOR ' + destino_final, f)

                    LOGGER.info(f"Archivo procesado correctamente | destino={destino_final}")

                except Exception as e:
                    LOGGER.exception(f"Error procesando archivo | error={e}")
                    if os.path.exists(local_file_path):
                        os.remove(local_file_path)
                    continue

    except ftplib.all_errors as e:
        LOGGER.exception(f"Error FTP general | error={e}")


if __name__ == '__main__':
    for pais in ['SV']:
        modificar_y_mover_ftp(pais, dias_atras=20)