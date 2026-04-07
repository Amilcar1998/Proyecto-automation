# -*- coding: utf-8 -*-
"""
Procesa todos los XML en una carpeta con patrón:
    InventoryDischarge_SV_YYYYMMDD*.xml

Para cada archivo:
  - Obtiene UPCs desde BD y reemplaza todos los <upc> del XML (ciclo).
  - Reemplaza SOLO el tag <transactionDate> por una fecha asignada
    en round-robin dentro del mes actual (del día 1 al día de hoy).
  - Renombra el archivo reemplazando la fecha de su nombre por la misma fecha asignada.
  - Guarda el resultado en una carpeta de salida, evitando colisiones de nombres.

Requisitos:
    - pyodbc
    - Conectividad al DSN configurado
"""

import os
import re
import sys
import logging
from pathlib import Path
from itertools import cycle
from datetime import date
import xml.etree.ElementTree as ET
import pyodbc

# ==============================
# CONFIGURACIÓN BD
# ==============================
# Sugerencia: mueve PASSWORD a variable de entorno y usa os.getenv("RI_PASSWORD")
DSN = "RI_TEST"
USUARIO = "ELOPEZ"
PASSWORD = "MAY2024"

SQL_UPC = """
SELECT DISTINCT TRIM(UPC_CORP) AS UPC
FROM ri11db.CONSULRP3 c
INNER JOIN ri11db.kskup b ON c.sku_corp = b.skusk
WHERE b.b34sk > 0 AND b.srlsk <> 'D'
"""

# ==============================
# CONFIGURACIÓN DE CARPETAS Y ARCHIVOS
# ==============================
# Carpeta de entrada y salida por defecto (puedes sobreescribirlas por CLI)
INPUT_DIR = "ArchivosEliseo enviar"
OUTPUT_DIR = "XML MODIFICADO"

# Patrón base esperado: InventoryDischarge_SV_YYYYMMDD + posible sufijo + .xml
FILENAME_PREFIX = "InventoryDischarge_SV_"
FILE_GLOB = f"{FILENAME_PREFIX}*.xml"

# Regex que captura: prefix, fecha (8 dígitos), sufijo opcional, extensión .xml
FILENAME_RE = re.compile(
    r"^(?P<prefix>InventoryDischarge_SV_)(?P<fecha>\d{8})(?P<sufijo>.*)\.xml$",
    re.IGNORECASE
)

# ==============================
# UTILIDADES
# ==============================
def localname(tag: str) -> str:
    """Devuelve el nombre local del tag, ignorando el namespace si existe."""
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag

# ==============================
# BD: OBTENER UPC DESDE BD
# ==============================
def obtener_upc_bd():
    conn = None
    try:
        conn = pyodbc.connect(f"DSN={DSN};UID={USUARIO};PWD={PASSWORD}")
        cursor = conn.cursor()
        cursor.execute(SQL_UPC)
        upcs = [str(row[0]).strip() for row in cursor.fetchall() if str(row[0]).strip()]
        if not upcs:
            raise RuntimeError("No se obtuvieron UPC desde la base de datos")
        return upcs
    finally:
        if conn:
            conn.close()

# ==============================
# FECHAS DEL MES ACTUAL (1 -> hoy)
# ==============================
def fechas_mes_actual_yyyymmdd():
    hoy = date.today()
    year, month = hoy.year, hoy.month
    dias = hoy.day  # del 1 al día actual
    return [f"{year:04d}{month:02d}{d:02d}" for d in range(1, dias + 1)]

# ==============================
# UTILIDADES DE NOMBRES DE ARCHIVOS
# ==============================
def parse_filename_parts(filename: str):
    m = FILENAME_RE.match(filename)
    if not m:
        return None
    return m.group("prefix"), m.group("fecha"), m.group("sufijo")

def build_output_filename(prefix: str, nueva_fecha: str, sufijo: str) -> str:
    return f"{prefix}{nueva_fecha}{sufijo}.xml"

# ==============================
# REEMPLAZOS EN XML
# ==============================
def reemplazar_upcs(root: ET.Element, lista_upc):
    """
    Reemplaza todos los nodos <upc> (ignora namespace).
    Devuelve la cantidad de nodos modificados.
    """
    upc_iter = cycle(lista_upc)
    count = 0
    for elem in root.iter():
        if localname(elem.tag).lower() == "upc":
            elem.text = next(upc_iter)
            count += 1
    return count

def reemplazar_transaction_date(root: ET.Element, nueva_fecha_yyyymmdd: str):
    """
    Reemplaza SOLO <transactionDate> (ignora namespace).
    Devuelve la cantidad de nodos modificados.
    """
    cambios = 0
    for elem in root.iter():
        if localname(elem.tag).lower() == "transactiondate":
            elem.text = nueva_fecha_yyyymmdd
            cambios += 1
    return cambios

def procesar_xml(xml_in_path: Path, xml_out_path: Path, upcs_bd, fecha_para_archivo: str):
    """
    Procesa un XML: reemplaza <upc> y <transactionDate>, guarda en xml_out_path.
    """
    tree = ET.parse(xml_in_path)
    root = tree.getroot()

    upc_count = reemplazar_upcs(root, upcs_bd)
    fecha_count = reemplazar_transaction_date(root, fecha_para_archivo)

    # Escribir XML
    xml_out_path.parent.mkdir(parents=True, exist_ok=True)
    tree.write(xml_out_path, encoding="utf-8", xml_declaration=True)
    return upc_count, fecha_count

# ==============================
# PIPELINE PRINCIPAL
# ==============================
def main(input_dir: str = INPUT_DIR, output_dir: str = OUTPUT_DIR):
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"La carpeta de entrada no existe: {in_dir.resolve()}")

    # Archivos candidatos
    files = sorted([p for p in in_dir.glob(FILE_GLOB) if p.is_file()])

    if not files:
        raise RuntimeError(
            f"No se encontraron archivos con patrón {FILE_GLOB} en {in_dir.resolve()}"
        )

    # Obtener UPCs de BD
    logging.info("Conectando a la BD para obtener UPCs...")
    upcs_bd = obtener_upc_bd()
    logging.info("UPCs obtenidos: %d", len(upcs_bd))

    # Fechas del mes actual (1 -> hoy)
    fechas = fechas_mes_actual_yyyymmdd()
    logging.info("Rango de fechas a repartir: %s -> %s (%d días)",
                 fechas[0], fechas[-1], len(fechas))

    total_ok = 0
    total_err = 0
    resumen = []

    for idx, in_path in enumerate(files):
        original_name = in_path.name
        parts = parse_filename_parts(original_name)
        if not parts:
            logging.warning("Nombre no coincide con el patrón esperado: %s", original_name)
            total_err += 1
            continue

        prefix, _fecha_en_nombre, sufijo = parts
        nueva_fecha = fechas[idx % len(fechas)]  # Reparto round-robin del 1 a hoy

        out_name = build_output_filename(prefix, nueva_fecha, sufijo)
        out_path = out_dir / out_name

        # Evitar colisiones si ya existe el nombre objetivo
        if out_path.exists():
            n = 1
            while True:
                candidato = out_dir / f"{out_path.stem}_{n:03d}{out_path.suffix}"
                if not candidato.exists():
                    out_path = candidato
                    break
                n += 1

        try:
            upc_count, fecha_count = procesar_xml(in_path, out_path, upcs_bd, nueva_fecha)
            total_ok += 1
            resumen.append({
                "in": original_name,
                "out": out_path.name,
                "fecha": nueva_fecha,
                "upcs_modificados": upc_count,
                "fechas_modificadas": fecha_count
            })
            logging.info("OK: %s -> %s | fecha=%s | upc=%d | transactionDate=%d",
                         original_name, out_path.name, nueva_fecha, upc_count, fecha_count)
        except Exception as ex:
            total_err += 1
            logging.error("ERROR procesando %s: %s", original_name, ex)

    logging.info("==== RESUMEN ====")
    logging.info("Procesados OK: %d | Con error: %d | Total: %d", total_ok, total_err, len(files))

    # Reporte corto al final (muestra hasta 10 filas)
    for r in resumen[:10]:
        logging.info("IN: %-40s OUT: %-40s FECHA: %s (UPC=%d, transactionDate=%d)",
                     r["in"], r["out"], r["fecha"], r["upcs_modificados"], r["fechas_modificadas"])

if __name__ == "__main__":
    # Uso:
    #   python script.py               -> usa INPUT_DIR/OUTPUT_DIR por defecto
    #   python script.py in_dir out_dir
    if len(sys.argv) >= 3:
        main(sys.argv[1], sys.argv[2])
    else:
        main()