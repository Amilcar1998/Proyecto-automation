import logging
import subprocess
import sys
import os
import glob
import shutil
from datetime import datetime
import json

# Función para instalar paquetes si no están presentes
def install_package(package):
    try:
        __import__(package)
    except ImportError:
        logging.info(f"El paquete '{package}' no está instalado. Procediendo a instalarlo...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        logging.info(f"El paquete '{package}' se ha instalado correctamente.")

# Verificar e instalar paquetes necesarios
required_packages = ['pandas', 'xlwt', 'openpyxl']
for package in required_packages:
    install_package(package)

# Importar paquetes después de asegurar que estén instalados
import pandas as pd
import xlwt

# Definir las carpetas a borrar
carpetas_a_borrar = ["local", "Importado"]

# Borrar las carpetas especificadas si existen
for carpeta in carpetas_a_borrar:
    if os.path.exists(carpeta):
        shutil.rmtree(carpeta)
        logging.info(f"La carpeta '{carpeta}' ha sido borrada.")

# (Eliminada la creación de carpetas de backup por solicitud del usuario)

# Cargar el archivo de entrada (Excel)
df = pd.read_excel("Formatos/OC compras.xlsx")  # Cambia el nombre del archivo a tu archivo real

# Imprimir las columnas del DataFrame para verificar
logging.info("Columnas disponibles en el DataFrame:")
logging.info(df.columns.tolist())

# Borrar archivos existentes que inicien con "Plantilla"
for file in glob.glob("Plantillas/Plantilla*.xls"):  # Cambia a .xls
    os.remove(file)

# Lista para guardar los datos del reporte PO
po_report_data = []

# Agrupar los datos por 'ID', 'TIPO', 'PAIS' y 'PREDISTRIBUIDO'
for (id_valor, tipo_valor, pais_valor, predistribuido_valor), group in df.groupby(
        ['ID', 'TIPO', 'PAIS', 'PREDISTRIBUIDO']):

    # Determinar el directorio de salida basado en PAIS (LOCAL o IMPORTADA)
    if str(pais_valor).strip().upper() == 'LOCAL':
        directory = f'ARCHIVOS/LOCAL/'
    else:
        directory = f'ARCHIVOS/IMPORTADA/'

    # Crear el directorio si no existe
    os.makedirs(directory, exist_ok=True)

    # Crear un nuevo libro de trabajo
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Datos')

    # Escribir encabezados
    sheet.write(0, 0, 'PAIS')
    sheet.write(0, 1, 'COMPANIA')
    sheet.write(0, 2, 'TIENDA')
    sheet.write(0, 3, 'SKU')
    sheet.write(0, 4, 'COLOR')
    sheet.write(0, 5, 'TALLA')
    sheet.write(0, 6, 'MISELANEO')
    sheet.write(0, 7, 'UNIDADES')
    sheet.write(0, 8, 'PRECIO')

    # Contador de filas, comenzamos desde 1 para que los datos empiecen desde la fila 2
    fila = 1

    detalles = []
    # Escribir los datos a partir de la fila 2
    for _, row in group.iterrows():
        sheet.write(fila, 0, 4)  # PAIS
        sheet.write(fila, 1, 2)  # COMPANIA
        sheet.write(fila, 2, row['TIENDA'])  # TIENDA
        sheet.write(fila, 3, row['SKU'])  # SKU
        sheet.write(fila, 4, "")  # COLOR
        sheet.write(fila, 5, "")  # TALLA
        sheet.write(fila, 6, "")  # MISELANEO
        sheet.write(fila, 7, row['CANTIDADES'])  # UNIDADES
        sheet.write(fila, 8, "")  # PRECIO
        
        # Guardar en detalles para el JSON
        detalle = {
            "SKU": str(row['SKU']),
            "TIENDA": str(row['TIENDA']),
            "CANTIDADES": int(row['CANTIDADES']) if pd.notna(row['CANTIDADES']) else 0
        }
        if 'ITEMS' in row:
            detalle["ITEMS"] = int(row['ITEMS']) if pd.notna(row['ITEMS']) else 0
            
        detalles.append(detalle)

        # Aumentar la fila para el siguiente registro
        fila += 1

    # Nombre del archivo de salida basado en ID
    file_name = f'{directory}{id_valor}_Plantilla.xls'

    # Guardar el libro de trabajo
    workbook.save(file_name)

    # Confirmación de creación de archivo
    logging.info(f"Archivo {file_name} creado exitosamente.")


    # Agregar información al JSON
    po_report_data.append({
        'ID': str(id_valor),
        'TIPO': str(tipo_valor),
        'PAIS': str(pais_valor),
        'PREDISTRIBUIDO': str(predistribuido_valor),
        'Archivo_Generado': file_name,
        'Detalles': detalles,
        'Estado': 'Y'
    })


# Generar archivo JSON po_report.json con los datos recolectados
po_report_file = 'po_report.json'
with open(po_report_file, 'w', encoding='utf-8') as f:
    json.dump(po_report_data, f, indent=4, ensure_ascii=False)
logging.info(f"Reporte JSON guardado en {po_report_file}.")