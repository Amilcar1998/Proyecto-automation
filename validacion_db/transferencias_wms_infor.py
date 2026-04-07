import os
import sys
# Añadir el directorio raíz al path de Python para permitir importaciones relativas
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import time
import pandas as pd
from datetime import datetime
import win32com.client as win32
import logging
from conexion_config.logging_config import get_logger
import traceback
import json
import re
import requests
import docx
# Importamos los módulos necesarios
from conexion_config.conexion import ConexionAS400
from scripts_as400.ejecutor_cl import EjecutorCL
from Jira_Utilidades.jira_utilidades import JiraClient
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT
from requests.auth import HTTPBasicAuth

logger = get_logger(__name__)

def esperar_finalizacion_trabajo(base,numero_orden,conexion,logger,tabla_origen,tabla_destino,job_name,timeout_minutos):
    """
    Espera a que un proceso batch finalice validando movimientos
    entre tablas dinámicas (multi–base / multi–tabla).
    """

    inicio_monitoreo = time.time()
    timeout_segundos = timeout_minutos * 60
    intervalo_verificacion = 5

    logger.info(
        f"Iniciando monitoreo | Orden={numero_orden} | "
        f"Origen={tabla_origen} → Destino={tabla_destino}"
    )

    try:
        logger.info("Obteniendo el conteo de la tabla origen")
        conteo_inicial_origen = obtener_conteo_registros(base,
            conexion, tabla_origen, numero_orden
        )
        logger.info("Obteniendo el conteo de la tabla Destino")
        conteo_inicial_destino = obtener_conteo_registros(base,
            conexion, tabla_destino, numero_orden
        )

        logger.info(
            f"Estado inicial - "
            f"{tabla_origen}: {conteo_inicial_origen}, "
            f"{tabla_destino}: {conteo_inicial_destino}"
        )

        if conteo_inicial_origen == -1 or conteo_inicial_destino == -1:
            logger.warning("Error en conteo inicial. Fallback 30s.")
            time.sleep(30)
            return True

        if conteo_inicial_origen == 0:
            logger.info("Tabla origen ya vacía. Proceso completado.")
            return True

        ultimo_conteo_origen = conteo_inicial_origen
        tiempo_sin_cambios = 0

        while time.time() - inicio_monitoreo < timeout_segundos:
            time.sleep(intervalo_verificacion)

            conteo_actual_origen = obtener_conteo_registros(base,
                conexion, tabla_origen, numero_orden
            )
            conteo_actual_destino = obtener_conteo_registros(base,
                conexion, tabla_destino, numero_orden
            )

            logger.info(
                f"Verificación - "
                f"{tabla_origen}: {conteo_actual_origen}, "
                f"{tabla_destino}: {conteo_actual_destino}"
            )

            # Finalización correcta
            if conteo_actual_origen == 0 and conteo_actual_destino > conteo_inicial_destino:
                logger.info("[OK] Registros procesados completamente.")
                return True

            # Progreso detectado
            if conteo_actual_origen < ultimo_conteo_origen:
                logger.info(
                    f"[PROGRESO] {ultimo_conteo_origen} → {conteo_actual_origen}"
                )
                ultimo_conteo_origen = conteo_actual_origen
                tiempo_sin_cambios = 0
            else:
                tiempo_sin_cambios += intervalo_verificacion

            # Validación de job si se proporciona
            if tiempo_sin_cambios >= 60 and job_name:
                logger.info(f"[VERIFY] Verificando job activo: {job_name}")
                try:
                    activo = verificar_trabajos_activos(conexion, job_name)
                    if activo is False:
                        logger.warning("No hay job activo y no hay progreso.")
                        return False
                except Exception as e:
                    logger.debug(f"No se pudo validar job: {e}")

                tiempo_sin_cambios = 0

        # Timeout
        logger.warning(f"⏰ Timeout alcanzado ({timeout_minutos} min)")
        return False

    except Exception as e:
        logger.error(f"Error en monitoreo: {e}")
        time.sleep(30)
        return True


def obtener_conteo_registros(base,conexion, tabla, numero_orden):
    """
    Obtiene el conteo de registros en una tabla para una orden específica
    """
    try:
        query = f"SELECT COUNT(*) as CONTEO FROM {base}.{tabla} WHERE NUMORDEN = '{numero_orden}'"
        cursor = conexion.cursor()
        cursor.execute(query)
        resultado = cursor.fetchone()
        return resultado[0] if resultado else 0
    except Exception as e:
        # logger.info(f"Error al obtener conteo de {tabla}: {e}") # Comentado para evitar flood en consola
        return -1

def verificar_trabajos_activos(conexion, nombre_trabajo_patron):
    """
    Verifica si hay trabajos activos que coincidan con el patrón especificado
    """
    try:
        # Consultar tabla de trabajos activos del sistema
        query = f"""
            SELECT COUNT(*) as CONTEO 
            FROM TABLE(QSYS2.ACTIVE_JOB_INFO()) 
            WHERE JOB_NAME LIKE '%{nombre_trabajo_patron}%'
            AND JOB_STATUS = 'ACTIVE'
        """
        cursor = conexion.cursor()
        cursor.execute(query)
        resultado = cursor.fetchone()
        return (resultado[0] > 0) if resultado else False
    except Exception as e:
        # logger.info(f"Error al verificar trabajos activos: {e}") # Comentado para evitar flood en consola
        return None

def ejecutar_query(conexion, query):
    """Ejecuta una consulta SQL y devuelve los resultados"""
    try:
        cursor = conexion.cursor()
        cursor.execute(query)
        
        # Si hay resultados, convertir a DataFrame
        if cursor.description:
            columns = [column[0] for column in cursor.description]
            results = cursor.fetchall()
            df = pd.DataFrame.from_records(results, columns=columns)
            return df
        else:
            return None
    except Exception as e:
        logger.info(f"Error al ejecutar consulta: {e}")
        return None

def ejecutar_stored_procedure(conexion, procedure_call):
    """Ejecuta un stored procedure"""
    try:
        cursor = conexion.cursor()
        cursor.execute(procedure_call)
        conexion.commit()
        return True
    except Exception as e:
        logger.info(f"Error al ejecutar stored procedure: {e}")
        return False

def ejecutar_comando_cl(comando, submit_job=False, job_name=None):
    """
    Ejecuta un comando CL en AS400
    """
    try:
        cl = EjecutorCL()
        
        if submit_job:
            respuesta = cl.submit_job(comando, job_name=job_name)
        else:
            respuesta = cl.ejecutar(comando)
            
        return respuesta
    except Exception as e:
        logger.info(f"Error al ejecutar comando CL: {e}")
        return None

# --- Funciones de Reporte y Correo ---

def generar_excel(df, filename):
    """Genera un archivo Excel a partir de un DataFrame (Función no usada en el flujo principal)"""
    try:
        directorio = os.path.dirname(filename)
        if not os.path.exists(directorio):
            os.makedirs(directorio)
            
        df.to_excel(filename, index=False)
        logger.info(f"Excel generado: {filename}")
        return filename
    except Exception as e:
        logger.info(f"Error al generar Excel: {e}")
        return None

def get_default_signature():
    """Devuelve la firma HTML por defecto cuando no se encuentra la imagen"""
    return """
    <hr style='border: 1px solid #dddddd;'>
    <div style='font-size: 11px; color: #777777;'>
        <p>Generado automáticamente por el Sistema de Automatización de Transferencias WMS-INFOR</p>
    </div>
    """

def enviar_correo(asunto, cuerpo, adjuntos=None):
    """Envía un correo con los archivos adjuntos de transferencias WMS-INFOR"""
    try:
        logger.info("Iniciando proceso de envío de correo...")
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        # Leer destinatarios desde archivo de configuración
        destinatarios = []
        config_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'archivos_config')
        destinatarios_path = os.path.join(config_dir, 'destinatarios.txt')
        
        logger.info(f"Buscando archivo de destinatarios en: {destinatarios_path}")
        if os.path.exists(destinatarios_path):
            with open(destinatarios_path, 'r', encoding='utf-8') as f:
                for line in f:
                    for correo in line.replace(';', ',').split(','):
                        correo = correo.strip()
                        if correo and not correo.startswith('#'):
                            destinatarios.append(correo)
            logger.info(f"Destinatarios encontrados en archivo: {destinatarios}")
        else:
            logger.info("Archivo de destinatarios no encontrado, usando destinatario por defecto")
            destinatarios = ["eliseo_lopezp@unicomer.com"] # Destinatario por defecto
        
        mail.To = ";".join(destinatarios)
        mail.Subject = asunto
        
        # Obtener ruta de la firma
        recursos_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Recursos')
        os.makedirs(recursos_dir, exist_ok=True)
        firma_path = os.path.join(recursos_dir, 'Firma.jpg')
        
        #slogger.info(f"Buscando firma en: {firma_path}")
        
        # Incluir la firma directamente en el cuerpo HTML del correo
        if os.path.exists(firma_path):
            import base64
            with open(firma_path, 'rb') as img_file:
                img_data = base64.b64encode(img_file.read()).decode('utf-8')
            
            firma_html = f"""
            <hr style='border: 1px solid #dddddd; margin-top: 20px;'>
            <div style='font-size: 11px; color: #777777;'>
                <img src="data:image/jpeg;base64,{img_data}" alt="Firma" style="max-width: 500px;" />
            </div>
            """
            logger.info("Firma convertida a base64 e incrustada en el HTML")
        else:
            logger.info(f"Firma no encontrada en: {firma_path}")
            firma_html = get_default_signature()
        
        # Combinar el cuerpo con la firma
        mail.HTMLBody = cuerpo + firma_html
        
        # Verificar y adjuntar archivos
        archivos_adjuntados = 0
        if adjuntos:
            for adjunto in adjuntos:
                if os.path.exists(adjunto):
                    mail.Attachments.Add(adjunto)
                    archivos_adjuntados += 1
                    logger.info(f"Adjunto agregado: {adjunto}")
                else:
                    logger.info(f"ADVERTENCIA: El archivo adjunto no existe: {adjunto}")
        
        logger.info(f"Enviando correo a {len(destinatarios)} destinatarios con {archivos_adjuntados} adjuntos...")
        mail.Send()
        logger.info(f"[OK] Correo enviado exitosamente a: {', '.join(destinatarios)}")
        return True
    except Exception as e:
        logger.info(f"[ERROR] ERROR al enviar correo: {e}")
        logger.info(traceback.format_exc())
        return False

# --- Funciones de Generación de Reporte Word ---

def dataframe_to_table(document, dataframe, titulo):
    """Convierte un DataFrame a una tabla en un documento Word para reportes de transferencias WMS-INFOR"""
    if dataframe is None or dataframe.empty:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(f"No hay datos disponibles para {titulo}")
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(169, 68, 66)  # Rojo oscuro
        return
    
    # Añadir título de sección con formato compacto
    heading = document.add_heading(titulo, level=2)
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    heading.paragraph_format.space_before = Pt(6)
    heading.paragraph_format.space_after = Pt(3)
    
    # Crear tabla
    table = document.add_table(rows=1, cols=len(dataframe.columns))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Encabezados
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(dataframe.columns):
        hdr_cells[i].text = str(column)
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        # Centrar texto de encabezado
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Añadir datos
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value) if value is not None else ''

# --- Funciones de Validación de Inventario ---

def validar_inventario(df_inv_inicial, df_inv_final, df_orden, logger):
    """Valida inventario con formato de tabla mostrando validaciones de origen y destino."""
    resultados = {
        'estado': 'OK',
        'mensaje': '',
        'resumen_inv_inicial': None,
        'metricas': {},
        'registros_inicial': int(0 if df_inv_inicial is None else len(df_inv_inicial)),
        'registros_final': int(0 if df_inv_final is None else len(df_inv_final)),
        'registros_winrc': int(0 if df_orden is None else len(df_orden)),
        'movimientos_inventario': [],
        'discrepancias': []
    }

    if df_inv_inicial is None or df_inv_final is None or df_orden is None:
        logger.error("Datos insuficientes para validación de inventario")
        resultados['estado'] = 'ERROR'
        resultados['mensaje'] = 'Faltan uno o más DataFrames requeridos.'
        return resultados

    try:
        tmp_inicial = df_inv_inicial.copy()
        tmp_final = df_inv_final.copy()
        
        # Normalizaciones
        for df in [tmp_inicial, tmp_final]:
            df['SKU_N'] = df['SKU'].astype(str).str.strip()
            df['ORIGEN_N'] = df['ORIGEN'].astype(str).str.strip()
            df['DESTINO_N'] = df['DESTINO'].astype(str).str.strip()
            
            for col in ['UNIDADES', 'INV_ORIGEN', 'INV_DESTINO']:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        logger.info(' VALIDACIÓN DE TRANSFERENCIAS CON TABLA')
        logger.info('='*100)
        
        # VALIDACIÓN DE ORÍGENES: INV_INICIAL_ORIGEN - TOTAL_TRANSFERIDO = INV_FINAL_ORIGEN
        df_origenes = (
            tmp_inicial.groupby(['SKU_N', 'ORIGEN_N'], as_index=False)
            .agg({
                'UNIDADES': 'sum',
                'INV_ORIGEN': 'first'
            })
            .rename(columns={
                'SKU_N': 'SKU', 
                'ORIGEN_N': 'ORIGEN',
                'UNIDADES': 'TRANSFER',
                'INV_ORIGEN': 'INV_INI'
            })
        )
        
        df_final_origen = (
            tmp_final.groupby(['SKU_N', 'ORIGEN_N'], as_index=False)
            .agg({'INV_ORIGEN': 'first'})
            .rename(columns={'SKU_N': 'SKU', 'ORIGEN_N': 'ORIGEN', 'INV_ORIGEN': 'INV_FIN'})
        )
        
        df_validacion_origen = df_origenes.merge(df_final_origen, on=['SKU', 'ORIGEN'], how='left')
        df_validacion_origen['CALC'] = df_validacion_origen['INV_INI'] - df_validacion_origen['TRANSFER']
        df_validacion_origen['VALIDO'] = df_validacion_origen['CALC'] == df_validacion_origen['INV_FIN']
        df_validacion_origen['STATUS'] = df_validacion_origen['VALIDO'].apply(lambda x: '[OK] OK' if x else '[ERROR] ERROR')
        
        # Convertir a enteros para mostrar
        for col in ['INV_INI', 'TRANSFER', 'CALC', 'INV_FIN']:
            df_validacion_origen[col] = df_validacion_origen[col].astype(int)
        
        logger.info('TABLA DE VALIDACIÓN - ORÍGENES')
        logger.info('='*80)
        tabla_origen = df_validacion_origen[['SKU', 'ORIGEN', 'INV_INI', 'TRANSFER', 'CALC', 'INV_FIN', 'STATUS']]
        
        
        # VALIDACIÓN DE DESTINOS: INV_INICIAL_DESTINO + TOTAL_RECIBIDO = INV_FINAL_DESTINO
        df_destinos = (
            tmp_inicial.groupby(['SKU_N', 'DESTINO_N'], as_index=False)
            .agg({
                'UNIDADES': 'sum',
                'INV_DESTINO': 'first'
            })
            .rename(columns={
                'SKU_N': 'SKU', 
                'DESTINO_N': 'DESTINO',
                'UNIDADES': 'RECIBIDO',
                'INV_DESTINO': 'INV_INI'
            })
        )
        
        df_final_destino = (
            tmp_final.groupby(['SKU_N', 'DESTINO_N'], as_index=False)
            .agg({'INV_DESTINO': 'first'})
            .rename(columns={'SKU_N': 'SKU', 'DESTINO_N': 'DESTINO', 'INV_DESTINO': 'INV_FIN'})
        )
        
        df_validacion_destino = df_destinos.merge(df_final_destino, on=['SKU', 'DESTINO'], how='left')
        df_validacion_destino['CALC'] = df_validacion_destino['INV_INI'] + df_validacion_destino['RECIBIDO']
        df_validacion_destino['VALIDO'] = df_validacion_destino['CALC'] == df_validacion_destino['INV_FIN']
        df_validacion_destino['STATUS'] = df_validacion_destino['VALIDO'].apply(lambda x: '[OK] OK' if x else '[ERROR] ERROR')
        
        # Convertir a enteros para mostrar
        for col in ['INV_INI', 'RECIBIDO', 'CALC', 'INV_FIN']:
            df_validacion_destino[col] = df_validacion_destino[col].astype(int)
        
        logger.info(' TABLA DE VALIDACIÓN - DESTINOS')
        logger.info('='*80)
        tabla_destino = df_validacion_destino[['SKU', 'DESTINO', 'INV_INI', 'RECIBIDO', 'CALC', 'INV_FIN', 'STATUS']]
        
        
        # RESUMEN FINAL
        total_origenes = len(df_validacion_origen)
        origenes_correctos = df_validacion_origen['VALIDO'].sum()
        total_destinos = len(df_validacion_destino)
        destinos_correctos = df_validacion_destino['VALIDO'].sum()
        
        logger.info('[PROGRESO] RESUMEN FINAL')
        logger.info('='*50)
        resumen_data = {
            'TIPO': ['ORÍGENES', 'DESTINOS', 'TOTAL'],
            'CORRECTOS': [origenes_correctos, destinos_correctos, origenes_correctos + destinos_correctos],
            'TOTAL': [total_origenes, total_destinos, total_origenes + total_destinos],
            'PORCENTAJE': [
                f'{(origenes_correctos/total_origenes*100):.0f}%' if total_origenes > 0 else '0%',
                f'{(destinos_correctos/total_destinos*100):.0f}%' if total_destinos > 0 else '0%',
                f'{((origenes_correctos + destinos_correctos)/(total_origenes + total_destinos)*100):.0f}%' if (total_origenes + total_destinos) > 0 else '0%'
            ],
            'STATUS': [
                '[OK] OK' if origenes_correctos == total_origenes else '[ERROR] ERROR',
                '[OK] OK' if destinos_correctos == total_destinos else '[ERROR] ERROR',
                '[OK] OK' if (origenes_correctos == total_origenes and destinos_correctos == total_destinos) else '[ERROR] ERROR'
            ]
        }
        
        df_resumen = pd.DataFrame(resumen_data)
        
        
        # Determinar estado general
        if origenes_correctos == total_origenes and destinos_correctos == total_destinos:
            logger.info(' RESULTADO: TODAS LAS TRANSFERENCIAS SON CORRECTAS')
            resultados['estado'] = 'OK'
            resultados['mensaje'] = 'Todas las validaciones de origen y destino son correctas'
        else:
            logger.info('[ERROR] RESULTADO: HAY ERRORES EN LAS TRANSFERENCIAS')
            resultados['estado'] = 'ERROR'
            resultados['mensaje'] = f'Errores encontrados - Orígenes: {origenes_correctos}/{total_origenes}, Destinos: {destinos_correctos}/{total_destinos}'
        
        # Guardar las tablas de validación para incluir en el reporte Word
        resultados['tablas_validacion'] = {
            'origen': df_validacion_origen[['SKU', 'ORIGEN', 'INV_INI', 'TRANSFER', 'CALC', 'INV_FIN', 'STATUS']],
            'destino': df_validacion_destino[['SKU', 'DESTINO', 'INV_INI', 'RECIBIDO', 'CALC', 'INV_FIN', 'STATUS']],
            'resumen': df_resumen
        }
        
        # Crear resumen para el reporte Word (formato compatible)
        df_resumen_origen_reporte = df_validacion_origen.copy()
        df_resumen_origen_reporte = df_resumen_origen_reporte.rename(columns={
            'INV_INI': 'INV_INICIAL_ORIGEN',
            'INV_FIN': 'INV_FINAL_ORIGEN',
            'CALC': 'DIFERENCIA_ORIGEN',
            'TRANSFER': 'TOTAL_TRANSFERIDO'
        })
        df_resumen_origen_reporte['DIFERENCIA_ORIGEN'] = df_resumen_origen_reporte['INV_FINAL_ORIGEN'] - df_resumen_origen_reporte['INV_INICIAL_ORIGEN']
        
        # Agregar columnas de destinos para compatibilidad (sin texto innecesario)
        df_resumen_origen_reporte['TOTAL_DESTINOS'] = 1  # Simplificado
        df_resumen_origen_reporte['DESTINOS'] = ''  # Campo vacío, no "Ver detalle..."
        
        resultados['resumen_inv_inicial'] = df_resumen_origen_reporte[['SKU', 'ORIGEN', 'INV_INICIAL_ORIGEN', 'INV_FINAL_ORIGEN', 'DIFERENCIA_ORIGEN', 'TOTAL_TRANSFERIDO', 'TOTAL_DESTINOS', 'DESTINOS']]
        resultados['metricas']['grupos_inicial'] = int(len(df_resumen_origen_reporte))
        resultados['metricas']['filas_inicial'] = int(len(tmp_inicial))
        
        # Crear movimientos de inventario para el reporte Word
        for _, row in df_validacion_origen.iterrows():
            movimiento = {
                'sku': str(row['SKU']),
                'origen': str(row['ORIGEN']),
                'destino': '',  # Campo vacío, no "Ver tabla destinos"
                'unidades_movidas': int(row['TRANSFER']),
                'inv_origen': int(row['INV_INI']),
                'inv_destino': 0,  # No aplicable en vista de origen
                'origen_ok': bool(row['VALIDO']),
                'destino_ok': True,  # Asumir OK por defecto
                'estado_general': 'OK' if row['VALIDO'] else 'ERROR',
                'cantidad_lineas': 1,
                'total_unidades_sku': int(row['TRANSFER'])
            }
            resultados['movimientos_inventario'].append(movimiento)
        
        logger.info(f"Validación completada: {len(df_validacion_origen)} orígenes y {len(df_validacion_destino)} destinos validados")
        
    except Exception as e:
        logger.error(f"Error en validación: {e}")
        logger.error(traceback.format_exc())
        resultados['estado'] = 'ERROR'
        resultados['mensaje'] = f'Error en validación: {str(e)}'

    return resultados


def mostrar_resumen_10_lineas(resultados_validacion):
    """Muestra un resumen de las 10 líneas con los campos específicos solicitados"""
    if not resultados_validacion.get('movimientos_inventario'):
        logger.info("[VERIFY] No hay movimientos de inventario para mostrar")
        return
    
    logger.info("" + "="*140)
    logger.info("[PROGRESO] RESUMEN DE VALIDACIÓN - 10 LÍNEAS")
    logger.info("="*140)
    # SKU, ORIGEN, DESTINO, INV_ORIGEN, INV_DESTINO, CANT_LINEAS, TOTAL_SKU, RESULTADO
    #logger.info(f"{'SKU':<14} {'ORIGEN':<8} {'DESTINO':<8} {'INV_ORIGEN':<12} {'INV_DESTINO':<12} {'CANT_LINEAS':<12} {'TOTAL_SKU':<10} {'RESULTADO':<10}")
    logger.info("-"*140)
    
    # Ordenar por SKU, ORIGEN, DESTINO para visualización consistente
    movimientos = sorted(
        resultados_validacion['movimientos_inventario'],
        key=lambda m: (str(m.get('sku','')), str(m.get('origen','')), str(m.get('destino','')))
    )
    
    # Solo mostrar los movimientos únicos de la tabla de validación (SKU + ORIGEN)
    movimientos_unicos = {}
    for mov in movimientos:
        clave = f"{mov['sku']}_{mov['origen']}"
        if clave not in movimientos_unicos:
            movimientos_unicos[clave] = mov

    for i, mov in enumerate(movimientos_unicos.values(), 1):
        sku = str(mov.get('sku',''))[:14]
        origen = str(mov.get('origen',''))[:8]
        # El destino en esta tabla es el destino principal del DF_ORDEN, se deja vacío o se busca el destino más común
        destino = 'N/A' # Mejor dejar N/A o buscar el destino del movimiento.
        
        # Para simplificar el resumen en consola, tomamos el destino de la primera línea de la orden
        if 'destino' in mov:
            destino = str(mov['destino'])[:8]
        elif 'tablas_validacion' in resultados_validacion and 'destino' in resultados_validacion['tablas_validacion']:
            df_dest = resultados_validacion['tablas_validacion']['destino']
            destino_match = df_dest[df_dest['SKU'] == mov['sku']]['DESTINO'].iloc[0] if not df_dest[df_dest['SKU'] == mov['sku']].empty else 'N/A'
            destino = str(destino_match)[:8]
        
        inv_origen = f"{mov.get('inv_origen', 0):.0f}" # Convertido a entero para visualización
        inv_destino = f"{mov.get('inv_destino', 0):.0f}" # Convertido a entero para visualización
        cant_lineas = str(mov.get('cantidad_lineas', 1))
        total_sku = f"{mov.get('total_unidades_sku', 0):.0f}"
        resultado = mov.get('estado_general', 'N/A')
        
        # Formatear resultado con emoji
        if resultado == 'OK':
            resultado_fmt = "[OK] OK"
        else:
            resultado_fmt = "[ERROR] ERROR"
        
        logger.info(f"{sku:<14} {origen:<8} {destino:<8} {inv_origen:<12} {inv_destino:<12} {cant_lineas:<12} {total_sku:<10} {resultado_fmt:<10}")
    
    logger.info("="*140)
    logger.info(f"Total de líneas procesadas (agrupadas por SKU+ORIGEN): {len(movimientos_unicos)}")
    logger.info("NOTA: INV_ORIGEN/DESTINO son saldos iniciales. TOTAL_SKU es el total transferido para ese SKU+ORIGEN.")
    


def generar_word_reporte_con_validacion(df_orden, df_inv_inicial, df_ktchpwinrc, df_inv_final, resultados_validacion, numero_orden, reporte_path):
    """Genera un documento Word con los datos de las transferencias WMS-INFOR y resultados de validación"""
    try:
        document = Document()
        
        # Configurar tamaño de página A3 (para mejor visualización de tablas)
        section = document.sections[0]
        section.page_width = Inches(16.5)
        section.page_height = Inches(11.7)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        
        # Configurar estilo del documento para eliminar espacios entre párrafos
        for style_name in document.styles:
            try:
                style = document.styles[style_name]
                if hasattr(style, 'paragraph_format'):
                    style.paragraph_format.space_before = Pt(0)
                    style.paragraph_format.space_after = Pt(0)
            except:
                pass
        
        # Configurar estilo de texto base
        style = document.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)
        
        # Título principal
        title = document.add_heading(f'Reporte de Transferencia WMS-INFOR - Orden {numero_orden}', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Información general - en la misma línea
        info_p = document.add_paragraph()
        info_p.add_run(f'Fecha de generación: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")} | ')
        info_p.add_run(f'Número de orden: {numero_orden}')
        
        # Sección 1: Datos de la orden
        dataframe_to_table(document, df_orden, "Datos de la Transferencia (KTCHPWASIN)")
        
        # Sección 2: Inventario Inicial
        dataframe_to_table(document, df_inv_inicial, "Inventario Inicial (INVENTARIO_INICIAL)")
        
        # Sección 3: Datos KTCHPWINRC
        dataframe_to_table(document, df_ktchpwinrc, "Datos KTCHPWINRC (KTCHPWINRC)")
        
        # Sección 4: Inventario Final
        dataframe_to_table(document, df_inv_final, "Inventario Final (INVENTARIO_FINAL)")
        
        # Sección 5: Resultados de la Validación
        document.add_heading('Resultados de la Validación', level=1)
        
        # Estado y mensaje de validación en un solo párrafo
        p = document.add_paragraph()
        p.add_run(f'Estado: ').bold = True
        estado_run = p.add_run(resultados_validacion['estado'])
        estado_run.bold = True
        
        if resultados_validacion['estado'] == 'OK':
            estado_run.font.color.rgb = RGBColor(0, 128, 0)  # Verde
        elif resultados_validacion['estado'] == 'ADVERTENCIA':
            estado_run.font.color.rgb = RGBColor(255, 165, 0)  # Naranja
        else:
            estado_run.font.color.rgb = RGBColor(255, 0, 0)  # Rojo
            
        p.add_run(' | ')
        p.add_run(f'Mensaje: ').bold = True
        p.add_run(resultados_validacion['mensaje'])
        
        # TABLAS DE VALIDACIÓN - Exactamente como aparecen en el log
        if 'tablas_validacion' in resultados_validacion:
            # Agregar tabla de validación de ORÍGENES
            if 'origen' in resultados_validacion['tablas_validacion']:
                df_validacion_origen = resultados_validacion['tablas_validacion']['origen']
                dataframe_to_table(document, df_validacion_origen, "TABLA DE VALIDACIÓN - ORÍGENES")
            
            # Agregar tabla de validación de DESTINOS
            if 'destino' in resultados_validacion['tablas_validacion']:
                df_validacion_destino = resultados_validacion['tablas_validacion']['destino']
                dataframe_to_table(document, df_validacion_destino, "TABLA DE VALIDACIÓN - DESTINOS")
            
            # Agregar resumen final
            if 'resumen' in resultados_validacion['tablas_validacion']:
                df_resumen = resultados_validacion['tablas_validacion']['resumen']
                dataframe_to_table(document, df_resumen, "RESUMEN FINAL")

        # TABLA: Resumen de Inventario Inicial Agrupado (SKU, ORIGEN)
        try:
            resumen_df = resultados_validacion.get('resumen_inv_inicial')
            if isinstance(resumen_df, pd.DataFrame) and not resumen_df.empty:
                resumen_ordenado = resumen_df.copy()
                if 'SKU' in resumen_ordenado.columns and 'ORIGEN' in resumen_ordenado.columns:
                    resumen_ordenado = resumen_ordenado.sort_values(['SKU', 'ORIGEN'])
                dataframe_to_table(document, resumen_ordenado, "Resumen Inventario Inicial (SKU, ORIGEN)")
        except Exception:
            pass

        # TABLA: Resumen de Movimientos por SKU
        if resultados_validacion.get('movimientos_inventario'):
            document.add_heading('Resumen de Movimientos por SKU', level=2)
            
            # Explicación del resumen con aclaración de fuente de datos
            p_resumen_intro = document.add_paragraph()
            p_resumen_intro.add_run("Detalle de movimientos reales de inventario por SKU (solo movimientos > 0, sin duplicados):").bold = True
            
            # Nota explicativa sobre los datos
            p_nota = document.add_paragraph()
            p_nota.add_run("NOTA: ").bold = True
            nota_inv_run = p_nota.add_run("INV_ORIGEN")
            nota_inv_run.bold = True
            nota_inv_run.font.color.rgb = RGBColor(128, 0, 128)  # Púrpura
            p_nota.add_run(" e ")
            nota_dest_run = p_nota.add_run("INV_DESTINO")
            nota_dest_run.bold = True
            nota_dest_run.font.color.rgb = RGBColor(0, 100, 139)  # Azul
            p_nota.add_run(" muestran los valores del inventario INICIAL antes de la transferencia. ")
            p_nota.add_run("TOTAL_SKU")
            p_nota.runs[-1].bold = True
            p_nota.runs[-1].font.color.rgb = RGBColor(128, 128, 0)  # Oliva
            p_nota.add_run(" es el total de unidades movidas para ese SKU en todas las transferencias.")
            
            # Agregar información específica sobre origen y destino
            if resultados_validacion.get('movimientos_inventario'):
                origenes = set()
                destinos = set()
                for mov in resultados_validacion['movimientos_inventario']:
                    if mov.get('unidades_movidas', 0) > 0:
                        origenes.add(mov['origen'])
                        destinos.add(mov['destino'])
                
                if origenes or destinos:
                    p_ubicaciones = document.add_paragraph()
                    p_ubicaciones.add_run("UBICACIONES: ").bold = True
                    
                    if origenes:
                        origen_run = p_ubicaciones.add_run("ORIGEN ")
                        origen_run.bold = True
                        origen_run.font.color.rgb = RGBColor(178, 34, 34)  # Rojo ladrillo
                        origenes_str = ", ".join(str(o) for o in sorted(origenes))
                        p_ubicaciones.add_run(f"({origenes_str})")
                    
                    if origenes and destinos:
                        p_ubicaciones.add_run(" → ")
                    
                    if destinos:
                        destino_run = p_ubicaciones.add_run("DESTINOS ")
                        destino_run.bold = True
                        destino_run.font.color.rgb = RGBColor(0, 100, 0)  # Verde oscuro
                        destinos_str = ", ".join(str(d) for d in sorted(destinos))
                        p_ubicaciones.add_run(f"({destinos_str})")
            
            # Crear tabla con las mismas columnas que el resumen de consola
            table = document.add_table(rows=1, cols=8)
            table.style = 'Table Grid'
            
            # Encabezados
            hdr_cells = table.rows[0].cells
            headers = ["SKU", "ORIGEN", "DESTINO", "INV_ORIGEN", "INV_DESTINO", "UNID_MV", "TOTAL_SKU", "RESULTADO"]
            colores = [
                RGBColor(0, 0, 139), RGBColor(178, 34, 34), RGBColor(0, 100, 0), RGBColor(128, 0, 128),
                RGBColor(0, 100, 139), RGBColor(255, 140, 0), RGBColor(128, 128, 0), RGBColor(0, 128, 0)
            ]
            
            for i, header in enumerate(headers):
                hdr_cells[i].text = ""
                p = hdr_cells[i].paragraphs[0]
                run = p.add_run(header)
                run.bold = True
                run.font.color.rgb = colores[i]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Filtrar duplicados y solo movimientos reales
            movimientos_unicos = {}
            for mov in resultados_validacion['movimientos_inventario']:
                if mov.get('unidades_movidas', 0) > 0:
                    clave = f"{mov['sku']}_{mov['origen']}_{mov['destino']}"
                    if clave not in movimientos_unicos:
                        movimientos_unicos[clave] = mov
            
            # Datos de la tabla - solo movimientos únicos y reales
            for mov in movimientos_unicos.values():
                row_cells = table.add_row().cells
                
                # SKU
                row_cells[0].text = ""
                p = row_cells[0].paragraphs[0]
                sku_run = p.add_run(str(mov['sku']))
                sku_run.bold = True
                sku_run.font.color.rgb = RGBColor(0, 0, 139)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # ORIGEN
                row_cells[1].text = ""
                p = row_cells[1].paragraphs[0]
                origen_run = p.add_run(str(mov['origen']))
                origen_run.bold = True
                origen_run.font.color.rgb = RGBColor(178, 34, 34)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # DESTINO
                row_cells[2].text = ""
                p = row_cells[2].paragraphs[0]
                destino_run = p.add_run(str(mov['destino']))
                destino_run.bold = True
                destino_run.font.color.rgb = RGBColor(0, 100, 0)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # INV_ORIGEN
                row_cells[3].text = ""
                p = row_cells[3].paragraphs[0]
                inv_origen_run = p.add_run(str(mov['inv_origen']))
                inv_origen_run.bold = True
                inv_origen_run.font.color.rgb = RGBColor(128, 0, 128)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # INV_DESTINO
                row_cells[4].text = ""
                p = row_cells[4].paragraphs[0]
                inv_destino_run = p.add_run(str(mov['inv_destino']))
                inv_destino_run.bold = True
                inv_destino_run.font.color.rgb = RGBColor(0, 100, 139)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # UNID_MV (unidades movidas)
                row_cells[5].text = ""
                p = row_cells[5].paragraphs[0]
                unidades_run = p.add_run(f"{mov['unidades_movidas']:.0f}")
                unidades_run.bold = True
                unidades_run.font.color.rgb = RGBColor(255, 140, 0)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # TOTAL_SKU
                row_cells[6].text = ""
                p = row_cells[6].paragraphs[0]
                total_sku = mov.get('total_unidades_sku', mov['unidades_movidas'])
                total_run = p.add_run(f"{total_sku:.0f}")
                total_run.bold = True
                total_run.font.color.rgb = RGBColor(128, 128, 0)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # RESULTADO
                row_cells[7].text = ""
                p = row_cells[7].paragraphs[0]
                if mov.get('origen_ok', True) and mov.get('destino_ok', True):
                    resultado_run = p.add_run("[OK] OK")
                    resultado_run.font.color.rgb = RGBColor(0, 128, 0)
                else:
                    resultado_run = p.add_run("[ERROR] ERROR")
                    resultado_run.font.color.rgb = RGBColor(255, 0, 0)
                resultado_run.bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
        # Métricas en un solo párrafo
        document.add_heading('Métricas', level=2)
        metrics_p = document.add_paragraph()
        
        metrics_p.add_run('Registros: ').bold = True
        metrics_p.add_run('Inicial: ').bold = True
        reg_ini = metrics_p.add_run(str(resultados_validacion['registros_inicial']))
        reg_ini.font.color.rgb = RGBColor(178, 34, 34)
        reg_ini.bold = True
        
        metrics_p.add_run(' | KTCHPWINRC: ').bold = True
        reg_winrc = metrics_p.add_run(str(resultados_validacion['registros_winrc']))
        reg_winrc.font.color.rgb = RGBColor(0, 0, 139)
        reg_winrc.bold = True
        
        metrics_p.add_run(' | Final: ').bold = True
        reg_final = metrics_p.add_run(str(resultados_validacion['registros_final']))
        reg_final.font.color.rgb = RGBColor(0, 100, 0)
        reg_final.bold = True

        # Discrepancias
        if resultados_validacion.get('discrepancias'):
            document.add_heading('Discrepancias Encontradas', level=2)
            
            # Lista compacta de discrepancias
            for idx, discrepancia in enumerate(resultados_validacion['discrepancias'], 1):
                p = document.add_paragraph(style='List Bullet')
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                
                tipo_run = p.add_run(f"{idx}. {discrepancia['tipo']}: ")
                tipo_run.bold = True
                
                if discrepancia['tipo'] == 'INVENTARIO':
                    tipo_run.font.color.rgb = RGBColor(255, 0, 0)
                elif discrepancia['tipo'] == 'SKU_PERDIDO':
                    tipo_run.font.color.rgb = RGBColor(255, 0, 0)
                elif discrepancia['tipo'] == 'CANTIDAD':
                    tipo_run.font.color.rgb = RGBColor(255, 165, 0)
                elif discrepancia['tipo'] == 'SKU_FALTANTE':
                    tipo_run.font.color.rgb = RGBColor(255, 165, 0)
                elif discrepancia['tipo'] == 'SIN_MOVIMIENTOS':
                    tipo_run.font.color.rgb = RGBColor(255, 165, 0)
                elif discrepancia['tipo'] == 'COLUMNAS_FALTANTES':
                    tipo_run.font.color.rgb = RGBColor(255, 0, 0)
                
                # Resaltar valores numéricos en la descripción
                descripcion = discrepancia['descripcion']
                import re
                
                partes = re.split(r'(\d+)', descripcion)
                for parte in partes:
                    if parte.isdigit():
                        num_run = p.add_run(parte)
                        num_run.bold = True
                        num_run.font.color.rgb = RGBColor(128, 0, 128)
                    else:
                        p.add_run(parte)

        # Conclusión en un párrafo compacto
        document.add_heading('Resumen de la Transferencia', level=1)
        conclusion_p = document.add_paragraph()
        conclusion_p.add_run('Este reporte fue generado automáticamente por el Sistema de Automatización de Transferencias WMS-INFOR. ')
        conclusion_p.add_run('El procesamiento incluyó la ejecución de PROCESAR_ORDENES, el comando SIWINTRACL y la consulta de tablas de inventario.')
        
        # Mensaje de estado
        if resultados_validacion['estado'] == 'OK':
            exito_run = conclusion_p.add_run(' [OK] ÉXITO: Se detectaron movimientos de saldos correctamente.')
            exito_run.bold = True
            exito_run.font.color.rgb = RGBColor(0, 128, 0)
        elif resultados_validacion['estado'] == 'ADVERTENCIA' and 'SIN_MOVIMIENTOS' in [d.get('tipo') for d in resultados_validacion.get('discrepancias', [])]:
            adv_run = conclusion_p.add_run(' ⚠️ ADVERTENCIA: No se detectaron movimientos de saldos entre inventarios.')
            adv_run.bold = True
            adv_run.font.color.rgb = RGBColor(255, 165, 0)
        elif resultados_validacion['estado'] == 'ERROR':
            error_run = conclusion_p.add_run(' ⚠️ ATENCIÓN: Se detectaron errores que requieren revisión manual.')
            error_run.bold = True
            error_run.font.color.rgb = RGBColor(255, 0, 0)
        
        # Guardar documento
        document.save(reporte_path)
        logger.info(f"Documento Word de transferencia WMS-INFOR generado: {os.path.abspath(reporte_path)}")
        return reporte_path
    except Exception as e:
        logger.info(f"Error al generar documento Word de transferencia WMS-INFOR: {e}")
        logger.info(traceback.format_exc())
        return None

def enviar_correo_sin_transacciones(logger):
    """
    Envía correo indicando que no se encontraron transacciones pendientes
    La fecha y hora de ejecución se generan automáticamente
    """
    fecha = datetime.now().strftime('%Y%m%d')
    hora = datetime.now().strftime('%H:%M:%S')

    asunto = f"WMS-INFOR | Sin transacciones pendientes - {fecha}"

    cuerpo = f"""
    <html>
    <body style="font-family: Calibri, Arial, sans-serif; font-size: 14px;">
        <h2 style="color:#2E74B5;">Ejecución automática WMS-INFOR</h2>

        <p>
            Se informa que durante la ejecución automática del proceso de
            <b>validación de transferencias WMS-INFOR</b>,
            realizada en la fecha y hora programada,
            <b>no se encontraron transacciones pendientes por procesar</b>.
        </p>

        <h3 style="color:#5B9BD5;">Resumen de la ejecución</h3>
        <ul>
            <li><b>Fecha:</b> {fecha}</li>
            <li><b>Hora de ejecución:</b> {hora}</li>
            <li><b>Proceso:</b> Validación de Transferencias WMS-INFOR</li>
            <li><b>Resultado:</b> Sin transacciones pendientes</li>
        </ul>

        <p>
            El sistema se ejecutó correctamente y permanecerá atento
            a nuevas transferencias en los siguientes ciclos programados.
        </p>

        <p style="margin-top:20px;">
            Saludos,<br>
            <b>Sistema de Validación de Transferencias WMS-INFOR</b><br>
            Automatización AS400
        </p>
    </body>
    </html>
    """

    enviado = enviar_correo(asunto, cuerpo)

    if enviado:
        logger.info("Correo enviado: sin transacciones pendientes")
    else:
        logger.error("Error al enviar correo: sin transacciones pendientes")

    return enviado


def enviar_correo_resumen_validacion(lista_resultados, logger=None):
    """
    Recibe una lista de diccionarios (cada uno con al menos las claves
    'orden', 'estado', 'observacion' y opcional 'ruta_reporte') y envía
    un único correo con la tabla resumen y todos los adjuntos.
    """
    if not lista_resultados:
        if logger: logger.warning("Lista de resultados vacía. No se enviará correo resumen.")
        return False

    fecha_hoy = datetime.now().strftime('%Y-%m-%d')

    # Contadores
    total = len(lista_resultados)
    errores = sum(1 for r in lista_resultados if r.get('estado') != 'OK')

    # Preparar adjuntos (solo rutas válidas)
    adjuntos_totales = [r.get('ruta_reporte') for r in lista_resultados if r.get('ruta_reporte') and os.path.exists(r.get('ruta_reporte'))]

    # Generar filas de la tabla HTML
    filas_html = ""
    for res in lista_resultados:
        orden = res.get('orden', '')
        estado = res.get('estado', '')

        # Formatear observaciones
        obs_raw = res.get('observacion', [])
        if hasattr(obs_raw, 'tolist'):
            try:
                obs_raw = obs_raw.tolist()
            except Exception:
                pass

        # Si es lista de discrepancias, sacar descripciones
        obs_list = []
        if isinstance(obs_raw, (list, tuple)):
            for o in obs_raw:
                if isinstance(o, dict) and 'descripcion' in o:
                    obs_list.append(str(o.get('descripcion')))
                else:
                    obs_list.append(str(o))
        else:
            obs_list = [str(obs_raw)]

        obs_str = ", ".join([x for x in obs_list if x and str(x).strip()])
        if not obs_str:
            obs_str = "Sin incidencias"

        color = "#D32F2F" if estado != "OK" else "#2E7D32"

        filas_html += f"""
        <tr>
            <td style="padding:5px; border-bottom:1px solid #ddd;">{orden}</td>
            <td style="padding:5px; border-bottom:1px solid #ddd; font-weight:bold; color:{color};">{estado}</td>
            <td style="padding:5px; border-bottom:1px solid #ddd; font-size:12px;">{obs_str}</td>
        </tr>
        """

    asunto_estado = "CON ERRORES" if errores > 0 else "EXITOSO"
    asunto = f"[RESUMEN TRANSFERENCIAS] Validación WMS-INFOR - {fecha_hoy}"

    cuerpo = f"""
    <html>
    <body style="font-family: Arial, sans-serif; font-size: 14px;">
        <h3 style="color:#2E74B5;">Resumen de Ejecución</h3>
        <p>Se han procesado <b>{total}</b> órdenes.</p>
        
        <table style="width:100%; border-collapse: collapse; border: 1px solid #ccc;">
            <tr style="background-color: #f2f2f2; text-align: left;">
                <th style="padding:8px; border-bottom:2px solid #ccc;">Orden</th>
                <th style="padding:8px; border-bottom:2px solid #ccc;">Estado</th>
                <th style="padding:8px; border-bottom:2px solid #ccc;">Observaciones</th>
            </tr>
            {filas_html}
        </table>
        
        <p style="margin-top:15px; color:#666; font-size:12px;">
            * Se adjuntan los reportes detallados de cada orden.
        </p>
    </body>
    </html>
    """

    if logger: logger.info(f"Enviando correo global con {len(adjuntos_totales)} adjuntos.")
    return enviar_correo(asunto, cuerpo, adjuntos=adjuntos_totales)

def procesar_transferencias_wms_infor(BASE='RI11DB'):
    """
    Procesa todas las transferencias WMS-INFOR según el flujo especificado.
    """
    logger = get_logger(__file__)
    logger.info("=== INICIANDO PROCESAMIENTO DE TRANSFERENCIAS WMS-INFOR ===")
    
    # ** CONFIGURACIÓN DE LÍMITE DE ÓRDENES (AJUSTAR PARA PRODUCCIÓN) **
    LIMITE_ORDENES = None # Configurado en 1 para pruebas. Poner None para procesar todas.
    # LIMITE_ORDENES = None  # En producción procesamos todas las órdenes
    
    tiempo_inicio = datetime.now()
    
    try:
        conexion_as400 = ConexionAS400()
        conexion = conexion_as400.conectar()
    except Exception as e:
        logger.error(f"Error al inicializar o conectar AS400: {e}")
        logger.info("[ERROR] No se pudo establecer conexión con AS400. Verificar configuración ODBC.")
        return False
    
    if not conexion:
        logger.error("No se pudo establecer conexión con AS400")
        logger.info("[ERROR] No se pudo establecer conexión con AS400. Verificar configuración ODBC.")
        return False
    
    try:
        logger.info("Obteniendo lista de órdenes de transferencias...")
        query_ordenes = f"SELECT DISTINCT trim(NUMORDEN) as NUMORDEN FROM {BASE}.KTCHPWASIN a"
        if LIMITE_ORDENES is not None and LIMITE_ORDENES > 0:
            query_ordenes += f" FETCH FIRST {LIMITE_ORDENES} ROWS ONLY"
        
        df_ordenes = ejecutar_query(conexion, query_ordenes)

        
        
        if df_ordenes is None or df_ordenes.empty:
            #enviar_correo_sin_transacciones(logger)
            logger.warning("No se encontraron órdenes para procesar")
            return False
        else:
            logger.info(f"Se encontraron {len(df_ordenes)} órdenes para procesar")
        
        fecha_actual = datetime.now().strftime('%Y%m%d')
        reporte_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Reportes', f'Transferencias_WMS_INFOR_{fecha_actual}')
        os.makedirs(reporte_dir, exist_ok=True)
        
        reportes_generados = []
        resultados_por_orden = {}
        
        for idx, row in df_ordenes.iterrows():
            numero_orden = row['NUMORDEN']
            logger.info(f"Procesando orden {idx+1}/{len(df_ordenes)}: {numero_orden}")
            
            # 1. Consultar datos iniciales de la orden
            logger.info(f"Consultando datos iniciales de la orden {numero_orden} (KTCHPWASIN)...")
            query_orden = f"""
                SELECT SFTTD, DOCTD, STTTD, SKUTD, U71TD, O43TD, D37TD, TIPO, DOCTDNEW, ASN, NUMORDEN 
                FROM {BASE}.KTCHPWASIN 
                WHERE NUMORDEN = '{numero_orden}'
            """
            df_orden = ejecutar_query(conexion, query_orden)
            
            if df_orden is None or df_orden.empty:
                logger.warning(f"No se encontraron datos para la orden {numero_orden}")
                continue
            b = BASE[2:4]

            
            
            # 2. Ejecutar primer llamado a PROCESAR_ORDENES (Captura Inicial)
            logger.info(f"Ejecutando primer PROCESAR_ORDENES (Captura Inicial) para {numero_orden}...")
            proc_call = f"CALL ELOPEZ.PROCESAR_ORDENES('{BASE}','{numero_orden}')"
            ejecutar_stored_procedure(conexion, proc_call)
            
            # 3. Ejecutar comando CL SIWINTRACL (Posteo a INFOR)
            logger.info(f"Ejecutando SIWINTRACL con parámetro '{b}' como trabajo...")
            comando_cl = f"CALL PGM(RIUNICOM63/SIWINTRACL) PARM('{b}')"
            job_info = ejecutar_comando_cl(comando_cl, submit_job=True, job_name=f"WMSI_TF{b}")
            
            if job_info:
                logger.info(f"Trabajo CL enviado: {job_info.get('job_name')} (ID: {job_info.get('job_id')})")
            
            # 4. Monitoreo de finalización
            USE_MONITOREO_INTELIGENTE = True
            TIMEOUT_MINUTOS = 15
            
            if USE_MONITOREO_INTELIGENTE:
                logger.info("Monitoreando finalización del trabajo SIWINTRACL...")
                tabla_origen ='KTCHPWASIN'
                tabla_destino='KTCHPWINRC'
                job_name=f"WMSI_TF{b}"
                #(numero_orden,conexion,logger,tabla_origen,tabla_destino,job_name=None,timeout_minutos=5)
                trabajo_finalizado = esperar_finalizacion_trabajo(BASE,numero_orden, conexion, logger,tabla_origen,tabla_destino,job_name, TIMEOUT_MINUTOS)
                
                if not trabajo_finalizado:
                    logger.warning(f"El trabajo no finalizó en el tiempo esperado ({TIMEOUT_MINUTOS} min). Continuando...")
                else:
                    logger.info("Trabajo SIWINTRACL finalizado exitosamente (por tabla o timeout con progreso).")
            else:
                logger.info(f"Esperando 30 segundos (espera fija) para procesamiento...")
                time.sleep(30)
            
            # 5. Consultar INVENTARIO_INICIAL (Ya capturado, solo se consulta)
            logger.info(f"Consultando INVENTARIO_INICIAL para orden {numero_orden}...")
            query_inv_inicial = f"SELECT * FROM {BASE}.INVENTARIO_INICIAL WHERE NUMORDEN = '{numero_orden}' order by sku desc "
            df_inv_inicial = ejecutar_query(conexion, query_inv_inicial)

            
        
            # 6. Consultar KTCHPWINRC (Registros transferidos)
            logger.info(f"Consultando KTCHPWINRC para orden {numero_orden}...")
            query_ktchpwinrc = f"""
                SELECT SFTTD, DOCTD, STTTD, SKUTD, U71TD, O43TD, D37TD, TIPO, DOCTDNEW, ASN, NUMORDEN
                FROM {BASE}.KTCHPWINRC 
                WHERE NUMORDEN = '{numero_orden}'
            """
            df_ktchpwinrc = ejecutar_query(conexion, query_ktchpwinrc)
            
            # 7. Ejecutar segundo llamado a PROCESAR_ORDENES (Captura Final)
            logger.info(f"Ejecutando segundo ELOPEZ.PROCESAR_ORDENES (Captura Final) para {numero_orden}...")
            all = ejecutar_stored_procedure(conexion, proc_call)

            logger.info(f"ejecución del procedimiento {all}")
                        
            # 8. Consultar INVENTARIO_FINAL
            logger.info(f"Consultando INVENTARIO_FINAL para orden {numero_orden}...")
            query_inv_final = f"SELECT * FROM {BASE}.INVENTARIO_FINAL WHERE NUMORDEN = '{numero_orden}' order by sku desc "

            logger.info(query_inv_final)
            df_inv_final = ejecutar_query(conexion, query_inv_final)
            # 9. Realizar Validación Matemática
            logger.info(f"Realizando validación matemática para orden {numero_orden}...")
            resultados_validacion = validar_inventario(df_inv_inicial, df_inv_final, df_orden, logger)
            mostrar_resumen_10_lineas(resultados_validacion)
            
            resultados_por_orden[numero_orden] = resultados_validacion
            
            # 10. Generar reporte Word
            reporte_nombre = f"Reporte_Orden_{numero_orden}_{fecha_actual}.docx"
            reporte_path = os.path.join(reporte_dir, reporte_nombre)
            
            word_path = generar_word_reporte_con_validacion(
                df_orden, 
                df_inv_inicial, 
                df_ktchpwinrc, 
                df_inv_final,
                resultados_validacion,
                numero_orden, 
                reporte_path
            )
            
            # 11. Subir evidencia a Jira
            if word_path:
                # Enlazar ruta de reporte y metadatos al resultado para resumen global
                resultados_validacion['ruta_reporte'] = word_path
                resultados_validacion['orden'] = numero_orden
                # Observaciones: preferir discrepancias si existen, sino mensaje
                if resultados_validacion.get('discrepancias'):
                    resultados_validacion['observacion'] = resultados_validacion.get('discrepancias')
                else:
                    resultados_validacion['observacion'] = [resultados_validacion.get('mensaje', '')]

                logger.info(f"Documento Word generado: {word_path}")
                reportes_generados.append(word_path)
                
                #subir_evidencia_a_jira(
                #    numero_orden=numero_orden,
                #    reporte_path=word_path,
                #    resultados_validacion=resultados_validacion,
                #    logger=logger,
                #    issue_key=None 
                #)

                #Maneja todo de relacionado a jira. 
                #Generación de la Tarea
                #Generación de comentario
                #Subida de las evidencias
                tipo = "transferencia"
                
                jira = JiraClient()
                jira.main_jira(numero_orden,reporte_path, tipo)
                

            else:
                logger.warning(f"No se pudo generar el documento Word para la orden {numero_orden}")
        
        # 12. Enviar Correo Electrónico
        tiempo_fin = datetime.now()
        duracion = tiempo_fin - tiempo_inicio
        duracion_str = str(duracion).split('.')[0]
        
        if reportes_generados:
            logger.info(f"Enviando correo con {len(reportes_generados)} reportes...")
            
            reportes_existentes = [r for r in reportes_generados if os.path.exists(r)]
            
            if not reportes_existentes:
                logger.error("Ninguno de los reportes generados existe físicamente. No se enviará correo.")
                return False
                
            # Métricas para el cuerpo del correo
            advertencias = 0
            errores = 0
            ordenes_procesadas = []
            productos_con_multiples_lineas = 0
            total_lineas_agrupadas = 0
            skus_duplicados = []
            total_origenes_validados = 0
            total_destinos_validados = 0
            total_registros_procesados = 0
            
            for numero_orden_res, resultado in resultados_por_orden.items():
                estado = resultado.get('estado', 'OK')
                if estado == 'ERROR': errores += 1
                elif estado == 'ADVERTENCIA': advertencias += 1
                
                if numero_orden_res in [row['NUMORDEN'] for _, row in df_ordenes.iterrows()]:
                    ordenes_procesadas.append(numero_orden_res)
                
                if 'tablas_validacion' in resultado:
                    total_origenes_validados += len(resultado['tablas_validacion'].get('origen', pd.DataFrame()))
                    total_destinos_validados += len(resultado['tablas_validacion'].get('destino', pd.DataFrame()))
                total_registros_procesados += resultado.get('registros_inicial', 0)
                
                movimientos = resultado.get('movimientos_inventario', [])
                for mov in movimientos:
                    cantidad_lineas = mov.get('cantidad_lineas', 1)
                    if cantidad_lineas > 1:
                        productos_con_multiples_lineas += 1
                        total_lineas_agrupadas += cantidad_lineas
                        skus_duplicados.append({
                            'sku': mov['sku'],
                            'lineas': cantidad_lineas,
                            'unidades_total': mov['unidades_movidas'],
                            'orden': numero_orden_res
                        })

            # Preparar lista de resultados para el correo resumen global
            lista_resultados = []
            for numero_orden_res, resultado in resultados_por_orden.items():
                ruta = resultado.get('ruta_reporte')
                orden = numero_orden_res
                estado = resultado.get('estado', 'OK')
                # Observaciones preferibles: discrepancias o mensaje
                if resultado.get('discrepancias'):
                    observ = resultado.get('discrepancias')
                else:
                    observ = [resultado.get('mensaje', '')]

                lista_resultados.append({
                    'orden': orden,
                    'estado': estado,
                    'observacion': observ,
                    'ruta_reporte': ruta
                })

            enviado = False
            intentos = 0
            max_intentos = 3

            while not enviado and intentos < max_intentos:
                intentos += 1
                logger.info(f"Intento {intentos} de envío de correo resumen global...")
                enviado = enviar_correo_resumen_validacion(lista_resultados, logger=logger)
                if not enviado and intentos < max_intentos:
                    logger.warning(f"Fallo en el intento {intentos}. Reintentando en 5 segundos...")
                    time.sleep(5)

            if enviado:
                logger.info("Correo resumen global enviado exitosamente")
            else:
                logger.error("Todos los intentos de envío de correo resumen fallaron")
        else:
            logger.warning("No se generaron reportes para enviar por correo")
        
        logger.info("=== PROCESAMIENTO DE TRANSFERENCIAS WMS-INFOR FINALIZADO ===")
        return True
        
    except Exception as e:
        logger.error(f"Error en el procesamiento de transferencias WMS-INFOR: {e}")
        logger.error(traceback.format_exc())
        return False
    finally:
        if conexion:
            conexion.close()
            logger.info("Conexión a AS400 cerrada")
            return logger

#procesar_ordenes = procesar_transferencias_wms_infor(Base)

if __name__ == "__main__":
    try:
        BASE ='RI12DB'
        logger.info("=== INICIANDO VALIDACIÓN WMS-INFOR ===")
        logger.info(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info("Procesando transferencias WMS-INFOR...")
        
        procesar_transferencias_wms_infor(BASE)
        
        logger.info("=== PROCESO COMPLETADO ===")
    except Exception as e:
        logger.info(f"ERROR en la ejecución: {e}")
        traceback.logger.info_exc()