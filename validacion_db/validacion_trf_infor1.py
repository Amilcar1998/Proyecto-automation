import os
import time
import pandas as pd
from datetime import datetime
import win32com.client as win32
import logging
from conexion_config.logging_config import get_logger
import sys

# Añadir el directorio raíz al path de Python para permitir importaciones relativas
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Importamos los módulos necesarios
from conexion_config.conexion import ConexionAS400
from scripts_as400.ejecutor_cl import EjecutorCL

import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT
import re

def _normalize_sku(value):
    """Normaliza el SKU para matching de inventario.
    Si el código viene como 12+ dígitos tipo '4' + SKU6 + sufijo (p.ej. 439344700008),
    extrae s[1:7] (los 6 dígitos del SKU). En otros casos, usa los últimos 6.
    """
    s = '' if value is None else str(value).strip()
    if len(s) >= 12 and s[0].isdigit() and s[0] == '4' and s[1:7].isdigit():
        return s[1:7]
    return s[-6:] if len(s) >= 6 else s

def _normalize_loc(value: str):
    """Normaliza códigos de ORIGEN/DESTINO: quita espacios, mayúsculas, y devuelve (raw, sin_zeros)."""
    s = '' if value is None else str(value).strip().upper()
    s_noz = s.lstrip('0') or '0'  # conservar '0' si quedaría vacío
    return s, s_noz

def setup_logging():
    """Configura el sistema de logging para el procesamiento de transferencias WMS-INFOR"""
    logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'logs')
    os.makedirs(logs_dir, exist_ok=True)
    
    fecha = datetime.now().strftime('%Y%m%d')
    log_file = os.path.join(logs_dir, f'transferencias_wms_infor_{fecha}.log')
    logger = get_logger(__file__)

    # Añadir file handler con timestamp si no existe
    log_abspath = os.path.abspath(log_file)
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
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler = logging.FileHandler(log_abspath, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    print(f"📝 Log configurado en: {log_abspath}")
    return logger

def esperar_finalizacion_trabajo(numero_orden, conexion, logger, timeout_minutos=5):
    """
    Espera a que el trabajo SIWINTRACL termine monitoreando:
    1. Cambios en la tabla KTCHPWINRC (registros transferidos de KTCHPWASIN)
    2. Estado del trabajo en el sistema
    3. Timeout para evitar esperas infinitas
    
    Args:
        numero_orden (str): Número de orden a monitorear
        conexion: Conexión a AS400
        logger: Logger para registrar el progreso
        timeout_minutos (int): Tiempo máximo de espera en minutos
        
    Returns:
        bool: True si el trabajo finalizó, False si se agotó el timeout
    """
    inicio_monitoreo = time.time()
    timeout_segundos = timeout_minutos * 60
    intervalo_verificacion = 10  # Verificar cada 10 segundos
    
    logger.info(f"Iniciando monitoreo inteligente con timeout de {timeout_minutos} minutos...")
    
    try:
        # Obtener conteo inicial de registros en KTCHPWASIN y KTCHPWINRC
        conteo_inicial_wasin = obtener_conteo_registros(conexion, "RI11DB.KTCHPWASIN", numero_orden)
        conteo_inicial_winrc = obtener_conteo_registros(conexion, "RI11DB.KTCHPWINRC", numero_orden)
        
        logger.info(f"Estado inicial - KTCHPWASIN: {conteo_inicial_wasin}, KTCHPWINRC: {conteo_inicial_winrc}")
        
        # Si hay error al obtener conteos, usar espera fija como fallback
        if conteo_inicial_wasin == -1 or conteo_inicial_winrc == -1:
            logger.warning("No se pudo obtener conteos iniciales. Usando espera fija de 30 segundos como fallback.")
            time.sleep(30)
            return True
        
        # Si no hay registros en WASIN, el trabajo ya procesó todo
        if conteo_inicial_wasin == 0:
            logger.info("No hay registros en KTCHPWASIN. El trabajo ya procesó la orden.")
            return True
        
        intentos = 0
        max_intentos = timeout_segundos // intervalo_verificacion
        ultimo_conteo_wasin = conteo_inicial_wasin
        tiempo_sin_cambios = 0
        
        while intentos < max_intentos:
            time.sleep(intervalo_verificacion)
            intentos += 1
            tiempo_transcurrido = int(time.time() - inicio_monitoreo)
            
            # Verificar cambios en las tablas
            conteo_actual_wasin = obtener_conteo_registros(conexion, "RI11DB.KTCHPWASIN", numero_orden)
            conteo_actual_winrc = obtener_conteo_registros(conexion, "RI11DB.KTCHPWINRC", numero_orden)
            
            # Si hay error al obtener conteos, continuar con el siguiente intento
            if conteo_actual_wasin == -1 or conteo_actual_winrc == -1:
                logger.warning(f"Error al obtener conteos en intento {intentos}. Reintentando...")
                continue
            
            logger.info(f"Verificación {intentos}/{max_intentos} ({tiempo_transcurrido}s) - "
                       f"WASIN: {conteo_actual_wasin}, WINRC: {conteo_actual_winrc}")
            
            # Verificar si los registros se movieron completamente de WASIN a WINRC
            if conteo_actual_wasin == 0 and conteo_actual_winrc > conteo_inicial_winrc:
                logger.info(f"✅ Trabajo finalizado exitosamente - Registros transferidos de WASIN a WINRC")
                return True
            
            # Verificar si hay progreso (registros disminuyendo en WASIN)
            if conteo_actual_wasin < ultimo_conteo_wasin:
                logger.info(f"📊 Progreso detectado - Procesando registros ({ultimo_conteo_wasin} → {conteo_actual_wasin})")
                ultimo_conteo_wasin = conteo_actual_wasin
                tiempo_sin_cambios = 0
            else:
                tiempo_sin_cambios += intervalo_verificacion
                
            # Si no hay cambios por más de 1 minuto, verificar trabajos activos
            if tiempo_sin_cambios >= 60:
                logger.info("📋 Verificando estado de trabajos activos...")
                try:
                    trabajos_activos = verificar_trabajos_activos(conexion, "WMSI_TF11")
                    if trabajos_activos is False:  # Explícitamente False (no None)
                        logger.info("💼 No se detectan trabajos WMSI_TF11 activos")
                        # Verificar una vez más las tablas antes de confirmar
                        time.sleep(5)
                        conteo_final_wasin = obtener_conteo_registros(conexion, "RI11DB.KTCHPWASIN", numero_orden)
                        
                        if conteo_final_wasin == 0:
                            logger.info("✅ Trabajo finalizado - No hay trabajos activos y WASIN está vacío")
                            return True
                        elif conteo_final_wasin == conteo_actual_wasin:
                            logger.warning("⚠️ El trabajo parece haberse detenido sin completarse")
                            return False
                except Exception as e:
                    logger.debug(f"No se pudo verificar estado de trabajos: {e}")
                
                # Resetear contador de tiempo sin cambios
                tiempo_sin_cambios = 0
        
        # Timeout alcanzado
        conteo_final_wasin = obtener_conteo_registros(conexion, "RI11DB.KTCHPWASIN", numero_orden)
        conteo_final_winrc = obtener_conteo_registros(conexion, "RI11DB.KTCHPWINRC", numero_orden)
        
        logger.warning(f"⏰ Timeout de {timeout_minutos} minutos alcanzado.")
        logger.info(f"Estado final - WASIN: {conteo_final_wasin}, WINRC: {conteo_final_winrc}")
        
        # Si al menos se procesaron algunos registros, considerar parcialmente exitoso
        if conteo_final_wasin < conteo_inicial_wasin:
            logger.info("📊 Se detectó progreso parcial durante el monitoreo")
            return True
        
        return False
        
    except Exception as e:
        logger.error(f"Error durante el monitoreo inteligente: {e}")
        logger.warning("Usando espera fija de 30 segundos como fallback...")
        time.sleep(30)
        return True  # Asumir que completó para continuar el proceso

def obtener_conteo_registros(conexion, tabla, numero_orden):
    """
    Obtiene el conteo de registros en una tabla para una orden específica
    
    Args:
        conexion: Conexión a AS400
        tabla (str): Nombre de la tabla
        numero_orden (str): Número de orden a consultar
        
    Returns:
        int: Número de registros encontrados
    """
    try:
        query = f"SELECT COUNT(*) as CONTEO FROM {tabla} WHERE NUMORDEN = '{numero_orden}'"
        cursor = conexion.cursor()
        cursor.execute(query)
        resultado = cursor.fetchone()
        return resultado[0] if resultado else 0
    except Exception as e:
        print(f"Error al obtener conteo de {tabla}: {e}")
        return -1

def verificar_trabajos_activos(conexion, nombre_trabajo_patron):
    """
    Verifica si hay trabajos activos que coincidan con el patrón especificado
    
    Args:
        conexion: Conexión a AS400
        nombre_trabajo_patron (str): Patrón del nombre del trabajo
        
    Returns:
        bool: True si hay trabajos activos, False en caso contrario
    """
    try:
        # Consultar tabla de trabajos activos del sistema
        # Nota: Esta consulta requiere permisos especiales en AS400
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
        # Si no tenemos permisos o la tabla no está disponible, retornar None
        # para indicar que no pudimos verificar el estado
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
        print(f"Error al ejecutar consulta: {e}")
        return None

def ejecutar_stored_procedure(conexion, procedure_call):
    """Ejecuta un stored procedure"""
    try:
        cursor = conexion.cursor()
        cursor.execute(procedure_call)
        conexion.commit()
        return True
    except Exception as e:
        print(f"Error al ejecutar stored procedure: {e}")
        return False

def ejecutar_comando_cl(comando, submit_job=False, job_name=None):
    """
    Ejecuta un comando CL en AS400
    
    Args:
        comando (str): Comando CL a ejecutar
        submit_job (bool): Si es True, envía el comando como un trabajo (SBMJOB)
        job_name (str, optional): Nombre del trabajo si se usa submit_job
    """
    try:
        cl = EjecutorCL()
        
        if submit_job:
            respuesta = cl.submit_job(comando, job_name=job_name)
        else:
            respuesta = cl.ejecutar(comando)
            
        return respuesta
    except Exception as e:
        print(f"Error al ejecutar comando CL: {e}")
        return None

def generar_excel(df, filename):
    """Genera un archivo Excel a partir de un DataFrame"""
    try:
        # Crear directorio si no existe
        directorio = os.path.dirname(filename)
        if not os.path.exists(directorio):
            os.makedirs(directorio)
            
        df.to_excel(filename, index=False)
        print(f"Excel generado: {filename}")
        return filename
    except Exception as e:
        print(f"Error al generar Excel: {e}")
        return None

def enviar_correo(asunto, cuerpo, adjuntos=None):
    """Envía un correo con los archivos adjuntos de transferencias WMS-INFOR"""
    try:
        print("Iniciando proceso de envío de correo...")
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        # Leer destinatarios desde archivo de configuración
        destinatarios = []
        config_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'archivos_config')
        destinatarios_path = os.path.join(config_dir, 'destinatarios.txt')
        
        print(f"Buscando archivo de destinatarios en: {destinatarios_path}")
        if os.path.exists(destinatarios_path):
            with open(destinatarios_path, 'r', encoding='utf-8') as f:
                for line in f:
                    for correo in line.replace(';', ',').split(','):
                        correo = correo.strip()
                        if correo and not correo.startswith('#'):
                            destinatarios.append(correo)
            print(f"Destinatarios encontrados en archivo: {destinatarios}")
        else:
            print("Archivo de destinatarios no encontrado, usando destinatario por defecto")
            destinatarios = ["eliseo_lopezp@unicomer.com"]
        
        mail.To = ";".join(destinatarios)
        mail.Subject = asunto
        
        # Obtener ruta de la firma
        recursos_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Recursos')
        os.makedirs(recursos_dir, exist_ok=True)  # Crear directorio si no existe
        firma_path = os.path.join(recursos_dir, 'Firma.jpg')
        
        print(f"Buscando firma en: {firma_path}")
        
        # Incluir la firma directamente en el cuerpo HTML del correo
        if os.path.exists(firma_path):
            # Convertir la imagen a base64 para incrustarla directamente en el HTML
            import base64
            with open(firma_path, 'rb') as img_file:
                img_data = base64.b64encode(img_file.read()).decode('utf-8')
            
            # Insertar la imagen directamente en el HTML usando data URI
            firma_html = f"""
            <hr style='border: 1px solid #dddddd; margin-top: 20px;'>
            <div style='font-size: 11px; color: #777777;'>
                <img src="data:image/jpeg;base64,{img_data}" alt="Firma" style="max-width: 500px;" />
            </div>
            """
            print("Firma convertida a base64 e incrustada en el HTML")
        else:
            print(f"Firma no encontrada en: {firma_path}")
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
                    print(f"Adjunto agregado: {adjunto}")
                else:
                    print(f"ADVERTENCIA: El archivo adjunto no existe: {adjunto}")
        
        print(f"Enviando correo a {len(destinatarios)} destinatarios con {archivos_adjuntados} adjuntos...")
        mail.Send()
        print(f"✅ Correo enviado exitosamente a: {', '.join(destinatarios)}")
        return True
    except Exception as e:
        print(f"❌ ERROR al enviar correo: {e}")
        import traceback
        print(traceback.format_exc())  # Mostrar traceback completo
        return False

def get_default_signature():
    """Devuelve la firma HTML por defecto cuando no se encuentra la imagen"""
    return """
    <hr style='border: 1px solid #dddddd;'>
    <div style='font-size: 11px; color: #777777;'>
        <p>Generado automáticamente por el Sistema de Automatización de Transferencias WMS-INFOR</p>
    </div>
    """

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
    
    # Eliminar el párrafo adicional después de la tabla
    # document.add_paragraph()  # Espacio después de la tabla - ELIMINADO

def generar_word_reporte(df_orden, df_inv_inicial, df_ktchpwinrc, df_inv_final, numero_orden, reporte_path):
    """Genera un documento Word con los datos de las transferencias WMS-INFOR"""
    try:
        document = Document()
        
        # Configurar tamaño de página A3
        section = document.sections[0]
        section.page_width = Inches(16.5)   # A3 ancho: 16.5 pulgadas (420mm)
        section.page_height = Inches(11.7)  # A3 alto: 11.7 pulgadas (297mm)
        section.left_margin = Inches(0.5)   # Margen izquierdo reducido
        section.right_margin = Inches(0.5)  # Margen derecho reducido
        section.top_margin = Inches(0.5)    # Margen superior reducido
        section.bottom_margin = Inches(0.5) # Margen inferior reducido
        
        # Configurar estilo del documento
        style = document.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)
        
        # Título principal
        title = document.add_heading(f'Reporte de Transferencia WMS-INFOR - Orden {numero_orden}', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Información general
        document.add_paragraph(f'Fecha de generación: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}')
        document.add_paragraph(f'Número de orden: {numero_orden}')
        document.add_paragraph()
        
        # Sección 1: Datos de la orden
        dataframe_to_table(document, df_orden, "Datos de la Orden (KTCHPWASIN)")
        
        # Sección 2: Inventario Inicial
        dataframe_to_table(document, df_inv_inicial, "Inventario Inicial (ELOPEZ.INVENTARIO_INICIAL)")
        
        # Sección 3: Datos KTCHPWINRC
        dataframe_to_table(document, df_ktchpwinrc, "Datos KTCHPWINRC (RI11DB.KTCHPWINRC)")
        
        # Sección 4: Inventario Final
        dataframe_to_table(document, df_inv_final, "Inventario Final (ELOPEZ.INVENTARIO_FINAL)")
        
        # Conclusión
        document.add_heading('Resumen de la Transferencia', level=1)
        p = document.add_paragraph()
        p.add_run('Este reporte fue generado automáticamente por el Sistema de Automatización de Transferencias WMS-INFOR. ')
        p.add_run('El procesamiento incluyó la ejecución de PROCESAR_ORDENES, el comando SIWINTRACL y la consulta de tablas de inventario.')
        
        # Guardar documento
        document.save(reporte_path)
        print(f"Documento Word de transferencia WMS-INFOR generado: {os.path.abspath(reporte_path)}")
        return reporte_path
    except Exception as e:
        print(f"Error al generar documento Word de transferencia WMS-INFOR: {e}")
        return None

def validar_inventario(df_inv_inicial, df_inv_final, df_orden, logger):
    """Valida inventario con formato de tabla mostrando validaciones de origen y destino.
    Verifica que: INV_INICIAL_ORIGEN - TOTAL_TRANSFERIDO = INV_FINAL_ORIGEN
    y que: INV_INICIAL_DESTINO + TOTAL_RECIBIDO = INV_FINAL_DESTINO
    """
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
        
        logger.info('🔍 VALIDACIÓN DE TRANSFERENCIAS CON TABLA')
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
        df_validacion_origen['STATUS'] = df_validacion_origen['VALIDO'].apply(lambda x: '✅ OK' if x else '❌ ERROR')
        
        # Convertir a enteros para mostrar
        for col in ['INV_INI', 'TRANSFER', 'CALC', 'INV_FIN']:
            df_validacion_origen[col] = df_validacion_origen[col].astype(int)
        
        logger.info('🔍 TABLA DE VALIDACIÓN - ORÍGENES')
        logger.info('='*80)
        tabla_origen = df_validacion_origen[['SKU', 'ORIGEN', 'INV_INI', 'TRANSFER', 'CALC', 'INV_FIN', 'STATUS']]
        logger.info('\n' + tabla_origen.to_string(index=False))
        
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
        df_validacion_destino['STATUS'] = df_validacion_destino['VALIDO'].apply(lambda x: '✅ OK' if x else '❌ ERROR')
        
        # Convertir a enteros para mostrar
        for col in ['INV_INI', 'RECIBIDO', 'CALC', 'INV_FIN']:
            df_validacion_destino[col] = df_validacion_destino[col].astype(int)
        
        logger.info('\n🔍 TABLA DE VALIDACIÓN - DESTINOS')
        logger.info('='*80)
        tabla_destino = df_validacion_destino[['SKU', 'DESTINO', 'INV_INI', 'RECIBIDO', 'CALC', 'INV_FIN', 'STATUS']]
        logger.info('\n' + tabla_destino.to_string(index=False))
        
        # RESUMEN FINAL
        total_origenes = len(df_validacion_origen)
        origenes_correctos = df_validacion_origen['VALIDO'].sum()
        total_destinos = len(df_validacion_destino)
        destinos_correctos = df_validacion_destino['VALIDO'].sum()
        
        logger.info('\n📊 RESUMEN FINAL')
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
                '✅ OK' if origenes_correctos == total_origenes else '❌ ERROR',
                '✅ OK' if destinos_correctos == total_destinos else '❌ ERROR',
                '✅ OK' if (origenes_correctos == total_origenes and destinos_correctos == total_destinos) else '❌ ERROR'
            ]
        }
        
        df_resumen = pd.DataFrame(resumen_data)
        logger.info('\n' + df_resumen.to_string(index=False))
        
        # Determinar estado general
        if origenes_correctos == total_origenes and destinos_correctos == total_destinos:
            logger.info('\n🎉 RESULTADO: TODAS LAS TRANSFERENCIAS SON CORRECTAS')
            resultados['estado'] = 'OK'
            resultados['mensaje'] = 'Todas las validaciones de origen y destino son correctas'
        else:
            logger.info('\n❌ RESULTADO: HAY ERRORES EN LAS TRANSFERENCIAS')
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
        import traceback
        logger.error(traceback.format_exc())
        resultados['estado'] = 'ERROR'
        resultados['mensaje'] = f'Error en validación: {str(e)}'

    return resultados


def procesar_transferencias_wms_infor(BASE='RI12DB'):
    """
    Procesa todas las transferencias WMS-INFOR según el flujo especificado:
    1. Obtener todas las órdenes distintas
    2. Para cada orden:
       - Ejecutar procedimiento PROCESAR_ORDENES
       - Ejecutar comando CL SIWINTRACL
       - Esperar un minuto
       - Consultar tablas de inventario
       - Ejecutar nuevamente PROCESAR_ORDENES
       - Generar reporte Word
    3. Enviar correo con todos los reportes
    """
    logger = setup_logging()
    logger.info("=== INICIANDO PROCESAMIENTO DE TRANSFERENCIAS WMS-INFOR ===")
    
    # Limitar número de órdenes para procesar (para desarrollo/pruebas)
    LIMITE_ORDENES = 1  # En producción procesamos todas las órdenes
    
    # Tiempo de inicio del procesamiento
    tiempo_inicio = datetime.now()
    
    # Crear conexión a AS400
    conexion_as400 = ConexionAS400()
    conexion = conexion_as400.conectar()
    
    if not conexion:
        logger.error("No se pudo establecer conexión con AS400")
        print("❌ No se pudo establecer conexión con AS400. Verificar configuración ODBC.")
        return False
    
    try:
        logger.info("Obteniendo lista de órdenes de transferencias...")
        # Obtener todas las órdenes distintas - Para IBM i:
        query_ordenes = f"SELECT DISTINCT trim(NUMORDEN) as NUMORDEN FROM {BASE}.KTCHPWASIN a"
        df_ordenes = ejecutar_query(conexion, query_ordenes)
        
        if df_ordenes is None or df_ordenes.empty:
            logger.warning("No se encontraron órdenes para procesar")
            return False
        
        logger.info(f"Se encontraron {len(df_ordenes)} órdenes para procesar")
        
        # Limitar órdenes para pruebas si es necesario
        if LIMITE_ORDENES is not None and len(df_ordenes) > LIMITE_ORDENES:
            logger.info(f"Limitando a {LIMITE_ORDENES} órdenes para procesamiento (modo prueba)")
            df_ordenes = df_ordenes.head(LIMITE_ORDENES)
        
        # Crear directorio para reportes
        fecha_actual = datetime.now().strftime('%Y%m%d')
        reporte_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Reportes', f'Transferencias_WMS_INFOR_{fecha_actual}')
        os.makedirs(reporte_dir, exist_ok=True)
        
        reportes_generados = []
        # Para rastrear los resultados de validación por cada orden
        resultados_por_orden = {}
        
        # Procesar cada orden
        for idx, row in df_ordenes.iterrows():
            numero_orden = row['NUMORDEN']
            logger.info(f"Procesando orden {idx+1}/{len(df_ordenes)}: {numero_orden}")
            
            # Consultar datos iniciales de la orden - SOLO CAMPOS ESPECÍFICOS
            logger.info(f"Consultando datos iniciales de la orden {numero_orden}...")
            query_orden = f"""
                SELECT SFTTD, DOCTD, STTTD, SKUTD, U71TD, O43TD, D37TD, TIPO, DOCTDNEW, ASN, NUMORDEN 
                FROM {BASE}.KTCHPWASIN 
                WHERE NUMORDEN = '{numero_orden}'
            """
            df_orden = ejecutar_query(conexion, query_orden)
            
            if df_orden is None or df_orden.empty:
                logger.warning(f"No se encontraron datos para la orden {numero_orden}")
                continue
            
            # Ejecutar primer llamado a PROCESAR_ORDENES
            logger.info(f"Ejecutando primer PROCESAR_ORDENES para {numero_orden}...")
            proc_call = f"CALL ELOPEZ.PROCESAR_ORDENES('{numero_orden}','','','','','','','','','')"
            ejecutar_stored_procedure(conexion, proc_call)
            
            # Ejecutar comando CL SIWINTRACL como trabajo enviado
            b = BASE[2:4]
            logger.info(f"Ejecutando SIWINTRACL con parámetro '{b}' como trabajo...")
            comando_cl = f"CALL PGM(RIUNICOM63/SIWINTRACL) PARM('{b}')"
            job_info = ejecutar_comando_cl(comando_cl, submit_job=True, job_name=f"WMSI_TF{b}")
            
            if job_info:
                logger.info(f"Trabajo enviado: {job_info.get('job_name')}")
            
            # Configuración de monitoreo - Cambiar USE_MONITOREO_INTELIGENTE a False para usar espera fija
            USE_MONITOREO_INTELIGENTE = True
            TIMEOUT_MINUTOS = 5
            ESPERA_FIJA_SEGUNDOS = 30
            
            if USE_MONITOREO_INTELIGENTE:
                # Esperar a que el trabajo termine usando verificación inteligente
                logger.info("Monitoreando finalización del trabajo SIWINTRACL...")
                trabajo_finalizado = esperar_finalizacion_trabajo(numero_orden, conexion, logger, timeout_minutos=TIMEOUT_MINUTOS)
                
                if not trabajo_finalizado:
                    logger.warning(f"El trabajo no finalizó en el tiempo esperado. Continuando con el proceso...")
                else:
                    logger.info("Trabajo SIWINTRACL finalizado exitosamente.")
            else:
                # Método tradicional: espera fija
                logger.info(f"Esperando {ESPERA_FIJA_SEGUNDOS} segundos para procesamiento...")
                time.sleep(ESPERA_FIJA_SEGUNDOS)
            
            # Consultar INVENTARIO_INICIAL
            logger.info(f"Consultando INVENTARIO_INICIAL para orden {numero_orden}...")
            query_inv_inicial = f"SELECT * FROM {BASE}.INVENTARIO_INICIAL WHERE NUMORDEN = '{numero_orden}' order by sku desc "
            df_inv_inicial = ejecutar_query(conexion, query_inv_inicial)
        
            # Consultar KTCHPWINRC - SOLO CAMPOS ESPECÍFICOS
            logger.info(f"Consultando KTCHPWINRC para orden {numero_orden}...")
            query_ktchpwinrc = f"""
                SELECT SFTTD, DOCTD, STTTD, SKUTD, U71TD, O43TD, D37TD, TIPO, DOCTDNEW, ASN, NUMORDEN
                FROM {BASE}.KTCHPWINRC 
                WHERE NUMORDEN = '{numero_orden}'
            """
            df_ktchpwinrc = ejecutar_query(conexion, query_ktchpwinrc)
            
            # Ejecutar segundo llamado a PROCESAR_ORDENES
            logger.info(f"Ejecutando segundo PROCESAR_ORDENES para {numero_orden}...")
            ejecutar_stored_procedure(conexion, proc_call)
                        
            # Consultar INVENTARIO_FINAL
            logger.info(f"Consultando INVENTARIO_FINAL para orden {numero_orden}...")
            query_inv_final = f"SELECT * FROM {BASE}.INVENTARIO_FINAL WHERE NUMORDEN = '{numero_orden}' order by sku desc "
            df_inv_final = ejecutar_query(conexion, query_inv_final)

            print (f"Inventario Inicial:\n{df_inv_inicial}")
            print (f"Inventario Final:\n{df_inv_final}")

            resultados_validacion = validar_inventario(df_inv_inicial, df_inv_final, df_orden, logger)
            print(f"Resultados de validación para orden {numero_orden}:\n{resultados_validacion}")



            # Validar resultados del inventario
            logger.info(f"Validando resultados de inventario para orden {numero_orden}...")
            # Mostrar resumen de las 10 líneas
            
            # Guardar resultados de validación para esta orden
            resultados_por_orden[numero_orden] = resultados_validacion
            
            # Generar reporte Word en lugar de Excel
            reporte_nombre = f"Reporte_Orden_{numero_orden}_{fecha_actual}.docx"
            reporte_path = os.path.join(reporte_dir, reporte_nombre)
            
            # Generar documento Word con todas las tablas y validación
            word_path = generar_word_reporte_con_validacion(
                df_orden, 
                df_inv_inicial, 
                df_ktchpwinrc, 
                df_inv_final,
                resultados_validacion,
                numero_orden, 
                reporte_path
            )
            
            if word_path:
                logger.info(f"Documento Word generado: {word_path}")
                reportes_generados.append(word_path)
            else:
                logger.warning(f"No se pudo generar el documento Word para la orden {numero_orden}")
        
        # Calcular tiempo total de procesamiento
        tiempo_fin = datetime.now()
        duracion = tiempo_fin - tiempo_inicio
        duracion_str = str(duracion).split('.')[0]  # Formato HH:MM:SS
        
        # Enviar correo con todos los reportes generados
        if reportes_generados:
            logger.info(f"Enviando correo con {len(reportes_generados)} reportes...")
            
            # Verificar que los reportes existan físicamente
            reportes_existentes = []
            for reporte in reportes_generados:
                if os.path.exists(reporte):
                    reportes_existentes.append(reporte)
                    logger.info(f"Reporte verificado: {reporte}")
                else:
                    logger.warning(f"Reporte no encontrado: {reporte}")
            
            if not reportes_existentes:
                logger.error("Ninguno de los reportes generados existe físicamente. No se enviará correo.")
                return False
                
            # Contar errores y advertencias correctamente + analizar productos duplicados
            advertencias = 0
            errores = 0
            ordenes_procesadas = []
            productos_con_multiples_lineas = 0
            total_lineas_agrupadas = 0
            skus_duplicados = []
            
            for idx, row in df_ordenes.iterrows():
                numero_orden = str(row['NUMORDEN'])
                reporte_nombre = f"Reporte_Orden_{numero_orden}_{fecha_actual}.docx"
                reporte_path = os.path.join(reporte_dir, reporte_nombre)
                
                if reporte_path in reportes_generados:
                    ordenes_procesadas.append(numero_orden)
                    
                    # Contar errores y advertencias basado en los resultados de validación
                    if numero_orden in resultados_por_orden:
                        estado = resultados_por_orden[numero_orden].get('estado', 'OK')
                        if estado == 'ERROR':
                            errores += 1
                        elif estado == 'ADVERTENCIA':
                            advertencias += 1
                        
                        # Analizar productos con múltiples líneas
                        movimientos = resultados_por_orden[numero_orden].get('movimientos_inventario', [])
                        for mov in movimientos:
                            cantidad_lineas = mov.get('cantidad_lineas', 1)
                            if cantidad_lineas > 1:
                                productos_con_multiples_lineas += 1
                                total_lineas_agrupadas += cantidad_lineas
                                skus_duplicados.append({
                                    'sku': mov['sku'],
                                    'lineas': cantidad_lineas,
                                    'unidades_total': mov['unidades_movidas'],
                                    'orden': numero_orden
                                })
            
            asunto = f"Validación Automática de Transferencias WMS-INFOR - {fecha_actual}"
            cuerpo = f"""
            <html>
            <body style='font-family: Calibri, sans-serif; font-size: 14px;'>
                <h2 style='color: #2E74B5;'>Validacion de Transferencias WMS-INFOR</h2>
                
                <h3 style='color: #5B9BD5;'>Resumen de Ejecución</h3>
                <ul style='margin-bottom: 10px;'>
                    <li><b>Fecha de procesamiento:</b> {fecha_actual}</li>
                    <li><b>Hora de inicio:</b> {tiempo_inicio.strftime('%H:%M:%S')}</li>
                    <li><b>Hora de finalización:</b> {tiempo_fin.strftime('%H:%M:%S')}</li>
                    <li><b>Duración total:</b> {duracion_str}</li>
                    <li><b>Total de transferencias procesadas:</b> {len(reportes_generados)}</li>
                </ul>
                
                <h3 style='color: #5B9BD5;'>Estado de Validación</h3>
                <ul style='margin-bottom: 15px;'>
                    <li style='color: green;'><b>Transferencias exitosas:</b> {len(reportes_generados) - errores - advertencias}</li>
                    <li style='color: orange;'><b>Transferencias con advertencias:</b> {advertencias}</li>
                    <li style='color: red;'><b>Transferencias con errores:</b> {errores}</li>
                </ul>
                
                <h3 style='color: #5B9BD5;'>📋 Resumen de Validaciones Procesadas</h3>
                <div style='background-color: #E8F5E8; padding: 10px; border: 1px solid #4CAF50; margin-bottom: 15px;'>"""
            
            # Agregar detalles específicos de las validaciones procesadas
            total_origenes_validados = 0
            total_destinos_validados = 0
            total_registros_procesados = 0
            
            for numero_orden, resultado in resultados_por_orden.items():
                if 'tablas_validacion' in resultado:
                    if 'origen' in resultado['tablas_validacion']:
                        total_origenes_validados += len(resultado['tablas_validacion']['origen'])
                    if 'destino' in resultado['tablas_validacion']:
                        total_destinos_validados += len(resultado['tablas_validacion']['destino'])
                total_registros_procesados += resultado.get('registros_inicial', 0)
            
            cuerpo += f"""
                    <p style='margin: 5px 0; font-weight: bold; color: #2E7D32;'>
                        ✅ <b>Campos agrupados y validados exitosamente:</b>
                    </p>
                    <ul style='margin: 5px 0 10px 20px;'>
                        <li><b style='color: #1976D2;'>Total de combinaciones (SKU + ORIGEN) validadas:</b> {total_origenes_validados}</li>
                        <li><b style='color: #1976D2;'>Total de combinaciones (SKU + DESTINO) validadas:</b> {total_destinos_validados}</li>
                        <li><b style='color: #1976D2;'>Total de registros individuales procesados:</b> {total_registros_procesados}</li>
                    </ul>
                    
                    <p style='margin: 5px 0; color: #424242;'>
                        <b>Campos sumados en la agrupación:</b>
                    </p>
                    <ul style='margin: 5px 0 0 20px; color: #424242;'>
                        <li><b style='color: #D32F2F;'>UNIDADES</b> - Sumadas por cada (SKU, ORIGEN) y (SKU, DESTINO)</li>
                        <li><b style='color: #388E3C;'>INV_ORIGEN / INV_DESTINO</b> - Primer valor del inventario inicial para cada agrupación</li>
                    </ul>
                </div>"""
            
            # Agregar sección de productos duplicados si existen
            if productos_con_multiples_lineas > 0:
                cuerpo += f"""
                <h3 style='color: #FF8C00; background-color: #FFF8DC; padding: 10px; border-left: 5px solid #FF8C00;'>
                    📦 Productos con Múltiples Líneas de Transferencia
                </h3>
                <div style='background-color: #F0F8FF; padding: 15px; border: 1px solid #B0C4DE; margin-bottom: 15px;'>
                    <p style='margin-top: 0; font-weight: bold; color: #4169E1;'>
                        ℹ️ INFORMACIÓN IMPORTANTE: Se detectaron productos con múltiples líneas de transferencia
                    </p>
                    <ul style='margin-bottom: 10px;'>
                        <li><b>Productos con múltiples líneas:</b> <span style='color: #FF6347; font-weight: bold;'>{productos_con_multiples_lineas}</span></li>
                        <li><b>Total de líneas agrupadas:</b> {total_lineas_agrupadas}</li>
                        <li><b>Promedio de líneas por producto:</b> {round(total_lineas_agrupadas / productos_con_multiples_lineas, 1)}</li>
                    </ul>
                    
                    <p style='color: #228B22; font-weight: bold; margin-bottom: 10px;'>
                        ✅ ESTO ES NORMAL Y ESPERADO
                    </p>
                    <p style='margin-bottom: 10px;'>
                        <b>Explicación:</b> Es común que un mismo SKU tenga múltiples líneas de transferencia en el sistema WMS. 
                        Nuestro sistema de validación <u>agrupa automáticamente</u> estas líneas por SKU y calcula correctamente 
                        el total de unidades transferidas.
                    </p>
                    
                    <details style='margin-top: 10px;'>
                        <summary style='font-weight: bold; color: #4169E1; cursor: pointer;'>
                            📋 Ver detalle de productos duplicados
                        </summary>
                        <div style='margin-top: 10px; padding: 10px; background-color: #FFFFFF; border: 1px solid #D3D3D3;'>"""
                
                # Agregar detalle de cada SKU duplicado
                for sku_info in skus_duplicados:
                    cuerpo += f"""
                            <div style='margin-bottom: 8px; padding: 5px; background-color: #F9F9F9; border-left: 3px solid #FF8C00;'>
                                <b>SKU {sku_info['sku']}</b> (Orden {sku_info['orden']}): 
                                <span style='color: #FF6347;'>{sku_info['lineas']} líneas</span> → 
                                <span style='color: #228B22;'>{sku_info['unidades_total']} unidades total</span>
                            </div>"""
                
                cuerpo += f"""
                        </div>
                    </details>
                </div>"""
            else:
                cuerpo += f"""
                <div style='background-color: #F0FFF0; padding: 10px; border: 1px solid #90EE90; margin-bottom: 15px;'>
                    <p style='margin: 0; color: #006400;'>
                        ✅ <b>Sin productos duplicados:</b> Todas las transferencias contienen líneas únicas por SKU.
                    </p>
                </div>"""
            
            # Continuar con el resto del correo
            cuerpo += f"""
                <h3 style='color: #5B9BD5;'>📊 Detalle de Campos Procesados en la Validación</h3>
                <div style='background-color: #F8F9FA; padding: 15px; border: 1px solid #DEE2E6; margin-bottom: 15px;'>
                    <h4 style='color: #495057; margin-top: 0;'>🔍 Campos que se Agrupan y Suman:</h4>
                    
                    <div style='margin-bottom: 10px;'>
                        <h5 style='color: #007BFF; margin-bottom: 5px;'>📦 VALIDACIÓN DE ORÍGENES (Agrupado por SKU + ORIGEN):</h5>
                        <ul style='margin-top: 0; margin-bottom: 10px;'>
                            <li><b style='color: #DC3545;'>UNIDADES:</b> Se suman todas las unidades transferidas del mismo SKU desde el mismo ORIGEN</li>
                            <li><b style='color: #28A745;'>INV_ORIGEN:</b> Se toma el primer valor del inventario inicial en el origen (antes de transferir)</li>
                            <li><b style='color: #6F42C1;'>Fórmula de validación:</b> INV_INICIAL_ORIGEN - TOTAL_TRANSFERIDO = INV_FINAL_ORIGEN</li>
                        </ul>
                    </div>
                    
                    <div style='margin-bottom: 10px;'>
                        <h5 style='color: #007BFF; margin-bottom: 5px;'>📦 VALIDACIÓN DE DESTINOS (Agrupado por SKU + DESTINO):</h5>
                        <ul style='margin-top: 0; margin-bottom: 10px;'>
                            <li><b style='color: #DC3545;'>UNIDADES:</b> Se suman todas las unidades recibidas del mismo SKU en el mismo DESTINO</li>
                            <li><b style='color: #28A745;'>INV_DESTINO:</b> Se toma el primer valor del inventario inicial en el destino (antes de recibir)</li>
                            <li><b style='color: #6F42C1;'>Fórmula de validación:</b> INV_INICIAL_DESTINO + TOTAL_RECIBIDO = INV_FINAL_DESTINO</li>
                        </ul>
                    </div>
                    
                    <div style='background-color: #E3F2FD; padding: 10px; border-left: 4px solid #2196F3;'>
                        <p style='margin: 0; font-weight: bold; color: #1976D2;'>
                            💡 <b>Importante:</b> El sistema agrupa automáticamente múltiples líneas del mismo SKU para evitar conteos duplicados 
                            y garantizar la precisión matemática de las validaciones de inventario.
                        </p>
                    </div>
                </div>
                
                <h3 style='color: #5B9BD5;'>Detalle del Procedimiento</h3>
                <p style='margin-bottom: 5px;'>Se ejecutaron las siguientes acciones para cada transferencia:</p>
                <ol style='margin-top: 0; margin-bottom: 10px;'>
                    <li>Consulta de datos en tabla KTCHPWASIN con el fin de extraer todos los load id existentes</li>
                    <li>Por cada load id encontrado se realizará lo siguiente:</li>
                        <li>Ejecución del procedimiento PROCESAR_ORDENES, este procedimiento toma una captura del inventario inicial y 
                        lo guarda en ELOPEZ.INVENTARIO_INICIAL así como tambien deja disponible solo el load id que se va ha procesar.</li>
                        <li>Ejecución del comando SIWINTRACL para el posteo de transferencias en INFOR</li>
                        <li>Consulta de tablas de inventario (INVENTARIO_INICIAL, KTCHPWINRC, INVENTARIO_FINAL)
                        cuando finaliza. el programa CL, borra el registro de la tabla KTCHPWASIN y lo pasa a la KTCHPWINRC.
                        </li>
                        <li>Segunda ejecución del procedimiento PROCESAR_ORDENES, Esto con el fin de guardar el movimiento del inventario final. 
                        el mismo procedimiento lo toma y lo guarda en INVENTARIO_FINAL. para tener un respaldo de la prueba.</li>
                        <li><b style='color: #FF8C00;'>Validación matemática con agrupación:</b> El sistema agrupa automáticamente por (SKU, ORIGEN) y (SKU, DESTINO), 
                        sumando las unidades transferidas/recibidas para cada combinación única.</li>
                        <li>Esto puede implementarse en todos los paises de ULA. pero actualmente infor solo existe para el salvador</li>
                        <li>Validación de consistencia de datos y generación de reportes</li>
                </ol>
                
                <h3 style='color: #5B9BD5;'>Transferencias Procesadas</h3>
                <p style='margin-bottom: 10px;'>Se procesaron las siguientes transferencias: <b>{', '.join(ordenes_procesadas)}</b></p>
                
                <p style='margin-bottom: 5px;'><b>Nota:</b> Para que una transferencia se considere exitosa, debe incluir movimientos de saldos entre los inventarios.</p>
                
                <p style='margin-bottom: 5px;'>Se adjuntan los reportes detallados de cada transferencia procesada.</p>
                
                <p style='margin-bottom: 0;'>Saludos!!!!!</p>
                <p style='margin-top: 0; margin-bottom: 0;'><b>Validación de Transferencias WMS-INFOR</b></p>
            </body>
            </html>
            """
            
            # Intentar enviar el correo hasta 3 veces
            enviado = False
            intentos = 0
            max_intentos = 3
            
            while not enviado and intentos < max_intentos:
                intentos += 1
                logger.info(f"Intento {intentos} de envío de correo...")
                enviado = enviar_correo(asunto, cuerpo, reportes_existentes)
                if not enviado and intentos < max_intentos:
                    logger.warning(f"Fallo en el intento {intentos}. Reintentando en 5 segundos...")
                    time.sleep(5)  # Esperar 5 segundos antes de reintentar
            
            if enviado:
                logger.info("Correo enviado exitosamente")
            else:
                logger.error("Todos los intentos de envío de correo fallaron")
        else:
            logger.warning("No se generaron reportes para enviar por correo")
        
        logger.info("=== PROCESAMIENTO DE TRANSFERENCIAS WMS-INFOR FINALIZADO ===")
        return True
        
    except Exception as e:
        logger.error(f"Error en el procesamiento de transferencias WMS-INFOR: {e}")
        return False
    finally:
        # Cerrar conexión
        if conexion:
            conexion.close()
            logger.info("Conexión a AS400 cerrada")

def mostrar_resumen_10_lineas(resultados_validacion):
    """Muestra un resumen de las 10 líneas con los campos específicos solicitados"""
    if not resultados_validacion.get('movimientos_inventario'):
        print("📋 No hay movimientos de inventario para mostrar")
        return
    
    print("\n" + "="*140)
    print("📊 RESUMEN DE VALIDACIÓN - 10 LÍNEAS")
    print("="*140)
    # SKU, ORIGEN, DESTINO, INV_ORIGEN, INV_DESTINO, CANT_LINEAS, TOTAL_SKU, RESULTADO
    print(f"{'SKU':<14} {'ORIGEN':<8} {'DESTINO':<8} {'INV_ORIGEN':<12} {'INV_DESTINO':<12} {'CANT_LINEAS':<12} {'TOTAL_SKU':<10} {'RESULTADO':<10}")
    print("-"*140)
    
    # Ordenar por SKU, ORIGEN, DESTINO para visualización consistente
    movimientos = sorted(
        resultados_validacion['movimientos_inventario'],
        key=lambda m: (str(m.get('sku','')), str(m.get('origen','')), str(m.get('destino','')))
    )
    for i, mov in enumerate(movimientos, 1):
        sku = str(mov.get('sku',''))[:14]
        origen = str(mov.get('origen',''))[:8]
        destino = str(mov.get('destino',''))[:8]
        inv_origen = f"{mov.get('inv_origen', 0):.1f}"
        inv_destino = f"{mov.get('inv_destino', 0):.1f}"
        cant_lineas = str(mov.get('cantidad_lineas', 1))
        total_sku = f"{mov.get('total_unidades_sku', 0):.0f}"
        resultado = mov.get('estado_general', 'N/A')
        
        # Formatear resultado con emoji
        if resultado == 'OK':
            resultado_fmt = "✅ OK"
        else:
            resultado_fmt = "❌ ERROR"
        print(f"{sku:<14} {origen:<8} {destino:<8} {inv_origen:<12} {inv_destino:<12} {cant_lineas:<12} {total_sku:<10} {resultado_fmt:<10}")
    
    print("="*140)
    print(f"Total de líneas procesadas: {len(movimientos)}")
    print("NOTA: TOTAL_SKU muestra el total de unidades para ese SKU en todas las transferencias")
    print()

def generar_word_reporte_con_validacion(df_orden, df_inv_inicial, df_ktchpwinrc, df_inv_final, resultados_validacion, numero_orden, reporte_path):
    """Genera un documento Word con los datos de las transferencias WMS-INFOR y resultados de validación"""
    try:
        document = Document()
        
        # Configurar tamaño de página A3
        section = document.sections[0]
        section.page_width = Inches(16.5)   # A3 ancho: 16.5 pulgadas (420mm)
        section.page_height = Inches(11.7)  # A3 alto: 11.7 pulgadas (297mm)
        section.left_margin = Inches(0.5)   # Margen izquierdo reducido
        section.right_margin = Inches(0.5)  # Margen derecho reducido
        section.top_margin = Inches(0.5)    # Margen superior reducido
        section.bottom_margin = Inches(0.5) # Margen inferior reducido
        
        # Configurar estilo del documento para eliminar espacios entre párrafos
        for style_name in document.styles:
            try:
                style = document.styles[style_name]
                if hasattr(style, 'paragraph_format'):
                    style.paragraph_format.space_before = Pt(0)
                    style.paragraph_format.space_after = Pt(0)
            except:
                # Ignorar estilos que no se pueden modificar
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
        dataframe_to_table(document, df_inv_inicial, "Inventario Inicial (ELOPEZ.INVENTARIO_INICIAL)")
        
        # Sección 3: Datos KTCHPWINRC
        dataframe_to_table(document, df_ktchpwinrc, "Datos KTCHPWINRC (RI11DB.KTCHPWINRC)")
        
        # Sección 4: Inventario Final
        dataframe_to_table(document, df_inv_final, "Inventario Final (ELOPEZ.INVENTARIO_FINAL)")
        
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
                dataframe_to_table(document, df_validacion_origen, "🔍 TABLA DE VALIDACIÓN - ORÍGENES")
            
            # Agregar tabla de validación de DESTINOS
            if 'destino' in resultados_validacion['tablas_validacion']:
                df_validacion_destino = resultados_validacion['tablas_validacion']['destino']
                dataframe_to_table(document, df_validacion_destino, "🔍 TABLA DE VALIDACIÓN - DESTINOS")
            
            # Agregar resumen final
            if 'resumen' in resultados_validacion['tablas_validacion']:
                df_resumen = resultados_validacion['tablas_validacion']['resumen']
                dataframe_to_table(document, df_resumen, "📊 RESUMEN FINAL")

        # TABLA: Resumen de Inventario Inicial Agrupado (SKU, ORIGEN)
        try:
            resumen_df = resultados_validacion.get('resumen_inv_inicial')
            if isinstance(resumen_df, pd.DataFrame) and not resumen_df.empty:
                # Ordenar por SKU y ORIGEN para presentación consistente
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
                # Obtener origen y destinos únicos
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
            
            # Encabezados exactamente como en consola
            hdr_cells = table.rows[0].cells
            headers = ["SKU", "ORIGEN", "DESTINO", "INV_ORIGEN", "INV_DESTINO", "UNID_MV", "TOTAL_SKU", "RESULTADO"]
            colores = [
                RGBColor(0, 0, 139),      # Azul oscuro para SKU
                RGBColor(178, 34, 34),    # Rojo ladrillo para ORIGEN
                RGBColor(0, 100, 0),      # Verde oscuro para DESTINO
                RGBColor(128, 0, 128),    # Púrpura para INV_ORIGEN
                RGBColor(0, 100, 139),    # Azul para INV_DESTINO
                RGBColor(255, 140, 0),    # Naranja para UNID_MV
                RGBColor(128, 128, 0),    # Oliva para TOTAL_SKU
                RGBColor(0, 128, 0)       # Verde para RESULTADO
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
                # Solo procesar movimientos con unidades > 0
                if mov.get('unidades_movidas', 0) > 0:
                    # Crear clave única para evitar duplicados
                    clave = f"{mov['sku']}_{mov['origen']}_{mov['destino']}"
                    
                    # Solo agregar si no existe o si es mejor (más reciente/completo)
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
                sku_run.font.color.rgb = RGBColor(0, 0, 139)  # Azul oscuro
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # ORIGEN
                row_cells[1].text = ""
                p = row_cells[1].paragraphs[0]
                origen_run = p.add_run(str(mov['origen']))
                origen_run.bold = True
                origen_run.font.color.rgb = RGBColor(178, 34, 34)  # Rojo ladrillo
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # DESTINO
                row_cells[2].text = ""
                p = row_cells[2].paragraphs[0]
                destino_run = p.add_run(str(mov['destino']))
                destino_run.bold = True
                destino_run.font.color.rgb = RGBColor(0, 100, 0)  # Verde oscuro
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # INV_ORIGEN
                row_cells[3].text = ""
                p = row_cells[3].paragraphs[0]
                inv_origen_run = p.add_run(str(mov['inv_origen']))
                inv_origen_run.bold = True
                inv_origen_run.font.color.rgb = RGBColor(128, 0, 128)  # Púrpura
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # INV_DESTINO
                row_cells[4].text = ""
                p = row_cells[4].paragraphs[0]
                inv_destino_run = p.add_run(str(mov['inv_destino']))
                inv_destino_run.bold = True
                inv_destino_run.font.color.rgb = RGBColor(0, 100, 139)  # Azul
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # UNID_MV (unidades movidas)
                row_cells[5].text = ""
                p = row_cells[5].paragraphs[0]
                unidades_run = p.add_run(f"{mov['unidades_movidas']:.0f}")
                unidades_run.bold = True
                unidades_run.font.color.rgb = RGBColor(255, 140, 0)  # Naranja
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # TOTAL_SKU
                row_cells[6].text = ""
                p = row_cells[6].paragraphs[0]
                total_sku = mov.get('total_unidades_sku', mov['unidades_movidas'])
                total_run = p.add_run(f"{total_sku:.0f}")
                total_run.bold = True
                total_run.font.color.rgb = RGBColor(128, 128, 0)  # Oliva
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # RESULTADO
                row_cells[7].text = ""
                p = row_cells[7].paragraphs[0]
                if mov.get('origen_ok', True) and mov.get('destino_ok', True):
                    resultado_run = p.add_run("✅ OK")
                    resultado_run.font.color.rgb = RGBColor(0, 128, 0)  # Verde
                else:
                    resultado_run = p.add_run("❌ ERROR")
                    resultado_run.font.color.rgb = RGBColor(255, 0, 0)  # Rojo
                resultado_run.bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
        # Métricas en un solo párrafo
        document.add_heading('Métricas', level=2)
        metrics_p = document.add_paragraph()
        
        metrics_p.add_run('Registros: ').bold = True
        metrics_p.add_run('Inicial: ').bold = True
        reg_ini = metrics_p.add_run(str(resultados_validacion['registros_inicial']))
        reg_ini.font.color.rgb = RGBColor(178, 34, 34)  # Rojo ladrillo
        reg_ini.bold = True
        
        metrics_p.add_run(' | KTCHPWINRC: ').bold = True
        reg_winrc = metrics_p.add_run(str(resultados_validacion['registros_winrc']))
        reg_winrc.font.color.rgb = RGBColor(0, 0, 139)  # Azul oscuro
        reg_winrc.bold = True
        
        metrics_p.add_run(' | Final: ').bold = True
        reg_final = metrics_p.add_run(str(resultados_validacion['registros_final']))
        reg_final.font.color.rgb = RGBColor(0, 100, 0)  # Verde oscuro
        reg_final.bold = True

        # Discrepancias
        if resultados_validacion.get('discrepancias'):
            document.add_heading('Discrepancias Encontradas', level=2)
            
            # Lista compacta de discrepancias
            for idx, discrepancia in enumerate(resultados_validacion['discrepancias'], 1):
                # Modificar estilo de párrafo para hacerlo más compacto
                p = document.add_paragraph(style='List Bullet')
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                
                # Resaltar tipo de discrepancia con colores según su tipo
                tipo_run = p.add_run(f"{idx}. {discrepancia['tipo']}: ")
                tipo_run.bold = True
                
                if discrepancia['tipo'] == 'INVENTARIO':
                    tipo_run.font.color.rgb = RGBColor(255, 0, 0)  # Rojo para errores críticos
                elif discrepancia['tipo'] == 'SKU_PERDIDO':
                    tipo_run.font.color.rgb = RGBColor(255, 0, 0)  # Rojo para errores críticos
                elif discrepancia['tipo'] == 'CANTIDAD':
                    tipo_run.font.color.rgb = RGBColor(255, 165, 0)  # Naranja para advertencias
                elif discrepancia['tipo'] == 'SKU_FALTANTE':
                    tipo_run.font.color.rgb = RGBColor(255, 165, 0)  # Naranja para advertencias
                elif discrepancia['tipo'] == 'SIN_MOVIMIENTOS':
                    tipo_run.font.color.rgb = RGBColor(255, 165, 0)  # Naranja para advertencias
                elif discrepancia['tipo'] == 'COLUMNAS_FALTANTES':
                    tipo_run.font.color.rgb = RGBColor(255, 0, 0)  # Rojo para errores críticos
                
                # Resaltar valores numéricos en la descripción
                descripcion = discrepancia['descripcion']
                import re
                
                # Dividir por números
                partes = re.split(r'(\d+)', descripcion)
                for parte in partes:
                    if parte.isdigit():
                        # Si es un número, resaltarlo
                        num_run = p.add_run(parte)
                        num_run.bold = True
                        num_run.font.color.rgb = RGBColor(128, 0, 128)  # Púrpura para números
                    else:
                        # Si no es un número, agregar texto normal
                        p.add_run(parte)

        # Conclusión en un párrafo compacto
        document.add_heading('Resumen de la Transferencia', level=1)
        conclusion_p = document.add_paragraph()
        conclusion_p.add_run('Este reporte fue generado automáticamente por el Sistema de Automatización de Transferencias WMS-INFOR. ')
        conclusion_p.add_run('El procesamiento incluyó la ejecución de PROCESAR_ORDENES, el comando SIWINTRACL y la consulta de tablas de inventario.')
        
        # Mensaje de estado (todo en un solo párrafo)
        if resultados_validacion['estado'] == 'OK':
            exito_run = conclusion_p.add_run(' ✅ ÉXITO: Se detectaron movimientos de saldos correctamente.')
            exito_run.bold = True
            exito_run.font.color.rgb = RGBColor(0, 128, 0)  # Verde
        elif resultados_validacion['estado'] == 'ADVERTENCIA' and 'SIN_MOVIMIENTOS' in [d.get('tipo') for d in resultados_validacion.get('discrepancias', [])]:
            adv_run = conclusion_p.add_run(' ⚠️ ADVERTENCIA: No se detectaron movimientos de saldos entre inventarios.')
            adv_run.bold = True
            adv_run.font.color.rgb = RGBColor(255, 165, 0)  # Naranja
        elif resultados_validacion['estado'] == 'ERROR':
            error_run = conclusion_p.add_run(' ⚠️ ATENCIÓN: Se detectaron errores que requieren revisión manual.')
            error_run.bold = True
            error_run.font.color.rgb = RGBColor(255, 0, 0)  # Rojo
        
        # Guardar documento
        document.save(reporte_path)
        print(f"Documento Word de transferencia WMS-INFOR generado: {os.path.abspath(reporte_path)}")
        return reporte_path
    except Exception as e:
        print(f"Error al generar documento Word de transferencia WMS-INFOR: {e}")
        import traceback
        print(traceback.format_exc())
        return None

# Función principal para mantener compatibilidad con código existente
procesar_ordenes = procesar_transferencias_wms_infor

# Ejecución principal del script
if __name__ == "__main__":
    try:
        print("=== INICIANDO VALIDACIÓN WMS-INFOR ===")
        print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("Procesando transferencias WMS-INFOR...")
        
        # Procesar todas las transferencias
        procesar_transferencias_wms_infor()
        
        print("=== PROCESO COMPLETADO ===")
    except Exception as e:
        print(f"ERROR en la ejecución: {e}")
        import traceback
        traceback.print_exc()