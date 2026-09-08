# ============================================================
# main.py
# Conciliación AS400 KTRHP vs KPUHP
# Regla de negocio:
# - Llave única: SKU
# - KTRHP: SUM(U71UA) por SKU
# - KPUHP: SUM(U45UH) por SKU
# - Reporte Word con 3 tablas
# ============================================================

import sys
import os

# Asegurar que el directorio raíz del proyecto esté en el PYTHONPATH para las importaciones
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from conexion_config.conexion import ConexionAS400
from datetime import datetime, timedelta
from conexion_config.logging_config import get_logger
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utilidades.notificaciones_correo import NotificadorCorreo
from jira.jira_utilidades import JiraClient

# ============================================================
# CONFIGURACIÓN DE LOGGER
# ============================================================

logger = get_logger(__file__)

# ============================================================
# UTILIDADES
# ============================================================

def obtener_fecha_proceso(dias_atras: int = 0) -> str:
    """
    Retorna la fecha del proceso en formato YYYYMMDD.
    Por defecto usa el día actual (dias_atras=0) porque las transferencias
    y las POs se generan en el mismo día.
    """
    ahora = datetime.now()
        
    fecha = (ahora - timedelta(days=dias_atras)).strftime("%Y%m%d")

    if not fecha.isdigit() or len(fecha) != 8:
        raise ValueError(f"Fecha inválida generada: {fecha}")

    return fecha


def ejecutar_consulta(query: str, params: tuple, nombre_proceso: str) -> list:
    """
    Ejecuta una consulta SQL en AS400.
    Retorna lista de tuplas (nunca None).
    """
    conexion = None
    try:
        logger.info(f"Iniciando consulta [{nombre_proceso}]")
        conexion = ConexionAS400().conectar()

        with conexion.cursor() as cursor:
            cursor.execute(query, params)
            resultados = cursor.fetchall()

        logger.info(
            f"Consulta [{nombre_proceso}] finalizada. "
            f"Registros obtenidos: {len(resultados)}"
        )
        return resultados

    except Exception as e:
        logger.exception(f"Error ejecutando [{nombre_proceso}]: {e}")
        return []

    finally:
        if conexion:
            try:
                conexion.close()
            except Exception:
                pass

# ============================================================
# EXTRACCIÓN DE DATOS
# ============================================================

def extraer_ktrhp(base: str, fecha: str) -> list:
    """Extrae totales KTRHP por SKU (U71UA) y el Owner (O43UA)"""
    query = f"SELECT SKUUA, O43UA, SUM(U71UA) AS TOTAL_U71UA FROM {base}.KTRHP WHERE D88UA = ? AND C86UA = 'R' GROUP BY SKUUA, O43UA"
    return ejecutar_consulta(query, (fecha,), "KTRHP_POR_SKU")

def extraer_kpuhp(base: str, fecha: str) -> list:
    """Extrae totales KPUHP por SKU (U45UH) dinámico y su PO (POÑUH)"""
    transpo = f"TRANSPO{base[2:4]}"
    query = f"SELECT SKUUH, POÑUH AS PO, SUM(U45UH) AS TOTAL_U45UH FROM {base}.KPUHP WHERE D88UH = ? AND JBDUH = ? GROUP BY SKUUH, POÑUH"
    return ejecutar_consulta(query, (fecha, transpo), "KPUHP_POR_SKU")

def actualizar_registros_dia_anterior(base: str, fecha: str, transpo_dinamico: str):
    """Actualiza las fechas de KTRHP y KPUHP enviando los registros procesados 1 día atrás."""
    try:
        fecha_obj = datetime.strptime(fecha, "%Y%m%d")
        fecha_nueva = (fecha_obj - timedelta(days=1)).strftime("%Y%m%d")
        
        query_ktrhp = f"UPDATE {base}.KTRHP SET D88UA = '{fecha_nueva}' WHERE D88UA = '{fecha}'"
        query_kpuhp = f"UPDATE {base}.KPUHP SET D88UH = '{fecha_nueva}' WHERE D88UH = '{fecha}' AND JBDUH = '{transpo_dinamico}'"
        
        logger.info(f"Actualizando registros a 1 día atrás (De {fecha} a {fecha_nueva})...")
        conexion = ConexionAS400().conectar()
        with conexion.cursor() as cursor:
            cursor.execute(query_ktrhp)
            cursor.execute(query_kpuhp)
        conexion.commit()
        conexion.close()
        logger.info("Registros actualizados correctamente.")
    except Exception as ex:
        logger.error(f"Error al actualizar las fechas: {ex}")

# ============================================================
# NORMALIZACIÓN A DICCIONARIOS
# ============================================================

def procesar_ktrhp(datos: list, base: str) -> tuple:
    """Separa KTRHP en válidas e ignoradas según O43UA dinámicamente basado en la base"""
    validas = {}
    ignoradas = {}
    owner_dinamico = base[2:4] # Extrae '12' de 'RI12DB'
    for sku, o43ua, total in datos:
        total = total or 0
        o43ua = str(o43ua).strip().upper() if o43ua else ""
        if o43ua == owner_dinamico:
            ignoradas[sku] = ignoradas.get(sku, 0) + total
        else:
            validas[sku] = validas.get(sku, 0) + total
    return validas, ignoradas

def dict_kpuhp_por_sku(datos: list) -> dict:
    return {sku: {"po": po, "total": total or 0} for sku, po, total in datos}

# ============================================================
# ANÁLISIS / CONCILIACIÓN
# ============================================================

def analizar_por_sku(ktrhp_validas: dict, kpuhp: dict) -> list:
    """Compara totales por SKU entre KTRHP Válidas y KPUHP"""
    resultado = []
    skus_totales = set(ktrhp_validas.keys()).union(set(kpuhp.keys()))
    for sku in skus_totales:
        total_ktrhp = ktrhp_validas.get(sku, 0)
        kpuhp_data = kpuhp.get(sku, {"po": "N/A", "total": 0})
        total_kpuhp = kpuhp_data["total"]
        po = kpuhp_data["po"]
        diferencia = total_ktrhp - total_kpuhp
        estado = "SIN DATOS" if sku not in kpuhp else "OK" if diferencia == 0 else "ERROR"
        resultado.append({"PO": po, "SKU": sku, "TOTAL_KTRHP": total_ktrhp, "TOTAL_KPUHP": total_kpuhp, "DIFERENCIA": diferencia, "ESTADO": estado})
    return resultado

# ============================================================
# REPORTE WORD (3 TABLAS)
# ============================================================

def generar_reporte_word(ktrhp_validas: dict, ktrhp_ignoradas: dict, kpuhp_por_sku: dict, analisis: list, fecha: str, base: str, ruta_salida: str):
    doc = Document()
    
    # --- ESTILOS BÁSICOS ---
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    # --- TÍTULO ---
    title = doc.add_heading("Reporte Ejecutivo de Conciliación AS400", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_info.add_run(f"Proceso: PO Vendor Storage | Fecha: {fecha}").bold = True
    
    # --- RESUMEN EJECUTIVO ---
    doc.add_heading("📊 Resumen:", level=1)
    
    todos_ok = all(r["ESTADO"] == "OK" for r in analisis)
    p_estado = doc.add_paragraph()
    p_estado.add_run("Estado Global de la Conciliación: ").bold = True
    if todos_ok and analisis:
        run_ok = p_estado.add_run("✅ CUADRE PERFECTO")
        run_ok.font.color.rgb = RGBColor(0, 128, 0)
    else:
        run_err = p_estado.add_run("❌ DIFERENCIAS DETECTADAS")
        run_err.font.color.rgb = RGBColor(255, 0, 0)
        
    doc.add_paragraph(f"Se procesaron {len(analisis)} SKUs en total.")
    
    # --- ANÁLISIS POR SKU (La parte más importante) ---
    doc.add_heading("📦 Análisis de Conciliación por SKU", level=1)
    doc.add_paragraph("Comparativa directa entre las unidades ordenadas a generar (KTRHP válidas) y las unidades finalmente generadas (KPUHP).")
    
    t4 = doc.add_table(rows=1, cols=6)
    t4.style = 'Medium Grid 1 Accent 1'
    encabezados = ["PO", "SKU", "Ordenadas (KTRHP)", "Generadas (KPUHP)", "Diferencia", "Estado"]
    for i, title in enumerate(encabezados):
        celda = t4.rows[0].cells[i]
        celda.text = title
        celda.paragraphs[0].runs[0].bold = True
        
    for r in analisis:
        row = t4.add_row().cells
        row[0].text = str(r.get("PO", "N/A"))
        row[1].text = str(r["SKU"])
        row[2].text = str(r["TOTAL_KTRHP"])
        row[3].text = str(r["TOTAL_KPUHP"])
        row[4].text = str(r["DIFERENCIA"])
        
        estado = r["ESTADO"]
        if estado == "OK":
            row[5].text = "✅ OK"
            row[5].paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 128, 0)
        else:
            row[5].text = f"❌ {estado}"
            row[5].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 0, 0)

    doc.add_page_break()

    # --- DETALLES TÉCNICOS ---
    doc.add_heading("📑 Anexo Técnico (Detalle de Bases de Datos)", level=1)
    
    owner_dinamico = base[2:4]
    valid_owner = f"V{base[3]}"

    doc.add_heading(f"1. KTRHP A Generar (Owner {valid_owner} o Vacío)", level=2)
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = 'Light Shading Accent 1'
    t1.rows[0].cells[0].text, t1.rows[0].cells[1].text = "SKU", "TOTAL U71UA"
    for sku, total in ktrhp_validas.items():
        row = t1.add_row().cells
        row[0].text, row[1].text = str(sku), str(total)

    doc.add_heading(f"2. KTRHP Ignoradas (Owner {owner_dinamico})", level=2)
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = 'Light Shading Accent 1'
    t2.rows[0].cells[0].text, t2.rows[0].cells[1].text = "SKU", "TOTAL U71UA"
    for sku, total in ktrhp_ignoradas.items():
        row = t2.add_row().cells
        row[0].text, row[1].text = str(sku), str(total)

    doc.add_heading("3. Resultados KPUHP (Generados)", level=2)
    t3 = doc.add_table(rows=1, cols=3)
    t3.style = 'Light Shading Accent 1'
    t3.rows[0].cells[0].text, t3.rows[0].cells[1].text, t3.rows[0].cells[2].text = "PO", "SKU", "TOTAL U45UH"
    for sku, info in kpuhp_por_sku.items():
        row = t3.add_row().cells
        row[0].text, row[1].text, row[2].text = str(info["po"]), str(sku), str(info["total"])

    doc.save(ruta_salida)
    logger.info(f"Reporte Word Ejecutivo generado: {ruta_salida}")

# ============================================================
# MAIN
# ============================================================

def main_vendor(BASE="RI14DB", fecha_forzada=None):
    logger.info("=== INICIO CONCILIACIÓN AS400 ===")
    fecha = fecha_forzada if fecha_forzada else obtener_fecha_proceso()
    owner_dinamico = BASE[2:4]
    valid_owner = f"V{BASE[3]}"
    transpo_dinamico = f"TRANSPO{owner_dinamico}"

    datos_ktrhp = extraer_ktrhp(BASE, fecha)
    ktrhp_validas, ktrhp_ignoradas = procesar_ktrhp(datos_ktrhp, BASE)
    
    kpuhp_por_sku = dict_kpuhp_por_sku(extraer_kpuhp(BASE, fecha))
    analisis = analizar_por_sku(ktrhp_validas, kpuhp_por_sku)

    if not analisis:
        logger.info(f"No se encontraron datos de conciliación para la fecha {fecha}. Omitiendo reporte, Jira y correo.")
        return

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    carpeta_reportes = os.path.join(base_dir, 'Reportes', 'PO AUTOMATICA')
    os.makedirs(carpeta_reportes, exist_ok=True)

    ruta_reporte = os.path.join(carpeta_reportes, f"Reporte_Conciliacion_SKU_{fecha}_{BASE}.docx")
    generar_reporte_word(ktrhp_validas, ktrhp_ignoradas, kpuhp_por_sku, analisis, fecha, BASE, ruta_reporte)
    
    # --- CREACIÓN DE TAREA EN JIRA ---
    try:
        jira = JiraClient()
        # Usamos la fecha como loadid para el ticket
        jira.main_jira(fecha, os.path.abspath(ruta_reporte), "vendor_storage", BASE)
    except Exception as e:
        logger.error(f"Error al generar tarea en Jira: {e}")
    
    # Envío de correo electrónico
    errores = sum(1 for r in analisis if r["ESTADO"] != "OK")
    estado_correo = "CON INCIDENCIAS" if errores > 0 else "EXITOSO"
    asunto = f"[VENDOR STORAGE] Conciliación PO - {fecha} ({estado_correo})"
    
    query_ktrhp = f"SELECT SKUUA, O43UA, SUM(U71UA) AS TOTAL_U71UA FROM {BASE}.KTRHP WHERE D88UA = '{fecha}' AND C86UA = 'R' GROUP BY SKUUA, O43UA"
    query_kpuhp = f"SELECT SKUUH, POÑUH AS PO, SUM(U45UH) AS TOTAL_U45UH FROM {BASE}.KPUHP WHERE D88UH = '{fecha}' AND JBDUH = '{transpo_dinamico}' GROUP BY SKUUH, POÑUH"
    # Construir filas de la tabla para el correo
    filas_html = ""
    for r in analisis:
        po_val = str(r.get("PO", "N/A"))
        sku_val = str(r["SKU"])
        estado_val = r["ESTADO"]
        diferencia = r["DIFERENCIA"]
        
        if estado_val == "OK":
            estado_html = '<b style="color: #2e7d32;">OK</b>'
            obs_html = "Cuadre perfecto entre KTRHP y KPUHP"
        else:
            estado_html = f'<b style="color: #d32f2f;">ERROR</b>'
            obs_html = f"Descuadre de {diferencia} unidades"
            
        filas_html += f"""
        <tr>
            <td style="padding: 10px; border-bottom: 1px solid #ddd; border-right: 1px solid #eee;">{po_val}</td>
            <td style="padding: 10px; border-bottom: 1px solid #ddd; border-right: 1px solid #eee;">{sku_val}</td>
            <td style="padding: 10px; border-bottom: 1px solid #ddd; border-right: 1px solid #eee;">{estado_html}</td>
            <td style="padding: 10px; border-bottom: 1px solid #ddd;">{obs_html}</td>
        </tr>
        """
        
    if not filas_html.strip():
        filas_html = '<tr><td colspan="4" style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd; color: #666;">No se encontraron registros procesados para esta fecha.</td></tr>'

    cuerpo = f"""
    <html>
    <body style="font-family: Arial, sans-serif; background-color: #f9f9f9; padding: 20px;">
    <div style="max-width: 800px; margin: auto; background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.05); border-top: 5px solid #1976d2;">
        <h2 style="color: #1976d2; margin-top: 0; padding-bottom: 10px; border-bottom: 1px solid #eee;">
            Conciliación PO Automática - Vendor Storage
        </h2>
        
        <p style="font-size: 15px; color: #444; line-height: 1.5;">
            Estimado equipo,<br><br>
            Se ha completado la conciliación automática correspondiente a la fecha <b>{fecha}</b>.
        </p>

        <table style="width: 100%; border-collapse: collapse; font-size: 14px; margin-top: 20px; border: 1px solid #ddd; border-radius: 5px; overflow: hidden;">
            <thead>
                <tr style="background-color: #f5f5f5; color: #333;">
                    <th style="padding: 12px 10px; text-align: left; border-bottom: 2px solid #ddd; border-right: 1px solid #eee;">PO</th>
                    <th style="padding: 12px 10px; text-align: left; border-bottom: 2px solid #ddd; border-right: 1px solid #eee;">SKU</th>
                    <th style="padding: 12px 10px; text-align: left; border-bottom: 2px solid #ddd; border-right: 1px solid #eee;">Estado</th>
                    <th style="padding: 12px 10px; text-align: left; border-bottom: 2px solid #ddd;">Observaciones</th>
                </tr>
            </thead>
            <tbody>
                {filas_html}
            </tbody>
        </table>
        <p style="font-size: 13px; color: #777; margin-top: 25px; padding: 10px; background-color: #f0f7ff; border-left: 4px solid #1976d2;">
            <b>Nota:</b> Se adjunta el reporte detallado en formato Word con el análisis técnico por SKU y desglose de las unidades.
        </p>

        <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #eee;">
            <h4 style="color:#555; margin-bottom: 10px;">Explicación del Cálculo Técnico:</h4>
            <ul style="font-size:13px; color:#666; line-height: 1.6; padding-left: 20px;">
                <li>De <b>KTRHP</b> se agrupan los SKU con Owner igual a {valid_owner} o Vacío (Excluyendo Owner {owner_dinamico}).</li>
                <li>De <b>KPUHP</b> se agrupan los SKU cuyo JBDUH es '{transpo_dinamico}'.</li>
                <li><b>Fórmula:</b> Diferencia = [Suma KTRHP a generar] - [Suma KPUHP generado].</li>
            </ul>
            <br>
            <div style="background-color:#f4f4f4; padding:10px; border-radius:5px; font-family:Consolas, monospace; font-size:12px; color:#000;">
                <b>Query Extracción KTRHP:</b><br>
                {query_ktrhp}<br><br>
                <b>Query Extracción KPUHP:</b><br>
                {query_kpuhp}
            </div>
        </div>
    </div>
    </body>
    </html>
    """
    
    NotificadorCorreo(flujo='wms_infor').enviar(asunto, cuerpo, adjuntos=[os.path.abspath(ruta_reporte)])
    
    # --- ENVIAR REGISTROS 1 DÍA ATRÁS ---
    actualizar_registros_dia_anterior(BASE, fecha, transpo_dinamico)
        
    logger.info("=== FIN CONCILIACIÓN AS400 ===")

# ============================================================
# EJECUCIÓN
# ============================================================

if __name__ == "__main__":
    import json
    
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_path = os.path.join(base_dir, 'archivos_config', 'destinatarios_flujos.json')
    bases = []
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config_flujos = json.load(f)
            bases = config_flujos.get('bases_trabajo', [])
    except Exception as e:
        logger.error(f"Error leyendo el archivo JSON: {e}")
        
    if not bases:
        logger.error("No se encontró la configuración de 'bases_trabajo' en el JSON. Por favor, asegúrese de colocar las bases como parametrización.")
    else:
        for b in bases:
            base_name = f"RI{b}DB"
            logger.info(f"Ejecutando PO Automática para la base: {base_name}")
            main_vendor(BASE=base_name)
