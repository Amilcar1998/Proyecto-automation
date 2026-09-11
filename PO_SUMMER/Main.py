import pathlib
import sys
import traceback
import time
import json
import logging
import pyodbc
import ctypes

# Reconfigurar stdout a UTF-8 para prevenir errores de charmap en la consola de Windows
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from Page.Page_Login import LoginPage
from Page.page_menu import MenuPage
from OCEANO.OCEANOMAIN import main_descripcion, llenar_po

# Directorio base del proyecto
BASE_DIR = pathlib.Path(__file__).resolve().parent

# Constantes de Execution State para Windows
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

# Configuración del navegador
EJECUTAR_VISIBLE = True  # Cambiar a True para ver el navegador, False para ejecutar en segundo plano (Headless)

def evitar_suspension_pantalla():
    """Le dice al sistema operativo Windows que mantenga la pantalla encendida."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_DISPLAY_REQUIRED | ES_SYSTEM_REQUIRED
        )
        logging.info("Bloqueo de suspensión de pantalla ACTIVADO exitosamente.")
    except Exception as e:
        logging.warning(f"No se pudo activar el bloqueo de pantalla: {e}")

def permitir_suspension_pantalla():
    """Restaura el estado de energía normal de Windows."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        logging.info("Bloqueo de suspensión de pantalla DESACTIVADO.")
    except Exception:
        pass


# Carpetas donde están los archivos .xls a subir
CARPETA_ARCHIVOS = pathlib.Path(r"c:\Users\eliseo_lopezp\automatizacion\ARCHIVOS")
CARPETA_IMPORTADAS = CARPETA_ARCHIVOS / "IMPORTADA"
CARPETA_LOCALES = CARPETA_ARCHIVOS / "LOCAL"


def obtener_archivos_con_tipo(po_report: list = None) -> list[tuple[pathlib.Path, str]]:
    """Retorna la lista de tuplas (archivo_path, tipo_po) buscando en el po_report
    solo los archivos que tengan Estado='Y'.
    """
    CARPETA_IMPORTADAS.mkdir(parents=True, exist_ok=True)
    CARPETA_LOCALES.mkdir(parents=True, exist_ok=True)

    archivos_con_tipo = []

    if not po_report:
        # 1. Archivos en carpetas específicas
        for f in sorted(CARPETA_IMPORTADAS.glob("*.xls")):
            archivos_con_tipo.append((f, "Importada"))
        for f in sorted(CARPETA_LOCALES.glob("*.xls")):
            archivos_con_tipo.append((f, "Local"))
        # 2. Archivos sueltos en la raíz de 'archivos'
        for f in sorted(CARPETA_ARCHIVOS.glob("*.xls")):
            archivos_con_tipo.append((f, "Importada"))
        return archivos_con_tipo

    # Procesar usando el JSON
    for item in po_report:
        if item.get("Estado") == "Y" and item.get("Archivo_Generado"):
            path_str = item["Archivo_Generado"]
            f = pathlib.Path(path_str)
            if not f.is_absolute():
                f = CARPETA_ARCHIVOS.parent / f
            
            if f.exists():
                tipo_po = "Local" if "LOCAL" in str(item.get("PAIS", "")).upper() else "Importada"
                archivos_con_tipo.append((f, tipo_po))

    # Ordenar para que las POs Importadas se procesen primero
    archivos_con_tipo.sort(key=lambda x: 0 if x[1] == "Importada" else 1)
    
    return archivos_con_tipo


def importar_sku_desde_archivo(menu: MenuPage, full_path: str):
    """
    Lógica robusta para subir un archivo .xls, manejando iframes,
    elementos ocultos y la espera de la respuesta del servidor.
    """
    driver = menu.driver
    logging.info("Iniciando proceso de importación de archivo SKU...")

    # 1. Cambiar al iframe de importación
    try:
        logging.info("Cambiando al iframe 'frameImporSkuXlsF1'...")
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("frameImporSkuXlsF1"))
        WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    except TimeoutException:
        logging.error("No se pudo encontrar o cambiar al iframe 'frameImporSkuXlsF1'.")
        raise

    # 2. Asignar la ruta del archivo al input
    try:
        logging.info("Buscando el campo de entrada de archivo (input[type='file'])...")
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'], input[name='file']"))
        )
        driver.execute_script(
            "arguments[0].style.display = 'block'; arguments[0].style.visibility = 'visible';",
            file_input
        )
        time.sleep(0.5)
        logging.info(f"Asignando ruta: '{full_path}'")
        file_input.send_keys(full_path)
        time.sleep(1)
        val_asignado = file_input.get_attribute("value")
        if val_asignado:
            logging.info(f"Archivo asignado en el input. Valor: '{val_asignado}'")
        else:
            logging.warning("El navegador no muestra el 'value' en el input, pero se continuará con el envío.")
    except Exception as e:
        logging.exception("Error al asignar el archivo.")
        raise

    # 3. Hacer clic en el botón de 'Submit' para cargar el archivo
    try:
        logging.info("Buscando el botón para enviar el archivo (Submit)...")
        btn_submit = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Submit']"))
        )
        btn_submit.click()
        logging.info("Clic en el botón de envío 'Submit' realizado.")
    except Exception as e:
        logging.exception("Error al hacer clic en el botón 'Submit'.")
        raise

    # 4. Esperar la confirmación del servidor y cerrar el diálogo
    try:
        logging.info("Esperando respuesta del servidor...")
        success_xpath = "//*[contains(text(), 'Se termino el proceso exitosamente')]"
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, success_xpath))
        )
        logging.info("Confirmación recibida: El archivo se procesó en el backend.")

        # Salir del iframe de importación para poder interactuar con el diálogo
        driver.switch_to.parent_frame()

        logging.info("Buscando y haciendo clic en el botón 'Cerrar' del diálogo...")
        close_button_xpath = "//div[contains(@class, 'ui-dialog-buttonpane')]//button[normalize-space()='Cerrar']"

        try:
            # Intento 1: Clic normal en el botón "Cerrar"
            close_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, close_button_xpath))
            )
            close_button.click()
            logging.info("Diálogo de importación cerrado con clic normal.")
        except TimeoutException:
            logging.warning("No se pudo hacer clic en el botón 'Cerrar' de forma normal. Intentando forzar el clic con JavaScript.")
            # Intento 2: Forzar el clic con JavaScript si el clic normal falla
            driver.execute_script(f"document.evaluate(\"{close_button_xpath}\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();")
            logging.info("Diálogo de importación cerrado forzando el clic con JavaScript.")

        # Esperar explícitamente a que el cuadro de diálogo desaparezca
        try:
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, ".ui-dialog"))
            )
            logging.info("El cuadro de diálogo de importación se ha cerrado por completo.")
        except Exception:
            logging.warning("El cuadro de diálogo de importación tardó demasiado en desaparecer.")
        
        time.sleep(1) # Pequeña pausa adicional de seguridad

    except TimeoutException:
        logging.error("No se recibió confirmación de éxito del servidor después de 60 segundos.")
        raise
    except Exception as e:
        logging.exception("Error durante la espera de confirmación o al cerrar el diálogo.")
        raise

    # Regresar al contexto principal para los siguientes pasos
    driver.switch_to.default_content()


def guardar_reporte(po_report):
    """Guarda el diccionario de POs en el archivo JSON físico inmediatamente."""
    report_path = BASE_DIR / "po_report.json"
    try:
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(po_report, f, indent=4, ensure_ascii=False)
    except Exception as e:
        logging.error(f"Error al escribir en po_report.json: {e}")


def ejecutar_procedimiento_retaceo(po_numero=None):
    """
    Se conecta a DB2 (AS/400) vía pyodbc y ejecuta el procedimiento ELOPEZ.RETACEO().
    Si el procedimiento necesita el número de PO, se puede pasar como parámetro.
    """
    logging.info("Iniciando conexión a DB2 para ejecutar ELOPEZ.RETACEO()...")
    try:
        # Utilizando el origen de datos DSN 'RI_TEST' configurado en tu Windows para la BD
        conn_str = "DSN=RI_TEST;UID=ELOPEZ;PWD=MAY2024;"
        with pyodbc.connect(conn_str) as conn:
            with conn.cursor() as cursor:
                # Si el procedimiento recibe la PO, sería: cursor.execute("CALL ELOPEZ.RETACEO(?)", (po_numero,))
                cursor.execute("CALL ELOPEZ.RETACEO()")
                conn.commit()
                logging.info(f"[EXITO DB] Procedimiento ELOPEZ.RETACEO() ejecutado correctamente.")
                return True
    except Exception as e:
        logging.error(f"[ERROR DB] Ocurrió un error al ejecutar RETACEO en DB2: {e}")
        return False

def matar_sesion_activa_db2():
    """
    Se conecta a DB2 (AS/400) vía pyodbc y ejecuta el procedimiento ELOPEZ.dsession('ELOPEZ')
    para botar cualquier sesión activa en el servidor antes de iniciar.
    """
    logging.info("Iniciando conexión a DB2 para matar sesión activa (dsession)...")
    try:
        conn_str = "DSN=RI_TEST;UID=ELOPEZ;PWD=MAY2024;"
        with pyodbc.connect(conn_str) as conn:
            with conn.cursor() as cursor:
                cursor.execute("CALL ELOPEZ.dsession('ELOPEZ')")
                conn.commit()
                logging.info("[EXITO DB] Sesión activa en el servidor cerrada exitosamente (dsession).")
                return True
    except Exception as e:
        logging.warning(f"[ADVERTENCIA DB] No se pudo ejecutar dsession para botar la sesión (o no había sesión): {e}")
        return False


def crear_po_para_archivo(driver, menu: MenuPage, archivo: pathlib.Path, tipo_po: str, indice: int, total: int, po_report: list):
    """Flujo completo de creación de PO para un único archivo .xls."""
    logging.info(f"============================================================")
    logging.info(f"[PROCESANDO] [{indice}/{total}] ({tipo_po}): {archivo.name}")
    logging.info(f"============================================================")

    try:
        menu.driver.switch_to.default_content()
        menu.esperar_pagina_lista(timeout=25)
        time.sleep(2)

        menu.acceder_frame_menu()
        menu.esperar_pagina_lista()
        menu.crear_po_menu()
        time.sleep(1)
        menu.crear_po()
        time.sleep(3)

        menu.acceder_frame_trabajo()
        menu.esperar_pagina_lista()
        main_descripcion(driver)
        llenar_po(driver, tipo_po=tipo_po)

        try:
            logging.info("Buscando y haciendo clic en el botón 'Continuar'...")
            menu.acceder_frame_trabajo()
            try:
                continuar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "next")))
                continuar_button.click()
            except Exception:
                driver.execute_script("""
                    var btn = document.getElementById('next');
                    if (btn) btn.click();
                    else throw new Error("Botón 'next' no encontrado");
                """)
            logging.info("Clic en 'Continuar' (ID='next') realizado.")
            time.sleep(3)
        except Exception as e:
            logging.error(f"Fallaron todos los intentos de hacer clic en 'Continuar': {e}")
            raise

        logging.info("Buscando y haciendo clic en el botón 'Importar SKU'...")
        importar_exitoso = False
        for intento in range(10):  # Intentar por hasta ~30 segundos (10 intentos de 3 seg)
            try:
                menu.acceder_frame_trabajo()
                # Verificar primero que exista en el DOM
                importar_button = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "importarSkusXlsF1"))
                )
                # Scroll para que se vea y verificar que es clickeable
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", importar_button)
                importar_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "importarSkusXlsF1"))
                )
                importar_button.click()
                logging.info("Clic en 'Importar SKU' realizado con éxito.")
                importar_exitoso = True
                time.sleep(2)
                break
            except Exception as e:
                logging.warning(f"Intento {intento+1} de clic en 'Importar SKU' fallido. Reintentando en 3s...")
                time.sleep(3)

        if not importar_exitoso:
            logging.warning("No se pudo hacer clic en 'Importar SKU' normalmente. Intentando con JavaScript...")
            try:
                menu.acceder_frame_trabajo()
                driver.execute_script("""
                    var btn = document.getElementById('importarSkusXlsF1');
                    if (btn) btn.click();
                    else throw new Error("Botón 'importarSkusXlsF1' no encontrado en el DOM");
                """)
                logging.info("Se forzó el clic en 'Importar SKU' mediante JavaScript.")
                time.sleep(2)
            except Exception as e:
                logging.exception("No se pudo encontrar o hacer clic en el botón 'Importar SKU' de ninguna manera.")
                raise

        importar_sku_desde_archivo(menu, str(archivo.resolve()))

        po_numero = menu.obtener_numero_po()
        
        # Guardar en memoria inmediatamente por si falla lo siguiente
        if po_numero:
            logging.info(f"PO {po_numero} extraído. Guardando inmediatamente en JSON por si ocurre un fallo...")
            encontrado = False
            for item in po_report:
                if item.get("Archivo_Generado") and archivo.name in item["Archivo_Generado"]:
                    item["po_numero"] = po_numero
                    encontrado = True
                    break
            if not encontrado:
                po_report.append({"Archivo_Generado": str(archivo), "origen": tipo_po, "po_numero": po_numero, "Estado": "Pendiente"})
            guardar_reporte(po_report)

        menu.ir_a_resumen_orden()
        menu.enviar_a_aprobacion()
        
        if po_numero:
            logging.info(f"Procediendo a autorizar la PO {po_numero} desde el menú...")
            menu.autorizar_po_menu(po_numero)
            # Marcar como Procesado luego de autorizar
            for item in po_report:
                if item.get("Archivo_Generado") and archivo.name in item["Archivo_Generado"]:
                    item["Estado"] = "Procesado"
                    break
            guardar_reporte(po_report)

        logging.info(f"[EXITO] PO ({tipo_po}) N° Maestro [{po_numero}] procesada (Creada y Autorizada) para: {archivo.name}")
        time.sleep(3)
        return po_numero

    except Exception as e:
        logging.exception(f"Ocurrió un error fatal procesando '{archivo.name}'.")
        return None


def test_login_exitoso():
    log_file = BASE_DIR / "automation.log"
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, mode='w', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

    # Botar la sesión activa del servidor DB2 antes de hacer nada más
    matar_sesion_activa_db2()

    evitar_suspension_pantalla()

    options = webdriver.ChromeOptions()
    if not EJECUTAR_VISIBLE:
        # Ejecutar en segundo plano sin mostrar la ventana (Headless)
        options.add_argument('--headless=new')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-gpu') # Recomendado para evitar problemas en headless
    
    driver = webdriver.Chrome(options=options)
    # Aunque esté en background, simulamos maximizar la ventana
    driver.maximize_window()
    driver.get("http://was7tr1.siman.com/AccesoSUMMER/")

    login_page = LoginPage(driver)
    login_page.login("ELOPEZ", "MAY2024")
    time.sleep(10)

    menu = MenuPage(driver)
    
    # Cargar reporte existente como lista si existe
    report_path = BASE_DIR / "po_report.json"
    po_report = []
    if report_path.exists():
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                po_report = json.load(f)
                if not isinstance(po_report, list):
                    po_report = [] # Fallback por si no es lista
        except Exception:
            po_report = []

    menu.acceder_frame_menu()
    menu.hacer_click_en_summer()
    menu.click_oceano()
    time.sleep(2)

    archivos = obtener_archivos_con_tipo(po_report)
    if not archivos:
        logging.warning(f"No se encontraron archivos .xls en: {CARPETA_ARCHIVOS}")
        logging.info("Saltando Flujo 1 y pasando a revisar pendientes...")

    total = len(archivos)
    logging.info(f"Se encontraron {total} archivo(s) .xls para procesar.")

    # =========================================================================
    # FLUJO 1: CREAR Y AUTORIZAR POs
    # =========================================================================
    logging.info(f"--- INICIANDO FLUJO 1: CREACIÓN Y AUTORIZACIÓN DE POs ---")
    pos_creadas = False
    for indice, (archivo, tipo_po) in enumerate(archivos, start=1):
        po_numero = crear_po_para_archivo(driver, menu, archivo, tipo_po, indice, total, po_report)
        if po_numero:
            pos_creadas = True

    # Si no se creó ninguna PO, solo avisamos
    if not pos_creadas:
        logging.warning("No se pudo crear ni autorizar ninguna PO nueva en esta ejecución.")

    # =========================================================================
    # FLUJO 2: EJECUTAR PROCEDIMIENTO EN BASE DE DATOS
    # =========================================================================
    logging.info(f"--- INICIANDO FLUJO 2: EJECUCIÓN PROCEDIMIENTO DB2 ---")
    ejecutar_procedimiento_retaceo()

    # =========================================================================
    # FLUJO 3: CAMBIAR ESTADO DE POs Y EXTRAER RI
    # =========================================================================
    logging.info(f"--- INICIANDO FLUJO 3: CAMBIO DE ESTADO Y EXTRACCIÓN RI ---")
    
    # Procesamos los items que tienen PO pero aún no tienen RI
    items_a_procesar = [item for item in po_report if item.get("po_numero") and not item.get("RI")]
    total_creadas = len(items_a_procesar)
    
    # Primer intento
    for i, item in enumerate(items_a_procesar, start=1):
        po_numero = item["po_numero"]
        logging.info(f"[PROCESANDO ESTADO] [{i}/{total_creadas}] PO: {po_numero}")
        try:
            menu.cambiar_estado_po_menu()
            estado_ok, numero_ri = menu.buscar_y_cambiar_estado_po(po_numero)
            
            if estado_ok and numero_ri:
                logging.info(f"[EXITO] PO {po_numero} cambió de estado. RI extraído: {numero_ri}")
                item["RI"] = numero_ri
                item["Estado"] = "Completado"
            else:
                logging.warning(f"[ADVERTENCIA] Falló el cambio de estado de la PO {po_numero}. Marcando como Pendiente.")
                item["Estado"] = "Pendiente"
        except Exception as e:
            logging.error(f"[ERROR] Excepción cambiando estado de la PO {po_numero}: {e}. Marcando como Pendiente.")
            item["Estado"] = "Pendiente"
        
        guardar_reporte(po_report)

    # Segundo intento para los que quedaron Pendientes
    items_pendientes = [item for item in items_a_procesar if item.get("Estado") == "Pendiente"]
    if items_pendientes:
        logging.info(f"--- SE ENCONTRARON {len(items_pendientes)} PO(S) PENDIENTES. INICIANDO REINTENTO ---")
        logging.info("Ejecutando procedimiento de DB2 nuevamente antes del reintento...")
        ejecutar_procedimiento_retaceo()
        
        for i, item in enumerate(items_pendientes, start=1):
            po_numero = item["po_numero"]
            logging.info(f"[REINTENTO ESTADO] [{i}/{len(items_pendientes)}] PO: {po_numero}")
            try:
                menu.cambiar_estado_po_menu()
                estado_ok, numero_ri = menu.buscar_y_cambiar_estado_po(po_numero)
                
                if estado_ok and numero_ri:
                    logging.info(f"[REINTENTO EXITO] PO {po_numero} cambió de estado. RI extraído: {numero_ri}")
                    item["RI"] = numero_ri
                    item["Estado"] = "Completado"
                else:
                    logging.warning(f"[REINTENTO FALLIDO] Falló nuevamente el cambio de estado de la PO {po_numero}.")
            except Exception as e:
                logging.error(f"[REINTENTO ERROR] Excepción cambiando estado de la PO {po_numero}: {e}")
            
            guardar_reporte(po_report)

    logging.info(f"Proceso completado: se procesaron estados para {total_creadas} PO(s) en esta ejecución.")
    time.sleep(5)
    
    # Cerrar Sesión del sistema
    try:
        logging.info("Intentando cerrar sesión en el sistema...")
        # menu.cerrar_sesion() # Método no existe en MenuPage
    except Exception as e:
        logging.error(f"Error al cerrar sesión: {e}")
    
    permitir_suspension_pantalla()
    driver.quit()
    
    excel_path = generar_excel_final()
    
    # Validar si todas las POs están Completadas
    if po_report and all(item.get("Estado") == "Completado" for item in po_report):
        ya_registrado = all(item.get("Jira_Creado") == True for item in po_report)
        if not ya_registrado:
            logging.info("El PO_REPORT está al 100% completado. Procediendo a generar la tarea en Jira...")
        try:
            # Agregamos la ruta principal al sys.path si es necesario para importar jira
            import sys
            project_root = str(pathlib.Path(__file__).parent.parent)
            if project_root not in sys.path:
                sys.path.insert(0, project_root)
            from jira.jira_utilidades import JiraClient
            
            # Construir tabla para Jira con todo el detalle del JSON
            datos_tabla = [["ID", "Tipo", "País", "Predistribuido", "Archivo", "Origen", "PO Número", "Estado", "RI", "SKU", "Tienda", "Cantidades", "Items"]]
            for item in po_report:
                archivo_nombre = pathlib.Path(item.get("Archivo_Generado", "")).name if item.get("Archivo_Generado") else ""
                base_info = [
                    item.get("ID", ""),
                    item.get("TIPO", ""),
                    item.get("PAIS", ""),
                    item.get("PREDISTRIBUIDO", ""),
                    archivo_nombre,
                    item.get("origen", ""),
                    item.get("po_numero", ""),
                    item.get("Estado", ""),
                    item.get("RI", "")
                ]
                
                detalles = item.get("Detalles", [])
                if not detalles:
                    datos_tabla.append(base_info + ["", "", "", ""])
                else:
                    for det in detalles:
                        fila_completa = list(base_info)
                        fila_completa.extend([
                            det.get("SKU", ""),
                            det.get("TIENDA", ""),
                            det.get("CANTIDADES", ""),
                            det.get("ITEMS", "")
                        ])
                        datos_tabla.append(fila_completa)
            
            import datetime
            fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
            
            # Obtener base desde JSON de configuracion
            base_jira = ""
            config_path = BASE_DIR / "archivos_config" / "destinatarios_flujos.json"
            if config_path.exists():
                try:
                    with open(config_path, "r", encoding="utf-8") as f:
                        config = json.load(f)
                        bases = config.get("bases_trabajo", [])
                        if bases:
                            base_num = str(bases[0]).strip()
                            # Asumimos que si viene "12" se formatea a "RI12DB", si viene "RI12DB" se usa directo
                            if not base_num.startswith("RI"):
                                base_jira = f"RI{base_num}DB"
                            else:
                                base_jira = base_num
                except Exception as e:
                    logging.warning(f"No se pudo leer la base de {config_path}: {e}")

            jira = JiraClient()
            jira_key = jira.main_jira(
                load=fecha_actual,
                ruta_archivo=str(excel_path) if excel_path else "",
                tipo="po_summer_report",
                base=base_jira,
                datos_tabla=datos_tabla
            )
            
            if jira_key:
                # Subir el log de ejecución también como evidencia adicional
                log_file = BASE_DIR / "automation.log"
                if log_file.exists():
                    jira.subir_evidencia(jira_key, str(log_file))
                    logging.info("Log de ejecución (automation.log) subido a Jira como evidencia adicional.")
                    
                for item in po_report:
                    item["Jira_Creado"] = True
                guardar_reporte(po_report)
                logging.info("Reporte actualizado con la bandera Jira_Creado=True.")
                
        except Exception as e:
            logging.error(f"Error al generar la tarea en Jira: {e}", exc_info=True)
    else:
        logging.info("La tarea en Jira ya fue generada previamente para este reporte (bandera Jira_Creado activa).")

def generar_excel_final():
    """Genera un archivo de Excel con toda la información del po_report.json desglosada."""
    logging.info("Generando archivo Excel final con los resultados...")
    report_path = BASE_DIR / "po_report.json"
    if not report_path.exists():
        logging.warning("No se encontró po_report.json para generar el Excel.")
        return
        
    try:
        import pandas as pd
        with open(report_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            
        filas = []
        for item in data:
            base_info = {
                "ID": item.get("ID", ""),
                "TIPO": item.get("TIPO", ""),
                "PAIS": item.get("PAIS", ""),
                "PREDISTRIBUIDO": item.get("PREDISTRIBUIDO", ""),
                "Archivo_Generado": item.get("Archivo_Generado", ""),
                "Estado": item.get("Estado", ""),
                "PO_NUMERO": item.get("po_numero", ""),
                "RI": item.get("RI", "")
            }
            
            detalles = item.get("Detalles", [])
            if not detalles:
                filas.append(base_info)
            else:
                for det in detalles:
                    fila_completa = base_info.copy()
                    fila_completa["SKU"] = det.get("SKU", "")
                    fila_completa["TIENDA"] = det.get("TIENDA", "")
                    fila_completa["CANTIDADES"] = det.get("CANTIDADES", "")
                    fila_completa["ITEMS"] = det.get("ITEMS", "")
                    filas.append(fila_completa)
                    
        df_final = pd.DataFrame(filas)
        excel_path = BASE_DIR / "Reporte_Final_POs.xlsx"
        
        # Generar un nombre de pestaña basado en la fecha/hora
        import datetime
        sheet_name = datetime.datetime.now().strftime("Ejecucion_%Y%m%d_%H%M")
        
        if excel_path.exists():
            # Si el archivo existe, agregamos una nueva hoja
            with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a") as writer:
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)
            logging.info(f"Reporte añadido como nueva pestaña '{sheet_name}' en: {excel_path}")
        else:
            # Si no existe, creamos el archivo
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)
            logging.info(f"Reporte Excel generado por primera vez en: {excel_path}")
            
        return excel_path
    except Exception as e:
        logging.error(f"Error al generar el Excel final: {e}")
        return None

if __name__ == "__main__":
    test_login_exitoso()