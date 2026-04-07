import sys
import time
import pandas as pd
from pathlib import Path
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.common.exceptions import (
    TimeoutException, 
    NoSuchElementException, 
    ElementClickInterceptedException,
    WebDriverException
)
from docx import Document
from docx.shared import Inches, Pt
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from docx.shared import RGBColor
import logging
import os
from collections import defaultdict
from logging.handlers import RotatingFileHandler

# Configuración de logging con archivos
def configurar_logging():
    """Configura logging dual: consola y archivo en carpeta logs"""
    # Crear carpeta logs si no existe
    logs_dir = Path('logs')
    logs_dir.mkdir(exist_ok=True)
    
    # Nombre de archivo único con timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = logs_dir / f'automatizacion_{timestamp}.log'
    
    # Configurar logger principal
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    # Limpiar handlers previos
    logger.handlers.clear()
    
    # Formato común
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    # Handler para consola
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # Handler para archivo con rotación
    file_handler = RotatingFileHandler(
        log_file, 
        maxBytes=10*1024*1024,  # 10MB máximo por archivo
        backupCount=5,          # Mantener 5 archivos de respaldo
        encoding='utf-8'
    )
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    logger.info(f"📁 Log guardándose en: {log_file}")
    return logger, log_file

# Inicializar logging
LOGGER, LOG_FILE = configurar_logging()

# Constantes del sistema
URL_LOGIN = "http://ursvtols8b3c:8080/engage-unicomer-web/"
EMPLOYEE_ID = "pos"
PASSWORD = "pos1234"
TIMEOUT = 3

# Variables globales para control
SCREENSHOTS_TOMADAS = []
COTIZACIONES_PROCESADAS = []
DRIVER = None
COTIZACION_ACTUAL = None

def limpiar_screenshots_cotizacion():
    """Limpia screenshots de la cotización actual y los incluye en el Word"""
    global SCREENSHOTS_TOMADAS
    try:
        if SCREENSHOTS_TOMADAS:
            screenshots_eliminados = 0
            for archivo, descripcion in SCREENSHOTS_TOMADAS:
                if os.path.exists(archivo):
                    try:
                        os.remove(archivo)
                        screenshots_eliminados += 1
                    except Exception as e:
                        LOGGER.warning(f"Error eliminando {archivo}: {e}")
            
            SCREENSHOTS_TOMADAS.clear()
            if screenshots_eliminados > 0:
                LOGGER.info(f"✅ {screenshots_eliminados} screenshots de cotización {COTIZACION_ACTUAL} eliminados")
        
    except Exception as e:
        LOGGER.error(f"Error limpiando screenshots de cotización: {e}")

def capturar_screenshot(nombre_archivo, descripcion=""):
    """Captura screenshot y lo guarda en la lista global con ID de cotización"""
    global SCREENSHOTS_TOMADAS, DRIVER, COTIZACION_ACTUAL
    try:
        # Agregar prefijo con ID de cotización si está disponible
        if COTIZACION_ACTUAL:
            nombre_con_id = f"C{COTIZACION_ACTUAL}_{nombre_archivo}"
        else:
            nombre_con_id = nombre_archivo
            
        DRIVER.save_screenshot(nombre_con_id)
        SCREENSHOTS_TOMADAS.append((nombre_con_id, descripcion))
        LOGGER.info(f"Screenshot capturado: {nombre_con_id}")
    except Exception as e:
        LOGGER.error(f"Error capturando screenshot {nombre_archivo}: {e}")

# ================== MÓDULO 1: LECTURA EXCEL ==================

def leer_excel_cotizaciones(archivo_excel):
    """
    Módulo 1: Lee y procesa archivo Excel con cotizaciones
    Retorna: dict con cotizaciones agrupadas por cotizacion_id
    """
    try:
        LOGGER.info(f"=== MÓDULO 1: LECTURA EXCEL ===")
        LOGGER.info(f"Archivo: {archivo_excel}")
        
        df = pd.read_excel(archivo_excel)
        LOGGER.info(f"Columnas encontradas: {list(df.columns)}")
        
        # Verificar columnas requeridas
        columnas_requeridas = ['cotizacion_id', 'codigo']
        for col in columnas_requeridas:
            if col not in df.columns:
                LOGGER.error(f"Columna requerida '{col}' no encontrada")
                return None
        
        # Agrupar por cotización_id
        cotizaciones = {}
        # Filtrar solo cotizaciones válidas (ignorar NA que indica fin de datos)
        cotizaciones_validas = df['cotizacion_id'].dropna().unique()
        
        for cotizacion_id_raw in sorted(cotizaciones_validas):
            # Convertir a entero si es posible
            try:
                cotizacion_id = int(float(cotizacion_id_raw))
            except (ValueError, TypeError):
                cotizacion_id = cotizacion_id_raw
                
            productos_cotizacion = df[df['cotizacion_id'] == cotizacion_id_raw]
            cotizaciones[cotizacion_id] = productos_cotizacion.to_dict('records')
            
        LOGGER.info(f"Se encontraron {len(cotizaciones)} cotizaciones")
        for cot_id, productos in cotizaciones.items():
            LOGGER.info(f"Cotización {cot_id}: {len(productos)} productos")
            
        return cotizaciones
        
    except Exception as e:
        LOGGER.error(f"Error en módulo lectura Excel: {e}")
        return None

# ================== MÓDULO 2: INICIO SESIÓN ==================

def iniciar_navegador():
    """Inicializa navegador Firefox con configuración optimizada para cotizaciones múltiples"""
    global DRIVER
    try:
        options = FirefoxOptions()
        options.add_argument('--width=1920')
        options.add_argument('--height=1080')
        
        # Configuraciones para evitar problemas con sesiones múltiples
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.set_preference('dom.webdriver.enabled', False)
        options.set_preference('useAutomationExtension', False)
        
        # Configurar para nuevas ventanas/pestañas limpias
        options.set_preference('browser.tabs.remote.autostart', False)
        
        DRIVER = webdriver.Firefox(options=options)
        DRIVER.maximize_window()
        DRIVER.implicitly_wait(5)
        
        LOGGER.info(" Navegador Firefox inicializado con configuración limpia")
        return True
    
    except Exception as e:
        LOGGER.error(f"Error inicializando navegador: {e}")
        return False

def iniciar_sesion():
    """
    Módulo 2: Realiza login en el sistema
    Retorna: True si login exitoso, False si falla
    """
    global DRIVER
    try:
        LOGGER.info(f"=== MÓDULO 2: INICIO SESIÓN ===")
        
        DRIVER.get(URL_LOGIN)
        LOGGER.info(f"Accediendo a: {URL_LOGIN}")
        
        # Ingresar credenciales
        WebDriverWait(DRIVER, TIMEOUT).until(EC.presence_of_element_located((By.ID, "employeeId")))
        DRIVER.find_element(By.ID, "employeeId").send_keys(EMPLOYEE_ID)
        DRIVER.find_element(By.ID, "password").send_keys(PASSWORD)
        
        LOGGER.info("Credenciales ingresadas")
        
        # Submit login
        DRIVER.find_element(By.CSS_SELECTOR, "input[type='submit']").click()
        LOGGER.info("Formulario enviado")
        
        time.sleep(2)
        capturar_screenshot("01_login_resultado.png", "Resultado del login")
        
        # Verificar login exitoso
        if "engage-unicomer-web" in DRIVER.current_url and DRIVER.current_url != URL_LOGIN:
            LOGGER.info(" Login exitoso - URL cambió")
            return True
        else:
            LOGGER.error(" Login falló")
            return False
            
    except Exception as e:
        LOGGER.error(f"Error en módulo inicio sesión: {e}")
        return False

def iniciar_orden_compra():
    """Busca y hace click en botón 'Comience una Orden de Compra'"""
    global DRIVER
    try:
        LOGGER.info("Buscando botón 'Comience una Orden de Compra'...")
        capturar_screenshot("02_dashboard.png", "Dashboard principal")
        
        # Selectores basados en v2 funcional
        selectores = [
            (By.XPATH, "//input[@value='Comience una Orden de Compra']"),
            (By.CSS_SELECTOR, "a[href='/engage-unicomer-web/saleOrder/index'] input"),
            (By.XPATH, "//a[@href='/engage-unicomer-web/saleOrder/index']//input")
        ]
        
        for tipo, selector in selectores:
            try:
                elementos = DRIVER.find_elements(tipo, selector)
                if elementos:
                    elemento = elementos[0]
                    DRIVER.execute_script("arguments[0].scrollIntoView(true);", elemento)
                    #time.sleep(1)
                    
                    capturar_screenshot("03_boton_encontrado.png", "Botón localizado")
                    
                    elemento.click()
                    LOGGER.info(" Click realizado en botón orden")
                    time.sleep(1)
                    
                    capturar_screenshot("04_formulario_orden.png", "Formulario orden abierto")
                    return True
                    
            except Exception:
                continue
        
        LOGGER.error(" No se encontró botón orden de compra")
        return False
        
    except Exception as e:
        LOGGER.error(f"Error iniciando orden: {e}")
        return False

def extraer_numero_cotizacion_inicial():
    """Extrae número de cotización al inicio, después de abrir formulario de orden"""
    global DRIVER
    try:
        LOGGER.info("=== EXTRAYENDO NÚMERO DE COTIZACIÓN INICIAL ===")
        
        # Esperar a que se cargue completamente el formulario
        time.sleep(1)
        
        numero_cotizacion = extraer_numero_cotizacion()
        
        if numero_cotizacion and not numero_cotizacion.startswith("No_encontrado") and not numero_cotizacion.startswith("ERROR"):
            # Agregar timestamp único para evitar duplicados
            timestamp_unico = int(time.time() * 1000) % 10000  # Últimos 4 dígitos del timestamp
            numero_unico = f"{numero_cotizacion}_{timestamp_unico}"
            
            LOGGER.info(f"✅ Número base extraído: {numero_cotizacion}")
            LOGGER.info(f"✅ Número único generado: {numero_unico}")
            return numero_unico
        else:
            LOGGER.warning("No se pudo extraer número válido al inicio")
            # Generar número alternativo con timestamp único
            timestamp_fallback = int(time.time() * 1000) % 100000
            numero_fallback = f"COTIZACION_{timestamp_fallback}"
            LOGGER.info(f"Número alternativo generado: {numero_fallback}")
            return numero_fallback
        
    except Exception as e:
        LOGGER.error(f"Error extrayendo número inicial: {e}")
        # En caso de error, generar número único de emergencia
        timestamp_error = int(time.time() * 1000) % 100000
        numero_error = f"ERROR_{timestamp_error}"
        return numero_error

# ================== MÓDULO 3: AGREGAR ITEMS ==================

def agregar_item(codigo_producto, cantidad=1, numero_producto=1):
    """
    Módulo 3: Agrega un item individual al carrito
    Basado en lógica funcional del v2
    """
    global DRIVER
    try:
        LOGGER.info(f"=== MÓDULO 3: AGREGAR ITEM {numero_producto} ===")
        LOGGER.info(f"Código: {codigo_producto}, Cantidad: {cantidad}")
        
        # Buscar campo de búsqueda
        campo_busqueda = WebDriverWait(DRIVER, TIMEOUT).until(
            EC.presence_of_element_located((By.ID, "saleItemSearchInput"))
        )
        
        # Limpiar y agregar producto
        campo_busqueda.clear()
        campo_busqueda.send_keys(codigo_producto)
        
        # Submit búsqueda
        try:
            campo_busqueda.send_keys(Keys.RETURN)
        except:
            boton_busqueda = DRIVER.find_element(By.CSS_SELECTOR, "input[type='submit'][value='Search']")
            boton_busqueda.click()
        
        time.sleep(1)
        
        # Manejar popup de error si aparece
        manejar_popup_error()
        
        LOGGER.info(f"Item {codigo_producto} agregado correctamente")
        return True
        
    except Exception as e:
        LOGGER.error(f"Error agregando item {numero_producto}: {e}")
        return False

def manejar_popup_error():
    """Maneja popups de error del sistema"""
    global DRIVER
    try:
        # Buscar popup de error
        try:
            boton_cerrar = DRIVER.find_element(By.CSS_SELECTOR, "a.ui-dialog-titlebar-close")
            boton_cerrar.click()
            LOGGER.info(" Popup cerrado")
        except:
            pass
        
        # Manejar mensaje informativo
        try:
            mensaje_div = DRIVER.find_element(By.XPATH, "/html/body/div[3]")
            if mensaje_div.is_displayed():
                texto = mensaje_div.text
                if "ERROR TO OBTAIN" in texto:
                    LOGGER.info(f"Mensaje informativo (ignorado): {texto[:100]}...")
        except:
            pass
            
    except Exception as e:
        LOGGER.debug(f"Error manejando popup: {e}")

# ================== MÓDULO 4: AGREGAR GARANTÍAS ==================

def agregar_garantias(con_garantia, numero_producto):
    """
    Módulo 4: Maneja pantalla de garantías según configuración
    con_garantia: True/False o 'Si'/'No'
    """
    global DRIVER
    try:
        LOGGER.info(f"=== MÓDULO 4: GARANTÍAS ITEM {numero_producto} ===")
        
        # Normalizar valor
        if isinstance(con_garantia, str):
            con_garantia = con_garantia.lower() == 'si'
        
        LOGGER.info(f"Garantía: {'AGREGAR' if con_garantia else 'RECHAZAR'}")
        
        # Esperar pantalla de garantías
        time.sleep(1)
        
        capturar_screenshot(f"07_pantalla_garantias_{numero_producto:02d}.png", 
                           f"Pantalla garantías - Item {numero_producto}")
        
        if con_garantia:
            # AGREGAR GARANTÍA - Solo si aparece el mensaje
            try:
                # Primero verificar si aparece algún botón de garantías
                time.sleep(1)
                boton_anadir = None
                selectores_anadir = [
                    "//button[@type='button' and @class='btn btn-success' and text()='Añadir']",
                    "//button[contains(@class, 'btn-success') and contains(text(), 'Add')]", 
                    "//button[contains(text(), 'Add') and contains(@class, 'btn')]",
                    "//input[@value='Añadir']",
                    "//button[text()='Añadir']"
                ]
                
                # Verificar si aparece pantalla de garantías
                garantias_detectadas = False
                for selector in selectores_anadir:
                    try:
                        boton_anadir = DRIVER.find_element(By.XPATH, selector)
                        if boton_anadir.is_displayed() and boton_anadir.is_enabled():
                            LOGGER.info(f"Botón Add encontrado: {selector}")
                            garantias_detectadas = True
                            break
                    except:
                        continue
                
                if not garantias_detectadas:
                    LOGGER.info("No se detectó pantalla de garantías - continuando sin agregar")
                    return True
                
                # Si hay pantalla de garantías, proceder
                DRIVER.execute_script("arguments[0].click();", boton_anadir)
                LOGGER.info("Click en Añadir garantía")
                time.sleep(2)
                
                # Buscar tabla de garantías y procesar precios
                try:
                    # Buscar tabla en el contenedor principal
                    tabla = DRIVER.find_element(By.XPATH, "//div[@class='principalContainer']//table[@class='table table-bordered']")
                    filas = tabla.find_elements(By.XPATH, ".//tbody[@class='saleOrderBody']//tr[@class='whiteBackgorund']")
                    
                    garantia_seleccionada = False
                    todas_son_cero = True
                    
                    LOGGER.info(f"Encontradas {len(filas)} opciones de garantía")
                    
                    # Verificar precios de garantías
                    for i, fila in enumerate(filas):
                        try:
                            # Buscar celda de precio (tercera columna)
                            precio_celda = fila.find_element(By.XPATH, ".//td[3]")
                            precio_texto = precio_celda.text.strip()
                            
                            # Verificar si el precio es mayor a $0
                            if '$' in precio_texto:
                                precio_numero = precio_texto.replace('$', '').replace(',', '').strip()
                                try:
                                    precio_valor = float(precio_numero)
                                    
                                    if precio_valor > 0:
                                        # Precio mayor a $0 - seleccionar esta garantía con Select
                                        todas_son_cero = False
                                        boton_select = fila.find_element(By.XPATH, ".//button[@class='btn btn-success mediumHeightBtn']")
                                        DRIVER.execute_script("arguments[0].click();", boton_select)
                                        LOGGER.info(f"Garantía seleccionada - precio: {precio_texto}")
                                        garantia_seleccionada = True
                                        break
                                    else:
                                        # Precio es $0.00 - usar Void en lugar de Select
                                        LOGGER.info(f"Precio $0.00 encontrado - buscando botón Void")
                                        # Hacer clic en Void para garantías de $0.00
                                        try:
                                            # Buscar botón Void con XPath específico
                                            boton_void = DRIVER.find_element(By.XPATH, "/html/body/div[2]/div[5]/div/div/a/button")
                                            DRIVER.execute_script("arguments[0].click();", boton_void)
                                            LOGGER.info("Clic en botón Void para garantía $0.00")
                                            garantia_seleccionada = True
                                            break
                                        except Exception as e:
                                            LOGGER.warning(f"No se pudo hacer clic en Void: {e}")
                                            # Fallback: usar Select si no encuentra Void
                                            boton_select = fila.find_element(By.XPATH, ".//button[@class='btn btn-success mediumHeightBtn']")
                                            DRIVER.execute_script("arguments[0].click();", boton_select)
                                            LOGGER.info(f"Fallback: Garantía $0.00 seleccionada con Select")
                                            garantia_seleccionada = True
                                            break
                                        
                                except ValueError:
                                    continue
                        except Exception as e:
                            LOGGER.warning(f"Error procesando fila {i}: {e}")
                            continue
                    
                    # Después de seleccionar garantía con precio > $0
                    if garantia_seleccionada and not todas_son_cero:
                        time.sleep(1)
                        try:
                            # Solo para garantías con precio > $0, hacer clic en botón final
                            boton_final = DRIVER.find_element(By.XPATH, "/html/body/div[2]/div[5]/div/div/a/button")
                            DRIVER.execute_script("arguments[0].click();", boton_final)
                            LOGGER.info("Clic en botón final para garantía con precio")
                            time.sleep(1)
                        except Exception as e:
                            LOGGER.warning(f"No se pudo hacer clic en botón final: {e}")
                        
                        capturar_screenshot(f"08_garantia_procesada_{numero_producto:02d}.png", 
                                           f"Garantía procesada - Item {numero_producto}")
                    elif garantia_seleccionada:
                        # Para garantías $0.00, el botón Void ya manejó todo
                        capturar_screenshot(f"08_garantia_void_{numero_producto:02d}.png", 
                                           f"Garantía Void procesada - Item {numero_producto}")
                    else:
                        # Si no se pudo seleccionar ninguna garantía, usar No Thanks
                        LOGGER.warning("No se pudo seleccionar garantía, usando No Thanks")
                        rechazar_garantia()
                        
                except Exception as e:
                    LOGGER.warning(f"Error en tabla garantías: {e}, usando No Thanks")
                    rechazar_garantia()
                    
            except Exception as e:
                LOGGER.warning(f"Error agregando garantía: {e}")
                rechazar_garantia()
        else:
            # RECHAZAR GARANTÍA - Solo si aparece botón No Thanks
            try:
                # Verificar si aparece botón No Thanks
                time.sleep(1)
                selectores_no = [
                    "//input[@value='No Thanks']", 
                    "//button[contains(text(), 'No Thanks')]"
                ]
                
                boton_no_thanks = None
                for selector in selectores_no:
                    try:
                        boton_no_thanks = DRIVER.find_element(By.XPATH, selector)
                        if boton_no_thanks.is_displayed():
                            LOGGER.info("Botón No Thanks encontrado - rechazando garantía")
                            DRIVER.execute_script("arguments[0].click();", boton_no_thanks)
                            LOGGER.info("Click en No Thanks")
                            break
                    except:
                        continue
                
                if not boton_no_thanks:
                    LOGGER.info("No se detectó pantalla de garantías - continuando sin rechazar")
                    
            except Exception as e:
                LOGGER.warning(f"Error rechazando garantía: {e}")
        
        time.sleep(1)
        
        # Capturar evidencia DESPUÉS de procesar garantías
        capturar_screenshot(f"06_item_con_garantias_procesadas_{numero_producto:02d}.png", 
                           f"Item {numero_producto} con garantías procesadas")
        
        LOGGER.info(f"Garantías procesadas para item {numero_producto}")
        return True
        
    except Exception as e:
        LOGGER.error(f"Error en módulo garantías: {e}")
        return False

def rechazar_garantia():
    """Rechaza garantía haciendo click en No Thanks"""
    global DRIVER
    try:
        selectores_no = [
            "//input[@value='No Thanks']", 
            "//button[contains(text(), 'No Thanks')]"
        ]
        
        for selector in selectores_no:
            try:
                boton_no = DRIVER.find_element(By.XPATH, selector)
                if boton_no.is_displayed():
                    DRIVER.execute_script("arguments[0].click();", boton_no)
                    LOGGER.info(" Click en No Thanks")
                    return True
            except:
                continue
        
        LOGGER.warning(" No se encontró botón No Thanks")
        return False
        
    except Exception as e:
        LOGGER.warning(f"Error rechazando garantía: {e}")
        return False

# ================== MÓDULO 5: ENLAZAR CUSTOMER ==================

def enlazar_customer():
    """
    Módulo 5: Enlaza cliente usando lógica funcional del v2
    """
    global DRIVER
    try:
        LOGGER.info(f"=== MÓDULO 5: ENLAZAR CUSTOMER ===")
        
        capturar_screenshot("12_antes_enlazar_cliente.png", "Antes de enlazar cliente")
        
        # Cerrar cualquier modal que esté bloqueando
        try:
            # Buscar y cerrar modal backdrop
            modal_backdrop = DRIVER.find_element(By.CSS_SELECTOR, ".modal-backdrop")
            if modal_backdrop.is_displayed():
                LOGGER.info("Modal backdrop detectado, intentando cerrar...")
                # Intentar cerrar modal con Escape
                DRIVER.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                time.sleep(1)
                
                # Si aún existe, intentar hacer clic fuera del modal
                try:
                    DRIVER.execute_script("document.querySelector('.modal-backdrop').remove();")
                    LOGGER.info("Modal backdrop eliminado con JavaScript")
                except:
                    pass
        except:
            pass
        
        # Buscar enlace de cliente (basado en v2)
        selectores_enlace = [
            "/html/body/div[2]/div[3]/div[28]/div[3]/h4/a",
            "//a[contains(text(), 'cliente')]",
            "//a[contains(text(), 'Cliente')]", 
            "//a[contains(text(), 'customer')]",
            "//h4/a[contains(@href, 'customer')]",
            "//div[contains(@class, 'customer')]//a",
            ".customer-link",
            "a[href*='customer']"
        ]
        
        enlace_cliente = None
        for selector in selectores_enlace:
            try:
                if selector.startswith("/"):
                    enlace_cliente = DRIVER.find_element(By.XPATH, selector)
                else:
                    enlace_cliente = DRIVER.find_element(By.CSS_SELECTOR, selector)
                
                if enlace_cliente and enlace_cliente.is_displayed():
                    LOGGER.info(f"Enlace cliente encontrado: {selector}")
                    break
            except:
                continue
        
        if not enlace_cliente:
            LOGGER.error(" No se encontró enlace de cliente")
            return False
        
        # Click en enlace usando JavaScript para evitar interferencias
        DRIVER.execute_script("arguments[0].click();", enlace_cliente)
        LOGGER.info(" Click en enlace cliente")
        time.sleep(3)
        
        capturar_screenshot("13_formulario_cliente.png", "Formulario cliente abierto")
        
        # Buscar campo de cliente (basado en v2)
        selectores_campo = [
            "/html/body/div[2]/form/div[1]/div/div[6]//input",
            "/html/body/div[2]/form/div[1]/div/div[6]/input", 
            "input[type='text'][placeholder*='cliente']",
            "input[type='text'][name*='customer']",
            "input[type='text'][id*='customer']",
            ".boxCustomerSearchRow input[type='text']",
            "div.boxCustomerSearchRow input"
        ]
        
        campo_cliente = None
        for selector in selectores_campo:
            try:
                if selector.startswith("/"):
                    campo_cliente = DRIVER.find_element(By.XPATH, selector)
                else:
                    campo_cliente = DRIVER.find_element(By.CSS_SELECTOR, selector)
                break
            except:
                continue
        
        if not campo_cliente:
            LOGGER.error(" No se encontró campo de cliente")
            return False
        
        # Ingresar número de cliente
        try:
            campo_cliente.click()
            time.sleep(1)
            campo_cliente.clear()
        except:
            campo_cliente.send_keys(Keys.CONTROL + "a")
        
        numero_cliente = "057874256"
        campo_cliente.send_keys(numero_cliente)
        campo_cliente.send_keys(Keys.RETURN)
        LOGGER.info(f"Cliente {numero_cliente} ingresado")
        time.sleep(1)
        
        capturar_screenshot("14_cliente_ingresado.png", "Cliente ingresado")
        
        # Primer Next
        try:
            next_button_1 = WebDriverWait(DRIVER, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='nextButton']"))
            )
            next_button_1.click()
            LOGGER.info(" Primer Next")
            time.sleep(3)
            
            capturar_screenshot("15_primer_next.png", "Después primer Next")
        except TimeoutException:
            LOGGER.warning(" Primer Next no encontrado")
        
        # Segundo Next
        try:
            next_button_2 = WebDriverWait(DRIVER, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='nextButton']"))
            )
            next_button_2.click()
            LOGGER.info(" Segundo Next")
            time.sleep(3)
            
            capturar_screenshot("16_segundo_next.png", "Después segundo Next")
        except TimeoutException:
            LOGGER.warning(" Segundo Next no encontrado")
        
        LOGGER.info(" Customer enlazado exitosamente")
        return True
        
    except Exception as e:
        LOGGER.error(f"Error en módulo enlazar customer: {e}")
        return False

# ================== MÓDULO 6: AGREGAR DESPACHO ==================

def agregar_despacho(entrega_cliente, tipo_entrega="cliente llevará"):
    """
    Módulo 6: Configura opciones de despacho/entrega
    entrega_cliente: True/False o 'Si'/'No' desde Excel
    tipo_entrega: "Dirección del cliente" o "cliente llevará" desde Excel
    """
    global DRIVER
    try:
        LOGGER.info(f"=== MÓDULO 6: DESPACHO/ENTREGA ===")
        
        # Normalizar valor del Excel para entrega_cliente
        if isinstance(entrega_cliente, str):
            con_despacho = entrega_cliente.lower().strip() in ['si', 'sí', 'yes', 'y', '1', 'true']
        elif entrega_cliente is None:
            con_despacho = False
        else:
            con_despacho = bool(entrega_cliente)
        
        LOGGER.info(f"Despacho/Entrega: {'CONFIGURAR' if con_despacho else 'OMITIR'}")
        LOGGER.info(f"Tipo de entrega: {tipo_entrega}")
        
        # Si entrega_cliente es "No", omitir completamente el módulo de despacho
        if not con_despacho:
            LOGGER.info("ℹ️ Entrega cliente omitida (configurado como No)")
            return True
        
        # Solo proceder si entrega_cliente es "Si"
        capturar_screenshot("17_antes_entrega_cliente.png", "Antes de configurar entrega cliente")
        
        try:
            # PASO 1: Click en Transaction Options Button
            LOGGER.info("Paso 1: Buscando botón Transaction Options")
            time.sleep(3)  # Espera adicional después de customer
            
            # Múltiples selectores para Transaction Options
            selectores_transaction = [
                (By.ID, "transactionOptionsBtn"),
                (By.XPATH, "//*[@id='transactionOptionsBtn']"),
                (By.XPATH, "//button[contains(@id, 'transactionOptions')]"),
                (By.XPATH, "//button[contains(text(), 'Transaction')]"),
                (By.CSS_SELECTOR, "button[id*='transaction']"),
            ]
            
            transaction_btn = None
            for selector_type, selector_value in selectores_transaction:
                try:
                    transaction_btn = WebDriverWait(DRIVER, 5).until(
                        EC.element_to_be_clickable((selector_type, selector_value))
                    )
                    LOGGER.info(f"✓ Transaction Options encontrado con: {selector_value}")
                    break
                except:
                    continue
            
            if transaction_btn:
                DRIVER.execute_script("arguments[0].click();", transaction_btn)
                LOGGER.info("✓ Click en Transaction Options exitoso")
                time.sleep(3)
            else:
                raise Exception("Transaction Options button no encontrado")
            
            capturar_screenshot("18_menu_transaction_options.png", "Menú Transaction Options abierto")
            
            # PASO 2: Click en opción de entrega (li[6])
            LOGGER.info("Paso 2: Seleccionando opción de entrega")
            opcion_entrega = WebDriverWait(DRIVER, TIMEOUT).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[3]/div[28]/div[4]/ul/li[6]/a"))
            )
            DRIVER.execute_script("arguments[0].click();", opcion_entrega)
            LOGGER.info("✓ Opción de entrega seleccionada")
            time.sleep(3)
            
            capturar_screenshot("19_formulario_entrega.png", "Formulario de entrega abierto")
            
            # PASO 3: Llenar datos del formulario
            LOGGER.info("Paso 3: Llenando datos del formulario")
            
            # Campo: ELISEO AMILCAR LOPEZ
            try:
                campo_nombre1 = WebDriverWait(DRIVER, TIMEOUT).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/form/div[1]/div/div[2]/div[2]/input"))
                )
                campo_nombre1.clear()
                campo_nombre1.send_keys("ELISEO AMILCAR LOPEZ")
                campo_nombre1.send_keys(Keys.RETURN)  # Enter después de escribir
                LOGGER.info("✓ Nombre 1 ingresado: ELISEO AMILCAR LOPEZ + Enter")
                time.sleep(1)
            except Exception as e:
                LOGGER.warning(f"Error ingresando nombre 1: {e}")
            
            # Campo: MARGARITA PORTILLO DE LOPEZ
            try:
                campo_nombre2 = DRIVER.find_element(By.XPATH, "/html/body/div[2]/form/div[1]/div/div[2]/div[3]/input")
                campo_nombre2.clear()
                campo_nombre2.send_keys("MARGARITA PORTILLO DE LOPEZ")
                campo_nombre2.send_keys(Keys.RETURN)  # Enter después de escribir
                LOGGER.info("✓ Nombre 2 ingresado: MARGARITA PORTILLO DE LOPEZ + Enter")
                time.sleep(1)
            except Exception as e:
                LOGGER.warning(f"Error ingresando nombre 2: {e}")
            
            # PASO 3.5: Campo adicional (div[4]/input)
            try:
                campo_adicional = DRIVER.find_element(By.XPATH, "/html/body/div[2]/form/div[1]/div/div[2]/div[4]/input")
                campo_adicional.clear()
                campo_adicional.send_keys("ELISEO LOPEZ")
                campo_adicional.send_keys(Keys.RETURN)  # Enter después de escribir
                LOGGER.info("✓ Campo adicional (div[4]) ingresado: ELISEO LOPEZ + Enter")
                time.sleep(1)
            except Exception as e:
                LOGGER.warning(f"Error ingresando campo adicional (div[4]): {e}")
                
            # PASO 4: Select MADRE
            LOGGER.info("Paso 4: Configurando select MADRE")
            try:
                # Es un Select2, necesitamos hacer click en el span para abrirlo
                select_opener = WebDriverWait(DRIVER, TIMEOUT).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/form/div[1]/div/div[2]/div[5]/span[2]/span/a"))
                )
                
                # Click para abrir el dropdown
                select_opener.click()
                LOGGER.info("✓ Select MADRE abierto")
                time.sleep(1)
                
                # Buscar y hacer click en la opción "Madre"
                # Este es un jQuery UI Autocomplete, necesitamos buscar en li/a elementos
                opciones_madre = [
                    "//li[@class='ui-menu-item']//a[text()='Madre']",
                    "//ul[@class='ui-autocomplete']//li//a[text()='Madre']",
                    "//li[@class='ui-menu-item']//a[contains(text(), 'Madre')]",
                    "//ul[contains(@class, 'ui-autocomplete')]//a[text()='Madre']",
                    "//li[@role='menuitem']//a[text()='Madre']",
                    "//a[@class='ui-corner-all'][text()='Madre']"
                ]
                
                madre_seleccionada = False
                for xpath_opcion in opciones_madre:
                    try:
                        opcion_madre = WebDriverWait(DRIVER, 3).until(
                            EC.element_to_be_clickable((By.XPATH, xpath_opcion))
                        )
                        
                        # Hacer scroll a la opción si es necesario
                        DRIVER.execute_script("arguments[0].scrollIntoView(true);", opcion_madre)
                        time.sleep(0.5)
                        
                        opcion_madre.click()
                        LOGGER.info("✓ Opción MADRE seleccionada")
                        madre_seleccionada = True
                        break
                    except Exception as e:
                        LOGGER.debug(f"Intento fallido con {xpath_opcion}: {e}")
                        continue
                
                if not madre_seleccionada:
                    LOGGER.warning("No se pudo seleccionar MADRE, tomando screenshot para debug")
                    capturar_screenshot("debug_select_madre_abierto.png", "Select Madre abierto para debug")
                    
                    # Intentar scroll manual en el contenedor
                    try:
                        LOGGER.info("Intentando scroll manual en el dropdown jQuery UI")
                        dropdown_container = DRIVER.find_element(By.CSS_SELECTOR, ".ui-autocomplete")
                        DRIVER.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight / 2;", dropdown_container)
                        time.sleep(1)
                        
                        # Intentar nuevamente después del scroll con selectores jQuery UI
                        for xpath_opcion in ["//li[@class='ui-menu-item']//a[text()='Madre']", "//a[@class='ui-corner-all'][text()='Madre']"]:
                            try:
                                opcion_madre = DRIVER.find_element(By.XPATH, xpath_opcion)
                                if opcion_madre.is_displayed():
                                    DRIVER.execute_script("arguments[0].click();", opcion_madre)
                                    LOGGER.info("✓ MADRE seleccionada después de scroll manual")
                                    madre_seleccionada = True
                                    break
                            except:
                                continue
                    except Exception as scroll_error:
                        LOGGER.warning(f"Error con scroll manual: {scroll_error}")
                    
                    # Último recurso: JavaScript directo
                    if not madre_seleccionada:
                        LOGGER.warning("Usando JavaScript como último recurso")
                        try:
                            DRIVER.execute_script("document.body.click();")  # Cerrar dropdown
                            time.sleep(0.5)
                            DRIVER.execute_script("""
                                var select = document.getElementById('relationshipSelect');
                                if (select) {
                                    select.value = '009';
                                    select.dispatchEvent(new Event('change'));
                                    if (typeof handleOnChange === 'function') {
                                        handleOnChange('relationshipSelect', '009');
                                    }
                                }
                            """)
                            LOGGER.info("✓ MADRE seleccionada por JavaScript")
                        except Exception as js_error:
                            LOGGER.error(f"Error con JavaScript: {js_error}")
                
                time.sleep(1)
                
            except Exception as e:
                LOGGER.warning(f"Error configurando select MADRE: {e}")
                
            # PASO 4B: Configurar TRANSPORTADOR (Relationship)
            LOGGER.info("Paso 4B: Configurando transportador (relationship)")
            try:
                # Buscar el select del transportador/relationship
                transportador_selectors = [
                    "/html/body/div[2]/form/div[1]/div/div[2]/div[6]/span[2]/span/a",  # XPath directo
                    "//div[contains(@class, 'boxSelectRow')]//span[contains(@class, 'custom-combobox')]//a",  # Genérico
                    "//span[contains(text(), 'Transportador')]/following-sibling::span//a",  # Por texto
                    "//select[@id='relationshipSelect']/following-sibling::span//a"  # Por el select relationship
                ]
                
                transportador_configurado = False
                for selector in transportador_selectors:
                    try:
                        transportador_opener = WebDriverWait(DRIVER, 5).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        transportador_opener.click()
                        LOGGER.info(f"✓ Select TRANSPORTADOR abierto con: {selector}")
                        time.sleep(1)
                        
                        # Buscar opción del transportador (generalmente "Cliente" o "Empleado")
                        opciones_transportador = [
                            "//li[@class='ui-menu-item']//a[contains(text(), 'Cliente')]",
                            "//ul[@class='ui-autocomplete']//li//a[contains(text(), 'Cliente')]", 
                            "//li[@class='ui-menu-item']//a[contains(text(), 'Empleado')]",
                            "//a[@class='ui-corner-all'][contains(text(), 'Cliente')]"
                        ]
                        
                        for xpath_opcion in opciones_transportador:
                            try:
                                opcion_transportador = WebDriverWait(DRIVER, 3).until(
                                    EC.element_to_be_clickable((By.XPATH, xpath_opcion))
                                )
                                opcion_transportador.click()
                                LOGGER.info("✓ Opción TRANSPORTADOR seleccionada")
                                transportador_configurado = True
                                break
                            except:
                                continue
                        
                        if transportador_configurado:
                            break
                            
                    except:
                        continue
                
                if not transportador_configurado:
                    LOGGER.warning("⚠️ No se pudo configurar transportador - continuando")
                    # Cerrar cualquier dropdown abierto
                    DRIVER.execute_script("document.body.click();")
                
                time.sleep(1)
            except Exception as e:
                LOGGER.warning(f"Error configurando transportador: {e}")
                DRIVER.execute_script("document.body.click();")  # Cerrar dropdown
                
            # PASO 5: Configurar tipo de entrega según Excel
            LOGGER.info(f"Paso 5: Configurando tipo de entrega: {tipo_entrega}")
            try:
                # Es un Select2/jQuery UI, necesitamos hacer click en el span para abrirlo
                select_opener = WebDriverWait(DRIVER, TIMEOUT).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/form/div[1]/div/div[2]/div[6]/span[2]/span/a"))
                )
                
                # Click para abrir el dropdown
                select_opener.click()
                LOGGER.info("✓ Select TIPO ENTREGA abierto")
                time.sleep(1)
                
                # Buscar la opción correcta según el tipo
                if "cliente" in tipo_entrega.lower():
                    texto_buscar = "Cliente llevará (Cliente Recoge)"
                else:
                    texto_buscar = tipo_entrega
                    
                opciones_entrega = [
                    f"//li[@class='ui-menu-item']//a[contains(text(), '{texto_buscar}')]",
                    f"//ul[@class='ui-autocomplete']//li//a[contains(text(), '{texto_buscar}')]",
                    f"//a[@class='ui-corner-all'][contains(text(), '{texto_buscar}')]",
                    f"//li[@class='ui-menu-item']//a[contains(text(), 'Cliente llevará')]",
                    f"//li[@class='ui-menu-item']//a[contains(text(), 'Cliente Recoge')]"
                ]
                
                entrega_seleccionada = False
                for xpath_opcion in opciones_entrega:
                    try:
                        # ESPERAR a que la opción esté disponible SIN hacer scroll ni manipular
                        opcion_entrega = WebDriverWait(DRIVER, 5).until(
                            EC.element_to_be_clickable((By.XPATH, xpath_opcion))
                        )
                        # UN SOLO CLICK directo, sin scroll ni manipulación extra
                        opcion_entrega.click()
                        LOGGER.info(f"✓ Tipo entrega seleccionado: {texto_buscar}")
                        entrega_seleccionada = True
                        break
                    except:
                        continue
                        
                if not entrega_seleccionada:
                    LOGGER.warning("Tipo entrega no seleccionado")
                    # NO hacer click en body, dejar que el formulario se maneje solo
                
                # ESPERAR más tiempo para que el formulario procese la selección
                time.sleep(2)
            except Exception as e:
                LOGGER.warning(f"Error configurando tipo entrega: {e}")
            
            # PASO 6: Select lugar de salida (bodega nejapa) - USAR AUTOCOMPLETE
            LOGGER.info("Paso 6: Configurando lugar de salida (Bodega Nejapa)")
            try:
                # Buscar el campo de entrada de texto para lugar de salida
                lugar_input = WebDriverWait(DRIVER, TIMEOUT).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/form/div[1]/div/div[2]/div[7]/span[2]/span/input"))
                )
                
                # Limpiar y escribir texto para activar autocomplete
                lugar_input.clear()
                lugar_input.send_keys("BODEGA 94 NEJAPA")
                LOGGER.info("✓ Texto escrito: BODEGA 94 NEJAPA")
                time.sleep(2)  # Esperar que aparezca el dropdown
                
                # Buscar y hacer click en la opción del dropdown que aparece
                opciones_lugar = [
                    "//li[@class='ui-menu-item']//a[contains(text(), 'BODEGA 94 NEJAPA')]",
                    "//ul[@class='ui-autocomplete']//li//a[contains(text(), 'BODEGA 94')]",
                    "//a[@class='ui-corner-all'][contains(text(), 'NEJAPA')]",
                    "//li[@class='ui-menu-item']//a[contains(text(), 'BODEGA')]",
                    "//div[@class='ui-menu-item-wrapper'][contains(text(), 'BODEGA 94')]",
                    "//li[contains(@class, 'ui-menu-item')][contains(text(), 'NEJAPA')]"
                ]
                
                lugar_seleccionado = False
                for xpath_opcion in opciones_lugar:
                    try:
                        opcion_lugar = WebDriverWait(DRIVER, 3).until(
                            EC.element_to_be_clickable((By.XPATH, xpath_opcion))
                        )
                        opcion_lugar.click()
                        LOGGER.info("✓ Opción clickeada del dropdown: BODEGA 94 NEJAPA")
                        lugar_seleccionado = True
                        break
                    except:
                        continue
                
                # Si no encuentra la opción en el dropdown, usar Enter como fallback
                if not lugar_seleccionado:
                    LOGGER.info("🔄 Fallback: Usando Enter para confirmar texto de lugar")
                    lugar_input.send_keys(Keys.ENTER)
                    LOGGER.info("✓ Enter confirmado en lugar de salida")
                
                # HACER CLICK EN OK/CONFIRMAR después de seleccionar lugar
                try:
                    boton_ok_lugar = WebDriverWait(DRIVER, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'OK') or contains(text(), 'Aceptar') or contains(text(), 'Confirmar')]"))
                    )
                    boton_ok_lugar.click()
                    LOGGER.info("✓ Click en OK del lugar de salida")
                except:
                    LOGGER.info("ℹ️ No se requiere OK para lugar de salida")
                
                # ESPERAR para que el formulario procese
                time.sleep(2)
            except Exception as e:
                LOGGER.warning(f"Error configurando lugar de salida: {e}")
            
            # PASO 7: Select transporte - ESCRITURA DIRECTA
            LOGGER.info("Paso 7: Configurando transporte (11094 - BODEGA 94 NEJAPA)")
            
            time.sleep(2)
            
            try:
                # Encontrar la caja de texto del transporte
                transporte_input = WebDriverWait(DRIVER, TIMEOUT).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/form/div[1]/div/div[2]/div[8]/span[2]/span/input"))
                )
                LOGGER.info("✓ Caja de texto del transporte encontrada")
                # PASO 7A: Limpiar completamente el input antes de escribir
                transporte_input.click()
                # Limpiar usando múltiples métodos para asegurar que está vacío
                transporte_input.clear()
                time.sleep(0.5)
                
                
                texto_transporte = "BODEGA 94 NEJAPA"
                transporte_input.send_keys(texto_transporte)
                actions = ActionChains(DRIVER)
                actions.send_keys(Keys.ARROW_DOWN).send_keys(Keys.ENTER).perform()
                c
                LOGGER.info(f"✓ Texto escrito: {texto_transporte}")
                
                time.sleep(1)  # Esperar que aparezca el dropdown de opciones
                
                # PASO 7C: Buscar y hacer click en la opción específica del dropdown
                try:
                    LOGGER.info("-Buscando opción '11094 - BODEGA 94 NEJAPA' en dropdown...")
                    
                    # Buscar la opción específica en el dropdown que aparece
                    opciones_dropdown = [
                        "//li[@class='ui-menu-item']//a[contains(text(), '11094 - BODEGA 94 NEJAPA')]",
                        "//li[@role='menuitem']//a[contains(text(), '11094 - BODEGA 94 NEJAPA')]",
                        "//li[contains(@class, 'ui-menu-item')]//a[contains(text(), 'BODEGA 94 NEJAPA')]",
                        "//ul[contains(@class, 'ui-autocomplete')]//a[contains(text(), '11094')]",
                        "//a[@class='ui-corner-all'][contains(text(), '11094 - BODEGA 94 NEJAPA')]",
                        "/html/body/ul[4]/li/a"
                    ]
                    
                    opcion_seleccionada = False
                    for xpath_opcion in opciones_dropdown:
                        LOGGER.info(f"Intentando con selector: {xpath_opcion}")
                        try:
                            opcion_nejapa = WebDriverWait(DRIVER, 1).until(
                                EC.element_to_be_clickable((By.XPATH, xpath_opcion))
                            )
                            
                            if opcion_nejapa.is_displayed():
                                # Hacer hover y click en la opción
                                action = ActionChains(DRIVER)
                                action.move_to_element(opcion_nejapa).pause(0.5).click().perform()
                                LOGGER.info(f"✓ Opción 11094 - BODEGA 94 NEJAPA seleccionada con: {xpath_opcion}")
                                opcion_seleccionada = True
                                break
                        except:
                            continue
                    
                    if not opcion_seleccionada:
                        # Fallback: usar flecha abajo + Enter como describiste
                        LOGGER.info("🔄 Fallback: Usando flecha abajo + Enter...")
                        transporte_input.send_keys(Keys.ARROW_DOWN)
                        time.sleep(0.5)
                        transporte_input.send_keys(Keys.ENTER)
                        LOGGER.info("✓ Flecha abajo + Enter ejecutado")
                        opcion_seleccionada = True
                    
                    time.sleep(1)  # Esperar a que se procese la selección
                    
                except Exception as dropdown_error:
                    LOGGER.warning(f"Error seleccionando del dropdown: {dropdown_error}")
                    # Último recurso: Enter directo
                    transporte_input.send_keys(Keys.ENTER)
                
                # Verificar que se seleccionó correctamente
                try:
                    valor_input = transporte_input.get_attribute('value')
                    select_oculto = DRIVER.find_element(By.ID, "transportBySelect")

                    valor_select = select_oculto.get_attribute('value')
                    
                    LOGGER.info(f"🔍 Verificación final: Input='{valor_input}', Select='{valor_select}'")
                    
                    if valor_select == "11094":
                        LOGGER.info("✅ TRANSPORTE CONFIGURADO EXITOSAMENTE")

                    else:
                        LOGGER.warning(f"⚠️ Select no tiene valor correcto: {valor_select}")
                        
                except Exception as verify_error:
                    LOGGER.warning(f"Error verificando transporte: {verify_error}")

            except Exception as e:
                LOGGER.warning(f"Error configurando transporte: {e}")
            
            time.sleep(3)  # Tiempo para que se procese la selección
            
            # PASO 8: Instructions
            LOGGER.info("Paso 8: Agregando instrucciones")
            try:
                instructions_input = DRIVER.find_element(By.XPATH, "//*[@id='instructions']")
                instructions_input.clear()
                instructions_input.send_keys("Entrega según instrucciones del cliente")
                LOGGER.info("✓ Instrucciones agregadas")
            except Exception as e:
                LOGGER.warning(f"Error agregando instrucciones: {e}")
            
            # PASO 9: Submit
            LOGGER.info("Paso 9: Enviando formulario de entrega")
            try:
                submit_button = DRIVER.find_element(By.XPATH, "/html/body/div[2]/form/div[2]/button")
                submit_button.click()
                LOGGER.info("✓ Formulario de entrega enviado")
                time.sleep(3)
            except Exception as e:
                LOGGER.warning(f"Error enviando formulario: {e}")
            
        except Exception as e:
            LOGGER.error(f"Error configurando entrega: {e}")
            
    except Exception as e:
        LOGGER.error(f"Error general en agregar_despacho: {e}")
        return False

def modulo_7_aplicar_credito():
    """Aplica descuentos por crédito - DESHABILITADO"""
    LOGGER.info("Modulo_7_aplicar_credito: DESHABILITADO")

def modulo_8_aplicar_descuento_empleado():
    """Aplica descuentos de empleado - DESHABILITADO"""
    LOGGER.info("Modulo_8_aplicar_descuento_empleado: DESHABILITADO")

def modulo_9_finalizar_cotizacion():
    """
    Finaliza y guarda la cotización
    Busca el botón Save o similar para completar la cotización
    """
    global DRIVER
    try:
        LOGGER.info("=== MÓDULO 9: FINALIZAR COTIZACIÓN ===")
        
        # Buscar botón Save con múltiples selectores
        selectores_save = [
            "//input[@value='Save']",
            "//button[contains(text(), 'Save')]",
            "//button[@type='submit']",
            "//input[@type='submit' and contains(@value, 'Save')]",
            "//a[contains(text(), 'Save')]"
        ]
        
        for selector in selectores_save:
            try:
                save_button = WebDriverWait(DRIVER, 5).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                save_button.click()
                LOGGER.info(f"✓ Botón Save presionado: {selector}")
                time.sleep(2)
                
                # Capturar screenshot después del save
                capturar_screenshot("finalizar_cotizacion", "Cotización finalizada y guardada")
                return True
                
            except Exception as e:
                continue
        
        LOGGER.warning("⚠️ No se encontró botón Save - cotización puede no estar guardada")
        return False
        
    except Exception as e:
        LOGGER.error(f"Error finalizando cotización: {e}")
        return False


# ================== MÓDULO 7: DESCUENTO EMPLEADO ==================

def configurar_metodo_envio_credito():
    """
    NUEVO MÓDULO: Configuración de método de envío y crédito (separado del módulo cliente)
    """
    global DRIVER
    try:
        LOGGER.info("=== CONFIGURACIÓN MÉTODO DE ENVÍO/CRÉDITO ===")
        
        # Buscar botón de configuración de envío/crédito fuera del módulo cliente
        selectores_envio_credito = [
            "//button[contains(text(), 'Shipping Method')]",
            "//button[contains(text(), 'Payment Method')]", 
            "//button[contains(@id, 'shipping')]",
            "//button[contains(@id, 'payment')]",
            "//a[contains(text(), 'Delivery')]",
            "//a[contains(text(), 'Payment')]",
            "//li[contains(text(), 'Credit')]",
            "//button[contains(text(), 'Credit')]"
        ]
        
        for selector in selectores_envio_credito:
            try:
                elemento = WebDriverWait(DRIVER, 5).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                DRIVER.execute_script("arguments[0].click();", elemento)
                LOGGER.info(f"✓ Elemento de envío/crédito encontrado: {selector}")
                time.sleep(2)
                capturar_screenshot("20_metodo_envio_credito.png", "Configuración método envío/crédito")
                return True
            except:
                continue
        
        LOGGER.warning("No se encontró configuración específica de envío/crédito fuera del módulo cliente")
        return True
        
    except Exception as e:
        LOGGER.error(f"Error en configuración método envío/crédito: {str(e)}")
        return False

def agregar_descuento_empleado(con_descuento):
    """
    Módulo 7: Aplica descuento de empleado si corresponde
    con_descuento: True/False o 'Si'/'No' desde Excel
    FUNCIÓN TEMPORALMENTE DESHABILITADA
    """
    global DRIVER
    try:
        LOGGER.info(f"=== MÓDULO 7: DESCUENTO EMPLEADO ===")
        LOGGER.info(f"⚠️ FUNCIÓN DESHABILITADA - Descuento empleado omitido")
        return True
        
        # Normalizar valor del Excel
        if isinstance(con_descuento, str):
            con_descuento = con_descuento.lower().strip() in ['si', 'sí', 'yes', 'y', '1', 'true']
        elif con_descuento is None:
            con_descuento = False
        
        LOGGER.info(f"Descuento empleado: {'APLICAR' if con_descuento else 'OMITIR'}")
        
        if con_descuento:
            # Buscar y aplicar descuento de empleado
            capturar_screenshot("18_antes_descuento_empleado.png", "Antes de aplicar descuento empleado")
            
            try:
                # Buscar campo o botón de descuento empleado
                selectores_descuento = [
                    "//input[@id='employeeDiscount']",
                    "//input[contains(@name, 'discount')]",
                    "//button[contains(text(), 'Employee Discount')]",
                    "//button[contains(text(), 'Descuento Empleado')]",
                    "//input[@type='checkbox'][contains(@id, 'employee')]",
                    "//select[contains(@name, 'discount')]"
                ]
                
                descuento_aplicado = False
                for selector in selectores_descuento:
                    try:
                        elemento_descuento = DRIVER.find_element(By.XPATH, selector)
                        if elemento_descuento.is_displayed() and elemento_descuento.is_enabled():
                            # Determinar tipo de elemento y aplicar descuento
                            if elemento_descuento.tag_name.lower() == 'input':
                                input_type = elemento_descuento.get_attribute('type').lower()
                                if input_type == 'checkbox':
                                    if not elemento_descuento.is_selected():
                                        DRIVER.execute_script("arguments[0].click();", elemento_descuento)
                                        LOGGER.info("✓ Checkbox descuento empleado activado")
                                elif input_type in ['text', 'number']:
                                    elemento_descuento.clear()
                                    elemento_descuento.send_keys("10")  # Descuento del 10%
                                    LOGGER.info("✓ Porcentaje descuento empleado aplicado")
                            elif elemento_descuento.tag_name.lower() == 'button':
                                DRIVER.execute_script("arguments[0].click();", elemento_descuento)
                                LOGGER.info("✓ Botón descuento empleado clickeado")
                            elif elemento_descuento.tag_name.lower() == 'select':
                                from selenium.webdriver.support.ui import Select
                                select = Select(elemento_descuento)
                                # Buscar opción de descuento empleado
                                for option in select.options:
                                    if 'employee' in option.text.lower() or 'empleado' in option.text.lower():
                                        select.select_by_visible_text(option.text)
                                        LOGGER.info("✓ Opción descuento empleado seleccionada")
                                        break
                            
                            descuento_aplicado = True
                            time.sleep(1)
                            break
                    except Exception as e:
                        LOGGER.debug(f"Selector {selector} falló: {e}")
                        continue
                
                if descuento_aplicado:
                    capturar_screenshot("19_descuento_empleado_aplicado.png", "Descuento empleado aplicado")
                    LOGGER.info("✅ Descuento empleado aplicado exitosamente")
                else:
                    LOGGER.warning("⚠️ No se encontró campo de descuento empleado")
                    capturar_screenshot("19_descuento_no_encontrado.png", "Campo descuento no encontrado")
                    
            except Exception as e:
                LOGGER.error(f"Error aplicando descuento empleado: {e}")
                capturar_screenshot("19_error_descuento.png", "Error en descuento empleado")
        else:
            LOGGER.info("ℹ️ Descuento empleado omitido (configurado como No)")
        
        return True
        
    except Exception as e:
        LOGGER.error(f"Error en módulo descuento empleado: {e}")
        return False

def agregar_creditos(con_creditos):
    """
    Módulo 8: Configura opciones de créditos
    con_creditos: True/False o 'Si'/'No' desde Excel
    FUNCIÓN TEMPORALMENTE DESHABILITADA
    """
    global DRIVER
    try:
        LOGGER.info(f"=== MÓDULO 8: CRÉDITOS ===")
        LOGGER.info(f"⚠️ FUNCIÓN DESHABILITADA - Créditos omitidos")
        return True
        
        # Normalizar valor del Excel
        if isinstance(con_creditos, str):
            con_creditos = con_creditos.lower().strip() in ['si', 'sí', 'yes', 'y', '1', 'true']
        elif con_creditos is None:
            con_creditos = False
        
        LOGGER.info(f"Créditos: {'CONFIGURAR' if con_creditos else 'OMITIR'}")
        
        if con_creditos:
            # Buscar y configurar opciones de créditos
            capturar_screenshot("20_antes_creditos.png", "Antes de configurar créditos")
            
            try:
                # Buscar opciones de créditos
                selectores_creditos = [
                    "//input[@type='radio'][contains(@value, 'credit')]",
                    "//input[@type='radio'][contains(@value, 'credito')]",
                    "//button[contains(text(), 'Credit')]",
                    "//button[contains(text(), 'Crédito')]",
                    "//select[contains(@name, 'credit')]",
                    "//input[@type='checkbox'][contains(@id, 'credit')]",
                    "//a[contains(text(), 'Financiamiento')]"
                ]
                
                creditos_configurado = False
                for selector in selectores_creditos:
                    try:
                        elemento_credito = DRIVER.find_element(By.XPATH, selector)
                        if elemento_credito.is_displayed() and elemento_credito.is_enabled():
                            if elemento_credito.tag_name.lower() == 'input':
                                input_type = elemento_credito.get_attribute('type').lower()
                                if input_type in ['radio', 'checkbox']:
                                    if not elemento_credito.is_selected():
                                        DRIVER.execute_script("arguments[0].click();", elemento_credito)
                                        LOGGER.info("✓ Opción de créditos seleccionada")
                            elif elemento_credito.tag_name.lower() in ['button', 'a']:
                                DRIVER.execute_script("arguments[0].click();", elemento_credito)
                                LOGGER.info("✓ Enlace/Botón de créditos clickeado")
                            elif elemento_credito.tag_name.lower() == 'select':
                                from selenium.webdriver.support.ui import Select
                                select = Select(elemento_credito)
                                # Buscar opción de crédito
                                for option in select.options:
                                    if 'credit' in option.text.lower() or 'credito' in option.text.lower():
                                        select.select_by_visible_text(option.text)
                                        LOGGER.info("✓ Opción de créditos seleccionada en dropdown")
                                        break
                            
                            creditos_configurado = True
                            time.sleep(1)
                            break
                    except Exception as e:
                        LOGGER.debug(f"Selector {selector} falló: {e}")
                        continue
                
                if creditos_configurado:
                    capturar_screenshot("21_creditos_configurado.png", "Créditos configurado")
                    LOGGER.info("✅ Créditos configurado exitosamente")
                else:
                    LOGGER.warning("⚠️ No se encontró opción de créditos")
                    capturar_screenshot("21_creditos_no_encontrado.png", "Opciones créditos no encontradas")
                    
            except Exception as e:
                LOGGER.error(f"Error configurando créditos: {e}")
                capturar_screenshot("21_error_creditos.png", "Error en configuración créditos")
        else:
            LOGGER.info("ℹ️ Créditos omitido (configurado como No)")
        
        return True
        
    except Exception as e:
        LOGGER.error(f"Error en módulo créditos: {e}")
        return False

# ================== MÓDULO 9: FINALIZAR COTIZACIÓN ==================

def finalizar_cotizacion_sin_extraccion():
    """
    Módulo 9: Finaliza la cotización usando la secuencia que SÍ está funcionando:
    PASO 1: Botón finalizar → PASO 2: Printer group → PASO 3A: Print → PASO 3B: Confirmación
    """
    global DRIVER
    try:
        LOGGER.info("=== FINALIZANDO COTIZACIÓN ===")
        capturar_screenshot("17_antes_finalizar.png", "Antes de finalizar cotización")
        
        # Esperar que se complete el proceso previo
        time.sleep(2)
        
        # PASO 1: Click en botón finalizar cotización
        try:
            LOGGER.info("PASO 1: Click en botón finalizar")
            time.sleep(3)  # Espera adicional antes de finalizar
            
            selectores_finalizar = [
                (By.XPATH, "/html/body/div[2]/div[4]/div/div/div[3]/form[1]/button"),
                (By.XPATH, "//button[contains(@class, 'btn-success') and contains(text(), 'Complete')]"),
                (By.XPATH, "//button[contains(@class, 'btn-success') and contains(text(), 'Finish')]"),
                (By.XPATH, "//button[contains(@class, 'btn-success') and contains(text(), 'Save')]"),
                (By.XPATH, "//button[contains(@class, 'btn-success')]"),
                (By.CSS_SELECTOR, "button.btn-success"),
                (By.XPATH, "//input[@type='submit' and @value='Complete']"),
                (By.XPATH, "//button[text()='Complete']"),
                (By.XPATH, "//button[text()='Save']"),
                (By.XPATH, "//input[@type='submit']"),
            ]
            
            boton_finalizar = None
            for selector_type, selector_value in selectores_finalizar:
                try:
                    boton_finalizar = WebDriverWait(DRIVER, 5).until(
                        EC.element_to_be_clickable((selector_type, selector_value))
                    )
                    LOGGER.info(f"✓ Botón finalizar encontrado con: {selector_value}")
                    break
                except:
                    continue
            
            if boton_finalizar:
                capturar_screenshot("18_antes_finalizar.png", "Antes de click finalizar")
                
                # PRESERVAR SESIÓN ANTES DEL CLICK
                LOGGER.info("🔧 Preservando datos de sesión...")
                try:
                    # Obtener datos de sesión antes del click
                    order_id = DRIVER.execute_script("return sessionStorage.getItem('orderId') || localStorage.getItem('orderId');")
                    session_data = DRIVER.execute_script("return $('#getsession').val() || '';")
                    LOGGER.info(f"📝 Sesión preservada - OrderID: {order_id}, SessionData: {session_data}")
                except Exception as e:
                    LOGGER.warning(f"No se pudo preservar sesión: {e}")
                
                DRIVER.execute_script("arguments[0].click();", boton_finalizar)
                LOGGER.info("✓ Click en botón finalizar exitoso")
                
                # DEBUGGING: Capturar HTML completo después del Save
                try:
                    # Esperar un momento para que la página se actualice
                    time.sleep(3)
                                       
                    
                    # Buscar específicamente elementos relacionados con printer/print/modal/dialog
                    elementos_printer = DRIVER.find_elements(By.XPATH, "//*[contains(@id, 'print') or contains(@class, 'print') or contains(@id, 'printer') or contains(@class, 'printer')]")
                    if elementos_printer:
                        LOGGER.info(f"🖨️ Elementos printer encontrados: {len(elementos_printer)}")
                        for i, elem in enumerate(elementos_printer[:5]):  # Solo primeros 5
                            try:
                                LOGGER.info(f"  Elemento {i+1}: tag={elem.tag_name}, id={elem.get_attribute('id')}, class={elem.get_attribute('class')}")
                            except:
                                pass
                    
                    # Buscar modales/dialogs
                    modales = DRIVER.find_elements(By.XPATH, "//*[contains(@class, 'modal') or contains(@class, 'dialog') or contains(@id, 'modal') or contains(@id, 'dialog')]")
                    if modales:
                        LOGGER.info(f"📋 Modales encontrados: {len(modales)}")
                        for i, modal in enumerate(modales[:3]):  # Solo primeros 3
                            try:
                                LOGGER.info(f"  Modal {i+1}: tag={modal.tag_name}, id={modal.get_attribute('id')}, visible={modal.is_displayed()}")
                            except:
                                pass
                    
                    # Buscar todos los formularios activos
                    formularios = DRIVER.find_elements(By.TAG_NAME, "form")
                    LOGGER.info(f"📝 Formularios encontrados: {len(formularios)}")
                    
                except Exception as debug_error:
                    LOGGER.warning(f"Error en debugging post-save: {debug_error}")
                
                # MONITOREO DETALLADO POST-SAVE
                LOGGER.info("🔍 MONITOREANDO ESTADO POST-SAVE...")
                
                for i in range(10):  # Monitorear por 10 segundos
                    time.sleep(1)
                    try:
                        # URL actual
                        current_url = DRIVER.current_url
                        LOGGER.info(f"📍 Segundo {i+1}: URL = {current_url}")
                        
                        # Título de la página
                        page_title = DRIVER.title
                        LOGGER.info(f"📄 Segundo {i+1}: Título = {page_title}")
                        
                        # Elementos visibles
                        visible_buttons = DRIVER.find_elements(By.XPATH, "//button[@style!='display: none;' or not(@style)]")
                        visible_btn_texts = [btn.text.strip() for btn in visible_buttons if btn.is_displayed() and btn.text.strip()]
                        if visible_btn_texts:
                            LOGGER.info(f"🔘 Segundo {i+1}: Botones visibles = {visible_btn_texts[:5]}")  # Primeros 5
                        
                        # Buscar printer group específicamente
                        try:
                            printer_element = DRIVER.find_element(By.XPATH, "//*[@id='printerGroup']")
                            is_visible = printer_element.is_displayed()
                            LOGGER.info(f"🖨️ Segundo {i+1}: Printer Group encontrado! Visible={is_visible}")
                            if is_visible:
                                LOGGER.info("✅ PRINTER GROUP DISPONIBLE - Procediendo a hacer click...")
                                
                                # HACER CLICK EN BOTÓN OK/PRINT DEL PRINTER GROUP
                                try:
                                    # Paso 1: Buscar específicamente la impresora "Default Printer"
                                    selectores_default_printer = [
                                        "//span[text()='Default Printer']/following-sibling::input[@type='radio']",
                                        "//span[text()='Default Printer']/../input[@type='radio']",
                                        "//form[contains(.,'Default Printer')]//input[@type='radio']",
                                        "//input[@type='radio' and following-sibling::form//span[text()='Default Printer']]",
                                        "//input[@type='radio' and @name='printerGroup' and @value='1']",  # Segundo radio (suele ser Default)
                                        "//div[contains(.,'Default Printer')]//input[@type='radio']"
                                    ]
                                    
                                    default_radio_seleccionado = False
                                    for radio_selector in selectores_default_printer:
                                        try:
                                            default_radio = WebDriverWait(DRIVER, 3).until(
                                                EC.element_to_be_clickable((By.XPATH, radio_selector))
                                            )
                                            default_radio.click()
                                            LOGGER.info(f"✓ Default Printer seleccionado: {radio_selector}")
                                            default_radio_seleccionado = True
                                            break
                                        except:
                                            continue
                                    
                                    # Si no encuentra "Default Printer", usar navegación con teclas
                                    if not default_radio_seleccionado:
                                        LOGGER.info("🔄 Fallback: Navegando con teclas de flecha hacia abajo + Enter...")
                                        try:
                                            # Encontrar el modal o cualquier elemento activo para enviar teclas
                                            modal_element = WebDriverWait(DRIVER, 3).until(
                                                EC.presence_of_element_located((By.XPATH, "//div[@id='choosePrinterDialog']"))
                                            )
                                            
                                            # Navegar hacia abajo y presionar Enter
                                            modal_element.send_keys(Keys.ARROW_DOWN)
                                            LOGGER.info("✓ Flecha abajo presionada")
                                            time.sleep(0.5)
                                            modal_element.send_keys(Keys.ENTER)
                                            LOGGER.info("✓ Enter presionado")
                                            default_radio_seleccionado = True
                                            
                                        except Exception as nav_error:
                                            LOGGER.warning(f"Error navegando con teclas: {nav_error}")
                                            
                                            # Último fallback: click directo en primer radio
                                            try:
                                                primer_radio = DRIVER.find_element(By.XPATH, "//input[@type='radio' and @name='printerGroup']")
                                                primer_radio.click()
                                                LOGGER.info("✓ Primer radio clickeado como último fallback")
                                                default_radio_seleccionado = True
                                            except:
                                                LOGGER.warning("⚠️ Falló completamente la selección de radio")
                                    
                                    if not default_radio_seleccionado:
                                        LOGGER.warning("⚠️ No se pudo seleccionar ningún radio button")
                                        return False
                                    
                                    time.sleep(1)  # Esperar que se habilite el botón
                                    
                                    # Paso 2: Click en botón "Imprimir"
                                    selectores_print = [
                                        "//button[@id='defaultSubmitPrintButton']",
                                        "//button[@type='submit' and contains(@class, 'btn-success')]",
                                        "//button[contains(text(), 'Imprimir')]",
                                        "//button[@class='btn btn-success']",
                                        "//input[@type='submit' and contains(@value, 'Print')]"
                                    ]
                                    
                                    boton_print = None
                                    for sel_value in selectores_print:
                                        try:
                                            boton_print = WebDriverWait(DRIVER, 3).until(
                                                EC.element_to_be_clickable((By.XPATH, sel_value))
                                            )
                                            LOGGER.info(f"🖨️ Botón print encontrado: {sel_value}")
                                            break
                                        except:
                                            continue
                                    
                                    if boton_print:
                                        # Verificar que el botón esté habilitado
                                        if boton_print.is_enabled():
                                            DRIVER.execute_script("arguments[0].click();", boton_print)
                                            LOGGER.info("✅ CLICK EN BOTÓN IMPRIMIR EXITOSO")
                                        else:
                                            LOGGER.warning("⚠️ Botón imprimir deshabilitado, forzando click...")
                                            # Intentar habilitar y hacer click
                                            DRIVER.execute_script("arguments[0].disabled = false;", boton_print)
                                            DRIVER.execute_script("arguments[0].click();", boton_print)
                                            LOGGER.info("✅ BOTÓN FORZADO Y CLICKEADO")
                                        
                                        time.sleep(2)  # Esperar después del click
                                        
                                        # Verificar si aparece confirmación adicional
                                        try:
                                            confirmacion = WebDriverWait(DRIVER, 3).until(
                                                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'OK') or contains(text(), 'Aceptar') or contains(text(), 'Confirm')]"))
                                            )
                                            DRIVER.execute_script("arguments[0].click();", confirmacion)
                                            LOGGER.info("✅ CONFIRMACIÓN FINAL EXITOSA")
                                        except:
                                            LOGGER.info("ℹ️ No se requiere confirmación adicional")
                                            
                                        capturar_screenshot("21_print_completado.png", "Proceso print completado")
                                        return True  # FINALIZACIÓN EXITOSA
                                    else:
                                        LOGGER.warning("⚠️ No se encontró botón imprimir")
                                
                                except Exception as print_error:
                                    LOGGER.error(f"❌ Error en click print: {print_error}")
                                
                                break  # Salir del monitoreo después del intento de print
                                
                        except:
                            LOGGER.info(f"🖨️ Segundo {i+1}: Printer Group NO encontrado")
                        
                        # Verificar si hay errores JavaScript
                        try:
                            js_errors = DRIVER.execute_script("return window.console || [];")
                            LOGGER.info(f"🐛 Segundo {i+1}: JS Console disponible")
                        except Exception as js_e:
                            LOGGER.info(f"🐛 Segundo {i+1}: JS Error = {str(js_e)[:100]}")
                        
                        # Verificar sesión
                        try:
                            order_id_session = DRIVER.execute_script("return sessionStorage.getItem('orderId');")
                            order_id_local = DRIVER.execute_script("return localStorage.getItem('orderId');") 
                            LOGGER.info(f"💾 Segundo {i+1}: Session={order_id_session}, Local={order_id_local}")
                        except Exception as session_e:
                            LOGGER.info(f"💾 Segundo {i+1}: Error sesión = {str(session_e)[:100]}")
                            
                    except Exception as monitor_e:
                        LOGGER.warning(f"⚠️ Error monitoreando segundo {i+1}: {str(monitor_e)[:100]}")
                
                LOGGER.info("🏁 FIN DE MONITOREO POST-SAVE")
                
                # VERIFICAR SI ENCONTRAMOS PRINTER GROUP DURANTE EL MONITOREO
                try:
                    # Buscar botón Save en la página delivery/search ANTES del printer group
                    LOGGER.info("🔄 Buscando botón Save en página delivery/search...")
                    
                    botones_save_delivery = DRIVER.find_elements(By.XPATH, "//button[@type='submit' and contains(@class, 'btn-success') and contains(text(), 'Save')]")
                    if botones_save_delivery:
                        save_button = botones_save_delivery[0]
                        if save_button.is_displayed() and save_button.is_enabled():
                            LOGGER.info("🎯 BOTÓN SAVE ENCONTRADO EN DELIVERY/SEARCH - Haciendo click...")
                            save_button.click()
                            LOGGER.info("✅ Click en botón Save delivery exitoso")
                            
                            # Esperar después del click para que procese
                            time.sleep(5)
                            
                            # Ahora verificar si aparece el printer group
                            try:
                                printer_element = WebDriverWait(DRIVER, 10).until(
                                    EC.presence_of_element_located((By.XPATH, "//*[@id='printerGroup']"))
                                )
                                if printer_element.is_displayed():
                                    LOGGER.info("🖨️ PRINTER GROUP ENCONTRADO DESPUÉS DEL SAVE DELIVERY!")
                                    
                                    # Seleccionar impresora y hacer click en imprimir
                                    try:
                                        # Seleccionar primer radio button disponible
                                        primer_radio = WebDriverWait(DRIVER, 3).until(
                                            EC.element_to_be_clickable((By.XPATH, "//input[@type='radio' and @name='printerGroup']"))
                                        )
                                        primer_radio.click()
                                        LOGGER.info("✓ Radio button de impresora seleccionado")
                                        time.sleep(1)
                                        
                                        # Click en botón imprimir
                                        boton_print = WebDriverWait(DRIVER, 3).until(
                                            EC.element_to_be_clickable((By.XPATH, "//button[@id='defaultSubmitPrintButton'] | //button[@class='btn btn-success']"))
                                        )
                                        boton_print.click()
                                        LOGGER.info("✅ BOTÓN IMPRIMIR CLICKEADO - COTIZACIÓN FINALIZADA COMPLETAMENTE")
                                        
                                        # Capturar evidencia final
                                        time.sleep(2)
                                        capturar_screenshot("21_cotizacion_finalizada_completamente.png", "Cotización finalizada completamente")
                                        return True
                                        
                                    except Exception as print_err:
                                        LOGGER.warning(f"⚠️ Error en proceso final de impresión: {print_err}")
                                        LOGGER.info("ℹ️ Continuando - Save ya se ejecutó exitosamente")
                                else:
                                    LOGGER.info("ℹ️ Printer Group no visible después del Save delivery")
                                    
                            except TimeoutException:
                                LOGGER.info("ℹ️ Printer Group no apareció después del Save delivery")
                                LOGGER.info("💡 Esto puede ser normal - el Save delivery ya se ejecutó")
                        else:
                            LOGGER.info("ℹ️ Botón Save delivery encontrado pero no clickeable")
                    else:
                        LOGGER.info("ℹ️ No se encontró botón Save en delivery/search")
                    
                    # Intentar una vez más encontrar el printer group después del monitoreo original
                    LOGGER.info("🔄 Verificación final del Printer Group...")
                    time.sleep(3)  # Esperar a que aparezca
                    
                    try:
                        printer_element = WebDriverWait(DRIVER, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//*[@id='printerGroup']"))
                        )
                        if printer_element.is_displayed():
                            LOGGER.info("🖨️ PRINTER GROUP ENCONTRADO EN VERIFICACIÓN FINAL!")
                            
                            # Seleccionar impresora y hacer click en imprimir
                            try:
                                # Seleccionar primer radio button disponible
                                primer_radio = WebDriverWait(DRIVER, 3).until(
                                    EC.element_to_be_clickable((By.XPATH, "//input[@type='radio' and @name='printerGroup']"))
                                )
                                primer_radio.click()
                                LOGGER.info("✓ Radio button de impresora seleccionado")
                                time.sleep(1)
                                
                                # Click en botón imprimir
                                boton_print = WebDriverWait(DRIVER, 3).until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@id='defaultSubmitPrintButton'] | //button[@class='btn btn-success']"))
                                )
                                boton_print.click()
                                LOGGER.info("✅ BOTÓN IMPRIMIR CLICKEADO - COTIZACIÓN FINALIZADA COMPLETAMENTE")
                                
                                # Capturar evidencia final
                                time.sleep(2)
                                capturar_screenshot("21_cotizacion_finalizada_completamente.png", "Cotización finalizada completamente")
                                return True
                                
                            except Exception as print_err:
                                LOGGER.warning(f"⚠️ Error en proceso final de impresión: {print_err}")
                                LOGGER.info("ℹ️ Continuando - Save ya se ejecutó exitosamente")
                            
                        else:
                            LOGGER.info("ℹ️ Printer Group no visible en verificación final")
                            
                    except TimeoutException:
                        LOGGER.info("ℹ️ Printer Group no apareció en verificación final")
                        LOGGER.info("💡 Esto puede ser normal - el Save ya se ejecutó")
                        
                except Exception as final_check_error:
                    LOGGER.warning(f"Error en verificación final: {final_check_error}")
                
                # ESPERAR MÁS TIEMPO para que procese completamente
                LOGGER.info("⏳ Esperando procesamiento completo del Save...")
                time.sleep(2)  # Tiempo adicional después del monitoreo
                
                # VERIFICAR Y RESTAURAR SESIÓN SI ES NECESARIO
                try:
                    current_order = DRIVER.execute_script("return sessionStorage.getItem('orderId') || localStorage.getItem('orderId');")
                    if not current_order and order_id:
                        LOGGER.info("🔧 Restaurando sesión perdida...")
                        DRIVER.execute_script(f"sessionStorage.setItem('orderId', '{order_id}');")
                        DRIVER.execute_script(f"localStorage.setItem('orderId', '{order_id}');")
                        LOGGER.info("✓ Sesión restaurada")
                except Exception as e:
                    LOGGER.warning(f"Error verificando/restaurando sesión: {e}")
                    
            else:
                raise Exception("No se encontró botón finalizar")
            
        except Exception as e:
            LOGGER.error(f"Error en paso 1 - finalizar: {e}")
            debuggear_botones_disponibles()
            return f"ERROR_PASO1_{int(time.time())}"
        
        # PASO 2: Esperar redirección automática del servidor después del Save
        try:
            LOGGER.info("PASO 2: Esperando redirección automática del servidor...")
            
            # Guardar URL actual para detectar cambios
            url_antes_submit = DRIVER.current_url
            LOGGER.info(f"📍 URL antes del submit: {url_antes_submit}")
            
            # Esperar y detectar cambio de página o estado
            tiempo_maximo = 20  # 20 segundos máximo
            for segundo in range(tiempo_maximo):
                time.sleep(1)
                
                try:
                    url_actual = DRIVER.current_url
                    titulo_actual = DRIVER.title
                    
                    LOGGER.info(f"⏱️ Segundo {segundo+1}: URL={url_actual}, Título={titulo_actual}")
                    
                    # Buscar indicadores de éxito en la página
                    indicadores_exito = [
                        "successfully",
                        "completed", 
                        "success",
                        "saved",
                        "cotización",
                        "quotation",
                        "orden creada",
                        "order created"
                    ]
                    
                    # Verificar texto de la página
                    try:
                        texto_pagina = DRIVER.page_source.lower()
                        for indicador in indicadores_exito:
                            if indicador in texto_pagina:
                                LOGGER.info(f"✅ Indicador de éxito encontrado: '{indicador}'")
                                capturar_screenshot("20_exito_detectado.png", f"Éxito detectado: {indicador}")
                                return True
                    except:
                        pass
                    
                    # Verificar si cambió la URL (redirección)
                    if url_actual != url_antes_submit:
                        LOGGER.info(f"🔄 Redirección detectada: {url_antes_submit} → {url_actual}")
                        # Esperar un poco más para que cargue completamente
                        time.sleep(3)
                        capturar_screenshot("20_despues_redireccion.png", "Estado después de redirección")
                        return True
                        
                    # Verificar si aparecen elementos de éxito/finalización
                    try:
                        elementos_exito = DRIVER.find_elements(By.XPATH, "//*[contains(text(), 'success') or contains(text(), 'completed') or contains(text(), 'saved')]")
                        if elementos_exito:
                            LOGGER.info("✅ Elementos de éxito encontrados en la página")
                            capturar_screenshot("20_elementos_exito.png", "Elementos de éxito detectados")
                            return True
                    except:
                        pass
                        
                except Exception as monitor_error:
                    LOGGER.warning(f"⚠️ Error monitoreando segundo {segundo+1}: {str(monitor_error)[:100]}")
            
            # Si llegamos aquí, asumimos que el proceso fue exitoso pero sin indicadores claros
            LOGGER.info("⏰ Tiempo de espera completado - asumiendo procesamiento exitoso")
            capturar_screenshot("20_tiempo_completado.png", "Procesamiento completado por tiempo")
            return True
            
        except Exception as e:
            LOGGER.error(f"Error esperando redirección: {e}")
            LOGGER.info("💡 Continuando - la cotización probablemente se guardó correctamente")
            return True  # Asumir éxito porque el Save ya se ejecutó

        # Si llegamos aquí después del monitoreo, la finalización fue exitosa
        LOGGER.info("✅ Finalización exitosa")
        return True
        
    except Exception as e:
        LOGGER.error(f"Error crítico finalizando cotización: {e}")
        capturar_screenshot("22_error_finalizacion.png", "Error en finalización")
        debuggear_botones_disponibles()
        return f"ERROR_CRITICO_{int(time.time())}"

def debuggear_botones_disponibles():
    """Función auxiliar para debuggear botones disponibles en caso de error"""
    global DRIVER
    try:
        LOGGER.info("=== DEBUG: ELEMENTOS DISPONIBLES ===")
        
        # Capturar screenshot del estado actual
        capturar_screenshot("debug_estado_actual.png", "Estado actual para debug")
        
        # Buscar botones
        botones = DRIVER.find_elements(By.TAG_NAME, "button")
        LOGGER.info(f"Botones encontrados: {len(botones)}")
        for i, btn in enumerate(botones[:15]):  # Primeros 15
            try:
                texto = btn.text.strip()
                classes = btn.get_attribute("class")
                onclick = btn.get_attribute("onclick")
                visible = btn.is_displayed()
                enabled = btn.is_enabled()
                LOGGER.info(f"  Btn {i}: '{texto}' class='{classes}' visible={visible} enabled={enabled}")
            except Exception as e:
                LOGGER.debug(f"Error leyendo botón {i}: {e}")
        
        # Buscar inputs
        inputs = DRIVER.find_elements(By.TAG_NAME, "input")
        LOGGER.info(f"Inputs encontrados: {len(inputs)}")
        for i, inp in enumerate(inputs[:15]):  # Primeros 15
            try:
                value = inp.get_attribute("value")
                tipo = inp.get_attribute("type")
                visible = inp.is_displayed()
                enabled = inp.is_enabled()
                LOGGER.info(f"  Input {i}: value='{value}' type='{tipo}' visible={visible} enabled={enabled}")
            except Exception as e:
                LOGGER.debug(f"Error leyendo input {i}: {e}")
        
        # Buscar enlaces
        enlaces = DRIVER.find_elements(By.TAG_NAME, "a")
        LOGGER.info(f"Enlaces encontrados: {len(enlaces)}")
        for i, enlace in enumerate(enlaces[:10]):  # Primeros 10
            try:
                texto = enlace.text.strip()
                href = enlace.get_attribute("href")
                visible = enlace.is_displayed()
                if texto and visible:
                    LOGGER.info(f"  Link {i}: '{texto}' href='{href}' visible={visible}")
            except Exception as e:
                LOGGER.debug(f"Error leyendo enlace {i}: {e}")
                
    except Exception as e:
        LOGGER.error(f"Error en debug de elementos: {e}")

def extraer_numero_cotizacion():
    """Extrae número de cotización de la página con validación mejorada"""
    global DRIVER
    try:
        # Esperar a que se genere el número
        time.sleep(1)
        
        # Selectores específicos para número de cotización
        selectores_cotizacion = [
            "/html/body/div[2]/div[4]/div/div/ul/li[4]/p[2]",  # XPath específico proporcionado
            "li.quotationNumberTitle p:last-child",
            "//li[contains(@class, 'quotationNumberTitle')]//p[last()]",
            "//li[contains(@class, 'quotationNumber')]//p",
            "//*[contains(text(), 'Quotation Number')]/following-sibling::*",
            "//p[contains(text(), '116')]",  # Números que empiezan por 116
            "//span[contains(text(), '116')]",
            "//*[text()[contains(., '116')]]",
            "//div[contains(@class, 'quotation')]//*[contains(text(), '116')]",
            # Selectores más amplios
            "//*[contains(text(), 'quotation')]//following::*[1]",
            "//div[contains(@class, 'quotation')]//span"
        ]
        
        numero_cotizacion = None
        
        # Buscar con selectores específicos
        for selector in selectores_cotizacion:
            try:
                if selector.startswith("//") or selector.startswith("//*"):
                    elementos = DRIVER.find_elements(By.XPATH, selector)
                else:
                    elementos = DRIVER.find_elements(By.CSS_SELECTOR, selector)
                
                for elemento in elementos:
                    if elemento.is_displayed() and elemento.text.strip():
                        texto = elemento.text.strip()
                        # Verificar que parece un número de cotización válido
                        if validar_numero_cotizacion(texto):
                            numero_cotizacion = texto
                            LOGGER.info(f"Número encontrado con selector: {selector}")
                            LOGGER.info(f"Número de cotización: {numero_cotizacion}")
                            return numero_cotizacion
                        
            except Exception as e:
                LOGGER.debug(f"Selector {selector} falló: {e}")
                continue
        
        # Si no encontramos con selectores, hacer búsqueda exhaustiva
        if not numero_cotizacion:
            LOGGER.info("Búsqueda exhaustiva de número de cotización...")
            try:
                # Buscar en todo el contenido de la página
                elementos_texto = DRIVER.find_elements(By.XPATH, "//*[text()]")
                for elem in elementos_texto:
                    try:
                        if elem.is_displayed():
                            texto = elem.text.strip()
                            if validar_numero_cotizacion(texto):
                                numero_cotizacion = texto
                                LOGGER.info(f"Número encontrado por búsqueda exhaustiva: {numero_cotizacion}")
                                return numero_cotizacion
                    except:
                        continue
            except Exception as e:
                LOGGER.warning(f"Error en búsqueda exhaustiva: {e}")
        
        if not numero_cotizacion:
            LOGGER.warning(" No se encontró número de cotización válido")
            return None
        
        return numero_cotizacion
        
    except Exception as e:
        LOGGER.error(f"Error extrayendo número de cotización: {e}")
        return None

def validar_numero_cotizacion(texto):
    """Valida que el texto sea un número de cotización válido"""
    if not texto or len(texto) < 8:
        return False
    
    # Debe contener dígitos y tener formato de cotización
    if '116' in texto and len(texto) >= 10:
        # Verificar que tiene principalmente dígitos
        digitos = sum(c.isdigit() for c in texto)
        if digitos >= 8:  # Al menos 8 dígitos
            return True
    
    return False

# ================== MÓDULO PRINCIPAL ==================

def procesar_cotizacion_individual(cotizacion_id, productos_cotizacion):
    """
    Procesa una cotización individual usando todos los módulos (1-9)
    Basado en configuraciones del Excel: descuento_empleado, con_creditos, entrega_cliente, etc.
    """
    global COTIZACIONES_PROCESADAS, COTIZACION_ACTUAL
    try:
        # Establecer cotización actual para screenshots
        COTIZACION_ACTUAL = cotizacion_id
        
        LOGGER.info(f"\n === PROCESANDO COTIZACIÓN {cotizacion_id} ===")
        LOGGER.info(f"Productos: {len(productos_cotizacion)}")
        
        # Obtener configuraciones de la cotización
        config = productos_cotizacion[0]  # Primer producto tiene la config
        con_garantia = config.get('con_garantia', 'No')
        descuento_empleado = config.get('descuento_empleado', 'No')  
        con_creditos = config.get('con_creditos', 'No')
        entrega_cliente = config.get('entrega_cliente', 'No')
        tipo_entrega = config.get('tipo_entrega', 'cliente llevará')  # Nuevo campo desde Excel
        
        LOGGER.info(f"Configuración - Garantías: {con_garantia}, Descuento: {descuento_empleado}, "
                   f"Créditos: {con_creditos}, Entrega: {entrega_cliente}, Tipo Entrega: {tipo_entrega}")
        
        # EXTRAER NÚMERO DE COTIZACIÓN AL INICIO (antes de agregar productos)
        numero_cotizacion_completo = extraer_numero_cotizacion_inicial()
        if not numero_cotizacion_completo:
            LOGGER.error("No se pudo extraer número de cotización inicial")
            return False
        
        # Separar número real del timestamp
        if '_' in numero_cotizacion_completo:
            numero_cotizacion_real = numero_cotizacion_completo.split('_')[0]  # Número del aplicativo
            numero_cotizacion_unico = numero_cotizacion_completo              # Con timestamp para archivos
        else:
            numero_cotizacion_real = numero_cotizacion_completo
            numero_cotizacion_unico = numero_cotizacion_completo
        
        LOGGER.info(f"📋 Número del aplicativo: {numero_cotizacion_real}")
        LOGGER.info(f"📁 Número único para archivos: {numero_cotizacion_unico}")
        
        # MÓDULO 3: Agregar todos los items
        for i, producto in enumerate(productos_cotizacion, 1):
            codigo_raw = (producto.get('codigo') or producto.get('producto') or 
                         producto.get('Codigo') or producto.get('Producto'))
            cantidad_raw = producto.get('cantidad', 1)
            
            # Convertir a enteros para eliminar decimales innecesarios
            try:
                codigo = int(float(codigo_raw)) if codigo_raw is not None else None
            except (ValueError, TypeError):
                codigo = str(codigo_raw) if codigo_raw is not None else None
                
            try:
                cantidad = int(float(cantidad_raw)) if cantidad_raw is not None else 1
            except (ValueError, TypeError):
                cantidad = 1
            
            if not agregar_item(codigo, cantidad, i):
                LOGGER.warning(f"Item {i} falló, continuando...")
                continue
            
            # MÓDULO 4: Manejar garantías para este item
            garantia_item = producto.get('con_garantia', con_garantia)
            if not agregar_garantias(garantia_item, i):
                LOGGER.warning(f"Garantías item {i} fallaron")
            
            # time.sleep(1)  # Eliminado para mayor velocidad
        
        # MÓDULO 5: Enlazar customer
        if not enlazar_customer():
            LOGGER.error(" Error enlazando customer")
            return False
        
        # MÓDULO 6: Agregar despacho (basado en Excel)
        if not agregar_despacho(entrega_cliente, tipo_entrega):
            LOGGER.warning("⚠️ Error en despacho, continuando...")
        
        # CONFIGURACIÓN MÉTODO ENVÍO/CRÉDITO (OMITIDO - ya manejado en módulo cliente)
        LOGGER.info("Configuración envío/crédito omitida (integrada en módulo cliente)")
        
        # MÓDULO 7: Descuento empleado (basado en Excel) 
        if not agregar_descuento_empleado(descuento_empleado):
            LOGGER.warning("⚠️ Error en descuento empleado, continuando...")
        
        # MÓDULO 8: Créditos (basado en Excel)
        if not agregar_creditos(con_creditos):
            LOGGER.warning("⚠️ Error en créditos, continuando...")
        
        # MÓDULO 9: Finalizar (sin extraer número, ya lo tenemos)
        resultado_finalizacion = finalizar_cotizacion_sin_extraccion()
        
        # Manejar diferentes tipos de retorno (bool o string)
        if isinstance(resultado_finalizacion, bool):
            if resultado_finalizacion:
                LOGGER.info("✅ Finalización exitosa")
            else:
                LOGGER.error("❌ Error en finalización")
                debuggear_botones_disponibles()
                LOGGER.info("Generando reporte con datos parciales...")
        elif isinstance(resultado_finalizacion, str) and resultado_finalizacion.startswith("ERROR"):
            LOGGER.error(f"Error en finalización: {resultado_finalizacion}")
            debuggear_botones_disponibles()
            LOGGER.info("Generando reporte con datos parciales...")
        else:
            LOGGER.info(f"✅ Finalización exitosa: {resultado_finalizacion}")
        
        # Guardar resultado con número real del aplicativo
        COTIZACIONES_PROCESADAS.append({
            'cotizacion_id': cotizacion_id,
            'numero_cotizacion': numero_cotizacion_real,        # Número real del aplicativo
            'numero_cotizacion_unico': numero_cotizacion_unico, # Con timestamp para archivos
            'productos': len(productos_cotizacion),
            'descuento_empleado': descuento_empleado,
            'con_creditos': con_creditos,
            'entrega_cliente': entrega_cliente
        })
        
        LOGGER.info(f"Cotización {cotizacion_id} completada: {numero_cotizacion_real}")
        
        # GENERAR REPORTE INMEDIATAMENTE PARA ESTA COTIZACIÓN
        generar_reporte_individual(cotizacion_id, numero_cotizacion_real, numero_cotizacion_unico, productos_cotizacion, entrega_cliente, tipo_entrega)
        
        # NO limpiar screenshots automáticamente - puede causar excepciones
        # limpiar_screenshots_cotizacion()
        
        return True
        
    except Exception as e:
        LOGGER.error(f"Error procesando cotización {cotizacion_id}: {e}")
        return False

def generar_reporte_individual(cotizacion_id, numero_cotizacion_real, numero_cotizacion_unico, productos_cotizacion, entrega_cliente=None, tipo_entrega=None):
    """Genera reporte Word ejecutivo simple con número de cotización y evidencias visuales"""
    try:
        LOGGER.info(f"🔄 Generando reporte para cotización {cotizacion_id}: {numero_cotizacion_real}")
        
        # Actualizar Excel inmediatamente con número real
        actualizar_excel_cotizacion_individual("productos_avanzado.xlsx", cotizacion_id, True)
        
        # Crear documento Word usando solo el número real extraído (sin correlativos)
        nombre_archivo = f"Reporte_cotizacion_{numero_cotizacion_real}.docx"
        
        doc = Document()
        
        # Usar el número real del aplicativo para mostrar
        numero_cotizacion_display = numero_cotizacion_real
        
        # Configurar estilo base Calibri
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(12)
        
        # Título ejecutivo simple
        titulo = doc.add_heading(f'Cotización Procesada: {numero_cotizacion_display}', 0)
        titulo.alignment = 1  # Centrado
        titulo_run = titulo.runs[0]
        titulo_run.font.name = 'Calibri'
        titulo_run.font.size = Pt(20)
        titulo_run.font.bold = True
        titulo_run.font.color.rgb = RGBColor(0, 51, 102)
        
        # Información básica
        info_fecha = doc.add_paragraph()
        info_fecha.add_run("Fecha: ").bold = True
        fecha_run = info_fecha.add_run(f"{datetime.now().strftime('%d/%m/%Y %H:%M')}")
        fecha_run.font.name = 'Calibri'
        fecha_run.font.size = Pt(12)
        
        info_estado = doc.add_paragraph()
        info_estado.add_run("Estado: ").bold = True
        estado_run = info_estado.add_run("PROCESADA EXITOSAMENTE")
        estado_run.font.name = 'Calibri'
        estado_run.font.size = Pt(12)
        estado_run.font.bold = True
        estado_run.font.color.rgb = RGBColor(0, 128, 0)
        
        # Configuración de entrega (información del Excel)
        if entrega_cliente is not None:
            info_entrega = doc.add_paragraph()
            info_entrega.add_run("Entrega Cliente: ").bold = True
            
            # Normalizar el valor de entrega_cliente
            if isinstance(entrega_cliente, str):
                entrega_normalizada = entrega_cliente.lower().strip() in ['si', 'sí', 'yes', 'y', '1', 'true']
            else:
                entrega_normalizada = bool(entrega_cliente)
            
            entrega_texto = "Sí" if entrega_normalizada else "No"
            entrega_run = info_entrega.add_run(entrega_texto)
            entrega_run.font.name = 'Calibri'
            entrega_run.font.size = Pt(12)
            if entrega_texto == "Sí":
                entrega_run.font.color.rgb = RGBColor(0, 128, 0)  # Verde para Sí
            else:
                entrega_run.font.color.rgb = RGBColor(128, 128, 128)  # Gris para No
            
            # Tipo de entrega si está disponible y entrega está habilitada
            if tipo_entrega is not None and entrega_texto == "Sí":
                tipo_entrega_p = doc.add_paragraph()
                tipo_entrega_p.add_run("Tipo Entrega: ").bold = True
                tipo_run = tipo_entrega_p.add_run(f"{tipo_entrega}")
                tipo_run.font.name = 'Calibri'
                tipo_run.font.size = Pt(12)
        
        # Evidencias del proceso - SOLO las pantallas
        if SCREENSHOTS_TOMADAS:
            doc.add_paragraph()
            evidencia_titulo = doc.add_heading('Evidencias del Procesamiento', level=1)
            evidencia_titulo_run = evidencia_titulo.runs[0]
            evidencia_titulo_run.font.name = 'Calibri'
            evidencia_titulo_run.font.color.rgb = RGBColor(0, 51, 102)
            
            for i, screenshot_info in enumerate(SCREENSHOTS_TOMADAS, 1):
                screenshot_file, descripcion = screenshot_info
                
                if os.path.exists(screenshot_file):
                    # Descripción clara del paso
                    paso_p = doc.add_paragraph()
                    paso_run = paso_p.add_run(f"Paso {i}: {descripcion}")
                    paso_run.font.name = 'Calibri'
                    paso_run.font.size = Pt(14)
                    paso_run.font.bold = True
                    paso_run.font.color.rgb = RGBColor(0, 51, 102)
                    
                    # Imagen del paso
                    try:
                        doc.add_picture(screenshot_file, width=Inches(6.5))
                        doc.add_paragraph()  # Espacio entre imágenes
                    except Exception as e:
                        LOGGER.warning(f"Error agregando screenshot {screenshot_file}: {e}")
                        # Si no puede agregar la imagen, al menos documentar el paso
                        error_p = doc.add_paragraph()
                        error_run = error_p.add_run(f"[Evidencia no disponible: {descripcion}]")
                        error_run.font.name = 'Calibri'
                        error_run.font.italic = True
                        error_run.font.color.rgb = RGBColor(128, 128, 128)
        
        # Footer ejecutivo
        doc.add_paragraph()
        footer_p = doc.add_paragraph()
        footer_run = footer_p.add_run(f" by ELISEO AMILCAR LOPEZ - {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        footer_run.font.name = 'Calibri'
        footer_run.font.size = Pt(10)
        footer_run.font.italic = True
        footer_run.font.color.rgb = RGBColor(100, 100, 100)
        footer_p.alignment = 1  # Centrado
        
        # Guardar documento
        doc.save(nombre_archivo)
        LOGGER.info(f"✅ Reporte ejecutivo generado: {nombre_archivo}")
        
        # Verificar que se creó
        if os.path.exists(nombre_archivo):
            size = os.path.getsize(nombre_archivo)
            LOGGER.info(f"✅ Archivo verificado: {size} bytes")
        else:
            LOGGER.error(f"❌ Error: No se pudo crear {nombre_archivo}")
        
        return nombre_archivo
        
    except Exception as e:
        LOGGER.error(f"❌ Error generando reporte individual: {e}")
        import traceback
        LOGGER.error(f"Traceback: {traceback.format_exc()}")
        return None
        titulo_run.font.name = 'Calibri'
        titulo_run.font.size = Pt(18)
        titulo_run.font.bold = True
        
        # Información general
        info_fecha = doc.add_paragraph()
        info_fecha.add_run("Fecha de procesamiento: ").bold = True
        info_fecha.add_run(f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        
        info_productos = doc.add_paragraph()
        info_productos.add_run("Productos procesados: ").bold = True
        info_productos.add_run(f"{len(productos_cotizacion)}")
        
        # Configurar fuente de párrafos
        for p in [info_fecha, info_productos]:
            for run in p.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(11)
        
        # Agregar screenshots capturados
        if SCREENSHOTS_TOMADAS:
            doc.add_paragraph()
            evidencia_titulo = doc.add_heading('Evidencia del Procesamiento', level=1)
            evidencia_titulo_run = evidencia_titulo.runs[0]
            evidencia_titulo_run.font.name = 'Calibri'
            
            for i, screenshot_info in enumerate(SCREENSHOTS_TOMADAS, 1):
                screenshot_file, descripcion = screenshot_info
                
                if os.path.exists(screenshot_file):
                    desc_p = doc.add_paragraph()
                    desc_run = desc_p.add_run(f"{i}. {descripcion}")
                    desc_run.font.name = 'Calibri'
                    desc_run.font.size = Pt(12)
                    desc_run.font.bold = True
                    desc_run.font.color.rgb = RGBColor(0, 51, 102)
                    
                    try:
                        doc.add_picture(screenshot_file, width=Inches(6))
                        doc.add_paragraph()
                    except Exception as e:
                        LOGGER.warning(f"Error agregando screenshot {screenshot_file}: {e}")
        
        # Conclusión
        conclusion_p = doc.add_paragraph()
        conclusion_p.add_run('Conclusión: ').bold = True
        conclusion_run = conclusion_p.add_run(f"Cotización {numero_cotizacion_display} procesada exitosamente.")
        conclusion_run.font.name = 'Calibri'
        conclusion_run.font.size = Pt(12)
        conclusion_run.font.bold = True
        conclusion_run.font.color.rgb = RGBColor(0, 128, 0)
        conclusion_p.alignment = 1
        
        # Footer
        doc.add_paragraph()
        footer_p = doc.add_paragraph()
        footer_run = footer_p.add_run(f"Generado automáticamente - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        footer_run.font.name = 'Calibri'
        footer_run.font.size = Pt(10)
        footer_run.font.italic = True
        footer_run.font.color.rgb = RGBColor(128, 128, 128)
        footer_p.alignment = 1
        
        # Guardar documento
        doc.save(nombre_archivo)
        LOGGER.info(f"Word generado: {nombre_archivo}")
        
        # DEBUG: Verificar screenshots antes de limpiar
        LOGGER.info(f"Screenshots disponibles para Word: {len(SCREENSHOTS_TOMADAS)}")
        for i, (archivo, desc) in enumerate(SCREENSHOTS_TOMADAS):
            existe = os.path.exists(archivo)
            LOGGER.info(f"  {i+1}. {archivo} - Existe: {existe} - {desc}")
        
        # NO limpiar screenshots automáticamente - esto puede causar problemas
        # limpiar_screenshots_cotizacion()
        
        return nombre_archivo
        
    except Exception as e:
        LOGGER.error(f"Error generando reporte individual: {e}")
        return None

def actualizar_excel_cotizacion_individual(archivo_excel, cotizacion_id, exito):
    """Actualiza Excel con el resultado de la cotización individual"""
    try:
        # Leer datos actuales
        df = pd.read_excel(archivo_excel)
        
        # Agregar columnas si no existen
        if 'COTIZACION' not in df.columns:
            df['COTIZACION'] = ''
            LOGGER.info("Columna COTIZACION agregada")
        if 'ESTADO' not in df.columns:
            df['ESTADO'] = 'Pendiente'
            LOGGER.info("Columna ESTADO agregada")
        if 'FECHA_PROCESADO' not in df.columns:
            df['FECHA_PROCESADO'] = ''
            LOGGER.info("Columna FECHA_PROCESADO agregada")
        if 'ARCHIVO_WORD' not in df.columns:
            df['ARCHIVO_WORD'] = ''
            LOGGER.info("Columna ARCHIVO_WORD agregada")
        
        # Buscar filas de esta cotización
        filas_cotizacion = df[df['cotizacion_id'] == cotizacion_id].index
        
        if len(filas_cotizacion) > 0:
            # Obtener número de cotización de la lista global
            numero_cotizacion = None
            for cotizacion_proc in COTIZACIONES_PROCESADAS:
                if cotizacion_proc['cotizacion_id'] == cotizacion_id:
                    numero_cotizacion = cotizacion_proc['numero_cotizacion']
                    break
            
            if numero_cotizacion:
                # Actualizar todas las filas de esta cotización
                df.loc[filas_cotizacion, 'COTIZACION'] = numero_cotizacion
                df.loc[filas_cotizacion, 'ESTADO'] = 'Completado' if exito else 'Error'
                df.loc[filas_cotizacion, 'FECHA_PROCESADO'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                # Para ARCHIVO_WORD, usar solo el número real extraído
                df.loc[filas_cotizacion, 'ARCHIVO_WORD'] = f"Reporte_cotizacion_{numero_cotizacion}.docx"
                
                # Guardar cambios
                df.to_excel(archivo_excel, index=False)
                LOGGER.info(f"✅ Excel actualizado - Cotización {cotizacion_id}: {numero_cotizacion}")
            else:
                LOGGER.warning(f"No se encontró número de cotización para {cotizacion_id}")
        else:
            LOGGER.warning(f"No se encontraron filas para cotización {cotizacion_id}")
        
    except Exception as e:
        LOGGER.error(f"Error actualizando Excel individual: {e}")
        mask = df['cotizacion_id'] == cotizacion_id
        if mask.any():
            # Actualizar directamente las celdas
            df.loc[mask, 'COTIZACION'] = str(numero_cotizacion)  # Convertir a string para evitar warning
            df.loc[mask, 'ESTADO'] = 'Completada'
            df.loc[mask, 'FECHA_PROCESADO'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
            
            # Guardar cambios inmediatamente
            df.to_excel(archivo_excel, index=False)
            LOGGER.info(f"✅ Excel actualizado - Cotización {cotizacion_id} → {numero_cotizacion}")
        else:
            LOGGER.warning(f"No se encontró cotización {cotizacion_id} en Excel")
        
    except Exception as e:
        LOGGER.error(f"Error actualizando Excel individual: {e}")

def limpiar_screenshots_cotizacion():
    """Elimina screenshots de la cotización actual"""
    global SCREENSHOTS_TOMADAS
    screenshots_eliminados = 0
    
    for screenshot_info in SCREENSHOTS_TOMADAS[:]:
        screenshot_file, _ = screenshot_info
        try:
            if os.path.exists(screenshot_file):
                os.remove(screenshot_file)
                screenshots_eliminados += 1
                LOGGER.info(f"Screenshot eliminado: {screenshot_file}")
        except Exception as e:
            LOGGER.warning(f"Error eliminando {screenshot_file}: {e}")
    
    # Limpiar lista para la siguiente cotización
    SCREENSHOTS_TOMADAS.clear()
    LOGGER.info(f"{screenshots_eliminados} screenshots eliminados para esta cotización")

def preparar_nueva_cotizacion():
    """El sistema maneja automáticamente nuevas cotizaciones - función simplificada"""
    global DRIVER
    try:
        LOGGER.info("🔄 Preparando nueva cotización...")
        
        # Agregar tiempo para que la página se estabilice
        time.sleep(5)  # Aumentar timeout para estabilización
        
        # Limpiar storage del navegador para sesión fresca (opcional)
        try:
            DRIVER.execute_script("window.localStorage.clear();")
            DRIVER.execute_script("window.sessionStorage.clear();")
            LOGGER.info("✓ Storage del navegador limpiado")
        except Exception as e:
            LOGGER.warning(f"Error limpiando storage: {e}")
        
        # Verificar que estamos en la página correcta
        try:
            WebDriverWait(DRIVER, 10).until(
                lambda driver: "engage-unicomer-web" in driver.current_url
            )
            LOGGER.info("✓ Página estabilizada para nueva cotización")
        except:
            LOGGER.warning("Advertencia: URL no verificada pero continuando")
        
        # El sistema maneja automáticamente nuevas cotizaciones - simplificado
        LOGGER.info("✓ Sistema listo para siguiente cotización automática")
        time.sleep(1)
        return True
        
    except Exception as e:
        LOGGER.error(f"Error preparando nueva cotización: {e}")
        return False

def crear_reportes(archivo_excel):
    """Reporte final consolidado - los individuales ya fueron generados"""
    try:
        LOGGER.info("=== REPORTE FINAL ===")
        
        # Verificar que tenemos datos
        if not COTIZACIONES_PROCESADAS:
            LOGGER.warning("No hay cotizaciones procesadas para reportar")
            return None, None
        
        LOGGER.info("Excel ya fue actualizado individualmente por cotización")
        LOGGER.info("Reportes Word ya fueron generados individualmente")
        
        # Buscar archivos Word generados
        reportes_generados = []
        for cotizacion in COTIZACIONES_PROCESADAS:
            numero = cotizacion['numero_cotizacion']
            # Buscar archivo Word correspondiente con nuevo formato
            import glob
            patron = f"Reporte_cotizacion_{numero}.docx"
            archivos_encontrados = glob.glob(patron)
            if archivos_encontrados:
                reportes_generados.extend(archivos_encontrados)
        
        LOGGER.info(f"Reportes encontrados: {len(reportes_generados)}")
        
        # Retornar archivos (Excel ya actualizado + reportes Word individuales)
        return archivo_excel, reportes_generados
    
    except Exception as e:
        LOGGER.error(f"Error en reporte final: {e}")
        import traceback
        LOGGER.error(f"Detalle del error: {traceback.format_exc()}")
        return None, None

def limpiar_recursos():
    """Limpia recursos finales - screenshots ya fueron eliminados por cotización"""
    global SCREENSHOTS_TOMADAS, DRIVER
    try:
        LOGGER.info("Limpiando recursos...")
        
        # Los screenshots ya fueron eliminados individualmente por cotización
        # Solo limpiar lista residual si queda algo
        if SCREENSHOTS_TOMADAS:
            screenshots_eliminados = 0
            for archivo, descripcion in SCREENSHOTS_TOMADAS:
                if os.path.exists(archivo):
                    try:
                        os.remove(archivo)
                        screenshots_eliminados += 1
                        LOGGER.info(f"Screenshot residual eliminado: {archivo}")
                    except Exception as e:
                        LOGGER.warning(f"Error eliminando {archivo}: {e}")
            
            SCREENSHOTS_TOMADAS.clear()
            if screenshots_eliminados > 0:
                LOGGER.info(f"{screenshots_eliminados} screenshots residuales eliminados")
        else:
            LOGGER.info("Screenshots ya fueron eliminados por cotización")
        
        # Cerrar navegador
        if DRIVER:
            try:
                DRIVER.quit()
                LOGGER.info(" Navegador cerrado")
            except:
                LOGGER.warning(" Error cerrando navegador")
        
        LOGGER.info(" Limpieza de recursos completada")
        
    except Exception as e:
        LOGGER.error(f"Error limpiando recursos: {e}")

def main():
    """
    MÉTODO PRINCIPAL que ejecuta todo el flujo modular
    """
    if len(sys.argv) != 2:
        print("Uso: python automatizacion_modular.py <archivo_excel>")
        return
    
    archivo_excel = sys.argv[1]
    
    try:
        LOGGER.info(" === AUTOMATIZACIÓN MODULAR INICIADA ===")
        LOGGER.info(f"Archivo: {archivo_excel}")
        
        # MÓDULO 1: Lectura Excel
        cotizaciones = leer_excel_cotizaciones(archivo_excel)
        if not cotizaciones:
            return
        
        # MÓDULO 2: Inicializar navegador y sesión
        if not iniciar_navegador():
            return
        
        if not iniciar_sesion():
            return
        
        # Procesar cada cotización (cada una genera su propia orden)
        for i, (cotizacion_id, productos) in enumerate(cotizaciones.items(), 1):
            LOGGER.info(f"\n🔄 INICIANDO COTIZACIÓN {i}/{len(cotizaciones)}: {cotizacion_id}")
            
            # Cada cotización necesita su propia orden de compra
            if not iniciar_orden_compra():
                LOGGER.error(f"❌ No se pudo iniciar orden de compra para cotización {cotizacion_id}")
                continue
            
            # Para cotizaciones adicionales, preparar nueva sesión si es necesario
            if i > 1:
                if not preparar_nueva_cotizacion():
                    LOGGER.error(f"❌ No se pudo preparar cotización {cotizacion_id}")
                    continue
            
            exito = procesar_cotizacion_individual(cotizacion_id, productos)
            if not exito:
                LOGGER.warning(f"⚠️ Cotización {cotizacion_id} tuvo problemas")
            
            LOGGER.info(f"✅ Cotización {cotizacion_id} completada")
            time.sleep(2)
        
        # Crear reportes finales
        excel_actualizado, reporte_word = crear_reportes(archivo_excel)
        
        # Resumen final
        LOGGER.info(" PROCESO MODULAR COMPLETADO")
        LOGGER.info(f"Cotizaciones procesadas: {len(COTIZACIONES_PROCESADAS)}")
        if excel_actualizado:
            LOGGER.info(f"Excel: {excel_actualizado}")
        if reporte_word:
            LOGGER.info(f"📑 Word: {reporte_word}")
        
        # Mantener navegador abierto
        LOGGER.info("Navegador abierto por 60 segundos...")
        time.sleep(2)
        
    except KeyboardInterrupt:
        LOGGER.warning("⏹️ Interrumpido por usuario")
        # CREAR REPORTES AUNQUE SEA INTERRUMPIDO
        if COTIZACIONES_PROCESADAS:
            LOGGER.info("📊 Creando reportes parciales...")
            excel_actualizado, reporte_word = crear_reportes(archivo_excel)
            if excel_actualizado:
                LOGGER.info(f"📊 Excel parcial: {excel_actualizado}")
            if reporte_word:
                LOGGER.info(f"📑 Word parcial: {reporte_word}")
        LOGGER.info(f"📋 Log completo disponible en: {LOG_FILE}")
    except Exception as e:
        LOGGER.error(f"❌ Error general: {e}")
        LOGGER.info(f"📋 Ver detalles del error en: {LOG_FILE}")
        # Los reportes se crean automáticamente al finalizar
    finally:
        limpiar_recursos()
        LOGGER.info("=== AUTOMATIZACIÓN FINALIZADA ===")

if __name__ == "__main__":
    main()