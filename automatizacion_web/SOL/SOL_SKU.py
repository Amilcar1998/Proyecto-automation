from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import random
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

class SOLSKUClass:
    """Clase para manejar la búsqueda de SKU en el módulo SOL"""
    
    def __init__(self, driver):
        self.driver = driver
        self.link_mantenimiento_sku = None
        self.doc = None
        self.evidencias_pasos = []
        self.carpeta_actual = None
        self.capturas = {
            'mantenimiento_inicial': None,
            'sku_ingresado': None,
            'tabla_resultados': None,
            'opciones_desplegadas': None,
            'opcion_seleccionada': None,
            'donde_resultado': None
        }
        self.locacion_seleccionada = None
        
    def acceder_frame_trabajo(self):
        """Cambia al frame de trabajo donde está el contenido principal"""
        try:
            # Primero volver al contenido principal
            self.driver.switch_to.default_content()
            
            # Intentar acceder al frame 'trabajo' si existe
            try:
                self.driver.switch_to.frame("trabajo")
                print("✅ Cambiado al frame 'trabajo'")
                return True
            except:
                # Si no existe el frame 'trabajo', intentar con 'center' o quedarse en default
                try:
                    self.driver.switch_to.frame("center")
                    print("✅ Cambiado al frame 'center'")
                    return True
                except:
                    # Si no hay frames específicos, quedarse en el contexto actual
                    print("✅ Usando contexto por defecto (sin frame específico)")
                    return True
                    
        except Exception as e:
            print(f"⚠️ Advertencia al cambiar frame: {e}")
            # Intentar continuar sin cambiar de frame
            return True
    
    def buscar_boton_sku(self):
        """Busca el enlace 'Mantenimiento de SKU' en la página"""
        try:
            time.sleep(2)
            
            # Primero intentamos en el frame west (menú lateral)
            try:
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame("west")
                print("✅ Cambiado al frame 'west' (menú)")
            except:
                print("⚠️ No se pudo acceder al frame 'west', continuando en contexto actual")
            
            # Estrategia 1: Por ID exacto
            try:
                link = self.driver.find_element(By.ID, "1032")
                self.link_mantenimiento_sku = link
                print("✅ Enlace 'Mantenimiento de SKU' encontrado por ID")
                return True
            except NoSuchElementException:
                pass
            
            # Estrategia 2: Por href exacto
            try:
                link = self.driver.find_element(By.XPATH, "//a[@href='/Sol/busquedaSKU.htm?comando=iniciarSesion']")
                self.link_mantenimiento_sku = link
                print("✅ Enlace 'Mantenimiento de SKU' encontrado por href")
                return True
            except NoSuchElementException:
                pass
            
            # Estrategia 3: Por texto exacto
            try:
                link = self.driver.find_element(By.LINK_TEXT, "Mantenimiento de SKU")
                self.link_mantenimiento_sku = link
                print("✅ Enlace 'Mantenimiento de SKU' encontrado por texto")
                return True
            except NoSuchElementException:
                pass
            
            # Estrategia 4: Por XPath con texto
            try:
                link = self.driver.find_element(By.XPATH, "//a[contains(text(), 'Mantenimiento de SKU')]")
                self.link_mantenimiento_sku = link
                print("✅ Enlace 'Mantenimiento de SKU' encontrado por XPath")
                return True
            except NoSuchElementException:
                pass
            
            # Estrategia 5: Por clase y texto
            try:
                link = self.driver.find_element(By.XPATH, "//a[@class='siman-menu-contenido-link ajax' and contains(text(), 'Mantenimiento de SKU')]")
                self.link_mantenimiento_sku = link
                print("✅ Enlace 'Mantenimiento de SKU' encontrado por clase y texto")
                return True
            except NoSuchElementException:
                pass
            
            print("❌ No se pudo encontrar el enlace 'Mantenimiento de SKU'")
            return False
            
        except Exception as e:
            print(f"❌ Error al buscar el enlace: {e}")
            return False
    
    def hacer_click_sku(self):
        """Hace click en el enlace 'Mantenimiento de SKU'"""
        try:
            if self.link_mantenimiento_sku is None:
                print("⚠️ Enlace 'Mantenimiento de SKU' no encontrado. Buscando...")
                if not self.buscar_boton_sku():
                    return False
            
            # Intentar con JavaScript primero (más confiable)
            try:
                print("🔧 Intentando click con JavaScript...")
                self.driver.execute_script("arguments[0].click();", self.link_mantenimiento_sku)
                print("✅ Click en 'Mantenimiento de SKU' realizado exitosamente (JavaScript)")
                time.sleep(5)
                return True
            except Exception as e:
                print(f"⚠️ Click JavaScript falló: {e}")
            
            # Si JavaScript falla, intentar con WebDriverWait
            print("🔧 Intentando click estándar con espera...")
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(self.link_mantenimiento_sku)
            )
            
            self.link_mantenimiento_sku.click()
            print("✅ Click en 'Mantenimiento de SKU' realizado exitosamente")
            time.sleep(5)
            return True
            
        except TimeoutException:
            print("❌ Tiempo de espera agotado. El enlace no está clickeable.")
            return False
        except NoSuchElementException:
            print("❌ No se encontró el enlace 'Mantenimiento de SKU'.")
            return False
        except Exception as e:
            print(f"❌ Error inesperado al hacer click: {e}")
            return False
    
    def buscar_sku_por_numero(self, numero_sku):
        """
        Busca un SKU específico ingresando el número en la caja de texto y presionando Enter
        
        Args:
            numero_sku (str): El número de SKU a buscar (ej: "123456")
            
        Returns:
            bool: True si la búsqueda fue exitosa, False en caso contrario
        """
        try:
            print(f"\n🔍 Buscando SKU: {numero_sku}")
            
            # Crear carpeta única "Evidencias" para todos los SKUs
            # Solo configurar si no está ya establecida (permite usar carpetas personalizadas)
            if not self.carpeta_actual:
                self.carpeta_actual = "Evidencias"
            
            if not os.path.exists(self.carpeta_actual):
                os.makedirs(self.carpeta_actual)
                print(f"📁 Carpeta creada: {self.carpeta_actual}")
            else:
                print(f"📁 Usando carpeta existente: {self.carpeta_actual}")
            
            # CAPTURA 1: Pantalla de Mantenimiento inicial
            time.sleep(2)
            ruta_captura = os.path.join(self.carpeta_actual, f"1_mantenimiento_inicial_{numero_sku}.png")
            self.capturar_pantalla(ruta_captura)
            self.capturas['mantenimiento_inicial'] = ruta_captura
            self.agregar_evidencia_paso(f"Ingreso a SOL - Pantalla de Mantenimiento de Item")
            
            # Cambiar al frame correcto - después de hacer click en Mantenimiento, el contenido está en frame "center"
            try:
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame("center")
                print("✅ Cambiado al frame 'center'")
            except:
                try:
                    self.driver.switch_to.default_content()
                    print("✅ Usando default_content")
                except:
                    print("⚠️ Usando contexto actual (sin cambio de frame)")
            
            time.sleep(2)
            
            # Buscar la caja de texto del SKU por name="skuNumber"
            input_sku = None
            
            # Estrategia 1: Por name "skuNumber"
            try:
                input_sku = self.driver.find_element(By.NAME, "skuNumber")
                print("✅ Caja de texto 'skuNumber' encontrada")
            except NoSuchElementException:
                pass
            
            # Estrategia 2: Por id "skuNumber"
            if not input_sku:
                try:
                    input_sku = self.driver.find_element(By.ID, "skuNumber")
                    print("✅ Caja de texto encontrada por id='skuNumber'")
                except NoSuchElementException:
                    pass
            
            # Estrategia 3: Por xpath con name o id skuNumber
            if not input_sku:
                try:
                    input_sku = self.driver.find_element(By.XPATH, "//input[@name='skuNumber' or @id='skuNumber']")
                    print("✅ Caja de texto encontrada por XPath")
                except NoSuchElementException:
                    pass
            
            if not input_sku:
                print("❌ No se pudo encontrar la caja de texto 'skuNumber'")
                self.capturar_pantalla(f"sku_input_no_encontrado_{numero_sku}.png")
                self.obtener_html_actual(f"sku_input_no_encontrado_{numero_sku}.html")
                return False
            
            # Limpiar la caja de texto de forma más agresiva
            try:
                # Método 1: Hacer click y seleccionar todo
                input_sku.click()
                time.sleep(0.5)
                input_sku.send_keys(Keys.CONTROL + "a")
                time.sleep(0.5)
                input_sku.send_keys(Keys.DELETE)
                time.sleep(0.5)
                
                # Método 2: Clear tradicional
                input_sku.clear()
                time.sleep(0.5)
                
                # Método 3: Backspace múltiple por si acaso
                for _ in range(10):
                    input_sku.send_keys(Keys.BACKSPACE)
                time.sleep(0.5)
                
                print("✅ Campo limpiado correctamente")
            except Exception as e:
                print(f"⚠️ Advertencia al limpiar campo: {e}")
            
            # Escribir el número
            print(f"📝 Ingresando número: {numero_sku}")
            input_sku.send_keys(numero_sku)
            
            # CAPTURA 2: SKU ingresado en el campo
            time.sleep(1)
            ruta_captura = os.path.join(self.carpeta_actual, f"2_sku_ingresado_{numero_sku}.png")
            self.capturar_pantalla(ruta_captura)
            self.capturas['sku_ingresado'] = ruta_captura
            self.agregar_evidencia_paso(f"Item {numero_sku} ingresado en el campo de búsqueda")
            
            # Presionar Enter
            input_sku.send_keys(Keys.ENTER)
            print("✅ Enter presionado, esperando resultado...")
            
            time.sleep(5)  # Esperar más tiempo para que cargue la tabla AJAX
            
            # Buscar la tabla de datos
            print("\n📊 Buscando tabla de resultados...")
            tabla = None
            
            # Estrategia 1: Esperar por el div contenedor tabla-datos
            try:
                contenedor = WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.tabla-datos"))
                )
                print("✅ Contenedor 'tabla-datos' encontrado")
                # Buscar la tabla dentro del contenedor
                tabla = contenedor.find_element(By.TAG_NAME, "table")
                print("✅ Tabla dentro del contenedor encontrada")
            except TimeoutException:
                pass
            
            # Estrategia 2: Buscar directamente la tabla jqgrid
            if not tabla:
                try:
                    tabla = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.ID, "list"))
                    )
                    print("✅ Tabla encontrada por ID 'list'")
                except TimeoutException:
                    pass
            
            # Estrategia 3: Buscar por clase ui-jqgrid-btable
            if not tabla:
                try:
                    tabla = self.driver.find_element(By.CSS_SELECTOR, "table.ui-jqgrid-btable")
                    print("✅ Tabla encontrada por clase 'ui-jqgrid-btable'")
                except NoSuchElementException:
                    pass
            
            # Estrategia 4: Buscar tabla con role grid
            if not tabla:
                try:
                    tabla = self.driver.find_element(By.CSS_SELECTOR, "table[role='grid']")
                    print("✅ Tabla encontrada por role='grid'")
                except NoSuchElementException:
                    pass
            
            if not tabla:
                print("❌ No se pudo encontrar la tabla de datos")
                self.capturar_pantalla(f"tabla_no_encontrada_{numero_sku}.png")
                self.obtener_html_actual(f"tabla_no_encontrada_{numero_sku}.html")
                return False
            
            # CAPTURA 3: Tabla de resultados cargada
            time.sleep(2)
            ruta_captura = os.path.join(self.carpeta_actual, f"3_tabla_resultados_{numero_sku}.png")
            self.capturar_pantalla(ruta_captura)
            self.capturas['tabla_resultados'] = ruta_captura
            self.agregar_evidencia_paso(f"Tabla de resultados cargada para SKU {numero_sku}")
            
            time.sleep(3)
            
            # Buscar la celda con "Locación Id" o el enlace de ubicaciones
            print("\n📍 Buscando celda 'Location Id'...")
            celda_location = None
            
            # Estrategia 1: Buscar el enlace con class="tip" y title="Locaciones Específicas"
            try:
                celda_location = self.driver.find_element(By.CSS_SELECTOR, "a.tip[title*='Locaciones']")
                print("✅ Enlace de 'Locaciones Específicas' encontrado")
            except NoSuchElementException:
                pass
            
            # Estrategia 2: Buscar cualquier enlace con class="tip"
            if not celda_location:
                try:
                    celda_location = self.driver.find_element(By.CSS_SELECTOR, "a.tip")
                    print("✅ Enlace con class='tip' encontrado")
                except NoSuchElementException:
                    pass
            
            # Estrategia 3: Buscar por el texto del código de ubicación (ej: C00P00F000T000)
            if not celda_location:
                try:
                    celda_location = self.driver.find_element(By.XPATH, "//a[contains(text(), 'C0') and contains(text(), 'P0')]")
                    print("✅ Celda con código de ubicación encontrada")
                except NoSuchElementException:
                    pass
            
            # Estrategia 4: Buscar en la tabla por el primer enlace de datos
            if not celda_location:
                try:
                    celda_location = tabla.find_element(By.XPATH, ".//tbody//tr[1]//td[5]//a")
                    print("✅ Primer enlace en columna de ubicación encontrado")
                except NoSuchElementException:
                    print("❌ No se pudo encontrar la celda de Location Id")
                    self.capturar_pantalla(f"celda_location_no_encontrada_{numero_sku}.png")
                    self.obtener_html_actual(f"debug_tabla_{numero_sku}.html")
                    return False
            
            # Mover el ratón sobre la celda para que aparezcan las opciones
            print("\n🖱️ Moviendo ratón sobre 'Location Id'...")
            
            actions = ActionChains(self.driver)
            actions.move_to_element(celda_location).perform()
            
            time.sleep(3)  # Esperar a que aparezca el tooltip
            
            # CAPTURA 4: Opciones desplegadas
            ruta_captura = os.path.join(self.carpeta_actual, f"4_opciones_location_desplegadas_{numero_sku}.png")
            self.capturar_pantalla(ruta_captura)
            self.capturas['opciones_desplegadas'] = ruta_captura
            self.agregar_evidencia_paso(f"Opciones de Location Id desplegadas para SKU {numero_sku}")
            
            # Buscar las opciones en el tooltip (dentro del cluetip)
            print("\n🔘 Buscando opciones en el tooltip...")
            
            try:
                # Buscar el contenedor del tooltip
                tooltip_container = None
                try:
                    tooltip_container = self.driver.find_element(By.ID, "cluetip-inner")
                    print("✅ Tooltip 'cluetip-inner' encontrado")
                except:
                    try:
                        tooltip_container = self.driver.find_element(By.ID, "cluetip")
                        print("✅ Tooltip 'cluetip' encontrado")
                    except:
                        print("⚠️ No se encontró el tooltip, buscando radio buttons directamente")
                
                # Buscar todos los radio buttons dentro del tooltip o en el documento
                if tooltip_container:
                    opciones = tooltip_container.find_elements(By.XPATH, ".//input[@type='radio']")
                    print(f"✅ Se encontraron {len(opciones)} radio buttons en el tooltip")
                else:
                    opciones = self.driver.find_elements(By.XPATH, "//input[@type='radio'][@name='_radio']")
                    print(f"✅ Se encontraron {len(opciones)} radio buttons en el documento")
                
                if len(opciones) < 2:
                    print(f"⚠️ Solo se encontraron {len(opciones)} opciones")
                    self.capturar_pantalla(f"opciones_insuficientes_{numero_sku}.png")
                    return False
                
                # Mostrar todas las opciones disponibles
                print("\n📋 Opciones de ubicación disponibles:")
                for i, opcion in enumerate(opciones, 1):
                    valor = opcion.get_attribute('value')
                    print(f"   {i}. {valor}")
                
                # Seleccionar aleatoriamente entre las opciones 2, 3 y 4
                opciones_validas = [2, 3, 4]
                # Filtrar solo las opciones que existen
                opciones_disponibles = [opt for opt in opciones_validas if opt <= len(opciones)]
                
                if not opciones_disponibles:
                    print("⚠️ No hay suficientes opciones disponibles (se necesitan al menos 2)")
                    self.capturar_pantalla(f"opciones_insuficientes_{numero_sku}.png")
                    return False
                
                # Seleccionar aleatoriamente
                indice_seleccionado = random.choice(opciones_disponibles)
                
                print(f"\n🎲 Selección aleatoria: Opción #{indice_seleccionado}")
                opcion_elegida = opciones[indice_seleccionado - 1]  # Ajustar a índice 0
                valor_elegido = opcion_elegida.get_attribute('value')
                print(f"   Seleccionando: {valor_elegido}")
                
                # Intentar click con JavaScript para mayor confiabilidad
                self.driver.execute_script("arguments[0].click();", opcion_elegida)
                print(f"✅ Click en opción #{indice_seleccionado} realizado exitosamente: {valor_elegido}")
                
                # Guardar la locación seleccionada
                self.locacion_seleccionada = valor_elegido
                
                time.sleep(2)
                
                # CAPTURA 5: Click en segunda pestaña Donde (sub-tab)
                ruta_captura = os.path.join(self.carpeta_actual, f"5_segunda_pestana_donde_{numero_sku}.png")
                self.capturar_pantalla(ruta_captura)
                self.capturas['opcion_seleccionada'] = ruta_captura
                self.agregar_evidencia_paso(f"Click en sub-tab Donde (segundo click)")
                
                # Ahora acceder a la pestaña "Donde"
                print("\n📍 Accediendo a la pestaña 'Donde'...")
                if not self.acceder_pestana_donde():
                    print("❌ No se pudo acceder a la pestaña 'Donde'")
                    self.capturar_pantalla(f"error_donde_{numero_sku}.png")
                    return False
                
                # CAPTURA 6: Resultado en Donde - Verificar si se muestra la compañía
                time.sleep(2)
                ruta_captura = os.path.join(self.carpeta_actual, f"6_validacion_compania_{numero_sku}.png")
                self.capturar_pantalla(ruta_captura)
                self.capturas['donde_resultado'] = ruta_captura
                
                # Verificar si la prueba es válida (si se muestran datos)
                # USAR LA MISMA LÓGICA QUE EN acceder_pestana_donde() para evitar inconsistencias
                prueba_valida = False
                valores_donde = []
                try:
                    # Buscar con múltiples selectores (igual que en acceder_pestana_donde)
                    gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#tabs-9 table.ui-jqgrid-btable td[role='gridcell']")
                    if len(gridcells) == 0:
                        gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#tabs-9 td[role='gridcell']")
                    if len(gridcells) == 0:
                        gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#tabs-9 tbody tr td[role='gridcell']")
                    if len(gridcells) == 0:
                        gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#list-selected td[role='gridcell']")
                    
                    # Si encontramos celdas, verificar que tengan texto
                    if len(gridcells) > 0:
                        for cell in gridcells:
                            texto = cell.text.strip()
                            if texto:
                                valores_donde.append(texto)
                        
                        if valores_donde:
                            info_donde = ", ".join(valores_donde)
                            self.agregar_evidencia_paso(f"✓ PRUEBA VÁLIDA - Compañía mostrada: {info_donde}")
                            print(f"✅ PRUEBA VÁLIDA - Se muestra la compañía en 'Donde': {info_donde}")
                            prueba_valida = True
                        else:
                            self.agregar_evidencia_paso(f"✗ PRUEBA NO VÁLIDA - Compañía no mostrada")
                            print(f"⚠️ PRUEBA NO VÁLIDA - No se muestra la compañía en 'Donde'")
                            prueba_valida = False
                    else:
                        self.agregar_evidencia_paso(f"✗ PRUEBA NO VÁLIDA - Compañía no mostrada")
                        print(f"⚠️ PRUEBA NO VÁLIDA - No se muestra la compañía en 'Donde'")
                        prueba_valida = False
                except Exception as e:
                    print(f"⚠️ Error al verificar datos en Donde: {e}")
                    prueba_valida = False
                
                # Guardar resultado de validación para uso externo
                self.prueba_valida = prueba_valida
                
                # Generar documento Word SOLO si la prueba es válida
                if prueba_valida:
                    print("\n📝 Generando documento de evidencias del proceso...")
                    self.generar_documento_evidencias(numero_sku)
                else:
                    print("\n⚠️ Documento NO generado - Prueba no válida (sin datos en tabla)")
                
                # NO generar documento adicional para el "Donde" - todo en uno solo
                
                # Regresar a Mantenimiento de SKU
                print("\n↩️ Regresando a Mantenimiento de SKU...")
                if self.regresar_mantenimiento_sku():
                    print("✅ Regreso exitoso a Mantenimiento de SKU")
                    # Limpiar evidencias para el siguiente SKU
                    self.evidencias_pasos = []
                    self.locacion_seleccionada = None
                else:
                    print("⚠️ No se pudo regresar automáticamente")
                
                print(f"✅ Procesamiento de SKU {numero_sku} completado")
                
                return True
                
            except Exception as e:
                print(f"❌ Error al buscar opciones: {e}")
                self.capturar_pantalla(f"error_opciones_{numero_sku}.png")
                self.obtener_html_actual(f"error_opciones_{numero_sku}.html")
                return False
            
        except Exception as e:
            print(f"❌ Error al buscar SKU {numero_sku}: {e}")
            self.capturar_pantalla(f"error_busqueda_sku_{numero_sku}.png")
            return False
    
    def procesar_lista_skus(self, lista_skus, intervalo_segundos=2):
        """
        Procesa una lista de SKUs uno por uno
        
        Args:
            lista_skus (list): Lista de números de SKU a buscar
            intervalo_segundos (int): Tiempo de espera entre cada búsqueda
            
        Returns:
            dict: Resumen con SKUs exitosos y fallidos
        """
        resultados = {
            "exitosos": [],
            "fallidos": [],
            "total": len(lista_skus)
        }
        
        print(f"\n{'='*70}")
        print(f"📋 PROCESANDO {len(lista_skus)} SKUs")
        print(f"{'='*70}\n")
        
        for i, sku in enumerate(lista_skus, 1):
            print(f"\n--- SKU {i}/{len(lista_skus)} ---")
            
            if self.buscar_sku_por_numero(sku):
                resultados["exitosos"].append(sku)
            else:
                resultados["fallidos"].append(sku)
            
            # Esperar entre búsquedas (excepto en la última)
            if i < len(lista_skus):
                print(f"⏳ Esperando {intervalo_segundos} segundos antes del siguiente SKU...\n")
                time.sleep(intervalo_segundos)
        
        # Mostrar resumen
        print(f"\n{'='*70}")
        print(f"📊 RESUMEN DE PROCESAMIENTO")
        print(f"{'='*70}")
        print(f"✅ Exitosos: {len(resultados['exitosos'])}/{resultados['total']}")
        print(f"❌ Fallidos: {len(resultados['fallidos'])}/{resultados['total']}")
        
        if resultados["exitosos"]:
            print(f"\n✅ SKUs procesados correctamente:")
            for sku in resultados["exitosos"]:
                print(f"   - {sku}")
        
        if resultados["fallidos"]:
            print(f"\n❌ SKUs con error:")
            for sku in resultados["fallidos"]:
                print(f"   - {sku}")
        
        print(f"{'='*70}\n")
        
        return resultados
    
    def capturar_pantalla(self, nombre_archivo="sol_sku_captura.png"):
        """Captura una pantalla de la página actual"""
        try:
            self.driver.save_screenshot(nombre_archivo)
            print(f"✅ Captura de pantalla guardada: {nombre_archivo}")
            return True
        except Exception as e:
            print(f"❌ Error al capturar pantalla: {e}")
            return False
    
    def obtener_html_actual(self, nombre_archivo="sol_sku_page.html"):
        """Guarda el HTML de la página actual para análisis"""
        try:
            html = self.driver.page_source
            with open(nombre_archivo, "w", encoding="utf-8") as f:
                f.write(html)
            print(f"✅ HTML guardado en: {nombre_archivo}")
            return True
        except Exception as e:
            print(f"❌ Error al guardar HTML: {e}")
            return False
    
    def acceder_pestana_donde(self):
        """
        Hace click en la pestaña 'Donde' principal y luego en el sub-tab 'Donde',
        verificando que exista al menos un gridcell
        
        Returns:
            bool: True si se accedió correctamente y existe al menos un gridcell, False en caso contrario
        """
        try:
            # Esperar un momento para que la página esté lista
            time.sleep(2)
            
            # PASO 1: Buscar y hacer click en la pestaña principal "Donde"
            print("🔍 Buscando pestaña principal 'Donde'...")
            
            link_donde = None
            
            # Estrategia 1: Por href exacto
            try:
                link_donde = self.driver.find_element(By.XPATH, "//a[@href='#where']")
                print("✅ Pestaña principal 'Donde' encontrada por href")
            except NoSuchElementException:
                pass
            
            # Estrategia 2: Por texto del span
            if not link_donde:
                try:
                    link_donde = self.driver.find_element(By.XPATH, "//a[.//span[text()='Donde']]")
                    print("✅ Pestaña principal 'Donde' encontrada por texto del span")
                except NoSuchElementException:
                    pass
            
            # Estrategia 3: Por texto directo
            if not link_donde:
                try:
                    link_donde = self.driver.find_element(By.LINK_TEXT, "Donde")
                    print("✅ Pestaña principal 'Donde' encontrada por texto directo")
                except NoSuchElementException:
                    pass
            
            # Estrategia 4: Por XPath más amplio
            if not link_donde:
                try:
                    link_donde = self.driver.find_element(By.XPATH, "//a[contains(text(), 'Donde') or contains(@href, 'where')]")
                    print("✅ Pestaña principal 'Donde' encontrada por XPath amplio")
                except NoSuchElementException:
                    print("❌ No se pudo encontrar la pestaña principal 'Donde'")
                    return False
            
            # Hacer click en la pestaña principal
            print("👆 Haciendo click en pestaña principal 'Donde'...")
            
            # Intentar con JavaScript primero
            try:
                self.driver.execute_script("arguments[0].click();", link_donde)
                print("✅ Click en pestaña principal 'Donde' realizado (JavaScript)")
            except:
                # Si falla, intentar click normal
                link_donde.click()
                print("✅ Click en pestaña principal 'Donde' realizado (normal)")
            
            time.sleep(3)  # Esperar a que cargue el contenido de la pestaña
            
            # PASO 2: Buscar y hacer click en el sub-tab "Donde"
            print("\n🔍 Buscando sub-tab 'Donde'...")
            
            link_subtab_donde = None
            
            # Estrategia 1: Por href exacto del subtab (tabs-9 según el HTML)
            try:
                link_subtab_donde = self.driver.find_element(By.XPATH, "//a[@href='#tabs-9']")
                print("✅ Sub-tab 'Donde' encontrado por href #tabs-9")
            except NoSuchElementException:
                pass
            
            # Estrategia 2: Por texto del span dentro del sub-tab-where
            if not link_subtab_donde:
                try:
                    link_subtab_donde = self.driver.find_element(By.XPATH, "//div[@id='sub-tab-where']//a[.//span[text()='Donde']]")
                    print("✅ Sub-tab 'Donde' encontrado dentro de sub-tab-where")
                except NoSuchElementException:
                    pass
            
            # Estrategia 3: Buscar dentro del contenedor where cualquier enlace con texto Donde
            if not link_subtab_donde:
                try:
                    link_subtab_donde = self.driver.find_element(By.XPATH, "//div[@id='where']//a[contains(text(), 'Donde')]")
                    print("✅ Sub-tab 'Donde' encontrado dentro del div #where")
                except NoSuchElementException:
                    pass
            
            # Estrategia 4: Buscar el tercer elemento li de los sub-tabs
            if not link_subtab_donde:
                try:
                    link_subtab_donde = self.driver.find_element(By.XPATH, "//div[@id='sub-tab-where']//ul//li[3]//a")
                    print("✅ Sub-tab 'Donde' encontrado como tercer elemento")
                except NoSuchElementException:
                    print("❌ No se pudo encontrar el sub-tab 'Donde'")
                    return False
            
            # Hacer click en el sub-tab
            print("👆 Haciendo click en sub-tab 'Donde'...")
            
            # Intentar con JavaScript primero
            try:
                self.driver.execute_script("arguments[0].click();", link_subtab_donde)
                print("✅ Click en sub-tab 'Donde' realizado (JavaScript)")
            except:
                # Si falla, intentar click normal
                link_subtab_donde.click()
                print("✅ Click en sub-tab 'Donde' realizado (normal)")
            
            time.sleep(3)  # Esperar a que cargue el contenido del sub-tab
            
            # PASO 3: Verificar que exista al menos un gridcell
            print("\n🔍 Verificando existencia de datos en la tabla 'Donde'...")
            
            try:
                # Esperar a que la tabla cargue
                time.sleep(2)
                
                # Buscar la tabla por múltiples métodos
                # Método 1: Buscar por clase de tabla jqGrid
                gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#tabs-9 table.ui-jqgrid-btable td[role='gridcell']")
                
                # Método 2: Si no encuentra, buscar cualquier gridcell dentro de tabs-9
                if len(gridcells) == 0:
                    gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#tabs-9 td[role='gridcell']")
                
                # Método 3: Buscar en toda la tabla dentro del tab
                if len(gridcells) == 0:
                    gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#tabs-9 tbody tr td[role='gridcell']")
                
                # Método 4: Buscar tabla list-selected específicamente
                if len(gridcells) == 0:
                    gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#list-selected td[role='gridcell']")
                
                if len(gridcells) == 0:
                    print("❌ No se encontraron datos en la tabla 'Donde' (tabla vacía)")
                    return False
                
                print(f"✅ Se encontraron {len(gridcells)} celdas con datos en la tabla")
                
                # Mostrar información de las celdas encontradas
                if gridcells:
                    datos_tabla = []
                    for i, cell in enumerate(gridcells, 1):
                        texto = cell.text.strip()
                        if texto:  # Solo mostrar celdas con texto
                            datos_tabla.append(texto)
                            print(f"📊 Celda {i}: '{texto}'")
                    
                    # Si encontramos datos, es válido
                    if datos_tabla:
                        print(f"✅ Tabla 'Donde' contiene datos válidos: {', '.join(datos_tabla)}")
                        return True
                    else:
                        print("❌ No se encontraron datos con texto en las celdas")
                        return False
                
                return True
                
            except Exception as e:
                print(f"❌ Error al verificar datos en tabla: {e}")
                return False
            
        except Exception as e:
            print(f"❌ Error al acceder a pestaña 'Donde': {e}")
            return False
    
    def agregar_evidencia_paso(self, descripcion):
        """Agrega un paso de evidencia a la lista"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.evidencias_pasos.append({
            'tiempo': timestamp,
            'descripcion': descripcion
        })
        print(f"📌 Evidencia registrada: {descripcion}")
    
    def generar_documento_evidencias(self, numero_sku):
        """Genera documento Word con todas las evidencias recopiladas e inserta las capturas"""
        try:
            doc = Document()
            
            # EVIDENCIAS CON CAPTURAS (sin portada ni resumen final)
            doc.add_heading("Evidencias del Proceso", 1)
            
            # Paso 1: Pantalla inicial de Mantenimiento
            doc.add_heading("1. Pantalla de Mantenimiento de SKU", 2)
            descripcion = doc.add_paragraph()
            descripcion.add_run("Acceso a la pantalla de Mantenimiento de SKU donde se buscará el ítem ").font.size = Pt(11)
            sku_run = descripcion.add_run(f"{numero_sku}")
            sku_run.font.size = Pt(11)
            sku_run.font.bold = True
            sku_run.font.color.rgb = RGBColor(0, 102, 204)
            descripcion.add_run(":").font.size = Pt(11)
            
            if self.capturas['mantenimiento_inicial'] and os.path.exists(self.capturas['mantenimiento_inicial']):
                doc.add_picture(self.capturas['mantenimiento_inicial'], width=Inches(6.0))
            doc.add_paragraph()
            
            # Paso 2: SKU ingresado
            doc.add_heading("2. Ingreso del SKU", 2)
            descripcion = doc.add_paragraph()
            descripcion.add_run("Se ingresa el número de SKU '").font.size = Pt(11)
            sku_run = descripcion.add_run(f"{numero_sku}")
            sku_run.font.size = Pt(11)
            sku_run.font.bold = True
            sku_run.font.color.rgb = RGBColor(0, 102, 204)
            descripcion.add_run("' en el campo de búsqueda ").font.size = Pt(11)
            campo_run = descripcion.add_run("skuNumber")
            campo_run.font.size = Pt(11)
            campo_run.font.italic = True
            campo_run.font.color.rgb = RGBColor(128, 128, 128)
            descripcion.add_run(":").font.size = Pt(11)
            
            if self.capturas['sku_ingresado'] and os.path.exists(self.capturas['sku_ingresado']):
                doc.add_picture(self.capturas['sku_ingresado'], width=Inches(6.0))
            doc.add_paragraph()
            
            # Paso 3: Tabla de resultados
            doc.add_heading("3. Tabla de Resultados", 2)
            descripcion = doc.add_paragraph()
            descripcion.add_run("La tabla muestra los resultados encontrados para el SKU ").font.size = Pt(11)
            sku_run = descripcion.add_run(f"{numero_sku}")
            sku_run.font.size = Pt(11)
            sku_run.font.bold = True
            sku_run.font.color.rgb = RGBColor(0, 102, 204)
            descripcion.add_run(":").font.size = Pt(11)
            
            if self.capturas['tabla_resultados'] and os.path.exists(self.capturas['tabla_resultados']):
                doc.add_picture(self.capturas['tabla_resultados'], width=Inches(6.0))
            doc.add_paragraph()
            
            # Paso 4: Opciones desplegadas
            doc.add_heading("4. Opciones de Locación Desplegadas", 2)
            descripcion = doc.add_paragraph()
            descripcion.add_run("Al pasar el cursor sobre ").font.size = Pt(11)
            locacion_run = descripcion.add_run("'Locaciones Específicas'")
            locacion_run.font.size = Pt(11)
            locacion_run.font.bold = True
            locacion_run.font.color.rgb = RGBColor(204, 102, 0)
            descripcion.add_run(", se despliegan las opciones disponibles:").font.size = Pt(11)
            
            if self.capturas['opciones_desplegadas'] and os.path.exists(self.capturas['opciones_desplegadas']):
                doc.add_picture(self.capturas['opciones_desplegadas'], width=Inches(6.0))
            doc.add_paragraph()
            
            # Paso 5: Opción seleccionada
            doc.add_heading("5. Selección de Locación", 2)
            descripcion = doc.add_paragraph()
            if self.locacion_seleccionada:
                descripcion.add_run("Se selecciona la locación: ").font.size = Pt(11)
                locacion_run = descripcion.add_run(f"{self.locacion_seleccionada}")
                locacion_run.font.size = Pt(11)
                locacion_run.font.bold = True
                locacion_run.font.color.rgb = RGBColor(0, 153, 0)
            else:
                descripcion.add_run("Se selecciona una locación aleatoriamente entre las opciones ").font.size = Pt(11)
                opciones_run = descripcion.add_run("2, 3 y 4")
                opciones_run.font.size = Pt(11)
                opciones_run.font.bold = True
                opciones_run.font.color.rgb = RGBColor(0, 153, 0)
                descripcion.add_run(":").font.size = Pt(11)
            
            if self.capturas['opcion_seleccionada'] and os.path.exists(self.capturas['opcion_seleccionada']):
                doc.add_picture(self.capturas['opcion_seleccionada'], width=Inches(6.0))
            doc.add_paragraph()
            
            # Paso 6: Resultado en Donde
            doc.add_heading("6. Resultado en Pestaña 'Donde'", 2)
            descripcion = doc.add_paragraph()
            descripcion.add_run("Se accede a la pestaña ").font.size = Pt(11)
            donde_run = descripcion.add_run("'Donde'")
            donde_run.font.size = Pt(11)
            donde_run.font.bold = True
            donde_run.font.color.rgb = RGBColor(153, 0, 153)
            descripcion.add_run(" y se verifica que se muestren datos en la tabla:").font.size = Pt(11)
            
            if self.capturas['donde_resultado'] and os.path.exists(self.capturas['donde_resultado']):
                doc.add_picture(self.capturas['donde_resultado'], width=Inches(6.0))
            
            # Verificar si hay datos en la tabla y mostrar resultado
            doc.add_paragraph()
            try:
                # Usar la misma lógica de validación que en buscar_sku_por_numero
                gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#tabs-9 table.ui-jqgrid-btable td[role='gridcell']")
                if len(gridcells) == 0:
                    gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#tabs-9 td[role='gridcell']")
                if len(gridcells) == 0:
                    gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#tabs-9 tbody tr td[role='gridcell']")
                if len(gridcells) == 0:
                    gridcells = self.driver.find_elements(By.CSS_SELECTOR, "#list-selected td[role='gridcell']")
                
                valores_donde = []
                if len(gridcells) > 0:
                    for cell in gridcells:
                        texto = cell.text.strip()
                        if texto:
                            valores_donde.append(texto)
                
                if valores_donde:
                    resultado_p = doc.add_paragraph()
                    resultado_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    resultado_run = resultado_p.add_run("✓ RESULTADO: PRUEBA Exitosa - Se muestran datos en la tabla")
                    resultado_run.font.bold = True
                    resultado_run.font.size = Pt(14)
                    resultado_run.font.color.rgb = RGBColor(0, 128, 0)
                else:
                    resultado_p = doc.add_paragraph()
                    resultado_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    resultado_run = resultado_p.add_run("✗ RESULTADO: PRUEBA NO VÁLIDA - No se muestran datos")
                    resultado_run.font.bold = True
                    resultado_run.font.size = Pt(14)
                    resultado_run.font.color.rgb = RGBColor(255, 0, 0)
            except:
                pass
            
            # Guardar en la carpeta Evidencias
            nombre = os.path.join(self.carpeta_actual, f"Evidencia_Proceso_SKU_{numero_sku}.docx")
            doc.save(nombre)
            print(f"✅ Documento de evidencias del proceso generado: {nombre}")
            
            # NO eliminar las imágenes PNG - las necesitamos para actualizar otros documentos
            print(f"📁 Imágenes PNG guardadas en: {self.carpeta_actual}")
            print(f"   - Las 6 capturas están disponibles para uso posterior")
            
            return nombre
            
        except Exception as e:
            print(f"❌ Error al generar documento de evidencias: {e}")
            return None
    
    def regresar_mantenimiento_sku(self):
        """Regresa a la pantalla de Mantenimiento de SKU para buscar el siguiente"""
        try:
            print("   🔄 Limpiando contexto actual...")
            time.sleep(3)
            
            # Paso 1: Volver al contenido principal
            try:
                self.driver.switch_to.default_content()
                print("   ✅ Contexto principal restaurado")
            except Exception as e:
                print(f"   ⚠️ Error al restaurar contexto: {e}")
            
            time.sleep(2)
            
            # Paso 2: Cambiar al frame del menú (west)
            try:
                self.driver.switch_to.frame("west")
                print("   ✅ Frame 'west' (menú) accesible")
            except Exception as e:
                print(f"   ⚠️ No se pudo acceder al frame 'west': {e}")
                # Intentar continuar sin el frame
            
            time.sleep(1)
            
            # Paso 3: Buscar el enlace de Mantenimiento de SKU
            print("   🔍 Buscando enlace 'Mantenimiento de SKU'...")
            if not self.buscar_boton_sku():
                print("   ❌ No se encontró el enlace")
                return False
            
            # Paso 4: Hacer click en Mantenimiento de SKU
            print("   👆 Haciendo click en 'Mantenimiento de SKU'...")
            try:
                self.driver.execute_script("arguments[0].click();", self.link_mantenimiento_sku)
                print("   ✅ Click realizado con JavaScript")
            except Exception as e:
                print(f"   ⚠️ Error al hacer click: {e}")
                return False
            
            # Paso 5: Esperar a que cargue la página
            print("   ⏳ Esperando carga completa de página...")
            time.sleep(7)  # Aumentado de 5 a 7 segundos
            
            # Paso 6: Cambiar al frame center donde está el formulario
            try:
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame("center")
                print("   ✅ Frame 'center' listo para nueva búsqueda")
            except Exception as e:
                print(f"   ⚠️ Usando default_content: {e}")
                self.driver.switch_to.default_content()
            
            # Paso 7: Esperar que el campo esté realmente disponible e interactuable
            print("   ⏳ Esperando que campo 'skuNumber' esté disponible...")
            time.sleep(3)  # Aumentado de 2 a 3 segundos
            
            try:
                # Esperar explícitamente que el campo sea visible y esté habilitado
                wait = WebDriverWait(self.driver, 10)
                campo = wait.until(EC.element_to_be_clickable((By.NAME, "skuNumber")))
                
                print("   ✅ Campo 'skuNumber' verificado y disponible")
                
                # Limpiar el campo por si tiene algo
                try:
                    # Scroll al elemento para asegurar visibilidad
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", campo)
                    time.sleep(0.5)
                    
                    campo.click()
                    time.sleep(0.5)
                    campo.send_keys(Keys.CONTROL + "a")
                    time.sleep(0.3)
                    campo.send_keys(Keys.DELETE)
                    time.sleep(0.3)
                    campo.clear()
                    print("   ✅ Campo limpiado preventivamente")
                except Exception as ex:
                    print(f"   ⚠️ Advertencia al limpiar campo preventivo: {ex}")
                
                return True
                
            except Exception as e:
                print(f"   ⚠️ Campo 'skuNumber' no está interactuable: {e}")
                return False
                
        except Exception as e:
            print(f"   ❌ Error general al regresar: {e}")
            return False


def ejecutar_busqueda_sku_sol(driver):
    """
    Función principal para ejecutar el acceso a Mantenimiento de SKU en SOL
    
    Args:
        driver: WebDriver de Selenium ya inicializado y con sesión activa
    
    Returns:
        bool: True si la operación fue exitosa, False en caso contrario
    """
    try:
        print("\n" + "="*60)
        print("INICIANDO ACCESO A MANTENIMIENTO DE SKU")
        print("="*60 + "\n")
        
        sol_sku = SOLSKUClass(driver)
        
        print("\n📍 Paso 1: Buscando enlace 'Mantenimiento de SKU'...")
        if not sol_sku.buscar_boton_sku():
            print("⚠️ Capturando evidencia para análisis...")
            sol_sku.capturar_pantalla("sol_sku_no_encontrado.png")
            sol_sku.obtener_html_actual("sol_sku_no_encontrado.html")
            return False
        
        # Paso 2: Hacer click en Mantenimiento de SKU
        print("\n📍 Paso 2: Haciendo click en 'Mantenimiento de SKU'...")
        if not sol_sku.hacer_click_sku():
            print("❌ No se pudo hacer click")
            sol_sku.capturar_pantalla("sol_sku_error_click.png")
            return False
        
        # Paso 3: Capturar resultado exitoso
        print("\n📍 Paso 3: Capturando resultado...")
        sol_sku.capturar_pantalla("sol_mantenimiento_sku_exitoso.png")
        
        print("\n" + "="*60)
        print("✅ ACCESO A MANTENIMIENTO DE SKU COMPLETADO")
        print("="*60 + "\n")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Error general en el proceso: {e}")
        return False
