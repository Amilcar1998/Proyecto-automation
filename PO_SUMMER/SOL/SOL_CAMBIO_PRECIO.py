import logging
import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Agregar el directorio raíz al path para importaciones
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from Page.Page_Login import LoginPage
from Page.page_menu import MenuPage


class SOLCambioPrecioClass:
    """Clase Page Object para manejar la funcionalidad de Cambio de Precio en el módulo SOL"""

    def __init__(self, driver):
        self.driver = driver
        # Locators del menú SOL
        self.header_cambio_precio = (By.XPATH, "//a[contains(text(), 'Cambio de precio')]")
        self.btn_crear_cambio_precio = (By.ID, "1072")           # Crear Cambio de Precio
        self.btn_mantenimiento_cambio_precio = (By.ID, "1071")   # Mantenimiento de Cambio de Precio
        self.btn_crear_cambio_precio_masivo = (By.ID, "1129")    # Crear Cambio de Precio Masivo

        # Locators del Formulario Crear Cambio de Precio (busquedaCambioPrecioForm / ajaxForm-whoPrice)
        self.txt_tipo_cambio_precio = (By.ID, "priceChangeCodesp")
        self.btn_lov_tipo_cambio = (By.ID, "bpriceChangeCodesp")
        self.txt_descripcion = (By.ID, "description")
        self.cmb_estatus = (By.ID, "status")
        self.txt_codigo_razon = (By.ID, "reasonCodesp")
        self.btn_lov_codigo_razon = (By.ID, "breasonCodesp")
        self.txt_numero_evento = (By.ID, "eventNumbersp")
        self.btn_lov_numero_evento = (By.ID, "beventNumbersp")
        self.rad_precio_regular = (By.ID, "calculationCurrent")
        self.rad_precio_credito = (By.ID, "calculationCredit")
        self.chk_preservar_precios_especiales = (By.ID, "preservePriceSpecial")
        self.chk_todas_las_zonas = (By.ID, "allZones")
        self.cmb_ordenamiento_reporte = (By.ID, "sortReport")
        self.txt_instrucciones_especiales = (By.ID, "specialInstruction")
        self.txt_fecha_inicio = (By.ID, "fechaInicio")
        self.txt_fecha_fin = (By.ID, "fechaFin")
        self.txt_hora_inicio_horas = (By.ID, "horaInicioHoras")
        self.txt_hora_inicio_minutos = (By.ID, "horaInicioMinutos")
        self.txt_hora_fin_horas = (By.ID, "horaFinHoras")
        self.txt_hora_fin_minutos = (By.ID, "horaFinMinutos")
        self.btn_siguiente_paso = (By.ID, "nextStep")

        # Locators de Importación de Excel
        self.btn_importacion_excel = (By.ID, "faceButton")            # Clic abre el diálogo de importación
        self.input_archivo_excel = (By.ID, "archivoid")               # input type="file"
        self.btn_submit_importar = (By.ID, "importarCambioDePrecio")  # Botón Importar submit

    def acceder_frame_menu_sol(self):
        """Accede al contexto donde se encuentra el menú accordion de SOL."""
        try:
            self.driver.switch_to.default_content()
            try:
                self.driver.switch_to.frame("west")
            except Exception:
                pass
            return True
        except Exception as e:
            logging.warning(f"Error cambiando al frame del menú SOL: {e}")
            return False

    def acceder_frame_contenido_sol(self):
        """Accede al frame/contenedor principal de contenido en SOL."""
        try:
            self.driver.switch_to.default_content()
            try:
                self.driver.switch_to.frame("center")
            except Exception:
                pass
            return True
        except Exception as e:
            logging.warning(f"Error cambiando al frame de contenido SOL: {e}")
            return False

    def desplegar_menu_cambio_precio(self):
        """Hace clic en la cabecera 'Cambio de precio' si aún no está desplegado."""
        try:
            self.acceder_frame_menu_sol()
            try:
                opcion = self.driver.find_element(*self.btn_crear_cambio_precio)
                if opcion.is_displayed():
                    logging.info("[OK] Menú 'Cambio de precio' ya está desplegado.")
                    return True
            except Exception:
                pass

            header = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(self.header_cambio_precio)
            )
            header.click()
            logging.info("Desplegado acordeón 'Cambio de precio'")
            time.sleep(1)
            return True
        except Exception as e:
            logging.error(f"Error al desplegar menú 'Cambio de precio': {e}")
            return False

    def acceder_crear_cambio_precio(self):
        """Accede a 'Crear Cambio de Precio' (ID 1072)."""
        try:
            self.desplegar_menu_cambio_precio()
            btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(self.btn_crear_cambio_precio)
            )
            btn.click()
            logging.info("Clic en 'Crear Cambio de Precio'(ID 1072) realizado exitosamente.")
            time.sleep(3)
            return True
        except Exception as e:
            logging.error(f"Error al acceder a 'Crear Cambio de Precio': {e}")
            return False

    def acceder_mantenimiento_cambio_precio(self):
        """Accede a 'Mantenimiento de Cambio de Precio' (ID 1071)."""
        try:
            self.desplegar_menu_cambio_precio()
            btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(self.btn_mantenimiento_cambio_precio)
            )
            btn.click()
            logging.info("Clic en 'Mantenimiento de Cambio de Precio'(ID 1071) realizado exitosamente.")
            time.sleep(3)
            return True
        except Exception as e:
            logging.error(f"Error al acceder a 'Mantenimiento de Cambio de Precio': {e}")
            return False

    def acceder_crear_cambio_precio_masivo(self):
        """Accede a 'Crear Cambio de Precio Masivo' (ID 1129)."""
        try:
            self.desplegar_menu_cambio_precio()
            btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(self.btn_crear_cambio_precio_masivo)
            )
            btn.click()
            logging.info("Clic en 'Crear Cambio de Precio Masivo'(ID 1129) realizado exitosamente.")
            time.sleep(3)
            return True
        except Exception as e:
            logging.error(f"Error al acceder a 'Crear Cambio de Precio Masivo': {e}")
            return False

    def _set_input_value(self, locator, valor, enviar_enter=True):
        """Método auxiliar seguro para asignar valor en inputs habilitando disabled de ser necesario."""
        self.acceder_frame_contenido_sol()
        elem = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(locator)
        )
        self.driver.execute_script("""
            var elem = arguments[0];
            var val = arguments[1];
            elem.removeAttribute('disabled');
            elem.removeAttribute('readonly');
            elem.value = val;
            elem.dispatchEvent(new Event('input', { bubbles: true }));
            elem.dispatchEvent(new Event('change', { bubbles: true }));
            elem.dispatchEvent(new Event('blur', { bubbles: true }));
            if (window.jQuery) {
                window.jQuery(elem).trigger('change').trigger('blur');
            }
        """, elem, valor)
        if enviar_enter:
            try:
                from selenium.webdriver.common.keys import Keys
                elem.send_keys(Keys.ENTER)
            except Exception:
                pass
        time.sleep(1)

    def _wait_and_click(self, locator_or_element, timeout=20, scroll=True, allow_js_fallback=True):
        """Click directo en un elemento con espera, scroll y fallback a JS si es necesario."""
        self.acceder_frame_contenido_sol()
        if isinstance(locator_or_element, tuple):
            elem = WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable(locator_or_element)
            )
        else:
            elem = locator_or_element

        if scroll:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center', inline:'nearest'});",
                elem
            )
            time.sleep(0.3)

        try:
            ActionChains(self.driver).move_to_element(elem).pause(0.2).click(elem).perform()
            return elem
        except Exception as e:
            try:
                elem.click()
                return elem
            except Exception:
                if allow_js_fallback:
                    self.driver.execute_script(
                        "var ev = document.createEvent('MouseEvents'); ev.initMouseEvent('click', true, true, window, 1, 0,0,0,0, false, false, false, false, 0, null); arguments[0].dispatchEvent(ev);",
                        elem
                    )
                    return elem
                raise e

    def _wait_for_tab(self, tab_selector, timeout=30):
        """Espera a que un panel de pestaña sea visible y activo."""
        self.acceder_frame_contenido_sol()
        return WebDriverWait(self.driver, timeout).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, tab_selector))
        )

    def llenar_formulario_crear_cambio_precio(
        self,
        tipo_cambio_precio="PMD",     # 'PMD' (Permanent Markdown), 'PMU' (Permanent Markup), 'TMD' (Temporary Markdown)
        descripcion="CAMBIO PRECIO AUTOMATIZADO",
        codigo_razon="03",           # '03' (NIVELACION), '04' (PROMOCION), '05' (LIQUIDACION)
        numero_evento="21",          # Por defecto 21
        tipo_calculo='C',            # 'C' = Precio Regular
        ordenamiento_reporte="DPT/SKU",
        instrucciones_especiales=None,
        clic_siguiente=True
    ):
        """
        Llena el formulario busquedaCambioPrecioForm siguiendo las reglas del negocio:
        - tipo_cambio_precio: 'PMD', 'PMU' o 'TMD' (envía ENTER para cargar datos)
        - codigo_razon: '03', '04' o '05' (envía ENTER)
        - numero_evento: por defecto '21' (envía ENTER)
        - fecha_inicio: hoy
        - fecha_fin: máximo 90 días a partir de hoy
        - todas_las_zonas: vacía salvo que sea permanente ('PMD'/'PMU' o empiece con 'P')
        - ordenamiento_reporte: 'DPT/SKU'
        """
        from datetime import datetime, timedelta
        try:
            logging.info("\n📋 Llenando formulario de Cambio de Precio...")
            self.acceder_frame_contenido_sol()

            # Cálculo de fechas: inicio hoy, fin hoy + 90 días
            fecha_hoy = datetime.now().strftime("%Y-%m-%d")
            fecha_90_dias = (datetime.now() + timedelta(days=90)).strftime("%Y-%m-%d")

            # Regla para 'todas_las_zonas': marcada si es tipo permanente (empieza con 'P', ej PMD/PMU)
            es_permanente = str(tipo_cambio_precio).upper().startswith("P")

            # 1. Tipo de Cambio de Precio (digita código y presiona ENTER)
            logging.info(f" -> Tipo de Cambio de Precio: {tipo_cambio_precio}")
            self._set_input_value(self.txt_tipo_cambio_precio, tipo_cambio_precio, enviar_enter=True)
            time.sleep(2)

            # 2. Descripción
            logging.info(f" -> Descripción: {descripcion}")
            self._set_input_value(self.txt_descripcion, descripcion, enviar_enter=False)

            # 3. Código de Razón (digita código y presiona ENTER)
            logging.info(f" -> Código de Razón: {codigo_razon}")
            self._set_input_value(self.txt_codigo_razon, codigo_razon, enviar_enter=True)

            # 4. Número de Evento (por defecto 21 y ENTER)
            logging.info(f" -> Número de Evento: {numero_evento}")
            self._set_input_value(self.txt_numero_evento, numero_evento, enviar_enter=True)

            # 5. Tipo de Cálculo (Radio: 'C' o 'P')
            rad_locator = self.rad_precio_credito if tipo_calculo == 'P' else self.rad_precio_regular
            elem = self.driver.find_element(*rad_locator)
            self.driver.execute_script("arguments[0].removeAttribute('disabled'); arguments[0].click();", elem)

            # 6. Todas las Zonas: marcado por defecto solo si es Permanente
            chk_zonas = self.driver.find_element(*self.chk_todas_las_zonas)
            if es_permanente:
                self.driver.execute_script("arguments[0].removeAttribute('disabled'); arguments[0].checked = true;", chk_zonas)
                logging.info(" -> Todas las Zonas: Marcado (Por ser cambio Permanente)")
            else:
                self.driver.execute_script("arguments[0].removeAttribute('disabled'); arguments[0].checked = false;", chk_zonas)
                logging.info(" -> Todas las Zonas: Desmarcado (Por ser cambio Temporal)")

            # 7. Ordenamiento de Reporte de Aviso ('DPT/SKU')
            if ordenamiento_reporte:
                from selenium.webdriver.support.ui import Select
                cmb = self.driver.find_element(*self.cmb_ordenamiento_reporte)
                self.driver.execute_script("arguments[0].removeAttribute('disabled');", cmb)
                Select(cmb).select_by_value(ordenamiento_reporte)
                logging.info(f" -> Ordenamiento de Reporte: {ordenamiento_reporte}")

            # 8. Instrucciones Especiales (si aplican)
            if instrucciones_especiales:
                self._set_input_value(self.txt_instrucciones_especiales, instrucciones_especiales, enviar_enter=False)

            # 9. Fechas (Inicio hoy, Fin 90 días)
            logging.info(f" -> Fecha Inicio: {fecha_hoy}")
            self._set_input_value(self.txt_fecha_inicio, fecha_hoy, enviar_enter=False)

            logging.info(f" -> Fecha Fin (máx 90 días): {fecha_90_dias}")
            self._set_input_value(self.txt_fecha_fin, fecha_90_dias, enviar_enter=False)

            # 10. Botón Siguiente Paso
            if clic_siguiente:
                btn_next = self._wait_and_click(self.btn_siguiente_paso, timeout=15, scroll=True)
                logging.info("Clic en 'Siguiente Paso'realizado exitosamente.")
                time.sleep(3)

            return True

        except Exception as e:
            logging.error(f"Error al llenar el formulario de Cambio de Precio: {e}")
            return False

    def importar_excel_cambio_precio(self, ruta_archivo_xls):
        """
        Realiza la importación del archivo Excel en la pestaña 'Dónde' (tabs-1):
        1. Hace clic en '#faceButton' para abrir el diálogo Facebox de importación
        2. Espera a que '#archivoid' sea visible (Facebox lo hace visible al abrirse)
        3. Envía la ruta del archivo al input type='file' (#archivoid)
        4. Hace clic en 'Importar' (#importarCambioDePrecio)
        5. Hace clic en 'Siguiente Paso' (#nextStepWhereCp)
        """
        try:
            logging.info(f"\n Importando archivo Excel: {ruta_archivo_xls}")
            self.acceder_frame_contenido_sol()
            time.sleep(2)

            # 1. Clic en '#faceButton' para abrir el diálogo Facebox
            logging.info("🖱️ Abriendo diálogo de importación (clic en 'Importación de Excel')...")
            btn_face = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(self.btn_importacion_excel)
            )
            btn_face.click()
            logging.info("Clic en 'Importación de Excel'(#faceButton) realizado.")

            # 2. Esperar a que #archivoid sea visible (el diálogo Facebox lo expone)
            logging.info("Esperando que el diálogo de importación sea visible...")
            file_input = WebDriverWait(self.driver, 15).until(
                EC.visibility_of_element_located(self.input_archivo_excel)
            )

            # 3. Enviar la ruta del archivo al input type="file"
            logging.info(f"📎 Adjuntando archivo: {ruta_archivo_xls}")
            file_input.send_keys(ruta_archivo_xls)
            logging.info(f"Archivo adjuntado correctamente.")
            time.sleep(1)

            # 4. Clic en 'Importar' (#importarCambioDePrecio)
            logging.info("📤 Enviando importación...")
            btn_importar = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(self.btn_submit_importar)
            )
            self.driver.execute_script("arguments[0].click();", btn_importar)
            logging.info("Clic en botón 'Importar'(#importarCambioDePrecio) realizado.")
            time.sleep(5)

            # 5. Clic en 'Siguiente Paso' (#nextStepWhereCp) de la pestaña 'Dónde'
            try:
                self._wait_and_click((By.ID, "nextStepWhereCp"), timeout=15, scroll=True)
                logging.info("Clic en 'Siguiente Paso'(#nextStepWhereCp) realizado.")

                # Esperar a que el AJAX de nextStepWhereCp termine antes de buscar el siguiente botón
                WebDriverWait(self.driver, 30).until(
                    lambda d: d.execute_script("return typeof jQuery !== 'undefined' ? jQuery.active == 0 : true")
                )
                logging.info("AJAX de 'Dónde'completado. Tabs-2 listo.")
                time.sleep(2)
            except Exception as e_next:
                logging.warning(f"No se pudo dar clic en 'nextStepWhereCp': {e_next}")

            # 6. Clic directo en 'Siguiente Paso' (#nextStep) de la pestaña 'Qué' (tabs-2)
            logging.info("📋 Avanzando pestaña 'Qué' → clic en 'Siguiente Paso'...")
            try:
                self._wait_for_tab("#tabs-2:not(.ui-tabs-hide)", timeout=20)
                self._wait_and_click((By.CSS_SELECTOR, "#tabs-2 #nextStep"), timeout=20, scroll=True)
                logging.info("Clic en 'Siguiente Paso'(#nextStep) de la pestaña 'Qué'realizado.")

                # Esperar a que el AJAX de nextStep termine antes de buscar Someter a Aprobacion
                WebDriverWait(self.driver, 30).until(
                    lambda d: d.execute_script("return typeof jQuery !== 'undefined' ? jQuery.active == 0 : true")
                )
                self._wait_for_tab("#tabs-3:not(.ui-tabs-hide)", timeout=30)
                logging.info("AJAX de 'Qué'completado. Tabs-3 listo.")
                time.sleep(2)
            except Exception as e_what:
                logging.warning(f"No se pudo dar clic en 'nextStep'(Qué): {e_what}")
                try:
                    self.driver.execute_script(
                        "if (typeof validarSiguientePaso === 'function') { validarSiguientePaso(); }"
                    )
                    logging.info("Invocado validarSiguientePaso() via JS como fallback.")
                    WebDriverWait(self.driver, 30).until(
                        lambda d: d.execute_script("return typeof jQuery !== 'undefined' ? jQuery.active == 0 : true")
                    )
                    self._wait_for_tab("#tabs-3:not(.ui-tabs-hide)", timeout=30)
                    logging.info("AJAX de 'Qué'completado tras fallback JS. Tabs-3 listo.")
                    time.sleep(2)
                except Exception as e_js:
                    logging.warning(f"No se pudo invocar validarSiguientePaso() via JS: {e_js}")

            # 7. Clic en 'Someter a Aprobacion' (tabs-3) para finalizar la creación
            try:
                logging.info("🏁 Sometiendo a aprobación...")
                btn_someter = WebDriverWait(self.driver, 30).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "#tabs-3 input[value='Someter a Aprobacion']")
                    )
                )
                ActionChains(self.driver).move_to_element(btn_someter).pause(0.2).click(btn_someter).perform()
                logging.info("Clic en 'Someter a Aprobacion'realizado. Cambio de Precio enviado a aprobación.")
                time.sleep(5)   # esperar que aparezca el mensaje de confirmación
            except Exception as e_someter:
                logging.warning(f"No se pudo dar clic en 'Someter a Aprobacion'directamente: {type(e_someter).__name__} {e_someter}")
                try:
                    btn_someter = WebDriverWait(self.driver, 30).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "#tabs-3 input[value='Someter a Aprobacion']")
                        )
                    )
                    self.driver.execute_script("arguments[0].click();", btn_someter)
                    logging.info("Clic en 'Someter a Aprobacion'via JS fallback realizado.")
                    time.sleep(5)
                except Exception as e_someter_js:
                    logging.warning(f"No se pudo dar clic en 'Someter a Aprobacion'via JS: {type(e_someter_js).__name__} {e_someter_js}")

            # 8. Clic en 'Cerrar' del mensaje de confirmación de éxito
            try:
                logging.info("🔒 Cerrando mensaje de confirmación...")
                cerrar_loc = (By.XPATH, "//a[normalize-space(text())='Cerrar' and not(contains(@style,'display:none'))]")
                btn_cerrar = WebDriverWait(self.driver, 20).until(
                    EC.element_to_be_clickable(cerrar_loc)
                )
                self._wait_and_click(btn_cerrar, timeout=15, scroll=True, allow_js_fallback=True)
                logging.info("Clic en 'Cerrar'realizado.")
                time.sleep(2)
            except Exception as e_cerrar:
                logging.warning(f"No se pudo dar clic en 'Cerrar': {type(e_cerrar).__name__} {e_cerrar}")

            logging.info("🎉 Creación de Cambio de Precio completada y sometida a aprobación.")
            return True

        except Exception as e:
            logging.error(f"Error durante la importación de Excel en Cambio de Precio: {e}")
            return False

    def obtener_archivos_por_tipo(self, tipo_cambio):
        """Retorna la lista de archivos .xls dentro de la carpeta del tipo correspondiente (PMD, PMU, TMD)."""
        import pathlib
        base_dir = pathlib.Path(r"c:\Users\eliseo_lopezp\automatizacion_web\archivos\cambio_precio") / tipo_cambio.upper()
        if not base_dir.exists():
            return []
        return sorted(base_dir.glob("*.xls"))

    def procesar_archivos_por_tipo(self, tipo_cambio, codigo_razon="03"):
        """
        Escanea la carpeta de un tipo de cambio (ej: PMD), llena el formulario por cada archivo .xls encontrado e importa el archivo.
        """
        archivos = self.obtener_archivos_por_tipo(tipo_cambio)
        if not archivos:
            logging.warning(f"No se encontraron archivos .xls en la carpeta: archivos/cambio_precio/{tipo_cambio}")
            return False

        logging.info(f"\n🗂️ Se encontraron {len(archivos)} archivo(s) .xls para el tipo '{tipo_cambio}'")
        for idx, archivo in enumerate(archivos, start=1):
            logging.info(f"\n--- [{idx}/{len(archivos)}] Procesando: {archivo.name} ---")
            self.acceder_crear_cambio_precio()
            self.llenar_formulario_crear_cambio_precio(
                tipo_cambio_precio=tipo_cambio,
                codigo_razon=codigo_razon,
                numero_evento="21",
                tipo_calculo='C',
                ordenamiento_reporte="DPT/SKU",
                clic_siguiente=True
            )
            self.importar_excel_cambio_precio(str(archivo.resolve()))
        return True

    def procesar_archivos_por_tipos(self, tipos_config):
        """
        Procesa múltiples tipos de cambio con sus códigos de razón y configuraciones.
        tipos_config debe ser una lista de diccionarios con al menos:
            - tipo_cambio: 'PMD', 'PMU' o 'TMD'
            - codigo_razon: '03', '04', '05', etc.
        Opcional:
            - descripcion
            - numero_evento
            - tipo_calculo
            - ordenamiento_reporte
        """
        if not tipos_config:
            logging.warning("No se recibió configuración de tipos de cambio para procesar.")
            return False

        resultado_global = True
        for tipo_config in tipos_config:
            tipo_cambio = str(tipo_config.get("tipo_cambio", "TMD")).upper()
            codigo_razon = str(tipo_config.get("codigo_razon", "05"))
            descripcion = tipo_config.get("descripcion", "CAMBIO PRECIO AUTOMATIZADO")
            numero_evento = str(tipo_config.get("numero_evento", "21"))
            tipo_calculo = tipo_config.get("tipo_calculo", "C")
            ordenamiento_reporte = tipo_config.get("ordenamiento_reporte", "DPT/SKU")

            archivos = self.obtener_archivos_por_tipo(tipo_cambio)
            if not archivos:
                logging.warning(f"No se encontraron archivos .xls en la carpeta: archivos/cambio_precio/{tipo_cambio}")
                resultado_global = False
                continue

            logging.info(f"\n🗂️ Se encontraron {len(archivos)} archivo(s) .xls para el tipo '{tipo_cambio}' con razón '{codigo_razon}'")
            for idx, archivo in enumerate(archivos, start=1):
                logging.info(f"\n--- [{idx}/{len(archivos)}] Procesando: {archivo.name} ---")
                self.acceder_crear_cambio_precio()
                self.llenar_formulario_crear_cambio_precio(
                    tipo_cambio_precio=tipo_cambio,
                    descripcion=descripcion,
                    codigo_razon=codigo_razon,
                    numero_evento=numero_evento,
                    tipo_calculo=tipo_calculo,
                    ordenamiento_reporte=ordenamiento_reporte,
                    clic_siguiente=True
                )
                self.importar_excel_cambio_precio(str(archivo.resolve()))
        return resultado_global


def ejecutar_sol_cambio_precio():
    """
    Script ejecutable independiente para la funcionalidad de Cambio de Precio en SOL.
    """
    driver = None
    try:
        logging.info("\n Iniciando automatización SOL - Cambio de Precio")
        logging.info("=" * 70)

        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        driver = webdriver.Chrome(options=options)
        driver.maximize_window()

        logging.info("\n Realizando login...")
        driver.get("http://was7tr1.siman.com/AccesoSUMMER/")

        login_page = LoginPage(driver)
        login_page.login("ELOPEZ", "MAY2025")
        logging.info("Login exitoso")

        time.sleep(10)

        logging.info("\n Navegando al menú SUMMER → SOL...")
        menu = MenuPage(driver)

        if not menu.acceder_frame_menu():
            raise Exception("No se pudo acceder al frame del menú")

        if not menu.hacer_click_en_summer():
            raise Exception("No se pudo hacer click en SUMMER")

        if not menu.click_SOL():
            raise Exception("No se pudo hacer click en SOL")

        logging.info("Acceso a módulo SOL completado.")

        # Instanciar módulo de Cambio de Precio
        sol_precio = SOLCambioPrecioClass(driver)
        if not sol_precio.acceder_crear_cambio_precio():
            raise Exception("No se pudo acceder a Crear Cambio de Precio")

        logging.info("\n Automatización de Cambio de Precio en SOL iniciada correctamente!")
        logging.info("=" * 70)

        time.sleep(30)
        return True

    except Exception as e:
        logging.error(f"\n ERROR en SOL Cambio de Precio: {e}")
        logging.info("=" * 70)
        return False


if __name__ == "__main__":
    ejecutar_sol_cambio_precio()
