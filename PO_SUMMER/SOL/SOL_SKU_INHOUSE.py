import logging
import os
import sys
import re
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Agregar el directorio raíz al path para importaciones
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from Page.Page_Login import LoginPage
from Page.page_menu import MenuPage


def _sanitizar_texto(texto: str) -> str:
    """
    Trunca el texto a un máximo de 20 caracteres, pero permite caracteres especiales
    para fines de validación.
    """
    if not isinstance(texto, str):
        return texto
    return texto[:20]


class SOLSkuInHouseClass:
    """Clase para manejar la navegación y acciones iniciales de SKU In-House en SOL."""

    def __init__(self, driver):
        self.driver = driver
        self.accordion_sku_inhouse_header = (By.XPATH, "//h3[contains(normalize-space(.), 'SKU In-House')]")
        self.link_ingresar_nuevo_sku = (By.ID, "1178")
        self.link_crear_sku_init = (By.ID, "1184")
        self.link_mantenimiento_usuarios = (By.ID, "1187")
        self.link_mantenimiento_source_vendor = (By.ID, "1896")
        self.link_mantenimiento_atributos_summer = (By.ID, "1897")

        # Locators para el formulario de "Ingresar nuevo SKU"
        self.txt_title = (By.ID, "title")
        self.select_importancy = (By.ID, "importancy")
        self.radio_dropshipping_no = (By.XPATH, "//input[@name='dropshiping' and @value='N']")
        self.txt_brand = (By.ID, "brand")
        self.txt_model = (By.ID, "model")
        self.textarea_description_short = (By.ID, "descriptionShort")
        self.txt_class_id = (By.ID, "classIdsp")
        self.txt_warranty = (By.ID, "warranty")
        self.txt_origin_country = (By.ID, "paisOrigen")
        self.txt_upc_vendor = (By.ID, "upcVendor")
        self.select_product_type = (By.ID, "productType")
        self.txt_production_time = (By.ID, "productionTime")
        self.select_sostenible = (By.ID, "sostenible")
        self.textarea_description_en = (By.ID, "descriptionEn")
        self.select_atributos = (By.ID, "attrSelect")
        self.btn_agregar_proveedor = (By.ID, "agregarVendor")
        self.select_prefijo = (By.ID, "prefix")

        # Locators para el modal de "Agregar Proveedor"
        self.modal_select_tipo_proveedor = (By.ID, "definitionVendor")
        self.modal_txt_proveedor = (By.ID, "vendor")
        self.modal_btn_agregar = (By.ID, "agregarVendorModal")
        self.modal_btn_cerrar = (By.ID, "cerrarVendor")
        self.modal_tabla_proveedores = (By.ID, "gbox_list-vendor")

    def acceder_menu_sku_inhouse(self):
        """Accede a la sección SKU In-House del menú SOL."""
        self.driver.switch_to.default_content()
        if not MenuPage(self.driver).acceder_frame_menu():
            logging.error("No se pudo acceder al frame de menú para SKU In-House.")
            return False

        try:
            header = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable(self.accordion_sku_inhouse_header)
            )
            header.click()
            # Espera dinámica: aguardar a que el primer link del menú sea visible.
            WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(self.link_ingresar_nuevo_sku)
            )
            logging.info("Accedió al menú SKU In-House.")
            return True
        except Exception as e:
            logging.error(f"Error al abrir el menú SKU In-House: {e}")
            return False

    def _click_link(self, locator, label):
        self.driver.switch_to.default_content()
        MenuPage(self.driver).acceder_frame_menu()

        try:
            link = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable(locator)
            )
            link.click()
            logging.info(f"Click en '{label}'realizado.")
            return True
        except TimeoutException:
            logging.error(f"Timeout al buscar el link '{label}'.")
            return False
        except NoSuchElementException:
            logging.error(f"No se encontró el link '{label}'.")
            return False
        except Exception as e:
            logging.error(f"Error al hacer click en '{label}': {e}")
            return False

    def _find_sku_ingreso_in_default_content(self, timeout: int = 20):
        """Intenta detectar la página de ingreso de SKU en el contenido principal sin usar frame."""
        self.driver.switch_to.default_content()
        locators = [
            (By.XPATH, "//form"),
            (By.XPATH, "//label[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sku')]/.."),
            (By.XPATH, "//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sku') or contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sku') or contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sku')]/.."),
            (By.XPATH, "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'guardar') or contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'save')]")
        ]

        for locator in locators:
            try:
                WebDriverWait(self.driver, timeout).until(
                    EC.presence_of_element_located(locator)
                )
                logging.info("Página de ingreso de SKU detectada en contenido principal sin frame.")
                return True
            except Exception:
                continue

        return False

    def _esperar_pagina_sku_ingreso(self, timeout: int = 20):
        """Espera a que la página de ingreso de SKU se abra dentro del frame trabajo o en contenido principal sin frame."""
        try:
            self.driver.switch_to.default_content()
            if MenuPage(self.driver).acceder_frame_trabajo():
                locators = [
                    (By.XPATH, "//form"),
                    (By.XPATH, "//label[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sku')]/.."),
                    (By.XPATH, "//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sku') or contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sku') or contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sku')]/.."),
                    (By.XPATH, "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'guardar') or contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'save')]")
                ]

                for locator in locators:
                    try:
                        WebDriverWait(self.driver, timeout).until(
                            EC.presence_of_element_located(locator)
                        )
                        logging.info("Página de ingreso de SKU detectada dentro del frame 'trabajo'.")
                        return True
                    except Exception:
                        continue

                logging.warning("No se detectó claramente la página de ingreso de SKU dentro del frame 'trabajo'. Intentando contenido principal sin frame.")
            else:
                logging.warning("Frame 'trabajo'no disponible después de ingresar nuevo SKU. Intentando contenido principal sin frame.")

            if self._find_sku_ingreso_in_default_content(timeout):
                return True

            logging.error("No se detectó la página de ingreso de SKU ni en frame 'trabajo'ni en contenido principal.")
            return False
        except Exception as e:
            logging.error(f"Error al esperar la página de ingreso de SKU: {e}")
            return False

    def ingresar_nuevo_sku(self):
        if not self._click_link(self.link_ingresar_nuevo_sku, "SKU In-House -> Ingresar nuevo SKU"):
            return False
        return self._esperar_pagina_sku_ingreso()

    def crear_sku_init(self):
        return self._click_link(self.link_crear_sku_init, "SKU In-House -> Crear SKU Init")

    def mantenimiento_usuarios(self):
        return self._click_link(self.link_mantenimiento_usuarios, "SKU In-House -> Mantenimiento de Usuarios")

    def mantenimiento_source_vendor(self):
        return self._click_link(self.link_mantenimiento_source_vendor, "SKU In-House -> Mantenimiento de Source Vendor")

    def mantenimiento_atributos_summer(self):
        return self._click_link(self.link_mantenimiento_atributos_summer, "SKU In-House -> Mantenimiento Atributos en SUMMER")

    def _manejar_modal_proveedores(self, proveedores: list):
        """
        Maneja la lógica de agregar proveedores primarios y secundarios en la ventana modal.
        """
        try:
            logging.info("Agregando proveedores...")
            self.driver.find_element(*self.btn_agregar_proveedor).click()

            # Esperar a que el modal y su contenido estén visibles
            WebDriverWait(self.driver, 15).until(
                EC.visibility_of_element_located(self.modal_tabla_proveedores)
            )
            logging.info("Modal de proveedores abierto.")

            for prov in proveedores:
                tipo = prov.get("tipo", "Primario")
                codigo = prov.get("codigo")

                if not codigo:
                    continue

                logging.info(f"   -> Agregando proveedor {tipo}: {codigo}")
                # Seleccionar tipo (Primario/Secundario)
                select_tipo = Select(self.driver.find_element(*self.modal_select_tipo_proveedor))
                select_tipo.select_by_visible_text(tipo)

                # Ingresar código y presionar Enter
                input_prov = self.driver.find_element(*self.modal_txt_proveedor)
                input_prov.clear()
                input_prov.send_keys(codigo)
                time.sleep(1) # Pausa para que el autocompletado reaccione
                input_prov.send_keys(Keys.ENTER)
                
                # Esperar a que el AJAX termine y la tabla se actualice
                WebDriverWait(self.driver, 10).until(lambda d: self.driver.execute_script("return jQuery.active == 0"))
                time.sleep(1)

                # Hacer clic en "Agregar proveedor" dentro del modal
                self.driver.find_element(*self.modal_btn_agregar).click()
                
                # Esperar a que la tabla se refresque con el nuevo proveedor
                WebDriverWait(self.driver, 15).until(
                    EC.text_to_be_present_in_element((By.ID, "list-vendor"), codigo)
                )
                logging.info(f"  Proveedor {codigo} agregado a la tabla.")

            # Cerrar el modal
            self.driver.find_element(*self.modal_btn_cerrar).click()
            logging.info("Modal de proveedores cerrado.")
            return True
        except Exception as e:
            logging.error(f"Error al manejar el modal de proveedores: {e}")
            return False

    def llenar_formulario_nuevo_sku(self, datos_sku: dict):
        """
        Llena el formulario de 'Ingresar nuevo SKU' con los datos proporcionados,
        sanitizando los campos de texto para evitar caracteres especiales.
        """
        try:
            logging.info("Llenando formulario de nuevo SKU con validación...")
            self.driver.switch_to.default_content()
            MenuPage(self.driver).acceder_frame_trabajo()

            # --- Llenado de campos con sanitización ---
            if 'title' in datos_sku:
                self.driver.find_element(*self.txt_title).send_keys(datos_sku['title'][:20])
            
            if 'model' in datos_sku:
                self.driver.find_element(*self.txt_model).send_keys(datos_sku['model'][:20])

            if 'descriptionShort' in datos_sku:
                self.driver.find_element(*self.textarea_description_short).send_keys(_sanitizar_texto(datos_sku['descriptionShort']))

            if 'warranty' in datos_sku:
                # Asume que garantía es numérica, pero sanitiza por si acaso
                self.driver.find_element(*self.txt_warranty).send_keys(re.sub(r'[^0-9]', '', str(datos_sku['warranty'])))

            if 'upcVendor' in datos_sku:
                self.driver.find_element(*self.txt_upc_vendor).send_keys(datos_sku['upcVendor'][:20])

            if 'productionTime' in datos_sku:
                self.driver.find_element(*self.txt_production_time).send_keys(re.sub(r'[^0-9.]', '', str(datos_sku['productionTime'])))

            if 'descriptionEn' in datos_sku:
                self.driver.find_element(*self.textarea_description_en).send_keys(datos_sku['descriptionEn'])

            # --- Llenado de campos con autocompletado y eventos ---
            if 'brand' in datos_sku:
                 campo_marca = self.driver.find_element(*self.txt_brand)
                 campo_marca.send_keys(datos_sku['brand'])
                 time.sleep(1)
                 campo_marca.send_keys(Keys.ENTER)

            if 'classId' in datos_sku:
                campo_clase = self.driver.find_element(*self.txt_class_id)
                campo_clase.send_keys(datos_sku['classId'])
                # Esperar a que el prefijo se cargue dinámicamente
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, "//select[@id='prefix']/option[string-length(text()) > 0]"))
                )
                logging.info("Prefijo cargado dinámicamente.")
                # Volver a encontrar el elemento para evitar StaleElementReferenceException
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable(self.txt_class_id)
                ).send_keys(Keys.ENTER)

            if 'originCountry' in datos_sku:
                campo_pais = self.driver.find_element(*self.txt_origin_country)
                campo_pais.send_keys(datos_sku['originCountry'])
                time.sleep(1)
                campo_pais.send_keys(Keys.ENTER)

            if 'dropshipping' in datos_sku and datos_sku['dropshipping'] == 'N':
                self.driver.find_element(*self.radio_dropshipping_no).click()

            # --- Llenado de campos SELECT ---
            if 'importancy' in datos_sku:
                Select(self.driver.find_element(*self.select_importancy)).select_by_value(datos_sku['importancy'])

            if 'productType' in datos_sku:
                Select(self.driver.find_element(*self.select_product_type)).select_by_value(datos_sku['productType'])

            if 'sostenible' in datos_sku:
                Select(self.driver.find_element(*self.select_sostenible)).select_by_value(datos_sku['sostenible'])

            # --- Interacción con el SELECT dinámico de Atributos ---
            if 'atributo' in datos_sku:
                logging.info(f"Seleccionando atributo dinámico: {datos_sku['atributo']}...")
                select_attr = Select(self.driver.find_element(*self.select_atributos))
                select_attr.select_by_visible_text(datos_sku['atributo'])
                
                # Espera crucial para que el contenido dinámico cargue
                if datos_sku['atributo'] == 'LOGISTICS VALUES':
                    WebDriverWait(self.driver, 15).until(
                        EC.visibility_of_element_located((By.ID, "divLogisticValues"))
                    )
                    logging.info("Contenido de 'LOGISTICS VALUES'cargado.")
                elif datos_sku['atributo'] == 'FEATURES AND BENEFITS':
                    WebDriverWait(self.driver, 15).until(
                        EC.visibility_of_element_located((By.ID, "divBenefits"))
                    )
                    logging.info("Contenido de 'FEATURES AND BENEFITS'cargado.")
                elif datos_sku['atributo'] == 'SUBSTITUTE':
                     WebDriverWait(self.driver, 15).until(
                        EC.visibility_of_element_located((By.ID, "divSustitutos"))
                    )
                     logging.info("Contenido de 'SUBSTITUTE'cargado.")
                # Pausa dinámica: espera a que se terminen las peticiones AJAX que cargan el contenido.
                WebDriverWait(self.driver, 10).until(
                    lambda d: self.driver.execute_script("return (typeof jQuery === 'undefined') || (jQuery.active === 0);")
                )
                logging.info("Contenido dinámico estabilizado (AJAX completado).")

            # --- Manejo del modal de proveedores ---
            if 'proveedores' in datos_sku:
                self._manejar_modal_proveedores(datos_sku['proveedores'])


            logging.info("Formulario principal y modal de proveedores completados.")
            return True
        except Exception as e:
            logging.error(f"Error al llenar el formulario de nuevo SKU: {e}")
            return False


def ejecutar_sku_inhouse_sol(opcion="ingresar_nuevo_sku"):
    """Ejecuta el flujo inicial de SKU In-House en SOL."""
    driver = None
    try:
        logging.info("\n Iniciando automatización SOL - SKU In-House")
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

        logging.info("\n Navegando al menú SUMMER → SOL...")
        menu = MenuPage(driver)
        # Espera dinámica: el método acceder_frame_menu ya contiene las esperas necesarias.
        if not menu.acceder_frame_menu():
            raise Exception("No se pudo acceder al frame del menú")
        if not menu.hacer_click_en_summer(): # Este método también tiene esperas internas.
            raise Exception("No se pudo hacer click en SUMMER")
        if not menu.click_SOL():
            raise Exception("No se pudo hacer click en SOL")

        sku_inhouse = SOLSkuInHouseClass(driver)
        if not sku_inhouse.acceder_menu_sku_inhouse():
            raise Exception("No se pudo acceder al menú SKU In-House")

        if opcion == "ingresar_nuevo_sku":
            if sku_inhouse.ingresar_nuevo_sku():
                # Datos de ejemplo para el nuevo SKU
                datos_a_ingresar = {
                    "title": "Televisor 55' 4K @#$% con un nombre muy largo que debe ser cortado",
                    "model": "TV55-ULTRA-!2024",
                    "descriptionShort": "Televisor inteligente con pantalla 4K y HDR.",
                    "warranty": "12",
                    "upcVendor": "123456789012-ABC",
                    "productionTime": "30",
                    "descriptionEn": "Smart TV with 4K screen and HDR support.",
                    "brand": "000709",
                    "classId": "55B",
                    "originCountry": "CHN",
                    "dropshipping": "N",
                    "importancy": "U", # 'N' para Normal, 'U' para Urgente
                    "productType": "R", # 'R', 'NR', 'CE'
                    "sostenible": "Y", # 'Y' o 'N'
                    "atributo": "LOGISTICS VALUES", # El texto visible de la opción
                    "proveedores": [
                        {"tipo": "Primario", "codigo": "010479"},
                        {"tipo": "Secundario", "codigo": "012713"}
                    ]
                }
                sku_inhouse.llenar_formulario_nuevo_sku(datos_a_ingresar)
        elif opcion == "crear_sku_init":
            sku_inhouse.crear_sku_init()
        elif opcion == "mantenimiento_usuarios":
            sku_inhouse.mantenimiento_usuarios()
        elif opcion == "mantenimiento_source_vendor":
            sku_inhouse.mantenimiento_source_vendor()
        elif opcion == "mantenimiento_atributos_summer":
            sku_inhouse.mantenimiento_atributos_summer()
        else:
            raise Exception(f"Opción desconocida de SKU In-House: {opcion}")

        logging.info("Flujo SKU In-House iniciado correctamente.")
        return True

    except Exception as e:
        logging.error(f"\n ERROR en SOL SKU In-House: {e}")
        logging.info("=" * 70)
        return False


if __name__ == "__main__":
    ejecutar_sku_inhouse_sol()
