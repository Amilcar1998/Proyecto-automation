import logging
import pathlib
import sys
import traceback
import time

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


# Carpetas donde están los archivos .xls a subir
CARPETA_ARCHIVOS = pathlib.Path(r"c:\Users\eliseo_lopezp\automatizacion_web\archivos")
CARPETA_IMPORTADAS = CARPETA_ARCHIVOS / "importadas"
CARPETA_LOCALES = CARPETA_ARCHIVOS / "locales"


def obtener_archivos_con_tipo() -> list[tuple[pathlib.Path, str]]:
    """Retorna la lista de tuplas (archivo_path, tipo_po) buscando en:
    - 'archivos/importadas' -> 'Importada'
    - 'archivos/locales' -> 'Local'
    - Archivos sueltos en 'archivos' -> 'Importada' (por defecto)
    """
    CARPETA_IMPORTADAS.mkdir(parents=True, exist_ok=True)
    CARPETA_LOCALES.mkdir(parents=True, exist_ok=True)

    archivos_con_tipo = []

    # 1. Archivos en carpetas específicas
    for f in sorted(CARPETA_IMPORTADAS.glob("*.xls")):
        archivos_con_tipo.append((f, "Importada"))

    for f in sorted(CARPETA_LOCALES.glob("*.xls")):
        archivos_con_tipo.append((f, "Local"))

    # 2. Archivos sueltos en la raíz de 'archivos'
    for f in sorted(CARPETA_ARCHIVOS.glob("*.xls")):
        archivos_con_tipo.append((f, "Importada"))

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
        # El contexto del driver ya está en 'trabajo'.
        # Cambiamos al iframe del diálogo de importación.
        logging.info("Cambiando al iframe 'frameImporSkuXlsF1'...")
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("frameImporSkuXlsF1"))
        WebDriverWait(driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    except TimeoutException:
        logging.error("ERROR: No se pudo encontrar o cambiar al iframe 'frameImporSkuXlsF1'.")
        raise

    # 2. Asignar la ruta del archivo al input
    try:
        # Esperar explícitamente a que aparezca el input de archivo
        logging.info("Buscando el campo de entrada de archivo (input[type='file'])...")
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'], input[name='file']"))
        )

        # TRUCO CLAVE 1: Forzar la visibilidad del input por si está oculto
        driver.execute_script(
            "arguments[0].style.display = 'block'; arguments[0].style.visibility = 'visible';",
            file_input
        )
        time.sleep(0.5)

        logging.info(f"Asignando ruta: '{full_path}'")
        # Inyectar la ruta directamente, SIN usar .clear()
        file_input.send_keys(full_path)
        time.sleep(1)  # Pausa para que el navegador procese la asignación

        val_asignado = file_input.get_attribute("value")
        if val_asignado:
            logging.info(f"Archivo asignado en el input. Valor: '{val_asignado}'")
        else:
            # A veces por seguridad los navegadores no devuelven el 'value', pero el archivo se asigna.
            logging.warning("El navegador no muestra el 'value'en el input, pero se continuará con el envío.")

    except TimeoutException:
        logging.error("ERROR: No se encontró el input de tipo 'file'en el iframe después de 10 segundos.")
        raise NoSuchElementException("No se pudo localizar el input para subir el archivo.")
    except Exception as e:
        logging.error(f"ERROR inesperado al asignar el archivo: {e}")
        raise

    # 3. Hacer clic en el botón de 'Submit' para cargar el archivo
    try:
        # TRUCO CLAVE 2: Buscar <input> o <button> de tipo submit o con texto relevante.
        logging.info("Buscando el botón para enviar el archivo...")
        xpath_submit = "//input[@type='submit'] | //button[@type='submit'] | //input[contains(translate(@value,'CARGAR','cargar'), 'cargar')] | //button[contains(translate(text(),'CARGAR','cargar'), 'cargar')]"
        btn_submit = driver.find_element(By.XPATH, xpath_submit)

        # Asegurarse de que el botón esté visible y clickeable
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_submit)
        time.sleep(0.5)

        btn_submit.click()
        logging.info("Clic en el botón de envío realizado.")

    except NoSuchElementException:
        logging.warning("No se encontró un botón de envío estándar. Intentando enviar el formulario vía JS.")
        driver.execute_script("document.forms[0] && document.forms[0].submit();")
        logging.info("Formulario enviado vía JS.")
    except Exception as e:
        logging.error(f"ERROR inesperado al hacer clic en el botón de envío: {e}")
        raise

    # 4. Esperar la confirmación del servidor (lógica mejorada)
    try:
        logging.info("Esperando respuesta del servidor...")
        # Esperar hasta 30 segundos por un elemento que contenga un mensaje de éxito.
        # Esto es más fiable que hacer un sondeo del código fuente completo.
        success_xpath = "//*[contains(text(), 'insertados') or contains(text(), 'exito') or contains(text(), 'éxito') or contains(text(), 'procesado') or contains(text(), 'OK')]"

        # Primero, verificar si hay un mensaje de error visible
        error_xpath = "//*[contains(text(), 'error') or contains(text(), 'exception')]"
        error_elements = driver.find_elements(By.XPATH, error_xpath)
        if any(el.is_displayed() for el in error_elements):
            error_message = next((el.text for el in error_elements if el.is_displayed()), "Error no especificado")
            logging.error(f"El servidor devolvió un error visible: {error_message}")
            raise Exception(f"El servidor reportó un error en la carga del archivo: {error_message}")

        # Si no hay errores visibles, esperar el mensaje de éxito
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, success_xpath))
        )
        logging.info("Confirmación recibida: El archivo se procesó en el backend.")

    except TimeoutException:
        logging.error("ERROR: No se recibió confirmación de éxito del servidor después de 30 segundos.")
        raise
    except Exception as e:
        # Capturar otros errores, como el que lanzamos manualmente por error del servidor
        logging.info(f"Error durante la espera de confirmación: {e}")
        raise

    # Regresar al contexto principal para los siguientes pasos
    driver.switch_to.default_content()


def crear_po_para_archivo(driver, menu: MenuPage, archivo: pathlib.Path, tipo_po: str, indice: int, total: int):
    """Flujo completo de creación de PO para un único archivo .xls."""
    logging.info(f"\n{'='*60}")
    logging.info(f"[PROCESANDO] [{indice}/{total}] ({tipo_po}): {archivo.name}")
    logging.info(f"{'='*60}")

    try:
        # Asegurar contexto principal y esperar estabilidad
        menu.driver.switch_to.default_content()
        menu.esperar_pagina_lista(timeout=25)
        time.sleep(2)

        # PASO 1: Click en "Digitar PO" desde el menú
        menu.acceder_frame_menu()
        menu.esperar_pagina_lista()
        menu.crear_po_menu()   # expande el submenú "Crear PO" (si no está desplegado)
        time.sleep(1)
        menu.crear_po()        # click en "Digitar PO / Crear PO Venta" (Titulo1Opcion69)
        time.sleep(3)          # Esperar a que el frame 'trabajo' inicie la carga del nuevo formulario PO

        # PASO 2: Ingresar descripción del PO
        menu.acceder_frame_trabajo()
        menu.esperar_pagina_lista()
        main_descripcion(driver)

        # PASO 3: Llenar encabezado del PO con su tipo (Importada o Local)
        llenar_po(driver, tipo_po=tipo_po)

        # PASO 3.1: Hacer clic en "Continuar" para ir a la sección de importación
        try:
            logging.info("Buscando y haciendo clic en el botón 'Continuar'...")
            # Asegurarse de que el driver está en el contexto correcto ('trabajo')
            menu.acceder_frame_trabajo()

            # Intento 1: Buscar por ID y hacer clic. Es el método más fiable.
            continuar_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "next"))
            )
            continuar_button.click()
            logging.info("Clic en 'Continuar'(ID='next') realizado.")
            time.sleep(3)  # Esperar a que la sección de importación cargue
        except TimeoutException:
            logging.warning("No se pudo hacer clic en el botón 'Continuar'por ID. Intentando con JavaScript.")
            try:
                # Intento 2: Forzar el clic en el botón 'next' usando JavaScript.
                menu.acceder_frame_trabajo()
                driver.execute_script("document.getElementById('next').click();")
                logging.info("Se forzó el clic en el botón con ID 'next'mediante JavaScript.")
                time.sleep(3)
            except Exception as e:
                logging.error(f"ERROR: Fallaron todos los intentos de hacer clic en 'Continuar'. Excepción: {e}")
                raise

        # PASO 3.2: Hacer clic en "Importar SKU" para abrir el diálogo de carga de archivos
        try:
            logging.info("Buscando y haciendo clic en el botón 'Importar SKU'...")
            # Asegurarse de que el driver está en el contexto correcto ('trabajo')
            menu.acceder_frame_trabajo()

            # Usar el ID único del botón para la máxima fiabilidad
            importar_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "importarSkusXlsF1"))
            )
            importar_button.click()
            logging.info("Clic en 'Importar SKU'realizado.")
            time.sleep(2) # Esperar a que el diálogo de importación aparezca
        except TimeoutException:
            logging.error("ERROR: No se pudo encontrar o hacer clic en el botón 'Importar SKU'.")
            raise

        # PASO 4: Importar SKU desde el archivo actual
        importar_sku_desde_archivo(menu, str(archivo.resolve()))

        # Extraer Número Maestro de PO generado
        po_numero = menu.obtener_numero_po()

        # PASO 5: Ir a Resumen de Orden de Compra
        menu.ir_a_resumen_orden()

        # PASO 6: Enviar a Aprobación
        menu.enviar_a_aprobacion()

        logging.info(f"[EXITO] PO ({tipo_po}) N° Maestro [{po_numero}] creada exitosamente para: {archivo.name}")
        time.sleep(3)

    except Exception as e:
        logging.info(f"\n[ERROR FATAL] Ocurrió un error procesando '{archivo.name}'. Deteniendo el flujo para este archivo.")
        traceback.print_exc()


def test_login_exitoso():
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get("http://was7tr1.siman.com/AccesoSUMMER/")

    login_page = LoginPage(driver)
    login_page.login("ELOPEZ", "MAY2025")
    time.sleep(10)

    menu = MenuPage(driver)

    # Navegar a OCEANO una sola vez antes del bucle
    menu.acceder_frame_menu()
    menu.hacer_click_en_summer()
    menu.click_oceano()
    time.sleep(2)

    # Leer archivos .xls clasificando por tipo (Importada / Local)
    archivos = obtener_archivos_con_tipo()
    if not archivos:
        logging.info(f"[AVISO] No se encontraron archivos .xls en: {CARPETA_ARCHIVOS}")
        driver.quit()
        return

    total = len(archivos)
    logging.info(f"\n[ARCHIVOS] Se encontraron {total} archivo(s) .xls para procesar.\n")

    # Bucle: una PO por archivo, cada una empieza desde "Crear PO" en el menú
    for indice, (archivo, tipo_po) in enumerate(archivos, start=1):
        crear_po_para_archivo(driver, menu, archivo, tipo_po, indice, total)

    logging.info(f"\n[FIN] Proceso completado: {total} PO(s) procesadas.")
    time.sleep(5)
    # driver.quit()


if __name__ == "__main__":
    test_login_exitoso()