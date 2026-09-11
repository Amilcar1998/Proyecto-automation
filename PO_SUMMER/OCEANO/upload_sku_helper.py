import logging
import pathlib
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def subir_excel(driver, nombre_excel="plantilla.xls"):
    """Sube el archivo Excel al formulario dentro del iframe 'frameImporSkuXlsF1'.
    El archivo debe estar en la carpeta <workspace>/archivos/<nombre_excel>."""
    # Cambiar al iframe
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "frameImporSkuXlsF1"))
    )
    # Localizar el input de archivo
    file_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'][name='file']"))
    )
    # Ruta del archivo
    base_dir = pathlib.Path(r"C:\\Users\\eliseo_lopezp\\automatizacion_web")
    excel_path = base_dir / "archivos" / nombre_excel
    if not excel_path.is_file():
        raise FileNotFoundError(f"Archivo no encontrado: {excel_path}")
    # Enviar la ruta al input
    file_input.send_keys(str(excel_path.resolve()))
    assert excel_path.name in file_input.get_attribute("value"), "Archivo no asignado"
    # Click en Submit
    submit_btn = driver.find_element(By.CSS_SELECTOR, "input[type='submit'][value='Submit']")
    submit_btn.click()
    # Esperar mensaje de éxito (ajustar XPath si es necesario)
    try:
        success_msg = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located(
                (By.XPATH, "//*[contains(text(),'Archivo cargado') or contains(text(),'Importación exitosa')]")
            )
        )
        logging.info("Importación completada:", success_msg.text)
    except Exception:
        logging.warning("No se detectó mensaje de éxito después de subir el archivo.")
    finally:
        driver.switch_to.default_content()
