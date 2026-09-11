import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
import time
import json
import os
import re

class MenuPage:
    def __init__(self, driver):
        self.driver = driver
        self.boton_summer = (By.ID, "Titulo3")  # ID del botón "SUMMER"

    def _wait_for_frame(self, frame_name, timeout=15):
        """Espera a que el frame esté disponible y cambia a él.
        Retorna True si el cambio fue exitoso, False en caso de timeout.
        """
        try:
            self.driver.switch_to.default_content()
            WebDriverWait(self.driver, timeout).until(
                EC.frame_to_be_available_and_switch_to_it(frame_name)
            )
            logging.info(f"[OK] Cambiado al frame '{frame_name}'")
            return True
        except TimeoutException:
            logging.info(f"[ERROR] Timeout esperando el frame '{frame_name}'")
            return False
        except Exception as e:
            logging.info(f"[ERROR] Error al cambiar al frame '{frame_name}': {e}")
            return False

    def esperar_pagina_lista(self, timeout=20):
        """Espera a que el DOM esté en readyState=complete y jQuery sin peticiones AJAX activas."""
        try:
            WebDriverWait(self.driver, timeout).until(
                lambda d: d.execute_script(
                    "return document.readyState === 'complete' "
                    "&& (typeof jQuery === 'undefined' || jQuery.active === 0);"
                )
            )
        except Exception:
            time.sleep(2)  # fallback si el script falla

    def _find_menu_in_default_content(self, timeout=10):
        try:
            self.driver.switch_to.default_content()
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#accordion, .siman-menu-header, .siman-menu-contenido-link"))
            )
            logging.info("[OK] No se usa frame para el menú; contenido principal detectado.")
            return True
        except Exception:
            return False

    def _find_menu_in_frames(self, timeout=10):
        self.driver.switch_to.default_content()
        frames = self.driver.find_elements(By.TAG_NAME, 'frame') + self.driver.find_elements(By.TAG_NAME, 'iframe')
        for index, frame in enumerate(frames):
            try:
                self.driver.switch_to.default_content()
                WebDriverWait(self.driver, timeout).until(
                    EC.frame_to_be_available_and_switch_to_it(index)
                )
                if self._find_menu_in_default_content(timeout=2):
                    logging.info(f"[OK] Menú detectado dentro del frame índice {index}.")
                    return True
            except Exception:
                continue
        self.driver.switch_to.default_content()
        return False

    def acceder_frame_menu(self):
        """Accede al frame 'menu' usando _wait_for_frame.
        Si no hay un frame 'menu', intenta usar el contenido principal donde puede vivir el menú.
        """
        # Try by name first
        if self._wait_for_frame("menu", timeout=30):
            return True

        # Fallback: try by index (second frame, index 1)
        logging.info("[WARNING] Frame 'menu' no encontrado por nombre; intentando fallback por índice.")
        if self._wait_for_frame(1, timeout=30):
            return True

        # Fallback a contenido principal cuando el menú ya está disponible directamente
        if self._find_menu_in_default_content(timeout=10):
            return True

        # Fallback a cualquier frame disponible que pueda contener el menú
        logging.info("[WARNING] No se encontró el menú en contenido principal; buscando en frames adicionales.")
        if self._find_menu_in_frames(timeout=10):
            return True

        logging.info("[ERROR] No se pudo acceder al menú en ningún frame ni en el contenido principal.")
        return False

    def hacer_click_en_summer(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(self.boton_summer)
            )
            boton = self.driver.find_element(*self.boton_summer)
            boton.click()
            logging.info("[OK] Click en 'SUMMER' realizado")
            # Minimal pause to allow UI update
            time.sleep(1)
            return True
        except TimeoutException:
            logging.info("[ERROR] Tiempo de espera agotado. No se encontró el botón 'SUMMER'.")
            return False
        except NoSuchElementException as e:
            logging.info("[ERROR] No se encontró el botón 'SUMMER'.")
            logging.info(e)
            return False
        except Exception as e:
            logging.info(f"[ERROR] Error inesperado al hacer clic en 'SUMMER': {e}")
            return False


    def click_oceano(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "Titulo3Opcion10"))
            )
            boton_oceano = self.driver.find_element(By.ID, "Titulo3Opcion10")
            boton_oceano.click()
            logging.info("[OK] Click en 'OCEANO' realizado")
            # Short pause to allow UI update
            time.sleep(1)
            return True
        except TimeoutException:
            logging.info("[ERROR] Tiempo de espera agotado. No se encontró el botón 'OCEANO'.")
            return False
        except NoSuchElementException as e:
            logging.info("[ERROR] No se encontró el botón 'OCEANO'.")
            logging.info(e)
            return False
        except Exception as e:
            logging.info(f"[ERROR] Error inesperado al hacer clic en 'OCEANO': {e}")
            return False
        
    def click_SOL(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "Titulo3Opcion9"))
            )
            boton_sol = self.driver.find_element(By.ID, "Titulo3Opcion9")
            boton_sol.click()
            logging.info("[OK] Click en 'SOL' realizado")
            # Short pause to allow UI update
            time.sleep(1)
            return True
        except TimeoutException:
            logging.info("[ERROR] Tiempo de espera agotado. No se encontró el botón 'SOL'.")
            return False
        except NoSuchElementException as e:
            logging.info("[ERROR] No se encontró el botón 'SOL'.")
            logging.info(e)
            return False
        except Exception as e:
            logging.info(f"[ERROR] Error inesperado al hacer clic en 'SOL': {e}")
            return False
    
    def click_swim(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "Titulo3Opcion7"))
            )
            boton_terra = self.driver.find_element(By.ID, "Titulo3Opcion7")
            boton_terra.click()
            logging.info("[OK] Click en 'swim' realizado")
            # Short pause to allow UI update
            time.sleep(1)
            return True
        except TimeoutException:
            logging.info("[ERROR] Tiempo de espera agotado. No se encontró el botón 'swim'.")
            return False
        except NoSuchElementException as e:
            logging.info("[ERROR] No se encontró el botón 'swim'.")
            logging.info(e)
            return False
        except Exception as e:
            logging.info(f"[ERROR] Error inesperado al hacer clic en 'swim': {e}")
            return False
    
    def crear_po_menu(self):
        try:
            # Si el submenú ("Titulo1Opcion69") ya está desplegado y visible, no hacer click en Titulo1 para evitar replegarlo
            try:
                opcion = self.driver.find_element(By.ID, "Titulo1Opcion69")
                if opcion.is_displayed():
                    logging.info("[OK] Submenú 'Crear PO' ya se encuentra desplegado.")
                    return True
            except Exception:
                pass

            # Esperar que el botón sea clickable (no solo presente)
            WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.ID, "Titulo1"))
            ).click()
            logging.info("[OK] Click en 'Crear PO' realizado")
            time.sleep(1)
            return True
        except TimeoutException:
            logging.info("[ERROR] Timeout esperando el botón 'Crear PO'.")
            return False
        except Exception as e:
            logging.info(f"[ERROR] Error inesperado al hacer clic en 'Crear PO': {e}")
            return False

    def crear_po(self):
        try:
            # Esperar a que cualquier capa de carga (blockUI) desaparezca antes de interactuar
            try:
                WebDriverWait(self.driver, 60).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, ".blockUI.blockOverlay"))
                )
            except Exception:
                logging.warning("Advertencia: blockUI overlay sigue presente después de 60s en crear_po")
                
            # Esperar que el submenú sea clickable (puede estar oculto si el menú no desplegó aún)
            WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.ID, "Titulo1Opcion69"))
            ).click()
            logging.info("[OK] Click en 'Digitar PO' (Crear PO Venta) realizado")
            time.sleep(1)
            return True
        except TimeoutException:
            logging.info("[ERROR] Timeout esperando el botón 'Digitar PO'. Reintentando desplegar menú...")
            try:
                self.driver.find_element(By.ID, "Titulo1").click()
                time.sleep(1)
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "Titulo1Opcion69"))
                ).click()
                logging.info("[OK] Click en 'Digitar PO' realizado en el reintento")
                return True
            except Exception as ex:
                logging.info(f"[ERROR] No se pudo hacer clic en 'Digitar PO' tras reintento: {ex}")
                return False
        except Exception as e:
            logging.info(f"[ERROR] Error inesperado al hacer clic en 'Digitar PO': {e}")
            return False
    
    def acceder_frame_trabajo(self):
        """Accede al frame 'trabajo' usando _wait_for_frame."""
        return self._wait_for_frame("trabajo")

    def _descartar_alertas(self):
        """Acepta inmediatamente cualquier alerta emergente JS abierta sin retardos innecesarios."""
        for _ in range(3):
            try:
                alert = self.driver.switch_to.alert
                text = alert.text
                if text:
                    logging.warning(f"Alerta pendiente detectada: '{text}'. Aceptando...")
                    alert.accept()
                else:
                    break
            except Exception:
                break

    def _switch_to_import_iframe(self):
        """
        Garantiza el cambio al iframe de importación de SKU (ej. id='frameImporSkuXlsF1', src='/OCEANO/uploadImportF1.jsp').
        Recorre 'trabajo', sub-iframes de país (country0, country3, etc.) y default_content.
        """
        # A) Buscar directamente en trabajo
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            frames = self.driver.find_elements(By.TAG_NAME, "iframe")
            for frame in frames:
                f_id = frame.get_attribute("id") or ""
                f_src = frame.get_attribute("src") or ""
                if "impor" in f_id.lower() or "uploadimport" in f_src.lower():
                    self.driver.switch_to.frame(frame)
                    logging.info(f"Cambiado a iframe de importación en 'trabajo': id='{f_id}'")
                    return True
        except Exception:
            pass

        # B) Buscar anidado dentro de sub-iframes (país: country3, country0, etc.)
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            sub_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            for sub_frame in sub_iframes:
                try:
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame("trabajo")
                    self.driver.switch_to.frame(sub_frame)
                    nested_frames = self.driver.find_elements(By.TAG_NAME, "iframe")
                    for nested in nested_frames:
                        n_id = nested.get_attribute("id") or ""
                        n_src = nested.get_attribute("src") or ""
                        if "impor" in n_id.lower() or "uploadimport" in n_src.lower():
                            self.driver.switch_to.frame(nested)
                            logging.info(f"Cambiado a iframe de importación anidado en país: id='{n_id}'")
                            return True
                except Exception:
                    continue
        except Exception:
            pass

        # C) Fallback directo por selector XPath en 'trabajo'
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            iframe_locator = (By.XPATH, "//iframe[contains(@id, 'frameImpor') or contains(@id, 'frameImport') or contains(@src, 'uploadImport')]")
            WebDriverWait(self.driver, 5).until(
                EC.frame_to_be_available_and_switch_to_it(iframe_locator)
            )
            logging.info("Cambiado a iframe de importación por selector XPath.")
            return True
        except Exception:
            pass

        return False

    def _hacer_click_continuar(self):
        """
        Localiza y hace click en el botón 'Continuar' (id='next', id='continuarPO', value='Continuar'),
        buscando dinámicamente en 'trabajo' y en los sub-iframes de país (country0, country3, etc.).
        """
        selectores_continuar = [
            (By.ID, "next"),
            (By.ID, "continuarPO"),
            (By.ID, "btnNext"),
            (By.XPATH, "//input[@id='next' or @id='continuarPO' or contains(@value,'Continuar') or contains(@value,'continuar')]"),
            (By.XPATH, "//button[contains(text(),'Continuar') or contains(@id,'next') or contains(@id,'continuar')]")
        ]

        # A) Probar en sub-iframes de país primero (country3, country0, etc.)
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            sub_frames = self.driver.find_elements(By.TAG_NAME, "iframe")
            for sub in sub_frames:
                try:
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame("trabajo")
                    self.driver.switch_to.frame(sub)
                    for selector in selectores_continuar:
                        elems = self.driver.find_elements(*selector)
                        for elem in elems:
                            if elem.is_displayed():
                                self.driver.execute_script("""
                                    var btn = arguments[0];
                                    btn.removeAttribute('disabled');
                                    btn.scrollIntoView({block: 'center', inline: 'nearest'});
                                """, elem)
                                try:
                                    elem.click()
                                except Exception:
                                    self.driver.execute_script("arguments[0].click();", elem)
                                logging.info("Click en 'Continuar'realizado (en iframe de país)")
                                return True
                except Exception:
                    continue
        except Exception:
            pass

        # B) Probar en 'trabajo' directamente
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            for selector in selectores_continuar:
                elems = self.driver.find_elements(*selector)
                for elem in elems:
                    if elem.is_displayed():
                        self.driver.execute_script("""
                            var btn = arguments[0];
                            btn.removeAttribute('disabled');
                            btn.scrollIntoView({block: 'center', inline: 'nearest'});
                        """, elem)
                        try:
                            elem.click()
                        except Exception:
                            self.driver.execute_script("arguments[0].click();", elem)
                        logging.info("Click en 'Continuar'realizado (en frame 'trabajo')")
                        return True
        except Exception:
            pass

        return False

    def _hacer_click_importar_sku(self):
        """
        Localiza y hace click en el botón 'Importar SKU' (id='importarSkusXlsF1', title='Importar Sku'),
        esperando dinámicamente a que la página 'trabajo.jsp' termine de cargar.
        """
        selectores_importar = [
            (By.ID, "importarSkusXlsF1"),
            (By.XPATH, "//input[@id='importarSkusXlsF1' or contains(@id,'importarSkusXls') or contains(@title,'Importar Sku') or contains(@onclick,'importSkusXls')]"),
            (By.XPATH, "//*[contains(@title,'Importar Sku') or contains(@title,'Importar SKU') or contains(@value,'Importar Sku') or contains(@value,'Importar SKU') or contains(text(),'Importar Sku')]"),
            (By.XPATH, "//input[@class='excel']")
        ]

        end_time = time.time() + 15
        while time.time() < end_time:
            self._descartar_alertas()

            # A) Probar en 'trabajo' directamente primero (donde está trabajo.jsp)
            try:
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame("trabajo")
                for selector in selectores_importar:
                    elems = self.driver.find_elements(*selector)
                    for elem in elems:
                        if elem.is_displayed():
                            self.driver.execute_script("""
                                var btn = arguments[0];
                                btn.removeAttribute('disabled');
                                btn.scrollIntoView({block: 'center', inline: 'nearest'});
                            """, elem)
                            try:
                                elem.click()
                            except Exception:
                                self.driver.execute_script("arguments[0].click();", elem)
                            logging.info("Click en 'Importar SKU'realizado exitosamente (en frame 'trabajo')")
                            return True
            except Exception:
                pass

            # B) Probar en sub-iframes de país
            try:
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame("trabajo")
                sub_frames = self.driver.find_elements(By.TAG_NAME, "iframe")
                for sub in sub_frames:
                    try:
                        self.driver.switch_to.default_content()
                        self.driver.switch_to.frame("trabajo")
                        self.driver.switch_to.frame(sub)
                        for selector in selectores_importar:
                            elems = self.driver.find_elements(*selector)
                            for elem in elems:
                                if elem.is_displayed():
                                    self.driver.execute_script("""
                                        var btn = arguments[0];
                                        btn.removeAttribute('disabled');
                                        btn.scrollIntoView({block: 'center', inline: 'nearest'});
                                    """, elem)
                                    try:
                                        elem.click()
                                    except Exception:
                                        self.driver.execute_script("arguments[0].click();", elem)
                                    logging.info("Click en 'Importar SKU'realizado exitosamente (en iframe de país)")
                                    return True
                    except Exception:
                        continue
            except Exception:
                pass

            time.sleep(0.5)

        logging.error("No se encontró el botón 'Importar SKU'tras esperar 15 segundos.")
        return False

    def click_continue_and_import(self, file_path):
        """Hace click en 'Continuar', maneja alertas y luego hace click en 'Importar Sku' para subir el Excel de forma rápida y dinámica."""
        try:
            # 1. Click 'Continuar'
            self._descartar_alertas()
            if not self._hacer_click_continuar():
                logging.warning("No se encontró explícitamente el botón 'Continuar', intentando avanzar...")

            # 2. Esperar dinámicamente que finalice el procesamiento de 'Continuar'
            self._descartar_alertas()
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.invisibility_of_element_located((By.ID, "loading"))
                )
            except Exception:
                pass

            try:
                ok_btns = self.driver.find_elements(By.XPATH, "//div[contains(@class,'ui-dialog')]//button[contains(text(),'Ok') or contains(text(),'Aceptar')]")
                for btn in ok_btns:
                    if btn.is_displayed():
                        btn.click()
                        logging.info("Botón de diálogo 'Ok'/ 'Aceptar'cliqueado.")
                        break
            except Exception:
                pass

            # 3. Click 'Importar Sku' (dinámico por ID parcial o texto para F1, F2, F3, F4, etc.)
            self._descartar_alertas()
            if not self._hacer_click_importar_sku():
                raise Exception("No se encontró el botón 'Importar SKU' en la página")

            # Asegurar la obtención y validación de la ruta absoluta física completa del sistema
            full_path = os.path.abspath(str(file_path))
            if not os.path.isfile(full_path):
                raise FileNotFoundError(f"No se encontró el archivo Excel en la ruta física especificada: '{full_path}'")

            # 4 y 5. Cambiar al iframe de importación y localizar el input file con reintento activo mientras carga uploadImportF1.jsp
            file_input = None
            ruta_asignada_ok = False
            end_time = time.time() + 20
            while time.time() < end_time:
                try:
                    if self._switch_to_import_iframe():
                        inputs = self.driver.find_elements(By.CSS_SELECTOR, "input[type='file'], input[name='file']")
                        if not inputs:
                            inputs = self.driver.find_elements(By.XPATH, "//input[@type='file' or @name='file']")
                        
                        if inputs:
                            candidate = inputs[0]
                            # Intentar limpiar el input antes de asignar la ruta
                            try:
                                candidate.clear()
                            except Exception:
                                pass
                            
                            logging.info(f"Asignando la ruta absoluta física completa al input de archivo: '{full_path}'")
                            candidate.send_keys(full_path)
                            
                            # Disparar eventos change/input en el elemento tipo file
                            self.driver.execute_script("""
                                var input = arguments[0];
                                input.dispatchEvent(new Event('input', { bubbles: true }));
                                input.dispatchEvent(new Event('change', { bubbles: true }));
                                if (window.jQuery) {
                                    window.jQuery(input).trigger('input').trigger('change');
                                }
                            """, candidate)
                            
                            # Validar la presencia de la asignación en el atributo del navegador
                            val_asignado = ""
                            for _ in range(5):
                                val_asignado = candidate.get_attribute("value") or ""
                                if val_asignado:
                                    break
                                time.sleep(0.2)
                            
                            if val_asignado:
                                file_input = candidate
                                ruta_asignada_ok = True
                                logging.info(f"Confirmación: Ruta completa '{full_path}'cargada en Selenium. (Valor visual retornado por el navegador: '{val_asignado}')")
                                break
                            else:
                                logging.warning("El valor del input de archivo sigue vacío, reintentando...")
                except Exception as ex_loop:
                    logging.warning(f"Error al localizar/asignar archivo, reintentando: {ex_loop}")
                
                time.sleep(1)

            if not ruta_asignada_ok or file_input is None:
                raise NoSuchElementException("No se pudo localizar o asignar la ruta al input de archivo tipo file dentro del iframe.")
            
            time.sleep(1.5)

            # Capturar el HTML del iframe antes de enviar el formulario para verificar el procesamiento
            html_antes_submit = ""
            try:
                html_antes_submit = self.driver.page_source
            except Exception:
                pass

            # 6. Click Submit / Disparar proceso de importación
            submit_exitoso = False

            # Diagnóstico del DOM dentro del iframe de importación
            try:
                info_dom = self.driver.execute_script("""
                    var info = [];
                    var inputs = document.getElementsByTagName('input');
                    for (var i = 0; i < inputs.length; i++) {
                        info.push('INPUT type=' + inputs[i].type + ' id=' + inputs[i].id + ' name=' + inputs[i].name + ' val=' + inputs[i].value);
                    }
                    var buttons = document.getElementsByTagName('button');
                    for (var j = 0; j < buttons.length; j++) {
                        info.push('BUTTON id=' + buttons[j].id + ' text=' + buttons[j].innerText);
                    }
                    return info.join(' | ');
                """)
                logging.info(f"Elementos dentro de uploadImportF1.jsp: {info_dom}")
            except Exception:
                pass

            # 1. Clic nativo de Selenium sobre el botón Submit
            try:
                inputs = self.driver.find_elements(By.TAG_NAME, "input")
                btn_submit = None
                for inp in inputs:
                    t = (inp.get_attribute("type") or "").lower()
                    v = (inp.get_attribute("value") or "").lower()
                    if (t == "submit" or "submit" in v or "cargar" in v or "import" in v) and t != "file":
                        btn_submit = inp
                        break

                if btn_submit:
                    self.driver.execute_script("""
                        var btn = arguments[0];
                        btn.style.display = 'inline-block';
                        btn.style.visibility = 'visible';
                        btn.removeAttribute('disabled');
                        btn.scrollIntoView({block: 'center', inline: 'nearest'});
                        btn.focus();
                    """, btn_submit)
                    time.sleep(0.3)
                    btn_submit.click()
                    submit_exitoso = True
                    logging.info("Clic NATIVO de Selenium en el botón Submit realizado exitosamente")
            except Exception as ex_nativo:
                logging.warning(f"Clic nativo en submit falló ({ex_nativo}), ejecutando submit directo en #uploadForm...")

            # 2. Invocación explícita del método submit() en #uploadForm para asegurar la transmisión multipart
            try:
                self.driver.execute_script("""
                    var form = document.getElementById('uploadForm') || document.forms[0];
                    if (form) {
                        if (window.jQuery) {
                            try { window.jQuery(form).trigger('submit'); } catch(e){}
                        }
                        if (form.onsubmit) { try { form.onsubmit(); } catch(e){} }
                        try { form.submit(); } catch(e){}
                    }
                """)
                logging.info("Formulario #uploadForm enviado exitosamente (/OCEANO/importformat1?cmd=upload)")
                submit_exitoso = True
            except Exception as ex_form:
                logging.warning(f"Error al invocar submit() en #uploadForm: {ex_form}")

            # 7. Esperar a que el servidor procese la plantilla de SKUs e inserte los datos en AS400/DB2
            logging.info("Esperando que el servidor procese y confirme la carga del archivo Excel...")
            end_time = time.time() + 20
            confirmado_servidor = False
            while time.time() < end_time:
                try:
                    src = self.driver.page_source
                    if src and src != html_antes_submit and ("Insert" in src or "exito" in src.lower() or "éxito" in src.lower() or "procesad" in src.lower() or "cargad" in src.lower() or "ok" in src.lower()):
                        confirmado_servidor = True
                        logging.info("Confirmación recibida: El servidor finalizó el procesamiento del Excel.")
                        break
                except Exception:
                    pass
                time.sleep(1)

            if not confirmado_servidor:
                logging.info("Esperando tiempo de procesamiento backend adicional para la grabación en DB2/AS400...")
                time.sleep(6)

            # Pausa de visualización previa al cierre
            time.sleep(2)

            # 8. Salir del iframe de importación y volver al frame 'trabajo'
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            time.sleep(1)

            try:
                cerrar_ejecutado = self.driver.execute_script("""
                    var clicked = false;
                    // 1. Clic en el botón Cerrar oficial de la ventana jQuery UI dialogImportSkuXlsF1
                    $("[id^='dialogImportSkuXls']").each(function() {
                        var $dialogPane = $(this).closest(".ui-dialog").find(".ui-dialog-buttonpane, .ui-dialog-buttonset");
                        $dialogPane.find("button").each(function() {
                            jQuery(this).trigger("click");
                            clicked = true;
                        });
                    });

                    // 2. Fallback general si no se encontró en el diálogo específico
                    if (!clicked) {
                        $(".ui-dialog-buttonpane button, .ui-dialog-titlebar-close").each(function() {
                            jQuery(this).trigger("click");
                            clicked = true;
                        });
                    }

                    // 3. Ejecutar la función oficial cargarTabs() que dibuja los SKUs en las pestañas del país
                    if (typeof cargarTabs === 'function') {
                        try { cargarTabs(); } catch(e){}
                    }

                    // 4. Limpieza de overlays residuales
                    if ($(".ui-widget-overlay").length > 0) {
                        $(".ui-widget-overlay").remove();
                        $(".ui-dialog").hide();
                    }
                    return clicked;
                """)

                if cerrar_ejecutado:
                    logging.info("Botón 'Cerrar'cliqueado y pestañas de SKUs actualizadas (cargarTabs)")
                else:
                    logging.warning("No se detectó botón 'Cerrar', forzando recarga de pestañas con cargarTabs()...")
                    self.driver.execute_script("if (typeof cargarTabs === 'function') { cargarTabs(); }")

                logging.info("Esperando que la rejilla de SKUs se renderice en las pestañas del país...")
                time.sleep(5)
            except Exception as ex:
                logging.warning(f"Excepción al cerrar el diálogo modal: {ex}")

            # 9. Volver al contexto principal
            self.driver.switch_to.default_content()
            return True
        except Exception as e:
            logging.error(f"Error durante el proceso de importación: {e}")
            self.driver.switch_to.default_content()
            return False

    def ir_a_resumen_orden(self):
        """Hace click en el botón 'Resumen de Orden de Compra' de forma inmediata al estar disponible."""
        try:
            self._descartar_alertas()
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")

            # Remoción de overlays residuales de jQuery UI (div.ui-widget-overlay) si persistieran
            try:
                self.driver.execute_script("""
                    if (window.jQuery) {
                        try { jQuery('.ui-dialog-content').dialog('close'); } catch(e){}
                        jQuery('.ui-widget-overlay').remove();
                        jQuery('.ui-dialog').hide();
                    }
                """)
            except Exception:
                pass

            # Esperar dinámicamente que el botón esté disponible
            btn_resumen = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "orderSummary"))
            )

            self.driver.execute_script("""
                var btn = arguments[0];
                btn.removeAttribute('disabled');
                btn.scrollIntoView({block: 'center', inline: 'nearest'});
            """, btn_resumen)

            try:
                btn_resumen.click()
            except Exception:
                self.driver.execute_script("arguments[0].click();", btn_resumen)

            logging.info("Click en 'Resumen de Orden de Compra'realizado")

            # Esperar dinámicamente que se oculte el botón o cargue el resumen
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.invisibility_of_element_located((By.ID, "orderSummary"))
                )
            except Exception:
                pass

            self._descartar_alertas()
            self.driver.switch_to.default_content()
            return True
        except Exception as e:
            logging.error(f"Error al ir a Resumen de Orden de Compra: {e}")
            self._descartar_alertas()
            self.driver.switch_to.default_content()
            return False

    def enviar_a_aprobacion(self):
        """Hace click en el botón 'Enviar a Aprobación' en la página de resumen."""
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")

            btn_aprobacion = None
            selectores = [
                (By.XPATH, "//input[@value='Enviar a Aprobación']"),
                (By.XPATH, "//input[contains(@value,'Aprobaci')]"),
                (By.XPATH, "//button[contains(text(),'Aprobaci')]"),
                (By.ID, "authorization"),
                (By.ID, "enviarAprobacion"),
                (By.NAME, "authorization"),
            ]
            for selector in selectores:
                try:
                    btn_aprobacion = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable(selector)
                    )
                    break
                except Exception:
                    continue

            if btn_aprobacion is None:
                logging.warning("No se encontró el botón 'Enviar a Aprobación' en la pantalla de resumen.")
                self.driver.switch_to.default_content()
                return False

            btn_aprobacion.click()
            logging.info("Click en 'Enviar a Aprobación' (resumen) realizado")
            time.sleep(3)

            try:
                confirmado = self.driver.execute_script("""
                    var clicked = false;
                    $(".ui-dialog-buttonpane button").each(function() {
                        var txt = $(this).text().trim();
                        if (txt === 'Aceptar' || txt === 'OK' || txt === 'Confirmar' || txt === 'Sí') {
                            $(this).trigger('click');
                            clicked = true;
                        }
                    });
                    return clicked;
                """)
                if confirmado:
                    logging.info("Confirmación de aprobación aceptada")
                    time.sleep(3)
            except Exception:
                pass

            self.driver.switch_to.default_content()
            return True
        except Exception as e:
            logging.error(f"Error al enviar a aprobación: {e}")
            self.driver.switch_to.default_content()
            return False



    def obtener_numero_po(self):
        """
        Extrae el Número maestro de PO generado de la pantalla (ej. '0200744981').
        Busca dinámicamente en 'trabajo' y en los sub-iframes de país.
        """
        selectores_po = [
            (By.XPATH, "//h3[contains(.,'Número maestro de PO')]//b"),
            (By.XPATH, "//h3[contains(.,'PO')]//b"),
            (By.XPATH, "//*[contains(text(),'Número maestro de PO')]/following-sibling::b"),
            (By.XPATH, "//*[contains(text(),'Número maestro de PO')]")
        ]

        # 1. Buscar directamente en 'trabajo'
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            for sel in selectores_po:
                elems = self.driver.find_elements(*sel)
                for elem in elems:
                    txt = elem.text.strip()
                    if txt:
                        match = re.search(r'\d{6,}', txt)
                        if match:
                            po_numero = match.group(0)
                            logging.info(f"📦 Número maestro de PO extraído exitosamente: '{po_numero}'")
                            return po_numero
        except Exception:
            pass

        # 2. Buscar en sub-iframes de 'trabajo' (country0, country3, etc.)
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            sub_frames = self.driver.find_elements(By.TAG_NAME, "iframe")
            for sub in sub_frames:
                try:
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame("trabajo")
                    self.driver.switch_to.frame(sub)
                    for sel in selectores_po:
                        elems = self.driver.find_elements(*sel)
                        for elem in elems:
                            txt = elem.text.strip()
                            if txt:
                                match = re.search(r'\d{6,}', txt)
                                if match:
                                    po_numero = match.group(0)
                                    logging.info(f"📦 Número maestro de PO extraído exitosamente (en sub-frame): '{po_numero}'")
                                    return po_numero
                except Exception:
                    continue
        except Exception:
            pass

        # 3. Fallback Regex directo en el código HTML de 'trabajo'
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            html_content = self.driver.page_source
            match = re.search(r'Número maestro de PO\s*:\s*<b[^>]*>\s*(\d+)\s*</b>', html_content, re.IGNORECASE)
            if match:
                po_numero = match.group(1)
                logging.info(f"📦 Número maestro de PO extraído vía Regex: '{po_numero}'")
                return po_numero
        except Exception:
            pass

        logging.warning("No se pudo extraer el Número maestro de PO de la pantalla actual.")
        return None

    def capturar_dom_de_iframes(self, paso_nombre="captura"):

        carpeta = "dom_capturas"
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)

        try:
            self.driver.switch_to.default_content()
            main_dom = self.driver.execute_script("return document.documentElement.outerHTML;")
            main_path = os.path.join(carpeta, f"{paso_nombre}_main.html")
            with open(main_path, "w", encoding="utf-8") as f:
                f.write(main_dom)
            logging.info(f"DOM principal guardado en '{main_path}'")

            frames = self.driver.find_elements(By.TAG_NAME, "iframe") + self.driver.find_elements(By.TAG_NAME, "frame")
            for idx, frame_elem in enumerate(frames):
                frame_name = frame_elem.get_attribute("name") or frame_elem.get_attribute("id") or f"frame_{idx}"
                try:
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame(frame_elem)
                    sub_dom = self.driver.execute_script("return document.documentElement.outerHTML;")
                    file_path = os.path.join(carpeta, f"{paso_nombre}_{frame_name}.html")
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.write(sub_dom)
                    logging.info(f"DOM del frame '{frame_name}'guardado en '{file_path}'")

                    # Revisar sub-iframes dentro de este frame
                    sub_frames = self.driver.find_elements(By.TAG_NAME, "iframe") + self.driver.find_elements(By.TAG_NAME, "frame")
                    for s_idx, s_frame_elem in enumerate(sub_frames):
                        s_name = s_frame_elem.get_attribute("name") or s_frame_elem.get_attribute("id") or f"subframe_{s_idx}"
                        try:
                            self.driver.switch_to.frame(s_frame_elem)
                            nested_dom = self.driver.execute_script("return document.documentElement.outerHTML;")
                            nested_path = os.path.join(carpeta, f"{paso_nombre}_{frame_name}_{s_name}.html")
                            with open(nested_path, "w", encoding="utf-8") as f:
                                f.write(nested_dom)
                            logging.info(f"DOM del sub-frame '{s_name}'guardado en '{nested_path}'")
                            self.driver.switch_to.parent_frame()
                        except Exception as e_sub:
                            logging.warning(f"No se pudo capturar el sub-frame '{s_name}': {e_sub}")
                except Exception as e_frame:
                    logging.warning(f"No se pudo capturar el frame '{frame_name}': {e_frame}")

            self.driver.switch_to.default_content()

            # Capturar logs de la consola
            try:
                logs = self.driver.get_log("browser")
                log_path = os.path.join(carpeta, f"{paso_nombre}_logs_consola.json")
                with open(log_path, "w", encoding="utf-8") as f:
                    json.dump(logs, f, indent=2)
                logging.info(f"Logs de consola guardados en '{log_path}'")
            except Exception:
                pass

        except Exception as e:
            logging.error(f"Error al capturar el árbol DOM: {e}")

    def autorizar_po_menu(self, po_numero):
        """
        Navega al menú de autorización y autoriza la PO.
        """
        try:
            self.acceder_frame_menu()
            
            # Desplegar 'ADMINISTRAR PO' (Titulo8)
            try:
                opcion71 = self.driver.find_element(By.ID, "Titulo8Opcion71")
                if not opcion71.is_displayed():
                    WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "Titulo8"))
                    ).click()
                    logging.info("Clic en 'ADMINISTRAR PO' (Titulo8) realizado.")
                    time.sleep(1)
            except:
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "Titulo8"))
                ).click()
                logging.info("Clic en 'ADMINISTRAR PO' (Titulo8) realizado.")
                time.sleep(1)

            btn_autorizar_menu = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "Titulo8Opcion71"))
            )
            btn_autorizar_menu.click()
            logging.info("Clic en menú 'Autorizar PO' (Titulo8Opcion71) realizado.")
            time.sleep(2)
            
            # Cambiar al frame de trabajo para la búsqueda
            self.acceder_frame_trabajo()
            
            # Buscar la PO
            input_q = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.NAME, "q"))
            )
            input_q.clear()
            input_q.send_keys(po_numero)
            time.sleep(0.5)
            input_q.send_keys(Keys.RETURN)
            
            # Contingencia: si no busca con Enter, dar clic al botón de búsqueda de la lupa
            try:
                lupa = self.driver.find_element(By.CSS_SELECTOR, ".pSearch.pButton")
                lupa.click()
            except:
                pass
                
            logging.info(f"Búsqueda de PO {po_numero} para autorizar enviada.")
            
            # Esperar a que la pantalla de carga (si aparece) desaparezca
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, ".blockUI.blockOverlay"))
                )
            except:
                pass
                
            time.sleep(1)
            
            # Seleccionar el checkbox de la PO, esperando hasta 30 segundos a que renderice
            checkbox_po = WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.ID, po_numero))
            )
            checkbox_po.click()
            logging.info(f"Checkbox de la PO {po_numero} seleccionado.")
            
            # Clic en Aceptar para autorizar con espera explícita
            btn_aceptar = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "acceptButton"))
            )
            # Asegurarnos de hacer scroll para que no esté oculto
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_aceptar)
            time.sleep(1)
            try:
                btn_aceptar.click()
            except:
                self.driver.execute_script("arguments[0].click();", btn_aceptar)
            
            logging.info("Clic en 'Aceptar' (acceptButton) realizado para autorizar (con espera explícita).")
            
            # Esperar a que aparezca el diálogo jQuery UI y hacer clic en Cerrar mediante JS
            try:
                # Damos un pequeño respiro para que el diálogo termine de renderizar
                time.sleep(2)
                cerrado = self.driver.execute_script("""
                    var clicked = false;
                    $(".ui-dialog-buttonpane button").each(function() {
                        var txt = $(this).text().trim();
                        if (txt === 'Cerrar' || txt === 'Aceptar' || txt === 'OK') {
                            $(this).trigger('click');
                            clicked = true;
                        }
                    });
                    return clicked;
                """)
                if cerrado:
                    logging.info("Clic en 'Cerrar' (jQuery UI dialog) realizado con éxito mediante JS.")
                    
                    # Esperar explícitamente a que el cuadro de diálogo desaparezca
                    try:
                        WebDriverWait(self.driver, 30).until(
                            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".ui-dialog"))
                        )
                        logging.info("El cuadro de diálogo se ha cerrado por completo.")
                    except:
                        logging.warning("El cuadro de diálogo tardó demasiado en desaparecer.")
                        
                    # Esperar también a la pantalla de carga (blockUI) que se dispara después de autorizar
                    try:
                        WebDriverWait(self.driver, 60).until(
                            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".blockUI.blockOverlay"))
                        )
                        logging.info("La capa de carga posterior a la autorización ha desaparecido.")
                    except:
                        logging.warning("La capa de carga posterior a la autorización tardó demasiado.")
                        
                else:
                    logging.warning("No se encontró el botón 'Cerrar' mediante JS.")
                time.sleep(1)
            except Exception as e:
                logging.warning(f"No se pudo cerrar la alerta de autorización: {e}")
            
            return True
        except Exception as e:
            logging.error(f"Error al autorizar la PO {po_numero}: {e}")
            return False

    def cambiar_estado_po_menu(self):
        """
        Navega a la pantalla de cambio de estado.
        """
        try:
            self.acceder_frame_menu()
            
            # Desplegar menú 'ADMINISTRAR PO' si es necesario
            try:
                opcion66 = self.driver.find_element(By.ID, "Titulo8Opcion66")
                if not opcion66.is_displayed():
                    WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "Titulo8"))
                    ).click()
                    logging.info("Clic en 'ADMINISTRAR PO' (Titulo8) realizado.")
                    time.sleep(1)
            except:
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "Titulo8"))
                ).click()
                logging.info("Clic en 'ADMINISTRAR PO' (Titulo8) realizado.")
                time.sleep(1)

            btn_cambiar = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "Titulo8Opcion66"))
            )
            btn_cambiar.click()
            logging.info("Clic en menú 'Cambiar Estado PO' (Titulo8Opcion66) realizado.")
            time.sleep(2)
            return True
        except Exception as e:
            logging.error(f"Error al navegar a Cambiar Estado PO: {e}")
            return False

    def buscar_y_cambiar_estado_po(self, po_numero):
        """
        Busca la PO por su número, cambia su estado y extrae el RI.
        Retorna (exito: bool, numero_ri: str|None)
        """
        try:
            self.acceder_frame_trabajo()
            
            # 1. Ingresar número de PO
            po_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "PONumber"))
            )
            po_input.clear()
            po_input.send_keys(po_numero)
            
            # 2. Borrar el campo Country
            country_input = self.driver.find_element(By.ID, "Country")
            country_input.clear()
            
            # 3. Click en Buscar
            btn_search = self.driver.find_element(By.ID, "botonSearch")
            btn_search.click()
            logging.info(f"Búsqueda ejecutada para PO: {po_numero}")
            
            # 4. Esperar a que el radio button aparezca y extraer RI del onclick
            radio_po = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "masterPONumber"))
            )
            
            onclick_text = radio_po.get_attribute("onclick") or ""
            numero_ri = None
            match_ri = re.search(r'pori=(\d+)', onclick_text)
            if match_ri:
                numero_ri = match_ri.group(1)
                logging.info(f"📦 RI extraído del atributo onclick: {numero_ri}")
            
            # 5. Seleccionar la PO y hacer clic en Cambiar Estado PO
            radio_po.click()
            logging.info("Radio button de la PO seleccionado.")
            
            btn_copy = self.driver.find_element(By.ID, "botonCopy")
            btn_copy.click()
            logging.info("Clic en 'Cambiar Estado PO' (botonCopy) realizado.")
            time.sleep(2)
            
            # 6. Manejar la alerta/diálogo y hacer clic en 'Cambiar estado de la orden'
            try:
                # Podría haber una alerta nativa antes o ser un botón en el DOM
                try:
                    alert = self.driver.switch_to.alert
                    alert.accept()
                    logging.info("Alerta JS aceptada al cambiar estado.")
                    time.sleep(1)
                except:
                    pass
                
                # Clic en el botón final de confirmación
                btn_change = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "changePOButton"))
                )
                btn_change.click()
                logging.info("Clic en 'Cambiar estado de la orden' (changePOButton) realizado.")
                time.sleep(2)
                
            except Exception as e:
                logging.warning(f"Advertencia al confirmar el cambio de estado: {e}")
            
            return True, numero_ri
            
        except Exception as e:
            logging.error(f"Error en buscar_y_cambiar_estado_po para {po_numero}: {e}")
            return False, None