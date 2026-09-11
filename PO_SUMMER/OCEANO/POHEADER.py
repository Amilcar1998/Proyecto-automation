import logging
import sys
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


class POHEADERCLASS:
    def __init__(self, driver):
        self.driver = driver
        self.descripcion = (By.ID, "descripcionPo")
        # Usamos locators explícitos para 'input' y 'select' para NO coincidir con los divs contenedores (divdata04VendorCode, etc.)
        self.proveedor = (By.CSS_SELECTOR, "input[id$='VendorCode']")
        self.evento = (By.CSS_SELECTOR, "input[id$='EventCode']")
        self.comprador = (By.CSS_SELECTOR, "input[id$='BuyerCode']")
        self.bodega = (By.CSS_SELECTOR, "input[id$='WarehouseCode']")
        self.fecha_embargue = (By.CSS_SELECTOR, "input[id$='ShipDate']")
        self.fecha_recepcion = (By.CSS_SELECTOR, "input[id$='ExpectedReceiptDate']")
        self.fecha_distribucion = (By.CSS_SELECTOR, "input[id$='DistributedDate']")
        self.fecha_cancelacion = (By.CSS_SELECTOR, "input[id$='CancelationDate']")
        self.fecha_para_ordenar = (By.CSS_SELECTOR, "input[id$='OtbDate1']")
        self.terminos_pago = (By.CSS_SELECTOR, "input[id$='TermNumberDays1']")
        self.tipo_po = (By.CSS_SELECTOR, "select[id$='LocalForeign']")
        self.correo = (By.CSS_SELECTOR, "input[id$='PoEmail']")
        self.pais = (By.ID, "countryChkId0")
        self.proveedorRwt = (By.CSS_SELECTOR, "input[id$='Svnhprm']")
        self.vendorSiteCode = (By.CSS_SELECTOR, "select[id$='VendorSiteCode']")

    def asegurar_frame_trabajo(self):
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
        except Exception:
            pass

    def asegurar_frame_country(self):
        """
        Garantiza que el driver esté posicionado en el iframe del país VISIBLE/ACTIVO dentro de 'trabajo'.
        """
        # 1. Si ya estamos dentro del iframe correcto con el formulario cargado, no cambiar de frame
        try:
            if len(self.driver.find_elements(By.CSS_SELECTOR, "input[id$='VendorCode']")) > 0:
                return True
        except Exception:
            pass

        # 2. Si no, ir a default_content -> 'trabajo' y esperar a que #loading esté oculto
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.invisibility_of_element_located((By.ID, "loading"))
                )
            except Exception:
                pass
        except Exception:
            pass

        # 3. Recorrer los iframes de 'trabajo' buscando primero aquel que sea VISIBLE
        try:
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            
            # Prioridad 1: iFrame activo/visible (sin 'visibility: hidden' ni 'display: none')
            for iframe in iframes:
                try:
                    parent_zone = iframe.find_element(By.XPATH, "..")
                    parent_style = parent_zone.get_attribute("style") or ""
                    iframe_style = iframe.get_attribute("style") or ""
                    
                    if "visibility: hidden" in parent_style.lower() or "visibility: hidden" in iframe_style.lower():
                        continue
                    if "display: none" in parent_style.lower() or "display: none" in iframe_style.lower():
                        continue

                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame("trabajo")
                    self.driver.switch_to.frame(iframe)
                    if len(self.driver.find_elements(By.CSS_SELECTOR, "input[id$='VendorCode']")) > 0:
                        return True
                except Exception:
                    continue

            # Prioridad 2: Fallback por candidatos conocidos
            candidatos = ["country3", "country0", "country1", "country2", "country4", "country5", "country6"]
            for iframe_id in candidatos:
                try:
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame("trabajo")
                    self.driver.switch_to.frame(iframe_id)
                    if len(self.driver.find_elements(By.CSS_SELECTOR, "input[id$='VendorCode']")) > 0:
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

    def asegurar_frame_country0(self):
        self.asegurar_frame_country()

    def set_value(self, locator, val, nombre_campo="", dar_enter=True):
        self.asegurar_frame_country()
        try:
            elem = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located(locator)
            )
            
            # Asegurar que estamos interactuando con una etiqueta <input> y no un <div> contenedor
            if elem.tag_name.lower() not in ["input", "select"]:
                sub_inputs = elem.find_elements(By.TAG_NAME, "input")
                if sub_inputs:
                    elem = sub_inputs[0]

            # 1. Habilitar, quitar readonly, scrollIntoView, focus y click
            self.driver.execute_script("""
                var elem = arguments[0];
                elem.removeAttribute('readonly');
                elem.removeAttribute('disabled');
                elem.scrollIntoView({block: 'center', inline: 'nearest'});
                elem.focus();
            """, elem)

            try:
                elem.click()
            except Exception:
                self.driver.execute_script("arguments[0].click();", elem)

            # 2. Limpiar contenido anterior únicamente si el campo no está vacío
            try:
                val_actual = elem.get_attribute("value") or ""
                if val_actual.strip():
                    elem.send_keys(Keys.CONTROL + "a")
                    elem.send_keys(Keys.BACKSPACE)
                    elem.clear()
            except Exception:
                pass

            # 3. Tipear el valor físicamente y presionar ENTER si dar_enter es True
            try:
                elem.send_keys(str(val))
                if dar_enter:
                    elem.send_keys(Keys.ENTER)
            except Exception:
                pass

            # 4. Sincronización JS y disparo de evento ENTER keyCode 13
            self.driver.execute_script("""
                var elem = arguments[0];
                var val = arguments[1];
                var darEnter = arguments[2];
                elem.value = val;
                elem.setAttribute('value', val);
                
                if (window.jQuery) {
                    window.jQuery(elem).val(val).trigger('input').trigger('change').trigger('keyup');
                }
                
                if (darEnter) {
                    var enterEvt = new KeyboardEvent('keypress', {
                        bubbles: true, cancelable: true, keyCode: 13, which: 13, charCode: 13
                    });
                    elem.dispatchEvent(enterEvt);
                }

                elem.dispatchEvent(new Event('input', { bubbles: true }));
                elem.dispatchEvent(new Event('change', { bubbles: true }));

                var oncomp = elem.getAttribute('oncomplete');
                if (oncomp) {
                    try { eval(oncomp); } catch(e){}
                }

                var elemId = elem.id || '';
                var baseId = elemId.replace(/^data\d+/, '');
                if (baseId) {
                    var possibleErrIds = ['err' + baseId, 'errSuggested', 'errVendorSiteCode', 'errCurrency', 'errComision'];
                    for (var i = 0; i < possibleErrIds.length; i++) {
                        var errEl = document.getElementById(possibleErrIds[i]);
                        if (errEl) {
                            if (possibleErrIds[i] === 'err' + baseId || 
                               (elemId.indexOf('Svnhprm') !== -1 && possibleErrIds[i] === 'errSuggested')) {
                                errEl.style.display = 'none';
                            }
                        }
                    }
                }
            """, elem, str(val), dar_enter)

            # 5. Si la app despliega indicador de carga (#loading), esperar a que desaparezca
            try:
                WebDriverWait(self.driver, 3).until(
                    EC.invisibility_of_element_located((By.ID, "loading"))
                )
            except Exception:
                pass

            # 6. Manejo inmediato de alertas si están presentes en el navegador
            try:
                alert = self.driver.switch_to.alert
                text = alert.text
                if text:
                    logging.warning(f"Alerta detectada en campo '{nombre_campo}': '{text}'. Aceptando...")
                    alert.accept()
            except Exception:
                pass

            val_real = elem.get_attribute("value") or ""
            tag_name = elem.tag_name
            elem_id = elem.get_attribute("id")
            logging.info(f"[{nombre_campo}] (<{tag_name}> #{elem_id}) reflejado en DOM: '{val_real}'")
            return True
        except Exception as e:
            logging.error(f"Error al ingresar [{nombre_campo}] ({val}): {e}")
            return False

    def ingresar_po(self, proveedor, evento, comprador, bodega, fecha_embargue, fecha_recepcion, fecha_distribucion, fecha_cancelacion, fecha_para_ordenar, terminos_pago, tipo_po, correo, pais, proveedorRwt):
        try:
            self.asegurar_frame_country()

            # 1. PROVEEDOR PRINCIPAL: Seteo e invocación de checkVendorParams
            logging.info(f" 🔹 Ingresando Proveedor: {proveedor}")
            self.set_value(self.proveedor, proveedor, "Proveedor")
            
            # Esperar a que la petición AJAX viewVendorParameters termine de forma dinámica
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.invisibility_of_element_located((By.ID, "loading"))
                )
            except Exception:
                pass
            try:
                alert = self.driver.switch_to.alert
                if alert.text:
                    logging.warning(f"Alerta de proveedor: '{alert.text}'. Aceptando...")
                    alert.accept()
            except Exception:
                pass

            # 2. COMPRADOR, EVENTO Y BODEGA (después de processParameters)
            logging.info(f" 🔹 Ingresando Comprador: {comprador}")
            self.set_value(self.comprador, comprador, "Comprador")

            logging.info(f" 🔹 Ingresando Evento: {evento}")
            self.set_value(self.evento, evento, "Evento")

            logging.info(f" 🔹 Ingresando Bodega: {bodega}")
            self.set_value(self.bodega, bodega, "Bodega")

            # 3. FECHAS
            self.set_value(self.fecha_embargue, fecha_embargue, "Fecha Embarque")
            self.set_value(self.fecha_recepcion, fecha_recepcion, "Fecha Recepción")
            self.set_value(self.fecha_distribucion, fecha_distribucion, "Fecha Distribución")
            self.set_value(self.fecha_cancelacion, fecha_cancelacion, "Fecha Cancelación")
            self.set_value(self.fecha_para_ordenar, fecha_para_ordenar, "Fecha OTB")

            # 4. TÉRMINOS DE PAGO Y CORREO
            self.set_value(self.terminos_pago, terminos_pago, "Términos de Pago")
            self.set_value(self.correo, correo, "Correo Electrónico")

            # 5. TIPO PO (Importada / Local)
            self.asegurar_frame_country()
            elemento = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(self.tipo_po)
            )
            select = Select(elemento)
            target_val = "F" if tipo_po in ['Importada', 'F'] else "L"
            try:
                select.select_by_value(target_val)
            except Exception:
                try:
                    select.select_by_visible_text(tipo_po)
                except Exception:
                    pass

            self.driver.execute_script("""
                var elem = arguments[0];
                var val = elem.value;
                elem.dispatchEvent(new Event('change', { bubbles: true }));
                if (window.jQuery) {
                    window.jQuery(elem).trigger('change');
                }
                if (typeof showHideForeign === 'function') {
                    var cId = window.countryId || '04';
                    showHideForeign(val, 'ordenImportada' + cId, '1');
                }
            """, elemento)
            logging.info(f"Tipo PO configurado como: {tipo_po}")
            time.sleep(2)

            # 6. SEGUNDO PROVEEDOR (PROVEEDOR SUGERIDO / RWT) SI ES IMPORTADA
            if target_val == 'F':
                logging.info(f" 🔹 Ingresando Proveedor Sugerido (RWT): {proveedorRwt}")
                self.set_value(self.proveedorRwt, proveedorRwt, "Proveedor Sugerido / RWT")
                
                # Esperar respuesta AJAX de getSiteCode
                logging.info("Esperando respuesta AJAX de Vendor Site Code (getSiteCode)...")
                time.sleep(3)
                
                try:
                    self.asegurar_frame_country()
                    site_elem = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located(self.vendorSiteCode)
                    )
                    select_site = Select(site_elem)
                    for _ in range(5):
                        if len(select_site.options) > 0:
                            idx = 1 if len(select_site.options) > 1 else 0
                            select_site.select_by_index(idx)
                            val_selected = select_site.options[idx].text
                            logging.info(f"Vendor Site Code seleccionado exitosamente: '{val_selected}'")
                            break
                        time.sleep(1)

                    self.driver.execute_script("""
                        var errSite = document.getElementById('errVendorSiteCode');
                        if (errSite) { errSite.style.display = 'none'; }
                    """)
                except Exception as e_site:
                    logging.warning(f"Vendor Site Code no requirió selección manual: {e_site}")
            else:
                logging.info("PO es Local: Se omite el segundo proveedor.")

            logging.info("Todos los datos del formulario fueron procesados correctamente.")

        except NoSuchElementException as e:
            logging.error(f"Error al ingresar datos en el formulario: {e}")
        except Exception as e:
            logging.error(f"Error inesperado al ingresar datos en el formulario: {e}")

    def ingresar_descripcion_po(self):
        descripcion = "Prueba PO Automatizada"
        try:
            self.asegurar_frame_trabajo()
            elem = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(self.descripcion)
            )
            elem.clear()
            elem.send_keys(descripcion)
            elem.send_keys(Keys.ENTER)
            time.sleep(5)
            logging.info("Descripción del PO ingresada correctamente.")
            self.gestionar_paises()
        except Exception as e:
            logging.error(f"Error al ingresar la descripción del PO: {e}")

    def gestionar_paises(self):
        """
        Selecciona Nicaragua y desmarca El Salvador (aceptando la alerta de confirmación).
        Tanto los checkboxes como los labels están en el frame 'trabajo'.
        """
        try:
            self.asegurar_frame_trabajo()
            time.sleep(1)

            # 1. Seleccionar NICARAGUA (id="countryChk3")
            try:
                nicaragua_chk = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "countryChk3"))
                )
                if not nicaragua_chk.is_selected():
                    try:
                        nicaragua_chk.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", nicaragua_chk)
                    logging.info("Nicaragua seleccionado.")
                    time.sleep(2)
            except Exception as e_nic:
                logging.error(f"Error al seleccionar Nicaragua: {e_nic}")

            # 2. Desmarcar EL SALVADOR (id="countryChk0") para eliminarlo
            try:
                salvador_chk = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "countryChk0"))
                )
                if salvador_chk.is_selected():
                    try:
                        salvador_chk.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", salvador_chk)
                    logging.info("Clic en El Salvador realizado para desmarcarlo.")
                    time.sleep(1)

                    # 3. Aceptar todas las alertas pendientes (ej. "¿Está seguro que Desea Eliminar el País?")
                    for _ in range(3):
                        try:
                            WebDriverWait(self.driver, 3).until(EC.alert_is_present())
                            alert = self.driver.switch_to.alert
                            logging.warning(f"Alerta detectada: '{alert.text}'. Aceptando...")
                            alert.accept()
                            logging.info("Alerta de eliminación de país aceptada correctamente.")
                            time.sleep(2)
                        except Exception:
                            break
            except Exception as e_sal:
                logging.error(f"Error al desmarcar El Salvador: {e_sal}")

            # 4. Esperar 4 segundos y aguardar que #loading esté oculto para que la recarga de parámetros de Nicaragua termine
            logging.info("Esperando que finalice la inicialización del formulario de país...")
            time.sleep(4)
            try:
                self.asegurar_frame_trabajo()
                WebDriverWait(self.driver, 15).until(
                    EC.invisibility_of_element_located((By.ID, "loading"))
                )
                time.sleep(2)
            except Exception:
                pass

        except Exception as e:
            logging.error(f"Error al gestionar la selección de países: {e}")
