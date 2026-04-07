from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import json

class MenuPage:
    def __init__(self, driver):
        self.driver = driver
        self.boton_summer = (By.ID, "Titulo3")  # ID del botón "SUMMER"

    def acceder_frame_menu(self):
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("menu")
            print("✅ Cambiado al frame 'menu'")
            return True
        except Exception as e:
            print(f"❌ Error al cambiar al frame 'menu': {e}")
            return False

    def hacer_click_en_summer(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(self.boton_summer)
            )
            boton = self.driver.find_element(*self.boton_summer)
            boton.click()
            print("✅ Click en 'SUMMER' realizado")
            time.sleep(5)  # Espera para observar el resultado del clic
            return True
        except TimeoutException:
            print("❌ Tiempo de espera agotado. No se encontró el botón 'SUMMER'.")
            return False
        except NoSuchElementException as e:
            print("❌ No se encontró el botón 'SUMMER'.")
            print(e)
            return False
        except Exception as e:
            print(f"❌ Error inesperado al hacer clic en 'SUMMER': {e}")
            return False


    def click_oceano(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "Titulo3Opcion10"))
            )
            boton_oceano = self.driver.find_element(By.ID, "Titulo3Opcion10")
            boton_oceano.click()
            print("✅ Click en 'OCEANO' realizado")
            time.sleep(20)  # Espera para observar el resultado del clic
            return True
        except TimeoutException:
            print("❌ Tiempo de espera agotado. No se encontró el botón 'OCEANO'.")
            return False
        except NoSuchElementException as e:
            print("❌ No se encontró el botón 'OCEANO'.")
            print(e)
            return False
        except Exception as e:
            print(f"❌ Error inesperado al hacer clic en 'OCEANO': {e}")
            return False
        
    def click_SOL(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "Titulo3Opcion9"))
            )
            boton_sol = self.driver.find_element(By.ID, "Titulo3Opcion9")
            boton_sol.click()
            print("✅ Click en 'SOL' realizado")
            time.sleep(20)  # Espera para observar el resultado del clic
            return True
        except TimeoutException:
            print("❌ Tiempo de espera agotado. No se encontró el botón 'SOL'.")
            return False
        except NoSuchElementException as e:
            print("❌ No se encontró el botón 'SOL'.")
            print(e)
            return False
        except Exception as e:
            print(f"❌ Error inesperado al hacer clic en 'SOL': {e}")
            return False
    
    def click_swim(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "Titulo3Opcion7"))
            )
            boton_terra = self.driver.find_element(By.ID, "Titulo3Opcion7")
            boton_terra.click()
            print("✅ Click en 'swim' realizado")
            time.sleep(20)  # Espera para observar el resultado del clic
            return True
        except TimeoutException:
            print("❌ Tiempo de espera agotado. No se encontró el botón 'swim'.")
            return False
        except NoSuchElementException as e:
            print("❌ No se encontró el botón 'swim'.")
            print(e)
            return False
        except Exception as e:
            print(f"❌ Error inesperado al hacer clic en 'swim': {e}")
            return False
    
    def crear_po_menu(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "Titulo1"))
            )
            boton_po = self.driver.find_element(By.ID, "Titulo1")
            boton_po.click()
            print("✅ Click en 'Crear PO' realizado")
            time.sleep(20)  # Espera para observar el resultado del clic
            return True
        except TimeoutException:
            print("❌ Tiempo de espera agotado. No se encontró el botón 'Crear PO'.")
            return False
        except NoSuchElementException as e:
            print("❌ No se encontró el botón 'Crear PO'.")
            print(e)
            return False
        except Exception as e:
            print(f"❌ Error inesperado al hacer clic en 'Crear PO': {e}")
            return False
    
    def crear_po(self):
        try:
            # Esperar hasta que el botón esté presente
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "Titulo1Opcion69"))
            )
            boton_po_venta = self.driver.find_element(By.ID, "Titulo1Opcion69")
            boton_po_venta.click()
            print("✅ Click en 'Crear PO' realizado")
            time.sleep(20)  # Espera para observar el resultado del clic
            return True
        except TimeoutException:
            print("❌ Tiempo de espera agotado. No se encontró el botón 'Crear PO Venta'.")
            return False
        except NoSuchElementException as e:
            print("❌ No se encontró el botón 'Crear PO Venta'.")
            print(e)
            return False
        except Exception as e:
            print(f"❌ Error inesperado al hacer clic en 'Crear PO Venta': {e}")
            return False
    
    def acceder_frame_trabajo(self):
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("trabajo")
            print("✅ Cambiado al frame 'trabajo'")
            return True
        except Exception as e:
            print(f"❌ Error al cambiar al frame 'trabajo': {e}")
            return False
        

    def capturar_dom_de_iframes(self):
        nombre_dom = "dom_iframe.html"
        nombre_logs = "logs_consola.json"
        
        try:
            

            # Capturar el DOM completo renderizado (como en DevTools)
            dom = self.driver.execute_script("return document.documentElement.outerHTML;")
            with open(nombre_dom, "w", encoding="utf-8") as f:
                f.write(dom)
            print(f"✅ Árbol DOM guardado en '{nombre_dom}'")

            # Capturar logs de la consola (como en DevTools → Console)
            logs = self.driver.get_log("browser")
            with open(nombre_logs, "w", encoding="utf-8") as f:
                json.dump(logs, f, indent=2)
            print(f"✅ Logs de consola guardados en '{nombre_logs}'")

        except Exception as e:
            print(f"❌ Error al guardar el árbol DOM o los logs: {e}")