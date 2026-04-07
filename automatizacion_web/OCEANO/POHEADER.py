from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
import time
import pyautogui
import calendar
from datetime import datetime

class POHEADERCLASS:
    def __init__(self, driver):
        self.driver = driver
        self.descripcion = (By.ID, "descripcionPo")
        self.proveedor = (By.ID, "data01VendorCode")
        self.evento = (By.ID, "data01EventCode") 
        self.comprador = (By.ID, "data01BuyerCode")
        self.bodega = (By.ID, "data01WarehouseCode")
        self.fecha_embargue = (By.ID, "data01ShipDate")
        self.fecha_recepcion = (By.ID, "data01ExpectedReceiptDate")
        self.fecha_distribucion = (By.ID, "data01DistributedDate")
        self.fecha_para_ordenar = (By.ID, "data01OtbDate1")
        self.terminos_pago = (By.ID, "data01TermNumberDays1")
        self.tipo_po = (By.ID, "data01LocalForeign")
        self.correo = (By.ID, "data01PoEmail")
        self.pais = (By.ID, "countryChkId0")
        self.proveedorRwt = (By.ID,'data01Svnhprm')

    def ingresar_po(self, proveedor, evento, comprador, bodega, fecha_embargue, fecha_recepcion, fecha_distribucion, fecha_para_ordenar, terminos_pago, tipo_po, correo, pais, proveedorRwt):
        try:
            #if pais:
            #    checkbox = self.driver.find_element(*self.pais)
            #    if not checkbox.is_selected():
            #        checkbox.click()

            self.driver.find_element(*self.proveedor).clear()
            self.driver.find_element(*self.proveedor).send_keys(proveedor)
            self.driver.find_element(*self.evento).clear()
            self.driver.find_element(*self.evento).send_keys(evento)
            self.driver.find_element(*self.comprador).clear()
            self.driver.find_element(*self.comprador).send_keys(comprador)
            self.driver.find_element(*self.bodega).clear()
            self.driver.find_element(*self.bodega).send_keys(bodega)
            self.driver.find_element(*self.fecha_embargue).clear()
            self.driver.find_element(*self.fecha_embargue).send_keys(fecha_embargue)
            self.driver.find_element(*self.fecha_recepcion).clear()
            self.driver.find_element(*self.fecha_recepcion).send_keys(fecha_recepcion)
            self.driver.find_element(*self.fecha_distribucion).clear()
            self.driver.find_element(*self.fecha_distribucion).send_keys(fecha_distribucion)
            self.driver.find_element(*self.fecha_para_ordenar).clear()
            self.driver.find_element(*self.fecha_para_ordenar).send_keys(fecha_para_ordenar)
            self.driver.find_element(*self.terminos_pago).clear()
            self.driver.find_element(*self.terminos_pago).send_keys(terminos_pago)
            self.driver.find_element(*self.correo).clear()
            self.driver.find_element(*self.correo).send_keys(correo)
            self.driver.find_element(*self.proveedorRwt).clear()
            

            elemento = self.driver.find_element(*self.tipo_po)
            select = Select(elemento)
            select.select_by_visible_text(tipo_po)  # o select_by_value(tipo_po)

            if elemento =='Importada':
                self.driver.find_element(*self.proveedorRwt).send_keys(proveedorRwt)

            time.sleep(20)  # Espera para observar el resultado de la acción
            print("✅ Datos ingresados en el formulario correctamente.")

        except NoSuchElementException as e:
            print(f"❌ Error al ingresar datos en el formulario: {e}")

    def ingresar_descripcion_po(self):
        descripcion = "Prueba PO Automatizada"
        try:
            self.driver.find_element(*self.descripcion).clear()
            self.driver.find_element(*self.descripcion).send_keys(descripcion)
            print("✅ Descripción del PO ingresada correctamente.")
        except NoSuchElementException as e:
            print(f"❌ Error al ingresar la descripción del PO: {e}")

    def dar_tab(cantidad):
        for _ in range(cantidad):
                pyautogui.press("tab")
                time.sleep(0.5)  # opcional para que sea visible y no se trabe



    def digitar_header_PO(self):
            #Se trabaja con pyautoqgui para interactuar con la interfaz gráfica ya que se tiene una limitación con frames  que se genera por javascript
            #y selenium no las reconoce por en el arbol DOM. por eso se utiliza pyoutogui

            pyautogui.click(x=400, y=180)
            
            for _ in range(14):
                pyautogui.press("tab")
                time.sleep(0.1)  # opcional para que sea visible y no se trabe

            pyautogui.write("010479-REGAL WORLDWIDE TRADING LLC")
            #pyautogui.press("enter")
            time.sleep(1)

            # Ir al campo "Evento"
            pyautogui.press("tab")
            pyautogui.press("tab")
            pyautogui.write("710 - COMPRADOR SUMMER")
            #pyautogui.press("enter")
            time.sleep(1)

            pyautogui.press("tab")
            pyautogui.press("tab")
            pyautogui.write("21 - AUTOMATIC")
            #pyautogui.press("enter")
            time.sleep(1)

            pyautogui.press("tab")
            pyautogui.press("tab")
            pyautogui.write("94 - CDD NEJAPA")
            pyautogui.press("enter")
            time.sleep(1)
            # Ir a "Fecha de Embarque"
            pyautogui.press("tab")
            pyautogui.press("tab")

            pyautogui.write("20251022")

            time.sleep(40)
        


