from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

class GoogleSearch:
    def buscar(self, termino):
        self.driver.get("http://was7tr1.siman.com/AccesoSUMMER/")
        #caja = self.driver.find_element(By.NAME, "q")
        #caja.send_keys(termino)
        #caja.send_keys(Keys.RETURN)
        time.sleep(20)
        

    def cerrar(self):
        self.driver.quit()
