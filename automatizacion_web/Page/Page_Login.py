from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

class LoginPage:
    def __init__(self, driver):
        self.driver = driver
        self.usuario_input = (By.ID, "Userid")
        self.clave_input = (By.ID, "Password")
        self.boton_login = (By.ID, "button_accept")

    def ingresar_usuario(self, usuario):
        self.driver.find_element(*self.usuario_input).clear()
        self.driver.find_element(*self.usuario_input).send_keys(usuario)

    def ingresar_clave(self, clave):
        self.driver.find_element(*self.clave_input).clear()
        self.driver.find_element(*self.clave_input).send_keys(clave)

    def hacer_click_login(self):
        self.driver.find_element(*self.boton_login).click()

    def login(self, usuario, clave):
        self.ingresar_usuario(usuario)
        self.ingresar_clave(clave)
        self.hacer_click_login()
