from selenium import webdriver
from Page.Page_Login import LoginPage
from Page.page_menu import MenuPage
import time
from OCEANO.OCEANOMAIN import *  # Asegúrate de que este import sea correcto
import pyautogui

def test_login_exitoso():
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get("http://was7tr1.siman.com/AccesoSUMMER/")  # ← URL real del login

    login_page = LoginPage(driver)
    login_page.login("ELOPEZ", "NOV2025*")

    # Espera para observar resultado (en la práctica usarías WebDriverWait)
    time.sleep(10)


    menu = MenuPage(driver)
    menu.acceder_frame_menu()

    menu.hacer_click_en_summer()
    menu.click_oceano()
    menu.acceder_frame_menu()
    menu.crear_po_menu()
    menu.crear_po()
    menu.acceder_frame_trabajo()

    main_descripcion(driver)  # Llama a la función para ingresar la descripción del PO

    digitar_header_PO(driver)  # Llama a la función para digitar el header del PO

    time.sleep(60)  # Espera para observar el resultado de la acción
    

    #driver.quit()


    
if __name__ == "__main__":
    test_login_exitoso()
#    print("Prueba de login exitosa.") esto se puede ser clase ?