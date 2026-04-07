from selenium import webdriver
from Page.Page_Login import LoginPage
from Page.page_menu import MenuPage
import time
import sys
import os

# Agregar el directorio raíz al path para importaciones
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def test_login_menu_scale001():
    """
    Función para ejecutar el login y acceder al menú SCALE001
    """
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get("http://was7tr1.siman.com/AccesoSUMMER/")

    # Login
    login_page = LoginPage(driver)
    login_page.login("ELOPEZ", "Amilcar2025*")
    print("✅ Login exitoso")
    
    time.sleep(5)

    # Navegación al menú SCALE001
    menu = MenuPage(driver)
    menu.acceder_frame_menu()
    menu.hacer_click_en_summer()
    menu.click_SCALE001()
    
    print("✅ Navegación a SCALE001 completada")
    
    # Esperar para observar resultado
    time.sleep(30)
    
    # Descomentar la siguiente línea cuando termines de probar
    # driver.quit()
    
    return driver

if __name__ == "__main__":
    test_login_menu_scale001()
    print("✅ Proceso completado. Puedes continuar con el mantenimiento en SCALE001.")
