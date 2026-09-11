import time
import logging
from selenium import webdriver
from Page.Page_Login import LoginPage
from Page.page_menu import MenuPage

def test_autorizar():
    # Configuración de logs
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    po_numero = "0200745228"
    logging.info(f"Iniciando prueba de autorización para la PO: {po_numero}")

    # Iniciar navegador
    options = webdriver.ChromeOptions()
    options.add_argument('--window-size=1920,1080')
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    try:
        # 1. Login
        driver.get("http://was7tr1.siman.com/AccesoSUMMER/")
        login_page = LoginPage(driver)
        login_page.login("ELOPEZ", "MAY2024")
        time.sleep(8)

        # 2. Navegar
        menu = MenuPage(driver)
        menu.acceder_frame_menu()
        menu.hacer_click_en_summer()
        menu.click_oceano()
        time.sleep(2)

        # 3. Intentar autorizar la PO
        logging.info("Llamando a menu.autorizar_po_menu(po_numero)...")
        menu.autorizar_po_menu(po_numero)

        time.sleep(10)
        logging.info("Prueba finalizada. Cierra la ventana cuando estés listo.")
        
    except Exception as e:
        logging.error(f"Error durante la prueba: {e}")
    
    finally:
        # driver.quit() # Comentado para que puedas ver el resultado en pantalla
        pass

if __name__ == "__main__":
    test_autorizar()
