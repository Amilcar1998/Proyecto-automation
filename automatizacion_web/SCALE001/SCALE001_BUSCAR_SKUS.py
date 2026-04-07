from selenium import webdriver
import sys
import os
import time

# Agregar el directorio raíz al path para importaciones
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from Page.Page_Login import LoginPage
from Page.page_menu import MenuPage
from SCALE001.SCALE001_SKU import SCALE001SKUClass, ejecutar_busqueda_sku_scale001

def buscar_skus_individuales(driver, lista_skus):
    """
    Busca múltiples SKUs uno por uno
    
    Args:
        driver: WebDriver de Selenium
        lista_skus: Lista de números de SKU a buscar
    """
    try:
        print("\n🚀 Iniciando automatización SCALE001 - Búsqueda de SKUs")
        print("="*70)
        
        # Login
        print("\n📝 Realizando login...")
        driver.get("http://was7tr1.siman.com/AccesoSUMMER/")
        login_page = LoginPage(driver)
        login_page.login("ELOPEZ", "NOV2025")
        print("✅ Login exitoso")
        
        time.sleep(10)
        
        # Navegación al menú
        print("\n📂 Navegando al menú SUMMER → SCALE001...")
        menu = MenuPage(driver)
        
        if not menu.acceder_frame_menu():
            raise Exception("No se pudo acceder al frame del menú")
        
        if not menu.hacer_click_en_summer():
            raise Exception("No se pudo hacer click en SUMMER")
        
        if not menu.click_SCALE001():
            raise Exception("No se pudo hacer click en SCALE001")
        
        print("✅ Navegación completada")
        
        # Acceder a Mantenimiento de SKU
        print("\n🔎 Accediendo a Mantenimiento de SKU...")
        if not ejecutar_busqueda_sku_scale001(driver):
            raise Exception("Error al acceder a Mantenimiento de SKU")
        
        print("✅ Acceso a Mantenimiento de SKU exitoso")
        
        # Crear instancia de SCALE001SKUClass para usar los métodos de búsqueda
        scale001_sku = SCALE001SKUClass(driver)
        
        # Procesar la lista de SKUs
        resultados = scale001_sku.procesar_lista_skus(lista_skus, intervalo_segundos=3)
        
        print("\n✅ Proceso completado!")
        print("="*70)
        
        return resultados
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        print("="*70)
        return None


def buscar_sku_individual(driver, numero_sku):
    """
    Busca un solo SKU (útil para llamadas individuales)
    
    Args:
        driver: WebDriver de Selenium
        numero_sku: Número de SKU a buscar
    """
    try:
        print("\n🚀 Iniciando automatización SCALE001 - Búsqueda de SKU")
        print("="*70)
        
        # Login
        print("\n📝 Realizando login...")
        driver.get("http://was7tr1.siman.com/AccesoSUMMER/")
        login_page = LoginPage(driver)
        login_page.login("ELOPEZ", "NOV2025")
        print("✅ Login exitoso")
        
        time.sleep(10)
        
        # Navegación al menú
        print("\n📂 Navegando al menú SUMMER → SCALE001...")
        menu = MenuPage(driver)
        
        if not menu.acceder_frame_menu():
            raise Exception("No se pudo acceder al frame del menú")
        
        if not menu.hacer_click_en_summer():
            raise Exception("No se pudo hacer click en SUMMER")
        
        if not menu.click_SCALE001():
            raise Exception("No se pudo hacer click en SCALE001")
        
        print("✅ Navegación completada")
        
        # Acceder a Mantenimiento de SKU
        print("\n🔎 Accediendo a Mantenimiento de SKU...")
        if not ejecutar_busqueda_sku_scale001(driver):
            raise Exception("Error al acceder a Mantenimiento de SKU")
        
        print("✅ Acceso a Mantenimiento de SKU exitoso")
        
        # Buscar el SKU
        scale001_sku = SCALE001SKUClass(driver)
        resultado = scale001_sku.buscar_sku_por_numero(numero_sku)
        
        if resultado:
            print("\n✅ Búsqueda completada exitosamente!")
        else:
            print("\n❌ La búsqueda falló")
        
        print("="*70)
        
        return resultado
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        print("="*70)
        return False


if __name__ == "__main__":
    print("""
    ╔══════════════════════════════════════════════════════╗
    ║     AUTOMATIZACIÓN SCALE001 - BÚSQUEDA DE SKUs       ║
    ╚══════════════════════════════════════════════════════╝
    """)
    
    # EJEMPLO 1: Buscar un solo SKU
    # driver = webdriver.Chrome()
    # driver.maximize_window()
    # buscar_sku_individual(driver, "123456")
    # time.sleep(30)
    # driver.quit()
    
    # EJEMPLO 2: Buscar múltiples SKUs de una lista
    driver = webdriver.Chrome()
    driver.maximize_window()
    
    # COLOCA AQUÍ TU LISTA DE SKUs
    lista_skus = [
        "121212",
        "274867",
        "274279",
        "177256",
        "288775",
        "350816",
        "350268"
    ]
    
    try:
        resultados = buscar_skus_individuales(driver, lista_skus)
        
        print("\n⏳ Navegador abierto por 60 segundos...")
        time.sleep(60)
        
    finally:
        driver.quit()
        print("\n👋 Fin del script")
