from selenium import webdriver
import sys
import os
import time

# Agregar el directorio raíz al path para importaciones
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from Page.Page_Login import LoginPage
from Page.page_menu import MenuPage
from SOL.SOL_SKU import ejecutar_busqueda_sku_sol

def ejecutar_mantenimiento_sol_sku():
    """
    Función principal que ejecuta el flujo completo:
    1. Login
    2. Navegar al menú SUMMER → SOL
    3. Buscar y hacer click en SKU
    """
    driver = None
    
    try:
        print("\n🚀 Iniciando automatización SOL - SKU")
        print("="*70)
        
        driver = webdriver.Chrome()
        driver.maximize_window()
        
        print("\n📝 Realizando login...")
        driver.get("http://was7tr1.siman.com/AccesoSUMMER/")
        
        login_page = LoginPage(driver)
        login_page.login("ELOPEZ", "NOV2025")
        print("✅ Login exitoso")
        
        time.sleep(10)
        
        print("\n📂 Navegando al menú SUMMER → SOL...")
        menu = MenuPage(driver)
        
        if not menu.acceder_frame_menu():
            raise Exception("No se pudo acceder al frame del menú")
        
        if not menu.hacer_click_en_summer():
            raise Exception("No se pudo hacer click en SUMMER")
        
        if not menu.click_SOL():
            raise Exception("No se pudo hacer click en SOL")
        
        print("✅ Navegación completada")
        
        print("\n🔎 Accediendo a Mantenimiento de SKU...")
        
        if not ejecutar_busqueda_sku_sol(driver):
            raise Exception("Error al acceder a Mantenimiento de SKU")
        
        print("\n✅ Proceso completado!")
        print("="*70)
        print("\n⏳ Navegador abierto por 60 segundos...")
        
        time.sleep(60)
        
        return True
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        print("="*70)
        
        if driver:
            print("\n⏳ Navegador abierto 30 segundos para inspección...")
            time.sleep(30)
        
        return False
        
    finally:
        pass


if __name__ == "__main__":
    print("""
    ╔══════════════════════════════════════════════════════╗
    ║     AUTOMATIZACIÓN SOL - BÚSQUEDA SKU                ║
    ╚══════════════════════════════════════════════════════╝
    """)
    
    time.sleep(2)
    ejecutar_mantenimiento_sol_sku()
    
    print("\n👋 Fin del script")
