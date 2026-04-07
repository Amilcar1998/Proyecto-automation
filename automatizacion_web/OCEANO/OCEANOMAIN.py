from selenium import webdriver
from .POHEADER import POHEADERCLASS
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
import time
from datetime import datetime, timedelta


def llenar_po(driver):
    proveedor="010479"
    evento="21"
    comprador="710"
    bodega="94"
    terminos_pago="30"
    tipo_po="Importada"
    correo="eliseo_lopezp@unicomer.com"
    pais="EL SALVADOR"  # Cambia a False si no quieres marcar el checkbox
    proveedorRwt="006227"
 

    fecha_hoy = datetime.now()
    fecha_3dias = fecha_hoy + timedelta(days=3)
    fecha_5dias = fecha_hoy + timedelta(days=5)
    fecha_10dias = fecha_hoy + timedelta(days=10)

    # Formato yyyymmdd
    fecha_embargue=fecha_hoy.strftime("%Y%m%d")
    fecha_recepcion = fecha_3dias.strftime("%Y%m%d")
    fecha_distribucion=fecha_5dias.strftime("%Y%m%d")
    fecha_para_ordenar= fecha_10dias.strftime("%Y%m%d")


    try:
            
            po_header = POHEADERCLASS(driver)
            po_header.ingresar_po(
                proveedor=proveedor,
                evento=evento,
                comprador=comprador,
                bodega=bodega,
                fecha_embargue=fecha_embargue,
                fecha_recepcion=fecha_recepcion,
                fecha_distribucion=fecha_distribucion,
                fecha_para_ordenar=fecha_para_ordenar,
                terminos_pago=terminos_pago,
                tipo_po=tipo_po,
                correo=correo,
                pais=pais,
                proveedorRwt=proveedorRwt
            ) 

            time.sleep(400)

            print("✅ PO ingresado exitosamente")
    except NoSuchElementException as e:
        print("❌ Error al ingresar PO: Elemento no encontrado.")
        print(e)

def main_descripcion(driver):
    time.sleep(20)  # Espera para asegurarte de que la página esté completamente cargada
    descripcion = "Prueba PO Automatizada"
    try:
        po_header = POHEADERCLASS(driver)
        po_header.ingresar_descripcion_po()
        print("✅ Descripción del PO ingresada correctamente.")
    except NoSuchElementException as e:
        print(f"❌ Error al ingresar la descripción del PO: {e}")


def digitar_header_PO(driver):
    try:
        po_header = POHEADERCLASS(driver)
        po_header.digitar_header_PO()  # Ajusta el número de tabulaciones según sea necesario
        print("✅ Header PO digitado correctamente.")
    except NoSuchElementException as e:
        print(f"❌ Error al digitar el header del PO: {e}")