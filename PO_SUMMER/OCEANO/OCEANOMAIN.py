import logging
from selenium import webdriver
from .POHEADER import POHEADERCLASS
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
import time
from datetime import datetime, timedelta


def llenar_po(driver, tipo_po="Importada"):
    proveedor = "010479"
    evento = "21"
    comprador = "710"
    bodega = "99"
    terminos_pago = "30"
    correo = "eliseo_lopezp@unicomer.com"
    pais = "EL SALVADOR"
    proveedorRwt = "006227"
 

    fecha_hoy = datetime.now()
    fecha_3dias = fecha_hoy + timedelta(days=3)
    fecha_5dias = fecha_hoy + timedelta(days=5)
    fecha_7dias = fecha_hoy + timedelta(days=7)
    fecha_10dias = fecha_hoy + timedelta(days=10)

    # Formato yyyymmdd
    fecha_embargue = fecha_hoy.strftime("%Y%m%d")
    fecha_recepcion = fecha_3dias.strftime("%Y%m%d")
    fecha_distribucion = fecha_5dias.strftime("%Y%m%d")
    fecha_cancelacion = fecha_7dias.strftime("%Y%m%d")
    fecha_para_ordenar = fecha_10dias.strftime("%Y%m")

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
            fecha_cancelacion=fecha_cancelacion,
            fecha_para_ordenar=fecha_para_ordenar,
            terminos_pago=terminos_pago,
            tipo_po=tipo_po,
            correo=correo,
            pais=pais,
            proveedorRwt=proveedorRwt
        ) 

        time.sleep(5)
        logging.info("PO ingresado exitosamente")
    except NoSuchElementException as e:
        logging.error("Error al ingresar PO: Elemento no encontrado.")
        logging.info(e)

def main_descripcion(driver):
    time.sleep(2)  # Short pause for page load stability
    try:
        po_header = POHEADERCLASS(driver)
        po_header.ingresar_descripcion_po()
    except NoSuchElementException as e:
        logging.error(f"Error al ingresar la descripción del PO: {e}")