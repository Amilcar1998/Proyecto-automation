import os
import sys

# Añadir el directorio raíz al path de Python para permitir importaciones
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from jira.jira_utilidades import JiraClient
from conexion_config.logging_config import get_logger

logger = get_logger(__name__)

def crear_tareas_pendientes():
    """
    Script de prueba para crear tareas en Jira de órdenes que ya fueron procesadas 
    y tienen su reporte Word generado, pero que no se enviaron a Jira.
    """
    print("=== INICIANDO CREACIÓN DE TAREAS PENDIENTES EN JIRA ===")
    
    # 1. Definir las órdenes y la ruta donde están los reportes de hoy
    ordenes_pendientes = ["0000016392", "0000016436"]
    fecha_reporte = "20260818" # Fecha de la carpeta (YYYYMMDD)
    
    # Ruta base de los reportes
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    carpeta_reportes = os.path.join(base_dir, "Reportes", "Transferencias_WMS_INFOR")
    
    # Instanciar el cliente de Jira
    jira = JiraClient()
    tipo = "transferencia"
    
    # Procesar cada orden
    for orden in ordenes_pendientes:
        print(f"\nProcesando orden: {orden}")
        
        # Construir el nombre del archivo Word esperado
        nombre_archivo = f"Reporte_Orden_{orden}_{fecha_reporte}.docx"
        reporte_path = os.path.join(carpeta_reportes, nombre_archivo)
        
        # Verificar si el archivo existe
        if os.path.exists(reporte_path):
            print(f"[OK] Reporte encontrado: {reporte_path}")
            print(f"-> Enviando a Jira...")
            try:
                # Enviar a Jira
                jira.main_jira(orden, reporte_path, tipo, "RI14DB")
                print(f"[ÉXITO] Tarea de Jira solicitada para la orden {orden}.")
            except Exception as e:
                print(f"[ERROR] Fallo al crear la tarea en Jira para la orden {orden}: {e}")
        else:
            print(f"[ADVERTENCIA] No se encontró el reporte en la ruta: {reporte_path}")

if __name__ == "__main__":
    crear_tareas_pendientes()
    print("\n=== PROCESO FINALIZADO ===")
