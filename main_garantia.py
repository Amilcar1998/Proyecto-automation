"""
Punto de entrada independiente para el flujo de garantías.
"""

import os
import sys
import logging
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Asegurar que el directorio raíz del proyecto esté en el PYTHONPATH
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from flujo_completo import flujo_completo
from conexion_config.logging_config import get_logger

# Crear carpeta logs y archivo de log para garantías
log_dir = os.path.join(os.path.dirname(__file__), 'logs')
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"garantias_{datetime.now().strftime('%Y%m%d')}.txt")

root_logger = logging.getLogger()
if not any(isinstance(h, logging.FileHandler) for h in root_logger.handlers):
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setFormatter(logging.Formatter("%(asctime)s - %(name)-25s - %(levelname)s - %(message)s"))
    root_logger.addHandler(fh)

logger = get_logger(__file__)

def main():
    print("\n" + "="*60)
    print("SISTEMA DE GARANTÍAS AS400")
    print("="*60)
    
    logger.info("Iniciando flujo completo de garantías desde script independiente...")
    try:
        flujo_completo(logger=logger)
        logger.info("Flujo de garantías finalizado exitosamente.")
    except Exception as e:
        logger.exception(f"Error durante la ejecución del flujo de garantías: {e}")
        
    print("="*60)

if __name__ == "__main__":
    # Crear directorio logs si no existe para evitar errores en logger
    if not os.path.exists('logs'):
        os.makedirs('logs')
    main()
