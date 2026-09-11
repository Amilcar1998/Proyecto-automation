import sys
import os

# Asegurar que la ruta base del proyecto está en sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from utilidades.notificaciones_correo import NotificadorCorreo
from datetime import datetime

def enviar_pendientes():
    fecha = datetime.now().strftime('%Y%m%d')
    asunto = f"WMS-INFOR | Resumen de Validación de Transferencias - {fecha} (Reenvío)"
    
    cuerpo = f"""
    <html>
    <body style="font-family: Arial, sans-serif; font-size: 14px;">
        <h3 style="color:#2E74B5;">Resumen de Ejecución</h3>
        <p>Se han procesado <b>2</b> órdenes.</p>
        
        <table style="width:100%; border-collapse: collapse; border: 1px solid #ccc;">
            <tr style="background-color: #f2f2f2; text-align: left;">
                <th style="padding:8px; border-bottom:2px solid #ccc;">Orden</th>
                <th style="padding:8px; border-bottom:2px solid #ccc;">Estado</th>
                <th style="padding:8px; border-bottom:2px solid #ccc;">Observaciones</th>
            </tr>
            <tr>
                <td style="padding:5px; border-bottom:1px solid #ddd;">0000016441</td>
                <td style="padding:5px; border-bottom:1px solid #ddd; font-weight:bold; color:#2E7D32;">OK</td>
                <td style="padding:5px; border-bottom:1px solid #ddd; font-size:12px;">Sin descuadres</td>
            </tr>
            <tr>
                <td style="padding:5px; border-bottom:1px solid #ddd;">0000016443</td>
                <td style="padding:5px; border-bottom:1px solid #ddd; font-weight:bold; color:#2E7D32;">OK</td>
                <td style="padding:5px; border-bottom:1px solid #ddd; font-size:12px;">Sin descuadres</td>
            </tr>
        </table>
        
        <p style="margin-top:15px; color:#666; font-size:12px;">
            * Se adjuntan los reportes detallados de cada orden.
        </p>
    </body>
    </html>
    """
    
    # Rutas de los reportes generados
    base_dir = r"C:\Users\eliseo_lopezp\proyecto1\Reportes\Transferencias_WMS_INFOR_20260818"
    adjuntos = [
        os.path.join(base_dir, "Reporte_Orden_0000016441_20260818.docx"),
        os.path.join(base_dir, "Reporte_Orden_0000016443_20260818.docx")
    ]
    
    # Filtrar solo los que existen (por seguridad)
    adjuntos_validos = [f for f in adjuntos if os.path.exists(f)]
    
    print(f"Archivos a adjuntar: {adjuntos_validos}")
    
    if not adjuntos_validos:
        print("No se encontraron los archivos físicos para enviar.")
        return

    notificador = NotificadorCorreo()
    enviado = notificador.enviar(asunto, cuerpo, adjuntos=adjuntos_validos)
    
    if enviado:
        print("¡Correo reenviado exitosamente!")
    else:
        print("Fallo al reenviar el correo.")

if __name__ == "__main__":
    enviar_pendientes()
