import os
import time
import traceback
import win32com.client as win32
from conexion_config.logging_config import get_logger

logger = get_logger(__file__)

class NotificadorCorreo:
    def __init__(self, destinatarios_path: str = None, flujo: str = 'default'):
        """
        Inicializa el notificador. Si no se provee la ruta del archivo de destinatarios,
        intentará buscarla en la ubicación por defecto (archivos_config/destinatarios_flujos.json).
        """
        self.flujo = flujo
        if destinatarios_path is None:
            # Asume que este archivo está en 'utilidades/', por lo que se sube un nivel para llegar a la raíz del proyecto
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            self.destinatarios_path = os.path.join(base_dir, 'archivos_config', 'destinatarios_flujos.json')
        else:
            self.destinatarios_path = destinatarios_path

    def _cargar_destinatarios(self) -> list:
        import json
        destinatarios = []
        logger.info(f"Buscando archivo de destinatarios en: {self.destinatarios_path}")
        if os.path.exists(self.destinatarios_path):
            try:
                if self.destinatarios_path.endswith('.json'):
                    with open(self.destinatarios_path, 'r', encoding='utf-8') as f:
                        config_flujos = json.load(f)
                        destinatarios = config_flujos.get(self.flujo, [])
                else:
                    with open(self.destinatarios_path, 'r', encoding='utf-8') as f:
                        for line in f:
                            for correo in line.replace(';', ',').split(','):
                                correo = correo.strip()
                                if correo and not correo.startswith('#'):
                                    destinatarios.append(correo)
                logger.info(f"Destinatarios encontrados en archivo: {destinatarios}")
            except Exception as e:
                logger.error(f"Error al leer el archivo de destinatarios: {e}")
        
        if not destinatarios:
            logger.info("Archivo de destinatarios no encontrado o vacío, usando destinatario por defecto")
            destinatarios = ["eliseo_lopezp@unicomer.com"]
            
        return destinatarios

    def _obtener_firma_outlook(self) -> str:
        """Intenta extraer la firma predeterminada de Outlook desde AppData."""
        import re
        try:
            appdata = os.environ.get('APPDATA')
            if not appdata:
                return ""
            sig_dir = os.path.join(appdata, 'Microsoft', 'Signatures')
            if not os.path.exists(sig_dir):
                return ""
            
            for f in os.listdir(sig_dir):
                if f.endswith('.htm'):
                    with open(os.path.join(sig_dir, f), 'r', encoding='utf-8', errors='ignore') as file:
                        html_content = file.read()
                        
                        # Reemplazar rutas relativas de imágenes por rutas absolutas locales
                        base_name = f.replace('.htm', '')
                        folder_name_escaped = base_name.replace(' ', '%20') + '_archivos'
                        abs_path = 'file:///' + os.path.join(sig_dir, base_name + '_archivos').replace('\\', '/').replace(' ', '%20')
                        
                        html_content = html_content.replace(folder_name_escaped, abs_path)
                        html_content = html_content.replace(base_name + '_archivos', abs_path)
                        
                        # Extraer solo el contenido dentro de <body>
                        body_match = re.search(r'<body[^>]*>(.*)</body>', html_content, re.IGNORECASE | re.DOTALL)
                        if body_match:
                            return f"<br><br>{body_match.group(1)}"
                        return f"<br><br>{html_content}"
        except Exception as e:
            logger.warning(f"No se pudo cargar la firma nativa de Outlook: {e}")
        return ""

    def enviar(self, asunto: str, cuerpo: str, adjuntos: list = None, destinatarios_explicitos: list = None) -> bool:
        """
        Envía un correo electrónico usando Outlook local vía COM de Windows.
        Se utiliza GetInspector para inyectar la firma predeterminada del usuario de forma nativa
        sin causar cierres inesperados.
        """
        import re
        import traceback
        import json

        logger.info("Iniciando proceso de envío de correo a través de NotificadorCorreo...")
        
        # Validar la bandera global de configuración de correo
        if os.path.exists(self.destinatarios_path) and self.destinatarios_path.endswith('.json'):
            try:
                with open(self.destinatarios_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if config.get('habilitar_correo', 'Y').upper() == 'N':
                        logger.info("Envío de correo deshabilitado desde configuración ('habilitar_correo' = 'N'). Omitiendo envío.")
                        return True
            except Exception as e:
                logger.error(f"Error verificando bandera habilitar_correo: {e}")

        try:
            outlook = win32.Dispatch('outlook.application')
            
            # Inicializar explícitamente el espacio de nombres MAPI
            # Usamos NewSession=False para NO bloquear el archivo .ost si el usuario
            # intenta abrir Outlook después.
            namespace = outlook.GetNamespace("MAPI")
            namespace.Logon("", "", False, False)
            
            mail = outlook.CreateItem(0)
            
            if destinatarios_explicitos:
                destinatarios = destinatarios_explicitos
            else:
                destinatarios = self._cargar_destinatarios()
                
            # Para que Outlook cargue la firma predeterminada CON todas sus imágenes
            # embebidas (cid:...), es necesario inicializar la vista del correo.
            # Usamos Display(False) porque GetInspector causa un bug al hacer Send().
            mail.Display(False)
            
            mail.To = ";".join(destinatarios)
            mail.Subject = asunto
            
            firma_nativa = mail.HTMLBody
            
            if firma_nativa and "<body" in firma_nativa.lower():
                # Limpiar el cuerpo de las etiquetas <html> y <body> para no corromper la firma
                cuerpo_limpio = re.sub(r'(?i)</?html[^>]*>', '', cuerpo)
                cuerpo_limpio = re.sub(r'(?i)</?body[^>]*>', '', cuerpo_limpio).strip()
                
                # Inyectar nuestro cuerpo limpio justo DESPUÉS de la etiqueta <body> de la firma
                cuerpo_con_firma = re.sub(r'(?i)(<body[^>]*>)', r'\g<1>' + cuerpo_limpio, firma_nativa, count=1)
                mail.HTMLBody = cuerpo_con_firma
                logger.info("Firma nativa con imágenes extraída exitosamente.")
            else:
                mail.HTMLBody = cuerpo
                logger.info("Firma nativa no encontrada, enviando cuerpo original.")
            
            # Verificar y adjuntar archivos
            archivos_adjuntados = 0
            if adjuntos:
                for adjunto in adjuntos:
                    if os.path.exists(adjunto):
                        mail.Attachments.Add(adjunto)
                        archivos_adjuntados += 1
                        logger.info(f"Adjunto agregado: {adjunto}")
                    else:
                        logger.info(f"ADVERTENCIA: El archivo adjunto no existe: {adjunto}")
            
            import time
            time.sleep(1) # Pequeña pausa para asegurar que COM termine de adjuntar
            
            logger.info(f"Enviando correo a {len(destinatarios)} destinatarios con {archivos_adjuntados} adjuntos...")
            mail.Send()
            
            # CRÍTICO: Darle tiempo al spooler de Outlook en segundo plano para 
            # sincronizar con Exchange y sacar el correo de la Bandeja de Salida
            time.sleep(4) 
            
            # Forzar explícitamente el envío/recepción en caso de que la opción 
            # "Enviar inmediatamente al estar conectado" esté desactivada en el Outlook del usuario
            try:
                for sync in namespace.SyncObjects:
                    sync.Start()
                logger.info("Sincronización manual disparada con éxito.")
            except Exception as e:
                logger.warning(f"No se pudo forzar la sincronización (normal si es headless): {e}")
            
            logger.info(f"[OK] Correo enviado exitosamente a: {', '.join(destinatarios)}")
            return True
        except Exception as e:
            logger.error(f"[ERROR] ERROR al enviar correo: {e}")
            logger.error(traceback.format_exc())
            
            # Si el correo falló y dejamos un inspector oculto abierto, 
            # forzamos su cierre para evitar bloquear el OST
            try:
                if 'mail' in locals():
                    mail.Close(1) # olDiscard = 1
            except:
                pass
                
            return False
        finally:
            # Liberación explícita de punteros COM para no dejar procesos Zombie de Outlook
            if 'mail' in locals(): del mail
            if 'namespace' in locals(): del namespace
            if 'outlook' in locals(): del outlook
