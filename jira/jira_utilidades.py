import os
import requests
from requests.auth import HTTPBasicAuth
from conexion_config.logging_config import get_logger


class JiraConfig:
    """Configuración para la conexión a Jira."""
    # Podrían ser cargadas desde variables de entorno (os.environ) para mayor seguridad.
    URL = os.environ.get("JIRA_URL", "https://grupounicomer.atlassian.net")
    USER = os.environ.get("JIRA_USER", "eliseo_lopezp@unicomer.com")
    TOKEN = os.environ.get("JIRA_TOKEN", "ATATT3xFfGF0AVjOUbXoEUUHF5WGZqYlOaPDMCFUXeoExn2772LdgxwCFFWzHrB3e8ONkqOmZ26-Kd6n9UniwxpVRecOwZzVSSvLGOEx_hba73_NT25iBmPfkou3uXoSncd8MAercwOlPEOoqrKqWUsgJZueewd-K9fLIspWbVZFNvil16jstiA=05557851")


class JiraPayloadBuilder:
    """Clase responsable de construir los payloads para las peticiones a Jira."""

    @staticmethod
    def construir_comentario(tipo: str) -> dict:
        if tipo == "transferencia":
            comentario = """
            1.	Consulta de datos en tabla KTCHPWASIN con el fin de extraer todos los load id existentes
            2.	Por cada load id encontrado se realizará lo siguiente:
            3.	Considerar que el DTAARA debe de tener una hora de cierre menor a la hora de ejecución.
            4.	Ejecución del procedimiento ELOPEZ.PROCESAR_ORDENES, este procedimiento toma una captura del inventario inicial y lo guarda en RI12DB.INVENTARIO_INICIAL así como tambien deja disponible solo el load id que se va ha procesar.
            5.	Ejecución del comando SIWINTRACL para el posteo de transferencias en INFOR
            6.	Consulta de tablas de inventario (INVENTARIO_INICIAL, KTCHPWINRC, INVENTARIO_FINAL) cuando finaliza. el programa CL, borra el registro de la tabla KTCHPWASIN y lo pasa a la KTCHPWINRC.
            7.	Segunda ejecución del procedimiento PROCESAR_ORDENES, Esto con el fin de guardar el movimiento del inventario final. el mismo procedimiento lo toma y lo guarda en INVENTARIO_FINAL. para tener un respaldo de la prueba.
            9.	Validación de consistencia de datos y generación de reportes
            """
        else:
            comentario = """
            1.	Consulta de datos en tabla KRWSPWASIN con el fin de extraer todos los load id existentes
            2.	Por cada load id encontrado se realizará lo siguiente:
            3.	Considerar que el DTAARA debe de tener una hora de cierre menor a la hora de ejecución.
            4.	Ejecución del procedimiento ELOPEZ.PROCESAR_PO, este procedimiento toma una captura del inventario inicial y lo guarda en RI12DB.INV_INICIAL así como tambien deja disponible solo el ASN que se va ha procesar.
            5.	Ejecución del comando SIWINTRACL para el posteo de transferencias en INFOR
            6.	Consulta de tablas de inventario (INV_INICIAL, KRWSPWASIN, INV_VINAL) cuando finaliza. el programa CL, borra el registro de la tabla KRWSPWASIN y lo pasa a la KRWSPWINRC.
            7.	Segunda ejecución del procedimiento ELOPEZ.PROCESAR_PO, Esto con el fin de guardar el movimiento del inventario final. el mismo procedimiento lo toma y lo guarda en INV_FINAL. para tener un respaldo de la prueba.
            8.	Validación de consistencia de datos y generación de reportes
            """

        return {
            "body": {
                "type": "doc",
                "version": 1,
                "content": [
                    {
                        "type": "paragraph",
                        "content": [
                            {"type": "text", "text": comentario}
                        ]
                    }
                ]
            }
        }

    @staticmethod
    def construir_tarea(loadid: str, tipo: str, base: str, datos_tabla: list = None) -> dict:
        if tipo == "transferencia":
            summary_prefix = "Validación de posteo de load id"
            description_title = f"Validación de Posteo {loadid}"
        elif tipo == "transferencia_directa":
            summary_prefix = "Generación de transferencias"
            description_title = "Solicitud de Generación de Transferencias"
        elif tipo == "vendor_storage":
            summary_prefix = "Conciliación PO Vendor Storage:"
            description_title = f"Validación de Conciliación {loadid}"
        elif tipo == "po_summer_report":
            summary_prefix = "Solicitud de creación de Órdenes de compra WMS INFOR:"
            description_title = f"Solicitud de creación de Órdenes de compra WMS INFOR {loadid}"
        else:
            summary_prefix = "Validación del ASN:"
            description_title = f"Validación del ASN {loadid}"
        
        if tipo == "transferencia":
            pasos_texto = [
                "Consulta de datos en tabla KTCHPWASIN con el fin de extraer todos los load id existentes.",
                "Por cada load id encontrado se realizará lo siguiente:",
                "Considerar que el DTAARA debe de tener una hora de cierre menor a la hora de ejecución.",
                f"Ejecución del procedimiento ELOPEZ.PROCESAR_ORDENES, este procedimiento toma una captura del inventario inicial y lo guarda en {base}.INVENTARIO_INICIAL así como también deja disponible solo el load id que se va a procesar.",
                "Ejecución del comando SIWINTRACL para el posteo de transferencias en INFOR.",
                f"Consulta de tablas de inventario ({base}.INVENTARIO_INICIAL, KTCHPWINRC, {base}.INVENTARIO_FINAL). Cuando finaliza el programa CL, borra el registro de la tabla KTCHPWASIN y lo pasa a la KTCHPWINRC.",
                f"Segunda ejecución del procedimiento PROCESAR_ORDENES. Esto con el fin de guardar el movimiento del inventario final. El mismo procedimiento lo toma y lo guarda en {base}.INVENTARIO_FINAL para tener un respaldo de la prueba.",
                "Validación de consistencia de datos y generación de reportes."
            ]
        elif tipo == "vendor_storage":
            pasos_texto = [
                "Extracción de datos desde la tabla KTRHP (unidades ordenadas a generar).",
                "Extracción de datos desde la tabla KPUHP (unidades reportadas como generadas).",
                "Agrupación de totales por SKU y consolidación matemática.",
                "Validación de diferencias: Diferencia = [Suma KTRHP] - [Suma KPUHP].",
                "Generación de reporte ejecutivo documentando los resultados de la conciliación para la fecha indicada."
            ]
        elif tipo == "po_summer_report":
            pasos_texto = [] # No usaremos pasos texto para este reporte
        else:
            pasos_texto = [
                "Consulta de datos en tabla KRWSPWASIN con el fin de extraer todos los load id existentes.",
                "Por cada load id encontrado se realizará lo siguiente:",
                "Considerar que el DTAARA debe de tener una hora de cierre menor a la hora de ejecución.",
                f"Ejecución del procedimiento ELOPEZ.PROCESAR_PO, este procedimiento toma una captura del inventario inicial y lo guarda en {base}.INV_INICIAL así como también deja disponible solo el ASN que se va a procesar.",
                "Ejecución del comando SIWINTRACL para el posteo de transferencias en INFOR.",
                f"Consulta de tablas de inventario ({base}.INV_INICIAL, KRWSPWASIN, {base}.INV_FINAL). Cuando finaliza el programa CL, borra el registro de la tabla KRWSPWASIN y lo pasa a la KRWSPWINRC.",
                f"Segunda ejecución del procedimiento ELOPEZ.PROCESAR_PO. Esto con el fin de guardar el movimiento del inventario final. El mismo procedimiento lo toma y lo guarda en {base}.INV_FINAL para tener un respaldo de la prueba.",
                "Validación de consistencia de datos y generación de reportes."
            ]

        if tipo == "transferencia_directa":
            description_adf = [
                {
                    "type": "panel",
                    "attrs": {"panelType": "info"},
                    "content": [
                        {
                            "type": "paragraph",
                            "content": [{"type": "text", "text": description_title, "marks": [{"type": "strong"}]}]
                        }
                    ]
                },
                {
                    "type": "paragraph",
                    "content": [{"type": "text", "text": "Se atendió la solicitud que contempla la generación de las siguientes transferencias:"}]
                }
            ]
        elif tipo == "po_summer_report":
            description_adf = [
                {
                    "type": "panel",
                    "attrs": {"panelType": "info"},
                    "content": [
                        {
                            "type": "paragraph",
                            "content": [{"type": "text", "text": description_title, "marks": [{"type": "strong"}]}]
                        }
                    ]
                },
                {
                    "type": "paragraph",
                    "content": [{"type": "text", "text": "Se generaron las siguientes ordenes de compra relacionadas con el proyecto WMS Infor:"}]
                }
            ]
        else:
            # Construir Filas de la Tabla
            table_rows = []
            
            # Fila de encabezado
            table_rows.append({
                "type": "tableRow",
                "content": [
                    {
                        "type": "tableHeader",
                        "content": [{"type": "paragraph", "content": [{"type": "text", "text": "Paso", "marks": [{"type": "strong"}]}]}]
                    },
                    {
                        "type": "tableHeader",
                        "content": [{"type": "paragraph", "content": [{"type": "text", "text": "Descripción de la Acción", "marks": [{"type": "strong"}]}]}]
                    }
                ]
            })
            
            # Filas de datos
            for idx, paso in enumerate(pasos_texto, 1):
                table_rows.append({
                    "type": "tableRow",
                    "content": [
                        {
                            "type": "tableCell",
                            "content": [{"type": "paragraph", "content": [{"type": "text", "text": str(idx)}]}]
                        },
                        {
                            "type": "tableCell",
                            "content": [{"type": "paragraph", "content": [{"type": "text", "text": paso}]}]
                        }
                    ]
                })

            description_adf = [
                {
                    "type": "panel",
                    "attrs": {"panelType": "info"},
                    "content": [
                        {
                            "type": "paragraph",
                            "content": [{"type": "text", "text": description_title, "marks": [{"type": "strong"}]}]
                        }
                    ]
                },
                {
                    "type": "table",
                    "attrs": {"isNumberColumnEnabled": False, "layout": "default"},
                    "content": table_rows
                }
            ]

        if datos_tabla and len(datos_tabla) > 0:
            custom_table_rows = []
            
            # Encabezado (primera fila de datos_tabla)
            header_row = {
                "type": "tableRow",
                "content": [
                    {
                        "type": "tableHeader",
                        "content": [{"type": "paragraph", "content": [{"type": "text", "text": str(col), "marks": [{"type": "strong"}]}]}]
                    } for col in datos_tabla[0]
                ]
            }
            custom_table_rows.append(header_row)
            
            # Filas de datos
            for row in datos_tabla[1:]:
                data_row = {
                    "type": "tableRow",
                    "content": [
                        {
                            "type": "tableCell",
                            "content": [{"type": "paragraph", "content": [{"type": "text", "text": str(cell) if cell is not None else ""}]}]
                        } for cell in row
                    ]
                }
                custom_table_rows.append(data_row)
            
            
            # Título de la tabla de datos adicionales (solo si no es transferencia_directa, ya que pusimos el texto antes)
            if tipo != "transferencia_directa":
                description_adf.append({
                    "type": "heading",
                    "attrs": {"level": 3},
                    "content": [{"type": "text", "text": "Datos del Documento"}]
                })
            
            # Añadimos la tabla al ADF
            description_adf.append(
                {
                    "type": "table",
                    "attrs": {"isNumberColumnEnabled": False, "layout": "default"},
                    "content": custom_table_rows
                }
            )

        return {
            "fields": {
                "project": {
                    "key": "ML"
                },
                "summary": f"{summary_prefix} {loadid}",
                "issuetype": {
                    "name": "Tarea"
                },
                "assignee": {
                    "id": "712020:327cc5e5-cdb4-4eaa-bcbc-710cde27d02a"
                },
                "customfield_10095": {
                    "id": "14456"
                },
                "customfield_10051": [
                    {
                        "id": "10088"
                    }
                ],
                "description": {
                    "type": "doc",
                    "version": 1,
                    "content": description_adf
                },
                "priority": {
                    "id": "3"
                }
            }
        }


class JiraClient:
    """Cliente encargado exclusivamente de la comunicación con la API de Jira."""

    def __init__(self):
        if not JiraConfig.USER or not JiraConfig.TOKEN:
            raise ValueError("Credenciales Jira no configuradas")

        self.base_url = JiraConfig.URL
        self.auth = HTTPBasicAuth(JiraConfig.USER, JiraConfig.TOKEN)
        self.headers_json = {
            "Accept": "application/json",
            "Content-Type": "application/json"
        }
        self.logger = get_logger(__file__)

    def subir_evidencia(self, issue_key: str, ruta_archivo: str) -> bool:
        if not os.path.exists(ruta_archivo):
            self.logger.error(f"No existe el archivo: {ruta_archivo}")
            return False

        url = f"{self.base_url}/rest/api/3/issue/{issue_key}/attachments"
        headers = {"X-Atlassian-Token": "no-check"}

        try:
            with open(ruta_archivo, "rb") as archivo:
                files = {"file": (os.path.basename(ruta_archivo), archivo)}
                response = requests.post(url=url, headers=headers, files=files, auth=self.auth, timeout=60)

            if response.status_code == 200:
                self.logger.info(f"Evidencia subida correctamente a {issue_key}")
                return True

            self.logger.error(f"Error al subir evidencia (HTTP {response.status_code}): {response.text}")
            return False

        except Exception as e:
            self.logger.exception(f"Excepción al subir evidencia: {e}")
            return False

    def cambiar_estado_a_done(self, issue_key: str) -> bool:
        estados_objetivo = ["In Progress", "Done"]

        try:
            while True:
                url = f"{self.base_url}/rest/api/3/issue/{issue_key}/transitions"
                response = requests.get(url, headers=self.headers_json, auth=self.auth, timeout=30)

                if response.status_code != 200:
                    self.logger.error(f"No se pudieron obtener transiciones (HTTP {response.status_code})")
                    return False

                transiciones = response.json().get("transitions", [])
                nombres = [t["name"] for t in transiciones]

                self.logger.info(f"Transiciones disponibles: {nombres}")

                siguiente = next((estado for estado in estados_objetivo if estado in nombres), None)

                if not siguiente:
                    self.logger.error(f"No se puede avanzar el ticket {issue_key}")
                    return False

                transition_id = next(t["id"] for t in transiciones if t["name"] == siguiente)
                payload = {"transition": {"id": transition_id}}

                ejecutar = requests.post(url, json=payload, headers=self.headers_json, auth=self.auth, timeout=30)

                if ejecutar.status_code != 204:
                    self.logger.error(f"Error al cambiar estado (HTTP {ejecutar.status_code}): {ejecutar.text}")
                    return False

                self.logger.info(f"Ticket {issue_key} movido a '{siguiente}'")

                if siguiente == "Done":
                    self.logger.info(f"Ticket {issue_key} finalizado")
                    return True

        except Exception as e:
            self.logger.exception(f"Excepción al cambiar estado: {e}")
            return False

    def enviar_comentario(self, issue_key: str, tipo: str) -> bool:
        payload = JiraPayloadBuilder.construir_comentario(tipo)
        url = f"{self.base_url}/rest/api/3/issue/{issue_key}/comment"

        try:
            response = requests.post(url, json=payload, headers=self.headers_json, auth=self.auth, timeout=30)

            if response.status_code == 201:
                self.logger.info(f"Comentario agregado a {issue_key}")
                return True

            self.logger.error(f"Error al comentar (HTTP {response.status_code}): {response.text}")
            return False

        except Exception as e:
            self.logger.exception(f"Excepción al enviar comentario: {e}")
            return False

    def crear_tarea_jira(self, loadid: str, tipo: str, base: str, datos_tabla: list = None) -> str:
        payload = JiraPayloadBuilder.construir_tarea(loadid, tipo, base, datos_tabla)
        url = f"{self.base_url}/rest/api/3/issue"

        try:
            response = requests.post(url, json=payload, headers=self.headers_json, auth=self.auth, timeout=30)

            self.logger.info(f"Status Code: {response.status_code}")
            
            if response.status_code == 201:
                key = response.json().get("key")
                self.logger.info(f"Clave generada correctamente: {key}")
                return key
                
            self.logger.error(f"Error al crear issue Jira - Status: {response.status_code} - Response: {response.text}")
            return None

        except requests.exceptions.RequestException as e:
            self.logger.error(f"Error de comunicación con Jira: {str(e)}")
            return None
        except Exception as e:
            self.logger.error(f"Excepción inesperada: {str(e)}", exc_info=True)
            return None

    def _is_jira_enabled(self) -> bool:
        import json
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'archivos_config', 'destinatarios_flujos.json')
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get('habilitar_jira', 'Y').upper() == 'Y'
        except Exception as e:
            self.logger.warning(f"No se pudo leer la configuración de Jira, asumiendo habilitado: {e}")
        return True

    def main_jira(self, load: str, ruta_archivo: str, tipo: str, base: str, datos_tabla: list = None):
        if not self._is_jira_enabled():
            self.logger.info("El procesamiento de Jira está deshabilitado en la configuración (habilitar_jira != 'Y'). Saltando creación de tarea.")
            return None

        self.logger.info(f"Iniciando a generar la Tarea en Jira para la base {base}")
        key = self.crear_tarea_jira(load, tipo, base, datos_tabla)
        
        if key:
            self.logger.info(f"Tarea creada correctamente: {key}")
            self.subir_evidencia(key, ruta_archivo)
            # self.enviar_comentario(key, tipo) # Removido a petición para incluir todo en la descripción
            self.cambiar_estado_a_done(key)
            return key
        else:
            self.logger.error("Se interrumpió el flujo porque falló la creación de la tarea en Jira.")
            return None


    
        