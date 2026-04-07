import os
from conexion_config.logging_config import get_logger
import logging
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime


class JiraClient:
    def __init__(self):
        # --- Configuración Jira ---
        #self.JIRA_URL = "https://grupounicomer.atlassian.net"
        #self.JIRA_USER = "eliseo_lopezp@unicomer.com"
        #self.JIRA_TOKEN ="ATATT3xFfGF09x0MveqTYaSYdETahg1I_iNdxbH3dlezMDNw_ArwztPjz9TfpryJZymCHoShD2tb7n9XxmCFzkkzgrugmT-DUgS1MdKmsrNCA816PcXoOXesNASerlrFTuKTrYJD5J-GtINpJadlzK171wfiiA-Q5MjI52rqRpBkAv6u27BhvQg=D90BF94F"


        if not self.JIRA_USER or not self.JIRA_TOKEN:
            raise ValueError("Credenciales Jira no configuradas")

        self.auth = HTTPBasicAuth(self.JIRA_USER, self.JIRA_TOKEN)
        self.headers_json = {
            "Accept": "application/json",
            "Content-Type": "application/json"
        }

        self.logger = self._configurar_logger()
    # --------------------------------------------------
    # Logger
    # --------------------------------------------------
    def _configurar_logger(self):
        return get_logger(__file__)
    # --------------------------------------------------
    # Subir evidencia
    # --------------------------------------------------
    def subir_evidencia(self, issue_key: str, ruta_archivo: str) -> bool:
        if not os.path.exists(ruta_archivo):
            self.logger.error(f"No existe el archivo: {ruta_archivo}")
            return False

        url = f"{self.JIRA_URL}/rest/api/3/issue/{issue_key}/attachments"
        headers = {"X-Atlassian-Token": "no-check"}

        try:
            with open(ruta_archivo, "rb") as archivo:
                files = {"file": (os.path.basename(ruta_archivo), archivo)}

                response = requests.post(
                    url=url,
                    headers=headers,
                    files=files,
                    auth=self.auth,
                    timeout=60
                )

            if response.status_code == 200:
                self.logger.info(f"Evidencia subida correctamente a {issue_key}")
                return True

            self.logger.error(
                f"Error al subir evidencia (HTTP {response.status_code}): {response.text}"
            )
            return False

        except Exception as e:
            self.logger.exception(f"Excepción al subir evidencia: {e}")
            return False
    # --------------------------------------------------
    # Cambiar estado (In Progress → Done)
    # --------------------------------------------------
    def cambiar_estado_a_done(self, issue_key: str) -> bool:
        estados_objetivo = ["In Progress", "Done"]

        try:
            while True:
                url = f"{self.JIRA_URL}/rest/api/3/issue/{issue_key}/transitions"
                response = requests.get(
                    url,
                    headers=self.headers_json,
                    auth=self.auth,
                    timeout=30
                )

                if response.status_code != 200:
                    self.logger.error(
                        f"No se pudieron obtener transiciones (HTTP {response.status_code})"
                    )
                    return False

                transiciones = response.json().get("transitions", [])
                nombres = [t["name"] for t in transiciones]

                self.logger.info(f"Transiciones disponibles: {nombres}")

                siguiente = None
                for estado in estados_objetivo:
                    if estado in nombres:
                        siguiente = estado
                        break

                if not siguiente:
                    self.logger.error(
                        f"No se puede avanzar el ticket {issue_key}"
                    )
                    return False

                transition_id = next(
                    t["id"] for t in transiciones if t["name"] == siguiente
                )

                payload = {"transition": {"id": transition_id}}

                ejecutar = requests.post(
                    url,
                    json=payload,
                    headers=self.headers_json,
                    auth=self.auth,
                    timeout=30
                )

                if ejecutar.status_code != 204:
                    self.logger.error(
                        f"Error al cambiar estado (HTTP {ejecutar.status_code}): {ejecutar.text}"
                    )
                    return False

                self.logger.info(
                    f"Ticket {issue_key} movido a '{siguiente}'"
                )

                if siguiente == "Done":
                    self.logger.info(f"Ticket {issue_key} finalizado")
                    return True

        except Exception as e:
            self.logger.exception(f"Excepción al cambiar estado: {e}")
            return False
    # --------------------------------------------------
    # Enviar comentario
    # --------------------------------------------------
    def enviar_comentario(self, issue_key: str,tipo) -> bool:
        if tipo =="transferencia":
            comentario ="""
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
            comentario ="""
            1.	Consulta de datos en tabla KRWSPWASIN con el fin de extraer todos los load id existentes
            2.	Por cada load id encontrado se realizará lo siguiente:
            3.	Considerar que el DTAARA debe de tener una hora de cierre menor a la hora de ejecución.
            4.	Ejecución del procedimiento ELOPEZ.PROCESAR_PO, este procedimiento toma una captura del inventario inicial y lo guarda en RI12DB.INV_INICIAL así como tambien deja disponible solo el ASN que se va ha procesar.
            5.	Ejecución del comando SIWINTRACL para el posteo de transferencias en INFOR
            6.	Consulta de tablas de inventario (INV_INICIAL, KRWSPWASIN, INV_VINAL) cuando finaliza. el programa CL, borra el registro de la tabla KRWSPWASIN y lo pasa a la KRWSPWINRC.
            7.	Segunda ejecución del procedimiento ELOPEZ.PROCESAR_PO, Esto con el fin de guardar el movimiento del inventario final. el mismo procedimiento lo toma y lo guarda en INV_FINAL. para tener un respaldo de la prueba.
            8.	Validación de consistencia de datos y generación de reportes
            """


        payload = {
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

        try:
            url = f"{self.JIRA_URL}/rest/api/3/issue/{issue_key}/comment"
            response = requests.post(
                url,
                json=payload,
                headers=self.headers_json,
                auth=self.auth,
                timeout=30
            )

            if response.status_code == 201:
                self.logger.info(f"Comentario agregado a {issue_key}")
                return True

            self.logger.error(
                f"Error al comentar (HTTP {response.status_code}): {response.text}"
            )
            return False

        except Exception as e:
            self.logger.exception(f"Excepción al enviar comentario: {e}")
            return False
    # --------------------------------------------------
    # Enviar comentario
    # --------------------------------------------------
    def crear_tarea_jira(self,loadid,tipo):
        if tipo =="transferencia":
            payload = {
                        "fields": {
                            "project": {
                                "key": "WG"
                            },
                            "summary": f"Validación de posteo de load id {loadid}",
                            "issuetype": {
                                "name": "Tarea"
                            },
                            "assignee": {
                                "id": "712020:327cc5e5-cdb4-4eaa-bcbc-710cde27d02a"
                            },
                            "customfield_10095": {
                                "id": "15241"
                            },
                            "customfield_10051": [
                                {
                                    "id": "10088"
                                }
                            ],
                            "description": {
                                "type": "doc",
                                "version": 1,
                                "content": [
                                    {
                                        "type": "paragraph",
                                        "content": [
                                            {
                                                "text": f"Validación de Posteo {loadid}",
                                                "type": "text"
                                            }
                                        ]
                                    }
                                ]
                            },
                            "priority": {
                                "id": "3"
                            }
                        }
                    }
        else:
            payload = {
                        "fields": {
                            "project": {
                                "key": "WG"
                            },
                            "summary": f"Validación del ASN:  {loadid}",
                            "issuetype": {
                                "name": "Tarea"
                            },
                            "assignee": {
                                "id": "712020:327cc5e5-cdb4-4eaa-bcbc-710cde27d02a"
                            },
                            "customfield_10095": {
                                "id": "15241"
                            },
                            "customfield_10051": [
                                {
                                    "id": "10088"
                                }
                            ],
                            "description": {
                                "type": "doc",
                                "version": 1,
                                "content": [
                                    {
                                        "type": "paragraph",
                                        "content": [
                                            {
                                                "text": f"Validación del ASN {loadid}",
                                                "type": "text"
                                            }
                                        ]
                                    }
                                ]
                            },
                            "priority": {
                                "id": "3"
                            }
                        }
                    }



        try:
            response = requests.post(
                f"{self.JIRA_URL}/rest/api/3/issue",
                json=payload,
                headers=self.headers_json,
                auth=self.auth,
                timeout=30
            )

            logging.info(f"Status Code: {response.status_code}")
            logging.info(f"Response Body: {response.text}")

            if response.status_code == 201:
                key = response.json().get("key")
                logging.info(f"Clave generada correctamente: {key}")
                return key
            else:
                logging.error(
                    f"Error al crear issue Jira - "
                    f"Status: {response.status_code} - "
                    f"Response: {response.text}"
                )
                return None

        except requests.exceptions.RequestException as e:
            logging.error(f"Error de comunicación con Jira: {str(e)}")
            return None

        except Exception as e:
            logging.error(f"Excepción inesperada: {str(e)}", exc_info=True)
            return None
    
    def main_jira(self, load,ruta_archivo,tipo):
        logging.info("iniciando a generar la Tarea")
        key = self.crear_tarea_jira(load,tipo)
        logging.info(f"tarea creada correctamente: {key}")

        self.subir_evidencia(key, ruta_archivo)
        self.enviar_comentario(key,tipo)
        self.cambiar_estado_a_done(key)


    
        