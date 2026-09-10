import sys
import os
import csv
import requests
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from jira.jira_utilidades import JiraClient

class ZephyrReportGenerator:
    def __init__(self):
        self.client = JiraClient()
        self.logger = self.client.logger
        self.zephyr_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImJhc2VVcmwiOiJodHRwczovL2dydXBvdW5pY29tZXIuYXRsYXNzaWFuLm5ldCIsInVzZXIiOnsiYWNjb3VudElkIjoiNzEyMDIwOjMyN2NjNWU1LWNkYjQtNGVhYS1iY2JjLTcxMGNkZTI3ZDAyYSIsInRva2VuSWQiOiI4MjEzZTdmNS0yZjJlLTRhYTQtODI5Ni00MjkwNjY0MGRhZWEifX0sImlzcyI6ImNvbS5rYW5vYWgudGVzdC1tYW5hZ2VyIiwic3ViIjoiOWNiODYxZGYtOTE4ZC0zZjFhLTgxODYtNjNkMjkxOGMxNTkwIiwiZXhwIjoxNzk1MjkzMjc1LCJpYXQiOjE3NjM3NTcyNzV9.xmr44h6gz_gdVioh9CFE_gwm69QlY9BXQaaV7i0QJmA"
        self.headers = {
            "Authorization": f"Bearer {self.zephyr_token}",
            "Accept": "application/json"
        }

    def obtener_historias_mas(self) -> list:
        url = f"{self.client.base_url}/rest/api/3/search/jql"
        jql = 'project = "MAS" AND issuetype IN (Story, EPIC) ORDER BY created DESC'
        payload = {"jql": jql, "maxResults": 100, "fields": ["summary", "status", "issuetype", "issuelinks", "attachment"]}
        try:
            response = requests.post(url, json=payload, headers=self.client.headers_json, auth=self.client.auth, timeout=30)
            if response.status_code == 200:
                return response.json().get("issues", [])
        except Exception as e:
            self.logger.error(f"Error consultando Jira: {e}")
        return []

    def obtener_estados_ciclos(self) -> dict:
        """Mapea ID de estado a Nombre para Test Cycles."""
        url = "https://api.zephyrscale.smartbear.com/v2/statuses"
        params = {"projectKey": "MAS", "statusType": "TEST_CYCLE"}
        try:
            resp = requests.get(url, headers=self.headers, params=params, timeout=30)
            if resp.status_code == 200:
                return {v["id"]: v["name"].upper() for v in resp.json().get("values", [])}
        except:
            pass
        return {13783088: 'NOT EXECUTED', 13783089: 'IN PROGRESS', 13783090: 'DONE'}

    def obtener_todos_los_ciclos_y_enlaces(self) -> dict:
        """
        Descarga todos los ciclos del proyecto y construye un mapa:
        issue_id -> [lista de ciclos a los que pertenece]
        """
        url = "https://api.zephyrscale.smartbear.com/v2/testcycles"
        params = {"projectKey": "MAS", "maxResults": 1000}
        
        mapa_issue_ciclos = {}
        
        try:
            resp = requests.get(url, headers=self.headers, params=params, timeout=30)
            if resp.status_code == 200:
                ciclos = resp.json().get("values", [])
                
                for c in ciclos:
                    cycle_info = {
                        "id": c.get("id"),
                        "key": c.get("key"),
                        "name": c.get("name"),
                        "status_id": c.get("status", {}).get("id") if c.get("status") else None
                    }
                    
                    enlaces = c.get("links", {}).get("issues", [])
                    for link in enlaces:
                        issue_id = str(link.get("issueId"))
                        if issue_id not in mapa_issue_ciclos:
                            mapa_issue_ciclos[issue_id] = []
                        mapa_issue_ciclos[issue_id].append(cycle_info)
                        
        except Exception as e:
            self.logger.error(f"Error obteniendo ciclos: {e}")
            
        return mapa_issue_ciclos

    def generar_reporte_xlsx(self, ruta_salida: str = "reporte_hu_zephyr_desglosado.xlsx"):
        self.logger.info("Iniciando generación de reporte por Estados de Ciclo...")
        
        historias = self.obtener_historias_mas()
        mapa_estados = self.obtener_estados_ciclos()
        mapa_issue_ciclos = self.obtener_todos_los_ciclos_y_enlaces()
        
        reporte_data = []
        
        for issue in historias:
            issue_key = issue["key"]
            issue_id = str(issue["id"])
            resumen = issue["fields"]["summary"]
            tipo = issue["fields"]["issuetype"]["name"]
            
            # Buscar tarea de análisis en los links
            tiene_tarea = "NO"
            issuelinks = issue["fields"].get("issuelinks", [])
            for link in issuelinks:
                linked_issue = link.get("outwardIssue") or link.get("inwardIssue")
                if linked_issue:
                    link_summary = linked_issue.get("fields", {}).get("summary", "")
                    if link_summary.startswith("An") and "lisis y dise" in link_summary and "casos de prueba" in link_summary:
                        tiene_tarea = "SI"
                        break
            
            # Buscar archivo Signoff en attachments
            tiene_signoff = "NO"
            attachments = issue["fields"].get("attachment", [])
            for att in attachments:
                if att.get("filename", "").lower().startswith("signoff"):
                    tiene_signoff = "SI"
                    break
            
            ciclos_vinculados = mapa_issue_ciclos.get(issue_id, [])
            total_ciclos = len(ciclos_vinculados)
            
            if total_ciclos == 0:
                reporte_data.append({
                    "HU/Epic": issue_key, "Tipo": tipo, "Resumen": resumen,
                    "Tiene_Tarea_Analisis": tiene_tarea,
                    "Tiene_Archivo_Signoff": tiene_signoff,
                    "Total_Ciclos": 0, "Ciclos_DONE": 0,
                    "Porcentaje_Ciclos": "N/A", "Cumple_100%": "NO",
                    "Ciclo_ID": "N/A",
                    "Ciclo_Nombre": "Sin Ciclos",
                    "Prefijo_Ciclo": "N/A",
                    "Estado_Ciclo": "N/A"
                })
                continue
                
            ciclos_done = 0
            
            for c in ciclos_vinculados:
                status_name = mapa_estados.get(c["status_id"], "UNKNOWN")
                if status_name in ["DONE", "PASS", "PASSED"]:
                    ciclos_done += 1
                    
            porcentaje = round((ciclos_done / total_ciclos) * 100, 2)
            cumple = "SI" if porcentaje == 100.0 else "NO"
            
            # Generar una fila por cada ciclo
            for c in ciclos_vinculados:
                status_name = mapa_estados.get(c["status_id"], "UNKNOWN")
                cycle_key = c.get("key", "")
                cycle_name = c.get("name", "")
                prefijo_ciclo = cycle_name.split("-")[0].strip() if "-" in cycle_name else cycle_name
                
                reporte_data.append({
                    "HU/Epic": issue_key,
                    "Tipo": tipo,
                    "Resumen": resumen,
                    "Tiene_Tarea_Analisis": tiene_tarea,
                    "Tiene_Archivo_Signoff": tiene_signoff,
                    "Total_Ciclos": total_ciclos,
                    "Ciclos_DONE": ciclos_done,
                    "Porcentaje_Ciclos": f"{porcentaje}%",
                    "Cumple_100%": cumple,
                    "Ciclo_ID": cycle_key,
                    "Ciclo_Nombre": cycle_name,
                    "Prefijo_Ciclo": prefijo_ciclo,
                    "Estado_Ciclo": status_name
                })
            
            self.logger.info(f"{issue_key}: {ciclos_done}/{total_ciclos} ciclos en DONE.")
            
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Reporte Zephyr"
            
            fieldnames = ["HU/Epic", "Tipo", "Resumen", "Tiene_Tarea_Analisis", "Tiene_Archivo_Signoff", "Total_Ciclos", "Ciclos_DONE", "Porcentaje_Ciclos", "Cumple_100%", "Ciclo_ID", "Ciclo_Nombre", "Prefijo_Ciclo", "Estado_Ciclo"]
            sheet.append(fieldnames)
            
            for row_data in reporte_data:
                row = [row_data[field] for field in fieldnames]
                sheet.append(row)
                
            # Añadir diseño de tabla bonito
            tab = Table(displayName="TablaZephyr", ref=sheet.dimensions)
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                   showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            sheet.add_table(tab)
            
            # Ajustar automáticamente el ancho de las columnas
            for col in sheet.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                # Agregar un pequeño margen al ancho
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[col_letter].width = adjusted_width
                
            workbook.save(ruta_salida)
            self.logger.info(f"Reporte generado en: {ruta_salida}")
            print(f"Reporte generado exitosamente en: {ruta_salida}")
        except Exception as e:
            self.logger.error(f"Error escribiendo XLSX: {e}")

if __name__ == "__main__":
    generador = ZephyrReportGenerator()
    generador.generar_reporte_xlsx()
