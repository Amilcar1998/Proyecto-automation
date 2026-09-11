# Automatización de Creación y Autorización de POs (PO_SUMMER)

Este proyecto contiene un orquestador completo de automatización web (Selenium) e integración de datos, diseñado para agilizar el flujo de creación de Órdenes de Compra (POs) en el sistema SUMMER/OCEANO, interactuar con DB2 (AS/400) y reportar evidencias automáticamente a Jira.

## 🚀 Arquitectura y Flujo de Trabajo

El proyecto está diseñado con mecanismos anti-fallos y persistencia en tiempo real, garantizando que si ocurre un problema de conexión, la data no se pierda. Se divide en los siguientes flujos principales:

### 1. `Compras.py` (Procesamiento Inicial)
* **Transformación de Datos:** Lee el archivo maestro de Excel y agrupa la información basándose en SKU, Tienda y Cantidades.
* **Separación de Plantillas:** Clasifica las POs en locales e importadas, generando múltiples plantillas `.xls` en `ARCHIVOS/LOCAL/` y `ARCHIVOS/IMPORTADA/`.
* **Persistencia Inteligente y Regla de Negocio:** Genera el archivo `po_report.json` como base de datos de control. Implementa la regla de negocio principal para asegurar que **todas las órdenes Importadas tengan máxima prioridad**, posicionándolas siempre al inicio de la cola.

### 2. `Main.py` (El Orquestador Web y Backend)
Es el núcleo de la automatización que procesa secuencialmente cada archivo listado en el JSON:

* **Fase 0: Preparación de Entorno**
  * Elimina sesiones colgadas o "fantasmas" en el servidor ejecutando silenciosamente el procedimiento DB2 `CALL ELOPEZ.dsession('ELOPEZ')`.
  * Inicia Selenium y activa llamadas a la API de Windows para evitar la suspensión de pantalla.

* **Fase 1: Creación y Extracción Segura**
  * Sube los archivos Excel de manera fluida sorteando cortinas de carga (jQuery `blockUI` y modales `.ui-dialog`) inyectando clics forzosos por JavaScript en caso de que la UI congele el botón de cerrar.
  * Captura el **Número Maestro de PO** desde el DOM y realiza un guardado de emergencia del JSON para evitar pérdidas de trabajo ante caídas inminentes.
  
* **Fase 2: Autorización Blindada**
  * Envía la PO a aprobación y la autoriza, esperando rigurosamente los tiempos largos de renderizado (hasta 60 segundos) de las tablas Ajax Flexigrid, evitando colisiones o clics interceptados.
  
* **Fase 3: Ejecución DB2 (AS/400)**
  * Se conecta vía ODBC al servidor remoto DSN `RI_TEST`.
  * Ejecuta el procedimiento almacenado `CALL ELOPEZ.RETACEO()` para que el negocio procese el retaceo inmediatamente en la base de datos.

* **Fase 4: Extracción de RI**
  * Vuelve al front-end, consulta el número de PO extraído, captura el código de recibo interno `RI` de OCEANO.
  * Sella el estado en el JSON como `"Completado"`.

### 3. Reportes y Cierre de Tareas (Jira)
Una vez finalizado el 100% de la cola:
* **Reporte Acumulativo en Excel:** Toma los resultados, desglosa los SKUs e imprime todo en una nueva pestaña (por ej: `Ejecucion_20260910_1738`) dentro de `Reporte_Final_POs.xlsx`. Esto permite condensar todo el historial diario de automatización en un solo documento sin sobrescribir información pasada.
* **Integración Jira:** Crea automáticamente un ticket en Jira informando de los resultados. Adjunta de manera nativa tanto el `Reporte_Final_POs.xlsx` (para control del negocio) como el `automation.log` (evidencia técnica bitácora de la ejecución) para el soporte en la nube.

## ⚙️ Requisitos y Dependencias

* **Python 3.x**
* **Bibliotecas Clave:** `selenium`, `pandas`, `openpyxl`, `xlwt`, `pyodbc`, `requests`
* **Base de datos:** Controlador ODBC de sistema configurado (`DSN=RI_TEST`).
* **Jira:** Variables de entorno (`JIRA_URL`, `JIRA_USER`, `JIRA_TOKEN`) configuradas en el entorno local para conectividad de su API.

## 📂 Estructura del Código

```text
.
├── ARCHIVOS/                  # Carpetas de plantillas autogeneradas (.xls)
├── Formatos/                  # Ubicación del archivo de carga de negocio original
├── Compras.py                 # Clasificador de datos inicial y regenerador de JSON
├── Main.py                    # Script maestro orquestador web y backend
├── po_report.json             # Estado persistente y base de datos (PO -> RI)
├── Reporte_Final_POs.xlsx     # Excel Acumulativo (múltiples pestañas por ejecución)
├── automation.log             # Archivo de bitácora técnica de todo el ciclo de vida
├── Page/                      # Clases Page Object Model para manejo inteligente de la UI
└── OCEANO/                    # Lógica de campos específicos del formulario OCEANO
```

## 🛠️ Instrucciones de Ejecución

1. **Prepara los datos:** Deposita tu matriz en formato `Formatos/OC compras.xlsx`.
2. **Genera las plantillas:** Ejecuta el generador inicial que ordenará las Importadas:
   ```powershell
   python Compras.py
   ```
3. **Inicia el proceso maestro:** Lanza el orquestador principal que hará la magia y levantará tu navegador Chrome:
   ```powershell
   python Main.py
   ```
4. **Verificación:** Al finalizar, el bot limpiará su rastro, emitirá reportes por consola. Entra a Jira y revisa que tu ticket fue generado con ambos adjuntos de validación.
