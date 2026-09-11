# Proyecto Automation - Ecosistema Principal

Este repositorio principal contiene el ecosistema completo de automatizaciones desarrolladas para la orquestación de datos, integraciones web (Selenium), procesamiento de archivos (FTP, Excel, JSON) e interacción con bases de datos legadas (AS/400 DB2) e INFOR WMS.

## 🚀 Arquitectura y Módulos Principales

El ecosistema está estructurado en carpetas y submódulos especializados para cubrir diferentes frentes de negocio:

### 1. 🏖️ `PO_SUMMER/`
* **Automatización de Compras:** Core de automatización web con Selenium encargado de inyectar y procesar plantillas de Excel masivas dentro del sistema SUMMER/OCEANO.
* **Flujo Híbrido:** Se encarga de capturar Números de PO, autorizar automáticamente, llamar al procedimiento DB2 (RETACEO) y capturar su RI asociado. Cuenta con un alto nivel de persistencia de datos y un `README.md` interno específico con mayor detalle.

### 2. 📦 `wms_infor/`
* **Módulo WMS:** Lógicas orientadas al entorno logístico WMS INFOR. 
* Se encarga de validaciones de transferencias e integraciones, ejecutando cruces de tablas de inventario para asegurar la consistencia.

### 3. 💵 `precios/`
* **Gestión de Precios:** Módulo (con estructura de API/Servicios: controladores, repositorios, modelos) encargado de procesar y estructurar las actualizaciones de cambios de precios y reportes asociados.

### 4. 🛒 `ordenes_compra/`
* Procesos independientes para la validación, extracción y procesamiento estructural de parámetros de Órdenes de Compra.

### 5. 🛡️ `garantias/`
* Procesamiento masivo de lógicas de retornos o garantías de artículos, con capacidades de exportación en Excel (ej. `garantias_export.xlsx`).

### 6. ⚙️ `as400_core/` & `conexion_config/`
* **Capa de Datos:** El núcleo duro para todo el ecosistema. Contiene las interfaces ODBC (`pyodbc`) y configuraciones de logs para comunicarse de forma unificada con el servidor AS/400 (DB2 - DSN: `RI_TEST`).

### 7. 🛠️ `utilidades/`
* **Herramientas Transversales:** Scripts auxiliares para manipulación y extracción de XMLs (UPCS), gestión y descargas automatizadas vía FTP, y envío automático de notificaciones de correo.

### 8. 🎫 `jira/`
* **Integración Atlassian API:** Librería propia de la empresa encargada de construir los "payloads" para Jira (Zephyr). Automatiza la apertura de tickets, publicación de tablas y resultados, y el envío de evidencias (`.log` y reportes) directamente a la nube al terminar cualquier proceso.

## 💻 Puntos de Entrada (Orquestadores)

En la raíz descansan los controladores globales del ecosistema:
* `main.py` y `flujo_completo.py`: Enlazan los módulos según sea el caso de uso que se necesite correr.
* `main_garantia.py`: Lanzador específico para el módulo de garantías.
* `main.bat`: Script de terminal nativo de Windows (Batch) diseñado para invocaciones programadas (cron/task scheduler).

## ⚙️ Dependencias e Instalación

Para asegurar compatibilidad, se recomienda el uso de **Python 3.x**. Las dependencias principales están en el archivo `requirements.txt`:
* **Navegación Web y APIs:** `selenium`, `requests`
* **Análisis de Datos:** `pandas`, `openpyxl`, `xlwt`
* **DB:** `pyodbc`

```bash
pip install -r requirements.txt
```

## 🔒 Notas Técnicas y Git

* Toda la arquitectura está centralizada en este repositorio de manera nativa (sin submódulos ocultos).
* **IMPORTANTE:** Dado que los scripts de validación pueden descargar archivos FTP con nombres extremadamente largos (que superan los 260 caracteres nativos de Windows), es obligatorio habilitar `core.longpaths` en el cliente Git local para evitar bloqueos durante los envíos (`git push`) o descargas (`git pull`):
  
  ```bash
  git config core.longpaths true
  ```
