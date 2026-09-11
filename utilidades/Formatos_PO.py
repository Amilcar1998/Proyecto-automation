import os
import re
import math
import logging
from datetime import datetime
import warnings

import pyodbc
import pandas as pd
import xlwt

warnings.filterwarnings(
    "ignore",
    message="pandas only supports SQLAlchemy connectable*"
)

LOG_FILE = f"proceso_oc_{datetime.now().strftime('%Y%m%d')}.log"

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.propagate = False

if not logger.handlers:
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    file_handler = logging.FileHandler(os.path.abspath(LOG_FILE), mode='w', encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)


def limpiar_tienda(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto


def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


class GeneradorPlantillasOC:
    def __init__(self, dsn, uid=None, pwd=None):
        self.dsn = dsn
        self.uid = uid
        self.pwd = pwd
        self.conn = None

    def conectar(self):
        try:
            if self.uid and self.pwd:
                self.conn = pyodbc.connect(
                    f"DSN={self.dsn};UID={self.uid};PWD={self.pwd}",
                    autocommit=False
                )
            else:
                self.conn = pyodbc.connect(
                    f"DSN={self.dsn}",
                    autocommit=False
                )

            logger.info("Conexión establecida correctamente")
            return self.conn

        except Exception as e:
            logger.error(f"Error conectando a la base de datos: {e}")
            raise

    def cerrar(self):
        try:
            if self.conn:
                self.conn.close()
                self.conn = None
                logger.info("Conexión cerrada correctamente")
        except Exception as e:
            logger.warning(f"Error cerrando conexión: {e}")

    def validar_conexion(self):
        try:
            if self.conn is None:
                return False

            cursor = self.conn.cursor()
            cursor.execute("SELECT 1 FROM SYSIBM.SYSDUMMY1")
            cursor.fetchone()
            cursor.close()
            return True

        except Exception as e:
            logger.warning(f"Conexión inválida: {e}")
            return False

    def ejecutar_query(self, sql, params=None):
        try:
            if not self.validar_conexion():
                raise Exception("No hay conexión activa con la base de datos")

            df = pd.read_sql(sql, self.conn, params=params)
            logger.info(f"Query ejecutada correctamente. Filas obtenidas: {len(df)}")
            return df

        except Exception as e:
            logger.error(f"Error ejecutando query: {e}")
            raise

    def obtener_datos_base(self, base):
        base = base.upper().strip()

        match = re.fullmatch(r"RI(\d{2})DB", base)
        if not match:
            raise ValueError(f"Formato de base inválido: {base}. Ejemplo válido: RI12DB")

        codigo = match.group(1)
        pais_digito = int(codigo[1])
        country_sql = f"0{pais_digito}"
        company_sql = "02"

        return {
            "base": base,
            "codigo": codigo,
            "pais_digito": pais_digito,
            "country_sql": country_sql,
            "company_sql": company_sql,
            "country_excel": pais_digito,
            "company_excel": 2
        }

    def obtener_items_locales(self, base, limite=None):
        datos = self.obtener_datos_base(base)
        base_sql = datos["base"]
        company = datos["company_sql"]
        country = datos["country_sql"]

        sql = f"""
SELECT c.UPC_CORP, c.SKU_CORP AS SKU, c.DES_ESPA, X.CURRENT_PRICE * 2 AS PRECIO, Y1.LIST_COST +10 AS COSTO, c.DES_CLA, v1.NAME_VENDOR AS VENDOR_1, v2.NAME_VENDOR AS VENDOR_2
FROM RI11DB.CONSULRP3 c INNER JOIN SUMMER.SKU A ON c.SKU_CORP = A.SKU_NUMBER INNER 
JOIN SUMMER.SKU_HIERARCHY sh ON A.SKU_CODE = SH.SKU_CODE
INNER JOIN SUMMER.SKU_INI_PREVIOUS_AUTORIZATION sipa ON c.SKU_CORP = sipa.SKU_NUMBER
LEFT JOIN SUMMER.SKU_INI_SKU_VENDOR v1 ON v1.DEFINITION_VENDOR = '1' AND v1.SKU_ID = sipa.SKU_ID 
LEFT JOIN SUMMER.SKU_INI_SKU_VENDOR v2 ON v2.DEFINITION_VENDOR = '2' AND v2.SKU_ID = sipa.SKU_ID 
LEFT JOIN SUMMER.SKU_INI_SOURCE_VENDOR sc ON sc.ID_VENDOR = v2.CODE_VENDOR 
INNER JOIN SUMMER.SKUBARET X ON A.SKU_CODE = X.SKU_CODE
INNER JOIN SUMMER.SKU_BASIC_COST Y1 ON Y1.SKU_CODE = X.SKU_CODE AND Y1.COMPANY_ID = X.COMPANY_ID AND Y1.COUNTRY_ID = X.COUNTRY_ID 
WHERE c.DES_ESPA <> ' ' AND v2.NAME_VENDOR IS NOT NULL AND SH.COMPANY_ID = X.COMPANY_ID AND SH.COUNTRY_ID = X.COUNTRY_ID AND X.COMPANY_ID = '{company}' 
AND X.COUNTRY_ID = '{country}' AND sh.SKU_TYPE_CODE IN ('A',' ') AND CREATE_DATE <='2025-10-01' AND CURRENT_PRICE > 0 AND v1.NAME_VENDOR <> 'REGAL WORLDWIDE TRADE'
GROUP BY c.UPC_CORP, c.SKU_CORP, c.DES_ESPA, c.DES_CLA, v1.NAME_VENDOR, v2.NAME_VENDOR, CURRENT_PRICE, Y1.LIST_COST
        """
        print(f"SQL: {sql}")
        if limite is not None and limite > 0:
            sql += f" FETCH FIRST {int(limite)} ROWS ONLY"

        return self.ejecutar_query(sql)

    def obtener_items_importados(self, base, limite=None):
        datos = self.obtener_datos_base(base)
        base_sql = datos["base"]
        company = datos["company_sql"]
        country = datos["country_sql"]

        sql = f"""
SELECT c.UPC_CORP, c.SKU_CORP AS SKU, c.DES_ESPA, X.CURRENT_PRICE * 2 AS PRECIO, Y1.LIST_COST AS COSTO, c.DES_CLA, v1.NAME_VENDOR AS VENDOR_1, v2.NAME_VENDOR AS VENDOR_2
FROM RI11DB.CONSULRP3 c INNER JOIN SUMMER.SKU A ON c.SKU_CORP = A.SKU_NUMBER 
INNER JOIN SUMMER.SKU_HIERARCHY sh ON A.SKU_CODE = SH.SKU_CODE
INNER JOIN SUMMER.SKU_INI_PREVIOUS_AUTORIZATION sipa ON c.SKU_CORP = sipa.SKU_NUMBER
LEFT JOIN SUMMER.SKU_INI_SKU_VENDOR v1 ON v1.DEFINITION_VENDOR = '1' AND v1.SKU_ID = sipa.SKU_ID 
LEFT JOIN SUMMER.SKU_INI_SKU_VENDOR v2 ON v2.DEFINITION_VENDOR = '2'
AND v2.SKU_ID = sipa.SKU_ID LEFT JOIN SUMMER.SKU_INI_SOURCE_VENDOR sc ON sc.ID_VENDOR = v2.CODE_VENDOR 
INNER JOIN SUMMER.SKUBARET X ON A.SKU_CODE = X.SKU_CODE
INNER JOIN SUMMER.SKU_BASIC_COST Y1 ON Y1.SKU_CODE = X.SKU_CODE AND Y1.COMPANY_ID = X.COMPANY_ID AND Y1.COUNTRY_ID = X.COUNTRY_ID 
WHERE c.DES_ESPA <> ' ' AND v2.NAME_VENDOR IS NOT NULL AND SH.COMPANY_ID = X.COMPANY_ID AND SH.COUNTRY_ID = X.COUNTRY_ID AND X.COMPANY_ID = '{company}' and 
SKU_TYPE_CODE IN ('A',' ') AND CREATE_DATE <='2025-10-01' AND X.COUNTRY_ID = '{country}' AND CURRENT_PRICE > 0 AND v1.NAME_VENDOR = 'REGAL WORLDWIDE TRADE'
GROUP BY c.UPC_CORP, c.SKU_CORP, c.DES_ESPA, c.DES_CLA, v1.NAME_VENDOR, v2.NAME_VENDOR, CURRENT_PRICE , Y1.LIST_COST
        """

        if limite is not None and limite > 0:
            sql += f" FETCH FIRST {int(limite)} ROWS ONLY"

        return self.ejecutar_query(sql)

    def obtener_tiendas_destino(self, base, limite_tiendas=None):
        datos = self.obtener_datos_base(base)
        codigo = datos["codigo"]
        base_sql = datos["base"]

        sql = f"""
            SELECT A.STRAS AS TIENDA FROM RIUNICOM63.KSTAP A INNER JOIN {base_sql}.KSTRP K ON K.STRKS = A.STRAS
            WHERE A.O08AS = 'A' AND K.O51KS <> 'W' AND A.RIDBS = '{codigo}'
        """

        if limite_tiendas is not None and limite_tiendas > 0:
            sql += f" FETCH FIRST {int(limite_tiendas)} ROWS ONLY"

        return self.ejecutar_query(sql)

    def construir_dataframe_salida(
        self,
        base,
        limite_importadas=None,
        limite_locales=None,
        limite_tiendas_predistribucion=None,
        incluir_locales=False,
        unidades_default=1
    ):
        datos = self.obtener_datos_base(base)
        pais_excel = datos["country_excel"]
        compania_excel = datos["company_excel"]

        df_locales = self.obtener_items_locales(base, limite=limite_locales)
        df_importados = self.obtener_items_importados(base, limite=limite_importadas)
        df_tiendas = self.obtener_tiendas_destino(
            base,
            limite_tiendas=limite_tiendas_predistribucion
        )

        resultado = []

        if incluir_locales:
            for _, row in df_locales.iterrows():
                resultado.append({
                    "PAIS": pais_excel,
                    "COMPANIA": compania_excel,
                    "TIENDA": "",
                    "SKU": limpiar_texto(row.get("SKU", "")),
                    "COLOR": "",
                    "TALLA": "",
                    "MISELANEO": "",
                    "UNIDADES": unidades_default,
                    "PRECIO": row.get("PRECIO", ""),
                    "COSTO": row.get("COSTO", "")
                })

        if df_tiendas.empty:
            logger.warning(f"No se encontraron tiendas para base={base}")
        else:
            for _, row in df_importados.iterrows():
                sku = limpiar_texto(row.get("SKU", ""))
                precio = row.get("PRECIO", "")
                costo = row.get("COSTO", "")
                unidades = row.get("CANTIDADES", unidades_default)

                for _, tienda_row in df_tiendas.iterrows():
                    tienda = limpiar_tienda(tienda_row.get("TIENDA", ""))

                    if not tienda:
                        continue

                    resultado.append({
                        "PAIS": pais_excel,
                        "COMPANIA": compania_excel,
                        "TIENDA": tienda,
                        "SKU": sku,
                        "COLOR": "",
                        "TALLA": "",
                        "MISELANEO": "",
                        "UNIDADES": unidades,
                        "PRECIO": precio,
                        "COSTO": costo
                    })

        df_resultado = pd.DataFrame(resultado, columns=[
            "PAIS",
            "COMPANIA",
            "TIENDA",
            "SKU",
            "COLOR",
            "TALLA",
            "MISELANEO",
            "UNIDADES",
            "PRECIO",
            "COSTO"
        ])

        logger.info(f"Total filas preparadas para salida en {base}: {len(df_resultado)}")
        return df_resultado

    def _escribir_encabezados_xls(self, sheet):
        sheet.write(0, 0, 'PAIS')
        sheet.write(0, 1, 'COMPANIA')
        sheet.write(0, 2, 'TIENDA')
        sheet.write(0, 3, 'SKU')
        sheet.write(0, 4, 'COLOR')
        sheet.write(0, 5, 'TALLA')
        sheet.write(0, 6, 'MISELANEO')
        sheet.write(0, 7, 'UNIDADES')
        sheet.write(0, 8, 'PRECIO')

    def _escribir_encabezados_costo_xls(self, sheet):
        sheet.write(0, 0, 'PAIS')
        sheet.write(0, 1, 'COMPAÑÍA')
        sheet.write(0, 2, 'SKU_NUMBER')
        sheet.write(0, 3, 'RETAIL')
        sheet.write(0, 4, 'COSTO')

    def _obtener_ruta_unica(self, output_dir, nombre_base):
        ruta = os.path.join(output_dir, nombre_base)

        if not os.path.exists(ruta):
            return ruta

        nombre_sin_ext, ext = os.path.splitext(nombre_base)
        consecutivo = 1

        while True:
            nuevo_nombre = f"{nombre_sin_ext}_{consecutivo}{ext}"
            nueva_ruta = os.path.join(output_dir, nuevo_nombre)
            if not os.path.exists(nueva_ruta):
                return nueva_ruta
            consecutivo += 1

    def _guardar_bloque_xls(self, bloque, output_dir, base, archivo_idx):
        wb = xlwt.Workbook()
        sheet = wb.add_sheet('DATOS')
        self._escribir_encabezados_xls(sheet)

        fila = 1
        for _, row in bloque.iterrows():
            sheet.write(fila, 0, row.get("PAIS", ""))
            sheet.write(fila, 1, row.get("COMPANIA", ""))
            sheet.write(fila, 2, limpiar_tienda(row.get("TIENDA", "")))
            sheet.write(fila, 3, limpiar_texto(row.get("SKU", "")))
            sheet.write(fila, 4, limpiar_texto(row.get("COLOR", "")))
            sheet.write(fila, 5, limpiar_texto(row.get("TALLA", "")))
            sheet.write(fila, 6, limpiar_texto(row.get("MISELANEO", "")))
            sheet.write(fila, 7, row.get("UNIDADES", ""))
            sheet.write(fila, 8, row.get("PRECIO", ""))
            fila += 1

        nombre = f"plantilla_{base}_{archivo_idx}.xls"
        ruta = self._obtener_ruta_unica(output_dir, nombre)
        wb.save(ruta)

        return ruta

    def _guardar_bloque_costo_xls(self, bloque, output_dir, base, archivo_idx):
        wb = xlwt.Workbook()
        sheet = wb.add_sheet('DATOS')
        self._escribir_encabezados_costo_xls(sheet)

        # Dejar solo 1 fila por SKU para no repetir información en la plantilla de costos
        bloque_unico = bloque.drop_duplicates(subset=['SKU'])

        fila = 1
        for _, row in bloque_unico.iterrows():
            sheet.write(fila, 0, row.get("PAIS", ""))
            sheet.write(fila, 1, row.get("COMPANIA", ""))
            sheet.write(fila, 2, limpiar_texto(row.get("SKU", "")))
            sheet.write(fila, 3, "")  # RETAIL NO LLEVARÁ NADA
            sheet.write(fila, 4, row.get("COSTO", ""))
            fila += 1

        nombre = f"plantilla_{base}_costo_{archivo_idx}.xls"
        ruta = self._obtener_ruta_unica(output_dir, nombre)
        wb.save(ruta)

        return ruta

    def exportar_xls_por_bloques(self, df_resultado, output_dir, base, limite_lineas=1000):
        os.makedirs(output_dir, exist_ok=True)

        if df_resultado.empty:
            logger.warning(f"No hay datos para exportar en base={base}")
            return []

        archivos_generados = []
        total_filas = len(df_resultado)

        logger.info(
            f"Exportando {total_filas} filas para {base}. "
            f"Límite por archivo: {limite_lineas}"
        )

        archivo_idx = 1

        for inicio in range(0, total_filas, limite_lineas):
            fin = inicio + limite_lineas
            bloque = df_resultado.iloc[inicio:fin].copy()
            ruta_precio = self._guardar_bloque_xls(bloque, output_dir, base, archivo_idx)
            ruta_costo = self._guardar_bloque_costo_xls(bloque, output_dir, base, archivo_idx)
            archivos_generados.extend([ruta_precio, ruta_costo])
            archivo_idx += 1

        return archivos_generados

    def exportar_xls_cantidad_fija(self, df_resultado, output_dir, base, cantidad_archivos):
        os.makedirs(output_dir, exist_ok=True)

        if df_resultado.empty:
            logger.warning(f"No hay datos para exportar en base={base}")
            return []

        if cantidad_archivos <= 0:
            raise ValueError("cantidad_archivos debe ser mayor que cero")

        total_filas = len(df_resultado)
        max_filas_xls = 65535  # 65536 menos 1 encabezado

        filas_por_archivo = math.ceil(total_filas / cantidad_archivos)

        if filas_por_archivo > max_filas_xls:
            cantidad_minima = math.ceil(total_filas / max_filas_xls)
            raise ValueError(
                f"No se puede exportar {total_filas} filas en {cantidad_archivos} archivo(s) .xls. "
                f"Cada archivo tendría {filas_por_archivo} filas y el máximo permitido es {max_filas_xls}. "
                f"Debes usar al menos {cantidad_minima} archivo(s) o cambiar a .xlsx."
            )

        logger.info(
            f"Exportando {total_filas} filas para {base} en exactamente "
            f"{cantidad_archivos} archivo(s). Filas aprox por archivo: {filas_por_archivo}"
        )

        archivos_generados = []
        inicio = 0

        for archivo_idx in range(1, cantidad_archivos + 1):
            fin = inicio + filas_por_archivo
            bloque = df_resultado.iloc[inicio:fin].copy()

            if bloque.empty:
                logger.warning(
                    f"No hay más filas para seguir creando archivos en {base}. "
                    f"Se generaron {len(archivos_generados)} archivo(s)."
                )
                break

            ruta_precio = self._guardar_bloque_xls(bloque, output_dir, base, archivo_idx)
            ruta_costo = self._guardar_bloque_costo_xls(bloque, output_dir, base, archivo_idx)
            archivos_generados.extend([ruta_precio, ruta_costo])
            inicio = fin

        return archivos_generados





    def procesar_base( self, base, output_dir, limite_importadas=None, limite_locales=None, limite_tiendas_predistribucion=None, limite_lineas_archivo=1000, incluir_locales=False, unidades_default=1, cantidad_archivos_fija=None ):
        logger.info(f"===== INICIO procesamiento base {base} =====")

        df_resultado = self.construir_dataframe_salida(
            base=base,
            limite_importadas=limite_importadas,
            limite_locales=limite_locales,
            limite_tiendas_predistribucion=limite_tiendas_predistribucion,
            incluir_locales=incluir_locales,
            unidades_default=unidades_default
        )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_base_dir = os.path.join(output_dir, base, timestamp)

        if cantidad_archivos_fija is not None:
            archivos = self.exportar_xls_cantidad_fija(
                df_resultado=df_resultado,
                output_dir=output_base_dir,
                base=base,
                cantidad_archivos=cantidad_archivos_fija
            )
        else:
            archivos = self.exportar_xls_por_bloques(
                df_resultado=df_resultado,
                output_dir=output_base_dir,
                base=base,
                limite_lineas=limite_lineas_archivo
            )

        logger.info(f"===== FIN procesamiento base {base} | archivos={len(archivos)} =====")
        return {
            "base": base,
            "filas": len(df_resultado),
            "archivos": archivos
        }

    def procesar_bases( self, bases, output_dir, limite_importadas=None, limite_locales=None, limite_tiendas_predistribucion=None, limite_lineas_archivo=1000, incluir_locales=False, unidades_default=1, cantidad_archivos_fija=None ):
        resultados = []

        for base in bases:
            try:
                resultado = self.procesar_base(
                    base=base,
                    output_dir=output_dir,
                    limite_importadas=limite_importadas,
                    limite_locales=limite_locales,
                    limite_tiendas_predistribucion=limite_tiendas_predistribucion,
                    limite_lineas_archivo=limite_lineas_archivo,
                    incluir_locales=incluir_locales,
                    unidades_default=unidades_default,
                    cantidad_archivos_fija=cantidad_archivos_fija
                )
                resultados.append(resultado)

            except Exception as e:
                logger.error(f"Error procesando base {base}: {e}")

        return resultados


def main():
    dsn = "RI_TEST"
    uid = "ELOPEZ"
    pwd = "MAY2024"

    bases = ["RI11DB", "RI12DB", "RI13DB","RI14DB"]

    salida_root = os.path.abspath("salida_plantillas")

    limite_importadas = 10
    limite_locales = 10
    limite_tiendas_predistribucion = 1
    limite_lineas_archivo = 1000
    incluir_locales = False
    unidades_default = 10

    # Si quieres exactamente 100 archivos por base, usa esto:
    cantidad_archivos_fija = None

    proceso = GeneradorPlantillasOC(dsn=dsn, uid=uid, pwd=pwd)

    try:
        proceso.conectar()

        resultados = proceso.procesar_bases(bases=bases,output_dir=salida_root,limite_importadas=limite_importadas,limite_locales=limite_locales,limite_tiendas_predistribucion=limite_tiendas_predistribucion,limite_lineas_archivo=limite_lineas_archivo,incluir_locales=incluir_locales,unidades_default=unidades_default,cantidad_archivos_fija=cantidad_archivos_fija)

        print("\nResumen:")
        for r in resultados:
            print(f"Base: {r['base']} | Filas: {r['filas']} | Archivos: {len(r['archivos'])}")
            for archivo in r["archivos"]:
                print(f"  - {archivo}")

    finally:
        proceso.cerrar()


if __name__ == "__main__":
    main()