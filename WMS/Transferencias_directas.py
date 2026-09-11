import random
import pyodbc
from datetime import datetime
from pathlib import Path
import pandas as pd
import sys
import os

# Añadir el directorio raíz al path para permitir importar módulo jira
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from jira.jira_utilidades import JiraClient

# Conectar usando el DSN configurado en ODBC
try:
    conn = pyodbc.connect(
        "DSN=RI_TEST;UID=ELOPEZ;PWD=MAY2024;CommitMode=0;Transaction isolation=No Commit"
    )
except pyodbc.Error as ex:
    sqlstate = ex.args[0]
    print(f"Error de conexión a la base de datos: {sqlstate}")
    exit()

def insert_ktchpwms(cursor, base, bodega, documento, linea, tienda, unidades, fecha, sku, usuario):  
    query_select = f"SELECT B34SK, DIVSK, DPÑSK, VSYSK, VNÑSK, CLSSK FROM {base}.KSKUP WHERE SKUSK = ?"
    cursor.execute(query_select, (sku,))
    fila = cursor.fetchone()
    if not fila:
        raise ValueError(f"No se encontraron datos en KSKUP para el SKU: {sku}")
    
    b34sk, divsk, dpñsk, vsysk, vnñsk, clssk = fila
    
    query_insert = f"""
    INSERT INTO {base}.KTCHPWMS (
        SFTTD, DOCTD, LN4TD, STTTD, SKUTD, O38TD, O39TD, O71TD, 
        U71TD, B96TD, B89TD, DIVTD, DPÑTD, O43TD, S55TD, O41TD, 
        O42TD, VSYTD, D37TD, ETKTD, C65TD, O08TD, O70TD, C104TD, 
        A20TD, VNÑTD, CLSTD, STYTD, C59TD, BP05TD, TYPTD
    ) 
    VALUES (?, ?, ?, ?, ?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    parameters = (bodega, documento, linea, tienda, sku,'', '', '', unidades,b34sk, 0, divsk, dpñsk, '', '', '', '', 
        vsysk, fecha, usuario, ' ', ' ', '  ', '  ',0.00, vnñsk, clssk, '', 'I', 1.1300, 'SH ')
    
    cursor.execute(query_insert, parameters)
    print("Inserción exitosa En KTCHPWMS.")

def insert_siftransf(cursor, base, bodega, documento, fecha, referencia):
    Nreferencia = f"{referencia} by ELISEO LOPEZ"
    if len(Nreferencia) > 50:
        Nreferencia = Nreferencia[:50]
    query_check = f"SELECT COUNT(*) FROM {base}.SIFTRANSF WHERE SFTTD = ? AND DOCTD = ?"
    cursor.execute(query_check, (bodega, documento))
    count = cursor.fetchone()[0]
    if count > 0:
        query_delete = f"DELETE FROM {base}.SIFTRANSF WHERE SFTTD = ? AND DOCTD = ?"
        cursor.execute(query_delete, (bodega, documento))
    
    query_insert = f"INSERT INTO {base}.SIFTRANSF (SFTTD, DOCTD, D37TD, REFTD, STATD) VALUES (?, ?, ?, ?, ?)"
    cursor.execute(query_insert, (bodega, documento, fecha, Nreferencia,'Y'))
    print('Inserción exitosa En SIFTRANSFE.')
        
def insert_bitac_info(cursor, base, documento, tienda,bodega,tipo):
    pk_template = f"{documento}||{tienda}"
    if tipo not in [66, 63,65,64, 70, 71, 72, 4]:
        pk_template += "||08"
    
    query2 = f"SELECT PKREGISTRO FROM {base}.BITAC_INFO bi WHERE TRIM(PKREGISTRO) = ?"
    cursor.execute(query2, (pk_template,))
    row = cursor.fetchone()
    if row:
        print(f"PKREGISTRO encontrado se debe ignorar: {row[0]}")
        return row[0]
    
    query_insert = f"INSERT INTO {base}.BITAC_INFO (IDBASEDATOS, IDTABLA, TIPOPROCESO, FECHAHORA, PKREGISTRO) VALUES (50114, 52282, 'I', CURRENT_TIMESTAMP, ?)"
    cursor.execute(query_insert, (pk_template,))
    print(f'Se realiza el insert en BITAC_INFO para PK: {pk_template}')
    return pk_template

def obtener_e_insertar_documento(cursor, bodega, base):
    fecha_actual = int(datetime.today().strftime('%Y%m%d'))
    query_select = f"SELECT DOCCD + 1 AS DOCCD FROM {base}.KDLTL WHERE STÑCD = ? AND C32CD = 'T' ORDER BY DOCCD DESC FETCH FIRST 1 ROWS ONLY"
    cursor.execute(query_select, (bodega,))
    row = cursor.fetchone()
    
    nuevo_doc = int(row[0]) if row else 1
    
    query_insert = f"INSERT INTO {base}.KDLTL (DCBCD, STÑCD, C32CD, DOCCD, C02CD, SAUCD, STTCD, D54CD) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    valores_insert = ('0', bodega, 'T', nuevo_doc, 'D', bodega, '0', fecha_actual)  
    cursor.execute(query_insert, valores_insert)
    return nuevo_doc

def obtener_skus(cursor, base):
    query_sku = f"""
    SELECT SUBSTRING(di.DSPSKU,2,6) AS SKU, di.DSPVENDOR  FROM {base}.DSP_ITEMS di INNER JOIN ELOPEZ.DSP_ITEMS di2 ON DI.DSPSKU =di2.DSPSKU 
    AND di.DSPSKU <> ' ' INNER JOIN {base}.KSK2P kp ON kp.skuk2=SUBSTRING(di2.DSPSKU,2,6) WHERE kp.O67K2 =di2.DSPVENDOR
    """
    cursor.execute(query_sku)
    rows = cursor.fetchall()
    dspvendor_empty = [row.SKU for row in rows if row.DSPVENDOR == ' ']
    dspvendor_y = [row.SKU for row in rows if row.DSPVENDOR == 'Y']
    return {"dspvendor_empty": dspvendor_empty, "dspvendor_y": dspvendor_y}

def _limpiar_valor_excel(valor):
    """Convierte un valor de Excel a string, eliminando '.0' si es un float."""
    if valor is None:
        return None
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()

def agrupar_transferencias_por_id(filas):
    grupos = {}
    for fila in filas:
        id_valor = fila.get("id")
        if id_valor is None or str(id_valor).strip() == "":
            print("Aviso: Se encontró una fila sin 'id' en el Excel y será ignorada.")
            continue

        id_transferencia = _limpiar_valor_excel(id_valor)

        tipo_val = fila.get("tipo")
        tipo = None
        if tipo_val is not None and str(tipo_val).strip() != '':
            try:
                tipo = int(float(tipo_val))
            except (ValueError, TypeError):
                raise ValueError(f"El valor de 'tipo' ('{tipo_val}') debe ser numérico en la fila con id '{id_transferencia}'.")

        bodega = _limpiar_valor_excel(fila.get("bodega"))
        referencia = _limpiar_valor_excel(fila.get("referencia")) or ""
        tienda = _limpiar_valor_excel(fila.get("tienda"))
        sku = _limpiar_valor_excel(fila.get("sku"))
        unidades = fila.get("unidades")

        if id_transferencia not in grupos:
            grupos[id_transferencia] = {
                "id": id_transferencia,
                "tipo": tipo,
                "bodega": bodega,
                "referencia": referencia,
                "lineas": []
            }

        grupo = grupos[id_transferencia]
        if grupo["tipo"] is None:
            grupo["tipo"] = tipo
        if grupo["bodega"] is None and bodega:
            grupo["bodega"] = bodega
        if not grupo["referencia"] and referencia:
            grupo["referencia"] = referencia
        
        linea_info = {"tienda": tienda, "sku": sku}
        if unidades is not None and str(unidades).strip() != '':
            try:
                linea_info["unidades"] = int(float(unidades))
            except (ValueError, TypeError):
                print(f"Aviso: La columna 'unidades' para el ID {id_transferencia} tiene un valor no numérico '{unidades}' y será ignorada.")

        grupo["lineas"].append(linea_info)
        
    return list(grupos.values())

def obtener_sku_para_tipo(tipo, dspvendor_empty_skus, dspvendor_y_skus, skus_usados):
    if tipo in [65, 64]:
        lista = dspvendor_y_skus
    else:
        lista = dspvendor_empty_skus
    
    while lista:
        sku = lista.pop(0) 
        if sku not in skus_usados:
            skus_usados.add(sku)
            return sku
    return None

def obtener_numero_aleatorio():
    return random.randint(1, 4)

def insert_base(cursor, base, bodega, documento, linea, tienda, unidades, fecha, sku, usuario,referencia,tipo):
    try:
        # Reutiliza el cursor para todas las operaciones
        insert_ktchpwms(cursor, base, bodega, documento, linea, tienda, unidades, fecha, sku, usuario)
        insert_siftransf(cursor, base, bodega, documento, fecha, referencia)
        insert_bitac_info(cursor, base, documento, tienda, bodega, tipo)
        
        conn.commit()
        print(f"[OK] Documento {documento} - Línea {linea} procesada correctamente.")
    
    except ValueError as e:
        print(f"[ERROR] Error de validación en Doc {documento} - Línea {linea}: {e}")
        conn.rollback()
    except Exception as e:
        print(f"[ERROR] Error al insertar en base para Doc {documento} - Línea {linea}: {e}")
        conn.rollback()

def leer_transferencias_desde_excel(ruta_archivo):
    """
    Lee un archivo de Excel y devuelve los datos como una lista de diccionarios.
    """
    try:
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
        # Convierte el DataFrame a una lista de diccionarios
        return df.to_dict('records')
    except FileNotFoundError:
        raise FileNotFoundError(f"El archivo '{ruta_archivo}' no fue encontrado.")
    except Exception as e:
        # Captura otras posibles excepciones de pandas y las relanza como un error genérico
        raise RuntimeError(f"No se pudo leer el archivo de Excel: {e}")
      
import json

def obtener_base_desde_config():
    ruta_config = os.path.join(os.path.dirname(__file__), '..', 'archivos_config', 'destinatarios_flujos.json')
    try:
        with open(ruta_config, 'r', encoding='utf-8') as f:
            config = json.load(f)
            bases = config.get('bases_trabajo', [])
            if bases:
                return f"RI{bases[0]}DB"
    except Exception as e:
        print(f"Error al leer configuración de base: {e}")
    return 'RI14DB' # Fallback por defecto

def main():
    base = obtener_base_desde_config()
    fecha_actual = datetime.today().strftime('%Y%m%d')
    usuario_default = 'ELOPEZ'
    ruta_excel = 'transferencias.xlsx'

    try:
        filas_excel = leer_transferencias_desde_excel(ruta_excel)
        transferencias_agrupadas = agrupar_transferencias_por_id(filas_excel)
    except (FileNotFoundError, ValueError, RuntimeError, ImportError) as e:
        print(f"Error al procesar el archivo de Excel: {e}")
        return

    cursor = conn.cursor()
    try:
        skus_data = obtener_skus(cursor, base)
    except Exception as e:
        print(f"Error al obtener SKUs de la base de datos: {e}")
        cursor.close()
        conn.close()
        return

    dspvendor_empty_skus = [sku for sku in skus_data["dspvendor_empty"]]
    dspvendor_y_skus = [sku for sku in skus_data["dspvendor_y"]]
    skus_usados = set()
    
    df_excel = pd.read_excel(ruta_excel, engine='openpyxl')
    if "documento_generado" not in df_excel.columns:
        df_excel["documento_generado"] = pd.Series(dtype='str')

    documentos_generados = []

    for transferencia in transferencias_agrupadas:
        tipo = transferencia["tipo"]
        bodega_origen = transferencia["bodega"]
        referencia = transferencia["referencia"]

        if not tipo or not bodega_origen:
            print(f"⚠️ Saltando transferencia con id '{transferencia['id']}' por falta de 'tipo' o 'bodega'.")
            continue

        print(f"\n--- Procesando transferencia ID: {transferencia['id']}, Tipo: {tipo}, Bodega: {bodega_origen} ---")

        try:
            documento = obtener_e_insertar_documento(cursor, bodega_origen, base)
            if not documento:
                print(f"❌ No se pudo generar un número de documento para la bodega {bodega_origen}. Saltando transferencia.")
                continue
        except Exception as e:
            print(f"❌ Error al generar documento para la bodega {bodega_origen}: {e}")
            continue
        
        # Guardar el documento generado en el DataFrame
        indices = df_excel[df_excel['id'].astype(str) == str(transferencia['id'])].index
        df_excel.loc[indices, 'documento_generado'] = str(documento)
        documentos_generados.append(documento)

        for i, linea_data in enumerate(transferencia["lineas"]):
            linea_num = i + 1
            tienda_destino = linea_data["tienda"]
            sku_asignado = linea_data["sku"]
            unidades = linea_data.get("unidades") or obtener_numero_aleatorio()

            if not tienda_destino:
                print(f"⚠️ Saltando línea {linea_num} de la transferencia {documento} por falta de 'tienda'.")
                continue

            if not sku_asignado:
                sku_asignado = obtener_sku_para_tipo(tipo, dspvendor_empty_skus, dspvendor_y_skus, skus_usados)
                if not sku_asignado:
                    print("❌ No hay más SKUs disponibles para asignar. Deteniendo el proceso.")
                    break 
                print(f"   -> Línea {linea_num}: SKU no especificado, asignando automáticamente: {sku_asignado}")

            pais_code = base[2:4]
            usuario_dinamico = f"RI{pais_code}CD001"
            usuario = usuario_dinamico if tipo in (63, 64) else usuario_default

            print(f"   -> Insertando línea {linea_num}: Doc: {documento}, Tienda: {tienda_destino}, SKU: {sku_asignado}, Unidades: {unidades}")
            
            insert_base(
                cursor,
                base=base,
                bodega=bodega_origen,
                documento=documento,
                linea=linea_num,
                tienda=tienda_destino,
                unidades=unidades,
                fecha=fecha_actual,
                sku=sku_asignado,
                usuario=usuario,
                referencia=referencia,
                tipo=tipo
            )
        else: # Se ejecuta si el for de lineas termina sin break
            continue
        break # Se ejecuta si el sku no se pudo asignar

    try:
        # ---------------- NUEVA COLUMNA ----------------
        def generar_identificador(row):
            tienda = str(row.get('tienda', '')).split('.')[0] if pd.notna(row.get('tienda')) else ''
            doc = str(row.get('documento_generado', '')).split('.')[0] if pd.notna(row.get('documento_generado')) else ''
            tipo_raw = row.get('tipo', '')
            
            if not doc:
                return ''
                
            try:
                tipo_str = str(int(float(tipo_raw))).zfill(2)
            except:
                tipo_str = str(tipo_raw).zfill(2)
                
            if tipo_str in ['08', '04', '8', '4']:
                return f"ASN_{tipo_str.zfill(2)}_{doc}_1"
            else:
                return f"{tienda}_{tipo_str.zfill(2)}_{doc}_1"

        df_excel['codigo_generado'] = df_excel.apply(generar_identificador, axis=1)
        # -----------------------------------------------

        df_excel.to_excel(ruta_excel, index=False, engine='openpyxl')
        print(f"\n✅ Archivo '{ruta_excel}' actualizado con los números de documento y códigos generados.")
        
        # ---------------- JIRA ----------------
        if documentos_generados:
            try:
                print("\n--- Generando tarea en Jira ---")
                jira = JiraClient()
                ruta_absoluta_excel = os.path.abspath(ruta_excel)
                docs_unicos = list(set(documentos_generados))
                
                # Generar correlativo y fecha
                ahora = datetime.now()
                correlativo = ahora.strftime('%Y%m%d%H%M%S')
                fecha_formateada = ahora.strftime('%d/%m/%Y %H:%M:%S')

                # Resumen para el título de la tarea (se usará como texto que acompaña a "Generación de transferencias")
                str_docs = ", ".join(map(str, docs_unicos))
                if len(str_docs) > 100:
                    str_docs = f"({len(docs_unicos)} documentos)"
                
                load_id = f"Ejecución {correlativo} [{fecha_formateada}] - Docs: {str_docs}"
                
                # Extraer los datos del excel para recrear la tabla en Jira (limitado a 100 por seguridad)
                df_to_jira = df_excel.head(100).fillna('')
                datos_tabla = [df_to_jira.columns.tolist()] + df_to_jira.values.tolist()
                
                print(f"-> Creando única tarea Jira para la ejecución: {load_id}")
                key_jira = jira.main_jira(load_id, ruta_absoluta_excel, "transferencia_directa", base, datos_tabla)
                if key_jira:
                    url_jira = f"https://grupounicomer.atlassian.net/browse/{key_jira}"
                    print(f"\n🔗 Link de la tarea creada: {url_jira}")
            except Exception as e:
                print(f"❌ Error al generar tarea en Jira: {e}")
                
            # ---------------- CORREO ----------------
            try:
                from utilidades.notificaciones_correo import NotificadorCorreo
                print("\n--- Enviando reporte por correo ---")
                
                ruta_absoluta_excel = os.path.abspath(ruta_excel)
                docs_unicos = list(set(documentos_generados))
                
                asunto = f"Reporte de Transferencias Directas - {datetime.now().strftime('%d/%m/%Y')}"
                
                # Crear tabla HTML recorriendo los datos con un FOR (como en los otros flujos)
                df_html = df_excel.fillna('')
                
                # 1. Generar encabezados
                encabezados_html = ""
                for col in df_html.columns:
                    encabezados_html += f'<th style="background-color: #f2f2f2; padding: 8px; border: 1px solid #ccc;">{col}</th>\n'
                
                # 2. Generar filas
                filas_html = ""
                for index, row in df_html.iterrows():
                    filas_html += "<tr>\n"
                    for col in df_html.columns:
                        filas_html += f'<td style="padding: 5px; border: 1px solid #ccc;">{row[col]}</td>\n'
                    filas_html += "</tr>\n"
                
                # 3. Ensamblar tabla
                tabla_html = f"""
                <table style="width:100%; border-collapse: collapse; border: 1px solid #ccc; text-align: center; font-size: 12px;">
                    <tr>
                        {encabezados_html}
                    </tr>
                    {filas_html}
                </table>
                """
                
                cuerpo = f"""
                <div style="font-family: Arial, sans-serif;">
                    <p>Estimado equipo,</p>
                    <p>Se adjunta el reporte de Transferencias Directas generado en la ejecución.</p>
                    <p>Documentos procesados: <b>{len(docs_unicos)}</b></p>
                    <br>
                    <h3 style="color: #2E74B5;">Detalle de Transferencias</h3>
                    {tabla_html}
                </div>
                """
                
                notificador = NotificadorCorreo(flujo='wms_infor')
                enviado = notificador.enviar(
                    asunto=asunto,
                    cuerpo=cuerpo,
                    adjuntos=[ruta_absoluta_excel]
                )
                if enviado:
                    print("✅ Reporte enviado por correo exitosamente.")
                else:
                    print("❌ Error al enviar el reporte por correo.")
            except Exception as e:
                print(f"❌ Excepción al intentar enviar el correo: {e}")
        # ---------------------------------------

    except Exception as e:
        print(f"\n❌ No se pudo guardar el archivo de Excel: {e}")

    print("\nProceso de transferencias desde Excel completado.")
    cursor.close()
    conn.close()

if __name__ == '__main__':
    main()
