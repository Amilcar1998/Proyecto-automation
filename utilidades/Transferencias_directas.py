import random
import pyodbc
from datetime import datetime


# Conectar usando el DSN configurado en ODBC

conn = pyodbc.connect(
    "DSN=RI_TEST;UID=ELOPEZ;PWD=MAY2024;CommitMode=0;Transaction isolation=No Commit"
)



def insert_ktchpwms(base, bodega, documento, linea, tienda, unidades, fecha, sku, usuario):  

    cursor = conn.cursor()

    query_select = f"""
        SELECT B34SK, DIVSK, DPÑSK, VSYSK, VNÑSK, CLSSK 
        FROM {base}.KSKUP 
        WHERE SKUSK = ?
    """

    # Ejecutar la consulta y obtener resultados
    cursor.execute(query_select, (sku,))
    fila = cursor.fetchone()

    # Validar si se encontraron resultados
    if fila:
        b34sk, divsk, dpñsk, vsysk, vnñsk, clssk = fila
    else:
        print("No se encontraron datos en KSKUP para el SKU:", sku)
        cursor.close()
        conn.close()
        exit()

    # Construcción de la consulta INSERT con parámetros
    query_insert = f"""
    INSERT INTO {base}.KTCHPWMS (
        SFTTD, DOCTD, LN4TD, STTTD, SKUTD, O38TD, O39TD, O71TD, 
        U71TD, B96TD, B89TD, DIVTD, DPÑTD, O43TD, S55TD, O41TD, 
        O42TD, VSYTD, D37TD, ETKTD, C65TD, O08TD, O70TD, C104TD, 
        A20TD, VNÑTD, CLSTD, STYTD, C59TD, BP05TD, TYPTD
    ) 
    VALUES (
        ?, ?, ?, ?, ?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
    )
    """
    # Parámetros de la consulta
    parameters = (bodega, documento, linea, tienda, sku,'', '', '', unidades,b34sk, 0, divsk, dpñsk, '', '', '', '', 
        vsysk, fecha, usuario, ' ', ' ', '  ', '  ',0.00, vnñsk, clssk, '', 'I', 1.1300, 'SH '
    )


    #print (f'insert{query_insert}: {parameters}')


    # Ejecutar la consulta INSERT
    try:
        cursor.execute(query_insert, parameters)
        #conn.commit()  # Confirmar la transacción
        print("Inserción exitosa En KTCHPWMS: ")
    except Exception as e:
        print(f"Error al insertar: {e}")

    # Cerrar cursor
    cursor.close()

def insert_siftransf(base, bodega, documento, fecha, referencia):
    cursor=conn.cursor()

    Nreferencia = f"{referencia} :AUTO-GENERADO ELISEO AMILCAR LOPEZ"

    try:
        query_check = f"SELECT COUNT(*) FROM {base}.SIFTRANSF WHERE SFTTD = ? AND DOCTD = ?"
        cursor.execute(query_check, (bodega, documento))
        count = cursor.fetchone()[0]
        
        query_delete = f"DELETE FROM {base}.SIFTRANSF WHERE SFTTD = ? AND DOCTD = ?"
        query_insert = f"""
        INSERT INTO {base}.SIFTRANSF (SFTTD, DOCTD, D37TD, REFTD, STATD)
        VALUES (?, ?, ?, ?, ?)
        """
        


        if count > 0:
            cursor.execute(query_delete, (bodega, documento))
        cursor.execute(query_insert, (bodega, documento, fecha, Nreferencia,'Y'))

        print('Inserción exitosa En SIFTRANSFE')
    except Exception as e:
        print(f"Error al procesar la transacción: {e}")
    
    finally:
        # Cerrar cursor y conexión
        cursor.close()
        
def insert_bitac_info(base, documento, tienda,bodega,tipo):
    cursor = conn.cursor()
    
    try:    
        if tipo in [66, 63,65,64, 70, 71, 72, 4]:
            query2 = f"""
            SELECT PKREGISTRO FROM {base}.BITAC_INFO bi WHERE TRIM(PKREGISTRO) = ?
            """
            pkregistro = f"{documento}||{tienda}"
            cursor.execute(query2, (pkregistro,))
            row = cursor.fetchone()

            if row:
                print(f"PKREGISTRO encontrado se debe ignorar: {row[0]}")
                return row[0]
            else:
                query =f""" 
                        INSERT INTO {base}.BITAC_INFO (IDBASEDATOS, IDTABLA, TIPOPROCESO, FECHAHORA, PKREGISTRO)
                        VALUES (50114, 52282, 'I', CURRENT_TIMESTAMP, ?)
                        """
                cursor.execute(query, (pkregistro,))
                row = cursor.fetchone()
                print('Se realiza el insert en BITAC_INFO')

                return row

        else:
            query2 = f"""
            SELECT PKREGISTRO FROM {base}.BITAC_INFO bi WHERE TRIM(PKREGISTRO) = ?
            """
            pkregistro = f"{documento}||{tienda}||08"

            # Ejecutar la consulta
            cursor.execute(query2, (pkregistro,))
            row = cursor.fetchone()

            if row:
                print(f"PKREGISTRO encontrado: {row[0]}")
                return row[0]
            else:                
                query =f""" 
                        INSERT INTO {base}.BITAC_INFO (IDBASEDATOS, IDTABLA, TIPOPROCESO, FECHAHORA, PKREGISTRO)
                        VALUES (50114, 52282, 'I', CURRENT_TIMESTAMP, ?)
                        """
                cursor.execute(query, (pkregistro,))
                row = cursor.fetchone()
                print (f'ingresado a bitac_info {row[0]}')
                return row[0]

        

    except Exception as e:
        print(f"Error al obtener PKREGISTRO: {e}")
        return None

    finally:
        # Cerrar cursor y conexión
        cursor.close()

def obtener_e_insertar_documento(bodega, base):
    cursor = conn.cursor()
    fecha_actual = int(datetime.today().strftime('%Y%m%d'))  # Fecha actual en formato yyyymmdd
    
    # Consulta para obtener el siguiente número de DOCCD
    query_select = f"""SELECT DOCCD + 1 AS DOCCD FROM {base}.KDLTL WHERE STÑCD = ? AND C32CD = 'T' 
    ORDER BY DOCCD DESC FETCH FIRST 1 ROWS ONLY
    """
    
    cursor.execute(query_select, (bodega,))
    row = cursor.fetchone()
    
    if row:
        #print(f'{row}')
        nuevo_doc = int(row[0])  # Convertir a entero para evitar Decimal
        
        # Insertar el nuevo documento en la tabla
        query_insert = f"""
        INSERT INTO {base}.KDLTL (DCBCD, STÑCD, C32CD, DOCCD, C02CD, SAUCD, STTCD, D54CD)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """
        
        valores_insert = ('0', bodega, 'T', nuevo_doc, 'D', bodega, '0', fecha_actual)  
        cursor.execute(query_insert, valores_insert)
        #conn.commit()  # Confirmar la transacción
        
        return nuevo_doc  # Retornar el DOCCD como entero
    else:
        # Si no existe el documento previo, insertar con DOCCD = 1
        query_insert = f"""INSERT INTO {base}.KDLTL (DCBCD, STÑCD, C32CD, DOCCD, C02CD, SAUCD, STTCD, D54CD)
                          VALUES (?, ?, ?, ?, ?, ?, ?, ?)"""
        valores_insert = ('0', bodega, 'T', 1, 'D', bodega, 0, fecha_actual)

        cursor.execute(query_insert, valores_insert)
        #conn.commit()

        # Obtener el nuevo DOCCD insertado
        query_select = f"""SELECT DOCCD AS DOCCD FROM {base}.KDLTL WHERE STÑCD = ? AND C32CD = 'T' 
        ORDER BY DOCCD DESC FETCH FIRST 1 ROWS ONLY
        """
        
        cursor.execute(query_select, (bodega,))
        row2 = cursor.fetchone()

        if row2:
            #print(f'{row2}')
            nuevo_doc = int(row2[0])  # Convertir a entero el valor obtenido
            return nuevo_doc

        # Si no hay registros previos o error, retornar None
        return None

def obtener_skus(cursor, base):
    
    cursor = conn.cursor()
    # Consulta para obtener los SKUs y DSPVENDOR
    #query_sku = f"""
    #SELECT DISTINCT SUBSTRING(sd.SODSKUPC, 2, 6) AS sku, DSPVENDOR FROM {base}.SO_DTL sd INNER JOIN 
    #{base}.SO_HDR sh ON SD.SODAOEID = SH.SOHEOKEY INNER JOIN RI11DB.DSP_ITEMS b ON 
    #SODSKUPC = b.dspsku WHERE SH.SOHSWMS ='S'
    #"""

    #query_sku = f"""
    #SELECT DISTINCT SUBSTRING(sd.SODSKUPC, 2, 6) AS sku, DSPVENDOR FROM {base}.SO_DTL sd INNER JOIN 
    #{base}.SO_HDR sh ON SD.SODAOEID = SH.SOHEOKEY INNER JOIN RI11DB.DSP_ITEMS b ON 
    #SODSKUPC = b.dspsku WHERE SH.SOHSWMS ='S' and dspsku in (SELECT DSPSKU  FROM ELOPEZ.DSP_ITEMS)
    #"""

   

    query_sku = f"""
    SELECT SUBSTRING(di.DSPSKU,2,6) AS SKU, di.DSPVENDOR  FROM {base}.DSP_ITEMS di INNER JOIN ELOPEZ.DSP_ITEMS di2 ON DI.DSPSKU =di2.DSPSKU 
    AND di.DSPSKU <> ' ' INNER JOIN {base}.KSK2P kp ON kp.skuk2=SUBSTRING(di2.DSPSKU,2,6) WHERE kp.O67K2 =di2.DSPVENDOR
    """

    
    cursor.execute(query_sku)
    rows = cursor.fetchall()
    
    # Separar los resultados según el valor de DSPVENDOR
    dspvendor_empty = []  # Para los DSPVENDOR = ' '
    dspvendor_y = []  # Para los DSPVENDOR = 'Y'
    
    for row in rows:
        sku, dspvendor = row
        if dspvendor == ' ':
            dspvendor_empty.append({"sku": sku, "dspvendor": dspvendor})
        elif dspvendor == 'Y':
            dspvendor_y.append({"sku": sku, "dspvendor": dspvendor})
    
    # Retornar los resultados como lista de diccionarios
    return {"dspvendor_empty": dspvendor_empty, "dspvendor_y": dspvendor_y}

def obtener_tiendas(base):
    cursor = conn.cursor()
    
    # Consulta para obtener los STRKS donde K.O51KS = ' '
    query = f"SELECT STRKS FROM {base}.KSTRP k INNER JOIN RIUNICOM63.KSTAP K2 ON K.STRKS=STRAS WHERE K.O51KS = ' ' AND K2.RIDBS ='12' AND O08AS ='A'"
    
    cursor.execute(query)
    rows = cursor.fetchall()
    
    # Retornar los resultados como una lista de valores STRKS
    tiendas = [row[0] for row in rows]  # Suponiendo que STRKS es la primera columna
    
    return tiendas

def obtener_numero_aleatorio():
    return random.randint(1,4)

def insert_base(base, bodega, documento, linea, tienda, unidades, fecha, sku, usuario,referencia,tipo):
    cursor = conn.cursor()
    """Ejecuta todas las inserciones en una sola transacción."""
   
    try:
        insert_ktchpwms(base,bodega, documento, linea, tienda, unidades, fecha, sku, usuario)
        #               base, bodega, documento, linea, tienda, unidades, fecha, sku, usuario
        insert_siftransf(base, bodega, documento, fecha, referencia)
                       #base, bodega, documento, fecha, referencia
        insert_bitac_info(base, documento, tienda,bodega,tipo)
        
        #cursor.commit()
        print(f"✅ Documento {documento} insertado correctamente.")
    except Exception as e:
        a = cursor.rollback()
        print(f'{a}')
        print(f"❌ Error al insertar: {documento}: {e}")
    finally:
        cursor.close()
      
def main():
    
    documento_origen = None 
    bodegas =  [99] 
    tipos_transferencias = [63, 66,65,64, 70, 71, 72, 8, 4]  # Tipos de transferencias
    transferencias_por_tipo = {
        63: 1,
        66: 1,
        65: 1,  # TIPO 66 VENDOR
        64: 1,  # TIPO 63 VENDOR
        70: 1,
        71: 1,
        72: 1,
        8: 1,
        4: 1
    }  # Cada tipo solo tendrá 1 transferencia
    #lineas_transferencias =obtener_numero_aleatorio()
    lineas_transferencias = 2  # Cantidad de líneas por transferencia
    base = 'RI12DB'

    # 1. Obtener los SKUs disponibles


    cursor = conn.cursor()
    skus = obtener_skus(cursor, base)
    dspvendor_empty_skus = [sku["sku"] for sku in skus["dspvendor_empty"]]
    dspvendor_y_skus = [sku["sku"] for sku in skus["dspvendor_y"]]
    todos_los_skus = dspvendor_empty_skus
    random.shuffle(todos_los_skus)  # Barajar los SKUs para asegurar aleatoriedad




    # 2. Obtener las tiendas
    tiendas = obtener_tiendas(base)
    if not tiendas:
        print("No se pudieron obtener las tiendas.")
        return


    fecha_actual = datetime.today().strftime('%Y%m%d')

    for bodega in bodegas:
        bodega2 ='98'
        for tipo in tipos_transferencias:
  
            cantidad_transferencias = transferencias_por_tipo.get(tipo, 0)  # Obtener la cantidad por tipo
            
            
            for i in range(cantidad_transferencias):
                if cantidad_transferencias > 0:  # Solo procesar si la cantidad es mayor a 0
                    # Generar el documento solo si es necesario (es decir, si aún no existe para este tipo)

                    for linea in range(1, lineas_transferencias + 1):  # Generar tantas líneas como se indique
                        print(f'inicia transferencias: tipo {tipo}')
                        unidades =obtener_numero_aleatorio()

                        if not todos_los_skus:
                            print("No hay más SKUs disponibles.")
                            return
                        
                        skus_usados = set()

                        # Dentro del bucle donde asignas los SKUs:
                        if tipo in [63, 66, 70, 71, 72, 8, 4]:  # No se deben mezclar los tipos de SKU
                            while True:
                                if dspvendor_empty_skus:
                                    sku_asignado = dspvendor_empty_skus.pop()  # Obtener un SKU vacío
                                else:
                                    print("No hay más SKUs disponibles.")
                                    return

                                if sku_asignado not in skus_usados:  # Verificar si ya se usó
                                    skus_usados.add(sku_asignado)  # Marcarlo como usado
                                    #print (f'{skus_usados}{linea}')
                                    break  # Salir del bucle cuando se encuentre un SKU no repetido
                        elif tipo in [65,64]:  #Hace referencia a las tipo 63 y 66 vendor por que no se puede colocarle una letra solo acepta numeros}

                            #print(f'sku_vendor: {dspvendor_y_skus}')

                            while True:
                                if dspvendor_y_skus:
                                    sku_asignado = dspvendor_y_skus.pop()  # Obtener un SKU  
                                    print(sku_asignado)
                                else:
                                    print("No hay más SKUs disponibles.")
                                    return

                                if sku_asignado not in skus_usados:  # Verificar si ya se usó
                                    skus_usados.add(sku_asignado)  # Marcarlo como usado
                                    #print (f'{skus_usados}{linea}')
                                    break  # Salir del bucle cuando se encuentre un SKU no repetido
                        
                    

                        tienda_aleatoria = random.choice(tiendas)  # Seleccionar tienda aleatoria
              
                        if tipo in (63, 64):
                            usuario = 'RI11CD01'
                        else:
                            usuario = 'ELOPEZ'

                        # Imprimir los resultados por tipo de transferencia
                        if tipo in [63, 66, 65, 64]:
                                #print (bodega)
                                # Generar un nuevo documento si es la primera iteración o si ya se procesaron todas las líneas del documento anterior
                                if documento_origen is None:
                                    documento_origen = obtener_e_insertar_documento(bodega, base)
                                    if not documento_origen:
                                        print(f"No se pudo generar el documento de origen {tienda_aleatoria}")
                                        continue
                                
                                #print(f"Tipo: {tipo}")
                                
                                #ejecutar_Procedimiento(cursor, bodega, documento_origen, linea, tienda_aleatoria, fecha_actual, sku_asignado, usuario,referencia)
                                insert_base(base, bodega, documento_origen, linea, tienda_aleatoria, unidades, fecha_actual, sku_asignado, usuario,bodega,tipo)
                                
                                # Si llegamos al final de las líneas de transferencia, forzamos la generación de un nuevo documento en la próxima iteración
                                if linea == lineas_transferencias:
                                    documento_origen = None
                        elif tipo == 70:
                            tienda_aleatoria='90'
                            # Generar un nuevo documento si es la primera iteración o si ya se procesaron todas las líneas del documento anterior
                            if documento_origen is None:
                                documento_origen = obtener_e_insertar_documento(bodega2, base)
                                if not documento_origen:
                                        print(f"No se pudo generar el documento de origen {tienda_aleatoria}")
                                        continue
                                
                            #ejecutar_Procedimiento(cursor, bodega2, documento_origen, linea, tienda_aleatoria, fecha_actual, sku_asignado, usuario,bodega)
                            insert_base(base, bodega2, documento_origen, linea, tienda_aleatoria, unidades, fecha_actual, sku_asignado, usuario,bodega,tipo)

                            if linea == lineas_transferencias:
                                    documento_origen = None
                        elif tipo == 71:
                                tienda_aleatoria ='92'
                                # Generar un nuevo documento si es la primera iteración o si ya se procesaron todas las líneas del documento anterior
                                if documento_origen is None:
                                    documento_origen = obtener_e_insertar_documento(bodega2, base)
                                    if not documento_origen:
                                        print(f"No se pudo generar el documento de origen {tienda_aleatoria}")
                                        continue

                                

                                #ejecutar_Procedimiento(cursor, bodega2, documento_origen, linea, tienda_aleatoria, fecha_actual, sku_asignado, usuario,bodega)
                                insert_base(base, bodega2, documento_origen, linea, tienda_aleatoria, unidades, fecha_actual, sku_asignado, usuario,bodega,tipo)

                                # Si llegamos al final de las líneas de transferencia, forzamos la generación de un nuevo documento en la próxima iteración
                                if linea == lineas_transferencias:
                                    documento_origen = None
                        elif tipo == 72:
                                # Generar un nuevo documento si es la primera iteración o si ya se procesaron todas las líneas del documento anterior
                                if documento_origen is None:
                                    documento_origen = obtener_e_insertar_documento(bodega2, base)
                                    if not documento_origen:
                                        print(f"No se pudo generar el documento de origen {tienda_aleatoria}")
                                        continue
                                    
                                #ejecutar_Procedimiento(cursor, bodega2, documento_origen, linea, tienda_aleatoria, fecha_actual, sku_asignado, usuario,bodega)
                                insert_base(base, bodega, documento_origen, linea, tienda_aleatoria, unidades, fecha_actual, sku_asignado, usuario,bodega,tipo)
                            

                                # Si llegamos al final de las líneas de transferencia, forzamos la generación de un nuevo documento en la próxima iteración
                                if linea == lineas_transferencias:
                                    documento_origen = None
                        elif tipo == 8:
                                
                                if documento_origen is None:
                                    
                                    tienda_aleatoria_origen = random.choice(tiendas)  # Tienda origen aleatoria
                                    tienda_aleatoria_destino = random.choice([t for t in tiendas if t != tienda_aleatoria_origen])  # Tienda destino aleatoria distinta a origen
                                    documento_origen = obtener_e_insertar_documento(tienda_aleatoria_origen, base)

                                    if not documento_origen:
                                        print(f"No se pudo generar el documento de origen {tienda_aleatoria}")
                                        continue

                                insert_base(base, tienda_aleatoria_origen, documento_origen, linea, tienda_aleatoria_destino, unidades, fecha_actual, sku_asignado, usuario,bodega,tipo)

                                if linea == lineas_transferencias:
                                    documento_origen = None  # Esto hará que se genere un nuevo documento en el próximo ciclo

                        elif tipo == 4:
                            # Generar documento con tienda aleatoria y mantenerlo constante para todas las líneas
                            if not todos_los_skus:
                                print("No hay más SKUs disponibles.")
                                return
                            
                            if(documento_origen is None):
                                tienda_aleatoria = random.choice(tiendas)  # Tienda aleatoria para el documento
                                documento_origen = obtener_e_insertar_documento(tienda_aleatoria, base)  # Generar el documento con tienda aleatoria

                                if not documento_origen:
                                    print(f"No se pudo generar el documento de origen {tienda_aleatoria}")
                                    continue

                            
                            sku_asignado = todos_los_skus.pop()  # Obtener un SKU disponible
                            insert_base(base, tienda_aleatoria, documento_origen, linea, bodega, unidades, fecha_actual, sku_asignado, usuario,bodega,tipo)
                            
                            if linea == lineas_transferencias:
                                    documento_origen = None

                print(f"Transferencia del tipo {tipo} completada. Generando")
main()


