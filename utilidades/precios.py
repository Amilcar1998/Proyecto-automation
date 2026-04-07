import os
import pyodbc
import pandas as pd
import xlwt
import warnings

warnings.simplefilter("ignore")

warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

def extraer_zonas(conn,pais_num):
    print(f"[DEBUG] Extrayendo zonas para RI{pais_num}DB")
    sql=f"""
    SELECT DISTINCT ZCOZC AS ZONA
    FROM RI{pais_num}DB.KZCDP
    WHERE O08ZC='E' and ZCOZC IN ('CRC','INT','RSO','TRO')
    """
    df=pd.read_sql_query(sql,conn)
    zonas=df['ZONA'].dropna().unique().tolist()
    print(f"[DEBUG] Zonas encontradas: {zonas}")
    return zonas

def extraer_skus(conn,pais_num):
    extrae_num=pais_num[-1]
    print(f"[DEBUG] Extrayendo SKUs país {pais_num}")

    sql=f"""
    SELECT B.SKU_NUMBER AS SKU, CURRENT_PRICE-1 AS PRECIO FROM SUMMER.SKUBARET A INNER JOIN SUMMER.SKU B ON A.SKU_CODE=B.SKU_CODE 
    INNER JOIN SUMMER.SKU_HIERARCHY SH ON B.SKU_CODE=SH.SKU_CODE inner join RI{pais}DB.KSKUP X ON B.SKU_CODE = X.SKUSK 
    WHERE A.COMPANY_ID='02' AND A.COUNTRY_ID='0{extrae_num}' AND CURRENT_PRICE>0  AND SH.COMPANY_ID=A.COMPANY_ID AND 
    SH.COUNTRY_ID=A.COUNTRY_ID AND SKU_TYPE_CODE<>'D' AND X.B34SK>0 AND B.SKU_NUMBER NOT IN(SELECT DISTINCT PBKSKU FROM RI{pais_num}DB.PCPRCBKP 
    INNER JOIN RI11DB.CONSULRP3 B ON PBKSKU= B.SKU_CORP WHERE COD_DIV ='Y') AND SH.DEPARTMENT_ID <> '725' """
    print(f"sql : {sql}")
    df=pd.read_sql_query(sql,conn)
    print(f"[DEBUG] SKUs encontrados: {len(df)}")
    return df

config_dir=os.path.join('ARCHIVOS_CONFIG','PRECIOS_CONFIG')
os.makedirs(config_dir,exist_ok=True)

plantillas_file=os.path.join(config_dir,'config_plantillas.txt')

print(f"[DEBUG] Leyendo configuración: {plantillas_file}")
print("[DEBUG] Ruta absoluta:", os.path.abspath(plantillas_file))

plantillas=[]

with open(plantillas_file) as f:

    for line in f:

        line=line.strip()

        if not line or line.startswith('#'):
            continue

        p=line.split(',')

        plantillas.append({
            'pais_num':p[0],
            'lineas':int(p[1]),
            'zonas':None if p[2]=='*' else p[2].split('|'),
            'num_plantillas':int(p[3]),
            'all_zones':p[4]=='1'
        })



conn=pyodbc.connect("DSN=RI_TEST;UID=ELOPEZ;PWD=MAY2024")

print("[DEBUG] Conexión BD establecida")

try:

    for conf in plantillas:

        pais=conf['pais_num']
        num_lineas=conf['lineas']
        zonas_conf=conf['zonas']
        num_plantillas=conf['num_plantillas']
        all_zones=conf['all_zones']

        df_skus=extraer_skus(conn,pais)
        zonas_db=extraer_zonas(conn,pais)

        if df_skus.empty or not zonas_db:
            print(f"[DEBUG] Sin datos país {pais}")
            continue

        if zonas_conf is None:
            zonas_usar=zonas_db
        else:
            zonas_usar=[z for z in zonas_conf if z in zonas_db]

        print(f"[DEBUG] País {pais}")
        print(f"[DEBUG] Zonas finales usadas: {zonas_usar}")

        output_dir=f'PLANTILLAS/PRECIOS_RI{pais}DB'
        os.makedirs(output_dir,exist_ok=True)

        country=f"0{pais[-1]}"

        for i in range(1,num_plantillas+1):

            if all_zones and i==1:
                zonas_loop=zonas_db
            else:
                zonas_loop=zonas_usar

            sample=df_skus.sample(
                n=min(num_lineas,len(df_skus)),
                random_state=i
            )

            for zona in zonas_loop:

                wb = xlwt.Workbook()
                ws = wb.add_sheet('DATOS')

                headers = ['COUNTRY','COMPANY','ZONE','SKU','PRECIOS']

                for col, h in enumerate(headers):
                    ws.write(0, col, h)

                fila = 1

                for _, r in sample.iterrows():

                    ws.write(fila, 0, country)
                    ws.write(fila, 1, '02')
                    ws.write(fila, 2, zona)
                    ws.write(fila, 3, r['SKU'])
                    ws.write(fila, 4, r['PRECIO'])

                    fila += 1

                sufijo = 'ALLZONES' if all_zones and i == 1 else zona
                nombre = f'plantilla_RI{pais}DB_{sufijo}_{i}.xls'


                ruta = os.path.join(output_dir, nombre)

                wb.save(ruta)

                print(f"[DEBUG] Archivo generado: {ruta} filas={fila-1}")

finally:

    conn.close()

    print("[DEBUG] Conexión cerrada")
