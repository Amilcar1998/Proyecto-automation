from datetime import datetime, timedelta
import decimal
import logging
import random
import time

from conexion_config.conexion import ConexionAS400


class PreciosRepository:
    ZONE_MAPPING_SCHEMA = "RI12DB"

    def __init__(self, bases):
        self.db = ConexionAS400()
        self.bases = bases
        self.logger = logging.getLogger("PreciosRepository")
        self.enie = "\u00D1"
        self.usuario_actjob = "ELOPEZ"
        self.tablas_monitoreo_job = ["QTEMPKILL"]

    def ajustar_hora_ejecucion(self, base):
        self.logger.info("INICIO ajustar_hora_ejecucion")
        cursor = None

        try:
            hora_sistema = datetime.now().strftime("%H:%M:%S")
            self.logger.info(f"Hora sistema detectada {hora_sistema}")
            conn = self.db.conectar()
            cursor = conn.cursor()
            self.logger.info(f"Procesando base {base}")

            pais = base[2:4]
            job = f"RI{pais}PC"
            sql = "SELECT REGHOU FROM RIUNICOM63.REGPARAM WHERE REGJOB = ? FETCH FIRST 1 ROW ONLY"

            cursor.execute(sql, (job,))
            row = cursor.fetchone()

            if not row:
                self.logger.warning(f"{base} REGHOU sin registros")
                return

            hora_actual = str(row[0]).strip()
            if hora_actual[:2] == hora_sistema[:2]:
                self.logger.info(f"{base} hora ya coincide con sistema")
                return

            sql_update = "UPDATE RIUNICOM63.REGPARAM SET REGHOU = ? WHERE REGJOB = ? AND REGHOU = ?"
            cursor.execute(sql_update, (hora_sistema, job, hora_actual))
            conn.commit()
            self.logger.info(f"{base} hora actualizada correctamente")

        finally:
            if cursor is not None:
                cursor.close()

        self.logger.info("FIN ajustar_hora_ejecucion")

    def get_item_promo(self, base):
        self.logger.info(f"INICIO get_item_promo | base={base}")
        conn = self.db.conectar()
        cursor = conn.cursor()
        query = f"""SELECT PBKPC{self.enie}, PBKZONE, PBKSKU, PBKTYPE, PBKSTATUS, PBKEFFDAT, PBKENDDAT,CHAR(PBKOLDPRC) AS PBKOLDPRC, CHAR(PBKNEWPRC) AS PBKNEWPRC FROM {base}.PCPRCBKP WHERE PBKTYPE='T' AND PBKSTATUS='A' """
        cursor.execute(query)
        columnas = [col[0] for col in cursor.description]
        datos = [dict(zip(columnas, fila)) for fila in cursor.fetchall()]
        cursor.close()
        self.logger.info(f"FIN get_item_promo | registros={len(datos)}")
        return datos

    def get_item_precio_permanente(self, base):
        self.logger.info(f"INICIO get_item_precio_permanente | base={base}")
        conn = self.db.conectar()
        cursor = conn.cursor()
        query = f"""SELECT PBKPC{self.enie}, PBKZONE, PBKSKU, PBKTYPE, PBKSTATUS, PBKEFFDAT, PBKENDDAT, CHAR(P.PBKOLDPRC) AS PBKOLDPRC, CHAR(P.PBKNEWPRC) AS PBKNEWPRC FROM {base}.PCPRCBKP P WHERE P.PBKTYPE='P' AND P.PBKSTATUS='A' """
        cursor.execute(query)
        columnas = [col[0] for col in cursor.description]
        datos = [dict(zip(columnas, fila)) for fila in cursor.fetchall()]
        cursor.close()
        self.logger.info(f"FIN get_item_precio_permanente | registros={len(datos)}")
        return datos

    def get_recepcion(self, base_ri):
        self.logger.info(f"INICIO obtener_items_primera_recepcion | base={base_ri}")
        conn = self.db.conectar()
        cursor = conn.cursor()
        query = f"""
            SELECT K.SKUSK
            FROM {base_ri}.KSKUP K
            WHERE K.SKUSK NOT IN (SELECT A.PBKSKU FROM {base_ri}.PCPRCBKP A)
              AND K.B34SK > 0
        """
        cursor.execute(query)
        datos = cursor.fetchall()
        cursor.close()
        self.logger.info(f"FIN obtener_items_primera_recepcion | registros={len(datos)}")
        return datos

    def actualizar_estado_cambio_precio(self, base):
        self.logger.info(f"INICIO actualizar_estado_cambio_precio | base={base}")
        conn = self.db.conectar()
        cursor = conn.cursor()

        try:
            hoy = datetime.now()
            tres_dias_atras = (hoy - timedelta(days=3)).strftime("%Y%m%d")
            un_dia_atras = (hoy - timedelta(days=1)).strftime("%Y%m%d")
            fecha_hoy = hoy.strftime("%Y%m%d")

            sql_no_kregp = f"""
                SELECT DISTINCT A.PBKPC{self.enie}, A.PBKZONE, A.PBKTYPE FROM {base}.PCPRCBKP A WHERE PBKPC{self.enie} NOT IN ( SELECT DISTINCT PBKPC{self.enie} FROM {base}.PCPRCBKP P WHERE P.PBKSKU IN ( SELECT REGSKU FROM {base}.KREGP K WHERE K.REGEFFDAT = '{fecha_hoy}' ) ) AND PBKSTATUS = 'A' """
            cursor.execute(sql_no_kregp)
            resultados = cursor.fetchall()
            if not resultados:
                self.logger.info("No hay cambios de precio disponibles")
                return None

            cambios_p = [registro for registro in resultados if registro[2] == "P"]
            cambios_t = [registro for registro in resultados if registro[2] == "T"]

            cambios_aplicados = []
            for cambio in [random.choice(cambios_p) if cambios_p else None, random.choice(cambios_t) if cambios_t else None]:
                if not cambio:
                    continue
                numero_cambio = cambio[0]
                sql_header = f"""
                    UPDATE {base}.PCHEADERP
                    SET PCHSTATUS='L', PCHCRTDAT=?, PCHLEAD=?, PCHEFFDAT=?, PCHENDDAT=?, PCHCHGDAT=?
                    WHERE PCHPC{self.enie}=?
                """
                cursor.execute(sql_header, (tres_dias_atras, un_dia_atras, fecha_hoy, fecha_hoy, fecha_hoy, numero_cambio))

                sql_detail = f"""
                    UPDATE {base}.PCPRCBKP
                    SET PBKSTATUS='L', PBKEFFDAT=?, PBKENDDAT=?, PBKCHGDAT=?
                    WHERE PBKPC{self.enie}=?
                """
                cursor.execute(sql_detail, (fecha_hoy, fecha_hoy, fecha_hoy, numero_cambio))
                sql_skus = f"""
                    SELECT DISTINCT PBKSKU
                    FROM {base}.PCPRCBKP
                    WHERE PBKPC{self.enie} = ?
                    FETCH FIRST 10 ROWS ONLY
                """
                cursor.execute(sql_skus, (numero_cambio,))
                skus = [str(fila[0]).strip() for fila in cursor.fetchall()]

                cambios_aplicados.append({
                    "cambio": numero_cambio,
                    "zona": str(cambio[1]).strip(),
                    "tipo": cambio[2],
                    "skus": skus,
                })

            conn.commit()
            self.logger.info(f"FIN actualizar_estado_cambio_precio | cambios={cambios_aplicados}")
            return cambios_aplicados
        finally:
            cursor.close()

    def borrar_kregp_items(self, base_ri):
        self.logger.info(f"INICIO borrar_kregp_items | base={base_ri}")
        conn = self.db.conectar()
        cursor = conn.cursor()
        try:
            cursor.execute(f"DELETE FROM {base_ri}.KREGL")
            conn.commit()
            self.logger.info(f"FIN borrar_kregp_items | registros_borrados={cursor.rowcount}")
            return True
        finally:
            cursor.close()

    def insert_kregp_items(self, base_ri, items):
        self.borrar_kregp_items(base_ri)
        self.logger.info(f"INICIO insert_kregp_items | base={base_ri} total_items={len(items)}")

        fecha_actual = datetime.today().strftime("%Y%m%d")
        fecha_manana = (datetime.today() + timedelta(days=1)).strftime("%Y%m%d")
        conn = self.db.conectar()
        cursor = conn.cursor()

        sql = f"""
            INSERT INTO {base_ri}.KREGL
            (REGPOSFLG, REGRCDOPT, REGST{self.enie}, REGSKU, REGO38, REGO39, REGO71, REGLEAD, REGEFFDAT,
             REGENDDAT, REGOLDPRC, REGNEWPRC, REGPCTYPE, REGPC{self.enie}, REGPCLVL, REGCRTUSR,
             REGCRTPGM, REGCRTDAT, REGCHGUSR, REGCHGPGM, REGCHGDAT)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        lote = []
        omitidos = 0
        for idx, item in enumerate(items, 1):
            sku_value = str(item.get("SKU") or item.get("PBKSKU") or "").strip()
            if not sku_value:
                self.logger.warning(f"Item {idx} sin SKU | item={item}")
                omitidos += 1
                continue
            lote.append((
                " ", "CHG", "0", sku_value, "", "", "", fecha_actual, fecha_manana, "99999999",
                "0", "0", "", "0", "0", "ELOPEZ", "MTSC01", fecha_manana, "ELOPEZ", "MTSC01", fecha_manana,
            ))

        self.logger.info(f"Lote preparado | registros_validos={len(lote)} omitidos={omitidos}")
        try:
            cursor.executemany(sql, lote)
            conn.commit()
            self.logger.info(f"FIN insert_kregp_items | insertados={len(lote)} omitidos={omitidos}")
            return True
        finally:
            cursor.close()

    def obtener_alias_prlkpabc(self, base_ri):
        self.logger.debug(f"INICIO obtener_alias_prlkpabc | base={base_ri}")
        conn = self.db.conectar()
        cursor = conn.cursor()
        query = f"""
            SELECT T.TABLE_NAME
            FROM SYSIBM.TABLES T
            WHERE T.TABLE_TYPE='ALIAS'
              AND T.TABLE_SCHEMA='{base_ri}'
              AND T.TABLE_NAME LIKE '%PRLKPABC%'
              AND T.TABLE_NAME NOT IN (
                  SELECT T2.TABLE_NAME
                  FROM SYSIBM.TABLES T2
                  WHERE T2.TABLE_TYPE='ALIAS'
                    AND T2.TABLE_SCHEMA='{base_ri}'
                    AND T2.TABLE_NAME LIKE '%PRLKPABCG%'
              )
        """
        
        cursor.execute(query)
        datos = cursor.fetchall()
        cursor.close()
        self.logger.debug(f"FIN obtener_alias_prlkpabc | total_alias={len(datos)}")
        return datos

    def obtener_prlkpabc(self, base, zona):
        self.logger.debug(f"INICIO obtener_prlkpabc | base={base}")
        zona = (zona or "").strip()
        if not zona:
            self.logger.debug(f"FIN obtener_prlkpabc | base={base} total_alias=0")
            return []

        conn = self.db.conectar()
        cursor = conn.cursor()
        query = f"""
            SELECT T.TABLE_NAME
            FROM SYSIBM.TABLES T
            WHERE T.TABLE_TYPE = 'ALIAS'
              AND T.TABLE_SCHEMA = '{base}'
              AND T.TABLE_NAME LIKE '%PRLKPABC%'
              AND T.TABLE_NAME NOT IN (
                  SELECT T2.TABLE_NAME
                  FROM SYSIBM.TABLES T2
                  WHERE T2.TABLE_TYPE = 'ALIAS'
                    AND T2.TABLE_SCHEMA = '{base}'
                    AND T2.TABLE_NAME LIKE '%PRLKPABCG%'
              )
              AND SUBSTRING(T.TABLE_NAME, 9, 3) IN (
                  SELECT TRIM(A.STRZS)
                  FROM {base}.KZSTP A
                  WHERE A.ZCOZS = ?
              )
        """
        try:
            cursor.execute(query, (zona,))
            datos = cursor.fetchall()
            aliases = [fila[0].strip() for fila in datos]
            self.logger.info(
                "FIN obtener_prlkpabc | base=%s zona=%s esquema_kzstp=%s total_alias=%s aliases=%s",
                base,
                zona,
                self.ZONE_MAPPING_SCHEMA,
                len(aliases),
                aliases,
            )
            return aliases
        finally:
            cursor.close()

    def ejecutar_alias(self, base_ri, alias_tabla):
        self.logger.debug(f"INICIO ejecutar_alias | base={base_ri} alias={alias_tabla}")
        conn = self.db.conectar()
        cursor = conn.cursor()
        try:
            sql = f"SELECT * FROM {base_ri}.{alias_tabla}"
            cursor.execute(sql)
            datos = cursor.fetchall()
            self.logger.debug(f"Alias ejecutado | alias={alias_tabla} registros={len(datos)}")
            return datos
        finally:
            cursor.close()

    def ejecutar_alias_por_sku(self, base_ri, alias_tabla, sku):
        self.logger.debug(f"INICIO ejecutar_alias_por_sku | base={base_ri} alias={alias_tabla} sku={sku}")
        conn = self.db.conectar()
        cursor = conn.cursor()
        try:
            sql = f"SELECT * FROM {base_ri}.{alias_tabla} P WHERE SUBSTR(CHAR(UPCABC), 2, 6) = ?"
            cursor.execute(sql, (str(sku).strip(),))
            datos = cursor.fetchall()
            self.logger.debug(f"Alias ejecutado por sku | alias={alias_tabla} sku={sku} registros={len(datos)}")
            return datos
        finally:
            cursor.close()

    def someter_programa(self, base_ri):
        parametro = base_ri[2:4]
        conn = self.db.conectar()
        cursor = conn.cursor()
        comando = f"SBMJOB CMD(CALL PGM(PROGRAMAS/REGMAIN) PARM('RI{parametro}PC')) JOB(RI{parametro}PC2) HOLD(*NO)"
        sql = "CALL QSYS2.QCMDEXC(?, ?)"
        try:
            cursor.execute(sql, (comando, decimal.Decimal(len(comando))))
            conn.commit()
            self.logger.info(f"comando={comando}")
            return True
        finally:
            cursor.close()

    def ejecutar_comando_cl(self, comando_cl):
        conn = self.db.conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("CALL QSYS2.QCMDEXC(?, ?)", (comando_cl, float(len(comando_cl))))
            conn.commit()
            self.logger.info(f"CL ejecutado | comando={comando_cl}")
            return True
        finally:
            cursor.close()

    def refrescar_actjob(self, espera_post_ejecucion=5):
        try:
            self.ejecutar_comando_cl("SBMJOB CMD(CALL PGM(ELOPEZ/ACTJOB)) JOB(ACTJOB)")
            time.sleep(espera_post_ejecucion)
            self.logger.info("ACTJOB ejecutado correctamente")
            return True
        except Exception as exc:
            self.logger.warning("No se pudo refrescar ACTJOB: %s", exc)
            return False


    def existe_job_activo(self, patron):
        conn = None
        cursor = None

        try:
            patron = str(patron or "").strip()
            if not patron:
                self.logger.warning("Patrón vacío en existe_job_activo")
                return False

            conn = self.db.conectar()

            for tabla in self.tablas_monitoreo_job:
                try:
                    tabla = str(tabla).strip()
                    if not tabla:
                        continue

                    cursor = conn.cursor()

                    sql = f""" SELECT QTEMPKILL FROM ELOPEZ.{tabla} WHERE QTEMPKILL LIKE ? FETCH FIRST 1 ROW ONLY """
                    params = (f"'%{patron}%'")

                    print(f"SQL: {sql} | params: {params}")

                    cursor.execute(sql, params)
                    fila = cursor.fetchone()

                    self.logger.debug( "Resultado monitoreo job | tabla=%s patron=%s fila=%s", tabla, patron, fila, )

                    if fila:
                        self.logger.info(
                            "Job activo | patron=%s tabla=%s registro=%s",
                            patron,
                            tabla,
                            fila,
                        )
                        return True
                    else:
                        self.logger.debug(
                            "No se encontró registro activo con LIKE en tabla %s | patron=%s",
                            tabla,
                            patron,
                        )

                except Exception as exc:
                    self.logger.warning(
                        "Error consultando monitoreo | tabla=%s patron=%s error=%s",
                        tabla, patron, exc
                    )
                finally:
                    try:
                        if cursor is not None:
                            cursor.close()
                    except Exception:
                        pass
                    cursor = None

            self.logger.info("Job no activo | patron=%s", patron)
            return False

        except Exception as exc:
            self.logger.error(
                "Error general validando job activo | patron=%s error=%s",
                patron, exc
            )
            return False

        finally:
            try:
                if conn is not None:
                    conn.close()
            except Exception:
                pass

    def esperar_fin_job_actjob( self, nombre_job, espera_inicial=10, espera_entre_refrescos=180, espera_post_refresco=5, max_espera_segundos=1800 ):
        self.logger.info(
            "INICIO esperar_fin_job_actjob | job=%s espera_inicial=%s espera_entre_refrescos=%s espera_post_refresco=%s max_espera=%s",
            nombre_job,
            espera_inicial,
            espera_entre_refrescos,
            espera_post_refresco,
            max_espera_segundos,
        )

        nombre_job = str(nombre_job or "").strip()
        if not nombre_job:
            raise ValueError("nombre_job es obligatorio")

        inicio = time.time()

        if espera_inicial > 0:
            time.sleep(espera_inicial)

        while True:
            transcurrido = int(time.time() - inicio)
            if transcurrido >= max_espera_segundos:
                raise TimeoutError(f"Tiempo de espera excedido para el job {nombre_job}")

            try:
                refrescado = self.refrescar_actjob(espera_post_ejecucion=espera_post_refresco)

                if not refrescado:
                    self.logger.warning(
                        "No se pudo refrescar ACTJOB para validar job=%s. Se reintentará en %s segundos.",
                        nombre_job,
                        espera_entre_refrescos,
                    )
                else:
                    activo = self.existe_job_activo(nombre_job)

                    self.logger.info(
                        "Validación job | job=%s activo=%s transcurrido=%s",
                        nombre_job,
                        activo,
                        transcurrido,
                    )

                    if not activo:
                        self.logger.info(
                            "FIN esperar_fin_job_actjob | job=%s finalizado=True tiempo_total=%s",
                            nombre_job,
                            transcurrido,
                        )
                        return {
                            "job": nombre_job,
                            "activo": False,
                            "tiempo_total_segundos": transcurrido,
                        }

            except Exception as exc:
                self.logger.warning(
                    "Error validando job %s: %s. Se reintentará en %s segundos.",
                    nombre_job,
                    exc,
                    espera_entre_refrescos,
                )

            transcurrido = int(time.time() - inicio)
            if transcurrido >= max_espera_segundos:
                raise TimeoutError(f"Tiempo de espera excedido para el job {nombre_job}")

            self.logger.info(
                "Job sigue activo | job=%s próximo_refresco_en=%s_segundos",
                nombre_job,
                espera_entre_refrescos,
            )
            time.sleep(espera_entre_refrescos)