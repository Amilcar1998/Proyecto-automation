import pyodbc
from conexion_config.conexion import ConexionAS400
from datetime import datetime, timedelta
import random 
import os
import shutil
from openpyxl import Workbook
import logging
from conexion_config.logging_config import get_logger
import win32com.client as win32
import atexit
import sys



# Elimina todos los __pycache__ al importar este módulo
for root, dirs, files in os.walk(os.path.dirname(os.path.abspath(__file__))):
    for d in dirs:
        if d == '__pycache__':
            shutil.rmtree(os.path.join(root, d), ignore_errors=True)

class GARANTIAS:
    def __init__(self):
        self.conexion_as400 = ConexionAS400()
        self.conexion = None
        
        # Obtener el logger ya configurado en main.py
        self.logger = get_logger(__file__)
        
        # Establecer conexión inicial
        self._verificar_y_reconectar()

    def _verificar_y_reconectar(self):
        """Verifica la conexión y reconecta si es necesario"""
        try:
            # Si no hay conexión, establecer una nueva
            if not self.conexion:
                self.logger.info("No hay conexión establecida, conectando...")
                self.conexion = self.conexion_as400.conectar()
                if self.conexion:
                    # Desactivar autocommit para mejor control de transacciones
                    self.conexion.autocommit = True
                    self.logger.info("[OK] Conexión establecida exitosamente (autocommit=True)")
                    return True
                else:
                    self.logger.error("[ERROR] No se pudo establecer conexión")
                    return False
            
            # Verificar si la conexión existente está activa
            cursor = None
            try:
                cursor = self.conexion.cursor()
                cursor.execute("SELECT 1 FROM SYSIBM.SYSDUMMY1")
                cursor.fetchone()
                self.logger.debug("Conexión existente verificada correctamente")
                return True
            finally:
                if cursor:
                    try:
                        cursor.close()
                    except:
                        pass
            
        except Exception as e:
            self.logger.warning(f"Conexión perdida o inválida ({e}), reconectando...")
            try:
                # Cerrar conexión existente si existe (usar el cierre del singleton si es posible)
                if self.conexion:
                    try:
                        self.conexion_as400.cerrar_conexion()
                    except Exception:
                        try:
                            self.conexion.close()
                        except Exception:
                            pass
                    finally:
                        self.conexion = None
            except:
                pass
            
            # Establecer nueva conexión
            self.conexion = self.conexion_as400.conectar()
            if self.conexion:
                # Desactivar autocommit para mejor control de transacciones
                self.conexion.autocommit = False
                self.logger.info("[OK] Reconexión exitosa (autocommit=False)")
                return True
            else:
                self.logger.error("[ERROR] Fallo en la reconexión")
                return False

    def _run_db_action(self, action, commit=False, *args, **kwargs):
        """
        Ejecuta una acción de base de datos segura.
        - Se asegura de estar conectado (reconecta si es necesario).
        - Provee un cursor a la función `action(cursor, *args, **kwargs)`.
        - Cierra el cursor y libera la conexión (usa el singleton) al terminar.
        - Si commit=True intentará commitear al finalizar la acción.
        Devuelve lo que retorne `action`, o None en caso de error.
        """
        # Verificar y reconectar antes de la acción
        if not self._verificar_y_reconectar():
            self.logger.error("No se pudo establecer conexión a la base de datos para ejecutar la acción")
            return None

        cursor = None
        try:
            cursor = self.conexion.cursor()
            result = action(cursor, *args, **kwargs)

            if commit and self.conexion:
                try:
                    self.conexion.commit()
                except Exception as e:
                    self.logger.error(f"Error al commitear la acción: {e}")

            return result

        except Exception as e:
            self.logger.error(f"Error durante la acción de BD: {type(e).__name__} - {e}")
            try:
                if self.conexion:
                    self.conexion.rollback()
            except Exception:
                pass
            return None

        finally:
            # Cerrar cursor
            try:
                if cursor:
                    cursor.close()
            except Exception as e:
                self.logger.debug(f"Error cerrando cursor: {e}")

            # Liberar la conexión a través del helper singleton cuando termine
            try:
                self.conexion_as400.cerrar_conexion()
            except Exception:
                try:
                    if self.conexion:
                        self.conexion.close()
                except Exception:
                    pass
            finally:
                self.conexion = None

    def get_items(self, base_RI):
        self.logger.info(f"Iniciando obtención de items para base {base_RI}")
        # Preparar la query (no modificar la lógica SQL)
        sql= (f"SELECT DISTINCT "
                f"PBKSKU AS SKU, "
                f"X.PBKTYPE AS TIPO_CAMBIO_PRECIO, "
                f"X.PBKSTATUS AS ESTADO, "
                f"A.CLSSK AS CLASE, "
                f"c.DES_CLA AS DESCRIPCION_CLASE, "
                f"UPC_CORP AS UPC, "
                f"DES_ESPA AS DESCRIPCION_ITEM, "
                f"X.PBKOLDPRC AS PRECIO_ANTERIOR, "
                f"PBKNEWPRC AS PRECIO_NUEVO, "
                f"'TEMPORAL' AS TIPO_REGISTRO, "
                f"PBKEFFDAT AS FECHA_INICIO, PBKENDDAT AS FECHA_FIN"
            f" FROM {base_RI}.KSKUP A"
            f" INNER JOIN {base_RI}.GA_EXT0003 ge ON A.CLSSK = ge.clase"
            f" LEFT JOIN {base_RI}.PCPRCBKP X ON A.SKUSK = X.PBKSKU"
            f" INNER JOIN {base_RI}.KSK2P kp ON KP.SKUK2 = A.SKUSK"
            f" INNER JOIN RI11DB.CONSULRP3 C ON A.SKUSK = C.SKU_CORP"
            f" WHERE A.SRLSK <> 'D' "
            f" AND KP.A03K2 > 0 "
            f" AND X.PBKTYPE = 'T' "
            f" AND PBKSTATUS = 'A' "
            f" AND X.PBKNEWPRC BETWEEN GE.RANGOA AND GE.RANGOB "
            f" AND PBKSKU NOT IN (SELECT SKU_CORP FROM RI11DB.CONSULRP3 A WHERE COD_DIV ='Y')"
            f" GROUP BY PBKSKU, X.PBKTYPE, X.PBKSTATUS, A.CLSSK, c.DES_CLA, UPC_CORP, DES_ESPA, X.PBKOLDPRC, PBKNEWPRC,PBKEFFDAT,PBKENDDAT"
            f" HAVING COUNT(*) = 1"
            f" UNION ALL"
            f" SELECT DISTINCT"
            f" PBKSKU AS SKU,"
            f" X.PBKTYPE AS TIPO_CAMBIO_PRECIO,"
            f" X.PBKSTATUS AS ESTADO,"
            f" A.CLSSK AS CLASE,"
            f" c.DES_CLA AS DESCRIPCION_CLASE,"
            f" UPC_CORP AS UPC,"
            f" DES_ESPA AS DESCRIPCION_ITEM,"
            f" X.PBKOLDPRC AS PRECIO_ANTERIOR,"
            f" PBKNEWPRC AS PRECIO_NUEVO,"
            f" 'PERMANENTE' AS TIPO_REGISTRO,"
            f" PBKEFFDAT AS FECHA_INICIO, PBKENDDAT AS FECHA_FIN"
            f" FROM {base_RI}.KSKUP A"
            f" INNER JOIN {base_RI}.GA_EXT0003 ge ON A.CLSSK = ge.clase"
            f" LEFT JOIN {base_RI}.PCPRCBKP X ON A.SKUSK = X.PBKSKU"
            f" INNER JOIN {base_RI}.KSK2P kp ON KP.SKUK2 = A.SKUSK"
            f" INNER JOIN RI11DB.CONSULRP3 C ON A.SKUSK = C.SKU_CORP"
            f" WHERE A.SRLSK <> 'D' "
            f" AND KP.A03K2 > 0 "
            f" AND X.PBKTYPE = 'P' "
            f" AND PBKSTATUS = 'A' "
            f" AND A.B34SK BETWEEN GE.RANGOA AND GE.RANGOB"
            f" AND PBKSKU NOT IN (SELECT SKU_CORP FROM RI11DB.CONSULRP3 A WHERE COD_DIV ='Y')"
            f" GROUP BY PBKSKU, X.PBKTYPE, X.PBKSTATUS, A.CLSSK, c.DES_CLA, UPC_CORP, DES_ESPA, X.PBKOLDPRC, PBKNEWPRC,PBKEFFDAT,PBKENDDAT"
            f" HAVING COUNT(*) = 1"
            f" UNION ALL"
            f" SELECT DISTINCT"
            f" A.SKUSK AS SKU,"
            f" 'N' AS TIPO_CAMBIO_PRECIO,"
            f" 'O' AS ESTADO,"
            f" A.CLSSK AS CLASE,"
            f" c.DES_CLA AS DESCRIPCION_CLASE,"
            f" UPC_CORP AS UPC,"
            f" DES_ESPA AS DESCRIPCION_ITEM,"
            f" A.B34SK AS PRECIO_ANTERIOR,"
            f" A.B34SK AS PRECIO_NUEVO,"
            f" 'PRIMERA_RECEPCION' AS TIPO_REGISTRO,"
            f" '20250101' AS FECHA_INICIO,'99999999' AS FECHA_FIN"
            f" FROM {base_RI}.KSKUP A"
            f" INNER JOIN {base_RI}.GA_EXT0003 ge ON A.CLSSK = ge.clase"
            f" LEFT JOIN {base_RI}.PCPRCBKP X ON A.SKUSK = X.PBKSKU"
            f" INNER JOIN {base_RI}.KSK2P kp ON KP.SKUK2 = A.SKUSK"
            f" INNER JOIN RI11DB.CONSULRP3 C ON A.SKUSK = C.SKU_CORP"
            f" WHERE A.SRLSK <> 'D' "
            f" AND KP.A03K2 > 0 "
            f" AND B34SK > 0 "
            f" AND A.SKUSK NOT IN (SELECT DISTINCT PBKSKU FROM {base_RI}.PCPRCBKP A) "
            f" AND A.SKUSK NOT IN (SELECT SKU_CORP FROM RI11DB.CONSULRP3 A WHERE COD_DIV ='Y') "
            f" AND A.B34SK BETWEEN GE.RANGOA AND GE.RANGOB"
            f" GROUP BY A.CLSSK, c.DES_CLA, UPC_CORP, DES_ESPA, X.PBKOLDPRC, PBKNEWPRC,B34SK,SKUSK")

        # Acción que ejecuta la query y devuelve resultados
        def _action(cursor):
            cursor.execute(sql)
            columnas = [col[0] for col in cursor.description]
            return [dict(zip(columnas, fila)) for fila in cursor.fetchall()]

        resultados = self._run_db_action(_action, commit=False)
        if resultados is None:
            return []
        self.logger.info(f"Obtenidos {len(resultados)} items de la base {base_RI}")
        return resultados

    def liberar_kregp(self):
        """Libera el objeto KREGP asegurando que no quede retenido."""
        try:
            if self.conexion:
                self.logger.info("Liberando objeto KREGP con ROLLBACK...")
                self.conexion.rollback()  # Asegura que cualquier transacción pendiente se deshaga
                self.logger.info("Objeto KREGP liberado correctamente.")
            else:
                self.logger.warning("No hay conexión activa para liberar KREGP.")
        except Exception as e:
            self.logger.error(f"Error al liberar KREGP: {e}")

    def copiar_objeto(self, base):
        self.logger.info(f"Iniciando copia de ELOPEZ.KREGP hacia {base}.KREGP")

        self.forzar_liberacion_objetos()

        comando = f"CPYF FROMFILE(ELOPEZ/KREGP) TOFILE({base}/KREGP) MBROPT(*REPLACE)"

        def _accion(cursor):
            self.logger.info(f"Ejecutando comando: {comando}")
            longitud = len(comando)
            cursor.execute("CALL QSYS2.QCMDEXC(?, ?)", (comando, float(longitud)))
            return True

        resultado = self._run_db_action(_accion, commit=True)

        if resultado:
            self.logger.info(f"CPYF ejecutado correctamente hacia {base}/KREGP")
            print(f"CPYF ejecutado correctamente hacia {base}/KREGP")
            return True

        self.logger.error(f"Error ejecutando CPYF hacia {base}/KREGP")
        print(f"\nError ejecutando CPYF hacia {base}/KREGP")
        return False

    def forzar_liberacion_objetos(self):
        """Fuerza la liberación de locks y el cierre real de la conexión."""
        try:
            if not self.conexion:
                self.logger.info("No hay conexión activa. No hay objetos por liberar.")
                return True

            self.logger.info("Forzando liberación de objetos bloqueados...")

            try:
                self.conexion.rollback()
                self.logger.info("ROLLBACK ejecutado correctamente.")
            except Exception as e:
                self.logger.warning(f"No fue posible ejecutar ROLLBACK: {e}")

            try:
                cursor = self.conexion.cursor()
                cursor.execute("COMMIT")
                cursor.close()
                self.logger.info("COMMIT ejecutado para liberar recursos del job.")
            except Exception as e:
                self.logger.warning(f"No fue posible ejecutar COMMIT de liberación: {e}")

            try:
                self.conexion_as400.cerrar_conexion()
            except Exception:
                try:
                    self.conexion.close()
                except Exception as e:
                    self.logger.warning(f"No fue posible cerrar conexión directamente: {e}")

            self.conexion = None
            self.logger.info("Conexión cerrada y objetos liberados correctamente.")
            return True

        except Exception as e:
            self.logger.error(f"Error al forzar liberación de objetos: {e}")
            return False


    def cerrar_conexion(self):
        """Cierra la conexión actual a la base de datos."""
        try:
            if self.conexion:
                try:
                    # Preferir el cierre centralizado del singleton
                    self.conexion_as400.cerrar_conexion()
                except Exception:
                    try:
                        self.conexion.close()
                    except Exception as e:
                        self.logger.error(f"Error cerrando conexión directamente: {e}")
                finally:
                    self.conexion = None
                    self.logger.info("Conexión cerrada correctamente.")
            else:
                self.logger.warning("No hay conexión activa para cerrar.")
        except Exception as e:
            self.logger.error(f"Error al cerrar la conexión: {e}")

    def __del__(self):
        """Destructor que asegura que la conexión se cierre al destruir la instancia"""
        try:
            if hasattr(self, 'conexion') and self.conexion:
                try:
                    self.conexion_as400.cerrar_conexion()
                except Exception:
                    try:
                        self.conexion.close()
                    except Exception:
                        pass
        except:
            pass

    def delete_kregp(self, base_RI):
        self.logger.info(f"Iniciando eliminación de registros en {base_RI}.KREGP")
        sql = f"DELETE FROM {base_RI}.KREGP"

        def _action(cursor):
            self.logger.info(f"Ejecutando DELETE sobre {base_RI}.KREGP")
            cursor.execute(sql)
            return cursor.rowcount

        registros_eliminados = self._run_db_action(_action, commit=True)

        if registros_eliminados is None:
            self.logger.error(f"Error al eliminar datos de {base_RI}.KREGP")
            print(f"Error al eliminar datos de {base_RI}.KREGP")
            return False

        self.logger.info(f"Eliminados {registros_eliminados} registros de {base_RI}.KREGP")
        print(f"Datos eliminados de {base_RI}.KREGP correctamente.")
        return True

    
    

    def insertar_items_kregp(self, base_ri, items):
        total_items = len(items)
        self.logger.info(f"INICIO insertar_items_kregp | base={base_ri} total_items={total_items}")

        fecha_actual = datetime.today().strftime("%Y%m%d")
        fecha_manana = (datetime.today() + timedelta(days=1)).strftime("%Y%m%d")
        tamano_bloque = 100

        def _mostrar_loader(insertados, omitidos, total, sku_actual=""):
            procesados = insertados + omitidos
            restantes = total - procesados
            porcentaje = (procesados / total * 100) if total > 0 else 100

            mensaje = (
                f"\rInsertando en {base_ri}.KREGP | "
                f"Procesados: {procesados}/{total} | "
                f"Insertados: {insertados} | "
                f"Omitidos: {omitidos} | "
                f"Restantes: {restantes} | "
                f"{porcentaje:6.2f}% | "
                f"SKU actual: {sku_actual}"
            )

            sys.stdout.write(mensaje)
            sys.stdout.flush()

        def _accion(cursor):
            registros_insertados = 0
            registros_omitidos = 0

            sql = f"""
            INSERT INTO {base_ri}.KREGP (
                REGPOSFLG, REGRCDOPT, REGSTÑ, REGSKU, REGO38, REGO39, REGO71,
                REGLEAD, REGEFFDAT, REGENDDAT, REGOLDPRC, REGNEWPRC,
                REGPCTYPE, REGPCÑ, REGPCLVL,
                REGCRTUSR, REGCRTPGM, REGCRTDAT,
                REGCHGUSR, REGCHGPGM, REGCHGDAT
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """

            lote = []
            ultimo_sku = ""

            for indice, item in enumerate(items, 1):
                try:
                    sku_valor = item.get("SKU") if "SKU" in item else item.get("PBKSKU")
                    sku_valor = str(sku_valor).strip() if sku_valor is not None else ""

                    if not sku_valor:
                        self.logger.warning(f"Item {indice} sin SKU | item={item}")
                        registros_omitidos += 1
                        _mostrar_loader(registros_insertados, registros_omitidos, total_items, "")
                        continue

                    ultimo_sku = sku_valor

                    lote.append((
                        " ",
                        "CHG",
                        "0",
                        sku_valor,
                        "",
                        "",
                        "",
                        fecha_actual,
                        fecha_manana,
                        "99999999",
                        "0",
                        "0",
                        "",
                        "0",
                        "0",
                        "ELOPEZ",
                        "MTSC01",
                        fecha_manana,
                        "ELOPEZ",
                        "MTSC01",
                        fecha_manana
                    ))
                    
                    # Imprimir el progreso en tiempo real registro por registro
                    _mostrar_loader(registros_insertados + len(lote), registros_omitidos, total_items, ultimo_sku)

                    if len(lote) >= tamano_bloque:
                        cursor.executemany(sql, lote)
                        registros_insertados += len(lote)
                        lote.clear()

                except Exception as e:
                    self.logger.warning(f"Error preparando SKU | item={item} | error={e}")
                    registros_omitidos += 1
                    _mostrar_loader(registros_insertados, registros_omitidos, total_items, "")
                    continue

            if lote:
                cursor.executemany(sql, lote)
                registros_insertados += len(lote)
                lote.clear()
                _mostrar_loader(registros_insertados, registros_omitidos, total_items, ultimo_sku)

            sys.stdout.write("\n")
            sys.stdout.flush()

            self.logger.info(
                f"INSERT completado en ELOPEZ.KREGP | insertados={registros_insertados} omitidos={registros_omitidos}"
            )

            return registros_insertados, registros_omitidos

        resultado = self._run_db_action(_accion, commit=True)

        if resultado is None:
            self.logger.error(f"ERROR insertar_items_kregp | base={base_ri}")
            return False

        registros_insertados, registros_omitidos = resultado

        self.logger.info(
            f"FIN insertar_items_kregp | base={base_ri} insertados={registros_insertados} omitidos={registros_omitidos}"
        )

        return registros_insertados > 0




    def importar_excel(self, data, filename="garantias_export.xlsx", logger=None):
        self.logger.info(f"Iniciando exportación a Excel: {filename}")
        
        try:
            wb = Workbook()
            ws = wb.active
            if not data:
                self.logger.warning("No hay datos para exportar a Excel")
                print("No hay datos para exportar.")
                return False
                
            # Escribir encabezados
            headers = list(data[0].keys())
            ws.append(headers)
            self.logger.info(f"Encabezados escritos: {headers}")
            
            # Escribir filas
            for row in data:
                ws.append([row.get(col, "") for col in headers])
            
            wb.save(filename)
            self.logger.info(f"Excel exportado exitosamente: {filename} ({len(data)} registros)")
            print(f"Datos exportados exitosamente a {filename}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al exportar a Excel {filename}: {e}")
            print(f"Error al exportar a Excel: {e}")
            return False

    def leer_configuracion_cantidad(self):
        """Lee la cantidad máxima de registros desde archivos_config/cantidad_registros.txt"""
        try:
            base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
            config_dir = os.path.join(base_dir, 'archivos_config')
            cantidad_path = os.path.join(config_dir, 'cantidad_registros.txt')
            
            print(f"[DEBUG] Buscando archivo de configuración en: {cantidad_path}")
            
            if not os.path.exists(cantidad_path):
                # Crear archivo con valor por defecto
                os.makedirs(config_dir, exist_ok=True)
                with open(cantidad_path, 'w', encoding='utf-8') as f:
                    f.write("# Cantidad máxima de registros sin promoción a procesar\n")
                    f.write("# Valor por defecto: 200\n")
                    f.write("# Cambia este número para ajustar la cantidad\n")
                    f.write("200\n")
                print(f"[OK] Archivo de configuración creado en: {cantidad_path}")
                return 200
            
            print(f"[OK] Archivo de configuración encontrado: {cantidad_path}")
            with open(cantidad_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        try:
                            cantidad = int(line)
                            if cantidad > 0:
                                print(f"[INFO] Cantidad leída desde configuración: {cantidad}")
                                return cantidad
                        except ValueError:
                            continue
            
            # Si no se encuentra valor válido, usar 200 por defecto
            print("[WARN] No se encontró valor válido, usando 200 por defecto")
            return 200
            
        except Exception as e:
            print(f"[ERROR] Error leyendo configuración de cantidad: {e}")
            return 200
        
    def copiar_objeto(self,base):
        comando = f"CPYF FROMFILE(ELOPEZ/KREGP) TOFILE({base}/KREGP) MBROPT(*REPLACE)"

        def _accion(cursor):
            longitud = len(comando)
            cursor.execute("CALL QSYS2.QCMDEXC(?, ?)", (comando, float(longitud)))
            return True

        resultado = self._run_db_action(_accion, commit=True)

        if resultado:
            print("CPYF ejecutado correctamente hacia RI11DB/KREGP")
            return True

        print("\nError ejecutando CPYF hacia RI11DB/KREGP")
        return False

    def main_garantias(self, base_RI, logger=None):
        self.logger.info(f"=== INICIANDO PROCESO PRINCIPAL DE GARANTÍAS PARA {base_RI} ===")

        try:
            self.logger.info("PASO 1: Obteniendo items de la base de datos")
            data = self.get_items(base_RI)

            self.logger.info("PASO 2: Eliminando registros existentes de ELOPEZ.KREGP")
            eliminacion_kregp = self.delete_kregp(base_RI)
            if not eliminacion_kregp:
                self.logger.error("No fue posible limpiar ELOPEZ.KREGP")
                return False

            self.logger.info("PASO 3: Separando items con y sin promoción")
            items_con_promocion = []
            items_sin_promocion_todos = []
            items_primera_recepcion = []

            for item in data:
                if item["TIPO_CAMBIO_PRECIO"] == 'T' and item["ESTADO"] == 'A':
                    items_con_promocion.append(item)
                elif item["TIPO_CAMBIO_PRECIO"] == 'N' and item["ESTADO"] == 'O':
                    items_primera_recepcion.append(item)
                else:
                    items_sin_promocion_todos.append(item)

            self.logger.info(f"Items en primera recepción: {len(items_primera_recepcion)}")
            self.logger.info(f"Items con promoción temporal: {len(items_con_promocion)}")
            self.logger.info(f"Items sin promoción temporal: {len(items_sin_promocion_todos)}")

            self.logger.info("PASO 4: Leyendo configuración de cantidad")
            cantidad_configurada = self.leer_configuracion_cantidad()

            items_sin_promocion = random.sample(
                items_sin_promocion_todos,
                min(cantidad_configurada, len(items_sin_promocion_todos))
            )

            items_primera_recepcion = random.sample(
                items_primera_recepcion,
                min(cantidad_configurada, len(items_primera_recepcion))
            )

            self.logger.info(f"Items sin promoción seleccionados aleatoriamente: {len(items_sin_promocion)}")
            print(f"Items con promoción: {len(items_con_promocion)}")
            print(f"Items sin promoción (aleatorios): {len(items_sin_promocion)}")
            print(f"Items en primera recepción: {len(items_primera_recepcion)}")

            self.logger.info("PASO 5: Unificando items para procesamiento")
            homologacion = items_con_promocion + items_sin_promocion + items_primera_recepcion
            self.logger.info(f"Total de items a procesar: {len(homologacion)}")

            self.logger.info(f"PASO 6: Insertando items en {base_RI}.KREGP")
            resultado_insert = self.insertar_items_kregp(base_RI, homologacion)
            
            # PASO CLAVE: CERRAR CONEXIÓN PARA LIBERAR KREGP Y QUE EL PROGRAMA CL NO SE RETENGA
            self.cerrar_conexion()
            
            if not resultado_insert:
                self.logger.error(f"Falló la inserción en {base_RI}.KREGP")
                return False

            self.logger.info("PASO 7: Omitido (Inserción directa sin CPYF)")

            self.logger.info("PASO 8: Generando archivo Excel")
            fecha_hoy = datetime.today().strftime('%Y-%m-%d')
            base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
            archivos_dir = os.path.join(base_dir, 'Archivos')
            os.makedirs(archivos_dir, exist_ok=True)

            excel_file = os.path.join(archivos_dir, f'garantias_export_{base_RI}_{fecha_hoy}.xlsx')
            self.importar_excel(homologacion, filename=excel_file, logger=logger)

            self.logger.info("PASO 9: Liberando objetos al final del proceso")
            self.forzar_liberacion_objetos()

            self.logger.info(f"=== PROCESO COMPLETADO EXITOSAMENTE PARA {base_RI} ===")
            self.logger.info(f"Archivo Excel generado: {excel_file}")

            return {
                'base': base_RI,
                'excel_file': excel_file,
                'cantidad_items': len(homologacion)
            }

        except Exception as e:
            self.logger.error(f"=== ERROR EN PROCESO PRINCIPAL PARA {base_RI} ===")
            self.logger.error(f"Error: {type(e).__name__} - {e}")

            try:
                self.forzar_liberacion_objetos()
            except Exception as force_error:
                self.logger.error(f"Error al forzar liberación en main_garantias: {force_error}")

            print(f"Error al procesar los items: {type(e).__name__} - {e}")
            return False



def leer_destinatarios_config():
    """
    Lee los destinatarios desde archivos_config/destinatarios.txt, ignorando líneas que empiezan con # o están vacías.
    Devuelve una lista de correos limpios.
    """
    import json
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    config_dir = os.path.join(base_dir, 'archivos_config')
    destinatarios_path = os.path.join(config_dir, 'destinatarios_flujos.json')
    if not os.path.exists(destinatarios_path):
        return []
    
    try:
        with open(destinatarios_path, 'r', encoding='utf-8') as f:
            config_flujos = json.load(f)
            return config_flujos.get('garantias', [])
    except Exception:
        return []

def enviar_correo_outlook(archivo_excel, destinatarios, asunto, cuerpo, logger=None):
    try:
        # Si destinatarios es None o string, leer desde config ignorando líneas con #
        if destinatarios is None or isinstance(destinatarios, str):
            destinatarios = leer_destinatarios_config()
            if logger:
                logger.info(f"Destinatarios cargados de configuración: {len(destinatarios)}")
        
        # Si no se pasa cuerpo, leerlo de archivos_config/cuerpo.txt
        if cuerpo is None:
            base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
            config_dir = os.path.join(base_dir, 'archivos_config')
            cuerpo_path = os.path.join(config_dir, 'cuerpo.txt')
            if not os.path.exists(cuerpo_path):
                # Crear archivo con formato de ejemplo, fondo radial y explicación del proceso
                ejemplo = (
                    "<html>\n"
                    "<body style='margin:0;padding:0;background:radial-gradient(circle at 50% 20%, #eaf1fb 0%, #c7d7ee 100%);font-family:Calibri,sans-serif;font-size:14px;color:#222;'>\n"
                    "<div style='background:rgba(255,255,255,0.85);padding:32px 24px 24px 24px;border-radius:12px;box-shadow:0 2px 8px #b0c4de;'>\n"
                    "<h2 style='color:#2a5699;margin-top:0;'>Reporte de Garantías</h2>\n"
                    "<p>Estimado equipo,</p>\n"
                    "<p>Adjunto encontrará el reporte de garantías generado automáticamente.</p>\n"
                    "<p style='margin-bottom:16px;'>\n"
                    "Este proceso realiza lo siguiente:<br>\n"
                    "<ol style='margin:8px 0 16px 24px;'>\n"
                    "  <li>Se Conecta a la base de datos AS400 y obtiene los ítems de garantías activos.</li>\n"
                    "  <li>Filtra y selecciona los registros relevantes, separando los que tienen promoción los items que tienen <span style='color:#e74c3c;font-weight:bold;'>promociones temporales</span> se extraen todos y los que tienen <span style='color:#3498db;font-weight:bold;'>promociones permanentes</span> solo se extraen una parte del global</li>\n"
                    "  <li>Inserta los ítems seleccionados en la tabla <span style='color:#2ecc71;font-weight:bold;'>KREGP</span> para su procesamiento.</li>\n"
                    "  <li>Ejecuta los programas correspondientes al flujo de garantias por porcentaje</li>\n"
                    "  <li>Genera un archivo Excel con el detalle de los ítems procesados para cada base.</li>\n"
                    "  <li>Envía este correo con todos los reportes adjuntos a los destinatarios configurados.</li>\n"
                    "  <li>Normalmente integración tiene el proceso de automático de procesamiento de los archivos itemImport pero si se requiere que se procesen de inmediato .</li>\n"
                    "  <li>Puede solicitar la ayuda a soporteIntegracion. o dejando un ticket en el portal</li>\n"
                    "  <a href='https://grupounicomer.atlassian.net/servicedesk/customer/portal/53' style='color:#3498db;'>Portal de Soporte</a> \n"
                    "</ol>\n"
                    "</p>\n"
                    "<p>Saludos!!!</p>\n"
                    "<p><b>Eliseo Amilcar López</b></p>\n"
                    "</div>\n"
                    "</body>\n"
                    "</html>"
                )
                with open(cuerpo_path, 'w', encoding='utf-8') as f:
                    f.write(ejemplo)
                cuerpo = ejemplo
            else:
                with open(cuerpo_path, 'r', encoding='utf-8') as f:
                    cuerpo = f.read()

        print("Iniciando proceso de envío de correo...")
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = ";".join(destinatarios)
        mail.Subject = asunto
        
        # Obtener ruta de la firma
        recursos_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Recursos')
        os.makedirs(recursos_dir, exist_ok=True)  # Crear directorio si no existe
        firma_path = os.path.join(recursos_dir, 'Firma.jpg')
        
        print(f"Buscando firma en: {firma_path}")
        
        # Permitir saltos de línea y estilos en el cuerpo
        # Si el cuerpo ya es HTML, no lo modificamos; si es texto plano, lo convertimos
        if '<' in cuerpo and '>' in cuerpo:
            cuerpo_html = cuerpo
        else:
            # Convertir saltos de línea a <br> y aplicar estilos básicos
            cuerpo_html = f"<div style='font-family:Calibri,sans-serif; font-size:14px; color:#222;'>"
            for linea in cuerpo.split('\n'):
                cuerpo_html += f"<p style='margin:0 0 8px 0;'>{linea.strip()}</p>"
            cuerpo_html += "</div>"
        
        # Verificar si existe la firma
        if os.path.exists(firma_path):
            try:
                # Para imágenes JPG, incluirlas como adjuntos y referenciarlas en el HTML
                img_id = "firma.jpg"
                mail.Attachments.Add(firma_path, 0, 0, img_id)
                firma_html = f"""
                <hr style='border: 1px solid #dddddd;'>
                <div style='font-size: 11px; color: #777777;'>
                    <img src="cid:{img_id}" alt="Firma" />
                </div>
                """
                print(f"Firma encontrada y adjuntada con ID: {img_id}")
            except Exception as e:
                print(f"Error al adjuntar imagen de firma: {e}")
                firma_html = """
                <hr style='border: 1px solid #dddddd;'>
                <div style='font-size: 11px; color: #777777;'>
                    <p>Generado automáticamente por el Sistema de Automatización</p>
                </div>
                """
        else:
            print(f"Archivo de firma no encontrado en {firma_path}, usando firma por defecto")
            firma_html = """
            <hr style='border: 1px solid #dddddd;'>
            <div style='font-size: 11px; color: #777777;'>
                <p>Generado automáticamente por el Sistema de Automatización</p>
            </div>
            """
        
        # Agregar la firma al cuerpo
        mail.HTMLBody = cuerpo_html + firma_html
        
        # Adjuntar el archivo Excel
        if os.path.exists(archivo_excel):
            mail.Attachments.Add(archivo_excel)
            print(f"Archivo Excel adjuntado: {archivo_excel}")
        else:
            print(f"ADVERTENCIA: El archivo Excel no existe: {archivo_excel}")
        
        print(f"Enviando correo a {len(destinatarios)} destinatarios...")
        mail.Send()
        print(f"[OK] Correo enviado exitosamente a: {', '.join(destinatarios)}")
        
        if logger:
            logger.info(f"Correo enviado a: {destinatarios}")
        return True
    except Exception as e:
        print(f"[ERROR] ERROR al enviar correo: {e}")
        if logger:
            logger.error(f"Error enviando correo: {e}")
        return False

def main_garantias_envio(logger=None):
    print("main_garantias_envio ya no envía correos ni genera Excel. El envío de correo debe hacerse al final del flujo completo.")
    return


# Asegurar cierre de cualquier conexión ODBC singleton al finalizar el proceso
def _cerrar_conexiones_al_terminar():
    try:
        conn_helper = ConexionAS400()
        conn_helper.cerrar_conexion()
        print("Cierre de conexiones ODBC ejecutado al terminar.")
    except Exception:
        pass

atexit.register(_cerrar_conexiones_al_terminar)