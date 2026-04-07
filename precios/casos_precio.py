import datetime
import re

class Configuracion:

    def __init__(self, conexion_helper, bases):
        self.db = conexion_helper
        self.bases = bases
        self.logger = logging.getLogger("Configuracion")

    def ajustar_hora_ejecucion(self):
        self.logger.info("INICIO ajustar_hora_ejecucion")
        try:
            hora_sistema = datetime.datetime.now().strftime("%H:%M:%S")
            self.logger.info(f"Hora sistema detectada {hora_sistema}")
            conn = self.db.conectar()
            cursor = conn.cursor()

            for base in self.bases:
                self.logger.info(f"Procesando base {base}")

                try:

                    cursor.execute(f"""SELECT REGHOU FROM {base}.REGHOU FETCH FIRST 1 ROW ONLY""")
                    row = cursor.fetchone()

                    if not row:
                        self.logger.warning(f"{base} REGHOU sin registros")
                        continue

                    hora_actual = str(row[0])

                    if hora_actual[:2] == hora_sistema[:2]:
                        self.logger.info(f"{base} hora ya coincide con sistema")
                        continue

                    cursor.execute(
                        f"""UPDATE {base}.REGHOU SET REGHOU=? WHERE RRN(REGHOU)=1""",
                        (hora_sistema,)
                    )

                    conn.commit()
                    self.logger.info(f"{base} hora actualizada {hora_sistema}")

                except Exception as e:
                    self.logger.error(f"Error procesando base {base} {str(e)}")

            cursor.close()

        except Exception as e:

            self.logger.error(f"Error ajustar_hora_ejecucion {str(e)}")
            raise

        self.logger.info("FIN ajustar_hora_ejecucion")
