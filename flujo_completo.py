import os
import sys
import time
from datetime import datetime

# Lightweight fallback mocks if real modules are not available
try:
    from scripts_as400.ejecutor_cl import EjecutorCL
except Exception:
    class EjecutorCL:
        def __init__(self):
            self.conexion = True

        def ejecutar(self, cmd):
            print(f"[MockEjecutorCL] ejecutar: {cmd}")
            return True

        def close(self):
            self.conexion = None


try:
    from validacion_db.garantias import GARANTIAS
except Exception:
    class GARANTIAS:
        def __init__(self):
            self.conexion = True

        def main_garantias(self, dbname):
            print(f"[MockGARANTIAS] main_garantias called for {dbname}")
            out = os.path.join(os.path.dirname(__file__), f"reporte_{dbname}.xlsx")
            with open(out, 'w', encoding='utf-8') as f:
                f.write('col1,col2\n1,2\n')
            return {'excel_file': out}

        def rollback(self):
            self.conexion = None

        def close(self):
            self.conexion = None


try:
    import win32com.client as win32
except Exception:
    class _MockMail:
        def __init__(self):
            self.To = ''
            self.Subject = ''
            self.HTMLBody = ''

            class _Attachments:
                def Add(self, p):
                    print(f"[MockMail] Attach: {p}")

            self.Attachments = _Attachments()

        def Send(self):
            print('[MockMail] Send called')

    class _MockOutlook:
        def CreateItem(self, _):
            return _MockMail()

    class _Win32Mock:
        @staticmethod
        def Dispatch(name):
            return _MockOutlook()

    win32 = _Win32Mock()


def flujo_completo(logger=None):
    fecha_hoy = datetime.today().strftime('%Y-%m-%d')

    config_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'archivos_config'))
    os.makedirs(config_dir, exist_ok=True)

    bases_path = os.path.join(config_dir, 'bases.txt')
    if not os.path.exists(bases_path):
        with open(bases_path, 'w', encoding='utf-8') as f:
            f.write('# Lista de bases habilitadas para el flujo de garantías\n13\n14\n')

    bases = []
    with open(bases_path, 'r', encoding='utf-8') as f:
        for line in f:
            base = line.strip()
            if base and not base.startswith('#'):
                bases.append(base)

    if not bases:
        print('No hay bases habilitadas en bases.txt')
        return

    excel_files = []
    cl = None

    def inicializar_cl():
        nonlocal cl
        try:
            if cl:
                try:
                    cl.close()
                except Exception:
                    pass

            cl = EjecutorCL()

            if logger:
                logger.info('EjecutorCL inicializado correctamente')

            return cl

        except Exception as e:
            cl = None
            if logger:
                logger.exception(f'No se pudo inicializar EjecutorCL: {e}')
            return None

    def asegurar_conexion_cl():
        """
        No usa atributos internos como cl.conexion.
        Solo garantiza que exista una instancia utilizable de EjecutorCL.
        """
        nonlocal cl

        try:
            if cl is None:
                if logger:
                    logger.warning('EjecutorCL no inicializado. Intentando crear instancia...')
                cl = EjecutorCL()

            return cl

        except Exception as e:
            if logger:
                logger.exception(f'Error asegurando EjecutorCL: {e}')
            try:
                cl = EjecutorCL()
                return cl
            except Exception:
                cl = None
                return None

    inicializar_cl()

    def obtener_comando_por_base(base):
        if base == '11':
            return 'CALL PGM(SALPGM63/PCUNESCL)'
        elif base == '12':
            return 'CALL PGM(SALPGM63/PCUNGUCL)'
        elif base == '13':
            return 'CALL PGM(SALPGM63/PCUNHNCL)'
        elif base == '14':
            return 'CALL PGM(SALPGM63/PCUNNICL)'
        else:
            return None

    def esperar_fin_job(nombre_job, espera_segundos=300, espera_inicial=100, logger=None):
        """
        Espera hasta que el job ya no aparezca activo.
        Muestra un loader en consola durante las esperas.
        """

        def mostrar_loader(mensaje_base, segundos):
            frames = ["|", "/", "-", "\\"]
            inicio = time.time()
            i = 0

            while (time.time() - inicio) < segundos:
                transcurrido = int(time.time() - inicio)
                restante = max(0, segundos - transcurrido)
                frame = frames[i % len(frames)]
                texto = f"\r{frame} {mensaje_base} | restante: {restante:>3}s "
                sys.stdout.write(texto)
                sys.stdout.flush()
                time.sleep(0.2)
                i += 1

            sys.stdout.write("\r" + " " * 120 + "\r")
            sys.stdout.flush()

        if logger:
            logger.info(f"Esperando arranque del job {nombre_job} por {espera_inicial} segundos")

        print(f"Iniciando monitoreo del job {nombre_job}")

        if espera_inicial > 0:
            mostrar_loader(f"Esperando arranque del job {nombre_job}", espera_inicial)

        while True:
            try:
                conexion_cl = asegurar_conexion_cl()
                if not conexion_cl:
                    raise Exception("No fue posible obtener instancia de EjecutorCL para validar jobs")

                if logger:
                    logger.info("Limpiando spool y ejecutando monitoreo ACTJOB para %s", nombre_job)

                conexion_cl.ejecutar("DLTSPLF FILE(*SELECT) SELECT(ELOPEZ)")
                conexion_cl.ejecutar("SBMJOB CMD(CALL PGM(ELOPEZ/ACTJOB)) JOB(ACTJOB)")

                mostrar_loader(f"Esperando generación de ACTJOB para {nombre_job}", 5)

                activo = job_sigue_activo(nombre_job)

            except Exception as e:
                mensaje = f"Error validando job {nombre_job}: {e}. Reintentando en {espera_segundos} segundos."
                print(mensaje)
                if logger:
                    logger.warning(mensaje)
                mostrar_loader(f"Reintentando validación de {nombre_job}", espera_segundos)
                continue

            if activo:
                mensaje = f"El job {nombre_job} sigue activo. Se volverá a validar."
                print(mensaje)
                if logger:
                    logger.info(mensaje)

                mostrar_loader(f"En espera para revalidar {nombre_job}", espera_segundos)
            else:
                mensaje = f"El job {nombre_job} ya no está activo. Continuando con el siguiente proceso."
                print(mensaje)
                if logger:
                    logger.info(mensaje)
                break
    
    for base in bases:
        print('Procesando base: ' + base)

        try:
            garantias = GARANTIAS()
            dbname = f'RI{base}DB'

            if logger:
                logger.info('Iniciando ' + dbname)

            resultado = garantias.main_garantias(dbname)

            try:
                if hasattr(garantias, 'conexion') and garantias.conexion:
                    garantias.close()
            except Exception:
                pass

            if resultado and isinstance(resultado, dict):
                excel_path = resultado.get('excel_file')
                if excel_path and os.path.exists(excel_path):
                    excel_files.append(excel_path)

            comando = obtener_comando_por_base(base)
            if not comando:
                mensaje = f'No existe comando configurado para la base {base}'
                print(mensaje)
                if logger:
                    logger.warning(mensaje)
                continue

            Gary = f'GARY3RD{base}'

            conexion_cl = asegurar_conexion_cl()
            if conexion_cl:
                comandos = [
                    {
                        'cmd': f"SBMJOB CMD(CALL PGM(RIUNICOM63/SFL021CL) PARM('{base}')) JOB({Gary})",
                        'job': Gary
                    },
                    {
                        'cmd': f"SBMJOB CMD({comando}) JOB(RI{base}PC)",
                        'job': f'RI{base}PC'
                    },
                    {
                        'cmd': f"SBMJOB CMD(QSH CMD('/integration/uc4/component/script/itemPrice.sh {base} 000 0')) JOB(RIORPOS{base})",
                        'job': f'RIORPOS{base}'
                    },
                    {
                        'cmd': f"SBMJOB CMD(CALL PGM(RIUNICOM63/SFL022CL) PARM('{base}')) JOB(MAEPRE{base})",
                        'job': f'MAEPRE{base}'
                    }
                ]

                try:
                    for item in comandos:
                        cmd = item['cmd']
                        nombre_job = item['job']

                        conexion_cl = asegurar_conexion_cl()
                        if not conexion_cl:
                            raise Exception('No fue posible obtener instancia de EjecutorCL antes de ejecutar jobs')

                        print(cmd)
                        if logger:
                            logger.info(f'Enviando job CL: {cmd}')

                        conexion_cl.ejecutar(cmd)
                        esperar_fin_job(nombre_job, espera_segundos=30, espera_inicial=10)

                except Exception:
                    if logger:
                        logger.exception('Fallo al enviar o validar job CL')
                    raise
            else:
                mensaje = 'No fue posible inicializar EjecutorCL, se omite ejecución de jobs CL.'
                print(mensaje)
                if logger:
                    logger.warning(mensaje)

        except Exception as e:
            print('Error procesando base ' + base + ': ' + str(e))
            try:
                if 'garantias' in locals() and hasattr(garantias, 'conexion') and garantias.conexion:
                    if hasattr(garantias, 'rollback'):
                        garantias.rollback()
                    garantias.close()
            except Exception:
                pass
            continue

    try:
        if cl:
            cl.close()
    except Exception:
        pass

    if not excel_files:
        print('No se encontraron archivos Excel para adjuntar.')
        return

    destinatarios_path = os.path.join(config_dir, 'destinatarios.txt')
    if not os.path.exists(destinatarios_path):
        with open(destinatarios_path, 'w', encoding='utf-8') as f:
            f.write('eliseo_lopezp@unicomer.com,jose_preza@unicomer.com\n')

    destinatarios = []
    with open(destinatarios_path, 'r', encoding='utf-8') as f:
        for line in f:
            for mail in line.replace(';', ',').split(','):
                mail = mail.strip()
                if mail and not mail.startswith('#'):
                    destinatarios.append(mail)

    if not destinatarios:
        print('No hay destinatarios configurados.')
        return

    cuerpo_path = os.path.join(config_dir, 'cuerpo.txt')
    cuerpo_default = '<p>Estimado equipo,</p><p>Adjunto los reportes de garantías.</p>'

    if not os.path.exists(cuerpo_path):
        with open(cuerpo_path, 'w', encoding='utf-8') as f:
            f.write(cuerpo_default)
        cuerpo = cuerpo_default
    else:
        with open(cuerpo_path, 'r', encoding='utf-8') as f:
            cuerpo = f.read().strip() or cuerpo_default

    asunto = 'Reportes de Garantías - ' + fecha_hoy

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = ';'.join(destinatarios)
        mail.Subject = asunto
        mail.HTMLBody = cuerpo

        for fpath in excel_files:
            if os.path.exists(fpath):
                try:
                    mail.Attachments.Add(fpath)
                except Exception:
                    pass

        recursos_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'validacion_db', 'Recursos')
        firma_path = os.path.join(recursos_dir, 'Firma Eliseo Lopez.jpg')

        if os.path.exists(firma_path):
            try:
                mail.Attachments.Add(firma_path)
            except Exception:
                pass

        mail.Send()
        print('Correo enviado a: ' + ', '.join(destinatarios))

    except Exception as e:
        print('No se pudo enviar el correo: ' + str(e))
    finally:
        if logger:
            logger.info('flujo_completo finalizado')


def main():
    print('Running flujo_completo as script...')
    flujo_completo()
    print('Done')


if __name__ == '__main__':
    main()