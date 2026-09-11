import os
import sys
import time
import json
from datetime import datetime

from as400_core.ejecutor_cl import EjecutorCL
from garantias.garantias import GARANTIAS

def flujo_completo(logger=None):
    config_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'archivos_config'))
    os.makedirs(config_dir, exist_ok=True)

    config_json_path = os.path.join(config_dir, 'destinatarios_flujos.json')
    bases = []
    if os.path.exists(config_json_path):
        try:
            with open(config_json_path, 'r', encoding='utf-8') as f:
                config_flujos = json.load(f)
                bases = config_flujos.get('bases_trabajo_garantias', [])
        except Exception as e:
            if logger: logger.error(f"Error leyendo {config_json_path}: {e}")
            print(f"Error leyendo {config_json_path}: {e}")

    if not bases:
        msg = f'No hay bases habilitadas en {config_json_path}'
        if logger: logger.warning(msg)
        print(msg)
        return

    def obtener_comando_por_base(base):
        if base == '11':
            return 'CALL PGM(PROGRAMAS/PCUNESCL)'
        elif base == '12':
            return 'CALL PGM(PROGRAMAS/PCUNGUCL)'
        elif base == '13':
            return 'CALL PGM(SALPGM63/PCUNHNCL)'
        elif base == '14':
            return 'CALL PGM(PROGRAMAS/PCUNNICL)'
        else:
            return None


    excel_files = []
    
    cl = None
    try:
        cl = EjecutorCL()
        if logger: logger.info('EjecutorCL inicializado correctamente')
    except Exception as e:
        if logger: logger.exception(f'No se pudo inicializar EjecutorCL: {e}')
        print(f'No se pudo inicializar EjecutorCL: {e}')
        return

    for base in bases:
        print('Procesando base: ' + base)
        garantias = None

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

            if not resultado:
                mensaje = f'El proceso de garantias fallo para la base {base}, se omite la ejecucion de jobs CL.'
                print(mensaje)
                if logger:
                    logger.error(mensaje)
                continue

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

            comandos = [
                {'cmd': f"SBMJOB CMD({comando}) JOB(RI{base}PC)", 'job': f'RI{base}PC'},
                {'cmd': f"SBMJOB CMD(CALL PGM(RIUNICOM63/SFL021CL) PARM('{base}')) JOB({Gary})", 'job': Gary},
                {'cmd': f"SBMJOB CMD({comando}) JOB(RI{base}PC)", 'job': f'RI{base}PC'},
                {'cmd': f"SBMJOB CMD(QSH CMD('/integration/uc4/component/script/itemPrice.sh {base} 000 0')) JOB(RIORPOS{base})", 'job': f'RIORPOS{base}'},
                {'cmd': f"SBMJOB CMD(CALL PGM(RIUNICOM63/SFL022CL) PARM('{base}')) JOB(MAEPRE{base})", 'job': f'MAEPRE{base}'}
            ]

            try:
                for item in comandos:
                    cmd = item['cmd']
                    nombre_job = item['job']
                    print(cmd)
                    if logger:
                        logger.info(f'Enviando job CL: {cmd}')
                    
                    cl.ejecutar(cmd)
                    cl.esperar_trabajo(nombre_job, intervalo_segundos=30)

            except Exception:
                if logger:
                    logger.exception('Fallo al enviar o validar job CL')
                raise

        except Exception as e:
            print('Error procesando base ' + base + ': ' + str(e))
            try:
                if garantias and hasattr(garantias, 'conexion') and garantias.conexion:
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
        mensaje = 'No se encontraron archivos Excel para adjuntar al correo.'
        print(mensaje)
        if logger: logger.info(mensaje)
    else:
        destinatarios = []
        habilitar_correo = "Y"
        try:
            with open(config_json_path, 'r', encoding='utf-8') as f:
                config_flujos = json.load(f)
                destinatarios = config_flujos.get('garantias', [])
                habilitar_correo = config_flujos.get('habilitar_correo', 'Y').upper()
        except Exception as e:
            if logger: logger.error(f"Error leyendo destinatarios: {e}")

        if habilitar_correo == "N":
            mensaje = "Envío de correo deshabilitado desde configuración ('habilitar_correo' = 'N'). Omitiendo envío."
            print(mensaje)
            if logger: logger.info(mensaje)
        elif destinatarios:
            fecha_hoy = datetime.today().strftime('%Y-%m-%d')
            asunto = f'Reportes de Garantías - {fecha_hoy}'
            cuerpo = '<p>Estimado equipo,</p><p>Adjunto los reportes de garantías.</p>'

            try:
                import win32com.client as win32
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = ';'.join(destinatarios)
                mail.Subject = asunto
                mail.HTMLBody = cuerpo

                for fpath in excel_files:
                    if os.path.exists(fpath):
                        mail.Attachments.Add(fpath)

                recursos_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ordenes_compra', 'Recursos')
                firma_path = os.path.join(recursos_dir, 'Firma Eliseo Lopez.jpg')
                if os.path.exists(firma_path):
                    try:
                        mail.Attachments.Add(firma_path)
                    except Exception:
                        pass

                mail.Send()
                mensaje_correo = f'Correo enviado exitosamente a: {", ".join(destinatarios)}'
                print(mensaje_correo)
                if logger: logger.info(mensaje_correo)
            except Exception as e:
                mensaje_err = f'No se pudo enviar el correo de garantías: {e}'
                print(mensaje_err)
                if logger: logger.exception(mensaje_err)
        else:
            mensaje = 'No hay destinatarios de garantías configurados en el JSON.'
            print(mensaje)
            if logger: logger.warning(mensaje)

    print('Flujo de garantías finalizado.')
    if logger:
        logger.info('flujo_completo finalizado')


def main():
    print('Running flujo_completo as script...')
    flujo_completo()
    print('Done')


if __name__ == '__main__':
    main()