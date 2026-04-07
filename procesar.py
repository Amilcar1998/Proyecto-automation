import logging
import sys
from datetime import datetime

from precios.controllers.precios_controller import PreciosController


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(name)s %(levelname)s %(message)s"
)


def _parsear_argumentos():
    base = "RI11DB"
    cps = None

    argumentos = sys.argv[1:]
    i = 0
    while i < len(argumentos):
        argumento = str(argumentos[i]).strip()
        argumento_upper = argumento.upper()

        if argumento_upper in ("--BASE", "-B") and (i + 1) < len(argumentos):
            base = str(argumentos[i + 1]).strip().upper()
            i += 2
            continue

        if argumento_upper in ("--CP", "--CPS", "-C") and (i + 1) < len(argumentos):
            cps = [
                cp.strip().upper()
                for cp in str(argumentos[i + 1]).split(",")
                if cp.strip()
            ]
            i += 2
            continue

        if argumento_upper.startswith("RI") and argumento_upper.endswith("DB"):
            base = argumento_upper
            i += 1
            continue

        if argumento_upper.startswith("SV"):
            cps = cps or []
            cps.append(argumento_upper)
            i += 1
            continue

        i += 1

    return base, cps


def main_precios():
    base, cps = _parsear_argumentos()
    bases = [base]

    for base in bases:
        controller = PreciosController([base])
        controller.ejecutar_base(base, cps=cps)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        controller.guardar(f"evidencia_pruebas_precios_{base}_{timestamp}.docx")


if __name__ == "__main__":
    main_precios()
