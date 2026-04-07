from precios.services.precios_service import PreciosService


class PreciosController:
    def __init__(self, bases):
        self.service = PreciosService(bases)

    def ejecutar_base(self, base, cps=None):
        self.service.ejecutar_base(base, cps=cps)

    def guardar(self, ruta_archivo):
        self.service.guardar(ruta_archivo)
