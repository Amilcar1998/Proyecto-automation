from precios.services.precios_service import PreciosService


class EjecutadorCP:
    """Fachada de compatibilidad sobre la nueva capa de servicio."""

    def __init__(self, bases):
        self.service = PreciosService(bases)

    def ejecutar_base(self, base, cps=None):
        return self.service.ejecutar_base(base, cps=cps)

    def guardar(self, ruta_archivo):
        return self.service.guardar(ruta_archivo)
