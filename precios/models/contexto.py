from dataclasses import dataclass, field
from typing import Any


@dataclass
class CpDefinition:
    id: str
    nombre: str


@dataclass
class ExecutionContext:
    promo: list[dict[str, Any]] = field(default_factory=list)
    permanente: list[dict[str, Any]] = field(default_factory=list)
    recepcion: list[tuple[Any, ...]] = field(default_factory=list)
    cambios_aplicados: list[dict[str, Any]] = field(default_factory=list)
    errores: list[str] = field(default_factory=list)
    programa_sometido: bool = False
    programa_finalizado: bool = False
    actjob_resultado: dict[str, Any] | None = None
