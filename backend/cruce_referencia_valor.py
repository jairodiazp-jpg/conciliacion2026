from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Any, Callable, Protocol


class MovimientoCruce(Protocol):
    sheet_name: str
    row: int
    value: float
    units: int
    reference_original: str
    matched: bool
    pendiente_revision: bool


@dataclass
class CruceReferenciaValor:
    origen: Any
    destino: Any
    referencia_origen: str
    referencia_destino: str
    referencia_normalizada: str
    valor_origen: float
    valor_destino: float
    unidades: int
    diferencia: int = 0


@dataclass
class PendienteReferenciaValor:
    movimientos: list[Any]
    motivo: str


def normalizar_referencia(value: Any) -> str | None:
    """Normaliza una referencia solo para comparar.

    Conserva ceros internos, guiones, barras, letras y prefijos.
    No convierte textos alfanuméricos a número.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        text = str(value)
    elif isinstance(value, float):
        if value.is_integer():
            text = str(int(value))
        else:
            return None
    else:
        text = str(value)

    normalized = text.strip()
    if not normalized:
        return None
    return normalized


def _clave(movimiento: MovimientoCruce) -> tuple[str, int] | None:
    referencia = normalizar_referencia(movimiento.reference_original)
    if referencia is None:
        return None
    return (referencia, abs(movimiento.units))


def _disponible(movimiento: MovimientoCruce) -> bool:
    return not movimiento.matched and not movimiento.pendiente_revision


def cruzar_por_referencia_y_valor(
    origenes: list[Any],
    destinos: list[Any],
    compatibles: Callable[[Any, Any], bool] | None = None,
) -> tuple[list[CruceReferenciaValor], list[PendienteReferenciaValor]]:
    """Cruza registros por referencia normalizada y valor exacto.

    Confirma únicamente coincidencias 1:1 inequívocas.
    El índice es temporal y vive solo en memoria.
    """
    indice_destino: dict[tuple[str, int], list[Any]] = defaultdict(list)
    for destino in destinos:
        if not _disponible(destino):
            continue
        clave = _clave(destino)
        if clave is None:
            continue
        indice_destino[clave].append(destino)

    confirmados: list[CruceReferenciaValor] = []
    pendientes: list[PendienteReferenciaValor] = []
    propuestas: list[tuple[Any, list[Any], str, int]] = []

    for origen in origenes:
        if not _disponible(origen):
            continue
        clave = _clave(origen)
        if clave is None:
            continue
        candidatos = [
            destino
            for destino in indice_destino.get(clave, [])
            if _disponible(destino) and (compatibles is None or compatibles(origen, destino))
        ]
        if not candidatos:
            continue
        propuestas.append((origen, candidatos, clave[0], clave[1]))

    destino_a_origenes: dict[tuple[str, int], list[Any]] = defaultdict(list)
    for origen, candidatos, _, _ in propuestas:
        if len(candidatos) == 1:
            destino = candidatos[0]
            destino_a_origenes[(destino.sheet_name, destino.row)].append(origen)

    for origen, candidatos, referencia_norm, unidades in propuestas:
        if not _disponible(origen):
            continue
        if len(candidatos) != 1:
            involucrados = [origen, *candidatos]
            for movimiento in involucrados:
                movimiento.pendiente_revision = True
            pendientes.append(
                PendienteReferenciaValor(
                    movimientos=involucrados,
                    motivo=(
                        "Pendiente de revisión: hay más de un candidato con la misma "
                        "referencia y el mismo valor."
                    ),
                )
            )
            continue

        destino = candidatos[0]
        if not _disponible(destino):
            continue
        reclamantes = destino_a_origenes[(destino.sheet_name, destino.row)]
        if len(reclamantes) != 1:
            involucrados = [*reclamantes, destino]
            for movimiento in involucrados:
                movimiento.pendiente_revision = True
            pendientes.append(
                PendienteReferenciaValor(
                    movimientos=involucrados,
                    motivo=(
                        "Pendiente de revisión: hay más de un candidato con la misma "
                        "referencia y el mismo valor."
                    ),
                )
            )
            continue

        origen.matched = True
        destino.matched = True
        confirmados.append(
            CruceReferenciaValor(
                origen=origen,
                destino=destino,
                referencia_origen=origen.reference_original,
                referencia_destino=destino.reference_original,
                referencia_normalizada=referencia_norm,
                valor_origen=origen.value,
                valor_destino=destino.value,
                unidades=unidades,
                diferencia=0,
            )
        )

    return confirmados, pendientes
