from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import date
from typing import Iterable


@dataclass(frozen=True)
class ValidacionTemporalConfig:
    tolerancia_dias: int = 0
    permitir_cruce_mes_anterior: bool = False
    permitir_cruce_ano_anterior: bool = False


@dataclass(frozen=True)
class RangoFechas:
    fecha_minima: date | None
    fecha_maxima: date | None

    def contiene(self, valor: date | None) -> bool:
        if valor is None:
            return False
        if self.fecha_minima is not None and valor < self.fecha_minima:
            return False
        if self.fecha_maxima is not None and valor > self.fecha_maxima:
            return False
        return True


@dataclass(frozen=True)
class EvaluacionTemporal:
    permitida: bool
    prioridad: int | None
    motivo: str
    diferencia_dias: int | None


def calcular_rango_fechas(fechas: Iterable[date | None]) -> RangoFechas:
    fechas_validas = [valor for valor in fechas if valor is not None]
    if not fechas_validas:
        return RangoFechas(fecha_minima=None, fecha_maxima=None)
    return RangoFechas(fecha_minima=min(fechas_validas), fecha_maxima=max(fechas_validas))


def inferir_periodo_principal(fechas: Iterable[date | None]) -> tuple[int, int] | None:
    conteo = Counter((valor.year, valor.month) for valor in fechas if valor is not None)
    if not conteo:
        return None
    periodo, _ = conteo.most_common(1)[0]
    return periodo


def _meses_diferencia(left: date, right: date) -> int:
    return abs((left.year * 12 + left.month) - (right.year * 12 + right.month))


def evaluar_temporal(
    left: date | None,
    right: date | None,
    config: ValidacionTemporalConfig | None = None,
    rango_operativo: RangoFechas | None = None,
) -> EvaluacionTemporal:
    cfg = config or ValidacionTemporalConfig()

    if left is None or right is None:
        return EvaluacionTemporal(False, None, "Fecha ausente", None)

    if rango_operativo is not None:
        if not rango_operativo.contiene(left) or not rango_operativo.contiene(right):
            return EvaluacionTemporal(
                False,
                None,
                "Fuera del rango operativo principal",
                abs((left - right).days),
            )

    diferencia_dias = abs((left - right).days)
    if left == right:
        return EvaluacionTemporal(True, 1, "Mismo dia", 0)

    mismo_ano = left.year == right.year
    mismo_mes = mismo_ano and left.month == right.month
    misma_semana = mismo_ano and left.isocalendar()[:2] == right.isocalendar()[:2]

    if mismo_mes:
        if misma_semana:
            return EvaluacionTemporal(True, 2, "Misma semana", diferencia_dias)
        return EvaluacionTemporal(True, 3, "Mismo mes", diferencia_dias)

    meses_diferencia = _meses_diferencia(left, right)
    if meses_diferencia == 1 and cfg.permitir_cruce_mes_anterior:
        if cfg.tolerancia_dias > 0 and diferencia_dias > cfg.tolerancia_dias:
            return EvaluacionTemporal(False, None, "Fuera de la tolerancia de cruce de mes", diferencia_dias)
        return EvaluacionTemporal(True, 4, "Cruce de mes habilitado", diferencia_dias)

    if abs(left.year - right.year) == 1 and cfg.permitir_cruce_ano_anterior:
        if meses_diferencia not in {1, 12}:
            return EvaluacionTemporal(False, None, "Cruce de ano no permitido por periodo", diferencia_dias)
        if cfg.tolerancia_dias > 0 and diferencia_dias > cfg.tolerancia_dias:
            return EvaluacionTemporal(False, None, "Fuera de la tolerancia de cruce de ano", diferencia_dias)
        return EvaluacionTemporal(True, 4, "Cruce de ano habilitado", diferencia_dias)

    return EvaluacionTemporal(False, None, "Año o mes incompatible", diferencia_dias)