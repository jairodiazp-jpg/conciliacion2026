from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any

UNIDAD_MINIMA = Decimal("0.01")


def parsear_decimal(value: Any) -> Decimal | None:
    """Convierte un valor a Decimal sin alterar el archivo origen."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    if not isinstance(value, str):
        return None

    cleaned = value.strip().replace("$", "").replace(" ", "").replace("\u00a0", "")
    if not cleaned:
        return None

    negativo = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        negativo = True
        cleaned = cleaned[1:-1]
    if cleaned.startswith("-"):
        negativo = True
        cleaned = cleaned[1:]

    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")

    try:
        resultado = Decimal(cleaned)
    except (InvalidOperation, ValueError):
        return None

    return -resultado if negativo else resultado


def a_unidades_minimas(value: Any) -> int | None:
    """Representa el valor en centavos exactos.

    Si hay residuo menor a un centavo, no redondea para forzar un cruce.
    """
    decimal_value = parsear_decimal(value)
    if decimal_value is None:
        return None

    scaled = decimal_value / UNIDAD_MINIMA
    integral = scaled.to_integral_value()
    if scaled != integral:
        return None
    return int(integral)


def valores_iguales(left: Any, right: Any) -> bool:
    left_units = a_unidades_minimas(left)
    right_units = a_unidades_minimas(right)
    if left_units is None or right_units is None:
        return False
    return left_units == right_units
