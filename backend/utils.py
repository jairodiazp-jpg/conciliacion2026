from __future__ import annotations
from datetime import date, datetime
from typing import Any

def parse_date(value: Any) -> date | None:
    """Parsea una fecha desde varios formatos comunes."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        cleaned = value.strip()
        if cleaned:
            for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y %H:%M:%S"):
                try:
                    return datetime.strptime(cleaned, fmt).date()
                except Exception:
                    continue
    return None

def parse_amount(value: Any) -> float | None:
    """Parsea un valor numérico desde string (manejando formatos de moneda y separadores)."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str):
        return None

    cleaned = value.strip().replace("$", "").replace(" ", "")
    if not cleaned:
        return None

    # Manejo de separadores de miles y decimales
    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        # Asumimos que la coma es decimal si es el único separador
        cleaned = cleaned.replace(",", ".")
    
    try:
        return float(cleaned)
    except ValueError:
        return None