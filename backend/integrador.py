from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from conciliador import ConciliadorContable
from pse_conciliador import PseConciliador


@dataclass
class ArchivoResultado:
    name: str
    content: str


class ProcesadorIntegrado:
    def __init__(
        self,
        *,
        contable_bytes: bytes | None = None,
        pse_bytes: bytes | None = None,
        cruces_bytes: bytes | None = None,
        date_tolerance_days: int = 1,
        value_tolerance: float = 0.01,
    ) -> None:
        self.contable_bytes = contable_bytes
        self.pse_bytes = pse_bytes
        self.cruces_bytes = cruces_bytes
        self.date_tolerance_days = date_tolerance_days
        self.value_tolerance = value_tolerance

    def procesar(self) -> dict[str, Any]:
        contable_result: dict[str, Any] | None = None
        pse_result: dict[str, Any] | None = None
        files: list[dict[str, str]] = []
        logs: list[dict[str, Any]] = []
        alertas: list[str] = []

        if self.contable_bytes is not None:
            contable_result = ConciliadorContable(self.contable_bytes).procesar()
            files.append(
                {
                    "name": "CONCILIACION_CONTABLE.xlsx",
                    "file": contable_result["file"],
                }
            )
            logs.extend(contable_result.get("logs", []))
            alertas.extend(contable_result.get("alertas", []))

        if self.pse_bytes is not None and self.cruces_bytes is not None:
            pse_result = PseConciliador(
                self.pse_bytes,
                self.cruces_bytes,
                date_tolerance_days=self.date_tolerance_days,
                value_tolerance=self.value_tolerance,
            ).procesar()
            files.append(
                {
                    "name": pse_result.get("output_name", "PSE_CONCILIADO.xlsx"),
                    "file": pse_result["file"],
                }
            )
            secondary_file = pse_result.get("secondary_file")
            if secondary_file:
                files.append(
                    {
                        "name": pse_result.get("secondary_output_name", "CRUCES_CONCILIADOS.xlsx"),
                        "file": secondary_file,
                    }
                )
            logs.extend(pse_result.get("logs", []))
            alertas.extend(pse_result.get("alertas", []))

        if contable_result is None and pse_result is None:
            raise ValueError("No se recibieron archivos para procesar")

        if contable_result is not None and pse_result is None:
            return contable_result
        if contable_result is None and pse_result is not None:
            return pse_result

        resumen = {
            "cruzados": int(contable_result.get("resumen", {}).get("cruzados", 0)) + int(pse_result.get("resumen", {}).get("cruzados", 0)),
            "posibles": int(contable_result.get("resumen", {}).get("posibles", 0)) + int(pse_result.get("resumen", {}).get("posibles", 0)),
            "precision_estimada": round(
                (
                    (
                        int(contable_result.get("resumen", {}).get("cruzados", 0))
                        + int(pse_result.get("resumen", {}).get("cruzados", 0))
                    )
                    /
                    max(
                        1,
                        int(contable_result.get("resumen", {}).get("total_movements", 0))
                        + int(pse_result.get("resumen", {}).get("total_movements", 0)),
                    )
                ),
                2,
            ),
        }

        return {
            "mode": "integrado",
            "contable": contable_result,
            "pse": pse_result,
            "files": files,
            "logs": logs,
            "alertas": alertas,
            "resumen": resumen,
        }
