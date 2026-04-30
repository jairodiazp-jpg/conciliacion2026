from __future__ import annotations

import base64
import unicodedata
from datetime import date, datetime
from dataclasses import dataclass
from io import BytesIO
from typing import Any

from conciliador import ConciliadorContable
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel
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

    def _normalizar_texto(self, value: str) -> str:
        normalized = unicodedata.normalize("NFKD", value)
        no_accents = "".join(char for char in normalized if not unicodedata.combining(char))
        return no_accents.strip().lower()

    def _index_annotation_columns(self, workbook) -> dict[str, tuple[int | None, int | None]]:
        indexed: dict[str, tuple[int | None, int | None]] = {}
        for sheet in workbook.worksheets:
            comment_col: int | None = None
            observation_col: int | None = None

            max_scan_row = min(sheet.max_row, 25)
            for row in sheet.iter_rows(min_row=1, max_row=max_scan_row, min_col=1, max_col=sheet.max_column):
                for cell in row:
                    if not isinstance(cell.value, str):
                        continue
                    header = self._normalizar_texto(cell.value)

                    if comment_col is None and header in {"comentario", "comentarios"}:
                        comment_col = cell.column
                    if observation_col is None and header in {"observacion", "observaciones"}:
                        observation_col = cell.column

                if comment_col is not None and observation_col is not None:
                    break

            indexed[sheet.title] = (comment_col, observation_col)

        return indexed

    def _append_text_once(self, current_value: Any, text: str) -> str:
        current = str(current_value).strip() if current_value is not None else ""
        if not current:
            return text
        if text.lower() in current.lower():
            return current
        return f"{current} | {text}"

    def _parse_row_date(self, row) -> datetime | None:
        for cell in row:
            value = cell.value
            if isinstance(value, datetime):
                return value
            if isinstance(value, date):
                return datetime.combine(value, datetime.min.time())
            if getattr(cell, "is_date", False) and isinstance(value, (int, float)):
                try:
                    parsed = from_excel(value)
                    if isinstance(parsed, datetime):
                        return parsed
                    if isinstance(parsed, date):
                        return datetime.combine(parsed, datetime.min.time())
                except Exception:
                    continue
        return None

    def _parse_row_value(self, row) -> float | None:
        for cell in row:
            value = cell.value
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                if abs(float(value)) >= 0.0001:
                    return float(value)
        return None

    def _row_contains_pse_marker(self, row) -> bool:
        texts = [self._normalizar_texto(str(cell.value)) for cell in row if isinstance(cell.value, str)]
        return any("pago virtual pse" in text for text in texts)

    def _annotate_pse_rows_in_contable(self, contable_b64: str, pse_dataset: list[dict[str, Any]]) -> str:
        workbook = load_workbook(filename=BytesIO(base64.b64decode(contable_b64)))
        annotation_columns = self._index_annotation_columns(workbook)

        for pse_row in pse_dataset:
            group_id = str(pse_row.get("id_grupo_conciliacion") or "").strip()
            if not group_id:
                continue

            try:
                pse_value = abs(float(pse_row.get("valor")))
            except (TypeError, ValueError):
                continue

            pse_date = str(pse_row.get("fecha") or "").strip()
            if not pse_date:
                continue

            pseudo_comment = str(pse_row.get("comentario_conciliacion") or "").strip()
            pseudo_values = str(pse_row.get("valores_asociados") or "").strip()
            group_text = f"[{group_id}]"

            for sheet in workbook.worksheets:
                comment_col, observation_col = annotation_columns.get(sheet.title, (None, None))
                for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
                    if not self._row_contains_pse_marker(row):
                        continue

                    row_date = self._parse_row_date(row)
                    if row_date is None:
                        continue
                    row_value = self._parse_row_value(row)
                    if row_value is None:
                        continue

                    row_date_text = row_date.date().isoformat() if isinstance(row_date, datetime) else row_date.isoformat()
                    if row_date_text != pse_date:
                        continue
                    if abs(abs(row_value) - pse_value) > self.value_tolerance:
                        continue

                    if comment_col is not None:
                        comment_cell = sheet.cell(row=row[0].row, column=comment_col)
                        comment_text = f"Cruce PSE {group_text} - {pseudo_comment if pseudo_comment else pseudo_values}"
                        comment_cell.value = self._append_text_once(comment_cell.value, comment_text)

                    if observation_col is not None:
                        observation_cell = sheet.cell(row=row[0].row, column=observation_col)
                        base_observation = pseudo_values if pseudo_values else "Reclasificacion PSE detectada"
                        observation_text = f"{group_text} {base_observation}".strip()
                        observation_cell.value = self._append_text_once(observation_cell.value, observation_text)

        output_stream = BytesIO()
        workbook.save(output_stream)
        return base64.b64encode(output_stream.getvalue()).decode("utf-8")

    def _merge_pse_comments_into_contable(self, contable_b64: str, dataset_cruces: list[dict[str, Any]]) -> str:
        workbook = load_workbook(filename=BytesIO(base64.b64decode(contable_b64)))
        annotation_columns = self._index_annotation_columns(workbook)

        for row in dataset_cruces:
            sheet_name = row.get("sheet")
            excel_row = row.get("row")
            estado = str(row.get("estado_conciliacion") or "").strip()
            comentario = str(row.get("comentario_conciliacion") or "").strip()
            valores = str(row.get("valores_asociados") or "").strip()

            if not sheet_name or sheet_name not in workbook.sheetnames:
                continue
            if estado == "Sin coincidencia":
                continue
            if not comentario and not valores:
                continue

            try:
                excel_row_int = int(excel_row)
            except (TypeError, ValueError):
                continue

            if excel_row_int < 1:
                continue

            sheet = workbook[sheet_name]
            comment_col, observation_col = annotation_columns.get(sheet_name, (None, None))
            group_id = str(row.get("id_grupo_conciliacion") or "").strip()
            group_text = f"[{group_id}]" if group_id else ""

            if comment_col is not None:
                comment_cell = sheet.cell(row=excel_row_int, column=comment_col)
                base_comment = comentario if comentario else valores
                prefix = f"Cruce PSE {group_text}".strip()
                comment_text = f"{prefix} - {base_comment}" if base_comment else prefix
                comment_cell.value = self._append_text_once(comment_cell.value, comment_text)

            if observation_col is not None:
                observation_cell = sheet.cell(row=excel_row_int, column=observation_col)
                base_observation = valores if valores else "Reclasificacion PSE detectada"
                observation_text = f"{group_text} {base_observation}".strip() if group_text else base_observation
                observation_cell.value = self._append_text_once(observation_cell.value, observation_text)

        output_stream = BytesIO()
        workbook.save(output_stream)
        return base64.b64encode(output_stream.getvalue()).decode("utf-8")

    def procesar(self) -> dict[str, Any]:
        contable_result: dict[str, Any] | None = None
        pse_result: dict[str, Any] | None = None
        logs: list[dict[str, Any]] = []
        alertas: list[str] = []

        if self.contable_bytes is not None:
            contable_result = ConciliadorContable(self.contable_bytes).procesar()
            logs.extend(contable_result.get("logs", []))
            alertas.extend(contable_result.get("alertas", []))

        if self.pse_bytes is not None and self.cruces_bytes is not None:
            pse_result = PseConciliador(
                self.pse_bytes,
                self.cruces_bytes,
                date_tolerance_days=self.date_tolerance_days,
                value_tolerance=self.value_tolerance,
            ).procesar()
            logs.extend(pse_result.get("logs", []))
            alertas.extend(pse_result.get("alertas", []))

        if contable_result is None and pse_result is None:
            raise ValueError("No se recibieron archivos para procesar")

        if contable_result is not None and pse_result is None:
            return contable_result
        if contable_result is None and pse_result is not None:
            return pse_result

        contable_result["file"] = self._merge_pse_comments_into_contable(
            contable_result["file"],
            pse_result.get("dataset_cruces", []),
        )
        contable_result["file"] = self._annotate_pse_rows_in_contable(
            contable_result["file"],
            pse_result.get("dataset", []),
        )

        files: list[dict[str, str]] = [
            {
                "name": "CONCILIACION_CONTABLE.xlsx",
                "file": contable_result["file"],
            },
            {
                "name": pse_result.get("output_name", "PSE_CONCILIADO.xlsx"),
                "file": pse_result["file"],
            },
        ]

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
