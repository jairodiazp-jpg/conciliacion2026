from __future__ import annotations

import base64
import re
import unicodedata
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet.worksheet import Worksheet

from validacion_temporal import (
    RangoFechas,
    ValidacionTemporalConfig,
    calcular_rango_fechas,
    evaluar_temporal,
    inferir_periodo_principal,
)
from utils import parse_date, parse_amount
from dinero import a_unidades_minimas
from cruce_referencia_valor import cruzar_por_referencia_y_valor


DATE_ALIASES = {
    "fecha",
    "date",
    "transaction date",
    "fecha movimiento",
    "fecha transaccion",
    "fecha transacción",
}
VALUE_ALIASES = {
    "valor",
    "monto",
    "importe",
    "amount",
    "valor transaccion",
    "valor transacción",
    "valor movimiento",
}
ACCOUNT_ALIASES = {
    "cuenta",
    "cuenta contable",
    "account",
    "codigo cuenta",
    "código cuenta",
    "numero cuenta",
    "número cuenta",
}
DESCRIPTION_ALIASES = {
    "descripcion",
    "descripción",
    "detalle",
    "concepto",
    "glosa",
}
PSE_DESCRIPTION_MARKER = "pago virtual pse"
DATE_ALIASES_PRIORITY = (
    "fecha documento",
    "fecha de transaccion",
    "fecha transaccion",
    "fecha movimiento",
    "transaction date",
    "fecha contabilizacion",
    "fecha",
    "date",
)
VALUE_ALIASES_PRIORITY = (
    "valor transaccion",
    "valor movimiento",
    "valor",
    "monto",
    "importe",
    "amount",
)
REFERENCE_ALIASES_PRIORITY = (
    "referencia",
    "reference",
    "ref",
    "id transaccion",
    "id transacción",
    "numero referencia",
    "número referencia",
    "nro referencia",
    "comprobante",
    "ticket",
)

VALUE_TOLERANCE_DEFAULT = 0.01
DATE_TOLERANCE_DEFAULT = 1
MAX_SUBSET_CANDIDATES = 16
MAX_SUBSET_SIZE = 6

HEADER_FILL = PatternFill(start_color="FFEAF2FF", end_color="FFEAF2FF", fill_type="solid")
PINK_FILL = PatternFill(start_color="FFFFC0CB", end_color="FFFFC0CB", fill_type="solid")


@dataclass
class SheetSchema:
    header_row: int
    date_col: int
    value_col: int
    account_col: int | None = None
    description_col: int | None = None
    reference_col: int | None = None


@dataclass
class Movimiento:
    sheet_name: str
    row: int
    raw_date: date | None
    value: float
    account_label: str
    units: int = 0
    reference_original: str = ""
    matched: bool = False
    pendiente_revision: bool = False


@dataclass
class MatchResult:
    pse_entry: Movimiento
    matched_entries: list[Movimiento]
    state: str
    group_id: str | None
    comment: str
    account_label: str
    values_associated: str
    difference: float | None


class PseConciliador:
    def __init__(
        self,
        pse_bytes: bytes,
        cruces_bytes: bytes,
        *,
        date_tolerance_days: int = DATE_TOLERANCE_DEFAULT,
        value_tolerance: float = VALUE_TOLERANCE_DEFAULT,
        temporal_config: ValidacionTemporalConfig | None = None,
    ) -> None:
        self.pse_workbook = load_workbook(filename=BytesIO(pse_bytes))
        self.cruces_workbook = load_workbook(filename=BytesIO(cruces_bytes))
        self.date_tolerance_days = max(0, int(date_tolerance_days))
        self.value_tolerance = max(0.0, float(value_tolerance))
        self.temporal_config = temporal_config or ValidacionTemporalConfig()
        self.logs: list[dict[str, Any]] = []
        self.alertas: list[str] = []
        self.dataset: list[dict[str, Any]] = []
        self.dataset_cruces: list[dict[str, Any]] = []
        self.group_counter = 1
        self.pse_schemas: dict[str, SheetSchema] = {}
        self.cruces_schemas: dict[str, SheetSchema] = {}
        self.pse_entries: list[Movimiento] = []
        self.cruces_entries: list[Movimiento] = []
        self.match_results: list[MatchResult] = []
        self.rango_operativo: RangoFechas | None = None

    def procesar(self) -> dict[str, Any]:
        self.pse_entries = self._extract_movements(self.pse_workbook, self.pse_schemas)
        self.cruces_entries = self._extract_movements(
            self.cruces_workbook,
            self.cruces_schemas,
            only_virtual_pse=True,
        )
        self._configurar_rango_temporal(self.pse_entries + self.cruces_entries)
        if self.value_tolerance > 0:
            self.logs.append(
                {
                    "tipo": "info",
                    "valor": self.value_tolerance,
                    "fecha": None,
                    "confianza": 1.0,
                    "detalle": (
                        "La tolerancia de valor se ignora en cruces PSE: "
                        "solo se confirman igualdades monetarias exactas."
                    ),
                }
            )
        self.match_results = self._conciliar(self.pse_entries, self.cruces_entries)

        self._escribir_enriquecimiento_pse(self.match_results)
        self._escribir_enriquecimiento_cruces(self.match_results)
        resumen = self._construir_resumen(self.match_results, self.pse_entries)
        self.alertas = self._construir_alertas(self.match_results, resumen)

        pse_output = BytesIO()
        self.pse_workbook.save(pse_output)
        pse_b64 = base64.b64encode(pse_output.getvalue()).decode("utf-8")

        cruces_output = BytesIO()
        self.cruces_workbook.save(cruces_output)
        cruces_b64 = base64.b64encode(cruces_output.getvalue()).decode("utf-8")

        return {
            "file": pse_b64,
            "output_name": "PSE_CONCILIADO.xlsx",
            "secondary_file": cruces_b64,
            "secondary_output_name": "CRUCES_CONCILIADOS.xlsx",
            "dataset": self.dataset,
            "dataset_cruces": self.dataset_cruces,
            "logs": self.logs,
            "resumen": resumen,
            "alertas": self.alertas,
        }

    def _normalizar_texto(self, value: str) -> str:
        normalized = unicodedata.normalize("NFKD", value)
        no_accents = "".join(char for char in normalized if not unicodedata.combining(char))
        return no_accents.strip().lower()

    def _configurar_rango_temporal(self, movements: list[Movimiento]) -> None:
        fechas = [movement.raw_date for movement in movements if movement.raw_date is not None]
        rango_general = calcular_rango_fechas(fechas)
        periodo_principal = inferir_periodo_principal(fechas)

        if periodo_principal is None:
            self.rango_operativo = rango_general
            return

        fechas_principales = [
            value
            for value in fechas
            if value is not None and (value.year, value.month) == periodo_principal
        ]
        if fechas_principales and len(fechas_principales) / max(1, len(fechas)) >= 0.6:
            self.rango_operativo = calcular_rango_fechas(fechas_principales)
            self.logs.append(
                {
                    "tipo": "info_temporal",
                    "valor": 0.0,
                    "fecha": self.rango_operativo.fecha_minima.isoformat() if self.rango_operativo.fecha_minima else None,
                    "confianza": 1.0,
                    "detalle": (
                        f"Periodo principal detectado: {periodo_principal[0]}-{periodo_principal[1]:02d}. "
                        f"Rango operativo {self.rango_operativo.fecha_minima} a {self.rango_operativo.fecha_maxima}"
                    ),
                }
            )
            return

        self.rango_operativo = rango_general

    def _parse_date(self, cell) -> date | None:
        parsed = parse_date(cell.value)
        if parsed:
            return parsed
        if getattr(cell, "is_date", False):
            if isinstance(cell.value, (int, float)):
                try:
                    parsed_excel = from_excel(cell.value)
                    return parsed_excel.date() if isinstance(parsed_excel, datetime) else parsed_excel
                except Exception:
                    return None
        return None

    def _parse_float(self, value: Any) -> float | None:
        return parse_amount(value)

    def _stringify(self, value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def _stringify_reference(self, value: Any) -> str:
        if value is None or isinstance(value, bool):
            return ""
        if isinstance(value, int):
            return str(value)
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        if isinstance(value, str):
            return value
        return str(value)

    def _sheet_account_label(self, sheet_name: str) -> str:
        match = re.search(r"\d{4,}", sheet_name)
        if match:
            return match.group(0)
        return sheet_name.strip()

    def _alias_priority(self, header: str, aliases: tuple[str, ...]) -> int | None:
        if header in aliases:
            return aliases.index(header)
        for index, alias in enumerate(aliases):
            if alias in header:
                return 100 + index
        return None

    def _detect_schema(self, sheet: Worksheet) -> SheetSchema | None:
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 30), min_col=1, max_col=sheet.max_column):
            date_candidates: list[tuple[int, int]] = []
            value_candidates: list[tuple[int, int]] = []
            reference_candidates: list[tuple[int, int]] = []
            account_col: int | None = None
            description_col: int | None = None

            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                header = self._normalizar_texto(cell.value)
                if not header:
                    continue

                date_priority = self._alias_priority(header, DATE_ALIASES_PRIORITY)
                if date_priority is not None:
                    date_candidates.append((date_priority, cell.column))

                value_priority = self._alias_priority(header, VALUE_ALIASES_PRIORITY)
                if value_priority is not None:
                    value_candidates.append((value_priority, cell.column))

                reference_priority = self._alias_priority(header, REFERENCE_ALIASES_PRIORITY)
                if reference_priority is not None:
                    reference_candidates.append((reference_priority, cell.column))

                if account_col is None and (header in ACCOUNT_ALIASES or any(alias in header for alias in ACCOUNT_ALIASES)):
                    account_col = cell.column
                if description_col is None and (
                    header in DESCRIPTION_ALIASES or any(alias in header for alias in DESCRIPTION_ALIASES)
                ):
                    description_col = cell.column

            if date_candidates and value_candidates:
                return SheetSchema(
                    header_row=row[0].row,
                    date_col=min(date_candidates)[1],
                    value_col=min(value_candidates)[1],
                    account_col=account_col,
                    description_col=description_col,
                    reference_col=min(reference_candidates)[1] if reference_candidates else None,
                )
        return None

    def _is_virtual_pse_row(self, row, schema: SheetSchema) -> bool:
        # Primero intenta buscar en la columna de descripción si existe
        if schema.description_col is not None and schema.description_col - 1 < len(row):
            description_text = self._normalizar_texto(self._stringify(row[schema.description_col - 1].value))
            if PSE_DESCRIPTION_MARKER in description_text:
                return True
        
        # Si no hay descripción o no se encontró el marcador, busca en todas las celdas
        texts = [self._normalizar_texto(self._stringify(cell.value)) for cell in row if isinstance(cell.value, str)]
        return any(PSE_DESCRIPTION_MARKER in text for text in texts)

    def _resolve_account_label(self, sheet: Worksheet, row, schema: SheetSchema) -> str:
        if schema.account_col is not None:
            cell_value = row[schema.account_col - 1].value if schema.account_col - 1 < len(row) else None
            cell_text = self._stringify(cell_value)
            if cell_text:
                return cell_text

        for cell in row:
            text = self._stringify(cell.value)
            if text and re.search(r"\d{4,}", text):
                return text

        return self._sheet_account_label(sheet.title)

    def _extract_movements(
        self,
        workbook,
        schema_registry: dict[str, SheetSchema],
        *,
        only_virtual_pse: bool = False,
    ) -> list[Movimiento]:
        movements: list[Movimiento] = []

        for sheet in workbook.worksheets:
            schema = self._detect_schema(sheet)
            if schema is None:
                continue

            schema_registry[sheet.title] = schema

            for row in sheet.iter_rows(min_row=schema.header_row + 1, max_row=sheet.max_row):
                if only_virtual_pse and not self._is_virtual_pse_row(row, schema):
                    continue

                date_cell = row[schema.date_col - 1] if schema.date_col - 1 < len(row) else None
                value_cell = row[schema.value_col - 1] if schema.value_col - 1 < len(row) else None
                if date_cell is None or value_cell is None:
                    continue

                raw_date = self._parse_date(date_cell)
                value = self._parse_float(value_cell.value)
                units = a_unidades_minimas(value_cell.value)
                if raw_date is None or value is None or units is None or units == 0:
                    continue

                reference_original = ""
                if schema.reference_col is not None and schema.reference_col - 1 < len(row):
                    reference_original = self._stringify_reference(row[schema.reference_col - 1].value)

                movements.append(
                    Movimiento(
                        sheet_name=sheet.title,
                        row=row[0].row,
                        raw_date=raw_date,
                        value=value,
                        account_label=self._resolve_account_label(sheet, row, schema),
                        units=units,
                        reference_original=reference_original,
                    )
                )

        return movements

    def _date_gap_days(self, left: date | None, right: date | None) -> int | None:
        if left is None or right is None:
            return None
        return abs((left - right).days)

    def _temporal_evaluation(self, left: date | None, right: date | None):
        return evaluar_temporal(left, right, self.temporal_config, self.rango_operativo)

    def _log_temporal_rejection(self, left: Movimiento, right: Movimiento, motivo: str) -> None:
        self.logs.append(
            {
                "tipo": "rechazado_temporal",
                "valor": round(abs(left.value), 2),
                "fecha": left.raw_date.isoformat() if left.raw_date else (right.raw_date.isoformat() if right.raw_date else None),
                "confianza": 0.0,
                "detalle": (
                    f"Movimiento rechazado: Fecha PSE: {left.raw_date.isoformat() if left.raw_date else None}. "
                    f"Fecha cruce: {right.raw_date.isoformat() if right.raw_date else None}. Motivo: {motivo}"
                ),
            }
        )

    def _same_absolute_units(self, left: Movimiento, right: Movimiento) -> bool:
        return abs(left.units) == abs(right.units)

    def _currency_string(self, value: float) -> str:
        return f"{abs(value):,.2f}"

    def _candidate_pool(self, entry: Movimiento, pool: list[Movimiento]) -> list[Movimiento]:
        candidates: list[Movimiento] = []
        for candidate in pool:
            if candidate.matched or candidate.pendiente_revision:
                continue
            if candidate.raw_date is None:
                continue
            decision = self._temporal_evaluation(entry.raw_date, candidate.raw_date)
            if not decision.permitida:
                if self._same_absolute_units(entry, candidate):
                    self._log_temporal_rejection(entry, candidate, decision.motivo)
                continue
            if abs(candidate.units) > abs(entry.units):
                continue
            candidates.append(candidate)

        candidates.sort(
            key=lambda item: (
                self._temporal_evaluation(entry.raw_date, item.raw_date).prioridad or 99,
                self._date_gap_days(entry.raw_date, item.raw_date) or 999,
                abs(abs(entry.units) - abs(item.units)),
                -abs(item.units),
            )
        )
        return candidates[:MAX_SUBSET_CANDIDATES]

    def _disponible(self, movement: Movimiento) -> bool:
        return not movement.matched and not movement.pendiente_revision

    def _temporalmente_compatible(self, left: Movimiento, right: Movimiento) -> bool:
        if left.raw_date is None or right.raw_date is None:
            return False
        decision = self._temporal_evaluation(left.raw_date, right.raw_date)
        if decision.permitida:
            return True
        if self._same_absolute_units(left, right):
            self._log_temporal_rejection(left, right, decision.motivo)
        return False

    def _audit_comment(
        self,
        *,
        titulo: str,
        left: Movimiento,
        right_entries: list[Movimiento],
        suma_unidades: int,
        comparado: Movimiento,
    ) -> str:
        refs_left = left.reference_original or "(sin referencia)"
        refs_right = ", ".join(entry.reference_original or "(sin referencia)" for entry in right_entries)
        valores_left = self._currency_string(left.value)
        valores_right = " + ".join(self._currency_string(entry.value) for entry in right_entries)
        suma = self._currency_string(suma_unidades / 100)
        comparado_txt = self._currency_string(comparado.value)
        return (
            f"{titulo}. Referencia origen: {refs_left}. Referencias destino: {refs_right}. "
            f"Valores originales: {valores_left} vs {valores_right}. "
            f"Suma calculada: {suma}. Valor comparado: {comparado_txt}. Diferencia: 0"
        )

    def _unique_exact_matches(self, entry: Movimiento, pool: list[Movimiento]) -> list[Movimiento]:
        return [
            candidate
            for candidate in pool
            if self._disponible(candidate)
            and candidate.raw_date is not None
            and self._temporal_evaluation(entry.raw_date, candidate.raw_date).permitida
            and self._same_absolute_units(entry, candidate)
        ]

    def _find_exact_subsets(self, entry: Movimiento, candidates: list[Movimiento]) -> list[list[Movimiento]]:
        if not candidates:
            return []

        target_units = abs(entry.units)
        exact_subsets: list[list[Movimiento]] = []

        def backtrack(start: int, chosen: list[Movimiento], current_sum: int) -> None:
            if chosen and current_sum == target_units:
                exact_subsets.append(chosen.copy())
            if len(chosen) >= MAX_SUBSET_SIZE or start >= len(candidates):
                return
            for index in range(start, len(candidates)):
                candidate = candidates[index]
                next_sum = current_sum + abs(candidate.units)
                if next_sum > target_units:
                    continue
                chosen.append(candidate)
                backtrack(index + 1, chosen, next_sum)
                chosen.pop()

        backtrack(0, [], 0)
        return exact_subsets

    def _build_account_string(self, entries: list[Movimiento]) -> str:
        labels = []
        for entry in entries:
            if entry.account_label and entry.account_label not in labels:
                labels.append(entry.account_label)
        return ", ".join(labels)

    def _build_associated_values(self, entries: list[Movimiento]) -> str:
        parts = []
        for entry in entries:
            parts.append(f"{entry.sheet_name}:fila {entry.row}={self._currency_string(entry.value)}")
        return " | ".join(parts)

    def _build_result_map(self, match_results: list[MatchResult]) -> dict[tuple[str, int], MatchResult]:
        result_map: dict[tuple[str, int], MatchResult] = {}
        for result in match_results:
            result_map[(result.pse_entry.sheet_name, result.pse_entry.row)] = result
            for entry in result.matched_entries:
                result_map[(entry.sheet_name, entry.row)] = result
        return result_map

    def _prepare_output_sheet(self, sheet: Worksheet, header_row: int, headers: list[str]) -> int:
        start_column = sheet.max_column + 1
        for offset, header in enumerate(headers):
            target_column = start_column + offset
            header_cell = sheet.cell(row=header_row, column=target_column)
            header_cell.value = header
            self._style_new_header(sheet, header_row, target_column)
        return start_column

    def _write_enrichment_rows(
        self,
        sheet: Worksheet,
        schema: SheetSchema,
        movements: list[Movimiento],
        result_map: dict[tuple[str, int], MatchResult],
        *,
        dataset_target: list[dict[str, Any]],
    ) -> None:
        start_column = self._prepare_output_sheet(
            sheet,
            schema.header_row,
            [
                "Estado_Conciliacion",
                "Cuenta_Contable",
                "Valores_Asociados",
                "Comentario_Conciliacion",
                "ID_Grupo_Conciliacion",
            ],
        )

        for movement in movements:
            if movement.sheet_name != sheet.title:
                continue

            result = result_map.get((movement.sheet_name, movement.row))
            if result is None:
                state = "Sin coincidencia"
                account_label = ""
                values_associated = ""
                comment = "Sin coincidencia en cruces contables"
                group_id = ""
                difference = None
            else:
                state = result.state
                account_label = result.account_label
                values_associated = result.values_associated
                comment = result.comment
                group_id = result.group_id or ""
                difference = result.difference

            sheet.cell(row=movement.row, column=start_column).value = state
            sheet.cell(row=movement.row, column=start_column + 1).value = account_label
            sheet.cell(row=movement.row, column=start_column + 2).value = values_associated
            sheet.cell(row=movement.row, column=start_column + 3).value = comment
            sheet.cell(row=movement.row, column=start_column + 4).value = group_id

            dataset_target.append(
                {
                    "sheet": movement.sheet_name,
                    "row": movement.row,
                    "fecha": movement.raw_date.isoformat() if movement.raw_date else None,
                    "valor": round(movement.value, 2),
                    "estado_conciliacion": state,
                    "cuenta_contable": account_label,
                    "valores_asociados": values_associated,
                    "comentario_conciliacion": comment,
                    "id_grupo_conciliacion": group_id,
                    "diferencia": round(difference, 2) if difference is not None else None,
                }
            )

    def _escribir_enriquecimiento_pse(self, match_results: list[MatchResult]) -> None:
        result_map = self._build_result_map(match_results)
        for sheet in self.pse_workbook.worksheets:
            schema = self.pse_schemas.get(sheet.title)
            if schema is None:
                continue
            self._write_enrichment_rows(
                sheet,
                schema,
                self.pse_entries,
                result_map,
                dataset_target=self.dataset,
            )

    def _escribir_enriquecimiento_cruces(self, match_results: list[MatchResult]) -> None:
        result_map = self._build_result_map(match_results)
        for sheet in self.cruces_workbook.worksheets:
            schema = self.cruces_schemas.get(sheet.title)
            if schema is None:
                continue
            self._write_enrichment_rows(
                sheet,
                schema,
                self.cruces_entries,
                result_map,
                dataset_target=self.dataset_cruces,
            )

    def _register_match(
        self,
        pse_entry: Movimiento,
        matched_entries: list[Movimiento],
        state: str,
        difference: float | None,
        comment: str | None = None,
    ) -> MatchResult:
        group_id = f"GRP-{self.group_counter:04d}"
        self.group_counter += 1

        for entry in matched_entries:
            entry.matched = True
            self._color_movement(entry, PINK_FILL)
        pse_entry.matched = True
        self._color_movement(pse_entry, PINK_FILL)

        account_label = self._build_account_string(matched_entries)
        values_associated = self._build_associated_values(matched_entries)
        total_match_value = sum(abs(entry.value) for entry in matched_entries)

        if comment is None:
            if state == "Conciliado 1:1":
                comment = (
                    f"Conciliado con cuenta {account_label} por valor total "
                    f"{self._currency_string(total_match_value)} (1 transacción contable). Diferencia: 0"
                )
                log_type = "match_1_1"
            elif state == "Conciliado 1:N":
                comment = (
                    f"Conciliado con cuenta {account_label} por valor total "
                    f"{self._currency_string(total_match_value)} ({len(matched_entries)} transacciones contables). "
                    f"Diferencia: 0"
                )
                log_type = "match_1_n"
            elif state == "Conciliado por referencia y valor":
                comment = self._audit_comment(
                    titulo="Conciliado por referencia y valor",
                    left=pse_entry,
                    right_entries=matched_entries,
                    suma_unidades=abs(pse_entry.units),
                    comparado=matched_entries[0] if matched_entries else pse_entry,
                )
                log_type = "match_referencia_valor"
            else:
                comment = f"{state}. Diferencia: 0"
                log_type = "match"
        else:
            log_type = "match_referencia_valor" if "referencia" in state.lower() else "match"

        self.logs.append(
            {
                "tipo": log_type,
                "valor": round(abs(pse_entry.value), 2),
                "fecha": pse_entry.raw_date.isoformat() if pse_entry.raw_date else None,
                "confianza": 0.99 if state.startswith("Conciliado") else 0.7,
                "detalle": f"{state} entre {pse_entry.sheet_name}:fila {pse_entry.row} y {values_associated}",
            }
        )

        return MatchResult(
            pse_entry=pse_entry,
            matched_entries=matched_entries,
            state=state,
            group_id=group_id,
            comment=comment,
            account_label=account_label,
            values_associated=values_associated,
            difference=difference,
        )

    def _register_match_n_to_1(
        self,
        cruce_entry: Movimiento,
        pse_entries: list[Movimiento],
        state: str,
        difference: float | None,
    ) -> list[MatchResult]:
        group_id = f"GRP-{self.group_counter:04d}"
        self.group_counter += 1

        cruce_entry.matched = True
        self._color_movement(cruce_entry, PINK_FILL)
        for entry in pse_entries:
            entry.matched = True
            self._color_movement(entry, PINK_FILL)

        total_pse = sum(abs(entry.value) for entry in pse_entries)
        pse_values = self._build_associated_values(pse_entries)
        cruce_reference = f"{cruce_entry.sheet_name}:fila {cruce_entry.row}={self._currency_string(cruce_entry.value)}"
        refs = ", ".join(entry.reference_original or "(sin referencia)" for entry in pse_entries)
        comment = (
            f"Conciliado N:1 contra {cruce_reference}. "
            f"Referencias PSE: {refs}. "
            f"Pagos PSE asociados ({len(pse_entries)}): {pse_values}. "
            f"Suma calculada: {self._currency_string(total_pse)}. "
            f"Valor contable comparado: {self._currency_string(cruce_entry.value)}. "
            f"Diferencia: 0"
        )

        self.logs.append(
            {
                "tipo": "match_n_1",
                "valor": round(abs(cruce_entry.value), 2),
                "fecha": cruce_entry.raw_date.isoformat() if cruce_entry.raw_date else None,
                "confianza": 0.98 if difference in (None, 0.0) else 0.7,
                "detalle": f"{state} para {cruce_reference} usando pagos PSE: {pse_values}",
            }
        )

        results: list[MatchResult] = []
        for pse_entry in pse_entries:
            results.append(
                MatchResult(
                    pse_entry=pse_entry,
                    matched_entries=[cruce_entry],
                    state=state,
                    group_id=group_id,
                    comment=comment,
                    account_label=cruce_entry.account_label,
                    values_associated=cruce_reference,
                    difference=difference,
                )
            )
        return results

    def _register_pending(
        self,
        pse_entries: list[Movimiento],
        related: list[Movimiento],
        comment: str,
    ) -> list[MatchResult]:
        results: list[MatchResult] = []
        seen: set[int] = set()
        for movement in [*pse_entries, *related]:
            if id(movement) in seen:
                continue
            seen.add(id(movement))
            movement.pendiente_revision = True

        for pse_entry in pse_entries:
            related_entries = [item for item in related if item is not pse_entry]
            results.append(
                MatchResult(
                    pse_entry=pse_entry,
                    matched_entries=related_entries,
                    state="Pendiente de revisión",
                    group_id=None,
                    comment=comment,
                    account_label=self._build_account_string(related_entries),
                    values_associated=self._build_associated_values(related_entries),
                    difference=None,
                )
            )
            self.logs.append(
                {
                    "tipo": "pendiente",
                    "valor": round(abs(pse_entry.value), 2),
                    "fecha": pse_entry.raw_date.isoformat() if pse_entry.raw_date else None,
                    "confianza": 0.0,
                    "detalle": f"Pendiente de revisión para {pse_entry.sheet_name}:fila {pse_entry.row}. {comment}",
                }
            )
        return results

    def _apply_unique_subsets(
        self,
        targets: list[Movimiento],
        pool: list[Movimiento],
        *,
        as_n_to_1: bool,
        results: list[MatchResult],
    ) -> None:
        for target in targets:
            if not self._disponible(target):
                continue
            available = [item for item in pool if self._disponible(item)]
            if not available:
                continue
            candidates = self._candidate_pool(target, available)
            subsets = self._find_exact_subsets(target, candidates)
            if not subsets:
                continue
            if len(subsets) > 1:
                involucrados = []
                for subset in subsets:
                    involucrados.extend(subset)
                if as_n_to_1:
                    results.extend(
                        self._register_pending(
                            involucrados,
                            [target],
                            "Pendiente de revisión: existen varias combinaciones PSE con la misma suma exacta.",
                        )
                    )
                else:
                    results.extend(
                        self._register_pending(
                            [target],
                            involucrados,
                            "Pendiente de revisión: existen varias combinaciones contables con la misma suma exacta.",
                        )
                    )
                continue

            subset = subsets[0]
            if as_n_to_1:
                state = "Conciliado N:1" if len(subset) > 1 else "Conciliado 1:1"
                results.extend(self._register_match_n_to_1(target, subset, state, 0.0))
            else:
                state = "Conciliado 1:N" if len(subset) > 1 else "Conciliado 1:1"
                comment = self._audit_comment(
                    titulo=state,
                    left=target,
                    right_entries=subset,
                    suma_unidades=sum(abs(item.units) for item in subset),
                    comparado=target,
                )
                results.append(self._register_match(target, subset, state, 0.0, comment=comment))

    def _conciliar(self, pse_entries: list[Movimiento], cruce_entries: list[Movimiento]) -> list[MatchResult]:
        results: list[MatchResult] = []
        remaining_pse = sorted(pse_entries, key=lambda item: abs(item.units), reverse=True)
        remaining_cruces = sorted(cruce_entries, key=lambda item: abs(item.units), reverse=True)
        pse_ids = {(item.sheet_name, item.row) for item in pse_entries}

        confirmados_ref, pendientes_ref = cruzar_por_referencia_y_valor(
            remaining_pse,
            remaining_cruces,
            compatibles=self._temporalmente_compatible,
        )
        for cruce in confirmados_ref:
            comment = self._audit_comment(
                titulo="Conciliado por referencia y valor",
                left=cruce.origen,
                right_entries=[cruce.destino],
                suma_unidades=abs(cruce.unidades),
                comparado=cruce.destino,
            )
            results.append(
                self._register_match(
                    cruce.origen,
                    [cruce.destino],
                    "Conciliado por referencia y valor",
                    0.0,
                    comment=comment,
                )
            )
        for pendiente in pendientes_ref:
            pse_pendientes = [
                item for item in pendiente.movimientos if (item.sheet_name, item.row) in pse_ids
            ]
            related = [item for item in pendiente.movimientos if (item.sheet_name, item.row) not in pse_ids]
            if pse_pendientes:
                results.extend(self._register_pending(pse_pendientes, related, pendiente.motivo))

        self._apply_unique_subsets(remaining_cruces, remaining_pse, as_n_to_1=True, results=results)
        self._apply_unique_subsets(remaining_pse, remaining_cruces, as_n_to_1=False, results=results)

        seen_pse = {(result.pse_entry.sheet_name, result.pse_entry.row) for result in results}
        for entry in remaining_pse:
            if (entry.sheet_name, entry.row) in seen_pse:
                continue
            state = "Pendiente de revisión" if entry.pendiente_revision else "Sin coincidencia"
            comment = (
                "Pendiente de revisión: no hay un cruce único comprobable."
                if entry.pendiente_revision
                else "Sin coincidencia en cruces contables"
            )
            results.append(
                MatchResult(
                    pse_entry=entry,
                    matched_entries=[],
                    state=state,
                    group_id=None,
                    comment=comment,
                    account_label="",
                    values_associated="",
                    difference=None,
                )
            )
            self.logs.append(
                {
                    "tipo": "pendiente" if entry.pendiente_revision else "unmatched",
                    "valor": round(abs(entry.value), 2),
                    "fecha": entry.raw_date.isoformat() if entry.raw_date else None,
                    "confianza": 0.0,
                    "detalle": f"{comment} para {entry.sheet_name}:fila {entry.row}",
                }
            )

        return results

    def _style_new_header(self, sheet: Worksheet, header_row: int, column_index: int) -> None:
        source_column = max(1, column_index - 1)
        source_cell = sheet.cell(row=header_row, column=source_column)
        target_cell = sheet.cell(row=header_row, column=column_index)
        target_cell._style = copy(source_cell._style)
        target_cell.font = copy(source_cell.font)
        target_cell.fill = HEADER_FILL
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)

    def _color_movement(self, movement: Movimiento, fill: PatternFill) -> None:
        wb = self.pse_workbook if movement.sheet_name in self.pse_schemas else self.cruces_workbook
        sheet = wb[movement.sheet_name]
        schema = self.pse_schemas.get(movement.sheet_name) or self.cruces_schemas.get(movement.sheet_name)
        if schema:
            cell = sheet.cell(row=movement.row, column=schema.value_col)
            cell.fill = fill

    def _construir_resumen(self, match_results: list[MatchResult], pse_entries: list[Movimiento]) -> dict[str, Any]:
        conciliados = len([result for result in match_results if result.state.startswith("Conciliado")])
        parciales = len([result for result in match_results if result.state == "Pendiente de revisión"])
        sin_coincidencia = len([result for result in match_results if result.state == "Sin coincidencia"])
        total = len(pse_entries)
        porcentaje_conciliados = round((conciliados / total * 100) if total else 0.0, 2)
        porcentaje_parciales = round((parciales / total * 100) if total else 0.0, 2)
        precision = round((conciliados / total) if total else 0.0, 2)

        return {
            "cruzados": conciliados,
            "posibles": parciales,
            "precision_estimada": precision,
            "total_movements": total,
            "porcentaje_cruzados": porcentaje_conciliados,
            "porcentaje_posibles": porcentaje_parciales,
            "sin_coincidencia": sin_coincidencia,
        }

    def _construir_alertas(self, match_results: list[MatchResult], resumen: dict[str, Any]) -> list[str]:
        alertas: list[str] = []
        conciliados = int(resumen["cruzados"])
        sin_coincidencia = int(resumen["sin_coincidencia"])

        if conciliados == 0:
            alertas.append("No se encontraron conciliaciones concluyentes entre PSE y cruces contables.")
        if sin_coincidencia:
            alertas.append(f"{sin_coincidencia} movimientos PSE quedaron sin coincidencia en cruces contables.")
        if any(result.state == "Pendiente de revisión" for result in match_results):
            alertas.append("Hay cruces PSE pendientes de revisión por ambigüedad o falta de unicidad.")

        return alertas
