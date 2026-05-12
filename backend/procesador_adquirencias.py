"""Procesador de Adquirencias: cruza valores de Adquirencias con Bancolombia 690."""

from __future__ import annotations

import base64
import re
import unicodedata
from datetime import date, datetime
from io import BytesIO
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.worksheet import Worksheet


LIGHT_BLUE_FILL = PatternFill(start_color="FFE8F4FF", end_color="FFE8F4FF", fill_type="solid")


class ProcesadorAdquirencias:
    def __init__(self, adquirencias_bytes: bytes, contable_b64: str, value_tolerance: float = 0.01) -> None:
        self.adquirencias_workbook = load_workbook(filename=BytesIO(adquirencias_bytes))
        self.contable_workbook = load_workbook(filename=BytesIO(base64.b64decode(contable_b64)))
        self.value_tolerance = value_tolerance
        self.logs: list[dict[str, Any]] = []
        self.adquirencia_counter = 1
        self.annotation_columns = self._index_annotation_columns()

    def procesar(self) -> str:
        """Procesa Adquirencias y retorna el archivo Contable actualizado en base64."""
        adquirencias_values = self._extraer_valores_adquirencias()
        self._cruzar_con_bancolombia(adquirencias_values)
        
        output_stream = BytesIO()
        self.contable_workbook.save(output_stream)
        return base64.b64encode(output_stream.getvalue()).decode("utf-8")

    def _normalizar_texto(self, value: str) -> str:
        normalized = unicodedata.normalize("NFKD", value)
        no_accents = "".join(char for char in normalized if not unicodedata.combining(char))
        return no_accents.strip().lower()

    def _index_annotation_columns(self) -> dict[str, tuple[int | None, int | None]]:
        indexed: dict[str, tuple[int | None, int | None]] = {}
        for sheet in self.contable_workbook.worksheets:
            comment_col: int | None = None
            observation_col: int | None = None

            max_scan_row = min(sheet.max_row, 25)
            for row in sheet.iter_rows(min_row=1, max_row=max_scan_row, min_col=1, max_col=sheet.max_column):
                for cell in row:
                    if not isinstance(cell.value, str):
                        continue
                    header = self._normalizar_texto(cell.value)

                    if comment_col is None and "coment" in header:
                        comment_col = cell.column
                    if observation_col is None and "observ" in header:
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

    def _parse_date(self, value: Any) -> date | None:
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

    def _extraer_valores_adquirencias(self) -> list[dict[str, Any]]:
        """Extrae valores de Adquirencias con fecha y autorizacion."""
        valores = []
        
        for sheet in self.adquirencias_workbook.worksheets:
            # Detectar columnas: valor, fecha, codigo autorizacion
            valor_col = None
            fecha_col = None
            auth_col = None
            
            max_scan = min(sheet.max_row, 25)
            for row in sheet.iter_rows(min_row=1, max_row=max_scan):
                for cell in row:
                    if not isinstance(cell.value, str):
                        continue
                    header = self._normalizar_texto(cell.value)
                    if valor_col is None and any(x in header for x in {"valor", "monto", "importe", "amount"}):
                        valor_col = cell.column
                    if fecha_col is None and any(x in header for x in {"fecha", "date", "transaction"}):
                        fecha_col = cell.column
                    if auth_col is None and any(x in header for x in {"autorizacion", "codigo", "authorization"}):
                        auth_col = cell.column
                if valor_col and fecha_col:
                    break
            
            if valor_col is None or fecha_col is None:
                continue
            
            # Extraer filas
            for row in sheet.iter_rows(min_row=max_scan + 1, max_row=sheet.max_row):
                try:
                    valor_cell = row[valor_col - 1] if valor_col - 1 < len(row) else None
                    fecha_cell = row[fecha_col - 1] if fecha_col - 1 < len(row) else None
                    
                    if valor_cell is None or fecha_cell is None:
                        continue
                    
                    valor = float(valor_cell.value or 0) if isinstance(valor_cell.value, (int, float)) else None
                    if valor is None or abs(valor) < 0.0001:
                        continue
                    
                    fecha = self._parse_date(fecha_cell.value)
                    if fecha is None:
                        continue
                    
                    auth_code = None
                    if auth_col is not None:
                        auth_cell = row[auth_col - 1] if auth_col - 1 < len(row) else None
                        if auth_cell is not None:
                            auth_code = str(auth_cell.value or "").strip()
                    
                    valores.append({
                        "sheet": sheet.title,
                        "row": row[0].row,
                        "valor": valor,
                        "fecha": fecha,
                        "autorizacion": auth_code,
                    })
                except Exception:
                    continue
        
        return valores

    def _cruzar_con_bancolombia(self, adquirencias: list[dict[str, Any]]) -> None:
        """Busca coincidencias en Bancolombia 690 y aplica coloring + anotaciones."""
        # Buscar hoja de Bancolombia 690
        bancolombia_sheet = None
        for sheet in self.contable_workbook.worksheets:
            if "1331" in sheet.title or "690" in sheet.title:
                bancolombia_sheet = sheet
                break
        
        if bancolombia_sheet is None:
            return
        
        # Buscar "Consignaciones sin registrar..." como marcador de inicio
        start_row = None
        marker_text = "consignaciones sin registrar"
        for row in bancolombia_sheet.iter_rows(min_row=1, max_row=bancolombia_sheet.max_row):
            for cell in row:
                if isinstance(cell.value, str) and marker_text in self._normalizar_texto(cell.value):
                    start_row = cell.row + 1
                    break
            if start_row:
                break
        
        if start_row is None:
            start_row = 2  # fallback
        
        # Detectar columna de valor
        valor_col = self._find_valor_column(bancolombia_sheet)
        
        # Procesar filas de Bancolombia desde el marcador
        for bancolombia_row in bancolombia_sheet.iter_rows(min_row=start_row, max_row=bancolombia_sheet.max_row):
            try:
                valor_cell = bancolombia_row[valor_col - 1] if valor_col - 1 < len(bancolombia_row) else None
                if valor_cell is None:
                    continue
                
                banco_valor = float(valor_cell.value or 0) if isinstance(valor_cell.value, (int, float)) else None
                if banco_valor is None or abs(banco_valor) < 0.0001:
                    continue
                
                # Buscar coincidencia en Adquirencias
                for adq in adquirencias:
                    if abs(abs(adq["valor"]) - abs(banco_valor)) <= self.value_tolerance:
                        # Encontró coincidencia
                        adq_tag = f"Adquirencia {self.adquirencia_counter}"
                        self.adquirencia_counter += 1
                        
                        # Aplicar coloring
                        valor_cell.fill = LIGHT_BLUE_FILL
                        
                        # Aplicar anotaciones recíprocas
                        self._anotar_cruce(bancolombia_sheet, bancolombia_row[0].row, adq_tag)
                        
                        self.logs.append({
                            "tipo": "adquirencia_cruzada",
                            "valor": round(banco_valor, 2),
                            "fecha": adq["fecha"].isoformat() if adq["fecha"] else None,
                            "confianza": 0.85,
                            "detalle": f"Adquirencia {self.adquirencia_counter - 1}: {banco_valor:,.2f} en {bancolombia_sheet.title}:fila {bancolombia_row[0].row}",
                        })
                        break
            except Exception:
                continue

    def _find_valor_column(self, sheet: Worksheet) -> int:
        """Detecta la columna de valor en una hoja."""
        candidates = ["valor", "monto", "importe", "amount", "consignacion"]
        for cell in sheet[1]:
            if isinstance(cell.value, str):
                h = self._normalizar_texto(cell.value)
                if any(c in h for c in candidates):
                    return cell.column
        # fallback
        return 3

    def _anotar_cruce(self, sheet: Worksheet, row_idx: int, tag: str) -> None:
        """Agrega anotación de cruce en comentario y observación."""
        comment_col, observation_col = self.annotation_columns.get(sheet.title, (None, None))
        
        if comment_col is not None:
            comment_cell = sheet.cell(row=row_idx, column=comment_col)
            comment_text = f"{tag} - Cruce con Adquirencias"
            current = str(comment_cell.value).strip() if comment_cell.value else ""
            if current and tag not in current:
                comment_cell.value = f"{current} | {comment_text}"
            elif not current:
                comment_cell.value = comment_text
        
        if observation_col is not None:
            observation_cell = sheet.cell(row=row_idx, column=observation_col)
            observation_text = f"{tag} encontrado en archivo de Adquirencias"
            current = str(observation_cell.value).strip() if observation_cell.value else ""
            if current and tag not in current:
                observation_cell.value = f"{current} | {observation_text}"
            elif not current:
                observation_cell.value = observation_text
