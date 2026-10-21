import unicodedata
from datetime import date, datetime
from typing import Any, Dict, List, Tuple

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.worksheet import Worksheet

# Colores
BROWN_FILL = PatternFill(start_color="FFCC9966", end_color="FFCC9966", fill_type="solid")
GRAY_FILL = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")

class QueryInterno:
    def __init__(self, workbook: Any) -> None:
        self.workbook = workbook
        self.entries: List[Dict[str, Any]] = []
        self.cross_id = 1
        self.duplicate_id = 1

    def _normalizar_texto(self, value: str) -> str:
        normalized = unicodedata.normalize("NFKD", str(value) if value else "")
        no_accents = "".join(char for char in normalized if not unicodedata.combining(char))
        return no_accents.strip().lower()

    def _parse_amount(self, value: Any) -> float | None:
        if value is None: return None
        try:
            if isinstance(value, (int, float)): return float(value)
            s = str(value).replace("$", "").replace(".", "").replace(",", ".").strip()
            return float(s)
        except: return None

    def _parse_date(self, value: Any) -> date | None:
        if isinstance(value, date): return value
        if isinstance(value, datetime): return value.date()
        if isinstance(value, str):
            try:
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try: return datetime.strptime(value, fmt).date()
                    except: continue
            except: pass
        return None

    def _detectar_columnas(self, sheet: Worksheet) -> Tuple[int, int, int, int]:
        valor_col, fecha_col, obs_col = None, None, None
        for row in sheet.iter_rows(min_row=1, max_row=5):
            for cell in row:
                header = self._normalizar_texto(str(cell.value))
                if "valor" in header and valor_col is None: valor_col = cell.column
                if "fecha" in header and fecha_col is None: fecha_col = cell.column
                if "observ" in header and obs_col is None: obs_col = cell.column
        
        if valor_col is None: valor_col = 4
        if fecha_col is None: fecha_col = 2
        if obs_col is None: obs_col = sheet.max_column
        
        return valor_col, fecha_col, obs_col, 1

    def procesar(self) -> None:
        self.entries = []
        for sheet in self.workbook.worksheets:
            valor_col, fecha_col, obs_col, header_row = self._detectar_columnas(sheet)
            
            for row in sheet.iter_rows(min_row=header_row + 1, max_row=sheet.max_row):
                valor = self._parse_amount(row[valor_col-1].value)
                fecha = self._parse_date(row[fecha_col-1].value)
                
                if valor is not None and fecha is not None:
                    self.entries.append({
                        "sheet": sheet,
                        "row": row[0].row,
                        "valor": round(valor, 2),
                        "fecha": fecha,
                        "obs_col": obs_col
                    })

        index: Dict[Tuple[date, float], List[Dict[str, Any]]] = {}
        for entry in self.entries:
            key = (entry["fecha"], entry["valor"])
            index.setdefault(key, []).append(entry)
        
        for key, matches in index.items():
            if len(matches) < 2: continue
            
            by_sheet: Dict[str, List[Dict[str, Any]]] = {}
            for m in matches:
                by_sheet.setdefault(m["sheet"].title, []).append(m)
            
            if len(by_sheet) < 2: continue
            
            if len(matches) > 2 or len(by_sheet) > 2:
                tag = f"DUPLICADO {self.duplicate_id}"
                self.duplicate_id += 1
                for m in matches:
                    self._marcar(m, tag, GRAY_FILL, is_duplicate=True, all_matches=matches)
            else:
                tag = f"CRUCE QUERY {self.cross_id}"
                self.cross_id += 1
                for m in matches:
                    self._marcar(m, tag, BROWN_FILL, is_duplicate=False, all_matches=matches)

    def _marcar(self, entry: Dict[str, Any], tag: str, fill: PatternFill, is_duplicate: bool, all_matches: List[Dict[str, Any]]) -> None:
        sheet = entry["sheet"]
        row = entry["row"]
        
        for cell in sheet[row]:
            cell.fill = fill
            
        obs_cell = sheet.cell(row=row, column=entry["obs_col"])
        current = str(obs_cell.value) if obs_cell.value else ""
        
        if is_duplicate:
            coincidencias = "; ".join([f"{m['sheet'].title} Fila {m['row']}" for m in all_matches if m != entry])
            text = f"{tag} | Valor: ${entry['valor']} | Fecha: {entry['fecha']} | Coincidencias: {coincidencias}"
        else:
            other = [m for m in all_matches if m != entry][0]
            text = f"{tag} | Valor: ${entry['valor']} | Fecha: {entry['fecha']} | Cuenta encontrada: {other['sheet'].title} Fila {other['row']}"
            
        if current and "CRUCE" not in current and "DUPLICADO" not in current:
            obs_cell.value = f"{current} | {text}"
        else:
            obs_cell.value = text
