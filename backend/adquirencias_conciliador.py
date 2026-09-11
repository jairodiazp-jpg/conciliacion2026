from __future__ import annotations

import base64
import unicodedata
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from utils import parse_date


# ============================================================
# COLOR ADQUIRENCIAS
# ============================================================

GREEN_FILL = PatternFill(
    fill_type="solid",
    fgColor="FF00FF00",
)


class AdquirenciasConciliador:

    # ========================================================
    # ENCABEZADOS ESPERADOS
    # ========================================================

    ADQ_AUTH_HEADER = "CODIGO AUTORIZACION"
    ADQ_VALUE_HEADER = "VALOR TOTAL"
    ADQ_OBS_HEADER = "Observacion"

    CCS_AUTH_HEADER = "N° De Aprobación"
    CCS_VALUE_HEADER = "Valor"
    CCS_OBS_HEADER = "Observaciones"

    def __init__(
        self,
        adquirencias_bytes: bytes,
        ccs_bytes: bytes,
    ) -> None:

        self.adquirencias_wb = load_workbook(
            filename=BytesIO(adquirencias_bytes)
        )

        self.ccs_wb = load_workbook(
            filename=BytesIO(ccs_bytes)
        )

    # ========================================================
    # DETECTAR COLUMNAS Y FILAS DINÁMICAMENTE
    # ========================================================

    def _normalize_header(self, value: Any) -> str:
        if value is None:
            return ""
        text = str(value).strip().lower()
        text = unicodedata.normalize("NFKD", text)
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        return "".join(ch for ch in text if ch.isalnum() or ch.isspace()).strip()

    def _detectar_estructura(self, ws, headers: list[str]) -> dict[str, int]:
        """Busca encabezados en las primeras 30 filas.

        Usa comparación tolerante (normalizando mayúsculas, acentos y símbolos)
        para detectar encabezados que pueden variar levemente en los archivos del
        cliente. Si no se encuentran coincidencias exactas, se intenta una
        búsqueda por palabras clave (subcadena) como fallback.
        """
        def _normalize_text(s: str) -> str:
            return self._normalize_header(s)

        # palabras clave por encabezado para fallback por subcadena
        keywords_map = {}
        for h in headers:
            hn = h.lower()
            if "autoriz" in hn or "aprob" in hn or "aprobaci" in hn or "n°" in hn or "nro" in hn or "codigo" in hn:
                keywords_map[h] = ["autoriz", "aprob", "codigo", "nsu", "ref", "nro"]
            elif "valor" in hn or "monto" in hn or "importe" in hn:
                keywords_map[h] = ["valor", "monto", "importe", "amount"]
            elif "observ" in hn or "coment" in hn:
                keywords_map[h] = ["observ", "coment", "nota"]
            else:
                keywords_map[h] = [hn]

        normalized_headers = {h: _normalize_text(h) for h in headers}
        found: dict[str, int] = {}

        # Primera pasada: coincidencia exacta en versión normalizada
        for fila in range(1, min(ws.max_row, 30) + 1):
            for cell in ws[fila]:
                if cell.value is None:
                    continue
                val_norm = _normalize_text(cell.value)
                for h, h_norm in normalized_headers.items():
                    if val_norm == h_norm and h not in found:
                        found[h] = cell.column
            if len(found) == len(headers):
                break

        # Fallback: búsqueda por subcadena usando keywords_map
        if len(found) < len(headers):
            for fila in range(1, min(ws.max_row, 30) + 1):
                for cell in ws[fila]:
                    if cell.value is None:
                        continue
                    val_norm = _normalize_text(cell.value)
                    for h, keys in keywords_map.items():
                        if h in found:
                            continue
                        for key in keys:
                            if key and key in val_norm:
                                found[h] = cell.column
                                break
                        if h in found:
                            break
                if len(found) == len(headers):
                    break

        return found

    def _detectar_inicio_datos(self, ws, auth_col: int, *, fallback: int = 19) -> int:
        """Busca la primera fila después del encabezado que tenga datos.

        Mejor manejo de fallback: si no se detecta encabezado textual, detectar la
        primera fila con un valor en la columna de autorización dentro de las
        primeras 30 filas y considerarla fila de datos. Si nada aparece, usar el
        fallback original.
        """
        encabezados_auth = {
            "n° de aprobación",
            "n de aprobacion",
            "codigo autorizacion",
            "código autorización",
        }
        max_scan = min(ws.max_row, 30) + 5
        for fila in range(1, max_scan):
            val = ws.cell(row=fila, column=auth_col).value
            if val is not None and str(val).strip().lower() in encabezados_auth:
                return fila + 1

        # Si no se encontró un encabezado textual, intentar detectar la primer fila
        # con datos en la columna de autorización dentro de las primeras filas.
        for fila in range(1, min(ws.max_row, 30) + 1):
            val = ws.cell(row=fila, column=auth_col).value
            if val is not None and str(val).strip() != "":
                # Si el valor detectado parece un encabezado (por ejemplo contiene letras),
                # intentar considerarlo encabezado y devolver la siguiente fila.
                text = str(val).strip().lower()
                if any(ch.isalpha() for ch in text) and len(text) > 2:
                    return fila + 1
                # En caso contrario, asumir que es la primera fila de datos.
                return fila

        return fallback

    # ========================================================
    # NORMALIZAR AUTORIZACIÓN
    # ========================================================

    def _normalizar_auth(self, value: Any) -> str:

        if value is None:
            return ""

        text = str(value)

        text = (
            text
            .replace("\u00A0", " ")
            .replace("\u200B", "")
            .replace("\u200C", "")
            .replace("\u200D", "")
            .replace("\ufeff", "")
        )

        text = text.strip()
        if text.endswith(".0") and text.replace(".", "", 1).isdigit():
            text = text[:-2]
        return text

    def _normalizar_fecha_igualdad(self, value: Any) -> str | None:
        parsed = parse_date(value)
        if parsed is not None:
            return parsed.isoformat()
        return None

    def _columnas_fecha(self, ws, preferencias: tuple[str, ...]) -> list[int]:
        """Detecta columnas de fecha ordenadas por prioridad de encabezado."""
        scored: list[tuple[int, int]] = []
        seen: set[int] = set()
        for fila in range(1, min(ws.max_row, 30) + 1):
            for cell in ws[fila]:
                if cell.value is None or cell.column in seen:
                    continue
                header = self._normalize_header(cell.value)
                if "fecha" not in header and "date" not in header:
                    continue
                if "canje" in header:
                    priority = 80
                else:
                    priority = 50
                    for index, preferencia in enumerate(preferencias):
                        if preferencia in header:
                            priority = index
                            break
                scored.append((priority, cell.column))
                seen.add(cell.column)
        scored.sort()
        preferred = [column for priority, column in scored if priority < 80]
        return preferred or [column for _, column in scored]

    def _fechas_fila(self, ws, row: int, date_cols: list[int]) -> list[str]:
        fechas: list[str] = []
        for column in date_cols:
            parsed = self._normalizar_fecha_igualdad(ws.cell(row=row, column=column).value)
            if parsed and parsed not in fechas:
                fechas.append(parsed)
        return fechas

    # ========================================================
    # NORMALIZAR VALOR
    # ========================================================

    def _normalizar_valor(
        self,
        value: Any,
    ) -> Decimal:

        if value is None:
            return Decimal("0.00")

        if isinstance(value, Decimal):

            return value.quantize(
                Decimal("0.01"),
                rounding=ROUND_HALF_UP,
            )

        if isinstance(value, int):

            return Decimal(value).quantize(
                Decimal("0.01"),
                rounding=ROUND_HALF_UP,
            )

        if isinstance(value, float):

            return Decimal(str(value)).quantize(
                Decimal("0.01"),
                rounding=ROUND_HALF_UP,
            )

        text = str(value).strip()

        if not text:
            return Decimal("0.00")

        text = (
            text
            .replace("$", "")
            .replace("COP", "")
            .replace("cop", "")
            .replace("\u00A0", "")
            .replace(" ", "")
        )

        negativo = False

        if text.startswith("(") and text.endswith(")"):

            negativo = True
            text = text[1:-1]

        if text.startswith("-"):

            negativo = True
            text = text[1:]

        # ---------------------------------------------
        # FORMATO 1.250.000,50
        # ---------------------------------------------

        if "," in text and "." in text:

            if text.rfind(",") > text.rfind("."):

                text = text.replace(".", "")
                text = text.replace(",", ".")

            else:

                text = text.replace(",", "")

        # ---------------------------------------------
        # FORMATO 1250000,50
        # ---------------------------------------------

        elif "," in text:

            partes = text.split(",")

            if len(partes[-1]) <= 2:

                text = (
                    "".join(partes[:-1])
                    + "."
                    + partes[-1]
                )

            else:

                text = "".join(partes)

        # ---------------------------------------------
        # FORMATO 1.250.000
        # ---------------------------------------------

        elif "." in text:

            partes = text.split(".")

            if len(partes) > 2:

                text = "".join(partes)

        try:

            resultado = Decimal(text)

            if negativo:
                resultado = -resultado

            return resultado.quantize(
                Decimal("0.01"),
                rounding=ROUND_HALF_UP,
            )

        except (InvalidOperation, ValueError):

            return Decimal("0.00")

    # ========================================================
    # VALOR EN CENTAVOS
    # ========================================================

    def _valor_key(
        self,
        value: Any,
    ) -> int:

        valor = self._normalizar_valor(value)

        return int(
            (
                valor * Decimal("100")
            ).quantize(
                Decimal("1"),
                rounding=ROUND_HALF_UP,
            )
        )

    # ========================================================
    # BUSCAR HOJA CCS
    # ========================================================

    def _buscar_hoja_ccs(self):

        for ws in self.ccs_wb.worksheets:
            nombre = str(ws.title).strip().lower()
            if "2490" in nombre:
                return ws

        raise ValueError(
            "No se encontró la hoja correspondiente a la cuenta CCS 2490. "
            f"Hojas disponibles: {self.ccs_wb.sheetnames}"
        )

    # ========================================================
    # DETECTAR COLUMNAS CCS
    # ========================================================

    def _detectar_columna_cuenta(self, ws) -> int | None:
        nombres = {
            "cuenta",
            "cta",
            "cuenta ccs",
            "cuenta bancaria",
            "numero cuenta",
        }
        for fila in range(1, min(ws.max_row, 30) + 1):
            for cell in ws[fila]:
                if cell.value is None:
                    continue
                nombre = str(cell.value).strip().lower()
                if nombre in nombres or "cuenta" in nombre:
                    return cell.column
        return None

    def _detectar_columnas_ccs(self, ws) -> tuple[int, int]:
        return self.CCS_AUTH_COL, self.CCS_VALUE_COL

    # ========================================================
    # BUSCAR COLUMNA POR ENCABEZADO
    # ========================================================

    def _buscar_columna(
        self,
        ws,
        nombres: set[str],
        filas: tuple[int, ...],
    ) -> int | None:

        for fila in filas:

            if fila > ws.max_row:
                continue

            for cell in ws[fila]:

                if cell.value is None:
                    continue

                nombre = (
                    str(cell.value)
                    .strip()
                    .lower()
                )

                if nombre in nombres:

                    return cell.column

        return None

    # ========================================================
    # ENCONTRAR COLUMNA OBSERVACIONES ADQUIRENCIAS
    # ========================================================

    def _obtener_columna_observaciones_adq(
        self,
        ws,
    ) -> int:

        nombres = {
            "observacion",
            "observaciones",
        }

        # Primero buscar por encabezado
        for fila in range(1, min(ws.max_row, 10) + 1):

            for cell in ws[fila]:

                if cell.value is None:
                    continue

                texto = self._normalize_header(cell.value)

                if texto in nombres or "observ" in texto:

                    return cell.column

        # Si no existe, crearla al final
        nueva_columna = ws.max_column + 1

        ws.cell(
            row=1,
            column=nueva_columna,
        ).value = "OBSERVACIONES"

        return nueva_columna

    # ========================================================
    # ENCONTRAR COLUMNA OBSERVACIONES ADQUIRENCIAS (AMPLIADO)
    # ========================================================

    def _obtener_columna_observaciones_adq_ampliado(
        self,
        ws,
    ) -> int:

        nombres = {
            "observacion",
            "observaciones",
        }

        # Buscar en las primeras 30 filas
        for fila in range(1, min(ws.max_row, 30) + 1):

            for cell in ws[fila]:

                if cell.value is None:
                    continue

                texto = self._normalize_header(cell.value)

                if texto in nombres or "observ" in texto:

                    return cell.column

        # Si no existe, crearla al final
        nueva_columna = ws.max_column + 1

        ws.cell(
            row=1,
            column=nueva_columna,
        ).value = "OBSERVACIONES"

        return nueva_columna

    # ========================================================
    # ENCONTRAR COLUMNA OBSERVACIONES CCS
    # ========================================================

    def _obtener_columna_observaciones_ccs(
        self,
        ws,
        ccs_auth_col: int,
        ccs_val_col: int,
    ) -> tuple[int, int]:

        nombres = {
            "observacion",
            "observaciones",
        }

        # Buscar en las filas donde normalmente
        # se encuentra el encabezado.
        for fila in range(
            1,
            min(ws.max_row, 30) + 1,
        ):

            for cell in ws[fila]:

                if cell.value is None:
                    continue

                texto = self._normalize_header(cell.value)

                if texto in nombres or "observ" in texto:

                    return cell.column, fila

        # Si no existe, crearla al final
        columna = ws.max_column + 1

        # Buscar la fila donde están Auth y Valor
        header_row = 1

        for fila in range(
            1,
            min(ws.max_row, 30) + 1,
        ):

            auth_header = str(
                ws.cell(
                    row=fila,
                    column=ccs_auth_col,
                ).value
                or ""
            ).strip().lower()

            value_header = str(
                ws.cell(
                    row=fila,
                    column=ccs_val_col,
                ).value
                or ""
            ).strip().lower()

            if (
                "aprob" in auth_header
                or "autoriz" in auth_header
            ) and "valor" in value_header:

                header_row = fila
                break

        ws.cell(
            row=header_row,
            column=columna,
        ).value = "OBSERVACIONES"

        return columna, header_row

    # ========================================================
    # ENCONTRAR PRIMERA FILA DE DATOS CCS
    # ========================================================

    def _obtener_inicio_ccs(
        self,
        ws,
        ccs_auth_col: int,
        ccs_val_col: int,
    ) -> int:
        return self.CCS_START_ROW

    # ========================================================
    # AGREGAR OBSERVACIÓN
    # ========================================================

    def _agregar_observacion(
        self,
        ws,
        row: int,
        col: int,
        texto: str,
    ) -> None:

        actual = str(
            ws.cell(
                row=row,
                column=col,
            ).value
            or ""
        ).strip()

        if not actual:

            ws.cell(
                row=row,
                column=col,
            ).value = texto

            return

        if texto not in actual:

            ws.cell(
                row=row,
                column=col,
            ).value = (
                actual
                + " | "
                + texto
            )

    # ========================================================
    # PROCESAR
    # ========================================================

    def procesar(self) -> dict[str, Any]:

        # ====================================================
        # HOJAS
        # ====================================================

        adq_sheet = self.adquirencias_wb.active

        ccs_sheet = self._buscar_hoja_ccs()

        # ====================================================
        # COLUMNAS
        # ====================================================

        adq_struct = self._detectar_estructura(adq_sheet, [self.ADQ_AUTH_HEADER, self.ADQ_VALUE_HEADER, self.ADQ_OBS_HEADER])
        adq_auth_col = adq_struct.get(self.ADQ_AUTH_HEADER, 23)
        adq_val_col = adq_struct.get(self.ADQ_VALUE_HEADER, 16)
        adq_obs_col = adq_struct.get(self.ADQ_OBS_HEADER) or self._obtener_columna_observaciones_adq_ampliado(adq_sheet)
        adq_date_cols = self._columnas_fecha(
            adq_sheet,
            (
                "fecha de transaccion",
                "fecha transaccion",
                "fecha de compensacion",
                "fecha compensacion",
                "fecha documento",
            ),
        )

        ccs_struct = self._detectar_estructura(ccs_sheet, [self.CCS_AUTH_HEADER, self.CCS_VALUE_HEADER, self.CCS_OBS_HEADER])
        ccs_auth_col = ccs_struct.get(self.CCS_AUTH_HEADER, 6)
        ccs_val_col = ccs_struct.get(self.CCS_VALUE_HEADER, 8)
        ccs_obs_col, ccs_header_row = self._obtener_columna_observaciones_ccs(ccs_sheet, ccs_auth_col, ccs_val_col)
        ccs_date_cols = self._columnas_fecha(
            ccs_sheet,
            (
                "fecha documento",
                "fecha de transaccion",
                "fecha transaccion",
                "fecha contabilizacion",
            ),
        )

        # ====================================================
        # INICIO DATOS CCS
        # ====================================================

        ccs_start_row = self._detectar_inicio_datos(ccs_sheet, ccs_auth_col)
        adq_start_row = self._detectar_inicio_datos(adq_sheet, adq_auth_col, fallback=2)

        # ====================================================
        # ÍNDICE CCS
        # ====================================================

        ccs_index: dict[
            tuple[str, str, int],
            list[int],
        ] = {}

        cantidad_ccs = 0

        for row in range(
            ccs_start_row,
            ccs_sheet.max_row + 1,
        ):


            auth = self._normalizar_auth(
                ccs_sheet.cell(
                    row=row,
                    column=ccs_auth_col,
                ).value
            )
            if not auth:
                continue

            raw_value = ccs_sheet.cell(
                row=row,
                column=ccs_val_col,
            ).value
            if raw_value is None:
                continue

            fechas_ccs = self._fechas_fila(ccs_sheet, row, ccs_date_cols)
            if not fechas_ccs:
                continue

            value_key = self._valor_key(raw_value)
            for fecha_value in fechas_ccs:
                key = (
                    auth,
                    fecha_value,
                    value_key,
                )
                ccs_index.setdefault(
                    key,
                    [],
                ).append(row)

            cantidad_ccs += 1

        # ====================================================
        # PROCESAR ADQUIRENCIAS
        # ====================================================

        cruce_count = 0
        cantidad_adq = 0
        cantidad_auth = 0
        used_ccs_rows: set[int] = set()
        dataset_adquirencias: list[dict[str, Any]] = []

        for row in range(
            adq_start_row,
            adq_sheet.max_row + 1,
        ):

            cantidad_adq += 1

            auth = self._normalizar_auth(
                adq_sheet.cell(
                    row=row,
                    column=adq_auth_col,
                ).value
            )

            if not auth:
                continue

            cantidad_auth += 1

            raw_value = adq_sheet.cell(
                row=row,
                column=adq_val_col,
            ).value
            if raw_value is None:
                continue

            fechas_adq = self._fechas_fila(adq_sheet, row, adq_date_cols)
            if not fechas_adq:
                continue

            value = self._normalizar_valor(raw_value)
            value_key = self._valor_key(raw_value)

            # =================================================
            # CRUCE REAL
            #
            # AUTORIZACIÓN + FECHA EXACTA + VALOR EXACTO + CUENTA 2490
            # =================================================

            matches: list[int] = []
            fecha_adq = fechas_adq[0]
            for fecha_candidata in fechas_adq:
                candidatos = [
                    ccs_row
                    for ccs_row in ccs_index.get((auth, fecha_candidata, value_key), [])
                    if ccs_row not in used_ccs_rows
                ]
                if candidatos:
                    fecha_adq = fecha_candidata
                    matches = candidatos
                    break

            if not matches:
                continue

            # =================================================
            # NUEVA ADQUIRENCIA
            # =================================================

            cruce_count += 1
            ccs_row_match = matches[0]
            used_ccs_rows.add(ccs_row_match)
            matches = [ccs_row_match]

            adquirencia_nombre = (
                f"ADQUIRENCIA {cruce_count}"
            )
            es_duplicado = len(
                [
                    ccs_row
                    for ccs_row in ccs_index.get((auth, fecha_adq, value_key), [])
                    if ccs_row != ccs_row_match
                ]
            ) > 0

            # =================================================
            # PINTAR ADQUIRENCIA
            # =================================================

            for cell in adq_sheet[row]:

                cell.fill = GREEN_FILL

            # =================================================
            # OBSERVACIÓN ADQUIRENCIA
            # =================================================

            filas_ccs = ", ".join(
                str(x)
                for x in matches
            )

            detalle_duplicado = (
                " | VALIDAR DUPLICADO: varios candidatos CCS 2490 para la misma aprobación, fecha y valor"
                if es_duplicado
                else ""
            )

            observacion_adq = (
                f"{adquirencia_nombre} encontrado en {ccs_sheet.title.strip()}:fila {filas_ccs}. "
                f"Autorización: {auth} | Fecha: {fecha_adq} | Valor: {value:.2f}{detalle_duplicado}"
            )

            self._agregar_observacion(
                adq_sheet,
                row,
                adq_obs_col,
                observacion_adq,
            )

            # =================================================
            # AGREGAR AL DATASET INTERNO (PARA LOGS)
            # =================================================

            dataset_adquirencias.append(
                {
                    "tipo": "adquirencia_cruzada",
                    "valor": float(value),
                    "fecha": fecha_adq,
                    "confianza": 0.95,
                    "detalle": observacion_adq,
                    "adquirencias_row": row,
                    "ccs_rows": matches,
                    "hoja_ccs": ccs_sheet.title,
                }
            )

            # =================================================
            # MARCAR CCS
            # =================================================

            for ccs_row in matches:

                # ---------------------------------------------
                # Pintar la fila correspondiente en CCS
                # ---------------------------------------------

                for cell in ccs_sheet[ccs_row]:

                    cell.fill = GREEN_FILL

                # ---------------------------------------------
                # Observación CCS
                # ---------------------------------------------

                observacion_ccs = (
                    f"{adquirencia_nombre} encontrado en Adquirencias ({adq_sheet.title}):fila {row}. "
                    f"Autorización: {auth} | Fecha: {fecha_adq} | Valor: {value:.2f}{detalle_duplicado}"
                )

                self._agregar_observacion(
                    ccs_sheet,
                    ccs_row,
                    ccs_obs_col,
                    observacion_ccs,
                )

        # ====================================================
        # GUARDAR ADQUIRENCIAS
        # ====================================================

        out_adq = BytesIO()

        self.adquirencias_wb.save(
            out_adq
        )

        adquirencias_result = (
            out_adq.getvalue()
        )

        # ====================================================
        # GUARDAR CCS
        # ====================================================

        out_ccs = BytesIO()

        self.ccs_wb.save(
            out_ccs
        )

        ccs_result = (
            out_ccs.getvalue()
        )

        # ====================================================
        # VERIFICACIÓN DEL ARCHIVO FINAL
        # ====================================================

        # Volver a abrir los archivos guardados para
        # garantizar que las modificaciones realmente
        # quedaron persistidas.

        verificacion_adq = load_workbook(
            filename=BytesIO(
                adquirencias_result
            )
        )

        verificacion_ccs = load_workbook(
            filename=BytesIO(
                ccs_result
            )
        )

        ver_adq_sheet = (
            verificacion_adq.active
        )

        ver_ccs_sheet = (
            verificacion_ccs[
                ccs_sheet.title
            ]
        )

        observaciones_adq_final = 0
        filas_verdes_adq = 0

        for row in range(
            adq_start_row,
            ver_adq_sheet.max_row + 1,
        ):

            obs = (
                ver_adq_sheet.cell(
                    row=row,
                    column=adq_obs_col,
                ).value
            )

            if obs and "ADQUIRENCIA" in str(obs):

                observaciones_adq_final += 1

                # Revisar si alguna celda de la fila
                # tiene el verde configurado.
                for cell in ver_adq_sheet[row]:

                    if (
                        cell.fill
                        and cell.fill.fill_type == "solid"
                        and cell.fill.fgColor.rgb
                        and cell.fill.fgColor.rgb.upper()
                        in {
                            "FF00FF00",
                            "0000FF00",
                        }
                    ):

                        filas_verdes_adq += 1
                        break

        observaciones_ccs_final = 0
        filas_verdes_ccs = 0

        for row in range(
            ccs_start_row,
            ver_ccs_sheet.max_row + 1,
        ):

            obs = (
                ver_ccs_sheet.cell(
                    row=row,
                    column=ccs_obs_col,
                ).value
            )

            if obs and "ADQUIRENCIA" in str(obs):

                observaciones_ccs_final += 1

                for cell in ver_ccs_sheet[row]:

                    if (
                        cell.fill
                        and cell.fill.fill_type == "solid"
                        and cell.fill.fgColor.rgb
                        and cell.fill.fgColor.rgb.upper()
                        in {
                            "FF00FF00",
                            "0000FF00",
                        }
                    ):

                        filas_verdes_ccs += 1
                        break

        # ====================================================
        # RESULTADO
        # ====================================================

        result = {

            "adquirencias_file": (
                base64.b64encode(
                    adquirencias_result
                ).decode("utf-8")
            ),

            "ccs_file": (
                base64.b64encode(
                    ccs_result
                ).decode("utf-8")
            ),

            "resumen": {

                "registros_adquirencias":
                    cantidad_adq,

                "autorizaciones_adquirencias":
                    cantidad_auth,

                "registros_ccs":
                    cantidad_ccs,

                "cruzados":
                    cruce_count,

                "observaciones_adquirencias":
                    observaciones_adq_final,

                "filas_verdes_adquirencias":
                    filas_verdes_adq,

                "observaciones_ccs":
                    observaciones_ccs_final,

                "filas_verdes_ccs":
                    filas_verdes_ccs,

                "hoja_ccs":
                    ccs_sheet.title,

                "criterio_cruce":
                    "AUTORIZACION + FECHA TRANSACCION/DOCUMENTO + VALOR EXACTO + CUENTA 2490",

                "adquirencias_autorizacion":
                    "W",

                "adquirencias_valor":
                    "P",

                "ccs_autorizacion":
                    "F",

                "ccs_valor":
                    "H",
            },
        }

        result["dataset"] = dataset_adquirencias

        return result