"""Update integrador.py to handle new Adquirencias return format."""

from pathlib import Path

content = Path("integrador.py").read_text(encoding="utf-8")

# Reemplazo 1: Sección contable-only (línea ~290-296)
old_1 = """            if self.adquirencias_bytes is not None:
                try:
                    procesador_adq = ProcesadorAdquirencias(
                        self.adquirencias_bytes,
                        contable_result["file"],
                        value_tolerance=self.value_tolerance,
                    )
                    contable_result["file"] = procesador_adq.procesar()
                    logs.extend(procesador_adq.logs)
                except Exception as e:
                    alertas.append(f"Fallo en procesamiento de Adquirencias: {e}")
            return contable_result"""

new_1 = """            if self.adquirencias_bytes is not None:
                try:
                    procesador_adq = ProcesadorAdquirencias(
                        self.adquirencias_bytes,
                        contable_result["file"],
                        value_tolerance=self.value_tolerance,
                    )
                    adq_result = procesador_adq.procesar()
                    contable_result["file"] = adq_result["contable_file"]
                    logs.extend(procesador_adq.logs)
                    contable_result_with_adq = {
                        "mode": "contable-only",
                        "contable": contable_result,
                        "files": [
                            {
                                "name": "CONCILIACION_CONTABLE.xlsx",
                                "file": adq_result["contable_file"],
                            },
                            {
                                "name": "ADQUIRENCIAS_PROCESADAS.xlsx",
                                "file": adq_result["adquirencias_file"],
                            },
                        ],
                        "logs": logs,
                        "alertas": alertas,
                        "resumen": contable_result.get("resumen", {}),
                    }
                    return contable_result_with_adq
                except Exception as e:
                    alertas.append(f"Fallo en procesamiento de Adquirencias: {e}")
            return contable_result"""

content = content.replace(old_1, new_1)

# Reemplazo 2: Sección integrada (línea ~342-348)
old_2 = """        # Fase adicional: Procesar Adquirencias si se proporciona
        if self.adquirencias_bytes is not None:
            try:
                procesador_adq = ProcesadorAdquirencias(
                    self.adquirencias_bytes,
                    contable_result["file"],
                    value_tolerance=self.value_tolerance,
                )
                contable_result["file"] = procesador_adq.procesar()
                logs.extend(procesador_adq.logs)
            except Exception as e:
                alertas.append(f"Fallo en procesamiento de Adquirencias: {e}")

        files: list[dict[str, str]] = [
            {
                "name": "CONCILIACION_CONTABLE.xlsx",
                "file": contable_result["file"],
            },"""

new_2 = """        # Fase adicional: Procesar Adquirencias si se proporciona
        adquirencias_file_b64 = None
        if self.adquirencias_bytes is not None:
            try:
                procesador_adq = ProcesadorAdquirencias(
                    self.adquirencias_bytes,
                    contable_result["file"],
                    value_tolerance=self.value_tolerance,
                )
                adq_result = procesador_adq.procesar()
                contable_result["file"] = adq_result["contable_file"]
                adquirencias_file_b64 = adq_result["adquirencias_file"]
                logs.extend(procesador_adq.logs)
            except Exception as e:
                alertas.append(f"Fallo en procesamiento de Adquirencias: {e}")

        files: list[dict[str, str]] = [
            {
                "name": "CONCILIACION_CONTABLE.xlsx",
                "file": contable_result["file"],
            },"""

content = content.replace(old_2, new_2)

# Reemplazo 3: Agregar archivo de adquirencias a la lista de retorno
old_3 = """        files: list[dict[str, str]] = [
            {
                "name": "CONCILIACION_CONTABLE.xlsx",
                "file": contable_result["file"],
            },
            {
                "name": pse_result.get("output_name", "PSE_CONCILIADO.xlsx"),
                "file": pse_result["file"],
            },
        ]"""

new_3 = """        files: list[dict[str, str]] = [
            {
                "name": "CONCILIACION_CONTABLE.xlsx",
                "file": contable_result["file"],
            },
            {
                "name": pse_result.get("output_name", "PSE_CONCILIADO.xlsx"),
                "file": pse_result["file"],
            },
        ]
        if adquirencias_file_b64 is not None:
            files.append({
                "name": "ADQUIRENCIAS_PROCESADAS.xlsx",
                "file": adquirencias_file_b64,
            })"""

content = content.replace(old_3, new_3)

Path("integrador.py").write_text(content, encoding="utf-8")
print("OK")
