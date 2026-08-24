"""Verificar que los títulos de hojas coincidan exactamente."""

import base64
from io import BytesIO
from pathlib import Path
from openpyxl import load_workbook
from procesador_adquirencias import ProcesadorAdquirencias

base = Path('..') / 'tmp_e2e'
adq = (base / 'ADQUIRENCIAS MAYO (1).xlsx').read_bytes()
memo = (base / 'CCS_Memorando Definitivo Ctas Bancarias Abril 2026.xlsx').read_bytes()

memo_b64 = base64.b64encode(memo).decode('utf-8')

processor = ProcesadorAdquirencias(adq, memo_b64)

print("=== TÍTULOS DE HOJAS ===\n")
print("annotation_columns_cont:")
for title in processor.annotation_columns_cont:
    print(f"  '{title}'")

print("\nannotation_columns_adq:")
for title in processor.annotation_columns_adq:
    print(f"  '{title}'")

print("\nAdquirencias workbook sheets:")
for sheet in processor.adquirencias_workbook.worksheets:
    print(f"  '{sheet.title}'")

print("\nContable workbook sheets:")
for sheet in processor.contable_workbook.worksheets:
    print(f"  '{sheet.title}'")
