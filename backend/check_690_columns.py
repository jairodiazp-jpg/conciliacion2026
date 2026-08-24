"""Ver exactamente qué hay en annotation_columns_cont para la hoja 690."""

import base64
from io import BytesIO
from pathlib import Path
from procesador_adquirencias import ProcesadorAdquirencias

base = Path('..') / 'tmp_e2e'
adq = (base / 'ADQUIRENCIAS MAYO (1).xlsx').read_bytes()
memo = (base / 'CCS_Memorando Definitivo Ctas Bancarias Abril 2026.xlsx').read_bytes()

memo_b64 = base64.b64encode(memo).decode('utf-8')

processor = ProcesadorAdquirencias(adq, memo_b64)

print("annotation_columns_cont para 690:")
key_690 = None
for key in processor.annotation_columns_cont:
    if '690' in key:
        key_690 = key
        break

if key_690:
    comment_col, obs_col = processor.annotation_columns_cont[key_690]
    print(f"  Key: '{key_690}'")
    print(f"  comment_col={comment_col}, obs_col={obs_col}")
    
    # Leer los headers para verificar
    sheet = processor.contable_workbook[key_690]
    print(f"  Header en col {comment_col}: {sheet.cell(1, comment_col).value if comment_col else 'N/A'}")
    print(f"  Header en col {obs_col}: {sheet.cell(1, obs_col).value if obs_col else 'N/A'}")
else:
    print("  No se encontró clave con '690'")
