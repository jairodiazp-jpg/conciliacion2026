"""
Test para validar el fix de detección PSE.
Verifica que se detecten cruces PSE incluso sin columna de descripción.
"""
from io import BytesIO
from openpyxl import Workbook
from pse_conciliador import PseConciliador


def test_pse_detection_without_description_column():
    """
    Prueba que _is_virtual_pse_row() detecta PSE en cualquier columna,
    incluso si no hay columna de descripción definida.
    """
    # Crear PSE simple
    pse_wb = Workbook()
    pse_ws = pse_wb.active
    pse_ws.title = "Movimientos"
    pse_ws.append(["Fecha", "Valor", "Concepto"])
    pse_ws.append(["2025-05-15", 50000, "Pago de servicios"])
    pse_ws.append(["2025-05-16", 75000, "Abono"])
    
    # Crear cruces CON PSE pero SIN columna descriptiva separada
    # Los datos PSE están en la columna "Concepto" (columna 3)
    cruces_wb = Workbook()
    cruces_ws = cruces_wb.active
    cruces_ws.title = "Cruces"
    cruces_ws.append(["Fecha", "Valor", "Concepto"])  # NO hay columna "Descripción"
    cruces_ws.append(["2025-05-15", 50000, "Pago virtual PSE"])  # ← Marcador en columna 3
    cruces_ws.append(["2025-05-16", 75000, "Pago virtual PSE"])
    
    pse_bytes = BytesIO()
    pse_wb.save(pse_bytes)
    pse_bytes.seek(0)
    
    cruces_bytes = BytesIO()
    cruces_wb.save(cruces_bytes)
    cruces_bytes.seek(0)
    
    # Procesar
    conciliador = PseConciliador(pse_bytes.getvalue(), cruces_bytes.getvalue())
    resultado = conciliador.procesar()
    
    # Validaciones
    print("=" * 60)
    print("TEST: Detección PSE sin columna de descripción")
    print("=" * 60)
    
    print(f"\nTotal movimientos PSE extraídos: {len(conciliador.pse_entries)}")
    print(f"Total cruces PSE extraídos: {len(conciliador.cruces_entries)}")
    print(f"Total dataset_cruces: {len(resultado.get('dataset_cruces', []))}")
    
    if len(conciliador.cruces_entries) > 0:
        print("\n✅ ÉXITO: Se detectaron cruces PSE correctamente")
        print(f"   Cruces extraídos:")
        for cruce in conciliador.cruces_entries:
            print(f"   - {cruce.sheet_name}:{cruce.row} = {cruce.value}")
    else:
        print("\n❌ FALLO: No se detectaron cruces PSE")
        return False
    
    if len(resultado.get('dataset_cruces', [])) > 0:
        print("\n✅ ÉXITO: dataset_cruces se llenó correctamente")
        print(f"   Registros en dataset_cruces:")
        for row in resultado.get('dataset_cruces', []):
            print(f"   - {row['sheet']}:{row['row']} Estado={row['estado_conciliacion']}")
    else:
        print("\n❌ FALLO: dataset_cruces está vacío")
        return False
    
    print("\n" + "=" * 60)
    print("RESULTADO FINAL: ✅ FIX VALIDADO EXITOSAMENTE")
    print("=" * 60)
    return True


if __name__ == "__main__":
    try:
        success = test_pse_detection_without_description_column()
        exit(0 if success else 1)
    except Exception as e:
        print(f"\n❌ ERROR en test: {e}")
        import traceback
        traceback.print_exc()
        exit(1)
