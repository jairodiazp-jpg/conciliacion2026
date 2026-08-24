import sys
import base64
from io import BytesIO
from openpyxl import load_workbook
from query_interno import QueryInterno

def main(file_path):
    try:
        # Cargar el archivo conciliado
        wb = load_workbook(file_path, data_only=False)
        
        # Ejecutar el Query Interno
        q = QueryInterno(wb)
        q.procesar()
        
        # Guardar el archivo
        wb.save(file_path)
        print(f"Procesamiento de Query Interno completado en {file_path}")
    except Exception as e:
        print(f"Error al procesar el archivo: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python run_query_interno.py <archivo_conciliado.xlsx>")
    else:
        main(sys.argv[1])
