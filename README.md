# Conciliador App

Proyecto fullstack stateless para conciliacion contable de archivos Excel (.xlsx), con procesamiento 100% en memoria.

## Arquitectura

- Frontend: React + Vite
- Backend: FastAPI + openpyxl
- Persistencia: no usa base de datos
- Archivos: no guarda archivos en disco

## Estructura

- backend/main.py: API FastAPI y endpoint POST /procesar
- backend/conciliador.py: motor de conciliacion por fases
- frontend/: dashboard web para upload, procesamiento y resultados

## Reglas implementadas

1. Deteccion de secciones:
   - "sin registrar en libros" => UPPER
   - "sin registrar en el extracto" => LOWER
2. Restriccion de cruce:
   - Solo UPPER vs LOWER
3. Fases:
   - Aprobacion por numero exacto (naranja)
   - Valor + fecha ±3 dias (verde)
   - Uno a muchos por suma (verde)
   - Posibles cruces por mismo valor (amarillo)
4. Entre cuentas:
   - Cruza movimientos entre hojas del workbook
5. Preservacion de formato:
   - Solo colorea celda de valor y etiqueta columna siguiente

## Conciliacion PSE

Se agrego un flujo opcional y desacoplado para conciliar un archivo PSE contra un archivo de cruces contables sin alterar la estructura original del PSE.

### Salida

- Archivos conciliados:
  - PSE_CONCILIADO.xlsx
  - CRUCES_CONCILIADOS.xlsx
- Columnas nuevas al final en ambos archivos:
  - Estado_Conciliacion
  - Cuenta_Contable
  - Valores_Asociados
  - Comentario_Conciliacion
  - ID_Grupo_Conciliacion
- Dataset de conciliacion con detalle por fila, diferencias y grupo asignado

### Tolerancias

- Fecha: ±1 dia por defecto
- Valor: configurable desde la UI o el endpoint

### Ejemplo de datos

| Archivo | Fecha | Valor | Resultado |
| --- | --- | --- | --- |
| PSE | 2026-04-30 | 1.500.000 | Conciliado 1:N |
| Cruces contables | 2026-04-29 | 500.000 | Parte del grupo |
| Cruces contables | 2026-04-30 | 500.000 | Parte del grupo |
| Cruces contables | 2026-05-01 | 500.000 | Parte del grupo |

Comentario esperado:

- Conciliado con cuenta 12345 por valor total 1.500.000,00 (3 transacciones contables)
- Conciliación parcial - diferencia de 20.000,00
- Sin coincidencia en cruces contables

## Backend: ejecutar local

1. Ir a carpeta backend
2. Instalar dependencias
3. Levantar servidor

Comandos sugeridos:

cd backend
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000

## Frontend: ejecutar local

1. Ir a carpeta frontend
2. Instalar dependencias
3. Levantar app

Comandos sugeridos:

cd frontend
npm install
npm run dev

## Variable de entorno frontend

- VITE_API_URL
  - Ejemplo local: http://localhost:8000

## Deploy

### Frontend en Cloudflare Pages

- Build command: npm run build
- Output directory: dist
- Environment variable: VITE_API_URL=https://TU_BACKEND

### Backend en Railway/Render

- Runtime: Python
- Start command:
  uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}
- Root directory recomendado: backend

## Contrato de respuesta del endpoint /procesar

{
  "file": "<base64-del-excel>",
  "logs": [
    {
      "tipo": "aprobacion|valor_fecha|posible",
      "valor": 123.45,
      "fecha": "2026-04-23",
      "confianza": 0.95,
      "detalle": "..."
    }
  ],
  "resumen": {
    "cruzados": 10,
    "posibles": 3,
    "precision_estimada": 0.77
  },
  "alertas": ["..."]
}
