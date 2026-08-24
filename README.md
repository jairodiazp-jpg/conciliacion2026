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

## Validacion temporal reforzada

El motor ahora aplica una capa estricta de validacion antes de intentar cualquier match:

- No cruza años diferentes por defecto.
- No cruza meses diferentes por defecto.
- Calcula rango operativo minimo y maximo para bloquear fechas fuera del periodo cargado.
- Prioriza coincidencias por mismo dia, misma semana y mismo mes.
- Registra descartes temporales en los logs para auditar falsos positivos.

### Flags disponibles

- `temporal_tolerance_days`: tolerancia adicional de dias, por defecto `0`.
- `allow_previous_month`: permite cruce con el mes anterior cuando se habilita de forma explicita.
- `allow_previous_year`: permite cruce con el ano anterior cuando se habilita de forma explicita.

Estas opciones estan disponibles en el endpoint `POST /procesar` y en el runner local `backend/run_integrado_local.py`.

## Conciliacion PSE

Se agrego un flujo opcional y desacoplado para conciliar un archivo PSE contra un archivo de cruces contables sin alterar la estructura original del PSE.
Ahora el backend permite ejecutar en una sola llamada la conciliacion contable existente y la conciliacion PSE/banco cuando se envian los archivos correspondientes juntos.

### Salida

- Archivos conciliados:
  - PSE_CONCILIADO.xlsx
  - CRUCES_CONCILIADOS.xlsx
- Si tambien se envía el archivo contable tradicional, la respuesta incluye CONCILIACION_CONTABLE.xlsx en la misma corrida.
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
3. Crear un archivo `.env` con la API key y limites de seguridad
4. Levantar servidor

Ejemplo de `.env`:

API_KEY=tu-clave-super-secreta-aqui
RATE_LIMIT_REQUESTS=60
RATE_LIMIT_SECONDS=60

Comandos sugeridos:

cd backend
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000

## Seguridad de API

Los endpoints de carga de archivos requieren el header `X-API-Key`.
Ejemplo:

curl -H "X-API-Key: tu-clave-super-secreta-aqui" -F "file=@archivo.xlsx" http://localhost:8000/procesar

En produccion, usa una clave fuerte y configurala como secreto del entorno; no la subas al repositorio ni la dejes en el codigo.

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
