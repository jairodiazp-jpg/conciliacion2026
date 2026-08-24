from __future__ import annotations

import os
import time
from collections import defaultdict, deque

from dotenv import load_dotenv
from fastapi import Depends, FastAPI, File, Form, HTTPException, Security, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import APIKeyHeader
from starlette.requests import Request
from starlette.responses import JSONResponse

load_dotenv()

from conciliador import ConciliadorContable
from integrador import ProcesadorIntegrado
from version import get_app_version
from validacion_temporal import ValidacionTemporalConfig


APP_VERSION = get_app_version()
API_KEY = os.getenv("API_KEY")
RATE_LIMIT_REQUESTS = int(os.getenv("RATE_LIMIT_REQUESTS", "60"))
RATE_LIMIT_SECONDS = int(os.getenv("RATE_LIMIT_SECONDS", "60"))

if not API_KEY:
    raise RuntimeError("La variable de entorno API_KEY no esta configurada. Define una clave fuerte antes de iniciar la API.")

app = FastAPI(title="Conciliador Contable API", version=APP_VERSION)
api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)
request_history: dict[str, deque[float]] = defaultdict(deque)


async def require_api_key(x_api_key: str | None = Security(api_key_header)) -> None:
    if x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key invalida o faltante")


@app.middleware("http")
async def rate_limit_middleware(request: Request, call_next):
    if request.url.path in {"/health", "/version"}:
        return await call_next(request)

    client_ip = request.client.host if request.client else "unknown"
    now = time.time()
    bucket = request_history[client_ip]

    while bucket and now - bucket[0] > RATE_LIMIT_SECONDS:
        bucket.popleft()

    if len(bucket) >= RATE_LIMIT_REQUESTS:
        return JSONResponse(
            status_code=429,
            content={"detail": "Demasiadas solicitudes. Intente nuevamente en unos minutos."},
        )

    bucket.append(now)
    return await call_next(request)

default_allowed_origins = {
    "https://conciliador-contable.netlify.app",
    "https://conciliacion2026-web.onrender.com",
    "http://localhost:5173",
    "http://127.0.0.1:5173",
}

allowed_origins_raw = os.getenv("ALLOWED_ORIGINS", "")
configured_allowed_origins = {origin.strip() for origin in allowed_origins_raw.split(",") if origin.strip()}
allow_all_origins = "*" in configured_allowed_origins
allowed_origins = sorted(default_allowed_origins | configured_allowed_origins)

cors_kwargs = {
    "allow_methods": ["*"],
    "allow_headers": ["*"],
}

if allow_all_origins:
    # Con credenciales, '*' no es valido en CORS. Usamos regex global y desactivamos credenciales.
    cors_kwargs["allow_origin_regex"] = ".*"
    cors_kwargs["allow_credentials"] = False
else:
    cors_kwargs["allow_origins"] = allowed_origins
    cors_kwargs["allow_credentials"] = False

app.add_middleware(CORSMiddleware, **cors_kwargs)

MAX_UPLOAD_BYTES = 20 * 1024 * 1024


def _validate_upload_bytes(filename: str | None, content: bytes, label: str) -> None:
    if content is None or len(content) == 0:
        raise HTTPException(status_code=400, detail=f"El archivo {label} esta vacio")
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"El archivo {label} excede el tamaño maximo permitido ({MAX_UPLOAD_BYTES // (1024 * 1024)} MB)",
        )
    safe_name = (filename or "").strip()
    if not safe_name or ".." in safe_name or "/" in safe_name or "\\" in safe_name:
        raise HTTPException(status_code=400, detail=f"Nombre de archivo invalido para {label}")
    if not safe_name.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail=f"Solo se permiten archivos .xlsx para {label}")
    if not content.startswith(b"PK"):
        raise HTTPException(status_code=400, detail=f"El archivo {label} no tiene un formato Excel valido")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok", "version": APP_VERSION}


@app.get("/version")
def version() -> dict[str, str]:
    return {"version": APP_VERSION}


async def _procesar_archivos(
    file: UploadFile | None = File(None),
    pse_file: UploadFile | None = File(None),
    cruces_file: UploadFile | None = File(None),
    query_interno_file: UploadFile | None = File(None),
    adquirencias_file: UploadFile | None = File(None),
    tolerance_days: int = Form(1),
    tolerance_value: float = Form(0.01),
    temporal_tolerance_days: int = Form(0),
    allow_previous_month: bool = Form(False),
    allow_previous_year: bool = Form(False),
) -> dict:
    if file is None and pse_file is None and cruces_file is None:
        raise HTTPException(status_code=400, detail="Debes enviar al menos un archivo para procesar")

    if file is not None and (not file.filename or not file.filename.lower().endswith(".xlsx")):
        raise HTTPException(status_code=400, detail="Solo se permiten archivos .xlsx")
    if pse_file is not None and (not pse_file.filename or not pse_file.filename.lower().endswith(".xlsx")):
        raise HTTPException(status_code=400, detail="El archivo PSE debe ser .xlsx")
    if cruces_file is not None and (not cruces_file.filename or not cruces_file.filename.lower().endswith(".xlsx")):
        raise HTTPException(status_code=400, detail="El archivo de cruces contables debe ser .xlsx")
    if query_interno_file is not None and (not query_interno_file.filename or not query_interno_file.filename.lower().endswith(".xlsx")):
        raise HTTPException(status_code=400, detail="El archivo de Query Interno debe ser .xlsx")
    if adquirencias_file is not None and (not adquirencias_file.filename or not adquirencias_file.filename.lower().endswith(".xlsx")):
        raise HTTPException(status_code=400, detail="El archivo de Adquirencias debe ser .xlsx")

    try:
        contable_content = None
        pse_content = None
        cruces_content = None
        query_interno_content = None
        adquirencias_content = None

        if file is not None:
            contable_content = await file.read()
            _validate_upload_bytes(file.filename, contable_content, "contable")

        if pse_file is not None:
            pse_content = await pse_file.read()
            _validate_upload_bytes(pse_file.filename, pse_content, "PSE")

        if cruces_file is not None:
            cruces_content = await cruces_file.read()
            _validate_upload_bytes(cruces_file.filename, cruces_content, "de cruces contables")

        if query_interno_file is not None:
            query_interno_content = await query_interno_file.read()
            _validate_upload_bytes(query_interno_file.filename, query_interno_content, "de Query Interno")

        if adquirencias_file is not None:
            adquirencias_content = await adquirencias_file.read()
            _validate_upload_bytes(adquirencias_file.filename, adquirencias_content, "de Adquirencias")

        if pse_content is not None and cruces_content is None:
            raise HTTPException(status_code=400, detail="Debes enviar el archivo PSE junto con el archivo de cruces contables")
        if cruces_content is not None and pse_content is None and contable_content is None:
            raise HTTPException(
                status_code=400,
                detail="Debes enviar el archivo de cruces contables junto con PSE o Contable",
            )

        engine = ProcesadorIntegrado(
            contable_bytes=contable_content,
            pse_bytes=pse_content,
            cruces_bytes=cruces_content,
            query_interno_bytes=query_interno_content,
            adquirencias_bytes=adquirencias_content,
            date_tolerance_days=tolerance_days,
            value_tolerance=tolerance_value,
            temporal_config=ValidacionTemporalConfig(
                tolerancia_dias=temporal_tolerance_days,
                permitir_cruce_mes_anterior=allow_previous_month,
                permitir_cruce_ano_anterior=allow_previous_year,
            ),
        )
        return engine.procesar()
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(status_code=500, detail="Error procesando los archivos")


@app.post("/procesar")
async def procesar(
    _: None = Depends(require_api_key),
    file: UploadFile | None = File(None),
    pse_file: UploadFile | None = File(None),
    cruces_file: UploadFile | None = File(None),
    query_interno_file: UploadFile | None = File(None),
    adquirencias_file: UploadFile | None = File(None),
    tolerance_days: int = Form(1),
    tolerance_value: float = Form(0.01),
    temporal_tolerance_days: int = Form(0),
    allow_previous_month: bool = Form(False),
    allow_previous_year: bool = Form(False),
) -> dict:
    return await _procesar_archivos(
        file=file,
        pse_file=pse_file,
        cruces_file=cruces_file,
        query_interno_file=query_interno_file,
        adquirencias_file=adquirencias_file,
        tolerance_days=tolerance_days,
        tolerance_value=tolerance_value,
        temporal_tolerance_days=temporal_tolerance_days,
        allow_previous_month=allow_previous_month,
        allow_previous_year=allow_previous_year,
    )


@app.post("/pse/conciliar")
async def conciliar_pse(
    _: None = Depends(require_api_key),
    file: UploadFile | None = File(None),
    pse_file: UploadFile | None = File(None),
    cruces_file: UploadFile | None = File(None),
    query_interno_file: UploadFile | None = File(None),
    adquirencias_file: UploadFile | None = File(None),
    tolerance_days: int = Form(1),
    tolerance_value: float = Form(0.01),
    temporal_tolerance_days: int = Form(0),
    allow_previous_month: bool = Form(False),
    allow_previous_year: bool = Form(False),
) -> dict:
    return await _procesar_archivos(
        file=file,
        pse_file=pse_file,
        cruces_file=cruces_file,
        query_interno_file=query_interno_file,
        adquirencias_file=adquirencias_file,
        tolerance_days=tolerance_days,
        tolerance_value=tolerance_value,
        temporal_tolerance_days=temporal_tolerance_days,
        allow_previous_month=allow_previous_month,
        allow_previous_year=allow_previous_year,
    )
