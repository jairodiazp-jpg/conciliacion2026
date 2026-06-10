from __future__ import annotations

import os

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

from conciliador import ConciliadorContable
from integrador import ProcesadorIntegrado
from version import get_app_version
from validacion_temporal import ValidacionTemporalConfig


APP_VERSION = get_app_version()

app = FastAPI(title="Conciliador Contable API", version=APP_VERSION)

default_allowed_origins = {
    "https://conciliador-contable.netlify.app",
    "https://conciliacion2026-web.onrender.com",
    "https://jairodiazp-jpg.github.io",
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
    if adquirencias_file is not None and (not adquirencias_file.filename or not adquirencias_file.filename.lower().endswith(".xlsx")):
        raise HTTPException(status_code=400, detail="El archivo de Adquirencias debe ser .xlsx")

    try:
        contable_content = None
        pse_content = None
        cruces_content = None
        adquirencias_content = None

        if file is not None:
            contable_content = await file.read()
            if not contable_content:
                raise HTTPException(status_code=400, detail="El archivo esta vacio")

        if pse_file is not None:
            pse_content = await pse_file.read()
            if not pse_content:
                raise HTTPException(status_code=400, detail="El archivo PSE esta vacio")

        if cruces_file is not None:
            cruces_content = await cruces_file.read()
            if not cruces_content:
                raise HTTPException(status_code=400, detail="El archivo de cruces contables esta vacio")

        if adquirencias_file is not None:
            adquirencias_content = await adquirencias_file.read()
            if not adquirencias_content:
                raise HTTPException(status_code=400, detail="El archivo de Adquirencias esta vacio")

        if pse_content is not None and cruces_content is None:
            raise HTTPException(status_code=400, detail="Debes enviar el archivo PSE junto con el archivo de cruces contables")
        if cruces_content is not None and pse_content is None and contable_content is None and adquirencias_content is None:
            raise HTTPException(
                status_code=400,
                detail="Debes enviar el archivo de cruces contables junto con PSE, Contable o Adquirencias",
            )

        engine = ProcesadorIntegrado(
            contable_bytes=contable_content,
            pse_bytes=pse_content,
            cruces_bytes=cruces_content,
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
    file: UploadFile | None = File(None),
    pse_file: UploadFile | None = File(None),
    cruces_file: UploadFile | None = File(None),
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
        adquirencias_file=adquirencias_file,
        tolerance_days=tolerance_days,
        tolerance_value=tolerance_value,
        temporal_tolerance_days=temporal_tolerance_days,
        allow_previous_month=allow_previous_month,
        allow_previous_year=allow_previous_year,
    )


@app.post("/pse/conciliar")
async def conciliar_pse(
    file: UploadFile | None = File(None),
    pse_file: UploadFile | None = File(None),
    cruces_file: UploadFile | None = File(None),
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
        adquirencias_file=adquirencias_file,
        tolerance_days=tolerance_days,
        tolerance_value=tolerance_value,
        temporal_tolerance_days=temporal_tolerance_days,
        allow_previous_month=allow_previous_month,
        allow_previous_year=allow_previous_year,
    )
