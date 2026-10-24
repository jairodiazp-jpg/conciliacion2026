from __future__ import annotations

import os
import logging

from fastapi import BackgroundTasks, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware

from conciliador import ConciliadorContable
from integrador import ProcesadorIntegrado
from version import get_app_version
from validacion_temporal import ValidacionTemporalConfig
from query_memorando_conciliador import conciliar_query_memorando
from voucher_busqueda import search_voucher_documents
from powerapps_jobs import create_job, get_job, run_job


APP_VERSION = get_app_version()
logger = logging.getLogger(__name__)
MAX_UPLOAD_BYTES = max(1, int(os.getenv("MAX_UPLOAD_BYTES", str(50 * 1024 * 1024))))
MAX_POWERAPPS_TOTAL_BYTES = max(
    MAX_UPLOAD_BYTES,
    int(os.getenv("MAX_POWERAPPS_TOTAL_BYTES", str(250 * 1024 * 1024))),
)

app = FastAPI(title="Conciliador Contable API", version=APP_VERSION)

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
            if len(contable_content) > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail="El archivo contable supera el límite permitido.")
            if not contable_content:
                raise HTTPException(status_code=400, detail="El archivo esta vacio")

        if pse_file is not None:
            pse_content = await pse_file.read()
            if len(pse_content) > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail="El archivo PSE supera el límite permitido.")
            if not pse_content:
                raise HTTPException(status_code=400, detail="El archivo PSE esta vacio")

        if cruces_file is not None:
            cruces_content = await cruces_file.read()
            if len(cruces_content) > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail="El archivo de cruces supera el límite permitido.")
            if not cruces_content:
                raise HTTPException(status_code=400, detail="El archivo de cruces contables esta vacio")

        if query_interno_file is not None:
            query_interno_content = await query_interno_file.read()
            if len(query_interno_content) > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail="El archivo Query Interno supera el límite permitido.")
            if not query_interno_content:
                raise HTTPException(status_code=400, detail="El archivo de Query Interno esta vacio")

        if adquirencias_file is not None:
            adquirencias_content = await adquirencias_file.read()
            if len(adquirencias_content) > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail="El archivo de Adquirencias supera el límite permitido.")
            if not adquirencias_content:
                raise HTTPException(status_code=400, detail="El archivo de Adquirencias esta vacio")

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
        logger.exception("Error procesando archivos en /procesar")
        raise HTTPException(status_code=500, detail="Error procesando los archivos")


@app.post("/procesar")
async def procesar(
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


@app.post("/query-memorando/conciliar")
async def conciliar_query_memo(
    query_file: UploadFile | None = File(None),
    memorando_file: UploadFile | None = File(None),
) -> dict:
    if query_file is None or memorando_file is None:
        raise HTTPException(status_code=400, detail="Debes enviar el Query y el memorando ya conciliado")
    for uploaded, label in ((query_file, "Query"), (memorando_file, "memorando")):
        if not uploaded.filename or not uploaded.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail=f"El archivo {label} debe ser .xlsx")
    query_content = await query_file.read()
    memorando_content = await memorando_file.read()
    if not query_content or not memorando_content:
        raise HTTPException(status_code=400, detail="Los archivos Query y memorando no pueden estar vacíos")
    try:
        return conciliar_query_memorando(query_content, memorando_content)
    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=422, detail=f"No se pudieron validar los archivos: {error}") from error


@app.post("/voucher/buscar")
async def buscar_valor_voucher(
    memorando_file: UploadFile | None = File(None),
    voucher_files: list[UploadFile] = File(default=[]),
    value: str = Form(""),
    exact_date: str | None = Form(None),
    store_code: str = Form(""),
    store: str | None = Form(None),
    reference: str | None = Form(None),
    account: str | None = Form(None),
) -> dict:
    if memorando_file is None:
        raise HTTPException(status_code=400, detail="Selecciona el memorando conciliado.")
    if not value.strip():
        raise HTTPException(status_code=400, detail="El valor exacto es obligatorio.")
    if not store_code.strip():
        raise HTTPException(status_code=400, detail="El código del punto de venta es obligatorio.")
    if not memorando_file.filename or not memorando_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="El memorando debe ser un archivo Excel.")
    try:
        memo_content = await memorando_file.read()
        vouchers = [(file.filename or "soporte", await file.read()) for file in voucher_files]
        return search_voucher_documents(
            memo_content,
            memorando_file.filename,
            vouchers,
            value,
            exact_date,
            store,
            reference,
            account,
            store_code=store_code,
        )
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error
    except Exception as error:
        raise HTTPException(status_code=422, detail=f"No se pudo consultar el documento: {error}") from error


def _validate_powerapps_key(value: str | None) -> None:
    expected = os.getenv("POWERAPPS_API_KEY", "").strip()
    if not expected:
        raise HTTPException(status_code=503, detail="La API de Power Apps no está configurada.")
    if value != expected:
        raise HTTPException(status_code=401, detail="Credencial de Power Apps inválida.")


@app.post("/powerapps/jobs", status_code=202)
async def create_powerapps_job(
    background_tasks: BackgroundTasks,
    module: str = Form("integrated"),
    file: UploadFile | None = File(None),
    pse_file: UploadFile | None = File(None),
    cruces_file: UploadFile | None = File(None),
    query_interno_file: UploadFile | None = File(None),
    adquirencias_file: UploadFile | None = File(None),
    query_file: UploadFile | None = File(None),
    memorando_file: UploadFile | None = File(None),
    tolerance_days: int = Form(1),
    tolerance_value: float = Form(0.01),
    temporal_tolerance_days: int = Form(0),
    allow_previous_month: bool = Form(False),
    allow_previous_year: bool = Form(False),
    x_powerapps_key: str | None = Header(None),
) -> dict:
    _validate_powerapps_key(x_powerapps_key)
    total_bytes = 0
    allowed_modules = {"integrated", "contable", "pse", "query", "adquirencias"}
    if module not in allowed_modules:
        raise HTTPException(status_code=400, detail="Tipo de cruce no permitido.")

    fields = {
        "file": file,
        "pse_file": pse_file,
        "cruces_file": cruces_file,
        "query_interno_file": query_interno_file,
        "adquirencias_file": adquirencias_file,
        "query_file": query_file,
        "memorando_file": memorando_file,
    }
    inputs: dict[str, bytes] = {"__module__": module.encode("utf-8")}
    for name, upload in fields.items():
        if upload is not None:
            content = await upload.read()
            if not content:
                raise HTTPException(status_code=400, detail=f"El archivo {name} está vacío.")
            if len(content) > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail=f"El archivo {name} supera el límite permitido.")
            total_bytes += len(content)
            if total_bytes > MAX_POWERAPPS_TOTAL_BYTES:
                raise HTTPException(status_code=413, detail="La solicitud supera el límite total permitido.")
            inputs[name] = content

    if module == "query":
        # Accept query_interno_file as a backwards-compatible alias for query_file.
        if "query_file" not in inputs and "query_interno_file" in inputs:
            inputs["query_file"] = inputs["query_interno_file"]
        if "query_file" not in inputs or "memorando_file" not in inputs:
            raise HTTPException(
                status_code=400,
                detail="El cruce Query requiere query_file y memorando_file.",
            )
    elif module == "pse" and ("pse_file" not in inputs or "cruces_file" not in inputs):
        raise HTTPException(status_code=400, detail="El cruce PSE requiere pse_file y cruces_file.")
    elif module == "adquirencias" and ("adquirencias_file" not in inputs or "file" not in inputs):
        raise HTTPException(
            status_code=400,
            detail="El cruce de adquirencias requiere adquirencias_file y file.",
        )
    elif module == "contable" and "file" not in inputs:
        raise HTTPException(status_code=400, detail="El cruce contable requiere file.")
    if module == "integrated" and "file" not in inputs and "pse_file" not in inputs and "cruces_file" not in inputs:
        raise HTTPException(status_code=400, detail="Debes enviar al menos un archivo para el cruce integrado.")
    job = create_job(
        inputs,
        options={
            "tolerance_days": tolerance_days,
            "tolerance_value": tolerance_value,
            "temporal_tolerance_days": temporal_tolerance_days,
            "allow_previous_month": allow_previous_month,
            "allow_previous_year": allow_previous_year,
        },
    )
    background_tasks.add_task(run_job, job)
    return {"job_id": job.job_id, "status": job.status, "created_at": job.created_at}


@app.get("/powerapps/jobs/{job_id}")
def powerapps_job_status(job_id: str, x_powerapps_key: str | None = Header(None)) -> dict:
    _validate_powerapps_key(x_powerapps_key)
    job = get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Trabajo no encontrado o expirado.")
    response: dict[str, object] = {
        "job_id": job.job_id,
        "status": job.status,
        "created_at": job.created_at,
    }
    if job.error:
        response["error"] = job.error
    if job.status == "completed":
        response["files"] = [
            {
                **file,
                "download_path": f"/powerapps/jobs/{job.job_id}/files/{file['file_id']}",
            }
            for file in job.files
        ]
    return response


@app.get("/powerapps/jobs/{job_id}/files/{file_id}")
def download_powerapps_file(
    job_id: str,
    file_id: str,
    x_powerapps_key: str | None = Header(None),
) -> Response:
    _validate_powerapps_key(x_powerapps_key)
    job = get_job(job_id)
    if job is None or job.status != "completed":
        raise HTTPException(status_code=404, detail="Trabajo no encontrado o aún no completado.")
    output = job.outputs.get(file_id)
    if output is None:
        raise HTTPException(status_code=404, detail="Archivo no encontrado o expirado.")
    name, content = output
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{name}"',
            "Cache-Control": "no-store",
        },
    )
