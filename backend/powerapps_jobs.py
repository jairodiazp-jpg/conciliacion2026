from __future__ import annotations

import base64
import logging
import os
import threading
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any

from adquirencias_conciliador import AdquirenciasConciliador
from conciliador import ConciliadorContable
from integrador import ProcesadorIntegrado
from query_memorando_conciliador import conciliar_query_memorando
from pse_conciliador import PseConciliador
from validacion_temporal import ValidacionTemporalConfig


@dataclass
class Job:
    job_id: str
    status: str = "pending"
    created_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    error: str | None = None
    files: list[dict[str, str]] = field(default_factory=list)
    outputs: dict[str, tuple[str, bytes]] = field(default_factory=dict, repr=False)
    inputs: dict[str, bytes] = field(default_factory=dict, repr=False)
    options: dict[str, int | float | bool] = field(default_factory=dict, repr=False)


_jobs: dict[str, Job] = {}
_jobs_lock = threading.Lock()
_logger = logging.getLogger(__name__)
_job_ttl_seconds = max(300, int(os.getenv("POWERAPPS_JOB_TTL_SECONDS", "3600")))


def create_job(
    inputs: dict[str, bytes],
    options: dict[str, int | float | bool] | None = None,
) -> Job:
    _cleanup_expired_jobs()
    job = Job(job_id=uuid.uuid4().hex, inputs=inputs, options=options or {})
    with _jobs_lock:
        _jobs[job.job_id] = job
    return job


def get_job(job_id: str) -> Job | None:
    _cleanup_expired_jobs()
    with _jobs_lock:
        return _jobs.get(job_id)


def _cleanup_expired_jobs() -> None:
    now = datetime.now(timezone.utc)
    with _jobs_lock:
        expired = [
            job_id
            for job_id, job in _jobs.items()
            if (now - datetime.fromisoformat(job.created_at)).total_seconds() > _job_ttl_seconds
        ]
        for job_id in expired:
            del _jobs[job_id]


def _set_job(job: Job, **values: Any) -> None:
    with _jobs_lock:
        for key, value in values.items():
            setattr(job, key, value)


def _decode_files(result: dict[str, Any]) -> list[tuple[str, bytes]]:
    output: list[tuple[str, bytes]] = []
    names: set[str] = set()

    def add_file(name: Any, content: Any) -> None:
        if not name or not content or str(name) in names:
            return
        decoded = base64.b64decode(content, validate=True)
        output.append((str(name), decoded))
        names.add(str(name))

    for item in result.get("files", []):
        name = item.get("name")
        content = item.get("file") or item.get("content")
        add_file(name, content)

    if result.get("file"):
        add_file(result.get("output_name", "resultado.xlsx"), result["file"])

    if result.get("secondary_file"):
        add_file(
            result.get("secondary_output_name", "resultado_secundario.xlsx"),
            result["secondary_file"],
        )

    # Query/Memorando returns two named workbooks instead of a `files` list.
    for key in ("query_file", "memorando_file"):
        item = result.get(key)
        if isinstance(item, dict):
            add_file(item.get("name"), item.get("file") or item.get("content"))

    return output


def _process(job: Job) -> dict[str, Any]:
    inputs = job.inputs
    options = job.options
    module = inputs.get("__module__", b"").decode("utf-8") or "integrated"
    temporal = ValidacionTemporalConfig(
        tolerancia_dias=int(options.get("temporal_tolerance_days", 0)),
        permitir_cruce_mes_anterior=bool(options.get("allow_previous_month", False)),
        permitir_cruce_ano_anterior=bool(options.get("allow_previous_year", False)),
    )
    if module == "contable":
        result = ConciliadorContable(inputs["file"], temporal_config=temporal).procesar()
    elif module == "pse":
        result = PseConciliador(
            inputs["pse_file"],
            inputs["cruces_file"],
            date_tolerance_days=int(options.get("tolerance_days", 1)),
            value_tolerance=float(options.get("tolerance_value", 0.01)),
            temporal_config=temporal,
        ).procesar()
    elif module == "query":
        result = conciliar_query_memorando(inputs["query_file"], inputs["memorando_file"])
    elif module == "adquirencias":
        result = AdquirenciasConciliador(inputs["adquirencias_file"], inputs["file"]).procesar()
        result = {
            "files": [
                {"name": "ADQUIRENCIAS_CONCILIADO.xlsx", "file": result["adquirencias_file"]},
                {"name": "CONCILIACION_CONTABLE.xlsx", "file": result["ccs_file"]},
            ]
        }
    else:
        result = ProcesadorIntegrado(
            contable_bytes=inputs.get("file"),
            pse_bytes=inputs.get("pse_file"),
            cruces_bytes=inputs.get("cruces_file"),
            query_interno_bytes=inputs.get("query_interno_file"),
            adquirencias_bytes=inputs.get("adquirencias_file"),
            date_tolerance_days=int(options.get("tolerance_days", 1)),
            value_tolerance=float(options.get("tolerance_value", 0.01)),
            temporal_config=temporal,
        ).procesar()
        if "query_file" in inputs and "memorando_file" in inputs:
            query_result = conciliar_query_memorando(
                inputs["query_file"],
                inputs["memorando_file"],
            )
            result.setdefault("files", []).extend(
                [
                    query_result["query_file"],
                    query_result["memorando_file"],
                ]
            )
    return result


def run_job(job: Job) -> None:
    _set_job(job, status="processing")
    try:
        result = _process(job)
        files: list[dict[str, str]] = []
        outputs: dict[str, tuple[str, bytes]] = {}
        for name, content in _decode_files(result):
            file_id = uuid.uuid4().hex
            outputs[file_id] = (name, content)
            files.append({"file_id": file_id, "name": name})
        if not files:
            raise RuntimeError("El procesamiento no generó archivos Excel descargables.")
        _set_job(job, status="completed", files=files, outputs=outputs, inputs={})
    except Exception as error:
        _logger.exception("Power Apps job failed: %s", job.job_id)
        _set_job(job, status="failed", error=str(error), inputs={})
