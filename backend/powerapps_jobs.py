from __future__ import annotations

import base64
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


_jobs: dict[str, Job] = {}
_jobs_lock = threading.Lock()


def create_job(inputs: dict[str, bytes]) -> Job:
    job = Job(job_id=uuid.uuid4().hex, inputs=inputs)
    with _jobs_lock:
        _jobs[job.job_id] = job
    return job


def get_job(job_id: str) -> Job | None:
    with _jobs_lock:
        return _jobs.get(job_id)


def _set_job(job: Job, **values: Any) -> None:
    with _jobs_lock:
        for key, value in values.items():
            setattr(job, key, value)


def _decode_files(result: dict[str, Any]) -> list[tuple[str, bytes]]:
    output: list[tuple[str, bytes]] = []
    for item in result.get("files", []):
        name = item.get("name")
        content = item.get("file") or item.get("content")
        if name and content:
            output.append((str(name), base64.b64decode(content)))
    if not output and result.get("file"):
        output.append((str(result.get("output_name", "resultado.xlsx")), base64.b64decode(result["file"])))
    return output


def _process(job: Job) -> dict[str, Any]:
    inputs = job.inputs
    module = inputs.get("__module__", b"").decode("utf-8") or "integrated"
    temporal = ValidacionTemporalConfig()
    if module == "contable":
        result = ConciliadorContable(inputs["file"], temporal_config=temporal).procesar()
    elif module == "pse":
        result = PseConciliador(inputs["pse_file"], inputs["cruces_file"], temporal_config=temporal).procesar()
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
            temporal_config=temporal,
        ).procesar()
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
        _set_job(job, status="failed", error=str(error), inputs={})
