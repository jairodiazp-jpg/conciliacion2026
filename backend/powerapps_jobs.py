from __future__ import annotations

import base64
import json
import os
import threading
import urllib.error
import urllib.parse
import urllib.request
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


class JobConfigurationError(RuntimeError):
    pass


class OneDriveClient:
    def __init__(self) -> None:
        self.tenant_id = os.getenv("MICROSOFT_TENANT_ID", "").strip()
        self.client_id = os.getenv("MICROSOFT_CLIENT_ID", "").strip()
        self.client_secret = os.getenv("MICROSOFT_CLIENT_SECRET", "").strip()
        self.drive_id = os.getenv("ONEDRIVE_DRIVE_ID", "").strip()
        self.root_folder_id = os.getenv("ONEDRIVE_RESULTS_FOLDER_ID", "").strip()
        if not all((self.tenant_id, self.client_id, self.client_secret, self.drive_id, self.root_folder_id)):
            raise JobConfigurationError(
                "OneDrive no está configurado. Define las variables de Microsoft Graph y la carpeta de resultados."
            )

    def _request(self, request: urllib.request.Request) -> dict[str, Any]:
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            detail = error.read().decode("utf-8", errors="replace")[:500]
            raise RuntimeError(f"Microsoft Graph rechazó la operación ({error.code}): {detail}") from error

    def _token(self) -> str:
        values = urllib.parse.urlencode(
            {
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            }
        ).encode()
        request = urllib.request.Request(
            f"https://login.microsoftonline.com/{urllib.parse.quote(self.tenant_id)}/oauth2/v2.0/token",
            data=values,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        token = self._request(request).get("access_token")
        if not token:
            raise RuntimeError("Microsoft Graph no devolvió un token de acceso.")
        return str(token)

    def upload(self, job_id: str, name: str, content: bytes) -> dict[str, str]:
        token = self._token()
        folder_request = urllib.request.Request(
            f"https://graph.microsoft.com/v1.0/drives/{urllib.parse.quote(self.drive_id, safe='')}/items/{urllib.parse.quote(self.root_folder_id, safe='')}/children",
            data=json.dumps({"name": job_id, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}).encode(),
            method="POST",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        )
        folder = self._request(folder_request)
        folder_id = str(folder["id"])
        safe_name = urllib.parse.quote(name, safe="")
        request = urllib.request.Request(
            f"https://graph.microsoft.com/v1.0/drives/{urllib.parse.quote(self.drive_id, safe='')}/items/{urllib.parse.quote(folder_id, safe='')}:/"
            f"{safe_name}:/content",
            data=content,
            method="PUT",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
        )
        response = self._request(request)
        return {
            "file_id": str(response["id"]),
            "name": name,
            "web_url": str(response.get("webUrl", "")),
        }


@dataclass
class Job:
    job_id: str
    status: str = "pending"
    created_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    error: str | None = None
    files: list[dict[str, str]] = field(default_factory=list)
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
    module = inputs.pop("__module__", b"").decode("utf-8") or "integrated"
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
        client = OneDriveClient()
        result = _process(job)
        uploaded = []
        for name, content in _decode_files(result):
            uploaded.append(client.upload(job.job_id, name, content))
        if not uploaded:
            raise RuntimeError("El procesamiento no generó archivos Excel descargables.")
        _set_job(job, status="completed", files=uploaded, inputs={})
    except Exception as error:
        _set_job(job, status="failed", error=str(error), inputs={})
