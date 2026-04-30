from __future__ import annotations

import os

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

from conciliador import ConciliadorContable
from pse_conciliador import PseConciliador


app = FastAPI(title="Conciliador Contable API", version="1.0.0")

allowed_origins = [
    origin.strip()
    for origin in os.getenv("ALLOWED_ORIGINS", "*").split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/procesar")
async def procesar(file: UploadFile = File(...)) -> dict:
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Solo se permiten archivos .xlsx")

    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="El archivo esta vacio")

        engine = ConciliadorContable(content)
        resultado = engine.procesar()
        return resultado
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(status_code=500, detail="Error procesando el archivo")


@app.post("/pse/conciliar")
async def conciliar_pse(
    pse_file: UploadFile = File(...),
    cruces_file: UploadFile = File(...),
    tolerance_days: int = Form(1),
    tolerance_value: float = Form(0.01),
) -> dict:
    if not pse_file.filename or not pse_file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="El archivo PSE debe ser .xlsx")
    if not cruces_file.filename or not cruces_file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="El archivo de cruces contables debe ser .xlsx")

    try:
        pse_content = await pse_file.read()
        cruces_content = await cruces_file.read()
        if not pse_content:
            raise HTTPException(status_code=400, detail="El archivo PSE esta vacio")
        if not cruces_content:
            raise HTTPException(status_code=400, detail="El archivo de cruces contables esta vacio")

        engine = PseConciliador(
            pse_content,
            cruces_content,
            date_tolerance_days=tolerance_days,
            value_tolerance=tolerance_value,
        )
        return engine.procesar()
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(status_code=500, detail="Error conciliando el PSE")
