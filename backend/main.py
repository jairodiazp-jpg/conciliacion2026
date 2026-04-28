from __future__ import annotations

import os

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

from conciliador import ConciliadorContable


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
