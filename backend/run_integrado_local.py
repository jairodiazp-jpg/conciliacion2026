from __future__ import annotations

import argparse
import base64
from pathlib import Path

from integrador import ProcesadorIntegrado
from validacion_temporal import ValidacionTemporalConfig


def _read_bytes(path: Path) -> bytes:
    data = path.read_bytes()
    if not data:
        raise ValueError(f"Archivo vacío: {path}")
    return data


def _write_b64_file(out_dir: Path, name: str, b64: str) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / name
    out_path.write_bytes(base64.b64decode(b64))
    return out_path


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Ejecuta el modo integrado (Contable + PSE + Cruces) y guarda los XLSX resultantes en disco."
        )
    )
    parser.add_argument("--contable", required=True, help="Ruta al XLSX contable (Bancolombia/contable).")
    parser.add_argument("--pse", required=True, help="Ruta al XLSX PSE.")
    parser.add_argument("--cruces", required=True, help="Ruta al XLSX de cruces contables.")
    parser.add_argument("--out-dir", default="out", help="Carpeta de salida (default: ./out)")
    parser.add_argument("--tolerance-days", type=int, default=1)
    parser.add_argument("--tolerance-value", type=float, default=0.01)
    parser.add_argument("--temporal-tolerance-days", type=int, default=0)
    parser.add_argument("--allow-previous-month", action="store_true")
    parser.add_argument("--allow-previous-year", action="store_true")
    args = parser.parse_args()

    contable_path = Path(args.contable).expanduser().resolve()
    pse_path = Path(args.pse).expanduser().resolve()
    cruces_path = Path(args.cruces).expanduser().resolve()
    out_dir = Path(args.out_dir).expanduser().resolve()

    engine = ProcesadorIntegrado(
        contable_bytes=_read_bytes(contable_path),
        pse_bytes=_read_bytes(pse_path),
        cruces_bytes=_read_bytes(cruces_path),
        date_tolerance_days=args.tolerance_days,
        value_tolerance=args.tolerance_value,
        temporal_config=ValidacionTemporalConfig(
            tolerancia_dias=args.temporal_tolerance_days,
            permitir_cruce_mes_anterior=args.allow_previous_month,
            permitir_cruce_ano_anterior=args.allow_previous_year,
        ),
    )

    result = engine.procesar()
    for f in result.get("files", []):
        name = f.get("name") or "output.xlsx"
        content_b64 = f.get("content")
        if not content_b64:
            continue
        out_path = _write_b64_file(out_dir, name, content_b64)
        print(f"OK: {out_path}")

    alertas = result.get("alertas") or []
    if alertas:
        print("\nAlertas:")
        for a in alertas:
            print(f"- {a}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
