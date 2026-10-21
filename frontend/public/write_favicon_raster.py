#!/usr/bin/env python3
"""Genera PNG e ICO a partir de los colores del favicon SVG existente, sin dependencias extra."""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

NAVY = (0x1B, 0x2A, 0x4A, 255)
WHITE = (255, 255, 255, 255)
BLUE = (0x7E, 0xB6, 0xFF, 255)
SIZE = 32


def _pixel(x: int, y: int) -> tuple[int, int, int, int]:
    # Aproximación legible a 32px de las dos ondas del SVG existente.
    if 10 <= y <= 16:
        wave = int(16 + 8 * __import__("math").sin((x - 4) / 5.2))
        if abs(y - wave) <= 1:
            return WHITE
    if 18 <= y <= 24:
        wave = int(22 + 7 * __import__("math").sin((x - 4) / 5.2))
        if abs(y - wave) <= 1:
            return BLUE
    return NAVY


def _png_chunk(tag: bytes, data: bytes) -> bytes:
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)


def write_png(path: Path) -> None:
    raw = b""
    for y in range(SIZE):
        raw += b"\x00"
        for x in range(SIZE):
            raw += bytes(_pixel(x, y))
    ihdr = struct.pack(">IIBBBBB", SIZE, SIZE, 8, 6, 0, 0, 0)
    png = b"\x89PNG\r\n\x1a\n" + _png_chunk(b"IHDR", ihdr) + _png_chunk(b"IDAT", zlib.compress(raw, 9)) + _png_chunk(b"IEND", b"")
    path.write_bytes(png)


def write_ico(png_bytes: bytes, path: Path) -> None:
    # ICO con PNG embebido (soportado por navegadores modernos).
    header = struct.pack("<HHH", 0, 1, 1)
    entry = struct.pack("<BBBBHHII", SIZE, SIZE, 0, 0, 1, 32, len(png_bytes), 22)
    path.write_bytes(header + entry + png_bytes)


def main() -> None:
    folder = Path(__file__).resolve().parent
    png_path = folder / "calypso-favicon.png"
    ico_path = folder / "favicon.ico"
    write_png(png_path)
    write_ico(png_path.read_bytes(), ico_path)
    print(f"Escrito {png_path.name} y {ico_path.name}")


if __name__ == "__main__":
    main()
