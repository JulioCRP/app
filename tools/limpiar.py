"""Limpia caches de Python y archivos temporales del proyecto.

Se invoca desde el Makefile y desde instalar.ps1 para no depender de comandos
de shell que difieren entre Windows y Linux/macOS.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
VENV = RAIZ / ".venv"

TEMPORALES = (
    "salida_prueba.xlsx",
    "reporte_con_promedios.xlsx",
    "reporte_sin_promedios.xlsx",
)


def main() -> int:
    print("Borrando caches de Python...")
    borrados = 0
    for carpeta in RAIZ.rglob("__pycache__"):
        if VENV in carpeta.parents or carpeta == VENV:
            continue
        try:
            shutil.rmtree(carpeta, ignore_errors=True)
            borrados += 1
        except OSError:
            pass
    print(f"  {borrados} carpeta(s) __pycache__ eliminada(s)")

    for nombre in TEMPORALES:
        ruta = RAIZ / nombre
        if ruta.exists():
            try:
                ruta.unlink()
                print(f"  Borrado {nombre}")
            except OSError:
                print(f"  No se pudo borrar {nombre}")

    print("Limpieza terminada.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
