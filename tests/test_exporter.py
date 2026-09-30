"""Prueba del exportador.

Verifica que el .xlsx generado conserve A-N intactos y que O/P/Q coincidan con
el archivo de referencia.
"""

from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import numpy as np  # noqa: E402
import openpyxl  # noqa: E402

from src import exporter, metrics  # noqa: E402
from src.parser import parsear_reporte  # noqa: E402

ORIGINAL = RAIZ / "Movimientos_con_promedios.xlsx"


def test_exporta_y_conserva_a_n() -> None:
    contenido = ORIGINAL.read_bytes()
    resultado = parsear_reporte(ORIGINAL)
    resumen = metrics.resumen_productos(resultado.movimientos, resultado.productos)

    salida = exporter.exportar_con_promedios(contenido, resultado, resumen)
    destino = RAIZ / "salida_prueba.xlsx"
    destino.write_bytes(salida)
    print(f"    generado: {len(salida):,} bytes (original {len(contenido):,})")

    original = openpyxl.load_workbook(ORIGINAL, data_only=True).worksheets[0]
    nuevo = openpyxl.load_workbook(destino, data_only=True).worksheets[0]

    assert original.max_row == nuevo.max_row, "no debe cambiar el numero de filas"

    # A-N intactas.
    diferencias = 0
    primera: str | None = None
    for fila in range(1, original.max_row + 1):
        for columna in range(1, 15):
            a = original.cell(fila, columna).value
            b = nuevo.cell(fila, columna).value
            if a != b:
                # Tolerar 0.0 vs None en celdas vacias del original.
                if (a is None and b == 0) or (b is None and a == 0):
                    continue
                diferencias += 1
                if primera is None:
                    primera = f"fila {fila} col {columna}: {a!r} -> {b!r}"
    assert diferencias == 0, f"A-N cambiaron en {diferencias} celdas. Primera: {primera}"
    print("ok  A-N intactas")

    # O/P/Q coinciden con la referencia.
    fallos = 0
    validas = 0
    fila_encabezado = resultado.fila_encabezado
    for fila in range(1, original.max_row + 1):
        if fila == fila_encabezado:
            # El encabezado no es un grupo: se valida aparte mas abajo.
            continue
        es_grupo = str(original.cell(fila, 1).value or "").lower().startswith("código producto")
        for columna in (15, 16, 17):
            a = original.cell(fila, columna).value
            b = nuevo.cell(fila, columna).value
            if es_grupo:
                if b is not None:
                    fallos += 1
                continue
            if a is None and b is None:
                continue
            if a is None or b is None:
                fallos += 1
                continue
            if abs(float(a) - float(b)) > 0.005:
                fallos += 1
            else:
                validas += 1
    print(f"    celdas O/P/Q verificadas: {validas:,}, discrepancias: {fallos}")
    assert fallos == 0, f"{fallos} celdas O/P/Q no coinciden con la referencia"
    print("ok  O/P/Q coinciden con la referencia")

    # Estilos y encabezado.
    estilos = openpyxl.load_workbook(destino).worksheets[0]
    assert estilos.cell(2, 15).value == "Promedio mensual", "falta el encabezado de O"
    assert estilos.cell(2, 17).value == "Promedio anual", "falta el encabezado de Q"
    assert estilos.cell(2, 15).fill.start_color.rgb == estilos.cell(2, 14).fill.start_color.rgb, (
        "O nohereda el relleno del encabezado"
    )
    print("ok  encabezados y formato")

    destino.unlink()
    print("ok  test_exporta_y_conserva_a_n")


def test_error_si_falta_plantilla() -> None:
    """Un archivo sin filas de detalle debe dar un error claro, no un crash."""
    import io

    import pandas as pd

    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja["A1"] = "Codigo producto"
    hoja["A2"] = "Codigo producto: X Producto"
    buffer = io.BytesIO()
    libro.save(buffer)

    class FalsoResultado:
        fila_encabezado = 1
        movimientos = pd.DataFrame({"codigo": [], "fila_excel": []})

    resumen = pd.DataFrame(
        {
            "codigo": [],
            "prom_mensual": [],
            "prom_semestral": [],
            "prom_anual": [],
        }
    )

    try:
        exporter.exportar_con_promedios(buffer.getvalue(), FalsoResultado(), resumen)
    except exporter.ErrorExportacion as e:
        assert "plantilla" in str(e)
        print("ok  test_error_si_falta_plantilla")
        return
    raise AssertionError("se esperaba ErrorExportacion")


def main() -> int:
    print("Ejecutando pruebas del exportador...")
    test_exporta_y_conserva_a_n()
    test_error_si_falta_plantilla()
    print("\nTodas las pruebas del exportador pasaron.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
