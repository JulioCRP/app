"""Prueba del caso real del usuario: el archivo sube solo con columnas A-N.

Se genera una copia sin O, P y Q, se procesa con el pipeline completo y se
comprueba que los promedios calculados coinciden con los del reporte original.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import numpy as np  # noqa: E402
import openpyxl  # noqa: E402

from src import exporter, metrics  # noqa: E402
from src.parser import parsear_reporte  # noqa: E402

ORIGINAL = RAIZ / "Movimientos_con_promedios.xlsx"


def crear_solo_a_n(origen: Path, destino: Path) -> None:
    """Copia el archivo borrando las columnas O, P y Q."""
    shutil.copyfile(origen, destino)
    libro = openpyxl.load_workbook(destino)
    hoja = libro.worksheets[0]
    for fila in range(1, hoja.max_row + 1):
        for columna in (15, 16, 17):
            hoja.cell(fila, columna).value = None
    libro.save(destino)
    libro.close()


def main() -> int:
    print("Ejecutando prueba del archivo sin promedios (solo A-N)...")
    temporal = Path(tempfile.mkdtemp(prefix="prueba_a_n_"))
    entrada = temporal / "reporte_sin_promedios.xlsx"
    salida = temporal / "reporte_con_promedios.xlsx"

    try:
        crear_solo_a_n(ORIGINAL, entrada)
        print(f"    archivo de entrada: {entrada.name} ({entrada.stat().st_size:,} bytes)")

        contenido = entrada.read_bytes()

        # 1. El parser debe leer un archivo de 14 columnas sin romperse.
        resultado = parsear_reporte(contenido)
        movimientos = resultado.movimientos
        productos = resultado.productos
        print(
            f"    leido: {len(productos):,} productos, {len(movimientos):,} movimientos, "
            f"encabezado en fila {resultado.fila_encabezado}"
        )
        assert len(productos) == 2883, f"productos inesperados: {len(productos)}"
        assert len(movimientos) == 37813, f"movimientos inesperados: {len(movimientos)}"
        assert len(resultado.encabezado) == 14, f"encabezado inesperado: {resultado.encabezado}"
        for aviso in resultado.avisos:
            print(f"    aviso: {aviso}")

        # 2. Sin columnas O/P/Q, la verificacion debe decir "no disponible".
        resumen = metrics.resumen_productos(movimientos, productos)
        verificacion = metrics.verificar_promedios(movimientos, resumen)
        assert verificacion["disponible"] is False, (
            "un archivo sin promedios no debe reportarse como verificable"
        )
        print("    verificacion: no disponible (correcto, el archivo no trae O/P/Q)")

        # 3. Los promedios calculados deben ser los mismos que en el reporte
        #    original, que si los trae.
        referencia = parsear_reporte(ORIGINAL)
        resumen_original = metrics.resumen_productos(
            referencia.movimientos, referencia.productos
        )
        verificacion_real = metrics.verificar_promedios(
            referencia.movimientos, resumen_original
        )
        assert verificacion_real["disponible"]
        assert int(verificacion_real["resumen"]["difieren"].sum()) == 0, (
            "los promedios calculados desde el archivo A-N no coinciden con la referencia"
        )
        print(
            f"    promedios correctos: {verificacion_real['coincidencia']:,} de "
            f"{verificacion_real['total']:,}"
        )

        # 4. Exportar debe escribir O/P/Q y no alterar A-N.
        exportado = exporter.exportar_con_promedios(contenido, resultado, resumen)
        salida.write_bytes(exportado)
        print(f"    archivo exportado: {salida.stat().st_size:,} bytes")

        original = openpyxl.load_workbook(ORIGINAL, data_only=True).worksheets[0]
        nuevo = openpyxl.load_workbook(salida, data_only=True).worksheets[0]
        assert nuevo.max_row == original.max_row

        # A-N identicas al archivo de entrada (sin promedios).
        entrada_hoja = openpyxl.load_workbook(entrada, data_only=True).worksheets[0]
        for fila in range(1, entrada_hoja.max_row + 1):
            for columna in range(1, 15):
                assert entrada_hoja.cell(fila, columna).value == nuevo.cell(fila, columna).value, (
                    f"A-N cambio en fila {fila} columna {columna}"
                )
        print("    A-N identicas a la entrada")

        # O/P/Q coinciden con el reporte de referencia.
        verificadas = 0
        for fila in range(1, original.max_row + 1):
            if fila == resultado.fila_encabezado:
                continue
            es_grupo = str(original.cell(fila, 1).value or "").lower().startswith("código producto")
            for columna in (15, 16, 17):
                a = original.cell(fila, columna).value
                b = nuevo.cell(fila, columna).value
                if es_grupo:
                    assert b is None, f"la fila de grupo {fila} no deberia tener promedio"
                    continue
                if a is None:
                    assert b is None, f"O/P/Q deberia quedar vacia en la fila {fila}"
                    continue
                assert b is not None, f"falta el promedio en la fila {fila}, columna {columna}"
                assert abs(float(a) - float(b)) <= 0.005, (
                    f"promedio incorrecto en fila {fila} columna {columna}: {b} != {a}"
                )
                verificadas += 1
        print(f"    O/P/Q verificadas: {verificadas:,}, todas correctas")

        # 5. Los productos sin salidas positivas deben quedar vacios.
        sin_salidas = resumen[resumen["prom_mensual"].isna()]
        assert len(sin_salidas) == 876, f"esperaba 876 productos sin promedio, hay {len(sin_salidas)}"
        ejemplo = sin_salidas.iloc[0]
        fila_ejemplo = int(movimientos[movimientos["codigo"] == ejemplo["codigo"]]["fila_excel"].iloc[0])
        assert nuevo.cell(fila_ejemplo, 15).value is None, (
            "un producto sin salidas positivas debe quedar con O vacia"
        )
        print(f"    productos sin promedio: {len(sin_salidas):,} (O vacia, no cero)")

        print("\nPrueba del archivo sin promedios superada.")
        return 0
    finally:
        shutil.rmtree(temporal, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
