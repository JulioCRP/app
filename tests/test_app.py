"""Prueba de la app con AppTest (sin navegador ni servidor).

Simula la carga del archivo real y comprueba que el dashboard se construye
completo, que los promedios se calculan y que el exportador responde.
"""

from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

ARCHIVO = RAIZ / "Movimientos_con_promedios.xlsx"


def test_app_arranca_sin_archivo() -> None:
    """Sin archivo, la app debe mostrar la portada y parar sin errores."""
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(str(RAIZ / "app.py"), default_timeout=120).run()
    assert not at.exception, f"excepcion en la portada: {at.exception}"
    assert len(at.title) == 1
    assert any("Promedios de salida de inventario" in t.value for t in at.title)
    assert len(at.file_uploader) == 1, "debe pedir el archivo"
    print("ok  test_app_arranca_sin_archivo")


def test_app_con_archivo_real() -> None:
    """Con el archivo real: KPIs, pestanas, tabla y descarga."""
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(str(RAIZ / "app.py"), default_timeout=600)
    at.run()
    assert not at.exception, f"excepcion al arrancar: {at.exception}"

    # Simular la carga del archivo.
    at.file_uploader[0].upload(ARCHIVO.name, ARCHIVO.read_bytes()).run()

    assert not at.exception, f"excepcion al procesar: {at.exception}"
    print(f"    metricas renderizadas: {len(at.metric)}")
    assert len(at.metric) >= 4, "se esperaban al menos 4 indicadores"
    etiquetas = {m.label for m in at.metric}
    for esperada in ("Productos", "Unidades salidas"):
        assert esperada in etiquetas, f"falta el indicador {esperada}"

    assert len(at.tabs) == 4, "se esperaban 4 pestanas"
    print(f"    pestanas: {[t.label for t in at.tabs]}")

    assert len(at.dataframe) >= 1, "se esperaba la tabla de promedios"
    assert len(at.button) >= 1, "se esperaba el boton de preparar archivo"

    # La verificacion contra el archivo de referencia debe pasar.
    if at.success:
        textos = " ".join(s.value for s in at.success)
        assert "Verificacion superada" in textos, f"verificacion no superada: {textos}"
        print("    verificacion superada en la interfaz")
    else:
        raise AssertionError("no se mostro la verificacion superada")

    # Preparar el archivo exportado.
    preparar = [b for b in at.button if b.label == "Preparar archivo"]
    assert preparar, "no se encontro el boton 'Preparar archivo'"
    preparar[0].click().run()
    assert not at.exception, f"excepcion al exportar: {at.exception}"

    descargas = [d for d in at.get("download_button") if d.label == "Descargar Excel con promedios"]
    assert descargas, (
        "no se genero el boton de descarga de Excel; "
        f"hay: {[d.label for d in at.get('download_button')]}"
    )
    # En AppTest el contenido del boton vive en session_state, no en .value
    # (que es un booleano de activacion).
    assert at.session_state["nombre_descarga"] == "Movimientos_con_promedios_con_promedios.xlsx"
    contenido = at.session_state["descarga"]
    assert contenido, "la descarga esta vacia"
    assert isinstance(contenido, (bytes, bytearray))
    print(f"    descarga generada: {len(contenido):,} bytes")

    # El archivo descargado debe ser un xlsx valido con promedios.
    import io

    import openpyxl

    libro = openpyxl.load_workbook(io.BytesIO(contenido), data_only=True)
    hoja = libro.worksheets[0]
    assert hoja.cell(2, 15).value == "Promedio mensual", "falta el encabezado de O"
    assert hoja.cell(2, 16).value == "Promedio semestral", "falta el encabezado de P"
    assert hoja.cell(2, 17).value == "Promedio anual", "falta el encabezado de Q"
    assert hoja.max_row == 40700, f"filas inesperadas: {hoja.max_row}"
    libro.close()
    print("ok  test_app_con_archivo_real")


def main() -> int:
    print("Ejecutando pruebas de la app...")
    test_app_arranca_sin_archivo()
    test_app_con_archivo_real()
    print("\nTodas las pruebas de la app pasaron.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
