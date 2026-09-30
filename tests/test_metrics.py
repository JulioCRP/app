"""Prueba de regresion contra el reporte original.

Reproduce los promedios de las columnas O, P y Q del archivo y exige coincidencia
total. Si el reporte de referencia cambia, esta prueba avisa.
"""

from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

from src import metrics  # noqa: E402
from src.parser import parsear_reporte  # noqa: E402

ARCHIVO = RAIZ / "Movimientos_con_promedios.xlsx"
TOL = 0.005


def test_redondear_excel() -> None:
    assert metrics.redondear_excel(25.125) == 25.13, "REDONDEAR(25.125) debe ser 25.13"
    assert metrics.redondear_excel(25.135) == 25.14
    assert metrics.redondear_excel(2.625) == 2.63
    assert metrics.redondear_excel(-2.625) == -2.63
    assert metrics.redondear_excel(1.005) == 1.01
    assert metrics.redondear_excel(0.0) == 0.0
    print("ok  test_redondear_excel")


def test_sin_salidas_queda_vacio() -> None:
    """Un producto sin ninguna salida debe quedar vacio, no en cero.

    Es la regla que aplica el reporte original: un promedio en cero seria un
    promedio inventado.
    """
    mov = pd.DataFrame(
        {
            "codigo": ["A", "A", "B"],
            "cant_salida": [0.0, 0.0, 3.0],
            "val_salida": [0.0, 0.0, 30.0],
            "fecha": pd.to_datetime(["2026-01-01", "2026-01-02", "2026-01-03"]),
            "mes": [1, 1, 1],
            "semestre": [1, 1, 1],
            "anio": [2026, 2026, 2026],
            "nombre": ["x", "x", "y"],
        }
    )
    r = metrics.resumen_productos(mov)
    assert len(r) == 2, "el producto sin salidas debe seguir apareciendo en el resumen"
    a = r[r["codigo"] == "A"].iloc[0]
    b = r[r["codigo"] == "B"].iloc[0]
    for columna in ("prom_mensual", "prom_semestral", "prom_anual"):
        assert np.isnan(a[columna]), f"{columna} debe quedar vacia sin salidas"
    assert a["total_salidas"] == 0.0
    assert b["prom_mensual"] == 3.0
    assert b["n_meses"] == 1
    print("ok  test_sin_salidas_queda_vacio")


def test_cuenta_solo_meses_con_salida() -> None:
    """Un mes con entradas pero sin salidas no debe entrar al divisor."""
    mov = pd.DataFrame(
        {
            "codigo": ["A", "A", "A"],
            "cant_salida": [0.0, 2.0, 4.0],
            "val_salida": [0.0, 0.0, 0.0],
            "fecha": pd.to_datetime(["2026-01-01", "2026-02-01", "2026-03-01"]),
            "mes": [1, 2, 3],
            "semestre": [1, 1, 1],
            "anio": [2026, 2026, 2026],
            "nombre": ["x", "x", "x"],
        }
    )
    r = metrics.resumen_productos(mov).iloc[0]
    assert r["n_meses"] == 2, "solo febrero y marzo tuvieron salidas"
    assert r["total_salidas"] == 6
    assert r["prom_mensual"] == 3.0, "6 salidas / 2 meses = 3"
    assert r["n_semestres"] == 1
    assert r["prom_semestral"] == 6.0, "6 salidas / 1 semestre = 6"
    print("ok  test_cuenta_solo_meses_con_salida")


def test_reproduce_columnas_originales() -> None:
    """El criterio de aceptacion: 100% de coincidencia con O, P y Q.

    Aqui "coincidir" incluye los productos cuyo promedio esta vacio tanto en el
    archivo como en el calculo, que es el caso de los 876 productos sin salidas
    positivas.
    """
    resultado = parsear_reporte(ARCHIVO)
    mov = resultado.movimientos
    prods = resultado.productos
    resumen = metrics.resumen_productos(mov, prods)

    print(
        f"    productos={len(resumen)}  movimientos={len(mov)}  "
        f"bloques={len(prods)}  con movimientos={int((prods['n_movimientos'] > 0).sum())}"
    )

    verificacion = metrics.verificar_promedios(mov, resumen)
    assert verificacion["disponible"], "el archivo de referencia deberia traer O, P y Q"

    fallos: list[str] = []
    for fila in verificacion["resumen"].itertuples():
        print(
            f"    col {fila.columna}  coinciden {fila.coinciden}/{fila.productos}  "
            f"difieren {fila.difieren}  ambos vacios {fila.ambos_vacios}  "
            f"dif_max={fila.dif_max:.4f}"
        )
        if fila.difieren:
            fallos.append(f"columna {fila.columna}: {fila.difieren} discrepancias")

    if fallos:
        detalle = verificacion["detalle"]
        for columna in ("col_14", "col_15", "col_16"):
            if columna in detalle:
                malas = detalle[detalle[f"dif_{columna.replace('col_', 'prom_')}"] > TOL]
                if not malas.empty:
                    print(f"    DETALLE {columna}:")
                    print(malas.head(10).to_string())

    assert not fallos, "; ".join(fallos)
    assert verificacion["coincidencia"] == verificacion["total"], "deben coincidir todos"
    print("ok  test_reproduce_columnas_originales")


def test_sin_columnas_de_promedio_aveces() -> None:
    """Un archivo que solo trae A-N debe verificarse como 'no disponible'."""
    mov = pd.DataFrame(
        {
            "codigo": ["A", "A"],
            "cant_salida": [1.0, 2.0],
            "val_salida": [0.0, 0.0],
            "fecha": pd.to_datetime(["2026-01-01", "2026-02-01"]),
            "mes": [1, 2],
            "semestre": [1, 1],
            "anio": [2026, 2026],
            "nombre": ["x", "x"],
            "col_14": [np.nan, np.nan],
            "col_15": [np.nan, np.nan],
            "col_16": [np.nan, np.nan],
        }
    )
    resumen = metrics.resumen_productos(mov)
    v = metrics.verificar_promedios(mov, resumen)
    assert v["disponible"] is False
    print("ok  test_sin_columnas_de_promedio_aveces")


def main() -> int:
    print("Ejecutando pruebas...")
    test_redondear_excel()
    test_sin_salidas_queda_vacio()
    test_cuenta_solo_meses_con_salida()
    test_sin_columnas_de_promedio_aveces()
    test_reproduce_columnas_originales()
    print("\nTodas las pruebas pasaron.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
