"""Prueba del parser contra el archivo real: valida la lectura y reimprime las
cifras de referencia del reporte (que corregiran un analisis previo con un
desfase de clasificacion)."""

from __future__ import annotations

import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from src.parser import parsear_reporte  # noqa: E402

ARCHIVO = RAIZ / "Movimientos_con_promedios.xlsx"


def main() -> None:
    inicio = time.perf_counter()
    resultado = parsear_reporte(ARCHIVO)
    elapsed = time.perf_counter() - inicio

    mov = resultado.movimientos
    prods = resultado.productos

    print(f"tiempo de lectura      : {elapsed:.2f} s")
    print(f"fila de encabezado     : {resultado.fila_encabezado}")
    print(f"encabezado detectado   : {resultado.encabezado}")
    print(f"bloques de producto    : {len(prods)}")
    print(f"filas de movimiento    : {len(mov)}")
    for aviso in resultado.avisos:
        print(f"  aviso: {aviso}")

    con_mov = prods["n_movimientos"] if "n_movimientos" in prods else pd_counts(prods)
    n_con = int((prods["n_movimientos"] > 0).sum()) if "n_movimientos" in prods else None
    print(f"bloques con movimiento : {n_con}")
    print(f"bloques sin movimiento: {len(prods) - (n_con or 0)}")

    print("\n--- codigos repetidos entre bloques ---")
    repetidos = prods["codigo"][prods["codigo"].duplicated(keep=False)]
    print(f"  filas implicadas: {len(repetidos)} | codigos distintos: {repetidos.nunique()}")

    print("\n--- tipos de comprobante ---")
    print(mov["tipo_comprobante"].value_counts().to_string())

    print("\n--- fechas ---")
    validas = mov["fecha"].notna().sum()
    print(f"  con fecha valida: {validas} de {len(mov)}")
    print(f"  rango: {mov['fecha'].min()} -> {mov['fecha'].max()}")
    print(f"  anios: {sorted(mov['anio'].dropna().unique().tolist())}")

    print("\n--- columnas de promedios presentes en el archivo ---")
    for col in ("col_14", "col_15", "col_16"):
        n = int(mov[col].notna().sum())
        print(f"  {col}: {n} filas con valor")

    print("\n--- integridad de saldos a nivel grupo ---")
    descq_cant = (prods["saldo_cant"] - (prods["cant_inicial"] + prods["cant_entrada"] - prods["cant_salida"])).abs() > 0.005
    descq_val = (prods["saldo_valor"] - (prods["val_inicial"] + prods["val_entrada"] - prods["val_salida"])).abs() > 0.05
    print(f"  descuadre de cantidad: {int(descq_cant.sum())} de {len(prods)}")
    print(f"  descuadre de valor  : {int(descq_val.sum())} de {len(prods)}")
    print(f"  saldo valor negativo: {int((prods['saldo_valor'] < 0).sum())}")
    print(f"  filas 100% vacias  : {int((prods[[c for c in prods.columns if c.startswith(('cant_', 'val_', 'saldo_'))]].fillna(0).sum(axis=1) == 0).sum())}")

    print("\n--- desacople cantidad / valor en salidas ---")
    s = mov[["cant_salida", "val_salida"]].fillna(0)
    con_salida = s["cant_salida"] > 0
    print(f"  filas de detalle     : {len(s)}")
    print(f"  con salida > 0       : {int(con_salida.sum())}")
    print(f"  salida sin cantidad  : {int((s['cant_salida'] == 0).sum())}")
    print(f"  salida cant>0 y val=0: {int((con_salida & (s['val_salida'] == 0)).sum())}")
    print(f"  salida cant=0 y val>0: {int(((s['cant_salida'] == 0) & (s['val_salida'] > 0)).sum())}")
    print(f"  sin ningun movimiento: {int(((s['cant_salida'] == 0) & (s['val_salida'] == 0)).sum())}")

    print("\n--- grupos (producto, anio, mes) con salidas ---")
    con = mov[mov["cant_salida"] > 0]
    grupos = con.groupby(["codigo", "anio", "mes"]).size()
    print(f"  grupos: {len(grupos)} | con una sola salida: {int((grupos == 1).sum())}")


def pd_counts(prods):
    return 0


if __name__ == "__main__":
    main()
