"""Diagnostico: por que 876 productos tienen O, P y Q vacias en el archivo
original? Comprueba si coinciden exactamente con los productos sin salidas
positivas, o si hay otro criterio de fondo.
"""

from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import pandas as pd  # noqa: E402

from src import metrics  # noqa: E402
from src.parser import parsear_reporte  # noqa: E402

ARCHIVO = RAIZ / "Movimientos_con_promedios.xlsx"


def main() -> None:
    r = parsear_reporte(ARCHIVO)
    mov, prods = r.movimientos, r.productos
    resumen = metrics.resumen_productos(mov, prods)

    original = mov.groupby("codigo", observed=True)[["col_14", "col_15", "col_16"]].first()
    t = resumen.set_index("codigo").join(original, how="left")
    t["tiene_prom"] = t["col_14"].notna()

    print(f"productos                       : {len(t)}")
    print(f"con O/P/Q en el archivo         : {int(t['tiene_prom'].sum())}")
    print(f"sin O/P/Q en el archivo         : {int((~t['tiene_prom']).sum())}")
    print(f"sin ninguna salida positiva (==0): {int((t['total_salidas'] == 0).sum())}")

    vacios = set(t.index[~t["tiene_prom"]])
    ceros = set(t.index[t["total_salidas"] == 0])
    print(f"\nblancos == sin salidas positivas? {vacios == ceros}")
    print(f"  solo en blanco (tiene salida) : {len(vacios - ceros)}")
    print(f"  solo en cero  (tiene blanco)  : {len(ceros - vacios)}")

    if vacios - ceros:
        print("\n--- productos con PROMEDIO pero SIN salidas positivas ---")
        cols = ["nombre", "total_salidas", "n_salidas", "n_movimientos", "col_14"]
        print(t.loc[sorted(vacios - ceros), cols].to_string())

    if ceros - vacios:
        print("\n--- productos SIN PROMEDIO pero con salidas ---")
        cols = ["nombre", "total_salidas", "n_salidas", "n_meses", "col_14", "prom_mensual"]
        print(t.loc[sorted(ceros - vacios), cols].to_string())

    # Que tienen exactamente los movimientos de estos productos?
    print("\n--- composicion de los productos con O/P/Q vacia ---")
    codigos_vacios = sorted(vacios)
    sub = prods[prods["codigo"].isin(codigos_vacios)]
    sub_mov = mov[mov["codigo"].isin(codigos_vacios)]
    print(f"  filas de grupo      : {len(sub)}")
    print(f"  con saldo cantidad>0: {int((sub['saldo_cant'] > 0).sum())}")
    print(f"  con saldo valor >0  : {int((sub['saldo_valor'] > 0).sum())}")
    print(f"  con cant_entrada >0 : {int((sub['cant_entrada'] > 0).sum())}")
    print(f"  con cant_salida  >0 : {int((sub['cant_salida'] > 0).sum())}")
    print(f"  filas de detalle    : {len(sub_mov)}")
    print(f"  detalle cant_salida>0: {int((sub_mov['cant_salida'].fillna(0) > 0).sum())}")
    print(f"  detalle val_salida >0: {int((sub_mov['val_salida'].fillna(0) > 0).sum())}")
    print("\n  tipos de comprobante en sus movimientos:")
    print(sub_mov["tipo_comprobante"].value_counts().to_string())

    print("\n--- en los productos CON promedio, que tienen en comun? ---")
    con_prom = set(t.index[t["tiene_prom"]])
    sub2 = mov[mov["codigo"].isin(con_prom)]
    print(f"  detalle cant_salida>0: {int((sub2['cant_salida'].fillna(0) > 0).sum())} de {len(sub2)}")
    print(f"  grupos con saldo>0   : "
          f"{int((prods[prods['codigo'].isin(con_prom)]['saldo_cant'] > 0).sum())} de {len(con_prom)}")


if __name__ == "__main__":
    main()
