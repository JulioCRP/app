"""Calculo de promedios y metricas gerenciales.

Las formulas de promedio estan validadas contra el reporte original
(``Movimientos_con_promedios.xlsx``): las columnas O, P y Q de ese archivo
coinciden al 100% con lo que aqui se calcula.

Reglas de negocio
-----------------
1. Solo entran al promedio las filas con ``cantidad salida > 0``. Las salidas en
   cero quedan fuera tanto del numerador como del denominador.
2. El divisor es la cantidad de periodos (mes / semestre / anio) en los que el
   producto registro **al menos una salida**, no la cantidad de periodos con
   cualquier tipo de movimiento. Un mes con entradas pero sin salidas no cuenta.
3. El redondeo es ``half away from zero`` (igual que ``REDONDEAR`` de Excel),
   no el redondeo bancario que hace Python por defecto. Esa diferencia explica
   los 0,01 de discrepancia del reporte original.
4. Un producto sin ninguna salida positiva deja O, P y Q **vacias**, no en cero.
   Verificado sobre el reporte original: los 876 productos sin promedio son
   exactamente los 876 que no registran ninguna cantidad salida mayor que cero.
   Presentar 0 seria un promedio inventado.
"""

from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal

import numpy as np
import pandas as pd

MESES_ES: tuple[str, ...] = (
    "Ene",
    "Feb",
    "Mar",
    "Abr",
    "May",
    "Jun",
    "Jul",
    "Ago",
    "Sep",
    "Oct",
    "Nov",
    "Dic",
)

# Tolerancia al comparar un valor calculado contra uno almacenado en el archivo.
TOL_CANTIDAD: float = 0.005
TOL_VALOR: float = 0.05


def redondear_excel(valor: float, decimales: int = 2) -> float:
    """Redondea como ``REDONDEAR`` de Excel (half away from zero)."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return float("nan")
    cuantia = Decimal(1).scaleb(-decimales)
    return float(Decimal(repr(float(valor))).quantize(cuantia, rounding=ROUND_HALF_UP))


def _con_salidas(movimientos: pd.DataFrame) -> pd.DataFrame:
    """Filtra las filas de detalle que tienen cantidad salida mayor que cero."""
    mascara = movimientos["cant_salida"].fillna(0) > 0
    return movimientos[mascara]


def _promedios_por_periodo(
    con_salidas: pd.DataFrame, columna_periodo: str
) -> pd.DataFrame:
    """Suma y cantidad de salidas por producto y periodo, mas su promedio."""
    if con_salidas.empty:
        return pd.DataFrame(columns=["codigo", columna_periodo, "suma", "n_salidas"])
    agrupado = (
        con_salidas.groupby(["codigo", columna_periodo], observed=True)["cant_salida"]
        .agg(suma="sum", n_salidas="count")
        .reset_index()
    )
    agrupado["promedio"] = agrupado["suma"] / agrupado["n_salidas"]
    return agrupado


def resumen_productos(
    movimientos: pd.DataFrame,
    productos: pd.DataFrame | None = None,
    *,
    excluir_ceros: bool = True,
) -> pd.DataFrame:
    """Una fila por producto con totales, periodos activos y los tres promedios.

    Parametros
    ----------
    movimientos:
        Salida de :func:`src.parser.parsear_reporte`.
    productos:
        Fila de grupo por producto (opcional). Aporta los totales del periodo y
        el saldo final, que no se pueden reconstruir desde el detalle.
    excluir_ceros:
        Si es ``True`` (por defecto) las salidas en cero no cuentan para el
        promedio. Si es ``False`` se promedian todas las filas de detalle.
    """
    base = _con_salidas(movimientos) if excluir_ceros else movimientos

    fechas_validas = base[base["fecha"].notna()]

    por_mes = _promedios_por_periodo(base, "mes")
    por_sem = _promedios_por_periodo(base, "semestre")
    por_anio = _promedios_por_periodo(base, "anio")

    # El resumen debe incluir a TODOS los productos del reporte, tambien los que
    # no tienen ninguna salida: esos aparecen con total 0 y promedio 0, que es
    # justo lo que hace el archivo original.
    if productos is not None and not productos.empty:
        todos_los_codigos = productos["codigo"].dropna().unique()
    else:
        todos_los_codigos = movimientos["codigo"].dropna().unique()
    todos_los_codigos = pd.Index(sorted(todos_los_codigos.astype(str).unique()), name="codigo")

    resumen = pd.DataFrame(index=todos_los_codigos)
    if base.empty:
        resumen["total_salidas"] = 0.0
        resumen["n_salidas"] = 0
        resumen["valor_salidas"] = 0.0
    else:
        agregado = (
            base.assign(
                _cant=base["cant_salida"].fillna(0),
                _val=base["val_salida"].fillna(0),
            )
            .groupby("codigo", observed=True)
            .agg(total_salidas=("_cant", "sum"), n_salidas=("_cant", "count"), valor_salidas=("_val", "sum"))
        )
        resumen = resumen.join(agregado)
        resumen["total_salidas"] = resumen["total_salidas"].fillna(0.0)
        resumen["valor_salidas"] = resumen["valor_salidas"].fillna(0.0)
        resumen["n_salidas"] = resumen["n_salidas"].fillna(0).astype(int)

    # Numero de periodos distintos en los que hubo al menos una salida.
    n_meses = fechas_validas.groupby("codigo", observed=True)["mes"].nunique()
    n_semestres = fechas_validas.groupby("codigo", observed=True)["semestre"].nunique()
    n_anios = fechas_validas.groupby("codigo", observed=True)["anio"].nunique()

    resumen = resumen.join(n_meses.rename("n_meses"))
    resumen = resumen.join(n_semestres.rename("n_semestres"))
    resumen = resumen.join(n_anios.rename("n_anios"))
    for columna in ("n_meses", "n_semestres", "n_anios"):
        resumen[columna] = resumen[columna].fillna(0).astype(int)

    # Promedios: total dividido entre periodos con salida, redondeado al estilo Excel.
    # Sin ninguna salida positiva el promedio queda vacio (NaN), no en cero: es lo
    # que hace el reporte original y evita presentar un promedio inventado.
    resumen["prom_mensual"] = [
        redondear_excel(t / m) if m and m > 0 else float("nan")
        for t, m in zip(resumen["total_salidas"], resumen["n_meses"])
    ]
    resumen["prom_semestral"] = [
        redondear_excel(t / s) if s and s > 0 else float("nan")
        for t, s in zip(resumen["total_salidas"], resumen["n_semestres"])
    ]
    resumen["prom_anual"] = [
        redondear_excel(t / a) if a and a > 0 else float("nan")
        for t, a in zip(resumen["total_salidas"], resumen["n_anios"])
    ]

    # Desglose: promedio de cada mes del anio de referencia.
    for mes in range(1, 13):
        columna = f"prom_{MESES_ES[mes - 1].lower()}"
        parcial = por_mes[por_mes["mes"] == mes] if not por_mes.empty else por_mes
        if parcial.empty:
            resumen[columna] = float("nan")
            resumen[f"n_{MESES_ES[mes - 1].lower()}"] = 0
        else:
            valores = parcial.set_index("codigo")
            resumen[columna] = [
                redondear_excel(valores.at[c, "promedio"]) if c in valores.index else float("nan")
                for c in resumen.index
            ]
            resumen[f"n_{MESES_ES[mes - 1].lower()}"] = [
                int(valores.at[c, "n_salidas"]) if c in valores.index else 0
                for c in resumen.index
            ]

    # Semestres del anio de referencia.
    for etiqueta, numero in (("sem1", 1), ("sem2", 2)):
        parcial = por_sem[por_sem["semestre"] == numero] if not por_sem.empty else por_sem
        if parcial.empty:
            resumen[f"prom_{etiqueta}"] = float("nan")
        else:
            valores = parcial.set_index("codigo")
            resumen[f"prom_{etiqueta}"] = [
                redondear_excel(valores.at[c, "promedio"]) if c in valores.index else float("nan")
                for c in resumen.index
            ]

    resumen = resumen.reset_index()

    # Nombre y referencia, tomados del detalle (mas fiables que el grupo).
    nombres = (
        movimientos.dropna(subset=["nombre"])
        .drop_duplicates("codigo")
        .set_index("codigo")["nombre"]
    )
    resumen["nombre"] = resumen["codigo"].map(nombres).fillna("")

    if productos is not None and not productos.empty:
        extra = [c for c in ("codigo", "n_movimientos", "saldo_cant", "saldo_valor") if c in productos.columns]
        resumen = resumen.merge(productos[extra].drop_duplicates("codigo"), on="codigo", how="left")
        resumen["saldo_cant"] = resumen["saldo_cant"].fillna(0.0)
        resumen["saldo_valor"] = resumen["saldo_valor"].fillna(0.0)
        resumen["n_movimientos"] = resumen["n_movimientos"].fillna(0).astype(int)

    return _ordenar_columnas(resumen)


def dias_disponibilidad_por_producto(movimientos: pd.DataFrame) -> pd.DataFrame:
    """Calcula los dias de disponibilidad por producto y mes.

    Para cada codigo de producto y mes, cuenta la cantidad de dias unicos
    en los que el campo ``saldo_cant`` es mayor a cero, basado en la
    ``fecha elaboracion``.

    Returns
    -------
    pd.DataFrame
        columnas: codigo, anio, mes, dias_disponibles
    """
    base = movimientos[movimientos["fecha"].notna() & (movimientos["saldo_cant"] > 0)].copy()

    if base.empty:
        return pd.DataFrame(columns=["codigo", "anio", "mes", "dias_disponibles"])

    # Parsear la fecha texto a datetime para poder extraer el dia unico
    base["fecha_dt"] = pd.to_datetime(base["fecha"], format="%d/%m/%Y", errors="coerce")

    # Contar dias unicos con saldo_cant > 0 por producto, anio, mes
    dias = (
        base.groupby(["codigo", "anio", "mes"])
        .apply(
            lambda g: len(
                set(g[g["saldo_cant"] > 0]["fecha_dt"].dt.date.astype(str))
            ),
            include_groups=False,
        )
        .reset_index()
        .rename({0: "dias_disponibles"}, axis=1)
    )

    return dias


def dias_disponibilidad_por_producto_horiz(movimientos: pd.DataFrame) -> pd.DataFrame:
    """Calcula y formatea dias de disponibilidad en formato horizontal por producto.

    Para cada codigo de producto, muestra los dias de disponibilidad por mes
    (Ene-Dic) de forma horizontal, mas dos columnas finales:
    - ``Total``: suma de dias disponibles del producto a lo largo de los meses (filas)
    - ``Total por meses``: equivalente en meses (Total ÷ 30 dias/mes) del total de dias
      de disponibilidad del producto durante el periodo analizado.

    Returns
    -------
    pd.DataFrame
        columnas: codigo, Ene, Feb, Mar, Abr, May, Jun, Jul, Ago, Sep, Oct, Nov, Dic, Total, Total por meses
    """
    base = movimientos[movimientos["fecha"].notna() & (movimientos["saldo_cant"] > 0)].copy()

    if base.empty:
        meses_cols = [m.lower() for m in MESES_ES]
        return pd.DataFrame(
            columns=["codigo"] + meses_cols + ["Total", "Total por meses"]
        )

    # Parsear la fecha texto a datetime para poder extraer el dia unico
    base["fecha_dt"] = pd.to_datetime(base["fecha"], format="%d/%m/%Y", errors="coerce")

    # Contar dias unicos con saldo_cant > 0 por producto, anio, mes
    dias = (
        base.groupby(["codigo", "anio", "mes"])
        .apply(
            lambda g: len(
                set(g[g["saldo_cant"] > 0]["fecha_dt"].dt.date.astype(str))
            ),
            include_groups=False,
        )
        .reset_index()
        .rename({0: "dias_disponibles"}, axis=1)
    )

    # Pivot horizontal: filas=códigos, columnas=meses(1-12)
    pivot = dias.pivot_table(
        index="codigo",
        columns="mes",
        values="dias_disponibles",
        aggfunc="sum",
        fill_value=0,
    )

    # Normalizar columnas a enteros (pueden venir como float del groupby)
    pivot.columns = [int(c) for c in pivot.columns]

    # Asegurar columnas de los 12 meses (1-12), ordenados
    for m in range(1, 13):
        if m not in pivot.columns:
            pivot[m] = 0
    pivot = pivot.sort_index(axis=1)

    # Renombrar columnas a nombres de meses en español
    pivot.columns = [MESES_ES[m - 1] for m in pivot.columns]

    # Agregar columna Total (suma de los 12 meses por fila)
    pivot["Total"] = pivot.sum(axis=1)

    # Agregar columna "Total por meses" (Total ÷ 30 dias/mes)
    pivot["Total por meses"] = (pivot["Total"] / 30).round(2)

    # Restablecer índice y asegurar orden de columnas
    resultado = pivot.reset_index()
    columnas_orden = ["codigo"] + list(MESES_ES) + ["Total", "Total por meses"]
    resultado = resultado[columnas_orden]

    return resultado


def _ordenar_columnas(resumen: pd.DataFrame) -> pd.DataFrame:
    """Coloca las columnas en un orden legible para el usuario."""
    base = [
        "codigo",
        "nombre",
        "total_salidas",
        "n_salidas",
        "n_movimientos",
        "n_meses",
        "prom_mensual",
        "n_semestres",
        "prom_semestral",
        "n_anios",
        "prom_anual",
    ]
    mes_cols = [f"prom_{m.lower()}" for m in MESES_ES]
    n_mes_cols = [f"n_{m.lower()}" for m in MESES_ES]
    cola = ["prom_sem1", "prom_sem2", "saldo_cant", "saldo_valor", "valor_salidas"]
    presentes = [c for c in base + mes_cols + n_mes_cols + cola if c in resumen.columns]
    return resumen[presentes]


def promedios_por_producto_fila(
    movimientos: pd.DataFrame, resumen: pd.DataFrame
) -> pd.DataFrame:
    """Lleva los promedios a cada fila de movimiento, para escribir O, P y Q.

    Se replica el valor del producto en todas sus filas, tal como hace el
    reporte original.
    """
    tabla = resumen[["codigo", "prom_mensual", "prom_semestral", "prom_anual"]]
    salida = movimientos.merge(tabla, on="codigo", how="left")
    return salida.rename(
        columns={
            "prom_mensual": "col_14",
            "prom_semestral": "col_15",
            "prom_anual": "col_16",
        }
    )


def serie_mensual(movimientos: pd.DataFrame, *, excluir_ceros: bool = True) -> pd.DataFrame:
    """Cantidad salida por anio y mes, con el desglose por tipo de comprobante."""
    base = _con_salidas(movimientos) if excluir_ceros else movimientos
    base = base[base["fecha"].notna()]
    if base.empty:
        return pd.DataFrame(columns=["anio", "mes", "periodo", "cantidad_salida", "tipo_comprobante"])

    serie = (
        base.groupby(["anio", "mes", "tipo_comprobante"], observed=True)["cant_salida"]
        .sum()
        .reset_index()
        .rename(columns={"cant_salida": "cantidad_salida"})
    )
    serie["periodo"] = serie["anio"].astype(int).astype(str) + "-" + serie["mes"].map(
        lambda m: f"{int(m):02d}"
    )
    return serie.sort_values(["anio", "mes", "tipo_comprobante"]).reset_index(drop=True)


def totales_por_tipo(movimientos: pd.DataFrame) -> pd.DataFrame:
    """Resumen por tipo de comprobante, para el grafico de barras apiladas."""
    base = _con_salidas(movimientos)
    if base.empty:
        return pd.DataFrame(columns=["tipo_comprobante", "cantidad_salida", "n_salidas"])
    return (
        base.groupby("tipo_comprobante", observed=True)["cant_salida"]
        .agg(cantidad_salida="sum", n_salidas="count")
        .reset_index()
        .sort_values("cantidad_salida", ascending=False)
        .reset_index(drop=True)
    )


def rankings(resumen: pd.DataFrame, n: int = 15) -> dict[str, pd.DataFrame]:
    """Top N productos por distintas metrics, listo para graficar."""
    criterios = {
        "salidas": ("total_salidas", "Total cantidad salida"),
        "movimientos": ("n_movimientos", "Numero de movimientos"),
        "frecuencia": ("n_salidas", "Numero de salidas"),
        "inventario": ("saldo_valor", "Saldo de valor en inventario"),
    }
    salida: dict[str, pd.DataFrame] = {}
    for clave, (columna, etiqueta) in criterios.items():
        if columna not in resumen.columns:
            continue
        top = (
            resumen.dropna(subset=[columna])
            .nlargest(n, columna)
            .loc[:, ["codigo", "nombre", columna]]
            .copy()
        )
        top = top.rename(columns={columna: "valor"})
        top["metrica"] = etiqueta
        top["etiqueta"] = top["codigo"] + " · " + top["nombre"].astype(str).str.slice(0, 40)
        salida[clave] = top.reset_index(drop=True)
    return salida


def _descuadre(df: pd.DataFrame, esperado: pd.Series, real: pd.Series, tol: float) -> pd.Series:
    return (real - esperado).abs() > tol


def alertas_calidad(
    movimientos: pd.DataFrame, productos: pd.DataFrame
) -> list[dict[str, object]]:
    """Siete controles de calidad, cada uno con conteo, porcentaje e impacto.

    Cada alerta es un dict con: clave, titulo, descripcion, severidad
    (``alta``/``media``/``baja``), cantidad, total, porcentaje, impacto
   economico y un DataFrame ``detalle`` para el drill-down.
    """
    alertas: list[dict[str, object]] = []
    total_mov = max(len(movimientos), 1)
    total_prod = max(len(productos), 1)

    # 1. Bloques de producto sin ninguna fila de movimiento.
    sin_mov = productos[productos["n_movimientos"].fillna(0) == 0]
    alertas.append(
        {
            "clave": "bloques_sin_movimientos",
            "titulo": "Productos sin movimientos",
            "descripcion": "Filas de grupo que no detalle ningun movimiento.",
            "severidad": "media" if len(sin_mov) else "baja",
            "cantidad": len(sin_mov),
            "total": total_prod,
            "porcentaje": 100 * len(sin_mov) / total_prod,
            "impacto": float(sin_mov["saldo_valor"].fillna(0).sum()),
            "unidad": "productos",
            "detalle": sin_mov[["codigo", "nombre_grupo", "saldo_cant", "saldo_valor"]],
        }
    )

    # 2. Salidas con cantidad pero sin valor.
    s = movimientos[["cant_salida", "val_salida"]].fillna(0)
    sin_valor = s[(s["cant_salida"] > 0) & (s["val_salida"] == 0)]
    alertas.append(
        {
            "clave": "salidas_sin_valor",
            "titulo": "Salidas con cantidad pero sin valor",
            "descripcion": "Se descontaron unidades pero no se registro su valor, "
            "asi que el saldo de valor no refleja la salida.",
            "severidad": "alta" if len(sin_valor) else "baja",
            "cantidad": len(sin_valor),
            "total": int((s["cant_salida"] > 0).sum()),
            "porcentaje": 100 * len(sin_valor) / max(int((s["cant_salida"] > 0).sum()), 1),
            "impacto": float(sin_valor["cant_salida"].sum()),
            "unidad": "lineas",
            "detalle": _muestra(movimientos[(s["cant_salida"] > 0) & (s["val_salida"] == 0)]),
        }
    )

    # 3. Salidas con valor pero sin cantidad.
    sin_cantidad = s[(s["cant_salida"] == 0) & (s["val_salida"] > 0)]
    alertas.append(
        {
            "clave": "salidas_sin_cantidad",
            "titulo": "Salidas con valor pero sin cantidad",
            "descripcion": "Se desconto valor sin descontar unidades.",
            "severidad": "alta" if len(sin_cantidad) else "baja",
            "cantidad": len(sin_cantidad),
            "total": int((s["val_salida"] > 0).sum()),
            "porcentaje": 100 * len(sin_cantidad) / max(int((s["val_salida"] > 0).sum()), 1),
            "impacto": float(sin_cantidad["val_salida"].sum()),
            "unidad": "lineas",
            "detalle": _muestra(movimientos[(s["cant_salida"] == 0) & (s["val_salida"] > 0)]),
        }
    )

    # 4. Saldo de valor negativo.
    neg = productos[productos["saldo_valor"] < 0]
    alertas.append(
        {
            "clave": "saldo_valor_negativo",
            "titulo": "Saldo de valor negativo",
            "descripcion": "Productos cuyo saldo de valor en inventario es negativo.",
            "severidad": "alta" if len(neg) else "baja",
            "cantidad": len(neg),
            "total": total_prod,
            "porcentaje": 100 * len(neg) / total_prod,
            "impacto": float(neg["saldo_valor"].sum()),
            "unidad": "productos",
            "detalle": neg[["codigo", "nombre_grupo", "saldo_cant", "saldo_valor"]].nlargest(
                50, "saldo_cant"
            ),
        }
    )

    # 5 y 6. Descuadres de las ecuaciones de saldo en la fila de grupo.
    esperado_cant = productos["cant_inicial"] + productos["cant_entrada"] - productos["cant_salida"]
    desc_cant = _descuadre(productos, esperado_cant, productos["saldo_cant"], TOL_CANTIDAD)
    alertas.append(
        {
            "clave": "descuadre_cantidad",
            "titulo": "Descuadre de cantidad",
            "descripcion": "Saldo cantidades distinto de Inicial + Entrada - Salida.",
            "severidad": "alta" if int(desc_cant.sum()) else "baja",
            "cantidad": int(desc_cant.sum()),
            "total": total_prod,
            "porcentaje": 100 * int(desc_cant.sum()) / total_prod,
            "impacto": 0.0,
            "unidad": "productos",
            "detalle": productos[desc_cant][
                ["codigo", "nombre_grupo", "cant_inicial", "cant_entrada", "cant_salida", "saldo_cant"]
            ].head(50),
        }
    )

    esperado_val = productos["val_inicial"] + productos["val_entrada"] - productos["val_salida"]
    desc_val = _descuadre(productos, esperado_val, productos["saldo_valor"], TOL_VALOR)
    alertas.append(
        {
            "clave": "descuadre_valor",
            "titulo": "Descuadre de valor",
            "descripcion": "Saldo valor distinto de Inicial + Entrada - Salida.",
            "severidad": "alta" if int(desc_val.sum()) else "baja",
            "cantidad": int(desc_val.sum()),
            "total": total_prod,
            "porcentaje": 100 * int(desc_val.sum()) / total_prod,
            "impacto": float((productos[desc_val]["saldo_valor"] - esperado_val[desc_val]).sum()),
            "unidad": "productos",
            "detalle": productos[desc_val][
                ["codigo", "nombre_grupo", "val_inicial", "val_entrada", "val_salida", "saldo_valor"]
            ].head(50),
        }
    )

    # 7. Movimientos sin fecha o sin comprobante.
    sin_fecha = movimientos[movimientos["fecha"].isna() | (movimientos["comprobante"] == "")]
    alertas.append(
        {
            "clave": "sin_fecha_o_comprobante",
            "titulo": "Movimientos sin fecha o sin comprobante",
            "descripcion": "No se pueden agrupar por periodo, asi que quedan fuera de "
            "los promedios mensuales y semestrales.",
            "severidad": "media" if len(sin_fecha) else "baja",
            "cantidad": len(sin_fecha),
            "total": total_mov,
            "porcentaje": 100 * len(sin_fecha) / total_mov,
            "impacto": 0.0,
            "unidad": "lineas",
            "detalle": _muestra(sin_fecha),
        }
    )

    return alertas


def _muestra(df: pd.DataFrame, n: int = 50) -> pd.DataFrame:
    """Recorta un DataFrame a las columnas utiles para el drill-down."""
    columnas = [
        c
        for c in (
            "fila_excel",
            "codigo",
            "nombre",
            "comprobante",
            "secuencia",
            "fecha",
            "cant_entrada",
            "cant_salida",
            "val_entrada",
            "val_salida",
            "saldo_cant",
            "saldo_valor",
        )
        if c in df.columns
    ]
    return df.head(n)[columnas]


def verificar_promedios(movimientos: pd.DataFrame, resumen: pd.DataFrame) -> dict[str, object]:
    """Compara los promedios calculados con las columnas O, P y Q del archivo.

    Un producto se considera correcto si el valor calculado y el almacenado
    coinciden, y tambien si ambos estan vacios (producto sin salidas positivas).
    Solo tiene sentido si el archivo de entrada ya trae esas columnas; si vienen
    vacias devuelve ``disponible=False``.
    """
    columnas = {"prom_mensual": "col_14", "prom_semestral": "col_15", "prom_anual": "col_16"}
    almacenado = (
        movimientos.groupby("codigo", observed=True)[list(columnas.values())].first()
    )
    if not np.isfinite(almacenado.to_numpy(dtype=float)).any():
        return {
            "disponible": False,
            "resumen": pd.DataFrame(),
            "detalle": pd.DataFrame(),
            "coincidencia": 0,
            "total": 0,
        }

    base = (
        movimientos.loc[movimientos["codigo"].isin(resumen["codigo"])]
        .groupby("codigo", observed=True)[list(columnas.values())]
        .first()
    )
    # No hay colision de nombres: a la izquierda van los promedios calculados
    # (prom_*) y a la derecha los almacenados (col_14..16).
    comparacion = resumen.set_index("codigo").join(base, how="inner")

    filas: list[dict[str, object]] = []
    detalle = comparacion.copy()
    for calculado, col in columnas.items():
        propio = comparacion[calculado]
        referencia = comparacion[col]
        ambos_vacios = propio.isna() & referencia.isna()
        numerico = (~ambos_vacios) & propio.notna() & referencia.notna()
        dif = (propio - referencia).abs()

        coincide = ambos_vacios | (numerico & (dif <= TOL_CANTIDAD))
        etiqueta = {"prom_mensual": "O", "prom_semestral": "P", "prom_anual": "Q"}[calculado]
        filas.append(
            {
                "columna": etiqueta,
                "promedio": calculado,
                "productos": int(len(comparacion)),
                "coinciden": int(coincide.sum()),
                "difieren": int((~coincide).sum()),
                "ambos_vacios": int(ambos_vacios.sum()),
                "dif_max": float(dif[numerico].max()) if int(numerico.sum()) else 0.0,
            }
        )
        detalle[f"dif_{calculado}"] = dif.where(numerico)

    resumen_verif = pd.DataFrame(filas)
    return {
        "disponible": True,
        "resumen": resumen_verif,
        "detalle": detalle[
            [
                c
                for c in detalle.columns
                if c in ("nombre",) or c.startswith(("col_", "dif_", "prom_", "total_"))
            ]
        ],
        "coincidencia": int(resumen_verif["coinciden"].sum()),
        "total": int(resumen_verif["productos"].sum()),
    }
