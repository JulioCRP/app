"""Dashboard gerencial de promedios de salida de inventario.

Uso:
    streamlit run app.py

La app es local: el archivo se procesa en memoria y nunca se sube a ningun
servidor. El archivo original del usuario no se modifica en disco.
"""

from __future__ import annotations

import sys
from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

RAIZ = Path(__file__).resolve().parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from src import exporter, metrics  # noqa: E402
from src.parser import ResultadoParseo, parsear_reporte  # noqa: E402

st.set_page_config(
    page_title="Promedios de inventario",
    page_icon=":material/query_stats:",
    layout="wide",
)

# Paleta sobria y consistente en todos los graficos.
COLOR_SALIDAS = "#2E6F9E"
COLOR_ENTRADAS = "#7FA8C4"
COLOR_ALERTA = "#B4482E"

MESES_ES = metrics.MESES_ES


# --------------------------------------------------------------------------
# Carga de datos (cacheada: el archivo de referencia tarda ~14 s en leerse)
# --------------------------------------------------------------------------
@st.cache_data(show_spinner="Leyendo y analizando el archivo...")
def cargar(contenido: bytes, nombre: str) -> tuple[ResultadoParseo, pd.DataFrame, dict]:
    """Parsea el archivo, calcula promedios y devuelve tambien las alertas."""
    resultado = parsear_reporte(contenido)
    resumen = metrics.resumen_productos(resultado.movimientos, resultado.productos)
    alertas = metrics.alertas_calidad(resultado.movimientos, resultado.productos)
    verificacion = metrics.verificar_promedios(resultado.movimientos, resumen)
    return resultado, resumen, {
        "alertas": alertas,
        "verificacion": verificacion,
        "serie": metrics.serie_mensual(resultado.movimientos),
        "por_tipo": metrics.totales_por_tipo(resultado.movimientos),
        "rankings": metrics.rankings(resumen),
        "nombre": nombre,
    }


def formato_numero(valor: float, decimales: int = 0) -> str:
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return "-"
    return f"{valor:,.{decimales}f}"


def formato_moneda(valor: float) -> str:
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return "-"
    return f"{valor:,.0f}"


# --------------------------------------------------------------------------
# Portada mientras no hay archivo
# --------------------------------------------------------------------------
st.title("Promedios de salida de inventario")
st.caption(
    "Calcula los promedios mensual, semestral y anual de cada producto y los "
    "agrega a las columnas O, P y Q de tu archivo. Todo se procesa de forma local."
)

archivo = st.file_uploader(
    "Archivo de Excel con los movimientos",
    type=["xlsx", "xlsm"],
    help="El archivo puede traer solo las columnas A-N; las columnas O, P y Q se completan solas.",
)

if archivo is None:
    st.info("Sube un archivo .xlsx para comenzar.")
    st.markdown(
        """
**Como se calcula cada promedio**

| Columna | Promedio | Formula |
|---|---|---|
| O | Mensual | Total de cantidad salida ÷ meses con al menos una salida |
| P | Semestral | Total de cantidad salida ÷ semestres con al menos una salida |
| Q | Anual | Total de cantidad salida ÷ anios con al menos una salida |

Solo cuentan las filas con cantidad salida mayor que cero. Un producto sin
ninguna salida deja O, P y Q vacias, en lugar de mostrar un cero.
        """
    )
    st.stop()

contenido = archivo.getvalue()
resultado, resumen, contexto = cargar(contenido, archivo.name)
movimientos = resultado.movimientos
productos = resultado.productos
alertas = contexto["alertas"]

# Avisos del parser, para que el usuario sepa que se descarto algo.
if resultado.avisos:
    with st.expander(f"Avisos de lectura ({len(resultado.avisos)})"):
        for aviso in resultado.avisos:
            st.caption(aviso)

# --------------------------------------------------------------------------
# Filtros en la barra lateral
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("Filtros")
    busqueda = st.text_input("Buscar producto", placeholder="Codigo o nombre")
    anios_disponibles = sorted(
        int(a) for a in movimientos["anio"].dropna().unique().tolist()
    )
    anios = st.pills(
        "Anio",
        anios_disponibles,
        default=anios_disponibles,
        selection_mode="multi",
    )
    tipos = sorted(t for t in movimientos["tipo_comprobante"].dropna().unique() if t)
    tipos_sel = st.pills(
        "Tipo de comprobante",
        tipos,
        default=tipos,
        selection_mode="multi",
    )
    solo_con_promedio = st.toggle(
        "Solo productos con promedio", value=False,
        help="Excluye los productos que no registran ninguna salida positiva.",
    )
    st.divider()
    st.caption(f"Archivo: {contexto['nombre']}")
    st.caption(
        f"{len(productos):,} productos · {len(movimientos):,} movimientos"
    )

# Aplica los filtros. Son baratas porque los datos ya estan en memoria.
mov_filtrado = movimientos
if anios:
    mov_filtrado = mov_filtrado[mov_filtrado["anio"].isin(anios)]
if tipos_sel:
    mov_filtrado = mov_filtrado[mov_filtrado["tipo_comprobante"].isin(tipos_sel)]

codigos_filtrados = set(mov_filtrado["codigo"].dropna())
resumen_filtrado = resumen[resumen["codigo"].isin(codigos_filtrados)].copy()

if busqueda:
    termino = busqueda.strip().lower()
    resumen_filtrado = resumen_filtrado[
        resumen_filtrado["codigo"].str.lower().str.contains(termino, na=False)
        | resumen_filtrado["nombre"].str.lower().str.contains(termino, na=False)
    ]

if solo_con_promedio:
    resumen_filtrado = resumen_filtrado[resumen_filtrado["prom_mensual"].notna()]

if anios or tipos_sel or busqueda or solo_con_promedio:
    st.caption(
        f"Mostrando {len(resumen_filtrado):,} de {len(resumen):,} productos "
        f"({len(mov_filtrado):,} movimientos)."
    )

if resumen_filtrado.empty:
    st.warning("Ningun producto cumple los filtros seleccionados.")
    st.stop()


# --------------------------------------------------------------------------
# Indicadores
# --------------------------------------------------------------------------
serie_filtrada = metrics.serie_mensual(mov_filtrado)
con_salida = mov_filtrado[mov_filtrado["cant_salida"].fillna(0) > 0]

total_salidas = float(con_salida["cant_salida"].sum())
saldo_inventario = float(resumen_filtrado["saldo_valor"].sum())
productos_sin_promedio = int(resumen_filtrado["prom_mensual"].isna().sum())

# Tendencia de salidas por mes, para las sparklines de los indicadores.
por_periodo = (
    serie_filtrada.groupby(["anio", "mes"], observed=True)["cantidad_salida"].sum().reset_index()
    if not serie_filtrada.empty
    else pd.DataFrame(columns=["anio", "mes", "cantidad_salida"])
)
tendencia = por_periodo["cantidad_salida"].tolist() if not por_periodo.empty else []

# Ultimo mes con datos, para comparar contra el mes anterior.
variacion = None
if len(por_periodo) >= 2:
    ultimo = float(por_periodo["cantidad_salida"].iloc[-1])
    anterior = float(por_periodo["cantidad_salida"].iloc[-2])
    if anterior > 0:
        variacion = (ultimo - anterior) / anterior * 100

st.subheader("Resumen general")
with st.container(horizontal=True):
    st.metric(
        "Productos",
        f"{len(resumen_filtrado):,}",
        border=True,
    )
    st.metric(
        "Unidades salidas",
        formato_numero(total_salidas),
        help="Suma de cantidad salida de los productos y movimientos filtrados.",
        border=True,
        chart_data=tendencia,
        chart_type="bar",
    )
    st.metric(
        "Saldo de valor en inventario",
        formato_moneda(saldo_inventario),
        help="Suma del saldo de valor de los productos filtrados.",
        border=True,
    )
    st.metric(
        "Sin promedio calculado",
        f"{productos_sin_promedio:,}",
        help="Productos sin ninguna cantidad salida mayor que cero; sus columnas O, P y Q quedan vacias.",
        border=True,
    )

if variacion is not None:
    tendencia_txt = "sube" if variacion >= 0 else "baja"
    st.caption(
        f"En el ultimo mes con datos la cantidad salida {tendencia_txt} "
        f"{abs(variacion):.1f}% frente al mes anterior."
    )


# --------------------------------------------------------------------------
# Vistas
# --------------------------------------------------------------------------
pestanas = st.tabs(
    ["Evolucion mensual", "Productos destacados", "Tabla de promedios", "Calidad de datos"]
)

with pestanas[0]:
    if serie_filtrada.empty:
        st.info("No hay movimientos con fecha en la seleccion actual.")
    else:
        grafico = (
            alt.Chart(serie_filtrada)
            .mark_bar()
            .encode(
                x=alt.X("periodo:N", title="Periodo", sort=None),
                y=alt.Y("sum(cantidad_salida):Q", title="Unidades salidas"),
                color=alt.Color(
                    "tipo_comprobante:N", title="Comprobante"
                ),
                tooltip=[
                    alt.Tooltip("periodo:N", title="Periodo"),
                    alt.Tooltip("tipo_comprobante:N", title="Comprobante"),
                    alt.Tooltip("cantidad_salida:Q", title="Unidades", format=","),
                ],
            )
            .properties(height=380)
        )
        st.altair_chart(grafico, width="stretch")

        por_tipo = metrics.totales_por_tipo(mov_filtrado)
        if not por_tipo.empty:
            izquierda, derecha = st.columns([3, 2])
            with izquierda:
                st.markdown("**Distribucion por tipo de comprobante**")
                st.altair_chart(
                    alt.Chart(por_tipo)
                    .mark_bar()
                    .encode(
                        y=alt.Y("tipo_comprobante:N", title="Comprobante", sort="-x"),
                        x=alt.X("cantidad_salida:Q", title="Unidades salidas"),
                        color=alt.Color(
                            "cantidad_salida:Q",
                            title="Unidades",
                            scale=alt.Scale(range=[COLOR_SALIDAS, COLOR_ENTRADAS]),
                        ),
                        tooltip=[
                            alt.Tooltip("tipo_comprobante:N", title="Comprobante"),
                            alt.Tooltip("cantidad_salida:Q", title="Unidades", format=","),
                            alt.Tooltip("n_salidas:Q", title="Lineas", format=","),
                        ],
                    )
                    .properties(height=260),
                    width="stretch",
                )
            with derecha:
                st.markdown("**Meses con mayor cantidad salida**")
                top_meses = por_periodo.nlargest(8, "cantidad_salida").copy()
                top_meses["etiqueta"] = (
                    top_meses["anio"].astype(int).astype(str)
                    + "-"
                    + top_meses["mes"].map(lambda m: f"{int(m):02d}")
                )
                st.altair_chart(
                    alt.Chart(top_meses)
                    .mark_bar()
                    .encode(
                        y=alt.Y("etiqueta:N", title="Periodo", sort="-x"),
                        x=alt.X("cantidad_salida:Q", title="Unidades"),
                        color=alt.value(COLOR_SALIDAS),
                        tooltip=[
                            alt.Tooltip("etiqueta:N", title="Periodo"),
                            alt.Tooltip("cantidad_salida:Q", title="Unidades", format=","),
                        ],
                    )
                    .properties(height=260),
                    width="stretch",
                )

        with st.expander("Ver datos mensuales"):
            st.dataframe(
                serie_filtrada,
                width="stretch",
                hide_index=True,
                column_config={
                    "periodo": "Periodo",
                    "tipo_comprobante": "Comprobante",
                    "cantidad_salida": st.column_config.NumberColumn(
                        "Unidades", format=","  # type: ignore[arg-type]
                    ),
                },
            )

with pestanas[1]:
    rankings = metrics.rankings(resumen_filtrado, n=15)
    etiquetas = {
        "salidas": "Mayor cantidad salida",
        "movimientos": "Mas movimientos registrados",
        "frecuencia": "Mas lineas de salida",
        "inventario": "Mayor saldo de valor en inventario",
    }
    if not rankings:
        st.info("No hay datos suficientes para armar los rankings.")
    else:
        for clave, titulo in etiquetas.items():
            if clave not in rankings:
                continue
            with st.container(border=True):
                st.markdown(f"**{titulo}**")
                st.altair_chart(
                    alt.Chart(rankings[clave])
                    .mark_bar()
                    .encode(
                        y=alt.Y("etiqueta:N", title=None, sort="-x"),
                        x=alt.X("valor:Q", title="Valor"),
                        color=alt.value(COLOR_SALIDAS),
                        tooltip=[
                            alt.Tooltip("codigo:N", title="Codigo"),
                            alt.Tooltip("nombre:N", title="Producto"),
                            alt.Tooltip("valor:Q", title="Valor", format=",.2f"),
                        ],
                    )
                    .properties(height=30 * len(rankings[clave]) + 60),
                    width="stretch",
                )

with pestanas[2]:
    st.caption(
        "Una fila por producto. Los promedios O, P y Q son los que se escriben en "
        "el archivo exportado; el desglose mensual se muestra solo como apoyo."
    )
    columnas_tabla = [
        "codigo",
        "nombre",
        "total_salidas",
        "n_movimientos",
        "n_meses",
        "prom_mensual",
        "n_semestres",
        "prom_semestral",
        "prom_anual",
        "saldo_cant",
        "saldo_valor",
    ]
    meses_tabla = [f"prom_{m.lower()}" for m in MESES_ES]
    mostrar_meses = st.toggle("Mostrar desglose mensual", value=False)

    vista = resumen_filtrado.copy()
    if not mostrar_meses:
        vista = vista.drop(columns=[c for c in meses_tabla if c in vista.columns])

    st.dataframe(
        vista[[c for c in columnas_tabla + (meses_tabla if mostrar_meses else []) if c in vista.columns]],
        width="stretch",
        hide_index=True,
        height=520,
        column_config={
            "codigo": "Codigo",
            "nombre": "Producto",
            "total_salidas": st.column_config.NumberColumn("Total salidas", format=","),
            "n_movimientos": st.column_config.NumberColumn("Movimientos", format=","),
            "n_meses": "Meses c/salida",
            "prom_mensual": st.column_config.NumberColumn("O · mensual", format="%.2f"),
            "n_semestres": "Semestres c/salida",
            "prom_semestral": st.column_config.NumberColumn("P · semestral", format="%.2f"),
            "prom_anual": st.column_config.NumberColumn("Q · anual", format="%.2f"),
            "saldo_cant": st.column_config.NumberColumn("Saldo cantidad", format="%.2f"),
            "saldo_valor": st.column_config.NumberColumn("Saldo valor", format="%.2f"),
            **{f"prom_{m.lower()}": m for m in MESES_ES},
        },
    )

    descarga_tabla = vista.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Descargar tabla en CSV",
        data=descarga_tabla,
        file_name="tabla_promedios.csv",
        mime="text/csv",
    )

with pestanas[3]:
    st.caption(
        "Controles sobre los datos de origen. No corrigen nada: señalan en que "
        "registros conviene revisar antes de tomar decisiones con los promedios."
    )
    for alerta in alertas:
        if alerta["cantidad"] == 0:
            continue
        severidad = alerta["severidad"]
        titulo = f"{alerta['titulo']} · {alerta['cantidad']:,} de {alerta['total']:,}"
        detalle = alerta["detalle"]
        with st.expander(titulo):
            st.caption(alerta["descripcion"])
            columnas = [c for c in detalle.columns if c in ("codigo", "nombre", "nombre_grupo")]
            referencia = "nombre_grupo" if "nombre_grupo" in detalle.columns else "nombre"
            st.caption(
                f"Afecta al {alerta['porcentaje']:.1f}% de los {alerta['total']:,} registros de control."
            )
            if not detalle.empty:
                st.dataframe(
                    detalle,
                    width="stretch",
                    hide_index=True,
                    height=260,
                    column_config={
                        columnas[0] if columnas else "codigo": "Codigo",
                        referencia: "Producto",
                    },
                )

    sin_problemas = sum(1 for a in alertas if a["cantidad"] == 0)
    if sin_problemas:
        st.success(f"{sin_problemas} de {len(alertas)} controles sin hallazgos.")


# --------------------------------------------------------------------------
# Verificacion y descarga
# --------------------------------------------------------------------------
st.divider()
verificacion = contexto["verificacion"]
if verificacion["disponible"]:
    con_datos = int(verificacion["resumen"]["difieren"].sum())
    total_celdas = int(verificacion["total"])
    if con_datos == 0:
        st.success(
            f"Verificacion superada: los {total_celdas:,} promedios calculados "
            "coinciden con las columnas O, P y Q que ya traia el archivo."
        )
    else:
        st.warning(
            f"{con_datos:,} de {total_celdas:,} promedios difieren de los que trae "
            "el archivo. Revisa el detalle antes de usar el resultado."
        )
    with st.expander("Detalle de la verificacion"):
        st.dataframe(
            verificacion["resumen"],
            width="stretch",
            hide_index=True,
        )

st.subheader("Descargar archivo con promedios")
st.caption(
    "Se conserva el archivo tal como lo subiste, con los mismos estilos, y se "
    "agregan las columnas O, P y Q con los promedios calculados."
)

if st.button("Preparar archivo", type="primary"):
    with st.spinner("Escribiendo los promedios..."):
        st.session_state["descarga"] = exporter.exportar_con_promedios(
            contenido, resultado, resumen
        )
        nombre_base = Path(contexto["nombre"]).stem
        st.session_state["nombre_descarga"] = f"{nombre_base}_con_promedios.xlsx"

if "descarga" in st.session_state:
    st.download_button(
        "Descargar Excel con promedios",
        data=st.session_state["descarga"],
        file_name=st.session_state.get("nombre_descarga", "reporte_con_promedios.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
