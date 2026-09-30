"""Exportacion del reporte con los promedios calculados.

La estrategia es abrir el archivo original con openpyxl en modo normal y
escribir unicamente las columnas O, P y Q. Asi se conservan intactos los estilos,
los anchos de columna, los filtros y el pie del reporte: el resultado es
indistinguible del archivo original mas los promedios.

No se reconstruye el archivo de cero, porque eso perderia el formato.
"""

from __future__ import annotations

import io
import re
from copy import copy

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.utils import column_index_from_string, get_column_letter

from src.parser import ResultadoParseo

# Formato con el que el reporte escribe los promedios.
FORMATO_PROMEDIO: str = "#,##0.00"

# Indices de columna de openpyxl (base 1) para O, P y Q. Se escriben con letras
# y no con los indices de src.parser porque aquellos son base 0.
COL_O: int = column_index_from_string("O")
COL_P: int = column_index_from_string("P")
COL_Q: int = column_index_from_string("Q")

ENCABEZADOS_PROMEDIO: dict[int, str] = {
    COL_O: "Promedio mensual",
    COL_P: "Promedio semestral",
    COL_Q: "Promedio anual",
}

COLUMNAS_RESULTADO: tuple[int, ...] = (COL_O, COL_P, COL_Q)

# Columna contigua usada como plantilla de estilo (saldo cantidades, columna J).
COLUMNAS_PLANTILLA: int = column_index_from_string("J")


class ErrorExportacion(RuntimeError):
    """Se lanza cuando el libro no se puede preparar para escribir O/P/Q."""


def _es_fila_grupo(valor: object) -> bool:
    return bool(re.match(r"^\s*c[oó]digo\s+producto\s*:", str(valor or ""), re.IGNORECASE))


def _encontrar_fila_plantilla(hoja: openpyxl.worksheet.worksheet.Worksheet) -> int | None:
    """Primera fila de detalle de la hoja, para tomarla como plantilla."""
    limite = min(hoja.max_row, 300)
    for fila in range(1, limite + 1):
        if not hoja.cell(fila, 1).value or _es_fila_grupo(hoja.cell(fila, 1).value):
            continue
        comprobante = hoja.cell(fila, 4).value
        if comprobante and re.match(r"^[A-Za-z]{1,3}-?\s?\d", str(comprobante).strip()):
            return fila
    return None


def _es_nulo(valor: object) -> bool:
    return valor is None or (isinstance(valor, (float, np.floating)) and np.isnan(valor))


def exportar_con_promedios(
    contenido: bytes,
    resultado: ResultadoParseo,
    resumen: pd.DataFrame,
    *,
    hoja_nombre: str | None = None,
) -> bytes:
    """Devuelve el .xlsx original con O, P y Q completadas.

    Parametros
    ----------
    contenido:
        Bytes del archivo subido. El archivo del disco no se toca nunca.
    resultado:
        Salida de :func:`src.parser.parsear_reporte`; de aqui sale el numero de
        fila real de Excel de cada movimiento.
    resumen:
        Salida de :func:`src.metrics.resumen_productos`; de aqui salen los tres
        promedios por codigo de producto.
    """
    promedios = resumen.set_index("codigo")[["prom_mensual", "prom_semestral", "prom_anual"]]

    libro = openpyxl.load_workbook(io.BytesIO(contenido))
    try:
        hoja = libro[hoja_nombre] if hoja_nombre else libro.worksheets[0]

        plantilla = _encontrar_fila_plantilla(hoja)
        if plantilla is None:
            raise ErrorExportacion(
                "No se encontro ninguna fila de detalle en la hoja para usar como "
                "plantilla de formato. Revisa que el archivo tenga la estructura "
                "esperada: encabezado con 'Comprobante' y debajo filas de "
                "movimiento que no empiecen por 'Codigo producto:'."
            )
        estilo = copy(hoja.cell(plantilla, COLUMNAS_PLANTILLA)._style)

        movimientos = resultado.movimientos
        codigos = movimientos["codigo"].to_numpy()
        filas = movimientos["fila_excel"].to_numpy()

        # Un solo acceso al indice por producto, en vez de uno por fila.
        valores = promedios.reindex(pd.Index(codigos)).to_numpy(dtype=float)

        escritas = 0
        for indice_fila, fila in zip(valores, filas):
            fila_excel = int(fila)
            for indice_columna, valor in zip(COLUMNAS_RESULTADO, indice_fila):
                celda = hoja.cell(fila_excel, indice_columna)
                celda._style = copy(estilo)
                celda.number_format = FORMATO_PROMEDIO
                if np.isnan(valor):
                    # Sin salidas positivas el promedio va vacio, no en cero.
                    celda.value = None
                else:
                    celda.value = float(valor)
                    escritas += 1

        # Encabezados de las columnas nuevas, solo si estan vacios.
        if resultado.fila_encabezado:
            origen = hoja.cell(resultado.fila_encabezado, 14)
            for indice_columna, etiqueta in ENCABEZADOS_PROMEDIO.items():
                celda = hoja.cell(resultado.fila_encabezado, indice_columna)
                if celda.value in (None, ""):
                    celda.value = etiqueta
                    celda.font = copy(origen.font)
                    celda.fill = copy(origen.fill)
                    celda.border = copy(origen.border)
                    celda.alignment = copy(origen.alignment)
            for indice_columna in ENCABEZADOS_PROMEDIO:
                letra = get_column_letter(indice_columna)
                actual = hoja.column_dimensions[letra].width or 0
                hoja.column_dimensions[letra].width = max(actual, 18)

        buffer = io.BytesIO()
        libro.save(buffer)
        return buffer.getvalue()
    finally:
        libro.close()
