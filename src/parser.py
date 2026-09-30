"""Estructura de referencia del reporte de movimientos.

El archivo de los ERP (tipo "Movimientos con promedios") tiene una forma muy
particular que la app debe tolerar:

- La fila 1 va vacia (solo con formato).
- El encabezado **no** esta en la fila 1: aparece unas filas mas abajo
  (``Código producto`` ... ``Saldo valor``).
- Cada producto ocupa un **bloque**: una fila de grupo
  (``Código producto: <codigo> <nombre>`` en la columna A, con los totales del
  periodo en G..N) seguida de sus filas de detalle.
- Las columnas de promedios (``O``, ``P``, ``Q``) solo se rellenan en las filas
  de detalle y valen lo mismo para todas las filas de un producto.
- La columna de fecha es **texto** ``dd/mm/aaaa``, no un tipo fecha.
- Al final hay filas de pie tipo ``Procesado en: ...``.

Este modulo traduce ese formato a un DataFrame plano, una fila por movimiento,
sin perder el numero de fila original (necesario para exportar el Excel).
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Final

import pandas as pd

# Nombre canonico de cada columna del reporte, por posicion (indice base 0).
COLUMNAS: Final[tuple[str, ...]] = (
    "codigo",  # A  Codigo producto
    "nombre",  # B  Nombre producto
    "referencia",  # C  Referencia de fabrica
    "comprobante",  # D  Comprobante
    "secuencia",  # E  Secuencia
    "fecha",  # F  Fecha elaboracion
    "cant_inicial",  # G  Cantidad inicial
    "cant_entrada",  # H  Cantidad entrada
    "cant_salida",  # I  Cantidad salida
    "saldo_cant",  # J  Saldo cantidades
    "val_inicial",  # K  Valor inicial
    "val_entrada",  # L  Valor entrada
    "val_salida",  # M  Valor salida
    "saldo_valor",  # N  Saldo valor
)

# Indice base 0 de las columnas de promedios que calcula la app.
COL_PROM_MENSUAL: Final[int] = 14  # O
COL_PROM_SEMESTRAL: Final[int] = 15  # P
COL_PROM_ANUAL: Final[int] = 16  # Q

TOTAL_COLUMNAS: Final[int] = max(COL_PROM_ANUAL + 1, len(COLUMNAS))

# Las columnas de promedios llegan vacias en el archivo de entrada.
PROMEDIOS_VACIOS: Final[tuple[str, ...]] = (
    f"col_{COL_PROM_MENSUAL}",
    f"col_{COL_PROM_SEMESTRAL}",
    f"col_{COL_PROM_ANUAL}",
)

# Etiquetas que identifican la fila de encabezado dentro del archivo. Se
# comparan sin acentos y en minusculas, porque el ERP puede variar la tilde.
_ETIQUETAS_ENCABEZADO: Final[frozenset[str]] = frozenset(
    {
        "codigo producto",
        "nombre producto",
        "referencia de fabrica",
        "comprobante",
        "secuencia",
        "fecha elaboracion",
        "cantidad inicial",
        "cantidad entrada",
        "cantidad salida",
        "saldo cantidades",
        "valor inicial",
        "valor entrada",
        "valor salida",
        "saldo valor",
    }
)

# "Código producto: AC0004 ACEITE PAG150 (8oz) BVA USA" -> "AC0004"
_RE_GRUPO: Final[re.Pattern[str]] = re.compile(
    r"^\s*c[oó]digo\s+producto\s*:\s*(?P<codigo>\S+)\s*(?P<nombre>.*)$",
    re.IGNORECASE,
)
_RE_PIE: Final[re.Pattern[str]] = re.compile(r"^\s*procesado\s+en\b", re.IGNORECASE)


def _sin_acentos(texto: str) -> str:
    """Quita acentos y pasa a minusculas, para comparar etiquetas."""
    substitutions = str.maketrans("áéíóúüñÁÉÍÓÚÜÑ", "aeiouunAEIOUUN")
    return texto.translate(substitutions).strip().lower()


def _normalizar(texto: object) -> str:
    """Pasa un valor de celda a texto limpio, o cadena vacia si esta vacio."""
    if texto is None:
        return ""
    if isinstance(texto, float) and pd.isna(texto):
        return ""
    if pd.isna(texto):
        return ""
    return str(texto).strip()


def _col_alfabetica(indice: int) -> str:
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA'."""
    letras = ""
    indice += 1
    while indice:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def _es_fila_encabezado(fila: list[object]) -> bool:
    """Una fila es encabezado si coincide con varias etiquetas conocidas.

    Se exige un minimo de 5 coincidencias para no confundirse con una fila de
    datos que por casualidad traiga textos parecidos.
    """
    textos = [_sin_acentos(_normalizar(v)) for v in fila]
    aciertos = sum(1 for t in textos if t in _ETIQUETAS_ENCABEZADO)
    return aciertos >= 5


def _es_fila_pie(fila: list[object]) -> bool:
    return bool(_RE_PIE.match(_normalizar(fila[0])))


def _a_numero(valor: object) -> float:
    """Convierte celda a float; vacio, texto o NaN pasan a ``nan``."""
    if valor is None:
        return float("nan")
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)
    texto = _normalizar(valor)
    if not texto:
        return float("nan")
    try:
        return float(texto.replace(" ", "").replace(" ", ""))
    except ValueError:
        return float("nan")


def _a_fecha(valor: object) -> pd.Timestamp | pd.NaT:
    """Interpreta la columna de fecha, que viene como texto ``dd/mm/aaaa``."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return pd.NaT
    if isinstance(valor, pd.Timestamp):
        return valor
    texto = _normalizar(valor)
    if not texto:
        return pd.NaT
    try:
        return pd.to_datetime(texto, format="%d/%m/%Y", errors="raise")
    except (ValueError, TypeError):
        # Respaldo por si el ERP cambiara el formato en algun reporte futuro.
        try:
            return pd.to_datetime(texto, dayfirst=True, errors="raise")
        except (ValueError, TypeError):
            return pd.NaT


@dataclass(frozen=True)
class ResultadoParseo:
    """Salida del parser: movimientos, bloques de producto y avisos de lectura."""

    movimientos: pd.DataFrame
    productos: pd.DataFrame
    fila_encabezado: int
    encabezado: list[str]
    avisos: list[str]
    formato: str


def _leer_bruto(ruta_o_bytes: str | bytes | io.BytesIO) -> pd.DataFrame:
    """Lee el .xlsx completo sin interpretar cabeceras (header=None).

    Acepta una ruta o los bytes del archivo subido. Los bytes se envuelven en un
    BytesIO porque pandas no los reconoce como archivo por si solos.
    """
    fuente: str | io.BytesIO
    if isinstance(ruta_o_bytes, (bytes, bytearray)):
        fuente = io.BytesIO(ruta_o_bytes)
    elif isinstance(ruta_o_bytes, io.BytesIO):
        fuente = ruta_o_bytes
    else:
        fuente = str(ruta_o_bytes)
    return pd.read_excel(
        fuente,
        header=None,
        dtype=object,
        engine="openpyxl",
    )


def calcular_semestre(serie: pd.Series) -> pd.Series:
    """Numero de semestre (1 o 2) de cada fecha; NaN donde no hay fecha."""
    mes = serie.dt.month
    return ((mes - 1) // 6 + 1).where(serie.notna())


def parsear_reporte(ruta_o_bytes: str | bytes) -> ResultadoParseo:
    """Convierte el .xlsx del ERP en DataFrames listos para calcular.

    Parametros
    ----------
    ruta_o_bytes:
        Ruta al archivo, o los bytes ya leidos (util desde un upload).

    Devuelve
    -------
    ResultadoParseo
        ``movimientos`` con una fila por linea de detalle (incluye las de
        cantidad salida en cero, que luego se filtran al calcular), ``productos``
        con una fila por fila de grupo y los avisos de la lectura.
    """
    avisos: list[str] = []
    bruto = _leer_bruto(ruta_o_bytes)

    # Recorta columnas sobrantes: el reporte usa A..Q como maximo.
    if bruto.shape[1] > TOTAL_COLUMNAS:
        bruto = bruto.iloc[:, :TOTAL_COLUMNAS]
        avisos.append(
            f"El archivo tiene {bruto.shape[1] + 0} columnas visibles; se toman las "
            f"primeras {TOTAL_COLUMNAS} (A..{_col_alfabetica(TOTAL_COLUMNAS - 1)})."
        )

    filas = bruto.to_numpy(dtype=object).tolist()

    # Localiza la fila de encabezado en lugar de asumir que es la primera.
    fila_encabezado = -1
    for i, fila in enumerate(filas):
        if _es_fila_encabezado(fila):
            fila_encabezado = i
            break

    if fila_encabezado < 0:
        raise ValueError(
            "No se encontro la fila de encabezado. Se busca una fila que contenga "
            "'Código producto'. Revisa que el archivo tenga la estructura del "
            "reporte de movimientos."
        )

    if fila_encabezado > 0:
        avisos.append(
            f"El encabezado no estaba en la primera fila: se encontro en la fila "
            f"{fila_encabezado + 1}."
        )

    encabezado = [_normalizar(v) for v in filas[fila_encabezado][: len(COLUMNAS)]]
    cuerpo = filas[fila_encabezado + 1 :]

    # Descarta filas totalmente vacias y el pie de pagina del ERP.
    registros: list[dict[str, object]] = []
    bloques: list[dict[str, object]] = []
    codigo_actual = ""
    nombre_actual = ""
    referencia_actual = ""
    n_pie = 0
    n_vacias = 0
    fila_grupo_actual = 0

    for offset, fila in enumerate(cuerpo):
        n_fila = fila_encabezado + 2 + offset  # 1-based, como lo ve Excel
        a = _normalizar(fila[0] if len(fila) > 0 else None)

        if not a and all(_normalizar(v) == "" for v in fila):
            n_vacias += 1
            continue
        if _es_fila_pie(fila):
            n_pie += 1
            continue

        def celda(idx: int) -> object:
            return fila[idx] if len(fila) > idx else None

        coincidencia = _RE_GRUPO.match(a)
        if coincidencia:
            codigo_actual = coincidencia.group("codigo")
            nombre_actual = _normalizar(coincidencia.group("nombre"))
            referencia_actual = ""
            fila_grupo_actual = n_fila
            bloques.append(
                {
                    "fila_excel": n_fila,
                    "codigo": codigo_actual,
                    "nombre_grupo": nombre_actual,
                    "cant_inicial": _a_numero(celda(6)),
                    "cant_entrada": _a_numero(celda(7)),
                    "cant_salida": _a_numero(celda(8)),
                    "saldo_cant": _a_numero(celda(9)),
                    "val_inicial": _a_numero(celda(10)),
                    "val_entrada": _a_numero(celda(11)),
                    "val_salida": _a_numero(celda(12)),
                    "saldo_valor": _a_numero(celda(13)),
                    "n_movimientos": 0,
                }
            )
            continue

        # Cualquier otra fila con datos es un movimiento del producto actual.
        if not codigo_actual:
            avisos.append(
                f"La fila {n_fila} tiene movimientos pero no pertenece a ningun "
                "bloque de producto; se ignora."
            )
            continue

        referencia_actual = _normalizar(celda(2)) or referencia_actual
        nombre_det = _normalizar(celda(1)) or nombre_actual
        fecha = _a_fecha(celda(5))
        registro: dict[str, object] = {
            "fila_excel": n_fila,
            "fila_grupo": fila_grupo_actual,
            "codigo": codigo_actual,
            "nombre": nombre_det,
            "referencia": referencia_actual,
            "comprobante": _normalizar(celda(3)),
            "secuencia": _normalizar(celda(4)),
            "fecha": fecha,
        }
        for nombre_col, idx in (
            ("cant_inicial", 6),
            ("cant_entrada", 7),
            ("cant_salida", 8),
            ("saldo_cant", 9),
            ("val_inicial", 10),
            ("val_entrada", 11),
            ("val_salida", 12),
            ("saldo_valor", 13),
        ):
            registro[nombre_col] = _a_numero(celda(idx))
        for idx in (COL_PROM_MENSUAL, COL_PROM_SEMESTRAL, COL_PROM_ANUAL):
            registro[f"col_{idx}"] = _a_numero(celda(idx))
        registros.append(registro)
        if bloques:
            bloques[-1]["n_movimientos"] += 1

    if n_pie:
        avisos.append(f"Se ignoraron {n_pie} fila(s) de pie ('Procesado en...').")
    if n_vacias:
        avisos.append(f"Se ignoraron {n_vacias} fila(s) vacia(s).")

    movimientos = pd.DataFrame.from_records(registros)
    if movimientos.empty:
        raise ValueError(
            "Se leyeron los encabezados pero no se encontro ninguna fila de "
            "movimiento. Revisa el contenido del archivo."
        )

    movimientos["tipo_comprobante"] = (
        movimientos["comprobante"].str.split("-").str[0].str.strip()
    )
    movimientos["semestre"] = calcular_semestre(movimientos["fecha"])
    movimientos["anio"] = movimientos["fecha"].dt.year
    movimientos["mes"] = movimientos["fecha"].dt.month

    productos = pd.DataFrame.from_records(bloques)

    return ResultadoParseo(
        movimientos=movimientos,
        productos=productos,
        fila_encabezado=fila_encabezado + 1,
        encabezado=encabezado,
        avisos=avisos,
        formato="bloques",
    )
