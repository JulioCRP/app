"""Inspeccion de referencia del XLSX: confirma la estructura real de filas y columnas."""

from pathlib import Path

from openpyxl import load_workbook

ARCHIVO = Path(__file__).resolve().parent.parent / "Movimientos_con_promedios.xlsx"


def main() -> None:
    wb = load_workbook(ARCHIVO, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    print(f"hoja            : {ws.title}")
    print(f"dimensiones     : {ws.max_row} filas x {ws.max_column} columnas")

    print("\n--- primeras 6 filas (celda a celda, None = vacio) ---")
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), start=1):
        celdas = [f"{chr(65 + j)}={v!r}" for j, v in enumerate(row) if v is not None]
        print(f"fila {i:>3}: " + (" | ".join(celdas) if celdas else "(vacia)"))

    print("\n--- ultimas 3 filas ---")
    for i, row in enumerate(
        ws.iter_rows(min_row=ws.max_row - 2, max_row=ws.max_row, values_only=True), start=ws.max_row - 2
    ):
        celdas = [f"{chr(65 + j)}={v!r}" for j, v in enumerate(row) if v is not None]
        print(f"fila {i:>3}: " + (" | ".join(celdas) if celdas else "(vacia)"))

    print("\n--- tipos de dato de la columna F (fecha elaboracion) ---")
    tipos: dict[str, int] = {}
    muestras: list[str] = []
    for row in ws.iter_rows(min_row=3, max_row=min(ws.max_row, 500), values_only=True):
        v = row[5] if len(row) > 5 else None
        if v is None:
            continue
        tipos[type(v).__name__] = tipos.get(type(v).__name__, 0) + 1
        if len(muestras) < 5:
            muestras.append(f"{v!r} ({type(v).__name__})")
    print("  conteo por tipo:", tipos)
    print("  muestras       :", ", ".join(muestras))

    print("\n--- conteo de filas de grupo vs detalle (desde la fila 3) ---")
    grupos = detalle = pie = 0
    primera_grupo = None
    for i, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
        a = row[0] if row else None
        if a is None:
            continue
        texto = str(a)
        if texto.lower().startswith("procesado en"):
            pie += 1
        elif texto.startswith("Código producto:") or texto.startswith("Codigo producto:"):
            grupos += 1
            if primera_grupo is None:
                primera_grupo = i
        else:
            detalle += 1
    print(f"  filas de grupo  : {grupos}   (primera en la fila {primera_grupo})")
    print(f"  filas de detalle: {detalle}")
    print(f"  filas de pie    : {pie}")
    wb.close()


if __name__ == "__main__":
    main()
