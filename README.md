# Promedios de inventario

Aplicacion local en Streamlit que lee un archivo `.xlsx` de movimientos de
inventario, calcula el promedio de salidas de cada producto y escribe los
resultados en las columnas **O**, **P** y **Q** del mismo archivo.

El archivo original **no se modifica**: la app trabaja sobre una copia en
memoria y te deja descargar el resultado. Nada sale de tu equipo.

---

## Como funciona

Sube un `.xlsx` que tenga las columnas A a N. La app calcula tres promedios por
producto y te devuelve el archivo con esas columnas agregadas.

| Columna | Promedio | Formula |
|---|---|---|
| **O** | Mensual | Total de cantidad salida ÷ meses con al menos una salida |
| **P** | Semestral | Total de cantidad salida ÷ semestres con al menos una salida |
| **Q** | Anual | Total de cantidad salida ÷ anios con al menos una salida |

Detalles importantes de la regla:

- Solo cuentan las filas con **cantidad salida mayor que cero**. Una salida en
  cero no suma al total ni cuenta como periodo.
- El divisor es la cantidad de **periodos con al menos una salida**, no de
  movimientos. Un mes con entradas pero sin salidas no cuenta.
- El redondeo es el de Excel (`REDONDEAR`, half away from zero), no el
  redondeo bancario que trae Python por defecto.
- Un producto **sin ninguna salida positiva deja O, P y Q vacias**, no en
  cero. Mostrar 0 seria un promedio inventado.
- El valor del promedio se repite en todas las filas de detalle del producto;
  las filas de grupo quedan vacias.

---

## Requisitos

- **Python 3.11 o superior** (probado en 3.11 y 3.13)
- Alrededor de 300 MB de disco para el entorno virtual
- 2 GB de RAM libres: leer un archivo de ~40.000 filas toma unos segundos

No necesitas Excel instalado ni conexion a internet (salvo la primera
instalacion de dependencias).

Las versiones de las librerias se fijan como **rangos minimos**, no como pines
exactos, para que la aplicacion tambien funcione en equipos con versiones algo
distintas de Python.

---

## Instalacion paso a paso

### 1. Verificar Python

```bash
python --version
```

Debe mostrar 3.11 o superior. Si no lo tienes, instalalo desde
<https://www.python.org/downloads/> marcando **"Add Python to PATH"** durante la
instalacion.

En Windows, si `python` no responde prueba con `py --version`.

### 2. Entrar a la carpeta del proyecto

```bash
cd ruta/a/este/proyecto
```

### 3. Crear el entorno virtual

**Windows (PowerShell o cmd):**

```bat
python -m venv .venv
```

**Linux o macOS:**

```bash
python3 -m venv .venv
```

### 4. Activar el entorno virtual

**Windows (PowerShell):**

```powershell
.\.venv\Scripts\Activate.ps1
```

**Windows (cmd):**

```bat
.venv\Scripts\activate.bat
```

**Linux o macOS:**

```bash
source .venv/bin/activate
```

Deberias ver `(.venv)` al inicio de la linea. Si PowerShell bloquea la
ejecucion por politica de seguridad, ejecuta una vez esto:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### 5. Instalar las dependencias

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### 6. Ejecutar la aplicacion

```bash
python -m streamlit run app.py
```

Se abre en <http://localhost:8501>. Si no se abre solo, copia esa direccion en
el navegador.

Para cerrarla, presiona `Ctrl+C` en la terminal.

---

## Atajos con Make

Si tienes `make` instalado, todos los pasos anteriores se reducen a:

```bash
make install    # crea el entorno e instala dependencias
make run        # levanta la aplicacion
make test       # corre todas las pruebas
```

En Windows `make` no viene incluido. Se instala con cualquiera de estos:

```powershell
winget install GnuWin32.Make
# o
choco install make
# o
scoop install make
```

Si prefieres no instalar `make`, el script `instalar.ps1` hace lo mismo en
PowerShell:

```powershell
.\instalar.ps1
.\instalar.ps1 run
.\instalar.ps1 test
```

---

## Como se usa

1. Abre <http://localhost:8501>.
2. Arrastra tu `.xlsx` al area de carga.
3. Espera unos segundos mientras se procesa. La primera carga es la mas lenta;
   despues Streamlit guarda el resultado en cache.
4. Revisa las pestanas:
   - **Evolucion mensual**: salidas por mes y por tipo de comprobante.
   - **Productos destacados**: rankings por salidas, movimientos e inventario.
   - **Tabla de promedios**: una fila por producto, con el desglose mensual
     opcional.
   - **Calidad de datos**: controles sobre los datos de origen.
5. Presiona **Preparar archivo** y descarga el Excel con los promedios en O, P y Q.

La barra lateral filtra por busqueda de producto, anio, tipo de comprobante, y
permite quedarse solo con los productos que sí tienen promedio.

---

## Controles de calidad

La pestana **Calidad de datos** revisa el archivo de origen. No corrige nada:
señala en que registros conviene mirar antes de tomar decisiones.

| Control | Que detecta |
|---|---|
| Productos sin movimientos | Filas de grupo sin ningun movimiento |
| Salidas con cantidad pero sin valor | Se descontaron unidades sin registrar su valor |
| Salidas con valor pero sin cantidad | Se descontó valor sin descontar unidades |
| Saldo de valor negativo | Productos con saldo de inventario negativo |
| Descuadre de cantidad | El saldo no cumple Inicial + Entrada - Salida |
| Descuadre de valor | El saldo de valor no cumple la ecuacion esperada |
| Movimientos sin fecha o comprobante | No se pueden agrupar por periodo |

---

## Estructura del proyecto

```
.
├── app.py                      Interfaz de la aplicacion
├── requirements.txt            Dependencias
├── Makefile                    Atajos de instalacion y pruebas
├── instalar.ps1                Alternativa a make en Windows
├── .streamlit/
│   └── config.toml             Configuracion local de Streamlit
├── src/
│   ├── parser.py               Lee el .xlsx y detecta la estructura
│   ├── metrics.py              Calcula los promedios y las alertas
│   └── exporter.py             Escribe O, P y Q en el archivo
├── tests/
│   ├── test_metrics.py         Fórmulas y coincidencia con el reporte
│   ├── test_exporter.py        Conservación de formato y columnas
│   ├── test_solo_a_n.py        Caso real: archivo sin O/P/Q
│   └── test_app.py             Pruebas de la interfaz
└── tools/                      Scripts de análisis y limpieza
```

---

## Pruebas

```bash
python tests/test_metrics.py     # fórmulas de los promedios
python tests/test_exporter.py    # A-N intactas, O/P/Q correctas
python tests/test_solo_a_n.py    # archivo de solo A-N, el caso real
python tests/test_app.py         # la interfaz
```

O todas juntas:

```bash
make test
```

`test_metrics.py` es la prueba de referencia: compara los promedios calculados
contra las columnas O, P y Q del archivo de ejemplo y exige coincidencia
total. Si el reporte de referencia cambia y las formulas ya no coinciden, esa
prueba falla y avisa.

---

## Formato esperado del archivo

El archivo tiene una hoja con esta estructura, repetida por producto:

- Fila 1: vacia
- Fila 2: encabezado (`Código producto`, `Nombre producto`,
  `Referencia de fabrica`, `Comprobante`, `Secuencia`, `Fecha elaboración`,
  `Cantidad inicial`, `Cantidad entrada`, `Cantidad salida`, `Saldo cantidades`,
  `Valor inicial`, `Valor entrada`, `Valor salida`, `Saldo valor`)
- Por cada producto, una fila que empieza con `Código producto: ...` seguida
  de sus movimientos
- Al final, una fila `Procesado en: ...`

El parser no depende de la fila exacta del encabezado: lo busca. Tampoco
importa si las columnas O, P y Q ya vienen llenas; si vienen, la app verifica
que sus calculos coincidan y te lo dice.

---

## Problemas frecuentes

**PowerShell dice que la ejecucion de scripts esta deshabilitada**

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

**"Se requiere Python 3.11 o superior"**

Python esta instalado pero es mas antiguo. Descarga la version actual desde
<https://www.python.org/downloads/>, marcando **"Add Python to PATH"**.

**"Address already in use" al levantar la app**

Hay otra instancia corriendo. Cierra la terminal anterior o usa otro puerto:

```bash
python -m streamlit run app.py --server.port 8502
```

**El archivo es muy grande y tarda mucho**

Sube el limite de tamaño en `.streamlit/config.toml`:

```toml
[server]
maxUploadSize = 500
```

**La app no abre el navegador**

Abre <http://localhost:8501> manualmente.

**"No se encontro ninguna fila de detalle"**

El archivo no tiene la estructura esperada. Revisa que el encabezado este en su
propia fila y que debajo haya filas cuyo comprobante empiece con el tipo de
documento (por ejemplo `FV-12345`).
