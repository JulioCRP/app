# Promedios de inventario - atajos de instalacion
#
# Uso basico:
#   make install    crea el entorno virtual e instala dependencias
#   make run        levanta la aplicacion
#   make test       corre todas las pruebas
#   make check      verifica la instalacion
#   make clean      borra caches
#   make todo       install + test
#
# En Windows make no viene incluido. Instalalo con:
#   winget install GnuWin32.Make
#   choco install make
#   scoop install make
# Si prefieres no instalarlo, usa instalar.ps1 en PowerShell.

# Windows y Unix usan el entorno virtual en carpetas distintas.
ifeq ($(OS),Windows_NT)
    PYBIN := .venv/Scripts/python.exe
    ACTIVAR := .venv\Scripts\activate
else
    PYBIN := .venv/bin/python
    ACTIVAR := source .venv/bin/activate
endif

PYTHON ?= python
PORT ?= 8501

.DEFAULT_GOAL := help
.PHONY: help venv verificar-python install run test test-metrics test-exporter test-solo-a-n test-app check clean reset todo

## Muestra esta ayuda
help:
	@echo Promedios de inventario
	@echo.
	@echo   make install    Crea el entorno virtual e instala dependencias
	@echo   make run        Levanta la aplicacion en http://localhost:$(PORT)
	@echo   make test       Corre todas las pruebas
	@echo   make check      Verifica que todo este instalado
	@echo   make clean      Borra caches
	@echo   make todo       install + test
	@echo   make reset      Borra el entorno y lo recrea desde cero

## Crea el entorno virtual
venv:
	@if exist "$(PYBIN)" ( \
		echo El entorno virtual ya existe. Usa make reset para rehacerlo. \
	) else ( \
		$(PYTHON) -m venv .venv && echo Entorno virtual creado en .venv \
	)

## Comprueba que la version de Python sea suficiente
verificar-python:
	@$(PYTHON) -c "import sys; v=sys.version_info; \
	print('Python detectado:', '.'.join(map(str, v[:3]))); \
	sys.exit(0) if v >= (3, 11) else sys.exit('ERROR: se requiere Python 3.11 o superior. Instala Python desde https://www.python.org/downloads/ y marca Add Python to PATH.')"

## Crea el entorno e instala las dependencias
install: verificar-python venv
	@echo Actualizando pip...
	@$(PYBIN) -m pip install --upgrade pip --quiet
	@echo Instalando dependencias...
	@$(PYBIN) -m pip install -r requirements.txt
	@echo.
	@echo Listo. Levanta la app con:  make run

## Levanta la aplicacion
run:
	@echo Iniciando en http://localhost:$(PORT)  (Ctrl+C para detener)
	@$(PYBIN) -m streamlit run app.py --server.port $(PORT)

## Corre todas las pruebas
test:
	@$(PYBIN) tests/test_metrics.py
	@$(PYBIN) tests/test_exporter.py
	@$(PYBIN) tests/test_solo_a_n.py
	@$(PYBIN) tests/test_app.py
	@echo.
	@echo Todas las pruebas pasaron.

test-metrics:
	@$(PYBIN) tests/test_metrics.py

test-exporter:
	@$(PYBIN) tests/test_exporter.py

test-solo-a-n:
	@$(PYBIN) tests/test_solo_a_n.py

test-app:
	@$(PYBIN) tests/test_app.py

## Verifica que el entorno este completo
check:
	@$(PYBIN) -c "import sys, streamlit, pandas, numpy, openpyxl, altair; print('Python    ', sys.version.split()[0]); print('streamlit ', streamlit.__version__); print('pandas    ', pandas.__version__); print('numpy     ', numpy.__version__); print('openpyxl  ', openpyxl.__version__); print('altair    ', altair.__version__)"

## Borra caches y archivos temporales
clean:
	@$(PYBIN) tools/limpiar.py

## Borra el entorno y lo recrea desde cero
reset:
	@echo Eliminando el entorno virtual...
	@if exist .venv rmdir /S /Q .venv
	@if exist .venv rm -rf .venv
	@$(MAKE) install

## install + test
todo: install test
