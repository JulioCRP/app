<#
.SYNOPSIS
    Instalador de la aplicacion de promedios de inventario.

.DESCRIPTION
    Crea el entorno virtual, instala las dependencias y deja la aplicacion
    lista para usar. Pensado para Windows con PowerShell, donde 'make' no viene
    instalado.

.PARAMETER Accion
    Accion a ejecutar: install (por defecto), run, test, check, clean, reset.

.EXAMPLE
    .\instalar.ps1
    Instala todo.

.EXAMPLE
    .\instalar.ps1 run
    Levanta la aplicacion.

.EXAMPLE
    .\instalar.ps1 test
    Corre todas las pruebas.
#>

param(
    [ValidateSet('install', 'run', 'test', 'check', 'clean', 'reset', 'todo')]
    [string]$Accion = 'install',

    [int]$Puerto = 8501
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

$Venv = Join-Path $PSScriptRoot '.venv'
$Python = Join-Path $Venv 'Scripts\python.exe'

function Escribir-Paso {
    param([string]$Mensaje)
    Write-Host ''
    Write-Host "==> $Mensaje" -ForegroundColor Cyan
}

function Comprobar-Python {
    $comando = $null
    foreach ($nombre in @('python', 'py')) {
        $encontrado = Get-Command $nombre -ErrorAction SilentlyContinue
        if ($encontrado) { $comando = $nombre; break }
    }
    if (-not $comando) {
        Write-Host 'No se encontro Python en el sistema.' -ForegroundColor Red
        Write-Host 'Descargalo de https://www.python.org/downloads/ y marca' -ForegroundColor Yellow
        Write-Host '"Add Python to PATH" durante la instalacion.' -ForegroundColor Yellow
        exit 1
    }
    $version = & $comando --version 2>&1
    Write-Host "  Python detectado: $version ($comando)"
    if ([version]($version -replace '^Python ', '') -lt [version]'3.11') {
        Write-Host 'Se requiere Python 3.11 o superior.' -ForegroundColor Red
        exit 1
    }
}

function Crear-Entorno {
    if (Test-Path $Python) {
        Write-Host '  El entorno virtual ya existe.'
        return
    }
    Escribir-Paso 'Creando el entorno virtual'
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $comando = 'py'
    } else {
        $comando = 'python'
    }
    & $comando -m venv $Venv
    if (-not (Test-Path $Python)) {
        Write-Host 'No se pudo crear el entorno virtual.' -ForegroundColor Red
        exit 1
    }
    Write-Host '  Entorno creado en .venv'
}

function Instalar-Dependencias {
    Escribir-Paso 'Instalando dependencias'
    & $Python -m pip install --upgrade pip --quiet
    & $Python -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) {
        Write-Host 'La instalacion de dependencias fallo.' -ForegroundColor Red
        exit 1
    }
}

switch ($Accion) {

    'install' {
        Escribir-Paso 'Comprobando Python'
        Comprobar-Python
        Crear-Entorno
        Instalar-Dependencias
        Write-Host ''
        Write-Host 'Instalacion completada.' -ForegroundColor Green
        Write-Host 'Levanta la aplicacion con:  .\instalar.ps1 run' -ForegroundColor Gray
    }

    'run' {
        if (-not (Test-Path $Python)) {
            Write-Host 'El entorno no existe. Ejecuta primero: .\instalar.ps1' -ForegroundColor Red
            exit 1
        }
        Write-Host "Iniciando en http://localhost:$Puerto  (Ctrl+C para detener)" -ForegroundColor Cyan
        & $Python -m streamlit run app.py --server.port $Puerto
    }

    'test' {
        if (-not (Test-Path $Python)) {
            Write-Host 'El entorno no existe. Ejecuta primero: .\instalar.ps1' -ForegroundColor Red
            exit 1
        }
        $fallos = 0
        foreach ($prueba in @('test_metrics', 'test_exporter', 'test_solo_a_n', 'test_app')) {
            Escribir-Paso "Ejecutando $prueba"
            & $Python "tests\$prueba.py"
            if ($LASTEXITCODE -ne 0) { $fallos++ }
        }
        if ($fallos -gt 0) {
            Write-Host "$fallos prueba(s) fallaron." -ForegroundColor Red
            exit 1
        }
        Write-Host ''
        Write-Host 'Todas las pruebas pasaron.' -ForegroundColor Green
    }

    'check' {
        if (-not (Test-Path $Python)) {
            Write-Host 'El entorno no existe. Ejecuta primero: .\instalar.ps1' -ForegroundColor Red
            exit 1
        }
        & $Python -c "import sys, streamlit, pandas, numpy, openpyxl, altair; print('Python    ', sys.version.split()[0]); print('streamlit ', streamlit.__version__); print('pandas    ', pandas.__version__); print('numpy     ', numpy.__version__); print('openpyxl  ', openpyxl.__version__); print('altair    ', altair.__version__)"
    }

    'clean' {
        Escribir-Paso 'Limpiando'
        Get-ChildItem -Path $PSScriptRoot -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -notlike "*\.venv\*" } |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
        $temporal = Join-Path $PSScriptRoot 'salida_prueba.xlsx'
        if (Test-Path $temporal) { Remove-Item -LiteralPath $temporal -Force }
        Write-Host 'Limpieza terminada.'
    }

    'reset' {
        Escribir-Paso 'Eliminando el entorno virtual'
        if (Test-Path $Venv) { Remove-Item -LiteralPath $Venv -Recurse -Force }
        & $PSCommandPath install
    }

    'todo' {
        & $PSCommandPath install
        & $PSCommandPath test
    }
}
