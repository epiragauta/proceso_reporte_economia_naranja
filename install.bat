@echo off
REM Script de instalación rápida - Windows
REM Instala dependencias y prepara el entorno

echo ========================================================================
echo   INSTALACION - REPORTE ECONOMIA NARANJA
echo ========================================================================
echo.

REM Verificar Python
echo [1/4] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo   X Python no esta instalado
    echo.
    echo   Por favor instala Python 3.10 o superior desde:
    echo   https://www.python.org/downloads/
    pause
    exit /b 1
)
python --version
echo   OK Python instalado
echo.

REM Crear entorno virtual
echo [2/4] Creando entorno virtual...
if exist .venv (
    echo   - Entorno virtual ya existe, omitiendo creacion
) else (
    python -m venv .venv
    if errorlevel 1 (
        echo   X Error al crear entorno virtual
        pause
        exit /b 1
    )
    echo   OK Entorno virtual creado
)
echo.

REM Activar entorno virtual
echo [3/4] Activando entorno virtual...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo   X Error al activar entorno virtual
    pause
    exit /b 1
)
echo   OK Entorno virtual activado
echo.

REM Instalar dependencias
echo [4/4] Instalando dependencias...
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo   X Error al instalar dependencias
    pause
    exit /b 1
)
echo   OK Dependencias instaladas
echo.

echo ========================================================================
echo   INSTALACION COMPLETADA
echo ========================================================================
echo.
echo   Para ejecutar el programa:
echo     1. Activa el entorno virtual: .venv\Scripts\activate
echo     2. Ejecuta: python main.py
echo.
echo   O directamente: .venv\Scripts\python main.py
echo.
echo ========================================================================
pause
