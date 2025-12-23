#!/bin/bash
# Script de instalación rápida - Linux/Mac
# Instala dependencias y prepara el entorno

set -e  # Salir si hay error

echo "========================================================================"
echo "  INSTALACION - REPORTE ECONOMIA NARANJA"
echo "========================================================================"
echo ""

# Verificar Python
echo "[1/4] Verificando Python..."
if ! command -v python3 &> /dev/null; then
    echo "  ✗ Python3 no está instalado"
    echo ""
    echo "  Por favor instala Python 3.10 o superior"
    exit 1
fi
python3 --version
echo "  ✓ Python instalado"
echo ""

# Crear entorno virtual
echo "[2/4] Creando entorno virtual..."
if [ -d ".venv" ]; then
    echo "  - Entorno virtual ya existe, omitiendo creación"
else
    python3 -m venv .venv
    echo "  ✓ Entorno virtual creado"
fi
echo ""

# Activar entorno virtual
echo "[3/4] Activando entorno virtual..."
source .venv/bin/activate
echo "  ✓ Entorno virtual activado"
echo ""

# Instalar dependencias
echo "[4/4] Instalando dependencias..."
python -m pip install --upgrade pip
pip install -r requirements.txt
echo "  ✓ Dependencias instaladas"
echo ""

echo "========================================================================"
echo "  INSTALACION COMPLETADA"
echo "========================================================================"
echo ""
echo "  Para ejecutar el programa:"
echo "    1. Activa el entorno virtual: source .venv/bin/activate"
echo "    2. Ejecuta: python main.py"
echo ""
echo "  O directamente: .venv/bin/python main.py"
echo ""
echo "========================================================================"
