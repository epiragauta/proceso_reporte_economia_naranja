"""
Configuración centralizada para el proceso de generación del
Reporte Consolidado de Economía Naranja

Este módulo contiene todas las rutas, nombres de archivos y configuraciones
utilizadas en el proceso de generación de reportes mensuales.
"""

import os
from pathlib import Path

# ============================================
# DIRECTORIOS BASE
# ============================================

# Directorio raíz del proyecto
DIR_BASE = Path(r'C:\ws\sena\data')

# Directorio del proceso
DIR_PROCESO = DIR_BASE / 'PROCESO_REPORTE_ECONOMIA_NARANJA'

# Directorios de componentes existentes
DIR_PE04 = DIR_BASE / 'PE-04'
DIR_METAS = DIR_BASE / 'metas'
DIR_APRENDICES = DIR_BASE / 'aprendices'
DIR_REPORTE_ECONOMIA_NARANJA = DIR_BASE / 'REPORTE_ECONOMIA_NARANJA'

# ============================================
# MAPEO DE MESES
# ============================================

MESES = {
    'ENERO': {'corto': 'Ene', 'numero': 1},
    'FEBRERO': {'corto': 'Feb', 'numero': 2},
    'MARZO': {'corto': 'Mar', 'numero': 3},
    'ABRIL': {'corto': 'Abr', 'numero': 4},
    'MAYO': {'corto': 'May', 'numero': 5},
    'JUNIO': {'corto': 'Jun', 'numero': 6},
    'JULIO': {'corto': 'Jul', 'numero': 7},
    'AGOSTO': {'corto': 'Ago', 'numero': 8},
    'SEPTIEMBRE': {'corto': 'Sep', 'numero': 9},
    'OCTUBRE': {'corto': 'Oct', 'numero': 10},
    'NOVIEMBRE': {'corto': 'Nov', 'numero': 11},
    'DICIEMBRE': {'corto': 'Dic', 'numero': 12}
}

# ============================================
# AÑO DE TRABAJO
# ============================================

ANIO_TRABAJO = 2025

# ============================================
# FUNCIONES DE CONFIGURACIÓN POR MES
# ============================================

def obtener_config_mes(mes_nombre):
    """
    Obtiene la configuración completa para un mes específico

    Args:
        mes_nombre (str): Nombre del mes en mayúsculas (ej: 'SEPTIEMBRE')

    Returns:
        dict: Diccionario con toda la configuración del mes
    """
    mes_nombre = mes_nombre.upper()

    if mes_nombre not in MESES:
        raise ValueError(f"Mes inválido: {mes_nombre}. Debe ser uno de: {', '.join(MESES.keys())}")

    mes_info = MESES[mes_nombre]
    mes_corto = mes_info['corto']
    mes_numero = mes_info['numero']

    # Directorios del mes
    dir_mes = DIR_PROCESO / mes_nombre
    dir_datos_intermedios = dir_mes / 'datos_intermedios'
    dir_datos_finales = dir_mes / 'datos_finales'

    # Directorio fuente de archivos originales
    dir_fuente = DIR_BASE / str(ANIO_TRABAJO) / f'{mes_numero:02d}-{mes_nombre.capitalize()}'

    config = {
        # Información del mes
        'mes_nombre': mes_nombre,
        'mes_corto': mes_corto,
        'mes_numero': mes_numero,
        'anio': ANIO_TRABAJO,

        # Directorios
        'dir_mes': dir_mes,
        'dir_datos_intermedios': dir_datos_intermedios,
        'dir_datos_finales': dir_datos_finales,
        'dir_fuente': dir_fuente,

        # ARCHIVOS DE ENTRADA (fuentes originales)
        'archivos_entrada': {
            'pe04_formacion': dir_fuente / f'PE-04_FORMACION NACIONAL {mes_nombre.upper()} {ANIO_TRABAJO}.xlsb',
            'avance_cupos': dir_fuente / f'PRIMER AVANCE CUPOS DE FORMACION {mes_nombre.upper()} {ANIO_TRABAJO}.xlsb',
            'avance_aprendices': dir_fuente / f'PRIMER AVANCE EN APRENDICES {mes_nombre.upper()} {ANIO_TRABAJO}.xlsb',
            'metas_sena': DIR_BASE / '2025' / '09-Septiembre' / 'Metas SENA 2025 V5 26092025_CLEAN.xlsx'  # Único archivo anual
        },

        # ARCHIVOS INTERMEDIOS
        'archivos_intermedios': {
            'bd_formacion': dir_datos_intermedios / f'sena_formacion_{mes_nombre.lower()}.db',
            'bd_metas': dir_datos_intermedios / 'metas_sena_2025.db',
            'cupos_disponibles_xlsx': dir_datos_intermedios / 'cupos_disponibles_por_regional_2025.xlsx',
            'cupos_disponibles_csv': dir_datos_intermedios / 'cupos_disponibles_por_regional_2025.csv',
            'reporte_aprendices': dir_datos_intermedios / f'SENA Mensual Nacional {mes_corto} {ANIO_TRABAJO}.xlsx'
        },

        # ARCHIVOS FINALES
        'archivos_finales': {
            'reporte_consolidado': dir_datos_finales / f'Reporte Consolidado Economía Naranja {mes_corto} {ANIO_TRABAJO}.xlsx'
        },

        # SCRIPTS A EJECUTAR (rutas absolutas)
        'scripts': {
            'importar_pe04': DIR_PE04 / 'importar_pe_04_mes.py',
            'crear_tabla_economia_naranja': DIR_REPORTE_ECONOMIA_NARANJA / 'crear_tabla_economia_naranja.sql',
            'normalizar_metas': DIR_METAS / 'normalizar_metas_sena.py',
            'cruce_metas_avance': DIR_METAS / 'cruce_metas_avance_final.py',
            'generar_reporte_aprendices': DIR_APRENDICES / 'generar_reporte_mensual.py',
            'generar_reporte_consolidado': DIR_REPORTE_ECONOMIA_NARANJA / 'generar_reporte_consolidado.py'
        },

        # NOMBRES DE TABLAS EN BD
        'tablas_bd': {
            'economia_naranja': f'ECONOMIA_NARANJA_{mes_nombre.upper()}_{ANIO_TRABAJO}'
        }
    }

    return config


def validar_archivos_entrada(config):
    """
    Valida que todos los archivos de entrada existan

    Args:
        config (dict): Configuración del mes

    Returns:
        tuple: (bool, list) - (Todo OK, lista de archivos faltantes)
    """
    archivos_faltantes = []

    for nombre, ruta in config['archivos_entrada'].items():
        if not ruta.exists():
            archivos_faltantes.append(f"{nombre}: {ruta}")

    return len(archivos_faltantes) == 0, archivos_faltantes


def crear_directorios_mes(config):
    """
    Crea los directorios necesarios para un mes

    Args:
        config (dict): Configuración del mes
    """
    config['dir_mes'].mkdir(parents=True, exist_ok=True)
    config['dir_datos_intermedios'].mkdir(parents=True, exist_ok=True)
    config['dir_datos_finales'].mkdir(parents=True, exist_ok=True)

    print(f"✓ Directorios creados para {config['mes_nombre']}")
    print(f"  - {config['dir_mes']}")
    print(f"  - {config['dir_datos_intermedios']}")
    print(f"  - {config['dir_datos_finales']}")


# ============================================
# FUNCIÓN DE AYUDA
# ============================================

def imprimir_config(config):
    """Imprime la configuración de forma legible"""
    print(f"\n{'='*60}")
    print(f"CONFIGURACIÓN - {config['mes_nombre']} {config['anio']}")
    print(f"{'='*60}\n")

    print(f"📅 Mes: {config['mes_nombre']} ({config['mes_corto']})")
    print(f"📅 Año: {config['anio']}")
    print(f"\n📁 Directorio del mes: {config['dir_mes']}")
    print(f"📁 Datos intermedios: {config['dir_datos_intermedios']}")
    print(f"📁 Datos finales: {config['dir_datos_finales']}")

    print(f"\n📥 Archivos de entrada:")
    for nombre, ruta in config['archivos_entrada'].items():
        existe = "✓" if ruta.exists() else "✗"
        print(f"  {existe} {nombre}: {ruta.name}")

    print(f"\n📊 Archivos intermedios:")
    for nombre, ruta in config['archivos_intermedios'].items():
        print(f"  • {nombre}: {ruta.name}")

    print(f"\n📄 Archivos finales:")
    for nombre, ruta in config['archivos_finales'].items():
        print(f"  • {nombre}: {ruta.name}")

    print(f"\n{'='*60}\n")


# ============================================
# EJEMPLO DE USO
# ============================================

if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("Uso: python configuracion.py <MES>")
        print(f"Meses válidos: {', '.join(MESES.keys())}")
        sys.exit(1)

    mes = sys.argv[1].upper()

    try:
        config = obtener_config_mes(mes)
        imprimir_config(config)

        # Validar archivos
        print("Validando archivos de entrada...")
        todo_ok, faltantes = validar_archivos_entrada(config)

        if todo_ok:
            print("✓ Todos los archivos de entrada están disponibles")
        else:
            print("✗ Archivos faltantes:")
            for archivo in faltantes:
                print(f"  - {archivo}")

    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)
