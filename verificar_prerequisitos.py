"""
Verificador de Prerequisitos - Reporte Economía Naranja

Este script valida que todos los archivos de entrada necesarios
estén disponibles antes de iniciar el proceso de generación del reporte.
"""

import sys
from pathlib import Path
from configuracion import obtener_config_mes, MESES

def verificar_python():
    """Verifica la versión de Python"""
    print("\n1. Verificando versión de Python...")
    version = sys.version_info
    if version.major >= 3 and version.minor >= 10:
        print(f"   ✓ Python {version.major}.{version.minor}.{version.micro}")
        return True
    else:
        print(f"   ✗ Python {version.major}.{version.minor}.{version.micro}")
        print(f"   Se requiere Python 3.10 o superior")
        return False


def verificar_dependencias():
    """Verifica que las dependencias de Python estén instaladas"""
    print("\n2. Verificando dependencias de Python...")

    dependencias = {
        'pandas': 'Manipulación de datos',
        'openpyxl': 'Lectura/escritura de Excel XLSX',
        'pyxlsb': 'Lectura de Excel XLSB',
        'sqlite3': 'Base de datos SQLite'
    }

    todas_ok = True

    for modulo, descripcion in dependencias.items():
        try:
            if modulo == 'sqlite3':
                import sqlite3
            else:
                __import__(modulo)
            print(f"   ✓ {modulo:12s} - {descripcion}")
        except ImportError:
            print(f"   ✗ {modulo:12s} - {descripcion} (FALTANTE)")
            todas_ok = False

    if not todas_ok:
        print("\n   Para instalar las dependencias faltantes:")
        print("   pip install pandas openpyxl pyxlsb")

    return todas_ok


def verificar_archivos_entrada(config):
    """Verifica que todos los archivos de entrada existan"""
    print(f"\n3. Verificando archivos de entrada para {config['mes_nombre']}...")

    archivos = config['archivos_entrada']
    todos_existen = True

    for nombre, ruta in archivos.items():
        if ruta.exists():
            tamano_mb = ruta.stat().st_size / (1024 * 1024)
            print(f"   ✓ {nombre:20s} ({tamano_mb:.1f} MB)")
        else:
            print(f"   ✗ {nombre:20s} - NO ENCONTRADO")
            print(f"      Ruta esperada: {ruta}")
            todos_existen = False

    return todos_existen


def verificar_scripts():
    """Verifica que todos los scripts necesarios existan"""
    print("\n4. Verificando scripts del sistema...")

    # Obtener una configuración de ejemplo para verificar scripts
    config = obtener_config_mes('SEPTIEMBRE')
    scripts = config['scripts']

    todos_existen = True

    for nombre, ruta in scripts.items():
        if ruta.exists():
            print(f"   ✓ {nombre}")
        else:
            print(f"   ✗ {nombre} - NO ENCONTRADO")
            print(f"      Ruta esperada: {ruta}")
            todos_existen = False

    return todos_existen


def verificar_directorios_base():
    """Verifica que los directorios base existan"""
    print("\n5. Verificando estructura de directorios...")

    from configuracion import DIR_BASE, DIR_PE04, DIR_METAS, DIR_APRENDICES, DIR_REPORTE_ECONOMIA_NARANJA

    directorios = {
        'DIR_BASE': DIR_BASE,
        'DIR_PE04': DIR_PE04,
        'DIR_METAS': DIR_METAS,
        'DIR_APRENDICES': DIR_APRENDICES,
        'DIR_REPORTE_ECONOMIA_NARANJA': DIR_REPORTE_ECONOMIA_NARANJA
    }

    todos_existen = True

    for nombre, ruta in directorios.items():
        if ruta.exists():
            print(f"   ✓ {nombre:30s} - {ruta}")
        else:
            print(f"   ✗ {nombre:30s} - NO ENCONTRADO")
            todos_existen = False

    return todos_existen


def verificar_espacio_disco(config):
    """Verifica que haya suficiente espacio en disco"""
    print("\n6. Verificando espacio en disco...")

    import shutil

    try:
        stat = shutil.disk_usage(config['dir_mes'])
        libre_gb = stat.free / (1024**3)

        if libre_gb >= 5:
            print(f"   ✓ Espacio disponible: {libre_gb:.1f} GB")
            return True
        else:
            print(f"   ⚠ Espacio disponible: {libre_gb:.1f} GB")
            print(f"   Se recomienda al menos 5 GB libres")
            return False
    except Exception as e:
        print(f"   ⚠ No se pudo verificar espacio en disco: {e}")
        return True  # No bloquear por esto


def main():
    print("="*70)
    print(" VERIFICADOR DE PREREQUISITOS - REPORTE ECONOMÍA NARANJA")
    print("="*70)

    # Validar argumentos
    if len(sys.argv) < 2:
        print("\nUso: python verificar_prerequisitos.py <MES>")
        print(f"\nMeses válidos:")
        for mes in MESES.keys():
            print(f"  - {mes}")
        sys.exit(1)

    mes_nombre = sys.argv[1].upper()

    # Validar mes
    if mes_nombre not in MESES:
        print(f"\n✗ Mes inválido: {mes_nombre}")
        print(f"\nMeses válidos: {', '.join(MESES.keys())}")
        sys.exit(1)

    # Obtener configuración
    try:
        config = obtener_config_mes(mes_nombre)
    except Exception as e:
        print(f"\n✗ Error al obtener configuración: {e}")
        sys.exit(1)

    # Ejecutar verificaciones
    resultados = []

    resultados.append(("Python", verificar_python()))
    resultados.append(("Dependencias", verificar_dependencias()))
    resultados.append(("Archivos de entrada", verificar_archivos_entrada(config)))
    resultados.append(("Scripts del sistema", verificar_scripts()))
    resultados.append(("Directorios base", verificar_directorios_base()))
    resultados.append(("Espacio en disco", verificar_espacio_disco(config)))

    # Resumen
    print("\n" + "="*70)
    print(" RESUMEN DE VERIFICACIÓN")
    print("="*70)

    todas_ok = True
    for nombre, resultado in resultados:
        simbolo = "✓" if resultado else "✗"
        print(f"   {simbolo} {nombre}")
        if not resultado:
            todas_ok = False

    print("="*70)

    if todas_ok:
        print("\n✓ TODOS LOS PREREQUISITOS ESTÁN LISTOS")
        print(f"\nPuedes ejecutar:")
        print(f"   python generar_reporte_completo.py {mes_nombre}")
        sys.exit(0)
    else:
        print("\n✗ ALGUNOS PREREQUISITOS FALTAN")
        print("\nPor favor, corrige los problemas antes de continuar.")
        sys.exit(1)


if __name__ == '__main__':
    main()
