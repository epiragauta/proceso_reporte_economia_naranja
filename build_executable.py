#!/usr/bin/env python3
"""
Script de Compilación - Reporte Economía Naranja

Este script compila el proyecto en un ejecutable standalone usando PyInstaller.
El ejecutable generado será un archivo .exe (Windows) o binario (Linux/Mac)
que puede distribuirse y ejecutarse sin necesidad de tener Python instalado.

Uso:
    python build_executable.py
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path


def limpiar_directorios_build():
    """Limpia directorios de builds anteriores"""
    print("\n→ Limpiando builds anteriores...")

    dirs_limpiar = ['build', 'dist', '__pycache__']
    archivos_limpiar = ['*.spec']

    for dir_name in dirs_limpiar:
        if Path(dir_name).exists():
            shutil.rmtree(dir_name)
            print(f"  ✓ Eliminado: {dir_name}/")

    # Limpiar archivos .spec
    for spec_file in Path('.').glob('*.spec'):
        spec_file.unlink()
        print(f"  ✓ Eliminado: {spec_file}")


def verificar_dependencias():
    """Verifica que PyInstaller esté instalado"""
    print("\n→ Verificando dependencias...")

    try:
        import PyInstaller
        print(f"  ✓ PyInstaller {PyInstaller.__version__} instalado")
        return True
    except ImportError:
        print("  ✗ PyInstaller no está instalado")
        print("\n  Para instalar:")
        print("    pip install pyinstaller")
        return False


def compilar_ejecutable():
    """Compila el proyecto con PyInstaller"""
    print("\n→ Compilando ejecutable...")
    print("  Esto puede tomar varios minutos...")

    # Nombre del ejecutable
    nombre_exe = "reporte-economia-naranja"

    # Archivos adicionales a incluir
    data_files = [
        ('scripts', 'scripts'),  # Incluir todos los scripts
        ('CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx', '.'),  # Catálogo en raíz
    ]

    # Archivos ocultos de imports (algunos módulos los necesitan)
    hidden_imports = [
        'pandas',
        'openpyxl',
        'pyxlsb',
        'sqlite3',
        'tqdm',
    ]

    # Construir comando de PyInstaller
    comando = [
        'pyinstaller',
        '--name', nombre_exe,
        '--onefile',  # Un solo archivo ejecutable
        '--console',  # Aplicación de consola (no GUI)
        '--clean',    # Limpiar cache antes de compilar
        # '--icon', 'icon.ico',  # Descomentar si tienes un ícono
    ]

    # Agregar archivos de datos
    for origen, destino in data_files:
        if Path(origen).exists():
            comando.extend(['--add-data', f'{origen}{os.pathsep}{destino}'])
        else:
            print(f"  ⚠ Advertencia: No se encontró {origen}, será omitido")

    # Agregar imports ocultos
    for imp in hidden_imports:
        comando.extend(['--hidden-import', imp])

    # Archivo principal
    comando.append('main.py')

    print(f"\n  Comando: {' '.join(comando)}\n")

    try:
        resultado = subprocess.run(
            comando,
            check=True,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )

        print("  ✓ Compilación exitosa")
        return True

    except subprocess.CalledProcessError as e:
        print("  ✗ Error en la compilación")
        print("\n  STDOUT:")
        print(e.stdout)
        print("\n  STDERR:")
        print(e.stderr)
        return False


def verificar_ejecutable():
    """Verifica que el ejecutable se haya creado correctamente"""
    print("\n→ Verificando ejecutable generado...")

    dist_dir = Path('dist')
    if not dist_dir.exists():
        print("  ✗ Directorio dist/ no existe")
        return False

    # Buscar ejecutable
    ejecutables = list(dist_dir.glob('reporte-economia-naranja*'))

    if not ejecutables:
        print("  ✗ No se encontró el ejecutable en dist/")
        return False

    ejecutable = ejecutables[0]
    tamano_mb = ejecutable.stat().st_size / (1024 * 1024)

    print(f"  ✓ Ejecutable generado: {ejecutable.name}")
    print(f"  ✓ Tamaño: {tamano_mb:.1f} MB")
    print(f"  ✓ Ubicación: {ejecutable.absolute()}")

    return True


def crear_carpeta_distribucion():
    """Crea una carpeta lista para distribuir"""
    print("\n→ Creando carpeta de distribución...")

    dist_dir = Path('dist')
    release_dir = Path('release')

    # Crear directorio release
    if release_dir.exists():
        shutil.rmtree(release_dir)
    release_dir.mkdir()

    # Copiar ejecutable
    ejecutables = list(dist_dir.glob('reporte-economia-naranja*'))
    if ejecutables:
        ejecutable = ejecutables[0]
        shutil.copy2(ejecutable, release_dir / ejecutable.name)
        print(f"  ✓ Copiado: {ejecutable.name}")

    # Copiar archivos esenciales
    archivos_copiar = [
        'README.md',
        'CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx',
        'Nueva_Hoja_SENA.xlsx',
    ]

    for archivo in archivos_copiar:
        if Path(archivo).exists():
            shutil.copy2(archivo, release_dir / archivo)
            print(f"  ✓ Copiado: {archivo}")

    # Crear estructura de directorios necesaria
    (release_dir / 'MESES').mkdir(exist_ok=True)
    print("  ✓ Creado: MESES/")

    print(f"\n  ✓ Carpeta de distribución lista: {release_dir.absolute()}")
    return True


def main():
    print("=" * 70)
    print(" COMPILADOR - REPORTE ECONOMÍA NARANJA")
    print("=" * 70)

    # Paso 1: Verificar dependencias
    if not verificar_dependencias():
        sys.exit(1)

    # Paso 2: Limpiar builds anteriores
    limpiar_directorios_build()

    # Paso 3: Compilar
    if not compilar_ejecutable():
        print("\n✗ La compilación falló")
        sys.exit(1)

    # Paso 4: Verificar ejecutable
    if not verificar_ejecutable():
        print("\n✗ No se pudo verificar el ejecutable")
        sys.exit(1)

    # Paso 5: Crear carpeta de distribución
    crear_carpeta_distribucion()

    print("\n" + "=" * 70)
    print(" COMPILACIÓN COMPLETADA EXITOSAMENTE")
    print("=" * 70)
    print("\n✓ El ejecutable está listo en la carpeta 'release/'")
    print("\n📦 Para distribuir:")
    print("  1. Comprime la carpeta 'release/' en un archivo ZIP")
    print("  2. Envía el ZIP a los usuarios")
    print("  3. Los usuarios solo necesitan descomprimir y ejecutar")
    print("\n" + "=" * 70)


if __name__ == "__main__":
    main()
