#!/usr/bin/env python3
"""
Reporte Economía Naranja SENA - Punto de Entrada Principal

Este es el programa principal para generar reportes consolidados
de Economía Naranja del SENA.

Uso:
    python main.py                    # Menú interactivo
    python main.py SEPTIEMBRE         # Generar reporte directamente
    python main.py --verificar SEPTIEMBRE  # Solo verificar prerequisitos
"""

import sys
import os
from pathlib import Path

# Agregar directorio de scripts al path
SCRIPT_DIR = Path(__file__).resolve().parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from configuracion import obtener_config_mes, MESES
from verificar_prerequisitos import main as verificar_prerequisitos_main
from generar_reporte_completo import main as generar_reporte_main


def limpiar_pantalla():
    """Limpia la pantalla de la consola"""
    os.system('cls' if os.name == 'nt' else 'clear')


def mostrar_banner():
    """Muestra el banner del programa"""
    print("=" * 70)
    print(" GENERADOR DE REPORTES - ECONOMÍA NARANJA SENA".center(70))
    print("=" * 70)
    print()


def mostrar_menu_principal():
    """Muestra el menú principal y retorna la opción seleccionada"""
    print("\nMENÚ PRINCIPAL:")
    print("-" * 70)
    print("  1. Generar Reporte Completo")
    print("  2. Verificar Prerequisitos")
    print("  3. Ver Configuración de un Mes")
    print("  4. Listar Meses Disponibles")
    print("  5. Salir")
    print("-" * 70)

    while True:
        try:
            opcion = input("\nSeleccione una opción (1-5): ").strip()
            if opcion in ['1', '2', '3', '4', '5']:
                return opcion
            else:
                print("⚠ Opción inválida. Por favor ingrese un número del 1 al 5.")
        except (KeyboardInterrupt, EOFError):
            print("\n\nPrograma interrumpido por el usuario.")
            sys.exit(0)


def seleccionar_mes():
    """Permite al usuario seleccionar un mes"""
    print("\nMESES DISPONIBLES:")
    print("-" * 70)

    meses_lista = list(MESES.keys())
    for i, mes in enumerate(meses_lista, 1):
        print(f"  {i:2d}. {mes}")

    print("-" * 70)

    while True:
        try:
            seleccion = input(f"\nSeleccione un mes (1-{len(meses_lista)}) o escriba el nombre: ").strip().upper()

            # Si es un número
            if seleccion.isdigit():
                idx = int(seleccion) - 1
                if 0 <= idx < len(meses_lista):
                    return meses_lista[idx]
                else:
                    print(f"⚠ Número inválido. Debe estar entre 1 y {len(meses_lista)}.")
            # Si es un nombre de mes
            elif seleccion in MESES:
                return seleccion
            else:
                print("⚠ Mes inválido. Intente de nuevo.")

        except (KeyboardInterrupt, EOFError):
            print("\n\nOperación cancelada.")
            return None


def opcion_generar_reporte():
    """Opción 1: Generar reporte completo"""
    limpiar_pantalla()
    mostrar_banner()
    print("GENERAR REPORTE COMPLETO")
    print("=" * 70)

    mes = seleccionar_mes()
    if mes is None:
        return

    print(f"\n→ Generando reporte para: {mes}")
    print("-" * 70)

    confirmacion = input("\n¿Desea continuar? (S/N): ").strip().upper()
    if confirmacion != 'S':
        print("Operación cancelada.")
        return

    # Llamar al script principal de generación
    sys.argv = ['generar_reporte_completo.py', mes]
    try:
        generar_reporte_main()
    except SystemExit as e:
        if e.code != 0:
            print(f"\n✗ El proceso finalizó con errores (código: {e.code})")
            input("\nPresione Enter para continuar...")


def opcion_verificar_prerequisitos():
    """Opción 2: Verificar prerequisitos"""
    limpiar_pantalla()
    mostrar_banner()
    print("VERIFICAR PREREQUISITOS")
    print("=" * 70)

    mes = seleccionar_mes()
    if mes is None:
        return

    print(f"\n→ Verificando prerequisitos para: {mes}")
    print("-" * 70)

    # Llamar al script de verificación
    sys.argv = ['verificar_prerequisitos.py', mes]
    try:
        verificar_prerequisitos_main()
    except SystemExit:
        pass

    input("\nPresione Enter para continuar...")


def opcion_ver_configuracion():
    """Opción 3: Ver configuración de un mes"""
    limpiar_pantalla()
    mostrar_banner()
    print("VER CONFIGURACIÓN DE MES")
    print("=" * 70)

    mes = seleccionar_mes()
    if mes is None:
        return

    try:
        from configuracion import imprimir_config, validar_archivos_entrada

        config = obtener_config_mes(mes)
        imprimir_config(config)

        # Validar archivos
        print("\nValidando archivos de entrada...")
        todo_ok, faltantes = validar_archivos_entrada(config)

        if todo_ok:
            print("✓ Todos los archivos de entrada están disponibles")
        else:
            print("\n✗ Archivos faltantes:")
            for archivo in faltantes:
                print(f"  - {archivo}")

    except Exception as e:
        print(f"\n✗ Error al obtener configuración: {e}")

    input("\nPresione Enter para continuar...")


def opcion_listar_meses():
    """Opción 4: Listar meses disponibles"""
    limpiar_pantalla()
    mostrar_banner()
    print("MESES DISPONIBLES")
    print("=" * 70)

    for mes, info in MESES.items():
        print(f"  {info['numero']:2d}. {mes:12s} ({info['corto']})")

    input("\nPresione Enter para continuar...")


def modo_interactivo():
    """Modo interactivo con menú"""
    while True:
        limpiar_pantalla()
        mostrar_banner()

        opcion = mostrar_menu_principal()

        if opcion == '1':
            opcion_generar_reporte()
        elif opcion == '2':
            opcion_verificar_prerequisitos()
        elif opcion == '3':
            opcion_ver_configuracion()
        elif opcion == '4':
            opcion_listar_meses()
        elif opcion == '5':
            print("\n¡Hasta luego!")
            sys.exit(0)


def modo_linea_comandos():
    """Modo de línea de comandos"""
    if len(sys.argv) < 2:
        print("Uso: reporte-naranja <MES> [--verificar]")
        print("\nEjemplos:")
        print("  reporte-naranja SEPTIEMBRE")
        print("  reporte-naranja --verificar OCTUBRE")
        sys.exit(1)

    # Parsear argumentos
    verificar_solo = False
    mes = None

    for arg in sys.argv[1:]:
        if arg == '--verificar':
            verificar_solo = True
        elif arg.upper() in MESES:
            mes = arg.upper()

    if mes is None:
        print(f"✗ Error: Debe especificar un mes válido")
        print(f"\nMeses válidos: {', '.join(MESES.keys())}")
        sys.exit(1)

    if verificar_solo:
        sys.argv = ['verificar_prerequisitos.py', mes]
        verificar_prerequisitos_main()
    else:
        sys.argv = ['generar_reporte_completo.py', mes]
        generar_reporte_main()


def main():
    """Función principal"""
    # Si no hay argumentos, modo interactivo
    if len(sys.argv) == 1:
        try:
            modo_interactivo()
        except KeyboardInterrupt:
            print("\n\nPrograma interrumpido por el usuario.")
            sys.exit(0)
    else:
        # Si hay argumentos, modo línea de comandos
        modo_linea_comandos()


if __name__ == "__main__":
    main()
