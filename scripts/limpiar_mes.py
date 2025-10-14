"""
Script de Limpieza - Reporte Economía Naranja

Este script elimina todos los archivos generados para un mes específico,
útil para volver a ejecutar el proceso desde cero.

ADVERTENCIA: Esta operación NO se puede deshacer.
"""

import sys
import shutil
from pathlib import Path
from configuracion import obtener_config_mes, MESES

def confirmar_eliminacion(mes_nombre):
    """Pide confirmación al usuario antes de eliminar"""
    print("\n" + "!"*70)
    print(" ADVERTENCIA: ESTA OPERACIÓN NO SE PUEDE DESHACER")
    print("!"*70)
    print(f"\nEstás a punto de eliminar TODOS los archivos generados para {mes_nombre}")
    print("\nEsto incluye:")
    print("  • Bases de datos generadas")
    print("  • Archivos intermedios")
    print("  • Reportes finales")
    print("  • Copias de archivos de entrada")
    print("\nLos archivos originales en el directorio fuente NO se eliminarán.")

    respuesta = input("\n¿Deseas continuar? (escribe 'SI' para confirmar): ")

    return respuesta.strip().upper() == 'SI'


def eliminar_directorio(directorio, descripcion):
    """Elimina un directorio y todo su contenido"""
    if not directorio.exists():
        print(f"  ⊘ {descripcion}: No existe (omitiendo)")
        return

    try:
        # Contar archivos
        archivos = list(directorio.rglob('*'))
        num_archivos = len([f for f in archivos if f.is_file()])

        # Eliminar
        shutil.rmtree(directorio)
        print(f"  ✓ {descripcion}: Eliminado ({num_archivos} archivos)")

    except Exception as e:
        print(f"  ✗ {descripcion}: Error al eliminar - {e}")


def main():
    print("="*70)
    print(" LIMPIEZA DE ARCHIVOS GENERADOS")
    print("="*70)

    # Validar argumentos
    if len(sys.argv) < 2:
        print("\nUso: python limpiar_mes.py <MES>")
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

    # Mostrar resumen de lo que se va a eliminar
    print(f"\nDirectorio del mes: {config['dir_mes']}")

    # Pedir confirmación
    if not confirmar_eliminacion(mes_nombre):
        print("\n✗ Operación cancelada por el usuario")
        sys.exit(0)

    # Eliminar directorios
    print(f"\nEliminando archivos de {mes_nombre}...\n")

    eliminar_directorio(config['dir_datos_intermedios'], "Datos intermedios")
    eliminar_directorio(config['dir_datos_finales'], "Datos finales")

    # Opcionalmente eliminar el directorio del mes completo
    if config['dir_mes'].exists():
        contenido = list(config['dir_mes'].iterdir())
        if len(contenido) == 0:
            config['dir_mes'].rmdir()
            print(f"  ✓ Directorio del mes: Eliminado (estaba vacío)")
        else:
            print(f"  ⊘ Directorio del mes: Mantiene {len(contenido)} elementos")

    print("\n" + "="*70)
    print(" LIMPIEZA COMPLETADA")
    print("="*70)
    print(f"\nPuedes ejecutar el proceso nuevamente con:")
    print(f"   python generar_reporte_completo.py {mes_nombre}")
    print()


if __name__ == '__main__':
    main()
