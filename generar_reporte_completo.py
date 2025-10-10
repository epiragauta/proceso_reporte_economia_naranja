"""
Script Maestro - Generación Completa del Reporte de Economía Naranja

Este script orquesta todo el proceso de generación del reporte consolidado,
desde la importación de datos hasta la generación del reporte final.

Uso:
    python generar_reporte_completo.py <MES>

Ejemplo:
    python generar_reporte_completo.py SEPTIEMBRE
"""

import sys
import os
import shutil
import subprocess
import sqlite3
from pathlib import Path
from datetime import datetime
from configuracion import obtener_config_mes, crear_directorios_mes, MESES

# ============================================
# FUNCIONES AUXILIARES
# ============================================

def log_paso(numero, total, mensaje):
    """Imprime un mensaje de log con formato consistente"""
    print(f"\n{'='*70}")
    print(f" PASO {numero}/{total}: {mensaje}")
    print(f"{'='*70}")


def ejecutar_comando(comando, descripcion, check=True):
    """
    Ejecuta un comando del sistema y maneja errores

    Args:
        comando (list): Comando a ejecutar
        descripcion (str): Descripción del comando
        check (bool): Si True, lanza excepción en caso de error

    Returns:
        bool: True si exitoso, False si falló
    """
    print(f"\n→ {descripcion}")
    print(f"  Comando: {' '.join(str(c) for c in comando)}")

    try:
        result = subprocess.run(
            comando,
            check=check,
            capture_output=True,
            text=True
        )

        if result.returncode == 0:
            print(f"✓ {descripcion} - EXITOSO")
            if result.stdout:
                # Mostrar solo las últimas 10 líneas de output
                lineas = result.stdout.strip().split('\n')
                if len(lineas) > 10:
                    print("  [...]")
                    for linea in lineas[-10:]:
                        print(f"  {linea}")
                else:
                    for linea in lineas:
                        print(f"  {linea}")
            return True
        else:
            print(f"✗ {descripcion} - FALLÓ")
            if result.stderr:
                print(f"  Error: {result.stderr}")
            return False

    except subprocess.CalledProcessError as e:
        print(f"✗ {descripcion} - ERROR")
        print(f"  Código de salida: {e.returncode}")
        if e.stderr:
            print(f"  Error: {e.stderr}")
        if check:
            raise
        return False


def copiar_archivo(origen, destino, descripcion):
    """Copia un archivo con validación"""
    print(f"\n→ Copiando {descripcion}")
    print(f"  Origen: {origen}")
    print(f"  Destino: {destino}")

    if not origen.exists():
        print(f"✗ Archivo origen no existe")
        return False

    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(origen, destino)
        tamano_mb = destino.stat().st_size / (1024 * 1024)
        print(f"✓ Archivo copiado ({tamano_mb:.1f} MB)")
        return True
    except Exception as e:
        print(f"✗ Error al copiar archivo: {e}")
        return False


# ============================================
# PASOS DEL PROCESO
# ============================================

def paso_1_crear_directorios(config):
    """PASO 1: Crear estructura de directorios"""
    log_paso(1, 8, "Crear estructura de directorios")

    try:
        crear_directorios_mes(config)
        return True
    except Exception as e:
        print(f"✗ Error al crear directorios: {e}")
        return False


def paso_2_copiar_archivos_entrada(config):
    """PASO 2: Copiar archivos de entrada al directorio del mes"""
    log_paso(2, 8, "Copiar archivos de entrada")

    # Copiar archivos de entrada a datos_intermedios para trazabilidad
    archivos = config['archivos_entrada']
    dir_destino = config['dir_datos_intermedios']

    todo_ok = True

    # Copiar PE-04 Formación
    todo_ok &= copiar_archivo(
        archivos['pe04_formacion'],
        dir_destino / archivos['pe04_formacion'].name,
        "PE-04 Formación"
    )

    # Copiar Avance Cupos
    todo_ok &= copiar_archivo(
        archivos['avance_cupos'],
        dir_destino / archivos['avance_cupos'].name,
        "Avance Cupos"
    )

    # Copiar Avance Aprendices
    todo_ok &= copiar_archivo(
        archivos['avance_aprendices'],
        dir_destino / archivos['avance_aprendices'].name,
        "Avance Aprendices"
    )

    # Copiar Metas SENA (si no existe ya)
    metas_destino = dir_destino / archivos['metas_sena'].name
    if not metas_destino.exists():
        todo_ok &= copiar_archivo(
            archivos['metas_sena'],
            metas_destino,
            "Metas SENA"
        )
    else:
        print(f"\n→ Metas SENA ya existe en destino (omitiendo copia)")

    return todo_ok


def paso_3_generar_bd_formacion(config):
    """PASO 3: Generar base de datos de formación"""
    log_paso(3, 8, "Generar base de datos de formación")

    # Preparar variables de entorno
    env = os.environ.copy()
    env['ARCHIVO_ENTRADA'] = str(config['dir_datos_intermedios'] / config['archivos_entrada']['pe04_formacion'].name)
    env['ARCHIVO_SALIDA'] = str(config['archivos_intermedios']['bd_formacion'])
    env['MES'] = config['mes_nombre']

    # Ejecutar script de importación
    return ejecutar_comando(
        ['python', str(config['scripts']['importar_pe04'])],
        f"Importar datos de PE-04 para {config['mes_nombre']}",
        check=True
    )


def paso_4_crear_tabla_economia_naranja(config):
    """PASO 4: Crear tabla de economía naranja"""
    log_paso(4, 8, "Crear tabla de economía naranja")

    # Leer SQL template
    sql_file = config['scripts']['crear_tabla_economia_naranja']

    if not sql_file.exists():
        print(f"✗ Archivo SQL no encontrado: {sql_file}")
        return False

    try:
        with open(sql_file, 'r', encoding='utf-8') as f:
            sql_template = f.read()

        # Reemplazar variables en el SQL
        # Cambiar SEPTIEMBRE por el mes actual en todos los lugares
        sql_ajustado = sql_template.replace('SEPTIEMBRE', config['mes_nombre'].upper())
        sql_ajustado = sql_ajustado.replace('AGOSTO', config['mes_nombre'].upper())

        # Conectar a la BD y ejecutar
        bd_formacion = config['archivos_intermedios']['bd_formacion']

        if not bd_formacion.exists():
            print(f"✗ Base de datos de formación no existe: {bd_formacion}")
            return False

        print(f"\n→ Conectando a BD: {bd_formacion}")
        conn = sqlite3.connect(bd_formacion)
        cursor = conn.cursor()

        print(f"→ Ejecutando SQL para tabla: {config['tablas_bd']['economia_naranja']}")
        cursor.executescript(sql_ajustado)
        conn.commit()

        # Verificar que la tabla se creó
        cursor.execute(f"SELECT COUNT(*) FROM {config['tablas_bd']['economia_naranja']}")
        count = cursor.fetchone()[0]

        conn.close()

        print(f"✓ Tabla creada con {count} registros")
        return True

    except Exception as e:
        print(f"✗ Error al crear tabla: {e}")
        return False


def paso_5_generar_bd_metas(config):
    """PASO 5: Generar base de datos de metas"""
    log_paso(5, 8, "Generar base de datos de metas")

    # Preparar variables de entorno
    env = os.environ.copy()
    env['ARCHIVO_METAS'] = str(config['dir_datos_intermedios'] / config['archivos_entrada']['metas_sena'].name)
    env['BD_SALIDA'] = str(config['archivos_intermedios']['bd_metas'])

    # Cambiar al directorio de metas para ejecutar el script
    cwd_original = os.getcwd()
    os.chdir(config['scripts']['normalizar_metas'].parent)

    try:
        resultado = ejecutar_comando(
            ['python', 'normalizar_metas_sena.py'],
            "Normalizar metas SENA",
            check=True
        )
        return resultado
    finally:
        os.chdir(cwd_original)


def paso_6_calcular_cupos_disponibles(config):
    """PASO 6: Calcular cupos disponibles"""
    log_paso(6, 8, "Calcular cupos disponibles (META - AVANCE)")

    # Preparar variables de entorno
    env = os.environ.copy()
    env['BD_METAS'] = str(config['archivos_intermedios']['bd_metas'])
    env['ARCHIVO_AVANCE'] = str(config['dir_datos_intermedios'] / config['archivos_entrada']['avance_cupos'].name)
    env['ARCHIVO_SALIDA'] = str(config['archivos_intermedios']['cupos_disponibles_xlsx'])

    # Cambiar al directorio de metas
    cwd_original = os.getcwd()
    os.chdir(config['scripts']['cruce_metas_avance'].parent)

    try:
        resultado = ejecutar_comando(
            ['python', 'cruce_metas_avance_final.py'],
            "Calcular cupos disponibles",
            check=True
        )

        # Mover archivos generados a datos_intermedios
        if resultado:
            for ext in ['.xlsx', '.csv']:
                archivo_origen = Path(f'cupos_disponibles_por_regional_2025{ext}')
                if archivo_origen.exists():
                    shutil.move(
                        str(archivo_origen),
                        str(config['archivos_intermedios'] / archivo_origen.name)
                    )

        return resultado
    finally:
        os.chdir(cwd_original)


def paso_7_generar_reporte_aprendices(config):
    """PASO 7: Generar reporte de aprendices"""
    log_paso(7, 8, "Generar reporte mensual de aprendices")

    # Preparar variables de entorno
    env = os.environ.copy()
    env['ARCHIVO_APRENDICES'] = str(config['dir_datos_intermedios'] / config['archivos_entrada']['avance_aprendices'].name)
    env['ARCHIVO_SALIDA'] = str(config['archivos_intermedios']['reporte_aprendices'])

    # Cambiar al directorio de aprendices
    cwd_original = os.getcwd()
    os.chdir(config['scripts']['generar_reporte_aprendices'].parent)

    try:
        resultado = ejecutar_comando(
            ['python', 'generar_reporte_mensual_aprendices.py'],
            "Generar reporte de aprendices",
            check=True
        )

        # Mover archivo generado a datos_intermedios
        if resultado:
            archivo_generado = Path(f"SENA Mensual Nacional {config['mes_corto']} {config['anio']}.xlsx")
            if archivo_generado.exists():
                shutil.move(
                    str(archivo_generado),
                    str(config['archivos_intermedios']['reporte_aprendices'])
                )

        return resultado
    finally:
        os.chdir(cwd_original)


def paso_8_generar_reporte_consolidado(config):
    """PASO 8: Generar reporte consolidado final"""
    log_paso(8, 8, "Generar reporte consolidado final")

    # Preparar variables de entorno
    env = os.environ.copy()
    env['BD_FORMACION'] = str(config['archivos_intermedios']['bd_formacion'])
    env['CUPOS_DISPONIBLES'] = str(config['archivos_intermedios']['cupos_disponibles_xlsx'])
    env['REPORTE_APRENDICES'] = str(config['archivos_intermedios']['reporte_aprendices'])
    env['ARCHIVO_SALIDA'] = str(config['archivos_finales']['reporte_consolidado'])
    env['MES_TRABAJO'] = config['mes_nombre']
    env['MES_CORTO'] = config['mes_corto']
    env['ANIO'] = str(config['anio'])

    # Cambiar al directorio del reporte consolidado
    cwd_original = os.getcwd()
    os.chdir(config['scripts']['generar_reporte_consolidado'].parent)

    try:
        resultado = ejecutar_comando(
            ['python', 'generar_reporte_consolidado.py'],
            "Generar reporte consolidado",
            check=True
        )

        # Mover archivo generado a datos_finales
        if resultado:
            archivo_generado = Path(f"Reporte Consolidado Economía Naranja {config['mes_corto']} {config['anio']}.xlsx")
            if archivo_generado.exists():
                shutil.move(
                    str(archivo_generado),
                    str(config['archivos_finales']['reporte_consolidado'])
                )
                print(f"\n✓ Reporte final guardado en:")
                print(f"  {config['archivos_finales']['reporte_consolidado']}")

        return resultado
    finally:
        os.chdir(cwd_original)


# ============================================
# FUNCIÓN PRINCIPAL
# ============================================

def main():
    print("="*70)
    print(" GENERACIÓN COMPLETA DEL REPORTE DE ECONOMÍA NARANJA")
    print("="*70)
    print(f" Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*70)

    # Validar argumentos
    if len(sys.argv) < 2:
        print("\n✗ Error: Falta especificar el mes")
        print("\nUso: python generar_reporte_completo.py <MES>")
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
        print(f"\n✓ Configuración cargada para {config['mes_nombre']} {config['anio']}")
    except Exception as e:
        print(f"\n✗ Error al cargar configuración: {e}")
        sys.exit(1)

    # Ejecutar pasos
    inicio = datetime.now()
    pasos = [
        paso_1_crear_directorios,
        paso_2_copiar_archivos_entrada,
        paso_3_generar_bd_formacion,
        paso_4_crear_tabla_economia_naranja,
        paso_5_generar_bd_metas,
        paso_6_calcular_cupos_disponibles,
        paso_7_generar_reporte_aprendices,
        paso_8_generar_reporte_consolidado
    ]

    for paso in pasos:
        try:
            if not paso(config):
                print(f"\n✗ ERROR EN {paso.__name__}")
                print("Proceso interrumpido.")
                sys.exit(1)
        except Exception as e:
            print(f"\n✗ EXCEPCIÓN EN {paso.__name__}: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)

    # Resumen final
    fin = datetime.now()
    duracion = (fin - inicio).total_seconds()

    print("\n" + "="*70)
    print(" PROCESO COMPLETADO EXITOSAMENTE")
    print("="*70)
    print(f"\n✓ Reporte generado: {config['archivos_finales']['reporte_consolidado']}")
    print(f"\n⏱ Tiempo total: {duracion:.1f} segundos ({duracion/60:.1f} minutos)")
    print(f"\n📁 Archivos intermedios en: {config['dir_datos_intermedios']}")
    print(f"📁 Reporte final en: {config['dir_datos_finales']}")
    print("\n" + "="*70)


if __name__ == '__main__':
    main()
