"""
Script para importar y normalizar datos de formación SENA desde archivos Excel (.xlsb)
a bases de datos SQLite independientes por mes.

Uso:
    python importar_mes.py SEPTIEMBRE
    python importar_mes.py AGOSTO
"""

import pyxlsb
import sqlite3
import sys
from pathlib import Path
from datetime import datetime
import re
import io
import pandas as pd

# Configurar codificación para Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

class ImportadorFormacionSENA:
    """Importa y normaliza datos de formación SENA desde Excel a SQLite"""

    def __init__(self, directorio, mes):
        self.mes = mes.upper()
        self.directorio = Path(directorio)
        self.archivo_excel = self.directorio / f"PE-04_FORMACION NACIONAL {self.mes} 2025.xlsb"
        self.archivo_db = self.directorio / f"sena_formacion_{self.mes.lower()}.db"

        # Buscar el catálogo de economía naranja (puede estar en varios lugares)
        self.catalogo_eco_naranja = self._buscar_catalogo_economia_naranja()

        # Mapeo de columnas según la estructura de la BD existente
        self.tablas = {
            'regionales': set(),
            'centros': set(),
            'niveles_formacion': set(),
            'jornadas': set(),
            'sectores_programa': set(),
            'ocupaciones': set(),
            'programas': set(),
            'ubicaciones': set(),
            'convenios': set(),
            'programas_especiales': set(),
            'empresas': set(),
            'programas_economia_naranja': set()
        }

    def _buscar_catalogo_economia_naranja(self):
        """Busca el archivo de catálogo de economía naranja en ubicaciones conocidas"""
        posibles_rutas = [
            self.directorio / "CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx",
            self.directorio.parent / "CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx",
            Path("C:/ws/sena/data/REPORTE_ECONOMIA_NARANJA/CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx"),
            Path("C:/ws/sena/data/2025/09-Septiembre/CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx")
        ]

        for ruta in posibles_rutas:
            if ruta.exists():
                print(f"[OK] Catálogo de Economía Naranja encontrado: {ruta}")
                return ruta

        print("[!] ADVERTENCIA: No se encontró el catálogo CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx")
        return None

    def _cargar_catalogo_economia_naranja(self, cursor):
        """Carga los programas de economía naranja desde el catálogo Excel"""
        try:
            print(f"\n→ Cargando catálogo de Economía Naranja...")

            # Leer el archivo Excel
            df = pd.read_excel(self.catalogo_eco_naranja)

            # Validar que existan las columnas necesarias
            columnas_requeridas = ['CODIGO', 'VERSION', 'NOMBRE DE PROGRAMA']
            columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]

            if columnas_faltantes:
                print(f"[!] El catálogo no tiene las columnas esperadas: {', '.join(columnas_faltantes)}")
                return

            # Insertar cada programa en la tabla
            programas_insertados = 0
            for _, row in df.iterrows():
                codigo = row['CODIGO']
                version = row['VERSION']
                nombre = row['NOMBRE DE PROGRAMA']

                if pd.notna(codigo) and pd.notna(version) and pd.notna(nombre):
                    cursor.execute("""
                        INSERT OR IGNORE INTO programas_economia_naranja (CODIGO_PROGRAMA, VERSION_PROGRAMA, NOMBRE_PROGRAMA)
                        VALUES (?, ?, ?)
                    """, (int(codigo), int(version), str(nombre)))
                    programas_insertados += 1

            print(f"[OK] Programas de Economía Naranja cargados: {programas_insertados}")

        except Exception as e:
            print(f"[!] Error al cargar catálogo de Economía Naranja: {e}")

    def validar_archivo(self):
        """Valida que el archivo Excel exista"""
        if not self.archivo_excel.exists():
            raise FileNotFoundError(f"No se encontro el archivo: {self.archivo_excel}")
        print(f"[OK] Archivo encontrado: {self.archivo_excel.name}")

    def crear_base_datos(self):
        """Crea la estructura de la base de datos SQLite"""
        conn = sqlite3.connect(self.archivo_db)
        cursor = conn.cursor()

        # Tabla regionales
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS regionales (
                CODIGO_REGIONAL INTEGER PRIMARY KEY,
                NOMBRE_REGIONAL TEXT
            )
        """)

        # Tabla centros
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS centros (
                CODIGO_CENTRO INTEGER PRIMARY KEY,
                NOMBRE_CENTRO TEXT,
                CODIGO_REGIONAL INTEGER,
                FOREIGN KEY (CODIGO_REGIONAL) REFERENCES regionales(CODIGO_REGIONAL)
            )
        """)

        # Tabla niveles_formacion
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS niveles_formacion (
                CODIGO_NIVEL_FORMACION INTEGER PRIMARY KEY,
                NOMBRE_NIVEL_FORMACION TEXT
            )
        """)

        # Tabla jornadas
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS jornadas (
                CODIGO_JORNADA INTEGER PRIMARY KEY,
                NOMBRE_JORNADA TEXT
            )
        """)

        # Tabla sectores_programa
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sectores_programa (
                CODIGO_SECTOR_PROGRAMA INTEGER PRIMARY KEY,
                NOMBRE_SECTOR_PROGRAMA TEXT
            )
        """)

        # Tabla ocupaciones
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ocupaciones (
                CODIGO_OCUPACION INTEGER PRIMARY KEY,
                NOMBRE_OCUPACION TEXT
            )
        """)

        # Tabla programas
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS programas (
                CODIGO_PROGRAMA INTEGER,
                VERSION_PROGRAMA INTEGER,
                NOMBRE_PROGRAMA TEXT,
                TIPO_PROGRAMA TEXT,
                ESTADO_PROGRAMA TEXT,
                CODIGO_OCUPACION INTEGER,
                CODIGO_SECTOR_PROGRAMA INTEGER,
                PRIMARY KEY (CODIGO_PROGRAMA, VERSION_PROGRAMA),
                FOREIGN KEY (CODIGO_OCUPACION) REFERENCES ocupaciones(CODIGO_OCUPACION),
                FOREIGN KEY (CODIGO_SECTOR_PROGRAMA) REFERENCES sectores_programa(CODIGO_SECTOR_PROGRAMA)
            )
        """)

        # Tabla ubicaciones
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ubicaciones (
                CODIGO_PAIS INTEGER,
                CODIGO_DEPARTAMENTO INTEGER,
                CODIGO_MUNICIPIO INTEGER,
                NOMBRE_PAIS TEXT,
                NOMBRE_DEPARTAMENTO TEXT,
                NOMBRE_MUNICIPIO TEXT,
                PRIMARY KEY (CODIGO_PAIS, CODIGO_DEPARTAMENTO, CODIGO_MUNICIPIO)
            )
        """)

        # Tabla convenios
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS convenios (
                CODIGO_CONVENIO REAL PRIMARY KEY,
                NOMBRE_CONVENIO TEXT
            )
        """)

        # Tabla programas_especiales
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS programas_especiales (
                CODIGO_PROGRAMA_ESPECIAL INTEGER PRIMARY KEY,
                NOMBRE_PROGRAMA_ESPECIAL TEXT
            )
        """)

        # Tabla empresas
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS empresas (
                NUMERO_IDENTIFICACION_EMPRESA TEXT PRIMARY KEY,
                NOMBRE_EMPRESA TEXT
            )
        """)

        # Tabla programas_economia_naranja
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS programas_economia_naranja (
                CODIGO_PROGRAMA INTEGER,
                VERSION_PROGRAMA INTEGER,
                NOMBRE_PROGRAMA TEXT,
                PRIMARY KEY (CODIGO_PROGRAMA, VERSION_PROGRAMA)
            )
        """)

        # Cargar datos desde el catálogo si existe
        if self.catalogo_eco_naranja:
            self._cargar_catalogo_economia_naranja(cursor)

        # Tabla fichas (tabla principal con todas las fichas de formación)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS fichas (
                IDENTIFICADOR_FICHA INTEGER PRIMARY KEY,
                IDENTIFICADOR_UNICO_FICHA INTEGER,
                ESTADO_CURSO TEXT,
                CODIGO_NIVEL_FORMACION INTEGER,
                CODIGO_JORNADA INTEGER,
                A_LA_MEDIDA TEXT,
                FECHA_INICIO_FICHA TIMESTAMP,
                FECHA_TERMINACION_FICHA TIMESTAMP,
                ETAPA_FICHA TEXT,
                MODALIDAD_FORMACION TEXT,
                NOMBRE_RESPONSABLE TEXT,
                CODIGO_CENTRO INTEGER,
                NUMERO_IDENTIFICACION_EMPRESA TEXT,
                CODIGO_PROGRAMA INTEGER,
                VERSION_PROGRAMA INTEGER,
                CODIGO_PAIS_CURSO INTEGER,
                CODIGO_DEPARTAMENTO_CURSO INTEGER,
                CODIGO_MUNICIPIO_CURSO INTEGER,
                CODIGO_CONVENIO REAL,
                AMPLICACION_COBERTURA TEXT,
                DESTINO_INFORMACION TEXT,
                CODIGO_PROGRAMA_ESPECIAL INTEGER,
                NUMERO_CURSOS INTEGER,
                TOTAL_APRENDICES_MASCULINOS INTEGER,
                TOTAL_APRENDICES_FEMENINOS INTEGER,
                TOTAL_APRENDICES_NO_BINARIO INTEGER,
                TOTAL_APRENDICES INTEGER,
                HORAS_PLANTA REAL,
                HORAS_CONTRATISTAS REAL,
                HORAS_CONTRATISTAS_EXTERNOS REAL,
                HORAS_MONITORES REAL,
                HORAS_INST_EMPRESA REAL,
                TOTAL_HORAS REAL,
                TOTAL_APRENDICES_ACTIVO INTEGER,
                DURACION_PROGRAMA INTEGER,
                NOMBRE_NUEVO_SECTOR TEXT,
                FOREIGN KEY (CODIGO_NIVEL_FORMACION) REFERENCES niveles_formacion(CODIGO_NIVEL_FORMACION),
                FOREIGN KEY (CODIGO_JORNADA) REFERENCES jornadas(CODIGO_JORNADA),
                FOREIGN KEY (CODIGO_CENTRO) REFERENCES centros(CODIGO_CENTRO),
                FOREIGN KEY (NUMERO_IDENTIFICACION_EMPRESA) REFERENCES empresas(NUMERO_IDENTIFICACION_EMPRESA),
                FOREIGN KEY (CODIGO_PROGRAMA, VERSION_PROGRAMA) REFERENCES programas(CODIGO_PROGRAMA, VERSION_PROGRAMA),
                FOREIGN KEY (CODIGO_CONVENIO) REFERENCES convenios(CODIGO_CONVENIO),
                FOREIGN KEY (CODIGO_PROGRAMA_ESPECIAL) REFERENCES programas_especiales(CODIGO_PROGRAMA_ESPECIAL)
            )
        """)

        conn.commit()
        print(f"[OK] Base de datos creada: {self.archivo_db.name}")
        return conn

    def leer_excel(self):
        """Lee el archivo Excel y retorna los datos"""
        print(f"Leyendo archivo Excel...")
        wb = pyxlsb.open_workbook(str(self.archivo_excel))

        # Obtener la primera hoja (hoja de datos principal)
        nombre_hoja = wb.sheets[0]
        print(f"[OK] Procesando hoja: {nombre_hoja}")

        ws = wb.get_sheet(nombre_hoja)
        rows = []

        print("Extrayendo filas...")
        for idx, row in enumerate(ws.rows()):
            if idx % 10000 == 0 and idx > 0:
                print(f"  Procesadas {idx} filas...")
            rows.append([cell.v if cell else None for cell in row])

        print(f"[OK] Total de filas leidas: {len(rows)}")
        return rows

    def normalizar_e_importar(self, rows):
        """Normaliza los datos y los importa a la base de datos"""
        conn = self.crear_base_datos()
        cursor = conn.cursor()

        if len(rows) < 2:
            raise ValueError("El archivo no contiene suficientes datos")

        # Buscar la fila de encabezados (la que contiene "IDENTIFICADOR_FICHA" o "CODIGO_REGIONAL")
        header_row_idx = None
        for idx, row in enumerate(rows[:20]):  # Buscar en las primeras 20 filas
            row_str = ' '.join([str(cell).upper() if cell else '' for cell in row])
            if 'IDENTIFICADOR_FICHA' in row_str or 'CODIGO_REGIONAL' in row_str:
                header_row_idx = idx
                print(f"[OK] Encabezados encontrados en fila {idx + 1}")
                break

        if header_row_idx is None:
            raise ValueError("No se encontraron los encabezados en el archivo")

        # Extraer encabezados y datos
        headers = rows[header_row_idx]
        data_rows = rows[header_row_idx + 1:]

        # Normalizar nombres de encabezados (remover espacios extra, convertir a mayúsculas)
        headers = [str(h).strip().upper() if h else None for h in headers]

        print(f"[OK] Columnas encontradas: {len([h for h in headers if h])}")
        print(f"[OK] Filas de datos: {len(data_rows)}")

        # Crear diccionario para almacenar índices de columnas
        col_idx = {header: idx for idx, header in enumerate(headers) if header}

        print("\nNormalizando e importando datos...")
        fichas_procesadas = 0

        # Procesar cada fila de datos
        for row_num, row in enumerate(data_rows, start=header_row_idx + 2):
            if row_num % 5000 == 0:
                print(f"  Procesando fila {row_num}...")
                conn.commit()  # Commit periódico

            try:
                # Extraer valores de la fila
                def get_val(col_name, default=None):
                    idx = col_idx.get(col_name)
                    if idx is not None and idx < len(row):
                        val = row[idx]
                        return val if val is not None else default
                    return default

                # Insertar/actualizar regional
                codigo_regional = get_val('CODIGO_REGIONAL')
                nombre_regional = get_val('NOMBRE_REGIONAL')
                if codigo_regional:
                    cursor.execute("""
                        INSERT OR IGNORE INTO regionales (CODIGO_REGIONAL, NOMBRE_REGIONAL)
                        VALUES (?, ?)
                    """, (codigo_regional, nombre_regional))

                # Insertar/actualizar centro
                codigo_centro = get_val('CODIGO_CENTRO')
                nombre_centro = get_val('NOMBRE_CENTRO')
                if codigo_centro:
                    cursor.execute("""
                        INSERT OR IGNORE INTO centros (CODIGO_CENTRO, NOMBRE_CENTRO, CODIGO_REGIONAL)
                        VALUES (?, ?, ?)
                    """, (codigo_centro, nombre_centro, codigo_regional))

                # Insertar/actualizar nivel de formación
                codigo_nivel = get_val('CODIGO_NIVEL_FORMACION')
                nombre_nivel = get_val('NIVEL_FORMACION')
                if codigo_nivel:
                    cursor.execute("""
                        INSERT OR IGNORE INTO niveles_formacion (CODIGO_NIVEL_FORMACION, NOMBRE_NIVEL_FORMACION)
                        VALUES (?, ?)
                    """, (codigo_nivel, nombre_nivel))

                # Insertar/actualizar jornada
                codigo_jornada = get_val('CODIGO_JORNADA')
                nombre_jornada = get_val('NOMBRE_JORNADA')
                if codigo_jornada:
                    cursor.execute("""
                        INSERT OR IGNORE INTO jornadas (CODIGO_JORNADA, NOMBRE_JORNADA)
                        VALUES (?, ?)
                    """, (codigo_jornada, nombre_jornada))

                # Insertar/actualizar sector programa
                codigo_sector = get_val('CODIGO_SECTOR_PROGRAMA')
                nombre_sector = get_val('NOMBRE_SECTOR_PROGRAMA')
                if codigo_sector:
                    cursor.execute("""
                        INSERT OR IGNORE INTO sectores_programa (CODIGO_SECTOR_PROGRAMA, NOMBRE_SECTOR_PROGRAMA)
                        VALUES (?, ?)
                    """, (codigo_sector, nombre_sector))

                # Insertar/actualizar ocupación
                codigo_ocupacion = get_val('CODIGO_OCUPACION')
                nombre_ocupacion = get_val('NOMBRE_OCUPACION')
                if codigo_ocupacion:
                    cursor.execute("""
                        INSERT OR IGNORE INTO ocupaciones (CODIGO_OCUPACION, NOMBRE_OCUPACION)
                        VALUES (?, ?)
                    """, (codigo_ocupacion, nombre_ocupacion))

                # Insertar/actualizar programa
                codigo_programa = get_val('CODIGO_PROGRAMA')
                version_programa = get_val('VERSION_PROGRAMA')
                nombre_programa = get_val('NOMBRE_PROGRAMA_FORMACION')
                tipo_programa = None  # No existe en el Excel
                estado_programa = None  # No existe en el Excel
                if codigo_programa and version_programa:
                    cursor.execute("""
                        INSERT OR IGNORE INTO programas
                        (CODIGO_PROGRAMA, VERSION_PROGRAMA, NOMBRE_PROGRAMA, TIPO_PROGRAMA,
                         ESTADO_PROGRAMA, CODIGO_OCUPACION, CODIGO_SECTOR_PROGRAMA)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (codigo_programa, version_programa, nombre_programa, tipo_programa,
                          estado_programa, codigo_ocupacion, codigo_sector))

                # Insertar/actualizar ubicación
                codigo_pais = get_val('CODIGO_PAIS_CURSO')
                codigo_depto = get_val('CODIGO_DEPARTAMENTO_CURSO')
                codigo_muni = get_val('CODIGO_MUNICIPIO_CURSO')
                nombre_pais = get_val('NOMBRE_PAIS_CURSO')
                nombre_depto = get_val('NOMBRE_DEPARTAMENTO_CURSO')
                nombre_muni = get_val('NOMBRE_MUNICIPIO_CURSO')
                if codigo_pais and codigo_depto and codigo_muni:
                    cursor.execute("""
                        INSERT OR IGNORE INTO ubicaciones
                        (CODIGO_PAIS, CODIGO_DEPARTAMENTO, CODIGO_MUNICIPIO,
                         NOMBRE_PAIS, NOMBRE_DEPARTAMENTO, NOMBRE_MUNICIPIO)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (codigo_pais, codigo_depto, codigo_muni,
                          nombre_pais, nombre_depto, nombre_muni))

                # Insertar/actualizar convenio
                codigo_convenio = get_val('CODIGO_CONVENIO')
                nombre_convenio = get_val('NOMBRE_CONVENIO')
                if codigo_convenio:
                    cursor.execute("""
                        INSERT OR IGNORE INTO convenios (CODIGO_CONVENIO, NOMBRE_CONVENIO)
                        VALUES (?, ?)
                    """, (codigo_convenio, nombre_convenio))

                # Insertar/actualizar programa especial
                codigo_prog_esp = get_val('CODIGO_PROGRAMA_ESPECIAL')
                nombre_prog_esp = get_val('NOMBRE_PROGRAMA_ESPECIAL')
                if codigo_prog_esp:
                    cursor.execute("""
                        INSERT OR IGNORE INTO programas_especiales
                        (CODIGO_PROGRAMA_ESPECIAL, NOMBRE_PROGRAMA_ESPECIAL)
                        VALUES (?, ?)
                    """, (codigo_prog_esp, nombre_prog_esp))

                # Insertar/actualizar empresa
                num_id_empresa = get_val('NUMERO_IDENTIFICACION_EMPRESA')
                nombre_empresa = get_val('NOMBRE_EMPRESA')
                if num_id_empresa:
                    cursor.execute("""
                        INSERT OR IGNORE INTO empresas
                        (NUMERO_IDENTIFICACION_EMPRESA, NOMBRE_EMPRESA)
                        VALUES (?, ?)
                    """, (num_id_empresa, nombre_empresa))

                # No hay columna de economía naranja directa, se infiere de otro modo si es necesario

                # Insertar ficha principal
                identificador_ficha = get_val('IDENTIFICADOR_FICHA')
                if identificador_ficha:
                    cursor.execute("""
                        INSERT OR REPLACE INTO fichas (
                            IDENTIFICADOR_FICHA, IDENTIFICADOR_UNICO_FICHA, ESTADO_CURSO,
                            CODIGO_NIVEL_FORMACION, CODIGO_JORNADA, A_LA_MEDIDA,
                            FECHA_INICIO_FICHA, FECHA_TERMINACION_FICHA, ETAPA_FICHA,
                            MODALIDAD_FORMACION, NOMBRE_RESPONSABLE, CODIGO_CENTRO,
                            NUMERO_IDENTIFICACION_EMPRESA, CODIGO_PROGRAMA, VERSION_PROGRAMA,
                            CODIGO_PAIS_CURSO, CODIGO_DEPARTAMENTO_CURSO, CODIGO_MUNICIPIO_CURSO,
                            CODIGO_CONVENIO, AMPLICACION_COBERTURA, DESTINO_INFORMACION,
                            CODIGO_PROGRAMA_ESPECIAL, NUMERO_CURSOS,
                            TOTAL_APRENDICES_MASCULINOS, TOTAL_APRENDICES_FEMENINOS,
                            TOTAL_APRENDICES_NO_BINARIO, TOTAL_APRENDICES,
                            HORAS_PLANTA, HORAS_CONTRATISTAS, HORAS_CONTRATISTAS_EXTERNOS,
                            HORAS_MONITORES, HORAS_INST_EMPRESA, TOTAL_HORAS,
                            TOTAL_APRENDICES_ACTIVO, DURACION_PROGRAMA, NOMBRE_NUEVO_SECTOR
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                  ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        identificador_ficha, get_val('IDENTIFICADOR_UNICO_FICHA'),
                        get_val('ESTADO_CURSO'), codigo_nivel, codigo_jornada,
                        get_val('A_LA_MEDIDA'), get_val('FECHA_INICIO_FICHA'),
                        get_val('FECHA_TERMINACION_FICHA'), get_val('ETAPA_FICHA'),
                        get_val('MODALIDAD_FORMACION'), get_val('NOMBRE_RESPONSABLE'),
                        codigo_centro, num_id_empresa, codigo_programa, version_programa,
                        codigo_pais, codigo_depto, codigo_muni, codigo_convenio,
                        get_val('AMPLICACION_COBERTURA'), get_val('DESTINO INFORMACIÓN'),
                        codigo_prog_esp, get_val('NUMERO_CURSOS'),
                        get_val('TOTAL_APRENDICES_MASCULINOS'), get_val('TOTAL_APRENDICES_FEMENINOS'),
                        get_val('TOTAL_APRENDICES_NO_BINARIO'), get_val('TOTAL_APRENDICES'),
                        get_val('HORAS_PLANTA'), get_val('HORAS_CONTRATISTAS'),
                        get_val('HORAS_CONTRATISTAS_EXTERNOS'), get_val('HORAS_MONITORES'),
                        get_val('HORAS_INST_EMPRESA'), get_val('TOTAL_HORAS'),
                        get_val('TOTAL_APRENDICES_ACTIVO'), get_val('DURACION_PROGRAMA'),
                        get_val('NOMBRE_NUEVO_SECTOR')
                    ))
                    fichas_procesadas += 1

            except Exception as e:
                print(f"\n[!] Error en fila {row_num}: {e}")
                continue

        # Commit final
        conn.commit()
        print(f"\n[OK] Fichas importadas: {fichas_procesadas}")

        # Mostrar estadísticas
        self.mostrar_estadisticas(cursor)

        conn.close()

    def mostrar_estadisticas(self, cursor):
        """Muestra estadísticas de la importación"""
        print("\n" + "="*60)
        print(f"ESTADÍSTICAS DE IMPORTACIÓN - {self.mes} 2025")
        print("="*60)

        tablas_stats = [
            ('regionales', 'regionales'),
            ('centros', 'centros de formación'),
            ('niveles_formacion', 'niveles de formación'),
            ('jornadas', 'jornadas'),
            ('sectores_programa', 'sectores de programa'),
            ('ocupaciones', 'ocupaciones'),
            ('programas', 'programas'),
            ('ubicaciones', 'ubicaciones'),
            ('convenios', 'convenios'),
            ('programas_especiales', 'programas especiales'),
            ('empresas', 'empresas'),
            ('programas_economia_naranja', 'programas economía naranja'),
            ('fichas', 'fichas de formación')
        ]

        for tabla, descripcion in tablas_stats:
            cursor.execute(f"SELECT COUNT(*) FROM {tabla}")
            count = cursor.fetchone()[0]
            print(f"  {descripcion.capitalize()}: {count:,}")

        print("="*60)

    def ejecutar(self):
        """Ejecuta el proceso completo de importación"""
        print("\n" + "="*60)
        print(f"IMPORTADOR DE FORMACIÓN SENA - {self.mes} 2025")
        print("="*60 + "\n")

        try:
            # Validar archivo
            self.validar_archivo()

            # Leer Excel
            rows = self.leer_excel()

            # Normalizar e importar
            self.normalizar_e_importar(rows)

            print(f"\n[OK] Importacion completada exitosamente")
            print(f"[*] Base de datos: {self.archivo_db}\n")

        except Exception as e:
            print(f"\n[ERROR] Error durante la importacion: {e}\n")
            raise

def main():
    """Función principal"""
    if len(sys.argv) < 3:
        print("Uso: python importar_mes.py <DIRECTORIO> <MES>")
        print("Ejemplo: python importar_mes.py c:\\ws\\sena\\data\\ SEPTIEMBRE")
        sys.exit(1)

    directorio = Path(sys.argv[1])
    mes = sys.argv[2]
    importador = ImportadorFormacionSENA(directorio, mes)
    importador.ejecutar()

if __name__ == "__main__":
    main()
