import pandas as pd
import sqlite3
from datetime import datetime

# Configuración
excel_file = r'C:\ws\sena\data\2025\09-Septiembre\Metas SENA 2025 V5 26092025_CLEAN.xlsx'
db_file = r'C:\ws\sena\data\metas_sena_2025.db'

# Leer el archivo Excel
print("Leyendo archivo Excel...")
df = pd.read_excel(excel_file, sheet_name='METAS FORMACION X REGIONAL', header=None)

# Crear conexión a SQLite
print("Conectando a base de datos SQLite...")
conn = sqlite3.connect(db_file)
cursor = conn.cursor()

# ===== TABLAS NORMALIZADAS =====

# 1. Tabla de Regionales
cursor.execute('''
CREATE TABLE IF NOT EXISTS regionales (
    id_regional INTEGER PRIMARY KEY,
    codigo_regional INTEGER UNIQUE NOT NULL,
    nombre_regional TEXT NOT NULL,
    fecha_carga TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
''')

# 2. Tabla de Categorías de Formación
cursor.execute('''
CREATE TABLE IF NOT EXISTS categorias_formacion (
    id_categoria INTEGER PRIMARY KEY AUTOINCREMENT,
    categoria_principal TEXT NOT NULL,
    subcategoria TEXT,
    descripcion TEXT,
    tipo_medida TEXT DEFAULT 'Cupos',
    UNIQUE(categoria_principal, subcategoria)
)
''')

# 3. Tabla de Modalidades
cursor.execute('''
CREATE TABLE IF NOT EXISTS modalidades (
    id_modalidad INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre_modalidad TEXT UNIQUE NOT NULL,
    tipo_formacion TEXT
)
''')

# 4. Tabla de Metas (Cupos)
cursor.execute('''
CREATE TABLE IF NOT EXISTS metas_cupos (
    id_meta INTEGER PRIMARY KEY AUTOINCREMENT,
    id_regional INTEGER NOT NULL,
    id_categoria INTEGER NOT NULL,
    anio INTEGER NOT NULL,
    valor INTEGER,
    FOREIGN KEY (id_regional) REFERENCES regionales(id_regional),
    FOREIGN KEY (id_categoria) REFERENCES categorias_formacion(id_categoria)
)
''')

# 5. Tabla de Metas de Retención
cursor.execute('''
CREATE TABLE IF NOT EXISTS metas_retencion (
    id_retencion INTEGER PRIMARY KEY AUTOINCREMENT,
    id_regional INTEGER NOT NULL,
    tipo_formacion TEXT NOT NULL,
    modalidad TEXT,
    anio INTEGER NOT NULL,
    valor INTEGER,
    FOREIGN KEY (id_regional) REFERENCES regionales(id_regional)
)
''')

# 6. Tabla de Metas de Certificación
cursor.execute('''
CREATE TABLE IF NOT EXISTS metas_certificacion (
    id_certificacion INTEGER PRIMARY KEY AUTOINCREMENT,
    id_regional INTEGER NOT NULL,
    tipo_formacion TEXT NOT NULL,
    anio INTEGER NOT NULL,
    valor INTEGER,
    FOREIGN KEY (id_regional) REFERENCES regionales(id_regional)
)
''')

# 7. Tabla de Programas Especiales
cursor.execute('''
CREATE TABLE IF NOT EXISTS programas_especiales (
    id_programa INTEGER PRIMARY KEY AUTOINCREMENT,
    id_regional INTEGER NOT NULL,
    nombre_programa TEXT NOT NULL,
    anio INTEGER NOT NULL,
    valor INTEGER,
    FOREIGN KEY (id_regional) REFERENCES regionales(id_regional)
)
''')

print("Tablas creadas exitosamente.")

# ===== INSERCIÓN DE DATOS =====

# Insertar Regionales
print("\nCargando regionales...")
regionales_data = []
for idx in range(3, len(df)):
    codigo = df.iloc[idx, 0]
    nombre = df.iloc[idx, 1]
    if pd.notna(codigo) and pd.notna(nombre):
        regionales_data.append((int(codigo), nombre))

cursor.executemany('INSERT OR IGNORE INTO regionales (codigo_regional, nombre_regional) VALUES (?, ?)',
                   regionales_data)
print(f"  {len(regionales_data)} regionales cargadas.")

# Mapeo de categorías de formación (columnas 2-37)
categorias_map = [
    # EDUCACIÓN SUPERIOR
    (2, 'EDUCACION SUPERIOR', 'Tecnologos Regular - Presencial', 'Cupos'),
    (3, 'EDUCACION SUPERIOR', 'Tecnólogos Regular - Virtual', 'Cupos'),
    (4, 'EDUCACION SUPERIOR', 'Tecnólogos Regular - A Distancia', 'Cupos'),
    (5, 'EDUCACION SUPERIOR', 'Tecnólogos CampeSENA', 'Cupos'),
    (6, 'EDUCACION SUPERIOR', 'Tecnólogos Full Popular', 'Cupos'),
    (7, 'EDUCACION SUPERIOR', 'Total Tecnólogos', 'Cupos'),
    (8, 'EDUCACION SUPERIOR', 'TOTAL EDUCACION SUPERIOR', 'Cupos'),

    # FORMACIÓN LABORAL
    (9, 'FORMACION LABORAL', 'Operarios Regular', 'Cupos'),
    (10, 'FORMACION LABORAL', 'Operarios CampeSENA', 'Cupos'),
    (11, 'FORMACION LABORAL', 'Operarios Full Popular', 'Cupos'),
    (12, 'FORMACION LABORAL', 'Total Operarios', 'Cupos'),
    (13, 'FORMACION LABORAL', 'Auxiliares Regular', 'Cupos'),
    (14, 'FORMACION LABORAL', 'Auxiliares CampeSENA', 'Cupos'),
    (15, 'FORMACION LABORAL', 'Auxiliares Full Popular', 'Cupos'),
    (16, 'FORMACION LABORAL', 'Total Auxiliares', 'Cupos'),
    (17, 'FORMACION LABORAL', 'Técnico Laboral Regular - Presencial', 'Cupos'),
    (18, 'FORMACION LABORAL', 'Técnico Laboral Regular - Virtual', 'Cupos'),
    (19, 'FORMACION LABORAL', 'Técnico Laboral CampeSENA', 'Cupos'),
    (20, 'FORMACION LABORAL', 'Técnico Laboral Full Popular', 'Cupos'),
    (21, 'FORMACION LABORAL', 'Técnico Laboral Articulación con la Media', 'Cupos'),
    (22, 'FORMACION LABORAL', 'Total Técnico Laboral', 'Cupos'),
    (23, 'FORMACION LABORAL', 'Total Profundización Técnica', 'Cupos'),
    (24, 'FORMACION LABORAL', 'TOTAL FORMACIÓN LABORAL', 'Cupos'),

    # TOTALES Y COMPLEMENTARIA
    (25, 'FORMACION TITULADA', 'TOTAL FORMACION TITULADA', 'Cupos'),
    (26, 'FORMACION COMPLEMENTARIA', 'Formación Complementaria - Virtual (Sin Bilingüismo)', 'Cupos'),
    (27, 'FORMACION COMPLEMENTARIA', 'Formación Complementaria - Presencial (Sin Bilingüismo)', 'Cupos'),
    (28, 'PROGRAMA DE BILINGUISMO', 'Programa de Bilingüismo - Virtual', 'Cupos'),
    (29, 'PROGRAMA DE BILINGUISMO', 'Programa de Bilingüismo - Presencial', 'Cupos'),
    (30, 'PROGRAMA DE BILINGUISMO', 'Total Programa de Bilingüismo', 'Cupos'),
    (31, 'FORMACION COMPLEMENTARIA', 'Formación Complementaria CampeSENA', 'Cupos'),
    (32, 'FORMACION COMPLEMENTARIA', 'Formación Complementaria Full Popular', 'Cupos'),
    (33, 'FORMACION COMPLEMENTARIA', 'TOTAL FORMACION COMPLEMENTARIA', 'Cupos'),
    (34, 'FORMACION PROFESIONAL INTEGRAL', 'TOTAL FORMACION PROFESIONAL INTEGRAL', 'Cupos'),

    # PROGRAMAS RELEVANTES
    (35, 'PROGRAMAS RELEVANTES', 'Total Formación Profesional CampeSENA', 'Cupos'),
    (36, 'PROGRAMAS RELEVANTES', 'Total Formación Profesional Full Popular', 'Cupos'),
    (37, 'PROGRAMAS RELEVANTES', 'Total Formación Profesional Integral - Virtual', 'Cupos'),
]

print("\nCargando categorías de formación...")
categorias_insert = []
cat_id_map = {}
for col_idx, cat_principal, subcat, tipo_medida in categorias_map:
    categorias_insert.append((cat_principal, subcat, tipo_medida))
    cat_id_map[col_idx] = (cat_principal, subcat)

cursor.executemany('''INSERT OR IGNORE INTO categorias_formacion
                     (categoria_principal, subcategoria, tipo_medida)
                     VALUES (?, ?, ?)''', categorias_insert)
print(f"  {len(categorias_insert)} categorías cargadas.")

# Obtener IDs de categorías
cat_ids = {}
for col_idx, (cat_principal, subcat) in cat_id_map.items():
    cursor.execute('''SELECT id_categoria FROM categorias_formacion
                     WHERE categoria_principal = ? AND subcategoria = ?''',
                   (cat_principal, subcat))
    result = cursor.fetchone()
    if result:
        cat_ids[col_idx] = result[0]

# Insertar Metas de Cupos
print("\nCargando metas de cupos...")
metas_cupos_data = []
for row_idx in range(3, len(df)):
    codigo_regional = df.iloc[row_idx, 0]
    if pd.notna(codigo_regional):
        # Validar que sea un número (omitir filas de totales)
        try:
            codigo_int = int(codigo_regional)
        except (ValueError, TypeError):
            continue

        # Obtener id_regional
        cursor.execute('SELECT id_regional FROM regionales WHERE codigo_regional = ?',
                      (codigo_int,))
        regional_result = cursor.fetchone()
        if regional_result:
            id_regional = regional_result[0]

            # Insertar valores de cada categoría
            for col_idx in cat_ids.keys():
                valor = df.iloc[row_idx, col_idx]
                if pd.notna(valor):
                    try:
                        metas_cupos_data.append((id_regional, cat_ids[col_idx], 2025, int(valor)))
                    except (ValueError, TypeError):
                        continue

cursor.executemany('''INSERT INTO metas_cupos
                     (id_regional, id_categoria, anio, valor)
                     VALUES (?, ?, ?, ?)''', metas_cupos_data)
print(f"  {len(metas_cupos_data)} metas de cupos cargadas.")

# Mapeo de columnas de Retención (38-57)
retencion_map = [
    (38, 'FORMACION LABORAL', 'Presencial'),
    (39, 'FORMACION LABORAL', 'Virtual'),
    (40, 'FORMACION LABORAL', 'TOTAL'),
    (41, 'EDUCACION SUPERIOR', 'Presencial'),
    (42, 'EDUCACION SUPERIOR', 'Virtual'),
    (43, 'EDUCACION SUPERIOR', 'TOTAL'),
    (44, 'FORMACION TITULADA', 'Presencial'),
    (45, 'FORMACION TITULADA', 'Virtual'),
    (46, 'FORMACION TITULADA', 'TOTAL'),
    (47, 'COMPLEMENTARIA', 'Presencial'),
    (48, 'COMPLEMENTARIA', 'Virtual'),
    (49, 'COMPLEMENTARIA', 'TOTAL'),
    (50, 'FORMACION PROFESIONAL', 'Presencial'),
    (51, 'FORMACION PROFESIONAL', 'Virtual'),
    (52, 'FORMACION PROFESIONAL', 'TOTAL'),
    (53, 'PROGRAMA DE BILINGUISMO', 'Presencial'),
    (54, 'PROGRAMA DE BILINGUISMO', 'Virtual'),
    (55, 'PROGRAMA DE BILINGUISMO', 'TOTAL'),
    (56, 'CampeSENA', None),
    (57, 'Full Popular', None),
]

# Insertar Metas de Retención
print("\nCargando metas de retención...")
metas_retencion_data = []
for row_idx in range(3, len(df)):
    codigo_regional = df.iloc[row_idx, 0]
    if pd.notna(codigo_regional):
        # Validar que sea un número (omitir filas de totales)
        try:
            codigo_int = int(codigo_regional)
        except (ValueError, TypeError):
            continue

        cursor.execute('SELECT id_regional FROM regionales WHERE codigo_regional = ?',
                      (codigo_int,))
        regional_result = cursor.fetchone()
        if regional_result:
            id_regional = regional_result[0]

            for col_idx, tipo_formacion, modalidad in retencion_map:
                valor = df.iloc[row_idx, col_idx]
                if pd.notna(valor):
                    try:
                        metas_retencion_data.append((id_regional, tipo_formacion, modalidad, 2025, int(valor)))
                    except (ValueError, TypeError):
                        continue

cursor.executemany('''INSERT INTO metas_retencion
                     (id_regional, tipo_formacion, modalidad, anio, valor)
                     VALUES (?, ?, ?, ?, ?)''', metas_retencion_data)
print(f"  {len(metas_retencion_data)} metas de retención cargadas.")

# Mapeo de columnas de Certificación (58-65)
certificacion_map = [
    (58, 'FORMACION LABORAL'),
    (59, 'EDUCACION SUPERIOR'),
    (60, 'FORMACION TITULADA'),
    (61, 'FORMACION COMPLEMENTARIA'),
    (62, 'FORMACION PROFESIONAL INTEGRAL'),
    (63, 'ARTICULACION CON LA MEDIA'),
    (64, 'CampeSENA'),
    (65, 'Full Popular'),
]

# Insertar Metas de Certificación
print("\nCargando metas de certificación...")
metas_certificacion_data = []
for row_idx in range(3, len(df)):
    codigo_regional = df.iloc[row_idx, 0]
    if pd.notna(codigo_regional):
        # Validar que sea un número (omitir filas de totales)
        try:
            codigo_int = int(codigo_regional)
        except (ValueError, TypeError):
            continue

        cursor.execute('SELECT id_regional FROM regionales WHERE codigo_regional = ?',
                      (codigo_int,))
        regional_result = cursor.fetchone()
        if regional_result:
            id_regional = regional_result[0]

            for col_idx, tipo_formacion in certificacion_map:
                valor = df.iloc[row_idx, col_idx]
                if pd.notna(valor):
                    try:
                        metas_certificacion_data.append((id_regional, tipo_formacion, 2025, int(valor)))
                    except (ValueError, TypeError):
                        continue

cursor.executemany('''INSERT INTO metas_certificacion
                     (id_regional, tipo_formacion, anio, valor)
                     VALUES (?, ?, ?, ?)''', metas_certificacion_data)
print(f"  {len(metas_certificacion_data)} metas de certificación cargadas.")

# Crear índices para mejorar performance
print("\nCreando índices...")
cursor.execute('CREATE INDEX IF NOT EXISTS idx_metas_cupos_regional ON metas_cupos(id_regional)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_metas_cupos_categoria ON metas_cupos(id_categoria)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_metas_retencion_regional ON metas_retencion(id_regional)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_metas_certificacion_regional ON metas_certificacion(id_regional)')

# Crear vistas útiles
print("\nCreando vistas...")

cursor.execute('''
CREATE VIEW IF NOT EXISTS vista_metas_cupos_completa AS
SELECT
    r.codigo_regional,
    r.nombre_regional,
    c.categoria_principal,
    c.subcategoria,
    m.anio,
    m.valor
FROM metas_cupos m
JOIN regionales r ON m.id_regional = r.id_regional
JOIN categorias_formacion c ON m.id_categoria = c.id_categoria
ORDER BY r.codigo_regional, c.categoria_principal, c.subcategoria
''')

cursor.execute('''
CREATE VIEW IF NOT EXISTS vista_resumen_regional AS
SELECT
    r.codigo_regional,
    r.nombre_regional,
    SUM(CASE WHEN c.categoria_principal = 'EDUCACION SUPERIOR' AND c.subcategoria = 'TOTAL EDUCACION SUPERIOR' THEN m.valor ELSE 0 END) as total_educacion_superior,
    SUM(CASE WHEN c.categoria_principal = 'FORMACION LABORAL' AND c.subcategoria = 'TOTAL FORMACIÓN LABORAL' THEN m.valor ELSE 0 END) as total_formacion_laboral,
    SUM(CASE WHEN c.categoria_principal = 'FORMACION PROFESIONAL INTEGRAL' THEN m.valor ELSE 0 END) as total_formacion_integral
FROM metas_cupos m
JOIN regionales r ON m.id_regional = r.id_regional
JOIN categorias_formacion c ON m.id_categoria = c.id_categoria
GROUP BY r.codigo_regional, r.nombre_regional
ORDER BY r.codigo_regional
''')

# Commit y cerrar
conn.commit()
print("\n[OK] Base de datos creada exitosamente!")
print(f"Archivo: {db_file}")

# Mostrar estadísticas
cursor.execute('SELECT COUNT(*) FROM regionales')
print(f"\nEstadisticas:")
print(f"  - Regionales: {cursor.fetchone()[0]}")
cursor.execute('SELECT COUNT(*) FROM categorias_formacion')
print(f"  - Categorias: {cursor.fetchone()[0]}")
cursor.execute('SELECT COUNT(*) FROM metas_cupos')
print(f"  - Metas de cupos: {cursor.fetchone()[0]}")
cursor.execute('SELECT COUNT(*) FROM metas_retencion')
print(f"  - Metas de retencion: {cursor.fetchone()[0]}")
cursor.execute('SELECT COUNT(*) FROM metas_certificacion')
print(f"  - Metas de certificacion: {cursor.fetchone()[0]}")

conn.close()
print("\nProceso completado!")
