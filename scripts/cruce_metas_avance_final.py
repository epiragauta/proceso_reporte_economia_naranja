import pandas as pd
import sqlite3
from pathlib import Path
import sys

# Dicrectorio raiz del proyecto
DIR_BASE = Path(__file__).resolve().parent.parent.parent

# Recibe la ruta del archivo para el avance
if len(sys.argv) < 2:
    print("Uso: python cruce_metas_avance_final.py <ruta_archivo_avance_xlsb>")
    sys.exit(1)
# Archivos de entrada
db_file = DIR_BASE / "metas" / "metas_sena_2025.db"
excel_avance = Path(sys.argv[1])
print(excel_avance)
mes = Path(excel_avance).parent.name.upper()
mes = mes.split("-")[1].split()[0]

print(f"=== CRUCE METAS VS AVANCE {mes} 2025 ===\n")


# Función para generar código DIVIPOLA
def generar_divipola(codigo_regional):
    """Genera código DIVIPOLA: código con cero a la izquierda + 000"""
    codigo_str = str(int(codigo_regional)).zfill(2)
    return f"{codigo_str}000"


# 1. OBTENER METAS DESDE BASE DE DATOS
print("1. Leyendo metas desde base de datos...")
conn = sqlite3.connect(db_file)

query_metas = """
SELECT
    r.codigo_regional,
    r.nombre_regional,
    MAX(CASE
        WHEN c.subcategoria = 'Técnico Laboral Articulación con la Media'
        THEN m.valor
        ELSE 0
    END) as meta_tecnico_articulacion,
    MAX(CASE
        WHEN c.subcategoria = 'TOTAL FORMACION TITULADA'
        THEN m.valor
        ELSE 0
    END) as meta_formacion_titulada,
    MAX(CASE
        WHEN c.subcategoria = 'TOTAL FORMACION COMPLEMENTARIA'
        THEN m.valor
        ELSE 0
    END) as meta_formacion_complementaria,
    MAX(CASE
        WHEN c.subcategoria = 'TOTAL FORMACION PROFESIONAL INTEGRAL'
        THEN m.valor
        ELSE 0
    END) as meta_fpi_total,
    MAX(CASE
        WHEN c.subcategoria = 'Total Formación Profesional Integral - Virtual'
        THEN m.valor
        ELSE 0
    END) as meta_fpi_virtual,
    MAX(CASE
        WHEN c.subcategoria = 'Total Programa de Bilingüismo'
        THEN m.valor
        ELSE 0
    END) as meta_bilinguismo
FROM
    regionales r
LEFT JOIN
    metas_cupos m ON r.id_regional = m.id_regional
LEFT JOIN
    categorias_formacion c ON m.id_categoria = c.id_categoria
WHERE
    m.anio = 2025
    AND c.subcategoria IN (
        'Técnico Laboral Articulación con la Media',
        'TOTAL FORMACION TITULADA',
        'TOTAL FORMACION COMPLEMENTARIA',
        'TOTAL FORMACION PROFESIONAL INTEGRAL',
        'Total Formación Profesional Integral - Virtual',
        'Total Programa de Bilingüismo'
    )
GROUP BY
    r.codigo_regional,
    r.nombre_regional
ORDER BY
    r.codigo_regional
"""

df_metas = pd.read_sql_query(query_metas, conn)
conn.close()

# Agregar código DIVIPOLA
df_metas["codigo_divipola"] = df_metas["codigo_regional"].apply(generar_divipola)

print(f"   {len(df_metas)} regionales cargadas")

# 2. LEER DATOS DE AVANCE
print("\n2. Leyendo avances desde archivo XLSB...")

# 2.1 TEC ARTIC REG - Columna F (TOTAL)
print("   - Leyendo TEC ARTIC REG...")
df_tec_artic = pd.read_excel(
    excel_avance, sheet_name="TEC ARTIC REG", engine="pyxlsb", header=5
)
# La columna TOTAL está en la columna 5 (índice 5)
df_tec_artic.columns = [
    "codigo_regional",
    "nombre_regional",
    "masculino",
    "femenino",
    "no_binario",
    "total",
]
df_tec_artic = df_tec_artic[["codigo_regional", "total"]].copy()
df_tec_artic = df_tec_artic[df_tec_artic["codigo_regional"].notna()]
# Convertir a numérico y filtrar filas no válidas
df_tec_artic["codigo_regional"] = pd.to_numeric(
    df_tec_artic["codigo_regional"], errors="coerce"
)
df_tec_artic = df_tec_artic[df_tec_artic["codigo_regional"].notna()].copy()
df_tec_artic["codigo_regional"] = df_tec_artic["codigo_regional"].astype(int)
df_tec_artic["total"] = (
    pd.to_numeric(df_tec_artic["total"], errors="coerce").fillna(0).astype(int)
)
df_tec_artic.rename(columns={"total": "avance_tec_articulacion"}, inplace=True)

# 2.2 NIVEL REGIONAL - Columnas AW (48), BO (66), BU (72)
print("   - Leyendo NIVEL REGIONAL...")
df_nivel = pd.read_excel(
    excel_avance, sheet_name="NIVEL REGIONAL ", engine="pyxlsb", header=6
)

# Extraer código regional (columna 0) y las columnas de interés
# AW = columna 48, BO = columna 66, BU = columna 72
df_nivel_extract = pd.DataFrame(
    {
        "codigo_regional": df_nivel.iloc[:, 0],
        "avance_formacion_titulada": df_nivel.iloc[:, 48],  # Columna AW
        "avance_formacion_complementaria": df_nivel.iloc[:, 66],  # Columna BO
        "avance_fpi_total": df_nivel.iloc[:, 72],  # Columna BU
    }
)

# Limpiar datos
df_nivel_extract = df_nivel_extract[df_nivel_extract["codigo_regional"].notna()].copy()
df_nivel_extract["codigo_regional"] = pd.to_numeric(
    df_nivel_extract["codigo_regional"], errors="coerce"
)
df_nivel_extract = df_nivel_extract[df_nivel_extract["codigo_regional"].notna()]
df_nivel_extract["codigo_regional"] = df_nivel_extract["codigo_regional"].astype(int)

# Convertir valores a numéricos
for col in [
    "avance_formacion_titulada",
    "avance_formacion_complementaria",
    "avance_fpi_total",
]:
    df_nivel_extract[col] = pd.to_numeric(
        df_nivel_extract[col], errors="coerce"
    ).fillna(0)

# 2.3 VIRTUAL GENER REG - Columna F (TOTAL)
print("   - Leyendo VIRTUAL GENER REG...")
df_virtual = pd.read_excel(
    excel_avance, sheet_name="VIRTUAL GENER REG", engine="pyxlsb", header=6
)
# Buscar las columnas correctas
df_virtual_extract = pd.DataFrame(
    {
        "codigo_regional": df_virtual.iloc[:, 0],
        "avance_fpi_virtual": df_virtual.iloc[:, 5],  # Columna F (índice 5)
    }
)
df_virtual_extract = df_virtual_extract[
    df_virtual_extract["codigo_regional"].notna()
].copy()
df_virtual_extract["codigo_regional"] = pd.to_numeric(
    df_virtual_extract["codigo_regional"], errors="coerce"
)
df_virtual_extract = df_virtual_extract[df_virtual_extract["codigo_regional"].notna()]
df_virtual_extract["codigo_regional"] = df_virtual_extract["codigo_regional"].astype(
    int
)
df_virtual_extract["avance_fpi_virtual"] = pd.to_numeric(
    df_virtual_extract["avance_fpi_virtual"], errors="coerce"
).fillna(0)

# 2.4 BILINGÜISMO REG - Columna F (TOTAL)
print("   - Leyendo BILINGÜISMO REG...")
df_bilinguismo = pd.read_excel(
    excel_avance, sheet_name="BILINGÜISMO REG", engine="pyxlsb", header=6
)
df_bilinguismo_extract = pd.DataFrame(
    {
        "codigo_regional": df_bilinguismo.iloc[:, 0],
        "avance_bilinguismo": df_bilinguismo.iloc[:, 5],  # Columna F (índice 5)
    }
)
df_bilinguismo_extract = df_bilinguismo_extract[
    df_bilinguismo_extract["codigo_regional"].notna()
].copy()
df_bilinguismo_extract["codigo_regional"] = pd.to_numeric(
    df_bilinguismo_extract["codigo_regional"], errors="coerce"
)
df_bilinguismo_extract = df_bilinguismo_extract[
    df_bilinguismo_extract["codigo_regional"].notna()
]
df_bilinguismo_extract["codigo_regional"] = df_bilinguismo_extract[
    "codigo_regional"
].astype(int)
df_bilinguismo_extract["avance_bilinguismo"] = pd.to_numeric(
    df_bilinguismo_extract["avance_bilinguismo"], errors="coerce"
).fillna(0)

print("   [OK] Todos los avances leidos")

# 3. COMBINAR DATOS
print("\n3. Combinando metas y avances...")

resultado = df_metas.copy()

# Merge con cada fuente de avance
resultado = resultado.merge(df_tec_artic, on="codigo_regional", how="left")
resultado = resultado.merge(df_nivel_extract, on="codigo_regional", how="left")
resultado = resultado.merge(df_virtual_extract, on="codigo_regional", how="left")
resultado = resultado.merge(df_bilinguismo_extract, on="codigo_regional", how="left")

# Llenar NaN con 0
cols_avance = [
    "avance_tec_articulacion",
    "avance_formacion_titulada",
    "avance_formacion_complementaria",
    "avance_fpi_total",
    "avance_fpi_virtual",
    "avance_bilinguismo",
]

for col in cols_avance:
    resultado[col] = resultado[col].fillna(0).astype(int)

# 4. CALCULAR CUPOS DISPONIBLES (META - AVANCE)
print("\n4. Calculando cupos disponibles...")

resultado["Cupos Doble Titulación"] = (
    resultado["meta_tecnico_articulacion"] - resultado["avance_tec_articulacion"]
)
resultado["Cupos en formación titulada"] = (
    resultado["meta_formacion_titulada"] - resultado["avance_formacion_titulada"]
)
resultado["Cupos en formación complementaria"] = (
    resultado["meta_formacion_complementaria"]
    - resultado["avance_formacion_complementaria"]
)
resultado["Total Cupos en formación profesional integral"] = (
    resultado["meta_fpi_total"] - resultado["avance_fpi_total"]
)
resultado["Cupos de virtualidad"] = (
    resultado["meta_fpi_virtual"] - resultado["avance_fpi_virtual"]
)
resultado["Cupos en Bilingüismo"] = (
    resultado["meta_bilinguismo"] - resultado["avance_bilinguismo"]
)

# 5. PREPARAR DATAFRAME FINAL
print("\n5. Preparando resultado final...")

# Seleccionar y renombrar columnas según especificación
df_final = resultado[
    [
        "codigo_regional",
        "nombre_regional",
        "codigo_divipola",
        "Cupos Doble Titulación",
        "Cupos en formación titulada",
        "Cupos en formación complementaria",
        "Total Cupos en formación profesional integral",
        "Cupos de virtualidad",
        "Cupos en Bilingüismo",
    ]
].copy()

# Renombrar columnas finales
df_final.columns = [
    "Código Regional",
    "Nombre de la Regional",
    "Código DIVIPOLA DANE",
    "Cupos Doble Titulación",
    "Cupos en formación titulada",
    "Cupos en formación complementaria",
    "Total Cupos en formación profesional integral",
    "Cupos de virtualidad",
    "Cupos en Bilingüismo",
]

# 6. EXPORTAR RESULTADOS
print("\n6. Exportando resultados...")

# Exportar a Excel
output_excel = DIR_BASE / "metas" / "cupos_disponibles_por_regional_2025.xlsx"
output_excel_intermedio = (
    DIR_BASE
    / "proceso_reporte_economia_naranja"
    / "MESES"
    / mes
    / "datos_intermedios"
    / "cupos_disponibles_por_regional_2025.xlsx"
)
df_final.to_excel(output_excel, index=False, sheet_name="Cupos Disponibles")
df_final.to_excel(output_excel_intermedio, index=False, sheet_name="Cupos Disponibles")

# Exportar a CSV
output_csv = DIR_BASE / "metas" / "cupos_disponibles_por_regional_2025.csv"
output_csv_intermedio = (
    DIR_BASE
    / "proceso_reporte_economia_naranja"
    / "MESES"
    / mes
    / "datos_intermedios"
    / "cupos_disponibles_por_regional_2025.csv"
)
df_final.to_csv(output_csv, index=False, encoding="utf-8-sig")
df_final.to_csv(output_csv_intermedio, index=False, encoding="utf-8-sig")

print("\n[OK] Resultados exportados:")
print(f"   - Excel: {output_excel}")
print(f"   - CSV: {output_csv}")

# 7. MOSTRAR RESULTADOS
print("\n" + "=" * 120)
print(f"CUPOS DISPONIBLES POR REGIONAL - {mes} 2025")
print("=" * 120)
print(df_final.to_string(index=False))

# 8. TOTALES
print("\n" + "=" * 120)
print("TOTALES NACIONALES DE CUPOS DISPONIBLES")
print("-" * 120)

totales = df_final.select_dtypes(include=["int64", "float64"]).sum()
print(
    f"Cupos Doble Titulación:                           {totales['Cupos Doble Titulación']:>15,}"
)
print(
    f"Cupos en formación titulada:                      {totales['Cupos en formación titulada']:>15,}"
)
print(
    f"Cupos en formación complementaria:                {totales['Cupos en formación complementaria']:>15,}"
)
print(
    f"Total Cupos en formación profesional integral:   {totales['Total Cupos en formación profesional integral']:>15,}"
)
print(
    f"Cupos de virtualidad:                             {totales['Cupos de virtualidad']:>15,}"
)
print(
    f"Cupos en Bilingüismo:                             {totales['Cupos en Bilingüismo']:>15,}"
)

print("\nProceso completado exitosamente!")
