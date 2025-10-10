-- ====================================================================
-- CREAR TABLA: ECONOMIA_NARANJA_SEPTIEMBRE_2025
-- ====================================================================
--
-- Esta consulta crea una tabla permanente con los datos de programas de
-- economía naranja agrupados por nivel de formación y departamento
--
-- Estructura de la tabla:
-- - CODIGO_NIVEL_FORMACION: Código del nivel de formación
-- - nombre_departamento: Nombre del departamento donde se dicta
-- - nombre_programa_formacion: Nombre del programa de formación
-- - COMPLEMENTARIA: Aprendices en niveles 8 y 9 (Curso Especial y Evento)
-- - TITULADA: Aprendices en otros niveles (Auxiliar, Técnico, Tecnólogo, etc.)
-- - TOTAL: Total de aprendices
-- - fecha_creacion: Timestamp de cuando se creó el registro
-- ====================================================================

-- Eliminar la tabla si existe
DROP TABLE IF EXISTS ECONOMIA_NARANJA_SEPTIEMBRE_2025;

-- Crear la tabla con los datos de economía naranja
CREATE TABLE ECONOMIA_NARANJA_SEPTIEMBRE_2025 AS
WITH programas_eco_naranja AS (
    -- Crear dataset con fichas que corresponden a programas de economía naranja
    SELECT
        f.codigo_nivel_formacion,
        f.total_aprendices,
        u.nombre_departamento as nombre_departamento,
        p.nombre_programa
    FROM fichas f
    -- Join con ubicaciones para obtener nombre del departamento
    JOIN ubicaciones u ON f.codigo_pais_curso = u.codigo_pais
                       AND f.codigo_departamento_curso = u.codigo_departamento
                       AND f.codigo_municipio_curso = u.codigo_municipio
    -- Join con programas para obtener nombre del programa
    JOIN programas p ON f.codigo_programa = p.codigo_programa
                     AND f.version_programa = p.version_programa
    -- Filtrar solo programas que existen en el catálogo de economía naranja
    -- La clave de matching es la concatenación de CODIGO_PROGRAMA + VERSION_PROGRAMA
    WHERE EXISTS (
        SELECT 1
        FROM programas_economia_naranja pen
        WHERE CAST(pen.codigo AS TEXT) || CAST(pen.version AS TEXT) =
              CAST(f.codigo_programa AS TEXT) || CAST(f.version_programa AS TEXT)
    )
    -- Excluir programas especiales específicos
    AND f.codigo_programa NOT IN (1013, 2176, 2295, 2315, 2317, 2375, 2356, 2357, 2377, 2316, 2335, 2258, 2395)
)

SELECT
    pen.nombre_departamento,
    pen.nombre_programa,
    -- COMPLEMENTARIA: Suma de aprendices en niveles 8 (Curso Especial) y 9 (Evento)
    SUM(CASE WHEN pen.codigo_nivel_formacion IN (8, 9) THEN pen.total_aprendices ELSE 0 END) AS COMPLEMENTARIA,
    -- TITULADA: Suma de aprendices en todos los otros niveles (1=Auxiliar, 2=Técnico, 6=Tecnólogo, 10=Operario, 223=Profundización)
    SUM(CASE WHEN pen.codigo_nivel_formacion NOT IN (8, 9) THEN pen.total_aprendices ELSE 0 END) AS TITULADA,
    -- TOTAL: Suma total de aprendices
    SUM(pen.total_aprendices) AS TOTAL,
    -- Agregar timestamp de creación
    CURRENT_TIMESTAMP AS fecha_creacion
FROM programas_eco_naranja pen
GROUP BY
    pen.nombre_departamento,
    pen.nombre_programa
ORDER BY
    pen.nombre_departamento,
    pen.nombre_programa;

-- Crear índices para optimizar consultas
CREATE INDEX IF NOT EXISTS idx_economia_naranja_nivel ON ECONOMIA_NARANJA_SEPTIEMBRE_2025(CODIGO_NIVEL_FORMACION);
CREATE INDEX IF NOT EXISTS idx_economia_naranja_departamento ON ECONOMIA_NARANJA_SEPTIEMBRE_2025(nombre_departamento);
CREATE INDEX IF NOT EXISTS idx_economia_naranja_programa ON ECONOMIA_NARANJA_SEPTIEMBRE_2025(nombre_programa_formacion);
CREATE INDEX IF NOT EXISTS idx_economia_naranja_total ON ECONOMIA_NARANJA_SEPTIEMBRE_2025(TOTAL);

-- Mostrar estadísticas de la tabla creada
SELECT 'Tabla ECONOMIA_NARANJA_SEPTIEMBRE_2025 creada exitosamente' as mensaje;
SELECT COUNT(*) as total_registros FROM ECONOMIA_NARANJA_SEPTIEMBRE_2025;
SELECT SUM(COMPLEMENTARIA) as total_complementaria, SUM(TITULADA) as total_titulada, SUM(TOTAL) as gran_total FROM ECONOMIA_NARANJA_SEPTIEMBRE_2025;