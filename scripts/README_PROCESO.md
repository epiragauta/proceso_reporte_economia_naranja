# Documentación Técnica: Proceso Automatizado de Reporte de Economía Naranja

## Propósito del Proceso

Este procedimiento automatiza la generación del **Reporte Consolidado de Economía Naranja**. El proceso integra datos de múltiples fuentes institucionales para producir un reporte mensual consolidado que permite el seguimiento de programas de formación en áreas de economía naranja.

El sistema centraliza la gestión de datos de formación, metas institucionales y avance de aprendices, reduciendo el tiempo de generación de reportes de varios días a minutos y eliminando errores manuales en el procesamiento de información.

## Flujo de Trabajo: 8 Pasos Automatizados

El proceso ejecuta de manera secuencial los siguientes pasos:

### Paso 1: Creación de Estructura de Directorios
- **Función**: Establece la organización de archivos para el mes
- **Directorios creados**:
  - `PROCESO_REPORTE_ECONOMIA_NARANJA/{MES}/`
  - `PROCESO_REPORTE_ECONOMIA_NARANJA/{MES}/datos_intermedios/`
  - `PROCESO_REPORTE_ECONOMIA_NARANJA/{MES}/datos_finales/`
- **Script responsable**: `configuracion.py::crear_directorios_mes()`

### Paso 2: Copia de Archivos de Entrada
- **Función**: Centraliza los archivos fuente en el directorio del mes
- **Archivos copiados**:
  - PE-04 Formación Nacional (XLSB)
  - Primer Avance Cupos de Formación (XLSB)
  - Primer Avance en Aprendices (XLSB)
  - Metas SENA (XLSX - anual)
- **Destino**: `datos_intermedios/`
- **Script responsable**: `generar_reporte_completo.py::paso_2_copiar_archivos_entrada()`

### Paso 3: Generación de Base de Datos de Formación
- **Función**: Convierte el archivo PE-04 a formato SQLite para procesamiento eficiente
- **Entrada**: `PE-04_FORMACION NACIONAL {MES} {AÑO}.xlsb`
- **Salida**: `sena_formacion_{mes}.db`
- **Script externo**: `importar_pe_04_mes.py`
- **Parámetros requeridos**:
  - Directorio de trabajo (donde se encuentra el archivo PE-04)
  - Mes del archivo (ejemplo: SEPTIEMBRE)
- **Procesamiento adicional**:
  - Carga automática del catálogo de programas de economía naranja desde `CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx`
  - Crea tabla `programas_economia_naranja` con código, versión y nombre de programa
  - Búsqueda inteligente del catálogo en múltiples ubicaciones
- **Tecnología**: Python + pandas + pyxlsb + sqlite3

### Paso 4: Creación de Tabla de Economía Naranja
- **Función**: Filtra programas de economía naranja mediante SQL usando catálogo precargado
- **Proceso**:
  - Lee template SQL: `crear_tabla_economia_naranja.sql`
  - Aplica join con tabla `programas_economia_naranja` (cargada en Paso 3)
  - Filtra fichas activas de programas de economía creativa
  - Genera tabla precalculada: `ECONOMIA_NARANJA_{MES}_{AÑO}`
- **Criterios de filtrado**:
  - Programas listados en catálogo de economía naranja (783 programas)
  - Estados de curso: EJECUCION, POR INICIAR
  - Cruce por código y versión de programa
- **Script responsable**: `generar_reporte_completo.py::paso_4_crear_tabla_economia_naranja()`

### Paso 5: Generación de Base de Datos de Metas
- **Función**: Normaliza y estructura las metas institucionales
- **Entrada**: `Metas SENA 2025 V5 26092025_CLEAN.xlsx`
- **Salida**: `metas_sena_2025.db`
- **Script externo**: `metas/normalizar_metas_sena.py`
- **Procesamiento**: Extracción y normalización de metas por regional y línea de formación

### Paso 6: Cálculo de Cupos Disponibles
- **Función**: Calcula diferencial META - AVANCE por regional
- **Entradas**:
  - `metas_sena_2025.db` (metas institucionales)
  - `PRIMER AVANCE CUPOS DE FORMACION {MES} {AÑO}.xlsb` (ejecución real)
- **Salidas**:
  - `cupos_disponibles_por_regional_2025.xlsx`
  - `cupos_disponibles_por_regional_2025.csv`
- **Ubicación salida**: Los archivos generados se copian de `metas/` a `datos_intermedios/`
- **Script externo**: `metas/cruce_metas_avance_final.py`
- **Cálculo**: Cupos Disponibles = Meta Anual - Avance Acumulado
- **Mejora**: Copia archivos usando rutas absolutas para evitar problemas de directorio de trabajo

### Paso 7: Generación de Reporte de Aprendices
- **Función**: Consolida estadísticas de aprendices por regional
- **Entrada**: `PRIMER AVANCE EN APRENDICES {MES} {AÑO}.xlsb`
- **Salida**: `SENA Mensual Nacional {MES_CORTO} {AÑO}.xlsx`
- **Ubicación salida**: El archivo generado se copia de `aprendices/` a `datos_intermedios/`
- **Script externo**: `aprendices/generar_reporte_mensual_aprendices.py`
- **Métricas**: Aprendices matriculados, en formación, certificados
- **Mejora**: Copia archivo usando ruta absoluta para evitar problemas de directorio de trabajo

### Paso 8: Consolidación Final del Reporte
- **Función**: Integra todas las fuentes en un reporte Excel maestro de 3 hojas
- **Entradas** (desde `datos_intermedios/`):
  - `sena_formacion_{mes}.db` (tabla precalculada ECONOMIA_NARANJA_{MES}_{AÑO})
  - `cupos_disponibles_por_regional_2025.xlsx`
  - `SENA Mensual Nacional {MES_CORTO} {AÑO}.xlsx`
- **Salida**: `Reporte Consolidado Economía Naranja {MES_CORTO} {AÑO}.xlsx`
- **Ubicación salida**: `datos_finales/`
- **Script externo**: `SCRIPTS/generar_reporte_consolidado.py`
- **Contenido del reporte**:
  - Hoja 1: Datos de Economía Naranja (3998 registros)
  - Hoja 2: Oferta Disponible por Regional (33 regionales)
  - Hoja 3: SENA Mensual Nacional (copiada con formato preservado)
- **Variables de entorno utilizadas**:
  - `BD_FORMACION`: Ruta a base de datos de formación
  - `CUPOS_DISPONIBLES`: Ruta a archivo de cupos
  - `REPORTE_APRENDICES`: Ruta a reporte de aprendices
  - `ARCHIVO_SALIDA`: Ruta del reporte final
  - `MES_TRABAJO`, `MES_CORTO`, `ANIO`: Parámetros del mes
- **Mejora**: Script completamente parametrizable mediante variables de entorno, sin rutas hardcodeadas

## Arquitectura del Sistema

### Organización de Archivos

```
C:\ws\sena\data\
├── PROCESO_REPORTE_ECONOMIA_NARANJA\          # Directorio raíz del proceso
│   ├── scripts\                               # Scripts del sistema
│   │   ├── configuracion.py                   # Configuración centralizada
│   │   ├── generar_reporte_completo.py       # Script maestro (orquestador)
│   │   ├── importar_pe_04_mes.py             # Importación PE-04 a SQLite
│   │   ├── crear_tabla_economia_naranja.sql  # Template SQL para filtrado
│   │   ├── verificar_prerequisitos.py        # Validador de requisitos
│   │   ├── limpiar_mes.py                    # Limpieza de datos del mes
│   │   └── README_PROCESO.md                 # Este documento
│   │
│   └── MESES\                                 # Directorios de procesamiento mensual
│       └── {MES}\                             # Directorio por mes (ej: SEPTIEMBRE)
│           ├── datos_intermedios\             # Archivos de procesamiento
│           │   ├── *.xlsb                    # Copias de archivos fuente
│           │   ├── sena_formacion_{mes}.db   # Base de datos de formación
│           │   ├── metas_sena_2025.db        # Base de datos de metas
│           │   ├── cupos_disponibles_*.xlsx  # Cálculos intermedios
│           │   └── SENA Mensual Nacional *.xlsx # Reporte de aprendices
│           │
│           └── datos_finales\                 # Producto final
│               └── Reporte Consolidado *.xlsx # Reporte maestro
│
├── metas\                                     # Componente: Gestión de metas
│   ├── normalizar_metas_sena.py              # Normalización de metas
│   └── cruce_metas_avance_final.py           # Cálculo cupos disponibles
│
├── aprendices\                                # Componente: Reportes aprendices
│   └── generar_reporte_mensual_aprendices.py # Consolidado de aprendices
│
├── REPORTE_ECONOMIA_NARANJA\                  # Componente: Reporte consolidado
│   ├── generar_reporte_consolidado.py        # Generador del reporte final
│   └── CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx # Catálogo oficial (783 programas)
│
└── {AÑO}\{MES}\                              # Archivos fuente institucionales
    ├── PE-04_FORMACION NACIONAL *.xlsb
    ├── PRIMER AVANCE CUPOS DE FORMACION *.xlsb
    ├── PRIMER AVANCE EN APRENDICES *.xlsb
    ├── Metas SENA *.xlsx
    └── CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx # Catálogo de programas
```

### Organización de Datos

#### datos_intermedios/
Contiene archivos de procesamiento temporal:
- **Bases de datos SQLite**: Formato optimizado para consultas y transformaciones
- **Copias de archivos fuente**: Trazabilidad y versionado
- **Reportes intermedios**: Productos de pasos previos al consolidado

#### datos_finales/
Contiene exclusivamente el producto final:
- **Reporte Consolidado**: Archivo Excel listo para distribución institucional

## Scripts del Sistema

### 1. configuracion.py
**Propósito**: Gestión centralizada de configuración

**Funciones principales**:
- `obtener_config_mes(mes_nombre)`: Retorna diccionario de configuración completo
- `validar_archivos_entrada(config)`: Verifica existencia de archivos fuente
- `crear_directorios_mes(config)`: Crea estructura de carpetas
- `imprimir_config(config)`: Visualización de configuración

**Configuración incluida**:
- Rutas de directorios base
- Mapeo de meses (nombre, abreviatura, número)
- Año de trabajo
- Nombres de archivos de entrada/salida
- Rutas de scripts externos
- Nombres de tablas de base de datos

**Variables clave**:
```python
DIR_BASE = C:\ws\sena\data
ANIO_TRABAJO = 2025
MESES = {'ENERO': {'corto': 'Ene', 'numero': 1}, ...}
```

### 2. verificar_prerequisitos.py
**Propósito**: Validación pre-ejecución del sistema

**Verificaciones realizadas**:
1. **Versión Python**: >= 3.10
2. **Dependencias Python**:
   - pandas (manipulación de datos)
   - openpyxl (Excel XLSX)
   - pyxlsb (Excel XLSB)
   - sqlite3 (base de datos)
3. **Archivos de entrada**: Existencia y tamaño
4. **Scripts del sistema**: Disponibilidad de componentes
5. **Directorios base**: Estructura del proyecto
6. **Espacio en disco**: Mínimo 5 GB recomendado

**Uso**:
```bash
python verificar_prerequisitos.py SEPTIEMBRE
```

**Salida**: Reporte completo de estado OK/FALTANTE

### 3. generar_reporte_completo.py
**Propósito**: Script maestro que orquesta todo el proceso

**Arquitectura**:
- **Orquestador**: Ejecuta los 8 pasos secuencialmente
- **Gestión de errores**: Detiene el proceso ante cualquier fallo
- **Logging detallado**: Información paso a paso con timestamps
- **Validación de salidas**: Verifica generación correcta de archivos

**Funciones auxiliares**:
- `log_paso(numero, total, mensaje)`: Formato de mensajes
- `ejecutar_comando(comando, descripcion, check=True, env=None)`: Wrapper mejorado para subprocess
  - Soporte para variables de entorno personalizadas
  - Codificación UTF-8 con manejo de errores
  - Logging detallado de STDOUT y STDERR en caso de fallo
  - Muestra salida completa para diagnóstico de errores
- `copiar_archivo(origen, destino, descripcion)`: Copia con validación

**Pasos implementados**:
- `paso_1_crear_directorios()`
- `paso_2_copiar_archivos_entrada()`
- `paso_3_generar_bd_formacion()`
- `paso_4_crear_tabla_economia_naranja()`
- `paso_5_generar_bd_metas()`
- `paso_6_calcular_cupos_disponibles()`
- `paso_7_generar_reporte_aprendices()`
- `paso_8_generar_reporte_consolidado()`

**Uso**:
```bash
python generar_reporte_completo.py SEPTIEMBRE
```

**Características**:
- Ejecuta scripts externos mediante subprocess con variables de entorno
- Gestiona cambios de directorio de trabajo de forma segura
- Copia archivos generados a `datos_intermedios/` usando rutas absolutas
- Pasa parámetros a scripts mediante variables de entorno (Pasos 5, 6, 7, 8)
- Pasa parámetros a scripts mediante argumentos de línea de comandos (Paso 3)
- Reporte de tiempo total de ejecución
- Manejo robusto de errores con logs detallados

### 4. limpiar_mes.py
**Propósito**: Limpieza de archivos generados para re-ejecución

**Funcionalidad**:
- Elimina `datos_intermedios/` completo
- Elimina `datos_finales/` completo
- Mantiene intactos archivos fuente originales
- Solicita confirmación explícita ("SI")

**Uso**:
```bash
python limpiar_mes.py SEPTIEMBRE
```

**Advertencias**:
- Operación irreversible
- Requiere confirmación manual
- Reporta cantidad de archivos eliminados

## Dependencias Técnicas

### Software Requerido

| Componente | Versión Mínima | Propósito |
|------------|----------------|-----------|
| Python | 3.10+ | Lenguaje de ejecución |
| pandas | Última estable | Manipulación de datos tabulares |
| openpyxl | Última estable | Lectura/escritura Excel XLSX |
| pyxlsb | Última estable | Lectura Excel XLSB (formato binario) |
| sqlite3 | Incluido en Python | Base de datos relacional |

### Instalación de Dependencias

```bash
pip install pandas openpyxl pyxlsb
```

Nota: `sqlite3` viene incluido con Python 3.10+

### Requisitos del Sistema

- **Sistema Operativo**: Windows (rutas configuradas para Windows)
- **Memoria RAM**: Mínimo 8 GB recomendado
- **Espacio en disco**: 5 GB libres por mes procesado
- **Procesador**: Multi-core recomendado para pandas

## Convenciones y Estándares

### Nomenclatura de Archivos

**Archivos de entrada**:
- Formato: `{TIPO}_{DESCRIPCION} {MES} {AÑO}.{ext}`
- Ejemplo: `PE-04_FORMACION NACIONAL SEPTIEMBRE 2025.xlsb`

**Bases de datos**:
- Formato: `{sistema}_{tipo}_{mes}.db`
- Ejemplo: `sena_formacion_septiembre.db`

**Reportes finales**:
- Formato: `Reporte Consolidado Economía Naranja {Mes} {Año}.xlsx`
- Ejemplo: `Reporte Consolidado Economía Naranja Sep 2025.xlsx`

### Convenciones de Código

- **Encoding**: UTF-8 con BOM para compatibilidad Windows
- **Estilo Python**: PEP 8
- **Documentación**: Docstrings en español
- **Manejo de rutas**: `pathlib.Path` para portabilidad
- **Variables de entorno**: Comunicación entre scripts

## Mantenimiento y Soporte

### Logs del Sistema

El script `generar_reporte_completo.py` genera logs detallados:
- Hora de inicio/fin de cada paso
- Comandos ejecutados
- Errores y excepciones
- Tiempo total de procesamiento
- Ubicaciones de archivos generados

### Resolución de Problemas Comunes

**Error: "Archivo no encontrado"**
- Ejecutar: `python verificar_prerequisitos.py {MES}`
- Verificar que archivos fuente estén en `{AÑO}\{MES}\`

**Error: "Módulo no encontrado"**
- Ejecutar: `pip install pandas openpyxl pyxlsb`

**Error: "Tabla ya existe"**
- Ejecutar: `python limpiar_mes.py {MES}`
- Re-ejecutar proceso completo

**Error: "Sin espacio en disco"**
- Liberar espacio (mínimo 5 GB)
- Ejecutar limpieza de meses antiguos

### Actualización Anual

Para cambiar de año (ejemplo: 2025 → 2026):

1. Editar `configuracion.py`:
   ```python
   ANIO_TRABAJO = 2026
   ```

2. Actualizar ruta del archivo de metas:
   ```python
   'metas_sena': DIR_BASE / '2026' / 'XX-Mes' / 'Metas SENA 2026.xlsx'
   ```

3. Verificar estructura de directorios en `{AÑO}\{MES}\`

## Contacto y Soporte

Para soporte técnico o modificaciones al sistema:
- Revisar logs detallados de ejecución
- Verificar prerequisitos con script de validación
- Documentar errores con capturas de pantalla
- Contactar al equipo de desarrollo institucional SENA

---

**Versión del documento**: 1.0
**Fecha**: Octubre 2025
**Institución**: SENA Colombia
**Sistema**: Proceso Automatizado de Reporte de Economía Naranja
