# Documentación Técnica: Proceso Automatizado de Reporte de Economía Naranja

## Propósito del Proceso

Este sistema automatiza la generación del **Reporte Consolidado de Economía Naranja** del SENA Colombia. El proceso integra datos de múltiples fuentes institucionales para producir un reporte mensual consolidado que permite el seguimiento de programas de formación en áreas de economía creativa y cultural.

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
- **Script externo**: `PE-04/importar_pe_04_mes.py`
- **Tecnología**: Python + pandas + pyxlsb + sqlite3

### Paso 4: Creación de Tabla de Economía Naranja
- **Función**: Filtra programas de economía naranja mediante SQL
- **Proceso**:
  - Lee template SQL: `REPORTE_ECONOMIA_NARANJA/crear_tabla_economia_naranja.sql`
  - Aplica filtros por líneas tecnológicas de economía creativa
  - Genera tabla: `ECONOMIA_NARANJA_{MES}_{AÑO}`
- **Criterios de filtrado**: Líneas tecnológicas 11, 12, 13, 14, 15, 16
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
- **Script externo**: `metas/cruce_metas_avance_final.py`
- **Cálculo**: Cupos Disponibles = Meta Anual - Avance Acumulado

### Paso 7: Generación de Reporte de Aprendices
- **Función**: Consolida estadísticas de aprendices por regional
- **Entrada**: `PRIMER AVANCE EN APRENDICES {MES} {AÑO}.xlsb`
- **Salida**: `SENA Mensual Nacional {MES_CORTO} {AÑO}.xlsx`
- **Script externo**: `aprendices/generar_reporte_mensual_aprendices.py`
- **Métricas**: Aprendices matriculados, en formación, certificados

### Paso 8: Consolidación Final del Reporte
- **Función**: Integra todas las fuentes en un reporte Excel maestro
- **Entradas**:
  - `sena_formacion_{mes}.db` (tabla ECONOMIA_NARANJA)
  - `cupos_disponibles_por_regional_2025.xlsx`
  - `SENA Mensual Nacional {MES_CORTO} {AÑO}.xlsx`
- **Salida**: `Reporte Consolidado Economía Naranja {MES_CORTO} {AÑO}.xlsx`
- **Ubicación**: `datos_finales/`
- **Script externo**: `REPORTE_ECONOMIA_NARANJA/generar_reporte_consolidado.py`
- **Contenido**: Vistas consolidadas por regional, programa y línea tecnológica

## Arquitectura del Sistema

### Organización de Archivos

```
C:\ws\sena\data\
├── PROCESO_REPORTE_ECONOMIA_NARANJA\          # Directorio raíz del proceso
│   ├── configuracion.py                       # Configuración centralizada
│   ├── generar_reporte_completo.py           # Script maestro (orquestador)
│   ├── verificar_prerequisitos.py            # Validador de requisitos
│   ├── limpiar_mes.py                        # Limpieza de datos del mes
│   │
│   ├── {MES}\                                # Directorio por mes (ej: SEPTIEMBRE)
│   │   ├── datos_intermedios\                # Archivos de procesamiento
│   │   │   ├── *.xlsb                       # Copias de archivos fuente
│   │   │   ├── *.db                         # Bases de datos SQLite
│   │   │   ├── cupos_disponibles_*.xlsx     # Cálculos intermedios
│   │   │   └── SENA Mensual Nacional *.xlsx # Reporte de aprendices
│   │   │
│   │   └── datos_finales\                    # Producto final
│   │       └── Reporte Consolidado *.xlsx   # Reporte maestro
│   │
├── PE-04\                                     # Componente: Importación PE-04
│   └── importar_pe_04_mes.py
│
├── metas\                                     # Componente: Gestión de metas
│   ├── normalizar_metas_sena.py
│   └── cruce_metas_avance_final.py
│
├── aprendices\                                # Componente: Reportes aprendices
│   └── generar_reporte_mensual_aprendices.py
│
├── REPORTE_ECONOMIA_NARANJA\                  # Componente: Reporte consolidado
│   ├── crear_tabla_economia_naranja.sql
│   └── generar_reporte_consolidado.py
│
└── {AÑO}\{MES}\                              # Archivos fuente institucionales
    ├── PE-04_FORMACION NACIONAL *.xlsb
    ├── PRIMER AVANCE CUPOS DE FORMACION *.xlsb
    ├── PRIMER AVANCE EN APRENDICES *.xlsb
    └── Metas SENA *.xlsx
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
- `ejecutar_comando(comando, descripcion)`: Wrapper para subprocess
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
- Ejecuta scripts externos mediante subprocess
- Gestiona cambios de directorio de trabajo
- Mueve archivos generados a ubicaciones correctas
- Reporte de tiempo total de ejecución

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
