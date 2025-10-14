# Guía de Uso: Generación del Reporte de Economía Naranja

**SENA Colombia - Procesamiento Automatizado de Reportes Mensuales (Economia Naranja)**

---

## Inicio Rápido (3 Pasos Simples)

Para usuarios experimentados, el proceso completo se resume en:

```bash
# 1. Verificar que todo esté listo
python verificar_prerequisitos.py SEPTIEMBRE

# 2. Generar el reporte completo
python generar_reporte_completo.py SEPTIEMBRE

# 3. Recoger el reporte final
# Ubicación: PROCESO_REPORTE_ECONOMIA_NARANJA\SEPTIEMBRE\datos_finales\
```

⏱ **Tiempo estimado**: 5-15 minutos dependiendo del tamaño de datos

---

## Prerequisitos

### Lista de Verificación Pre-Ejecución

Antes de ejecutar el proceso, asegúrese de cumplir con:

####  Software Instalado

- [ ] **Python 3.10 o superior** instalado
- [ ] **Dependencias Python** instaladas:
  ```bash
  pip install pandas openpyxl pyxlsb
  ```
- [ ] **Acceso a red institucional** SENA (si los archivos están en red)

####  Archivos de Entrada Disponibles

Para cada mes, debe tener estos 3 archivos en `C:\ws\sena\data\{AÑO}\{MES}\`:

1. **PE-04 Formación Nacional**
   - Nombre: `PE-04_FORMACION NACIONAL {MES} {AÑO}.xlsb`
   - Fuente: Sistema de Gestión de Formación
   - Formato: Excel Binary (.xlsb)

2. **Primer Avance Cupos de Formación**
   - Nombre: `PRIMER AVANCE CUPOS DE FORMACION {MES} {AÑO}.xlsb`
   - Fuente: Sistema de Seguimiento de Cupos
   - Formato: Excel Binary (.xlsb)

3. **Primer Avance en Aprendices**
   - Nombre: `PRIMER AVANCE EN APRENDICES {MES} {AÑO}.xlsb`
   - Fuente: Sistema de Gestión de Aprendices
   - Formato: Excel Binary (.xlsb)

####  Archivo Anual de Metas

- **Metas SENA (se usa todo el año)**
  - Ubicación: `C:\ws\sena\data\2025\09-Septiembre\`
  - Nombre: `Metas SENA 2025 V5 26092025_CLEAN.xlsx`
  - Formato: Excel (.xlsx)
  - Nota: Este archivo se comparte entre todos los meses del año

####  Espacio en Disco

- [ ] Mínimo **5 GB libres** en disco de trabajo
- [ ] Permisos de **lectura/escritura** en `C:\ws\sena\data\`

---

## Instrucciones Detalladas de Ejecución

### Paso 1: Validar Prerequisitos

**Propósito**: Verificar que todos los archivos y dependencias estén listos antes de iniciar.

**Comando**:
```bash
cd C:\ws\sena\data\PROCESO_REPORTE_ECONOMIA_NARANJA
python verificar_prerequisitos.py SEPTIEMBRE
```

**Sustituya `SEPTIEMBRE` por el mes que desea procesar**.

**Resultado esperado**:
```
======================================================================
 RESUMEN DE VERIFICACIÓN
======================================================================
   ✓ Python
   ✓ Dependencias
   ✓ Archivos de entrada
   ✓ Scripts del sistema
   ✓ Directorios base
   ✓ Espacio en disco
======================================================================

✓ TODOS LOS PREREQUISITOS ESTÁN LISTOS
```

**Si aparece ✗ (error)**:
- Revise la sección que falla
- Corrija el problema (instale dependencias, coloque archivos faltantes, etc.)
- Vuelva a ejecutar la verificación

### Paso 2: Generar el Reporte Completo

**Propósito**: Ejecutar el proceso automatizado de generación del reporte.

**Comando**:
```bash
python generar_reporte_completo.py SEPTIEMBRE
```

**Sustituya `SEPTIEMBRE` por el mes deseado**.

**¿Qué sucede durante la ejecución?**

El sistema ejecutará automáticamente 8 pasos:

1. **Creando directorios** → Organiza estructura de carpetas
2. **Copiando archivos de entrada** → Centraliza archivos fuente
3. **Generando base de datos de formación** → Importa PE-04 a SQLite
4. **Creando tabla de economía naranja** → Filtra programas relevantes
5. **Generando base de datos de metas** → Normaliza metas institucionales
6. **Calculando cupos disponibles** → META - AVANCE por regional
7. **Generando reporte de aprendices** → Consolida estadísticas de aprendices
8. **Generando reporte consolidado final** → Integra todas las fuentes

**Indicadores de progreso**:
```
======================================================================
 PASO 1/8: Crear estructura de directorios
======================================================================
✓ Directorios creados para SEPTIEMBRE
  - C:\ws\sena\data\PROCESO_REPORTE_ECONOMIA_NARANJA\SEPTIEMBRE
  - C:\ws\sena\data\PROCESO_REPORTE_ECONOMIA_NARANJA\SEPTIEMBRE\datos_intermedios
  - C:\ws\sena\data\PROCESO_REPORTE_ECONOMIA_NARANJA\SEPTIEMBRE\datos_finales
```

Cada paso mostrará:
- ✓ (marca verde) = Exitoso
- → (flecha) = En proceso
- ✗ (marca roja) = Error (el proceso se detiene)

**Resultado final esperado**:
```
======================================================================
 PROCESO COMPLETADO EXITOSAMENTE
======================================================================

✓ Reporte generado: C:\ws\sena\data\PROCESO_REPORTE_ECONOMIA_NARANJA\SEPTIEMBRE\datos_finales\Reporte Consolidado Economía Naranja Sep 2025.xlsx

⏱ Tiempo total: 287.3 segundos (4.8 minutos)

📁 Archivos intermedios en: C:\ws\sena\data\PROCESO_REPORTE_ECONOMIA_NARANJA\SEPTIEMBRE\datos_intermedios
📁 Reporte final en: C:\ws\sena\data\PROCESO_REPORTE_ECONOMIA_NARANJA\SEPTIEMBRE\datos_finales

======================================================================
```

### Paso 3: Recoger el Reporte Final

**Ubicación del archivo**:
```
C:\ws\sena\data\PROCESO_REPORTE_ECONOMIA_NARANJA\{MES}\datos_finales\
```

**Nombre del archivo**:
```
Reporte Consolidado Economía Naranja {Mes} {Año}.xlsx
```

**Ejemplo**:
```
Reporte Consolidado Economía Naranja Sep 2025.xlsx
```

**Contenido del reporte**:
- Consolidado por Regional
- Consolidado por Programa de Formación
- Consolidado por Línea Tecnológica
- Cupos disponibles (META - AVANCE)
- Estadísticas de aprendices
- Filtrado por programas de Economía Naranja

**Formato**: Excel (.xlsx) listo para distribución institucional

---

## Archivos de Entrada Necesarios por Mes

### Resumen Mensual

| Archivo | Ubicación | Frecuencia | Tamaño Aprox. |
|---------|-----------|------------|---------------|
| PE-04 Formación Nacional | `{AÑO}\{MES}\` | Mensual | 50-100 MB |
| Primer Avance Cupos | `{AÑO}\{MES}\` | Mensual | 10-30 MB |
| Primer Avance Aprendices | `{AÑO}\{MES}\` | Mensual | 20-50 MB |
| Metas SENA | `{AÑO}\09-Septiembre\` | Anual | 1-5 MB |

### Estructura de Directorio Fuente

```
C:\ws\sena\data\
└── 2025\
    ├── 01-Enero\
    │   ├── PE-04_FORMACION NACIONAL ENERO 2025.xlsb
    │   ├── PRIMER AVANCE CUPOS DE FORMACION ENERO 2025.xlsb
    │   └── PRIMER AVANCE EN APRENDICES ENERO 2025.xlsb
    ├── 02-Febrero\
    │   ├── PE-04_FORMACION NACIONAL FEBRERO 2025.xlsb
    │   ├── PRIMER AVANCE CUPOS DE FORMACION FEBRERO 2025.xlsb
    │   └── PRIMER AVANCE EN APRENDICES FEBRERO 2025.xlsb
    ├── ...
    └── 09-Septiembre\
        ├── PE-04_FORMACION NACIONAL SEPTIEMBRE 2025.xlsb
        ├── PRIMER AVANCE CUPOS DE FORMACION SEPTIEMBRE 2025.xlsb
        ├── PRIMER AVANCE EN APRENDICES SEPTIEMBRE 2025.xlsb
        └── Metas SENA 2025 V5 26092025_CLEAN.xlsx
```

**Importante**: Respete exactamente la nomenclatura de archivos y carpetas.

---

## Archivos de Salida Generados

### Directorio: datos_intermedios/

Archivos de procesamiento temporal:

| Archivo | Descripción | Formato |
|---------|-------------|---------|
| `sena_formacion_{mes}.db` | Base de datos de formación | SQLite |
| `metas_sena_2025.db` | Base de datos de metas | SQLite |
| `cupos_disponibles_por_regional_2025.xlsx` | Cálculo META - AVANCE | Excel |
| `cupos_disponibles_por_regional_2025.csv` | Cálculo META - AVANCE | CSV |
| `SENA Mensual Nacional {Mes} {Año}.xlsx` | Reporte de aprendices | Excel |
| Copias de archivos fuente | Trazabilidad | XLSB/XLSX |

### Directorio: datos_finales/

Producto final del proceso:

| Archivo | Descripción | Distribución |
|---------|-------------|--------------|
| `Reporte Consolidado Economía Naranja {Mes} {Año}.xlsx` | Reporte maestro consolidado | Nivel institucional |

---

## Resolución de Problemas Comunes

### ❌ Problema: "Mes inválido"

**Error**:
```
✗ Mes inválido: SETIEMBRE
Meses válidos: ENERO, FEBRERO, MARZO, ABRIL, MAYO, JUNIO, JULIO, AGOSTO, SEPTIEMBRE, OCTUBRE, NOVIEMBRE, DICIEMBRE
```

**Solución**:
- Escriba el mes en MAYÚSCULAS
- Use el nombre completo en español
- Verifique ortografía (ejemplo: SEPTIEMBRE no SETIEMBRE)

### ❌ Problema: "Archivo no encontrado"

**Error**:
```
✗ pe04_formacion - NO ENCONTRADO
   Ruta esperada: C:\ws\sena\data\2025\09-Septiembre\PE-04_FORMACION NACIONAL SEPTIEMBRE 2025.xlsb
```

**Solución**:
1. Verifique que el archivo existe en la ruta indicada
2. Confirme que el nombre del archivo es exacto (incluyendo mayúsculas/minúsculas)
3. Verifique permisos de lectura del archivo
4. Si el archivo está en otra ubicación, cópielo a la ruta esperada

### ❌ Problema: "Módulo no encontrado" (ModuleNotFoundError)

**Error**:
```
ModuleNotFoundError: No module named 'pandas'
```

**Solución**:
```bash
pip install pandas openpyxl pyxlsb
```

Si persiste el error:
```bash
python -m pip install --upgrade pandas openpyxl pyxlsb
```

### ❌ Problema: "Tabla ya existe"

**Error**:
```
sqlite3.OperationalError: table ECONOMIA_NARANJA_SEPTIEMBRE_2025 already exists
```

**Solución**:
Limpie los datos del mes y re-ejecute:
```bash
python limpiar_mes.py SEPTIEMBRE
python generar_reporte_completo.py SEPTIEMBRE
```

### ❌ Problema: "Sin espacio en disco"

**Error**:
```
OSError: [Errno 28] No space left on device
```

**Solución**:
1. Libere espacio en disco (mínimo 5 GB)
2. Limpie meses antiguos:
   ```bash
   python limpiar_mes.py ENERO
   python limpiar_mes.py FEBRERO
   ```
3. Elimine archivos temporales de sistema

### ❌ Problema: "Permiso denegado"

**Error**:
```
PermissionError: [Errno 13] Permission denied
```

**Solución**:
1. Cierre archivos Excel abiertos en `datos_finales/` o `datos_intermedios/`
2. Ejecute el comando con permisos de administrador
3. Verifique que tiene permisos de escritura en `C:\ws\sena\data\`

### ❌ Problema: El proceso se detiene sin mensaje claro

**Solución**:
1. Revise el último paso ejecutado en la consola
2. Verifique que los archivos de entrada no estén corruptos
3. Intente abrir manualmente los archivos .xlsb en Excel
4. Re-ejecute el script con más información:
   ```bash
   python generar_reporte_completo.py SEPTIEMBRE > log_ejecucion.txt 2>&1
   ```
5. Revise `log_ejecucion.txt` para detalles del error

---

## Calendario Mensual Sugerido

Para mantener un flujo de trabajo ordenado, sugerimos:

### Días 1-5 del mes

**Recopilación de datos**:
- Solicitar archivos de entrada al equipo de sistemas
- Verificar recepción de los 3 archivos mensuales
- Validar integridad de archivos (abrir en Excel)

### Días 5-10 del mes

**Generación del reporte**:
- Ejecutar `verificar_prerequisitos.py`
- Ejecutar `generar_reporte_completo.py`
- Revisar reporte consolidado generado

### Días 10-15 del mes

**Distribución y análisis**:
- Compartir reporte con áreas interesadas
- Realizar análisis de indicadores
- Generar presentaciones ejecutivas

### Días 15-30 del mes

**Seguimiento y planificación**:
- Toma de decisiones basada en datos
- Planificación de acciones correctivas
- Preparación para el siguiente mes

---

## Comandos de Referencia Rápida

### Verificación de Prerequisitos
```bash
python verificar_prerequisitos.py {MES}
```

### Generación de Reporte Completo
```bash
python generar_reporte_completo.py {MES}
```

### Limpieza de Datos del Mes
```bash
python limpiar_mes.py {MES}
```

### Ver Configuración del Mes
```bash
python configuracion.py {MES}
```

### Meses Válidos
```
ENERO, FEBRERO, MARZO, ABRIL, MAYO, JUNIO,
JULIO, AGOSTO, SEPTIEMBRE, OCTUBRE, NOVIEMBRE, DICIEMBRE
```

**Nota**: Siempre en MAYÚSCULAS

---

## Preguntas Frecuentes (FAQ)

### ¿Puedo ejecutar el proceso para varios meses a la vez?

No directamente. El proceso actual maneja un mes a la vez. Para procesar varios meses:

```bash
python generar_reporte_completo.py ENERO
python generar_reporte_completo.py FEBRERO
python generar_reporte_completo.py MARZO
```

### ¿Qué pasa si necesito regenerar un reporte?

Use el script de limpieza primero:
```bash
python limpiar_mes.py SEPTIEMBRE
python generar_reporte_completo.py SEPTIEMBRE
```

### ¿Puedo ejecutar el proceso en otro directorio?

Debe editar `configuracion.py` y cambiar:
```python
DIR_BASE = Path(r'C:\ws\sena\data')  # Cambiar a su ruta
```

### ¿Los archivos intermedios son necesarios después de generar el reporte?

No para el uso del reporte final, pero se recomienda conservarlos para:
- Auditoría y trazabilidad
- Re-generación rápida si hay errores menores
- Análisis detallados adicionales

### ¿Cuánto tiempo conservo los datos de meses anteriores?

Recomendaciones:
- **Reportes finales**: Todo el año (12 meses)
- **Datos intermedios**: 3 meses
- **Archivos fuente originales**: Todo el año + 2 años históricos

### ¿El proceso modifica los archivos fuente originales?

No. El proceso:
1. Copia los archivos originales a `datos_intermedios/`
2. Trabaja únicamente sobre las copias
3. Los archivos en `{AÑO}\{MES}\` permanecen intactos

### ¿Qué hago si el reporte tiene datos incorrectos?

1. Verifique que los archivos fuente sean correctos
2. Limpie el mes: `python limpiar_mes.py {MES}`
3. Reemplace archivos fuente incorrectos
4. Re-ejecute: `python generar_reporte_completo.py {MES}`

---

## Lista de Verificación Final

Antes de entregar el reporte, verifique:

- [ ] El reporte se generó sin errores
- [ ] El archivo Excel abre correctamente
- [ ] Las hojas contienen datos (no están vacías)
- [ ] Los totales son coherentes
- [ ] Los nombres de regionales son correctos
- [ ] El mes y año en el reporte son correctos
- [ ] El archivo está en `datos_finales/`

---
