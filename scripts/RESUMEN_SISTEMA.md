# Resumen - Procesamiento Automatizado de Reportes

## SENA Colombia - Reporte de Economía Naranja

---

## Implementación

Procesos funcionales para automatizar la generación del Reporte Consolidado de Economía Naranja del SENA.

---

## Archivos Creados

### Scripts Principales

```
PROCESO_REPORTE_ECONOMIA_NARANJA/
├── configuracion.py                   Configuración centralizada
├── verificar_prerequisitos.py         Validador de requisitos
├── generar_reporte_completo.py        Script maestro (orquestador)
├── limpiar_mes.py                    Limpieza de datos
├── README_PROCESO.md                  Documentación técnica
├── GUIA_USO.md                        Guía de usuario
└── RESUMEN_SISTEMA.md                 Este archivo
```

### Documentación Técnica (docs/)

```
docs/
├── 01_VISION_GENERAL_SISTEMA.md           Visión, alcance, stakeholders
├── 02_ARQUITECTURA_SISTEMA.md             Arquitectura técnica completa
├── 03_MODULO_METAS.md                     BD normalizada de metas
├── 04_MODULO_APRENDICES.md                Catálogo maestro municipios
├── 05_MODULO_ECONOMIA_NARANJA.md          Programas filtrados
├── 06_REPORTE_CONSOLIDADO.md              Integración 3 hojas
├── 07_CASOS_DE_USO.md                     12 casos de uso
├── 08_REQUERIMIENTOS_NO_FUNCIONALES.md    Performance, calidad
├── 09_DICCIONARIO_DATOS.md                DIVIPOLA, categorías
└── 10_MANUAL_OPERACION.md                 Operación mensual
```

---

##  Características del Procedimiento

### Un Solo Comando

```bash
python generar_reporte_completo.py SEPTIEMBRE
```

### 8 Pasos Automatizados

1.  Crear estructura de directorios
2.  Copiar archivos de entrada
3.  Generar BD de formación (PE-04)
4.  Crear tabla de economía naranja
5.  Generar BD de metas
6.  Calcular cupos disponibles (META - AVANCE)
7.  Generar reporte de aprendices
8.  Consolidar reporte final (3 hojas)

### Organización por Mes

```
{MES}/
├── datos_intermedios/    # BDs, cálculos, copias fuente
└── datos_finales/        # Reporte consolidado final
```

---

## 📊 Reporte Final Generado

**Nombre**: `Reporte Consolidado Economía Naranja {Mes} {Año}.xlsx`

**Estructura**:
- **Hoja 1**: Economía Naranja SEPTIEMBR 2025
  - 3,998 programas
  - Header azul #1F4E78
  - Fuente: Tabla ECONOMIA_NARANJA_SEPTIEMBRE_2025

- **Hoja 2**: Oferta Disponible SEPTIEMB 2025
  - 33 regionales
  - Header verde #548235
  - Cálculo: META - AVANCE
  - Fuente: cupos_disponibles_por_regional_2025.xlsx

- **Hoja 3**: SENA Mensual Nacional Sep 2025
  - 1,143 registros (33 deptos + 1,110 municipios)
  - Header verde #548235
  - Departamentos: Fondo naranja #F4B084 + negrita
  - Fuente: Reporte mensual de aprendices
  - 0% nombres vacíos (catálogo maestro)

---

## Uso Rápido

### Paso 1: Verificar
```bash
python verificar_prerequisitos.py SEPTIEMBRE
```

### Paso 2: Generar
```bash
python generar_reporte_completo.py SEPTIEMBRE
```

### Paso 3: Ubicación Reporte
```
Ubicación:
SEPTIEMBRE/datos_finales/Reporte Consolidado Economía Naranja Sep 2025.xlsx
```

---

## Archivos de Entrada Necesarios

Por cada mes, coloque en `C:\ws\sena\data\{AÑO}\{MES}\`:

1. `PE-04_FORMACION NACIONAL {MES} {AÑO}.xlsb`
2. `PRIMER AVANCE CUPOS DE FORMACION {MES} {AÑO}.xlsb`
3. `PRIMER AVANCE EN APRENDICES {MES} {AÑO}.xlsb`

**Archivo anual** (se usa todo el año):
- `C:\ws\sena\data\2025\09-Septiembre\Metas SENA 2025 V5 26092025_CLEAN.xlsx`

---
---

## 🔧 Comandos de Mantenimiento

### Verificar configuración de un mes
```bash
python configuracion.py SEPTIEMBRE
```

### Limpiar datos para re-generar
```bash
python limpiar_mes.py SEPTIEMBRE
```

### Ver versión de Python y dependencias
```bash
python --version
pip list | grep -E "pandas|openpyxl|pyxlsb"
```

---

##  Documentación Disponible

### Para Usuarios
- **GUIA_USO.md** - Instrucciones paso a paso, FAQ, troubleshooting
- **RESUMEN_SISTEMA.md** - Este documento (quick reference)

### Para Técnicos
- **README_PROCESO.md** - Documentación técnica detallada
- **docs/** - 10 documentos de especificación de requerimientos

### Para Desarrollo
- **configuracion.py** - Código con docstrings completos
- **generar_reporte_completo.py** - Script maestro comentado
- Scripts individuales con documentación inline

---

##  Mejoras Implementadas

### Corrección de Problemas Previos

1.  **Hoja 2 ahora muestra cupos disponibles** (META - AVANCE)
   - Antes: Solo mostraba metas
   - Ahora: Calcula diferencial real por regional

2.  **Catálogo maestro de municipios**
   - Antes: Municipios con nombres vacíos
   - Ahora: 0% vacíos (busca en todas las 5 hojas)

3.  **Formato de headers institucional**
   - Hoja 1: Azul #1F4E78
   - Hojas 2 y 3: Verde #548235
   - Departamentos en hoja 3: Naranja #F4B084

4.  **Validación de 31 caracteres en nombres de hojas**
   - Auto-trunca nombres largos
   - Mantiene coherencia: "SEPTIEMBRE" → "SEPTIEMBR"

---

##  Capacitación

### Nivel Básico (Usuarios)
1. Leer **GUIA_USO.md**
2. Practicar con un mes de prueba
3. Ejecutar verificar_prerequisitos.py
4. Generar primer reporte

### Nivel Intermedio (Administradores)
1. Comprender flujo de 8 pasos
2. Leer **README_PROCESO.md**
3. Conocer scripts intermedios
4. Manejar errores comunes

### Nivel Avanzado (Desarrolladores)
1. Estudiar arquitectura en **docs/**
2. Modificar configuracion.py para nuevos años
3. Adaptar scripts para cambios en estructura de datos
4. Contribuir mejoras al sistema

---


## Siguiente Paso

```bash
# Prueba el sistema ahora:
cd C:\ws\sena\data\PROCESO_REPORTE_ECONOMIA_NARANJA
python verificar_prerequisitos.py SEPTIEMBRE
python generar_reporte_completo.py SEPTIEMBRE
```
