# Generador de Reportes - Economía Naranja SENA

Sistema automatizado para generar reportes consolidados mensuales de programas de formación en Economía Naranja del SENA.

## Características

- **Proceso automatizado de 8 pasos** para generación de reportes
- **Interfaz interactiva** con menú fácil de usar
- **Compilable a ejecutable standalone** - no requiere Python instalado
- **Verificación de prerequisitos** antes de ejecutar
- **Gestión automática de rutas** - funciona desde cualquier ubicación

## Requisitos

### Para usar el ejecutable (usuarios finales)
- Sistema operativo Windows/Linux/Mac
- Archivos de entrada mensuales en las ubicaciones correctas
- Mínimo 5 GB de espacio libre en disco

### Para desarrollo (programadores)
- Python 3.10 o superior
- Dependencias listadas en `requirements.txt`

## Instalación

### Opción 1: Usar el ejecutable (Recomendado para usuarios)

1. Descarga y descomprime el archivo `reporte-economia-naranja-release.zip`
2. Coloca tus archivos de entrada en la estructura de carpetas adecuada
3. Ejecuta el programa:
   - **Windows**: Doble clic en `reporte-economia-naranja.exe`
   - **Linux/Mac**: `./reporte-economia-naranja` desde terminal

### Opción 2: Ejecutar desde código fuente (Desarrolladores)

```bash
# 1. Clonar o descargar el proyecto
cd proceso_reporte_economia_naranja

# 2. Crear entorno virtual (recomendado)
python -m venv .venv
source .venv/bin/activate  # En Windows: .venv\Scripts\activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar
python main.py
```

## Uso

### Modo Interactivo (con menú)

Simplemente ejecuta el programa sin argumentos:

```bash
python main.py
```

o haz doble clic en el ejecutable. Verás un menú con opciones:

```
MENÚ PRINCIPAL:
  1. Generar Reporte Completo
  2. Verificar Prerequisitos
  3. Ver Configuración de un Mes
  4. Listar Meses Disponibles
  5. Salir
```

### Modo Línea de Comandos

Para generar un reporte directamente:

```bash
python main.py SEPTIEMBRE
```

Para solo verificar prerequisitos:

```bash
python main.py --verificar OCTUBRE
```

## Estructura de Archivos

El programa espera la siguiente estructura:

```
📁 proceso_reporte_economia_naranja/
├── reporte-economia-naranja.exe    # Ejecutable compilado
├── main.py                         # Punto de entrada (desarrollo)
├── CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx
├── Nueva_Hoja_SENA.xlsx
├── scripts/                        # Scripts del sistema
│   ├── configuracion.py
│   ├── generar_reporte_completo.py
│   ├── verificar_prerequisitos.py
│   └── ...
└── MESES/                          # Resultados por mes
    ├── ENERO/
    ├── FEBRERO/
    └── ...

📁 ../2025/                         # Archivos de entrada (un nivel arriba)
├── 01-Enero/
│   ├── PE-04_FORMACION NACIONAL ENERO 2025.xlsb
│   ├── PRIMER AVANCE CUPOS DE FORMACION ENERO 2025.xlsb
│   ├── PRIMER AVANCE EN APRENDICES ENERO 2025.xlsb
│   └── Metas SENA 2025 V5 26092025_CLEAN.xlsx
├── 09-Septiembre/
│   └── ...
└── ...
```

## Proceso de Generación

El sistema ejecuta 8 pasos automáticamente:

1. **Crear estructura de directorios** - Organización del mes
2. **Copiar archivos de entrada** - Centralización de fuentes
3. **Generar BD de formación** - Conversión PE-04 a SQLite
4. **Crear tabla Economía Naranja** - Filtrado de programas
5. **Generar BD de metas** - Normalización de metas
6. **Calcular cupos disponibles** - META - AVANCE
7. **Generar reporte de aprendices** - Estadísticas mensuales
8. **Consolidar reporte final** - Excel con 3 hojas

### Resultados

Los archivos generados se guardan en:

- **Intermedios**: `MESES/{MES}/datos_intermedios/`
- **Final**: `MESES/{MES}/datos_finales/Reporte Consolidado Economía Naranja {Mes} {Año}.xlsx`

## Compilación a Ejecutable

Para desarrolladores que quieran crear el ejecutable:

```bash
# 1. Instalar dependencias incluyendo PyInstaller
pip install -r requirements.txt

# 2. Ejecutar script de compilación
python build_executable.py

# 3. El ejecutable estará en la carpeta release/
```

El script de compilación:
- Limpia builds anteriores
- Compila con PyInstaller
- Crea carpeta `release/` lista para distribuir
- Incluye todos los archivos necesarios

## Verificación de Prerequisitos

Antes de generar un reporte, puedes verificar que todo esté listo:

```bash
python main.py --verificar SEPTIEMBRE
```

Esto verifica:
- ✓ Versión de Python (>= 3.10)
- ✓ Dependencias instaladas (pandas, openpyxl, pyxlsb)
- ✓ Archivos de entrada disponibles
- ✓ Scripts del sistema presentes
- ✓ Estructura de directorios correcta
- ✓ Espacio en disco suficiente (>= 5 GB)

## Solución de Problemas

### Error: "Archivo no encontrado"
- Verifica que los archivos de entrada estén en `../2025/{MES}/`
- Ejecuta verificación de prerequisitos
- Revisa que los nombres coincidan exactamente

### Error: "Módulo no encontrado"
```bash
pip install -r requirements.txt
```

### Error: "Tabla ya existe"
Limpia los datos del mes y vuelve a ejecutar:
```bash
python scripts/limpiar_mes.py SEPTIEMBRE
```

### El ejecutable no inicia
- Ejecuta desde terminal para ver mensajes de error
- Verifica que todos los archivos de `release/` estén presentes
- En Windows: Verifica que no esté bloqueado por antivirus

## Configuración

El archivo `scripts/configuracion.py` centraliza toda la configuración:

- **Año de trabajo**: Cambiar `ANIO_TRABAJO = 2025`
- **Meses disponibles**: Diccionario `MESES`
- **Rutas base**: Auto-detectadas según entorno

## Soporte y Documentación

Documentación adicional disponible en:

- `scripts/README_PROCESO.md` - Documentación técnica detallada
- `scripts/GUIA_USO.md` - Guía de usuario paso a paso
- `scripts/DIAGRAMAS_PROCESO.md` - Diagramas de flujo del proceso

## Dependencias

### Librerías Python
- **pandas** - Procesamiento de datos tabulares
- **openpyxl** - Lectura/escritura Excel XLSX
- **pyxlsb** - Lectura Excel XLSB (formato binario)
- **sqlite3** - Base de datos (incluido en Python)
- **tqdm** - Barras de progreso
- **pyinstaller** - Compilación a ejecutable (solo desarrollo)

## Licencia

Proyecto interno SENA - Todos los derechos reservados

## Versión

1.0.0 - Diciembre 2024

---

**Desarrollado para**: SENA - Servicio Nacional de Aprendizaje
**Propósito**: Automatización de reportes de Economía Naranja
