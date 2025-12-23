# Guía de Compilación a Ejecutable

Esta guía te ayudará a compilar el proyecto en un ejecutable standalone (.exe en Windows) que puede distribuirse sin necesidad de tener Python instalado.

## Requisitos Previos

1. **Python 3.10 o superior** instalado
2. **Dependencias del proyecto** instaladas
3. **PyInstaller** instalado

## Instalación Rápida

### Windows

```cmd
# Ejecutar script de instalación automática
install.bat

# O manualmente:
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### Linux/Mac

```bash
# Ejecutar script de instalación automática
chmod +x install.sh
./install.sh

# O manualmente:
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Compilación

### Método Automático (Recomendado)

Simplemente ejecuta el script de compilación:

```bash
python build_executable.py
```

Este script:
1. Verifica que PyInstaller esté instalado
2. Limpia builds anteriores
3. Compila el proyecto con configuración optimizada
4. Verifica que el ejecutable se generó correctamente
5. Crea carpeta `release/` lista para distribuir

### Método Manual

Si prefieres compilar manualmente con PyInstaller:

```bash
pyinstaller --name reporte-economia-naranja \
            --onefile \
            --console \
            --add-data "scripts:scripts" \
            --add-data "CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx:." \
            --hidden-import pandas \
            --hidden-import openpyxl \
            --hidden-import pyxlsb \
            main.py
```

## Resultado de la Compilación

Después de compilar, encontrarás:

```
📁 dist/
└── reporte-economia-naranja.exe    # Ejecutable (Windows)
    o
    reporte-economia-naranja         # Ejecutable (Linux/Mac)

📁 release/                          # Carpeta lista para distribuir
├── reporte-economia-naranja.exe
├── README.md
├── CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx
├── Nueva_Hoja_SENA.xlsx
└── MESES/
```

## Tamaño del Ejecutable

- **Windows (.exe)**: ~50-80 MB
- **Linux**: ~50-80 MB
- **Mac**: ~60-90 MB

El tamaño es mayor porque incluye Python y todas las librerías empaquetadas.

## Distribución

### Preparar para distribuir

1. La carpeta `release/` ya está lista para distribuir
2. Comprímela en un archivo ZIP:

```bash
# Windows (PowerShell)
Compress-Archive -Path release -DestinationPath reporte-economia-naranja-v1.0.zip

# Linux/Mac
zip -r reporte-economia-naranja-v1.0.zip release/
```

3. Distribuye el archivo ZIP

### Instrucciones para usuarios finales

Incluye estas instrucciones con el ZIP:

```
INSTRUCCIONES DE USO:

1. Descomprimir el archivo ZIP en una carpeta
2. Colocar archivos de entrada mensuales en estructura correcta
3. Ejecutar reporte-economia-naranja.exe (doble clic o desde terminal)
4. Seguir el menú interactivo

No requiere instalación de Python ni dependencias.
```

## Solución de Problemas

### Error: "PyInstaller no está instalado"

```bash
pip install pyinstaller
```

### Error: "ModuleNotFoundError" al ejecutar

Agrega el módulo faltante a `hidden_imports` en `build_executable.py`:

```python
hidden_imports = [
    'pandas',
    'openpyxl',
    'pyxlsb',
    'tu_modulo_faltante',  # Agregar aquí
]
```

### El ejecutable es muy grande

Esto es normal. PyInstaller empaqueta Python completo y todas las librerías.

Alternativas:
- Usar `--onedir` en lugar de `--onefile` (genera carpeta con múltiples archivos)
- Usar UPX para comprimir el ejecutable (experimental)

### Antivirus bloquea el ejecutable

Algunos antivirus marcan ejecutables de PyInstaller como sospechosos (falso positivo).

Soluciones:
- Firma digital del ejecutable (requiere certificado)
- Agregar excepción en antivirus
- Distribuir desde fuente confiable

### Error en tiempo de ejecución

Ejecuta desde terminal para ver mensajes de error:

```bash
# Windows
.\reporte-economia-naranja.exe

# Linux/Mac
./reporte-economia-naranja
```

Los logs ayudarán a identificar el problema.

## Opciones Avanzadas de PyInstaller

### Agregar ícono personalizado

```bash
pyinstaller --icon=icon.ico ...
```

Coloca `icon.ico` en la raíz del proyecto.

### Compilar sin ventana de consola (solo Windows)

Para aplicaciones GUI (no recomendado para este proyecto):

```bash
pyinstaller --windowed ...
```

### Optimizar tamaño con UPX

```bash
pyinstaller --upx-dir=/path/to/upx ...
```

Descarga UPX desde: https://upx.github.io/

### Modo un directorio (más rápido de iniciar)

```bash
pyinstaller --onedir ...
```

Genera carpeta con múltiples archivos en lugar de un solo .exe

## Versionado

Para mantener versiones organizadas:

1. Actualiza versión en `README.md`
2. Compila con nombre versionado:

```bash
pyinstaller --name reporte-economia-naranja-v1.0 ...
```

3. Nombra el ZIP con la versión:
   - `reporte-economia-naranja-v1.0-windows.zip`
   - `reporte-economia-naranja-v1.0-linux.zip`

## Automatización con CI/CD (Avanzado)

Puedes automatizar la compilación con GitHub Actions:

```yaml
# .github/workflows/build.yml
name: Build Executable

on: [push, release]

jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v2
      - uses: actions/setup-python@v2
        with:
          python-version: '3.10'
      - run: pip install -r requirements.txt
      - run: python build_executable.py
      - uses: actions/upload-artifact@v2
        with:
          name: executable
          path: dist/
```

## Notas Importantes

1. **Siempre prueba el ejecutable** antes de distribuir
2. **Verifica en el sistema operativo objetivo** (Windows/Linux/Mac)
3. **Documenta la versión de Python** usada para compilar
4. **Guarda el entorno de compilación** para reproducibilidad

## Soporte

Si encuentras problemas durante la compilación:

1. Verifica versión de Python: `python --version`
2. Verifica versión de PyInstaller: `pyinstaller --version`
3. Revisa logs en `build/`
4. Consulta documentación oficial: https://pyinstaller.org/

---

**Última actualización**: Diciembre 2024
