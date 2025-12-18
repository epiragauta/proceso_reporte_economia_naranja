# Diagramas del Proceso Automatizado de Reporte de Economía Naranja

## Diagrama de Flujo del Proceso Completo

```mermaid
flowchart TD
    Start([Inicio: generar_reporte_completo.py]) --> Validate{Verificar<br/>prerequisitos}

    Validate -->|Error| EndError([Fin: Error de validación])
    Validate -->|OK| Step1[Paso 1: Crear estructura<br/>de directorios]

    Step1 --> Step2[Paso 2: Copiar archivos<br/>de entrada a datos_intermedios]

    Step2 --> Step3[Paso 3: Generar BD de formación<br/>importar_pe_04_mes.py]
    Step3 --> Step3a[Cargar catálogo de<br/>programas economía naranja]
    Step3a --> Step3b[Convertir PE-04 XLSB<br/>a SQLite]

    Step3b --> Step4[Paso 4: Crear tabla<br/>economía naranja]
    Step4 --> Step4a[Ejecutar SQL template:<br/>crear_tabla_economia_naranja.sql]
    Step4a --> Step4b[Filtrar 783 programas<br/>de economía creativa]

    Step4b --> Step5[Paso 5: Generar BD de metas<br/>normalizar_metas_sena.py]
    Step5 --> Step5a[Normalizar metas<br/>institucionales]

    Step5a --> Step6[Paso 6: Calcular cupos<br/>disponibles]
    Step6 --> Step6a[cruce_metas_avance_final.py:<br/>META - AVANCE]
    Step6a --> Step6b[Generar XLSX y CSV de<br/>cupos disponibles]
    Step6b --> Step6c[Copiar archivos a<br/>datos_intermedios]

    Step6c --> Step7[Paso 7: Generar reporte<br/>de aprendices]
    Step7 --> Step7a[generar_reporte_mensual_aprendices.py:<br/>Consolidar estadísticas]
    Step7a --> Step7b[Copiar reporte a<br/>datos_intermedios]

    Step7b --> Step8[Paso 8: Consolidación final]
    Step8 --> Step8a[generar_reporte_consolidado.py:<br/>Integrar 3 fuentes]
    Step8a --> Step8b{Validar<br/>archivos generados}

    Step8b -->|Error| EndError2([Fin: Error en generación])
    Step8b -->|OK| Step8c[Generar Excel con 3 hojas:<br/>1. Economía Naranja<br/>2. Oferta por Regional<br/>3. SENA Mensual]

    Step8c --> Step8d[Guardar reporte en<br/>datos_finales/]

    Step8d --> Success{Proceso<br/>exitoso?}
    Success -->|Si| EndSuccess([Fin: Reporte generado exitosamente])
    Success -->|No| EndError3([Fin: Error en proceso])

    style Start fill:#90EE90
    style EndSuccess fill:#90EE90
    style EndError fill:#FFB6C1
    style EndError2 fill:#FFB6C1
    style EndError3 fill:#FFB6C1
    style Step8c fill:#87CEEB
```

## Diagrama de Secuencia del Sistema

```mermaid
sequenceDiagram
    actor Usuario
    participant GRC as generar_reporte_completo.py<br/>(Orquestador)
    participant Config as configuracion.py
    participant FS as Sistema de Archivos
    participant PE04 as importar_pe_04_mes.py
    participant SQL as crear_tabla_economia_naranja.sql
    participant Metas as Componente Metas
    participant Aprendices as Componente Aprendices
    participant Consolidado as generar_reporte_consolidado.py
    participant BD as Bases de Datos SQLite

    Usuario->>GRC: Ejecutar proceso para MES

    rect rgb(240, 248, 255)
        Note over GRC,Config: PASO 1: Inicialización
        GRC->>Config: obtener_config_mes(MES)
        Config-->>GRC: Configuración completa
        GRC->>Config: validar_archivos_entrada()
        Config->>FS: Verificar archivos fuente
        FS-->>Config: Archivos OK
        Config-->>GRC: Validación exitosa
        GRC->>Config: crear_directorios_mes()
        Config->>FS: Crear MESES/{MES}/datos_intermedios/
        Config->>FS: Crear MESES/{MES}/datos_finales/
        FS-->>Config: Directorios creados
    end

    rect rgb(255, 250, 240)
        Note over GRC,FS: PASO 2: Copiar Archivos de Entrada
        GRC->>FS: Copiar PE-04 XLSB
        GRC->>FS: Copiar Avance Cupos XLSB
        GRC->>FS: Copiar Avance Aprendices XLSB
        GRC->>FS: Copiar Metas XLSX
        FS-->>GRC: Archivos copiados a datos_intermedios/
    end

    rect rgb(240, 255, 240)
        Note over GRC,BD: PASO 3: Generación BD Formación
        GRC->>PE04: subprocess.run(directorio, MES)
        PE04->>FS: Buscar CATALOGO_PROGRAMAS_ECONOMIA_NARANJA.xlsx
        FS-->>PE04: Catálogo encontrado
        PE04->>BD: Crear tabla programas_economia_naranja
        BD-->>PE04: Tabla creada (783 programas)
        PE04->>FS: Leer PE-04 XLSB
        FS-->>PE04: Datos PE-04
        PE04->>BD: Importar a SQLite: sena_formacion_MES.db
        BD-->>PE04: Importación completa
        PE04-->>GRC: BD generada exitosamente
    end

    rect rgb(255, 248, 240)
        Note over GRC,BD: PASO 4: Tabla Economía Naranja
        GRC->>SQL: Ejecutar template SQL
        SQL->>BD: JOIN con programas_economia_naranja
        SQL->>BD: Filtrar estados: EJECUCION, POR INICIAR
        BD-->>SQL: Datos filtrados
        SQL->>BD: CREATE TABLE ECONOMIA_NARANJA_{MES}_{AÑO}
        BD-->>SQL: Tabla creada
        SQL-->>GRC: 3998 registros procesados
    end

    rect rgb(255, 240, 245)
        Note over GRC,Metas: PASO 5 y 6: Metas y Cupos
        GRC->>Metas: subprocess.run(normalizar_metas_sena.py)
        Metas->>FS: Leer Metas SENA XLSX
        FS-->>Metas: Datos de metas
        Metas->>BD: Crear metas_sena_2025.db
        BD-->>Metas: BD creada
        Metas-->>GRC: Metas normalizadas

        GRC->>Metas: subprocess.run(cruce_metas_avance_final.py)
        Metas->>BD: Leer metas_sena_2025.db
        BD-->>Metas: Metas institucionales
        Metas->>FS: Leer Avance Cupos XLSB
        FS-->>Metas: Avance real
        Metas->>Metas: Calcular: META - AVANCE
        Metas->>FS: Generar cupos_disponibles_por_regional_2025.xlsx
        Metas->>FS: Generar cupos_disponibles_por_regional_2025.csv
        FS-->>Metas: Archivos generados
        GRC->>FS: Copiar archivos a datos_intermedios/
        Metas-->>GRC: Cupos calculados (33 regionales)
    end

    rect rgb(240, 255, 255)
        Note over GRC,Aprendices: PASO 7: Reporte de Aprendices
        GRC->>Aprendices: subprocess.run(generar_reporte_mensual_aprendices.py)
        Aprendices->>FS: Leer Avance Aprendices XLSB
        FS-->>Aprendices: Datos de aprendices
        Aprendices->>Aprendices: Consolidar estadísticas por regional
        Aprendices->>FS: Generar SENA Mensual Nacional {MES}.xlsx
        FS-->>Aprendices: Archivo generado
        GRC->>FS: Copiar archivo a datos_intermedios/
        Aprendices-->>GRC: Reporte generado
    end

    rect rgb(255, 245, 238)
        Note over GRC,Consolidado: PASO 8: Consolidación Final
        GRC->>Consolidado: subprocess.run(generar_reporte_consolidado.py)<br/>con variables de entorno

        Consolidado->>BD: Leer tabla ECONOMIA_NARANJA_{MES}_{AÑO}
        BD-->>Consolidado: 3998 registros

        Consolidado->>FS: Leer cupos_disponibles_por_regional_2025.xlsx
        FS-->>Consolidado: Datos de cupos (33 regionales)

        Consolidado->>FS: Leer SENA Mensual Nacional.xlsx
        FS-->>Consolidado: Datos de aprendices

        Consolidado->>Consolidado: Crear Excel con 3 hojas:<br/>1. Economía Naranja<br/>2. Oferta Disponible<br/>3. SENA Mensual

        Consolidado->>FS: Guardar en datos_finales/<br/>Reporte Consolidado Economía Naranja.xlsx
        FS-->>Consolidado: Archivo guardado

        Consolidado-->>GRC: Reporte consolidado exitoso
    end

    GRC->>Usuario: Proceso completado exitosamente<br/>Ubicación: datos_finales/Reporte Consolidado...
```

## Diagrama de Arquitectura de Componentes

```mermaid
graph TB
    subgraph "Archivos de Entrada"
        PE04[PE-04 XLSB<br/>Formación Nacional]
        AC[XLSB<br/>Avance Cupos]
        AA[XLSB<br/>Avance Aprendices]
        MS[XLSX<br/>Metas SENA]
        CAT[XLSX<br/>Catálogo 783 Programas]
    end

    subgraph "Scripts de Procesamiento"
        direction TB
        ORCH[generar_reporte_completo.py<br/>ORQUESTADOR]
        CONF[configuracion.py<br/>Configuración centralizada]

        subgraph "Procesamiento PE-04"
            IMP[importar_pe_04_mes.py]
            SQLF[crear_tabla_economia_naranja.sql]
        end

        subgraph "Componente Metas"
            NM[normalizar_metas_sena.py]
            CM[cruce_metas_avance_final.py]
        end

        subgraph "Componente Aprendices"
            GA[generar_reporte_mensual_aprendices.py]
        end

        subgraph "Consolidación"
            GC[generar_reporte_consolidado.py]
        end
    end

    subgraph "Datos Intermedios"
        BDFORM[(sena_formacion_MES.db<br/>+ tabla programas_economia_naranja)]
        BDMETAS[(metas_sena_2025.db)]
        CUPOS[cupos_disponibles_por_regional_2025.xlsx]
        RAPR[SENA Mensual Nacional.xlsx]
    end

    subgraph "Producto Final"
        FINAL[Reporte Consolidado<br/>Economía Naranja.xlsx<br/><br/>Hoja 1: Economía Naranja 3998 registros<br/>Hoja 2: Oferta Disponible 33 regionales<br/>Hoja 3: SENA Mensual Nacional]
    end

    PE04 --> IMP
    CAT --> IMP
    IMP --> BDFORM
    BDFORM --> SQLF
    SQLF --> BDFORM

    MS --> NM
    NM --> BDMETAS
    BDMETAS --> CM
    AC --> CM
    CM --> CUPOS

    AA --> GA
    GA --> RAPR

    BDFORM --> GC
    CUPOS --> GC
    RAPR --> GC
    GC --> FINAL

    ORCH -.controla.-> CONF
    ORCH -.ejecuta.-> IMP
    ORCH -.ejecuta.-> SQLF
    ORCH -.ejecuta.-> NM
    ORCH -.ejecuta.-> CM
    ORCH -.ejecuta.-> GA
    ORCH -.ejecuta.-> GC

    style ORCH fill:#FFD700
    style CONF fill:#87CEEB
    style FINAL fill:#90EE90
    style BDFORM fill:#DDA0DD
    style BDMETAS fill:#DDA0DD
```

## Diagrama de Estados del Proceso

```mermaid
stateDiagram-v2
    [*] --> Inicialización

    Inicialización --> ValidandoPrerequisitos: Ejecutar proceso
    ValidandoPrerequisitos --> CreandoDirectorios: Validación OK
    ValidandoPrerequisitos --> Error: Archivos faltantes

    CreandoDirectorios --> CopiandoArchivos: Estructura creada
    CopiandoArchivos --> GenerandoBDFormacion: Archivos copiados

    GenerandoBDFormacion --> CargandoCatalogo: Inicio importación
    CargandoCatalogo --> ImportandoPE04: Catálogo cargado (783 programas)
    ImportandoPE04 --> FiltrandoEconomiaNaranja: BD creada

    FiltrandoEconomiaNaranja --> GenerandoBDMetas: Tabla creada (3998 registros)
    GenerandoBDMetas --> CalculandoCupos: Metas normalizadas

    CalculandoCupos --> GenerandoReporteAprendices: Cupos calculados (33 regionales)
    GenerandoReporteAprendices --> ConsolidandoReporte: Reporte generado

    ConsolidandoReporte --> IntegrandoHoja1: Leyendo fuentes
    IntegrandoHoja1 --> IntegrandoHoja2: Economía Naranja
    IntegrandoHoja2 --> IntegrandoHoja3: Oferta Disponible
    IntegrandoHoja3 --> GuardandoReporte: SENA Mensual

    GuardandoReporte --> ProcesoCompletado: Archivo generado
    ProcesoCompletado --> [*]

    Error --> [*]

    note right of ValidandoPrerequisitos
        Verifica:
        - Python 3.10+
        - Dependencias
        - Archivos fuente
        - Espacio en disco
    end note

    note right of FiltrandoEconomiaNaranja
        Criterios:
        - Join con catálogo
        - Estados: EJECUCION, POR INICIAR
        - Cruce por código y versión
    end note

    note right of ConsolidandoReporte
        Variables de entorno:
        - BD_FORMACION
        - CUPOS_DISPONIBLES
        - REPORTE_APRENDICES
        - ARCHIVO_SALIDA
    end note
```

## Descripción de los Diagramas

### 1. Diagrama de Flujo del Proceso Completo
Muestra el flujo secuencial de los 8 pasos del proceso automatizado, incluyendo:
- Validaciones y puntos de decisión
- Generación de archivos intermedios
- Procesamiento de cada componente
- Consolidación final del reporte

### 2. Diagrama de Secuencia del Sistema
Ilustra la interacción temporal entre componentes:
- Orquestador (`generar_reporte_completo.py`) como coordinador central
- Llamadas entre scripts mediante `subprocess`
- Lectura/escritura de archivos y bases de datos
- Flujo de datos entre componentes
- Generación del reporte consolidado con 3 hojas

### 3. Diagrama de Arquitectura de Componentes
Presenta la estructura modular del sistema:
- Archivos de entrada (fuentes institucionales)
- Scripts de procesamiento organizados por función
- Bases de datos SQLite como almacenamiento intermedio
- Producto final consolidado
- Relaciones de dependencia entre componentes

### 4. Diagrama de Estados del Proceso
Describe los estados del sistema durante la ejecución:
- Transiciones entre estados
- Puntos de validación
- Notas con información clave de cada estado
- Manejo de errores y flujos alternativos

## Notas Técnicas

### Tecnologías Utilizadas en los Diagramas
- **Sintaxis**: Mermaid (compatible con GitHub, GitLab, y muchos visualizadores MD)
- **Renderización**: Los diagramas se renderizan automáticamente en plataformas compatibles

### Visualización
Para visualizar estos diagramas:
1. **GitHub/GitLab**: Automático al ver el archivo MD
2. **VS Code**: Instalar extensión "Markdown Preview Mermaid Support"
3. **Online**: https://mermaid.live/ (copiar y pegar el código)

### Colores en los Diagramas
- **Verde** (#90EE90): Estados exitosos
- **Rojo** (#FFB6C1): Estados de error
- **Azul claro** (#87CEEB): Procesos de consolidación
- **Amarillo** (#FFD700): Componente orquestador
- **Púrpura** (#DDA0DD): Bases de datos

---
