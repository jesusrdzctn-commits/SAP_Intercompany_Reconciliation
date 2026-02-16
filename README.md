# Sistema de Extracción y Consolidación de Documentos - Intercompañías

Sistema integral para la extracción automatizada de documentos SAP y consolidación de información de intercompañías **por sociedad individual**.

## 📋 Descripción

Esta aplicación facilita dos procesos principales:

1. **Descarga de Documentos SAP**: Extracción automatizada de reportes FBL1N, ZFIQ02 y FBL3N desde SAP **por cada sociedad seleccionada**
2. **Consolidación**: Procesamiento y consolidación de información de proveedores y clientes en matrices resumen **generando un archivo Excel por cada sociedad**

## 🏗️ Arquitectura del Proyecto

El proyecto sigue el patrón **MVC (Model-View-Controller)** para mantener una separación clara de responsabilidades:

```
Intercompañías/
│
├── main.py                      # Punto de entrada principal
├── interfaz_GUI.py              # Vista (UI)
├── controller.py                # Controlador (lógica de negocio)
├── FBL1_Intercompañias.py      # Módulo de extracción SAP
├── Consolidacion_V2.py          # Módulo de consolidación
├── requirements.txt             # Dependencias
│
└── Documents/Intercompañías/    # Estructura de carpetas de datos (se crea automáticamente)
    └── RDA_Intercompanias/
        └── src/
            ├── Input/
            │   ├── Proveedores/
            │   └── Clientes/
            └── Output/
```

## 🚀 Uso

### Inicio de la Aplicación

```bash
python main.py
```

### Flujo de Trabajo Recomendado

#### 1. Descarga de Documentos

1. **Configurar fechas**:
   - Fecha Desde: `DD.MM.YYYY` (ejemplo: 01.01.2025)
   - Fecha Hasta: `DD.MM.YYYY` (ejemplo: 31.12.2025)

2. **Agregar sociedades**:
   - Ingresar código de sociedad (ej: MX73)
   - Hacer clic en "Agregar" o presionar Enter
   - Repetir para cada sociedad adicional (ej: MX30, MX80)
   - **IMPORTANTE**: El sistema procesará cada sociedad por separado

3. **Ejecutar descarga**:
   - Clic en "⚡ Correr Descarga de Documentos"
   - Confirmar la operación
   - Esperar a que complete (SAP debe estar abierto)
   - **El proceso se ejecutará sociedad por sociedad**

**Archivos generados (por cada sociedad)**:

Si seleccionas **MX73, MX30, MX80**, se generarán:

```
Input/Proveedores/
├── FBL1_Intercompañias_MX73.xlsx
├── FBL1_Intercompañias_MX30.xlsx
├── FBL1_Intercompañias_MX80.xlsx
├── ZFIQ02_Intercompañias_MX73.xlsx
├── ZFIQ02_Intercompañias_MX30.xlsx
├── ZFIQ02_Intercompañias_MX80.xlsx
├── FBL3N_Proveedores_MX73.xlsx
├── FBL3N_Proveedores_MX30.xlsx
└── FBL3N_Proveedores_MX80.xlsx
```

#### 2. Consolidación

1. **Asegurar archivos de entrada**:
   - Verificar que existan todos los archivos necesarios en `Input/`
   - **Proveedores**: FBL1_Intercompañias_[SOCIEDAD].xlsx, ZFIQ02_Intercompañias_[SOCIEDAD].xlsx, FBL3N_Proveedores_[SOCIEDAD].xlsx
   - **Clientes**: fbl5n_[SOCIEDAD].xlsx, FBL3N_Clientes_[SOCIEDAD].xlsx, CLIENTES.xlsx (común)

2. **Ejecutar consolidación**:
   - Clic en "📊 Consolidación"
   - Confirmar la operación
   - **El sistema detecta automáticamente las sociedades disponibles**
   - Esperar a que complete (procesa cada sociedad secuencialmente)

**Archivos generados (uno por cada sociedad)**:

```
Output/
├── Intercompanias_Consolidado_MX73.xlsx
├── Intercompanias_Consolidado_MX30.xlsx
└── Intercompanias_Consolidado_MX80.xlsx
```

**Cada archivo consolidado contiene 8 hojas**:
- FBL1N
- Cat Proveedores
- FBL3N Proveedores
- Matriz Proveedores ⭐
- FBL5N
- Cat Clientes
- FBL3N Clientes
- Matriz Clientes ⭐

## 📁 Estructura de Carpetas

### Ejemplo con 3 sociedades (MX73, MX30, MX80):

```
C:\Users\{Usuario}\Documents\Intercompañias\
└── RDA_Intercompanias\
    └── src\
        ├── Input\
        │   ├── Proveedores\
        │   │   ├── FBL1_Intercompañias_MX73.xlsx       (generado)
        │   │   ├── FBL1_Intercompañias_MX30.xlsx       (generado)
        │   │   ├── FBL1_Intercompañias_MX80.xlsx       (generado)
        │   │   ├── ZFIQ02_Intercompañias_MX73.xlsx     (generado)
        │   │   ├── ZFIQ02_Intercompañias_MX30.xlsx     (generado)
        │   │   ├── ZFIQ02_Intercompañias_MX80.xlsx     (generado)
        │   │   ├── FBL3N_Proveedores_MX73.xlsx         (generado)
        │   │   ├── FBL3N_Proveedores_MX30.xlsx         (generado)
        │   │   └── FBL3N_Proveedores_MX80.xlsx         (generado)
        │   │
        │   └── Clientes\
        │       ├── fbl5n_MX73.xlsx                      (manual/otro proceso)
        │       ├── fbl5n_MX30.xlsx                      (manual/otro proceso)
        │       ├── fbl5n_MX80.xlsx                      (manual/otro proceso)
        │       ├── FBL3N_Clientes_MX73.xlsx             (manual/otro proceso)
        │       ├── FBL3N_Clientes_MX30.xlsx             (manual/otro proceso)
        │       ├── FBL3N_Clientes_MX80.xlsx             (manual/otro proceso)
        │       └── CLIENTES.xlsx                        (catálogo común)
        │
        └── Output\
            ├── Intercompanias_Consolidado_MX73.xlsx    (generado)
            ├── Intercompanias_Consolidado_MX30.xlsx    (generado)
            └── Intercompanias_Consolidado_MX80.xlsx    (generado)
```

**Nota**: 
- La estructura de carpetas se crea automáticamente
- Las rutas se adaptan dinámicamente al usuario actual del sistema
- **Un archivo consolidado por cada sociedad procesada**
- Los archivos de clientes deben tener el sufijo de la sociedad correspondiente

## 🔍 Detalles Técnicos

### Proceso de Descarga (Por Sociedad)

Para cada sociedad seleccionada:

1. **Conexión SAP**: Utiliza `win32com.client` para conectar con SAP GUI
2. **Navegación**: Ejecuta transacciones FBL1N, ZFIQ02, FBL3N
3. **Filtrado**: Aplica filtros de sociedad y fechas
4. **Exportación**: Descarga archivos Excel con nomenclatura: `[REPORTE]_[SOCIEDAD].xlsx`
5. **Iteración**: Repite el proceso para la siguiente sociedad

**Ejemplo de progreso**:
```
📥 Procesando sociedad MX73 (1/3)...
📥 Descargando FBL1 - MX73...
📥 Descargando ZFIQ02 - MX73...
📥 Descargando FBL3N - MX73...
✅ Sociedad MX73 completada (1/3)

📥 Procesando sociedad MX30 (2/3)...
...
```

### Proceso de Consolidación (Por Sociedad)

1. **Detección automática**: Busca archivos `FBL1_Intercompañias_*.xlsx` para identificar sociedades
2. **Carga de datos**: Lee archivos Excel de entrada para cada sociedad
3. **Limpieza**: Elimina caracteres especiales y espacios
4. **Enriquecimiento**: Realiza VLOOKUPs entre catálogos y transacciones
5. **Agregación**: Crea matrices pivote por proveedor/cliente y cuenta
6. **Exportación**: Genera archivo consolidado `Intercompanias_Consolidado_[SOCIEDAD].xlsx`
7. **Repetición**: Procesa la siguiente sociedad detectada

## ⚠️ Consideraciones Importantes

### Nomenclatura de Archivos

**CRÍTICO**: Todos los archivos deben seguir la convención:
```
[TIPO_REPORTE]_[SOCIEDAD].xlsx
```

Ejemplos válidos:
- ✅ `FBL1_Intercompañias_MX73.xlsx`
- ✅ `fbl5n_MX30.xlsx`
- ✅ `FBL3N_Clientes_MX80.xlsx`

Ejemplos inválidos:
- ❌ `FBL1_Intercompañias.xlsx` (falta código de sociedad)
- ❌ `FBL1_MX73_Intercompañias.xlsx` (sociedad en posición incorrecta)

### Performance

- El proceso de descarga puede tardar **varios minutos por sociedad**
- **Tiempo total = Tiempo por sociedad × Número de sociedades**
- La consolidación es generalmente rápida (<1 minuto por sociedad)
- No cerrar la aplicación durante la ejecución

## 💡 Casos de Uso

### Caso 1: Una sola sociedad
```
Sociedades seleccionadas: MX73
Archivos descargados: 3 archivos (FBL1, ZFIQ02, FBL3N)
Archivos consolidados: 1 archivo (Intercompanias_Consolidado_MX73.xlsx)
```

### Caso 2: Múltiples sociedades
```
Sociedades seleccionadas: MX73, MX30, MX80, MX01
Archivos descargados: 12 archivos (3 por sociedad)
Archivos consolidados: 4 archivos (1 por sociedad)
```

### Caso 3: Consolidación posterior
```
Descargas previas: MX73, MX30
Nueva descarga: MX80
Consolidación: Detecta y procesa MX73, MX30, MX80
Resultado: 3 archivos consolidados
```

## 📄 Licencia

Uso interno - Todos los derechos reservados