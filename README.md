# Sistema de Extracción y Consolidación de Documentos - Intercompañías

Sistema integral para la extracción automatizada de documentos SAP y consolidación de información de intercompañías (proveedores y clientes).

## 📋 Descripción

Esta aplicación facilita dos procesos principales:

1. **Descarga de Documentos SAP**: Extracción automatizada de reportes FBL1N, ZFIQ02, FBL3N y FBL5N desde SAP
2. **Consolidación**: Procesamiento y consolidación de información de proveedores y clientes en matrices resumen

## 🏗️ Arquitectura del Proyecto

El proyecto sigue el patrón **MVC (Model-View-Controller)** para mantener una separación clara de responsabilidades:

```
Intercompañías/
│
├── main.py                      # Punto de entrada principal
├── interfaz_GUI.py              # Vista (UI)
├── controller.py                # Controlador (lógica de negocio)
├── DescargaSAP.py               # Módulo de extracción SAP
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

## 🔧 Componentes

### 1. `main.py`
- Punto de entrada de la aplicación
- Inicializa la GUI y el controlador
- Conecta ambos componentes

### 2. `interfaz_GUI.py`
- **Responsabilidad**: Interfaz gráfica de usuario
- **Funciones**:
  - Gestión de sociedades
  - Selección de fechas
  - Muestra de estado del proceso
  - Botones de acción

### 3. `controller.py`
- **Responsabilidad**: Lógica de negocio
- **Funciones**:
  - `execute_download()`: Orquesta el proceso de descarga desde SAP (proveedores y clientes)
  - `execute_consolidation()`: Ejecuta el proceso de consolidación
  - Manejo de errores
  - Validaciones de datos

### 4. `DescargaSAP.py`
- **Responsabilidad**: Interacción con SAP GUI Scripting
- **Funciones**:
  - `FBL1N_Intercompañias()`: Descarga reporte FBL1N (proveedores)
  - `ZFIQ02_Intercompañias()`: Descarga catálogo de proveedores
  - `FBL3N()`: Descarga reporte FBL3N (cuentas contables)
  - `FBL5_Intercompañias()`: Descarga reporte FBL5N (clientes)

### 5. `Consolidacion_V2.py`
- **Responsabilidad**: Procesamiento y consolidación de datos
- **Funciones**:
  - `ejecutar_consolidacion_por_sociedad()`: Genera el archivo consolidado completo para una sociedad
  - Procesamiento de proveedores: FBL1N, ZFIQ02, FBL3N, Matriz Proveedores
  - Procesamiento de clientes: FBL5N, CLIENTES, FBL3N Clientes, Matriz Clientes

## 📦 Requisitos

### Software Necesario
- Python 3.8 o superior
- SAP GUI con Scripting habilitado
- Microsoft Excel

### Dependencias Python
```
pandas
pywin32
openpyxl
tkinter (incluido en Python estándar)
```

Instalar dependencias:
```bash
pip install -r requirements.txt
```

## 🚀 Uso

### Inicio de la Aplicación

```bash
python main.py
```

### Flujo de Trabajo Recomendado

#### 1. Descarga de Documentos

1. **Configurar fechas**:
   - Fecha Desde: `DD.MM.YYYY`
   - Fecha Hasta: `DD.MM.YYYY`

2. **Agregar sociedades**:
   - Ingresar código de sociedad (ej: MX73, MX30)
   - Hacer clic en "Agregar" o presionar Enter
   - Repetir para cada sociedad

3. **Ejecutar descarga**:
   - Clic en "⚡ Descargar Documentos"
   - Confirmar la operación
   - Esperar a que complete (SAP debe estar abierto)

**Archivos generados por sociedad en `Input/Proveedores/`**:
- `FBL1_Proveedores_{sociedad}.xlsx`
- `ZFIQ02_Proveedores_{sociedad}.xlsx`
- `FBL3N_Proveedores_{sociedad}.xlsx`

**Archivos generados por sociedad en `Input/Clientes/`**:
- `FBL5N_Clientes_{sociedad}.xlsx`
- `FBL3N_Clientes_{sociedad}.xlsx`

#### 2. Consolidación

1. **Asegurar archivos de entrada**:
   - Verificar que existan todos los archivos de proveedores y clientes en `Input/`
   - Confirmar que `CLIENTES.xlsx` esté en `Input/Clientes/`

2. **Ejecutar consolidación**:
   - Clic en "📊 Consolidación"
   - Confirmar la operación
   - Esperar a que complete

**Archivo generado por sociedad en `Output/`**:
- `Intercompanias_Consolidado_{sociedad}.xlsx` con 8 hojas:
  - `FBL1N`
  - `Cat Proveedores`
  - `FBL3N Proveedores`
  - `Matriz Proveedores`
  - `FBL5N`
  - `Cat Clientes`
  - `FBL3N Clientes`
  - `Matriz Clientes`

## 📁 Estructura de Carpetas

```
C:\Users\{Usuario}\Documents\Intercompañias\
└── RDA_Intercompanias\
    └── src\
        ├── Input\
        │   ├── Proveedores\
        │   │   ├── FBL1_Proveedores_{sociedad}.xlsx
        │   │   ├── ZFIQ02_Proveedores_{sociedad}.xlsx
        │   │   └── FBL3N_Proveedores_{sociedad}.xlsx
        │   └── Clientes\
        │       ├── FBL5N_Clientes_{sociedad}.xlsx
        │       ├── FBL3N_Clientes_{sociedad}.xlsx
        │       └── CLIENTES.xlsx          ← colocar manualmente
        └── Output\
            └── Intercompanias_Consolidado_{sociedad}.xlsx
```

**Nota**: La estructura de carpetas se crea automáticamente. Las rutas se adaptan dinámicamente al usuario actual del sistema (no requiere configuración manual).

## 🔍 Detalles Técnicos

### Proceso de Descarga

1. **Conexión SAP**: Utiliza `win32com.client` para conectar con SAP GUI
2. **Navegación**: Ejecuta transacciones FBL1N, ZFIQ02, FBL3N, FBL5N
3. **Filtrado**: Aplica filtros de sociedad y fechas
4. **Exportación**: Descarga archivos Excel por sociedad

### Proceso de Consolidación

1. **Carga de datos**: Lee archivos Excel de entrada (proveedores y clientes)
2. **Limpieza**: Elimina caracteres especiales y espacios
3. **Enriquecimiento**: Realiza VLOOKUPs entre catálogos y transacciones
4. **Filtrado por cuentas**: Aplica las cuentas contables configuradas por sociedad
5. **Agregación**: Crea matrices pivote por proveedor/cliente y cuenta
6. **Exportación**: Genera archivo consolidado multi-hoja por sociedad

### Cuentas Contables por Sociedad

#### Proveedores

| Sociedad | Cuentas |
|----------|---------|
| MX01 | 6600022, 7201000, 7204000 |
| MX05 | 7201000 |
| MX22 | 6600021, 2050000, 6600022, 6700040, 6700043, 6700048, 6900010 |
| MX30 | 6600022, 7204000, 6900010 |
| MX73 | 6600022 |
| *Resto* | 6600022 (default) |

#### Clientes

| Sociedad | Cuentas |
|----------|---------|
| MX01 | 7000005, 7000020, 7201000 |
| MX05 | 7201000 |
| MX22 | 4300010, 7000005, 7001002, 7010005, 7201000 |
| MX30 | 7001000, 7001002, 7001005, 7011000, 7500000 |
| MX31 | 7201000 |
| MX32 | 7201000 |
| MX73 | 7000005, 7001002, 7201000 |
| MX80 | 7201000 |
| *Resto* | 7201000 (default) |

### Validaciones

- Formato de fechas (DD.MM.YYYY)
- Existencia de sociedades seleccionadas
- Disponibilidad de todos los archivos de entrada (proveedores y clientes)
- Existencia de columnas requeridas

## ⚠️ Consideraciones Importantes

### SAP GUI Scripting

- Debe estar habilitado en SAP GUI
- SAP debe estar abierto y conectado durante la descarga
- No interactuar con SAP durante el proceso

### Archivos de Entrada

- Deben tener la estructura esperada (columnas específicas)
- Los nombres de archivo son sensibles
- Deben estar en las carpetas correctas
- `CLIENTES.xlsx` debe colocarse manualmente en `Input/Clientes/`

### Performance

- El proceso de descarga puede tardar varios minutos dependiendo del volumen de datos
- La consolidación es generalmente rápida (<1 minuto)
- No cerrar la aplicación durante la ejecución

## 🐛 Solución de Problemas

### Error: "Archivos no encontrados"
- Verificar que los archivos estén en las rutas correctas
- Revisar nombres exactos de archivos (son sensibles a mayúsculas)
- Ejecutar primero el proceso de descarga
- Confirmar que `CLIENTES.xlsx` esté en `Input/Clientes/`

### Error en SAP
- Verificar que SAP GUI Scripting esté habilitado
- Confirmar que SAP esté abierto y conectado
- Revisar que las transacciones estén disponibles (FBL1N, ZFIQ02, FBL3N, FBL5N)

### Error: "Formato de fecha inválido"
- Usar el formato DD.MM.YYYY exacto
- Ejemplo: 01.01.2025

### Sociedad sin datos en la matriz
- Verificar que la sociedad esté en el diccionario `CUENTAS_POR_SOCIEDAD` o `CUENTAS_CLIENTES_POR_SOCIEDAD`
- Si no está, se aplica el valor default (6600022 para proveedores, 7201000 para clientes)

## 📝 Notas de Desarrollo

### Agregar una Nueva Sociedad

1. Abrir `Consolidacion_V2.py`
2. Agregar la sociedad y sus cuentas en `CUENTAS_POR_SOCIEDAD` (proveedores)
3. Agregar la sociedad y sus cuentas en `CUENTAS_CLIENTES_POR_SOCIEDAD` (clientes)
4. Si tiene cuentas con conceptos nuevos, actualizar `CONCEPTO_POR_CUENTA_CLIENTE`

### Agregar una Nueva Cuenta

1. Agregar la cuenta al diccionario de la sociedad correspondiente en `CUENTAS_POR_SOCIEDAD` o `CUENTAS_CLIENTES_POR_SOCIEDAD`
2. Para cuentas de clientes, agregar el concepto en `CONCEPTO_POR_CUENTA_CLIENTE`

### Mantenimiento

- Revisar cuentas por sociedad si cambian las estructuras contables
- Actualizar nombres de columnas si SAP cambia estructura de reportes
- Mantener sincronizadas las versiones de pandas/openpyxl

## 📄 Licencia

Uso interno - Todos los derechos reservados