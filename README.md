# Sistema de Extracción y Consolidación de Documentos - Intercompañías

Sistema integral para la extracción automatizada de documentos SAP y consolidación de información de intercompañías.

## 📋 Descripción

Esta aplicación facilita dos procesos principales:

1. **Descarga de Documentos SAP**: Extracción automatizada de reportes FBL1N, ZFIQ02 y FBL3N desde SAP
2. **Consolidación**: Procesamiento y consolidación de información de proveedores y clientes en matrices resumen

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
  - `execute_download()`: Orquesta el proceso de descarga desde SAP
  - `execute_consolidation()`: Ejecuta el proceso de consolidación
  - Manejo de errores
  - Validaciones de datos

### 4. `FBL1_Intercompañias.py`
- **Responsabilidad**: Interacción con SAP GUI Scripting
- **Funciones**:
  - `FBL1_Intercompañias()`: Descarga reporte FBL1N
  - `ZFIQ02_Intercompañias()`: Descarga catálogo de proveedores
  - `FBL3N()`: Descarga reporte FBL3N

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
   - Clic en "⚡ Correr Descarga de Documentos"
   - Confirmar la operación
   - Esperar a que complete (SAP debe estar abierto)

**Archivos generados**:
- `FBL1_Intercompañias.xlsx`
- `ZFIQ02_Intercompañias.xlsx`
- Archivos FBL3N adicionales

#### 2. Consolidación

1. **Asegurar archivos de entrada**:
   - Verificar que existan todos los archivos necesarios en `Input/`
   - Proveedores: FBL1, ZFIQ02, FBL301-10
   - Clientes: FBL5N, FBL3N Clientes, CLIENTES.xlsx

2. **Ejecutar consolidación**:
   - Clic en "📊 Consolidación"
   - Confirmar la operación
   - Esperar a que complete

**Archivo generado**:
- `Output/Intercompanias_Consolidado.xlsx` con 8 hojas:
  - FBL1N
  - Cat Proveedores
  - FBL3N Proveedores
  - Matriz Proveedores
  - FBL5N
  - Cat Clientes
  - FBL3N Clientes
  - Matriz Clientes

## 📁 Estructura de Carpetas

```
C:\Users\{Usuario}\Documents\Intercompañias\
└── RDA_Intercompanias\
    └── src\
        ├── Input\
        │   ├── Proveedores\
        │   │   ├── FBL1_Intercompañias 1.xlsx
        │   │   ├── ZFIQ02_Intercompañias.xlsx
        │   │   └── FBL301-10.xlsx
        │   └── Clientes\
        │       ├── fbl5n 4.xlsx
        │       ├── FBL3N Clientes.xlsx
        │       └── CLIENTES.xlsx
        └── Output\
            └── Intercompanias_Consolidado.xlsx
```

**Nota**: La estructura de carpetas se crea automáticamente. Las rutas se adaptan dinámicamente al usuario actual del sistema (no requiere configuración manual).

## 🔍 Detalles Técnicos

### Proceso de Descarga

1. **Conexión SAP**: Utiliza `win32com.client` para conectar con SAP GUI
2. **Navegación**: Ejecuta transacciones FBL1N, ZFIQ02, FBL3N
3. **Filtrado**: Aplica filtros de sociedad y fechas
4. **Exportación**: Descarga archivos Excel

### Proceso de Consolidación

1. **Carga de datos**: Lee archivos Excel de entrada
2. **Limpieza**: Elimina caracteres especiales y espacios
3. **Enriquecimiento**: Realiza VLOOKUPs entre catálogos y transacciones
4. **Agregación**: Crea matrices pivote por proveedor/cliente y cuenta
5. **Exportación**: Genera archivo consolidado multi-hoja

### Validaciones

- Formato de fechas (DD.MM.YYYY)
- Existencia de sociedades seleccionadas
- Disponibilidad de archivos de entrada
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

### Performance

- El proceso de descarga puede tardar varios minutos dependiendo del volumen de datos
- La consolidación es generalmente rápida (<1 minuto)
- No cerrar la aplicación durante la ejecución

## 🐛 Solución de Problemas

### Error: "Archivos no encontrados"
- Verificar que los archivos estén en las rutas correctas
- Revisar nombres exactos de archivos
- Ejecutar primero el proceso de descarga

### Error en SAP
- Verificar que SAP GUI Scripting esté habilitado
- Confirmar que SAP esté abierto y conectado
- Revisar que las transacciones estén disponibles

### Error: "Formato de fecha inválido"
- Usar el formato DD.MM.YYYY exacto
- Ejemplo: 01.01.2025

## 📝 Notas de Desarrollo

### Mejoras Futuras

- [ ] Validación de archivos de entrada más robusta
- [ ] Logs detallados de proceso
- [ ] Configuración de rutas personalizable
- [ ] Opción de selección de archivos manual
- [ ] Manejo de múltiples periodos en batch

### Mantenimiento

- Revisar rutas hardcodeadas si cambia estructura de carpetas
- Actualizar nombres de columnas si SAP cambia estructura de reportes
- Mantener sincronizadas las versiones de pandas/openpyxl

## 👥 Autores

Sistema de Intercompañías - 2025

## 📄 Licencia

Uso interno - Todos los derechos reservados
