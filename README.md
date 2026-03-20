# Sistema de Extracción y Consolidación de Documentos - Intercompañías

Sistema integral para la extracción automatizada de documentos SAP y consolidación de información de intercompañías (proveedores y clientes).

## 📋 Descripción

Esta aplicación facilita dos procesos principales:

1. **Descarga de Documentos SAP**: Extracción automatizada de reportes FBL1N, ZFIQ02, FBL3N y FBL5N desde SAP mediante SAP GUI Scripting
2. **Consolidación**: Procesamiento y consolidación de información de proveedores y clientes en matrices resumen por sociedad

## 🏗️ Arquitectura del Proyecto

El proyecto sigue el patrón **MVC (Model-View-Controller)**:

```
Intercompañías/
│
├── main.py                      # Punto de entrada principal
├── interfaz_GUI.py              # Vista (UI con tkinter)
├── controller.py                # Controlador (lógica de negocio)
├── DescargaSAP.py               # Módulo de extracción SAP
├── Consolidacion_V2.py          # Módulo de consolidación
├── build_exe.py                 # Script para generar ejecutable
├── requirements.txt             # Dependencias Python
│
└── Documents/Intercompañías/    # Estructura de carpetas de datos (se crea automáticamente)
    └── RDA_Intercompanias/
        └── src/
            ├── Input/
            │   ├── Proveedores/
            │   └── Clientes/
            ├── Output/
            └── config/
                ├── cuentas_proveedores_por_sociedad.txt
                └── cuentas_clientes_por_sociedad.txt
```

## 🔧 Componentes

### `main.py`
Punto de entrada de la aplicación. Inicializa la ventana de tkinter, instancia la GUI y el controlador, y conecta ambos componentes antes de arrancar el loop de eventos.

### `interfaz_GUI.py`
Interfaz gráfica construida con **tkinter**. Organizada en dos pestañas:

**Pestaña Proceso:**
- Selector de fechas (Fecha Desde / Fecha Hasta) en formato `DD.MM.YYYY`
- Botones de acción: `⚡ Descargar Documentos` y `📊 Conciliación / Consolidación`
- Barra de estado en tiempo real

**Pestaña Configuración:**
- Gestión de sociedades (agregar / eliminar)
- Rangos de cuentas para FBL1N (proveedores) y FBL5N (clientes)
- Rutas de trabajo: entrada proveedores, entrada clientes y salida (con botón "Buscar" y normalización automática de separadores `\` / `/`)
- Tablas editables de cuentas por sociedad para proveedores y clientes
- Botones para persistir la configuración en archivos `.txt` dentro de la carpeta `config/`

La configuración de cuentas se guarda y carga automáticamente desde archivos `.txt` al iniciar la aplicación, por lo que los cambios sobreviven entre sesiones.

### `controller.py`
Controlador que orquesta los dos procesos principales:

**`execute_download()`**
- Valida fechas y sociedades antes de iniciar
- Itera sobre cada sociedad seleccionada ejecutando en orden: FBL1N → ZFIQ02 → FBL3N Proveedores → FBL5N → FBL3N Clientes
- Detecta automáticamente si una sociedad no tiene movimientos (archivo vacío) y omite los pasos dependientes
- Al finalizar cada sociedad guarda un archivo `_flags_{sociedad}.json` con los flags `sin_proveedores` y `sin_clientes` para que la consolidación los consuma
- Cierra workbooks de Excel abiertos por SAP antes de continuar

**`execute_consolidation()`**
- Lee los flags `_flags_{sociedad}.json` dejados por la descarga para saber si omitir proveedores o clientes
- Llama a `ejecutar_consolidacion_por_sociedad()` pasando las cuentas configuradas en la GUI
- Maneja errores de archivos faltantes con mensajes descriptivos

### `DescargaSAP.py`
Interacción directa con **SAP GUI Scripting** vía `win32com.client`. Contiene cuatro funciones principales:

**`FBL1N_Intercompañias()`** — Reporte de partidas abiertas de proveedores
- Navega a la transacción FBL1N
- Configura rango de cuentas, sociedades, fechas y variante de layout `/TAXVJG`
- Detecta el mensaje `MSITEM033` ("No se ha seleccionado ninguna partida") y crea un Excel vacío si no hay datos, devolviendo `False`
- Exporta el archivo y devuelve `True` si hay datos

**`FBL5_Intercompañias()`** — Reporte de partidas abiertas de clientes
- Navega a la transacción FBL5N
- Maneja popups de subsidiaria/central con `_cerrar_popup_subsidiaria()` en un loop hasta que desaparezcan
- Misma lógica de detección de sin-partidas y exportación que FBL1N

**`ZFIQ02_Intercompañias()`** — Catálogo de proveedores
- Navega a la transacción ZFIQ02 y filtra por sociedades
- Exporta a través del menú de SAP; detecta el Excel abierto en la instancia de Excel activa y lo guarda con `SaveAs` en la ruta indicada

**`FBL3N()`** — Reporte de partidas de cuentas contables (usado tanto para proveedores como para clientes)
- Recibe una lista de números de documento (`resultado`) y los carga en el filtro usando el portapapeles (`pyperclip`)
- Configura rango de fechas, sociedades y variante `RDA_FBL3N`
- Exporta el archivo en la ruta indicada

**Funciones auxiliares internas:**
- `_verificar_sin_partidas()`: detecta el mensaje de error SAP, cierra el popup, crea un Excel vacío y retorna `True`
- `_cerrar_popup_subsidiaria()`: itera cerrando popups de subsidiaria hasta que `wnd[1]` deje de existir

### `Consolidacion_V2.py`
Procesamiento de datos con **pandas** y **openpyxl**. Función principal: `ejecutar_consolidacion_por_sociedad()`.

**Proceso de proveedores:**
1. Lee `FBL1_Proveedores_{sociedad}.xlsx` y limpia caracteres especiales
2. Hace merge con el catálogo ZFIQ02 para agregar el nombre del proveedor
3. Lee `FBL3N_Proveedores_{sociedad}.xlsx`, filtra por las cuentas configuradas para la sociedad y enriquece con nombre, texto y texto de cabecera desde FBL1N
4. Genera la **Matriz Proveedores**: tabla pivote (proveedor × cuenta) con filas de totales, cuadre balanza y variaciones

**Proceso de clientes:**
1. Lee `FBL5N_Clientes_{sociedad}.xlsx` y hace merge con `Clientes_Catalogo.xls` para agregar nombre del cliente
2. Lee `FBL3N_Clientes_{sociedad}.xlsx`, filtra por las cuentas configuradas y enriquece con nombre, texto y asignación desde FBL5N
3. Genera la **Matriz Clientes** con la misma estructura pivote

**Manejo de sin movimientos:**
Si los flags `sin_proveedores` o `sin_clientes` vienen activos, se generan matrices vacías con una fila "SIN MOVIMIENTOS" en lugar de intentar procesar archivos sin datos.

**Archivos opcionales de enriquecimiento:**
- `Sociedad_Nombre.xls`: agrega el nombre de la empresa en la primera fila de cada matriz
- `Cuentas_Desc.xls`: agrega la descripción de cada cuenta contable en la fila de encabezado de la matriz

**Salida:** archivo `Intercompanias_Consolidado_{sociedad}.xlsx` con 8 hojas:

| Hoja | Contenido |
|------|-----------|
| `FBL1N` | Partidas de proveedores con nombre enriquecido |
| `Cat Proveedores` | Catálogo completo ZFIQ02 |
| `FBL3N Proveedores` | Partidas contables de proveedores filtradas |
| `Matriz Proveedores` | Pivote proveedor × cuenta |
| `FBL5N` | Partidas de clientes con nombre enriquecido |
| `Cat Clientes` | Catálogo de clientes |
| `FBL3N Clientes` | Partidas contables de clientes filtradas |
| `Matriz Clientes` | Pivote cliente × cuenta |

### `build_exe.py`
Script de empaquetado con **PyInstaller**:
- Verifica que todos los archivos fuente existen antes de compilar
- Instala PyInstaller automáticamente si no está disponible
- Genera un ejecutable `--onedir --windowed` con todos los hidden imports necesarios (`openpyxl`, `win32com`, `pyperclip`, `xlrd`, `pandas`)
- Reporta el tamaño del ejecutable generado e instrucciones para el usuario final

## 📦 Requisitos

### Software Necesario
- Python 3.8 o superior
- SAP GUI con Scripting habilitado
- Microsoft Excel (debe estar instalado; ZFIQ02 lo usa para guardar)

### Dependencias Python
```
pandas
pywin32
openpyxl
pyperclip
xlrd
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

#### 1. Configuración inicial (Pestaña Configuración)

1. Agregar las sociedades a procesar
2. Verificar o ajustar los rangos de cuentas para FBL1N y FBL5N
3. Ajustar las cuentas contables por sociedad (proveedores y clientes)
4. Guardar la configuración con los botones "Guardar TXT" para que persista entre sesiones
5. Verificar las rutas de trabajo (se crean automáticamente si no existen)

#### 2. Descarga de Documentos (Pestaña Proceso)

1. Configurar el intervalo de fechas (`DD.MM.YYYY`)
2. Asegurarse de que SAP esté abierto y conectado
3. Clic en `⚡ Descargar Documentos` y confirmar

**Archivos generados en `Input/Proveedores/`:**
- `FBL1_Proveedores_{sociedad}.xlsx`
- `ZFIQ02_Proveedores_{sociedad}.xlsx`
- `FBL3N_Proveedores_{sociedad}.xlsx`

**Archivos generados en `Input/Clientes/`:**
- `FBL5N_Clientes_{sociedad}.xlsx`
- `FBL3N_Clientes_{sociedad}.xlsx`

**Archivos de control generados en `Input/Proveedores/`:**
- `_flags_{sociedad}.json` (indica si hubo movimientos de proveedores/clientes)

#### 3. Consolidación (Pestaña Proceso)

1. Verificar que `Clientes_Catalogo.xls` esté en `Input/Clientes/` (colocarlo manualmente)
2. Clic en `📊 Conciliación / Consolidación` y confirmar

**Archivo generado en `Output/`:**
- `Intercompanias_Consolidado_{sociedad}.xlsx`

## 📁 Estructura de Carpetas

```
C:\Users\{Usuario}\Documents\Intercompañias\
└── RDA_Intercompanias\
    └── src\
        ├── Input\
        │   ├── Proveedores\
        │   │   ├── FBL1_Proveedores_{sociedad}.xlsx
        │   │   ├── ZFIQ02_Proveedores_{sociedad}.xlsx
        │   │   ├── FBL3N_Proveedores_{sociedad}.xlsx
        │   │   └── _flags_{sociedad}.json
        │   └── Clientes\
        │       ├── FBL5N_Clientes_{sociedad}.xlsx
        │       ├── FBL3N_Clientes_{sociedad}.xlsx
        │       ├── Clientes_Catalogo.xls      ← colocar manualmente
        │       ├── Sociedad_Nombre.xls        ← opcional
        │       └── Cuentas_Desc.xls           ← opcional
        ├── Output\
        │   └── Intercompanias_Consolidado_{sociedad}.xlsx
        └── config\
            ├── cuentas_proveedores_por_sociedad.txt
            └── cuentas_clientes_por_sociedad.txt
```

> Las rutas se crean automáticamente al iniciar la aplicación y se adaptan al usuario actual del sistema sin requerir configuración manual.

## ⚙️ Cuentas Contables por Sociedad

### Proveedores (valores por defecto)

| Sociedad | Cuentas |
|----------|---------|
| MX01 | 6600022, 7201000, 7204000 |
| MX05 | 7201000 |
| MX22 | 6600021, 2050000, 6600022, 6700040, 6700043, 6700048, 6900010 |
| MX30 | 6600022, 7204000, 6900010 |
| MX73 | 6600022 |
| *Resto* | 6600022 (default) |

### Clientes (valores por defecto)

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

> Estos valores pueden modificarse desde la pestaña **Configuración** de la aplicación y se persisten automáticamente en `config/`.

## 🔨 Generar Ejecutable

Para distribuir la aplicación sin requerir Python instalado:

```bash
python build_exe.py
```

El ejecutable se genera en `dist/Intercompanias/`. Solo es necesario distribuir esa carpeta; el usuario final únicamente necesita ejecutar `Intercompanias.exe`.

**Archivos y carpetas que NO deben incluirse en el repositorio** (agregar al `.gitignore`):

```gitignore
# PyInstaller
dist/
build/
*.spec

# Python
__pycache__/
*.pyc

# Entorno virtual
venv/
.env/

# Archivos generados por la aplicación
*.json
Input/
Output/
config/
```

## ⚠️ Consideraciones Importantes

### SAP GUI Scripting
- Debe estar habilitado en las opciones de SAP GUI
- SAP debe estar abierto y conectado antes de iniciar la descarga
- No interactuar con el escritorio durante el proceso de descarga

### Rutas de Trabajo
- Los botones "Buscar" normalizan automáticamente los separadores (`/` → `\`) para compatibilidad con SAP GUI Scripting
- Si se escribe la ruta manualmente en el campo de texto, usar `\` como separador o dejar que la aplicación lo normalice al ejecutar

### Archivos de Entrada
- Los nombres de archivo son sensibles (deben coincidir exactamente)
- `Clientes_Catalogo.xls` debe colocarse manualmente en `Input/Clientes/`
- `Sociedad_Nombre.xls` y `Cuentas_Desc.xls` son opcionales; si no existen, la consolidación continúa sin esos datos

## 🐛 Solución de Problemas

**"Archivos no encontrados"**
Verificar que los archivos estén en las rutas correctas y que se haya ejecutado primero la descarga. Confirmar que `Clientes_Catalogo.xls` esté en `Input/Clientes/`.

**Error en SAP al descargar**
Verificar que SAP GUI Scripting esté habilitado (`Opciones → Accesibilidad & Scripting`). Confirmar que las transacciones FBL1N, ZFIQ02, FBL3N y FBL5N estén disponibles para el usuario.

**"Formato de fecha inválido"**
Usar el formato exacto `DD.MM.YYYY`. Ejemplo: `01.01.2025`.

**La matriz aparece vacía o con "SIN MOVIMIENTOS"**
La sociedad no tuvo partidas en el periodo seleccionado. Verificar el rango de fechas y los rangos de cuentas configurados para FBL1N / FBL5N.

**ZFIQ02 no guarda el archivo correctamente**
Microsoft Excel debe estar instalado. Si hay múltiples workbooks abiertos en Excel al momento de la descarga, pueden generarse conflictos al guardar.

## 📝 Mantenimiento y Extensión

### Agregar una nueva sociedad
Desde la pestaña **Configuración** de la aplicación, ingresar la sociedad y sus cuentas tanto en la sección de proveedores como en la de clientes, y guardar con los botones "Guardar TXT".

### Agregar una nueva cuenta contable
Actualizar la entrada de la sociedad correspondiente en la pestaña **Configuración** y guardar.

### Actualizar nombres de columnas SAP
Si SAP cambia la estructura de los reportes, revisar los índices de columna usados en `controller.py` (por ejemplo, `colNDocument = df_FBL1.columns[6]` para FBL1N y `colNDocument_cli = df_FBL5.columns[8]` para FBL5N).

## 📄 Licencia

Uso interno — Todos los derechos reservados
