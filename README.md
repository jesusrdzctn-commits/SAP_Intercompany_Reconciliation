# Sistema de Extracción y Consolidación de Documentos - Intercompañías

Sistema integral para la extracción automatizada de documentos SAP y consolidación de información de intercompañías (proveedores y clientes).

## 📋 Descripción

Esta aplicación facilita dos procesos principales:

1. **Descarga de Documentos SAP**: Extracción automatizada de reportes FBL1N, ZFIQ02, FBL3N y FBL5N desde SAP, con soporte para sociedades de volumen normal y sociedades grandes (descarga por bloques de documentos).
2. **Consolidación**: Procesamiento y consolidación de información de proveedores y clientes en matrices resumen, con modo Manual (cuentas configuradas en la pestaña 'Configuración') o Automático (todas las cuentas del FBL3N).

## 🏗️ Arquitectura del Proyecto

El proyecto sigue el patrón **MVC (Model-View-Controller)**:

```
Intercompañías/
│
├── main.py                      # Punto de entrada principal
├── interfaz_GUI.py              # Vista (UI / Tkinter)
├── controller.py                # Controlador (lógica de negocio)
├── DescargaSAP.py               # Módulo de extracción SAP
├── Consolidacion_V2.py          # Módulo de consolidación
├── build_exe.py                 # Script para generar ejecutable con PyInstaller
├── requirements.txt             # Dependencias Python
│
└── Documents/Intercompañias/    # Estructura de datos (se crea automáticamente)
    └── RDA_Intercompanias/
        └── src/
            ├── Input/
            │   ├── Proveedores/
            │   └── Clientes/
            ├── Output/
            └── config/          # Persistencia de cuentas por sociedad (.txt)
```

## 🔧 Componentes

### 1. `main.py`
- Punto de entrada de la aplicación
- Inicializa la GUI y el controlador y los conecta

### 2. `interfaz_GUI.py`
- **Responsabilidad**: Interfaz gráfica de usuario (Tkinter, ventana 1000×750)
- **Pestaña Proceso**:
  - Campos de fecha Desde/Hasta (formato DD.MM.YYYY)
  - Botón **⚡ Descargar Sociedades Normales**: descarga completa para sociedades de volumen estándar
  - Grupo **Descarga Sociedades Grandes**: descarga FBL1N/FBL5N completos y FBL3N en bloques de documentos configurables
  - Grupo **Consolidación**: selector de modo (Manual / Automático) y botón **📊 Conciliación / Consolidación**
- **Pestaña Configuración**:
  - Lista de sociedades a procesar
  - Rangos de cuentas para FBL1N y FBL5N (configurables desde la GUI)
  - Rutas de trabajo (entrada proveedores, entrada clientes, salida)
  - Tablas editables de cuentas proveedores y clientes por sociedad, con persistencia en archivos `.txt` dentro de `config/`
- Barra de estado en el footer con progreso en tiempo real

### 3. `controller.py`
- **Responsabilidad**: Lógica de negocio y orquestación de procesos
- **Métodos principales**:
  - `execute_download()`: descarga normal (sin chunking)
  - `execute_download_large()`: descarga de sociedades grandes con FBL3N por bloques de documentos
  - `execute_consolidation()`: ejecuta el proceso de consolidación
  - `_run_chunked_download()`: flujo completo de descarga grande por sociedad
  - `_dividir_en_bloques()`: divide una lista de documentos en sub-listas de tamaño configurable
  - `_apilar_chunks()`: concatena los bloques de FBL3N en un único archivo final
  - `_reset_sesion_sap()`: navega SAP a `/n` para liberar memoria entre el bloque de proveedores y el de clientes
  - `_cerrar_workbook_excel()` / `_crear_excel_vacio()`: helpers de manejo de archivos

### 4. `DescargaSAP.py`
- **Responsabilidad**: Interacción con SAP GUI Scripting vía `win32com`
- **Funciones**:
  - `FBL1N_Intercompañias()`: descarga reporte de proveedores; rango de cuentas configurable
  - `ZFIQ02_Intercompañias()`: descarga catálogo de proveedores
  - `FBL3N()`: descarga reporte de cuentas contables a partir de una lista de números de documento
  - `FBL5_Intercompañias()`: descarga reporte de clientes; rango de cuentas configurable
  - `_verificar_sin_partidas()`: detecta mensajes `MSITEM030` / `MSITEM033` y crea archivo vacío placeholder
  - `_cerrar_popup_subsidiaria()`: cierra el popup de subsidiaria/central de FBL5N
  - `esperar_archivo()`: polling con timeout para confirmar que el archivo fue generado en disco

### 5. `Consolidacion_V2.py`
- **Responsabilidad**: Procesamiento y consolidación de datos
- **Función principal**: `ejecutar_consolidacion_por_sociedad()`
  - Soporta modo **Manual** (filtra por el diccionario de cuentas recibido) y **Automático** (usa todas las cuentas del FBL3N sin filtro)
  - Lee `Sociedad_Nombre.xls` y `Cuentas_Desc.xls` opcionales para enriquecer las matrices
  - Si `sin_proveedores=True` o `sin_clientes=True`, genera matrices vacías en lugar de fallar
  - Genera archivo consolidado multi-hoja por sociedad

## 📦 Requisitos

### Software necesario
- Python 3.8 o superior
- SAP GUI con Scripting habilitado
- Microsoft Excel (debe estar instalado; algunas funciones lo usan como intermediario COM)

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

### Inicio de la aplicación

```bash
python main.py
```

---

### Flujo de Trabajo — Pestaña Proceso

#### 1. Configurar fechas
- Fecha Desde y Fecha Hasta en formato `DD.MM.YYYY`
- La validación impide que "Desde" sea mayor que "Hasta"

#### 2. Agregar sociedades (Pestaña Configuración → sección Sociedades)
- Ingresar código de sociedad (ej: `MX73`, `MX30`) y hacer clic en **Agregar** o presionar Enter
- Se puede eliminar cualquier sociedad de la lista antes de ejecutar

#### 3a. Descarga normal — ⚡ Descargar Sociedades Normales
Para sociedades con volumen de documentos manejable.

**Orden de descarga por sociedad:**
1. FBL1N completo → `FBL1_Proveedores_{sociedad}.xlsx`
2. ZFIQ02 completo → `ZFIQ02_Proveedores_{sociedad}.xlsx`
3. FBL3N Proveedores (lista de docs del FBL1N) → `FBL3N_Proveedores_{sociedad}.xlsx`
4. FBL5N completo → `FBL5N_Clientes_{sociedad}.xlsx`
5. FBL3N Clientes (lista de docs del FBL5N) → `FBL3N_Clientes_{sociedad}.xlsx`

Si FBL1N o FBL5N no tienen movimientos, se crea un archivo vacío como placeholder y se omite el FBL3N correspondiente.

#### 3b. Descarga grandes — 🏢 Descargar Sociedades Grandes
Para sociedades con cientos de miles de documentos donde FBL3N completo agota la memoria de SAP.

**Diferencia clave:** FBL1N y FBL5N se descargan completos igual que antes. **Solo FBL3N** se divide en bloques de N documentos únicos para evitar el crash de sesión.

**Configuración de tamaño de bloque** (visible en la pantalla, encima del botón):
- **Docs por bloque — Proveedores**: default `15,000` (ajustable)
- **Docs por bloque — Clientes**: default `1,000` (ajustable)

**Flujo interno por sociedad:**
1. FBL1N completo (guardado en carpeta temporal `_chunks_{sociedad}/`)
2. ZFIQ02 completo
3. FBL3N Proveedores en bloques → archivos `FBL3N_Proveedores_{sociedad}_bloque001.xlsx`, `002`, …
4. Apilar todos los bloques de proveedores → `FBL3N_Proveedores_{sociedad}.xlsx` final
5. **Reset de sesión SAP** (`/n`) + pausa de 15 s para liberar memoria acumulada
6. FBL5N completo
7. FBL3N Clientes en bloques → archivos `FBL3N_Clientes_{sociedad}_bloque001.xlsx`, `002`, …
8. Apilar todos los bloques de clientes → `FBL3N_Clientes_{sociedad}.xlsx` final
9. Copiar FBL1N, ZFIQ02 y FBL5N a sus rutas finales
10. Eliminar carpetas temporales `_chunks_{sociedad}/`

> **¿Por qué el reset de sesión?**  
> SAP acumula estado interno (memoria de sesión) con cada bloque descargado. Sin el reset, el primer bloque de FBL3N Clientes crashea porque la sesión arrastra el peso de todos los bloques de proveedores. Navegar a `/n` limpia ese estado antes de continuar.

#### 4. Consolidación — 📊 Conciliación / Consolidación

**Modos de cuentas (selector de radio):**

| Modo | Comportamiento |
|------|----------------|
| **Manual** | Filtra FBL3N usando solo las cuentas configuradas en la pestaña Configuración para cada sociedad |
| **Automático** | Usa **todas** las cuentas presentes en los archivos FBL3N descargados, sin ningún filtro |

**Archivos de entrada requeridos por sociedad:**
- `Input/Proveedores/FBL1_Proveedores_{sociedad}.xlsx`
- `Input/Proveedores/ZFIQ02_Proveedores_{sociedad}.xlsx`
- `Input/Proveedores/FBL3N_Proveedores_{sociedad}.xlsx`
- `Input/Clientes/FBL5N_Clientes_{sociedad}.xlsx`
- `Input/Clientes/FBL3N_Clientes_{sociedad}.xlsx`
- `Input/Clientes/Clientes_Catalogo.xls` ← colocar manualmente (requerido)
- `Input/Sociedad_Nombre.xls` → agrega el nombre de la empresa en la fila de encabezado de las matrices (requerido)
- `Input/Cuentas_Desc.xls` → agrega la descripción de cada cuenta contable en la fila de encabezado (requerido)


**Archivo generado por sociedad en `Output/`:**
`Intercompanias_Consolidado_{sociedad}.xlsx` con 8 hojas:
- `FBL1N`
- `Cat Proveedores`
- `FBL3N Proveedores`
- `Matriz Proveedores`
- `FBL5N`
- `Cat Clientes`
- `FBL3N Clientes`
- `Matriz Clientes`

---

### Flujo de Trabajo — Pestaña Configuración

#### Sociedades y rangos de cuentas
- Lista de sociedades para los procesos de descarga
- **Rango FBL1N**: cuentas de proveedores a consultar (default: `4000000000` – `7399999999`)
- **Rango FBL5N**: cuentas de clientes a consultar (default: `200000` – `299999`)

#### Rutas de trabajo
Permite cambiar las carpetas de entrada (proveedores y clientes) y salida sin tocar el código. Las rutas se adaptan automáticamente al usuario del sistema al iniciar la app.

#### Cuentas proveedores / clientes por sociedad
Tablas editables donde se define qué cuentas contables se usan en la consolidación **Manual** para cada sociedad.

- **Agregar / Actualizar**: ingresa la sociedad y las cuentas separadas por coma → actualiza el diccionario en memoria
- **Eliminar Seleccionada**: borra la sociedad del diccionario
- **Guardar TXT Proveedores / Guardar TXT Clientes**: persiste el diccionario actual en:
  - `config/cuentas_proveedores_por_sociedad.txt`
  - `config/cuentas_clientes_por_sociedad.txt`

Al iniciar la aplicación, estos archivos `.txt` se cargan automáticamente. Si no existen, se usan los valores hardcoded por defecto.

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
        │   ├── Clientes\
        │   │   ├── FBL5N_Clientes_{sociedad}.xlsx
        │   │   ├── FBL3N_Clientes_{sociedad}.xlsx
        │   │   └── Clientes_Catalogo.xls   ← colocar manualmente (requerido)
        │   ├── Sociedad_Nombre.xls         ← colocar manualmente (requerido)
        │   └── Cuentas_Desc.xls            ← colocar manualmente (requerido)
        ├── Output\
        │   └── Intercompanias_Consolidado_{sociedad}.xlsx
        └── config\
            ├── cuentas_proveedores_por_sociedad.txt
            └── cuentas_clientes_por_sociedad.txt
```

> La estructura de carpetas (excepto los archivos manuales) se crea automáticamente al iniciar la aplicación.

## 🔍 Detalles Técnicos

### Manejo de sesión SAP y memoria

- Todas las transacciones se prefijan con `/n` (ej: `/nFBL1`, `/nFBL3N`) para forzar una navegación limpia.
- Entre el bloque de descargas de proveedores y el de clientes (modo grandes), se ejecuta `_reset_sesion_sap()` que navega a `/n` y espera 15 segundos para que SAP libere el estado acumulado.
- Los `time.sleep()` entre pasos son deliberados: SAP GUI Scripting no es síncrono y necesita tiempo para procesar exportaciones grandes.

### Chunking de FBL3N

- Solo FBL3N se divide; FBL1N y FBL5N siempre se descargan completos.
- Los bloques se generan dividiendo la lista de números de documento únicos del FBL1N/FBL5N en sub-listas de tamaño `chunk_prov` / `chunk_cli`.
- La verificación de éxito usa `fbl3n_p_ok and os.path.exists(path_p)` para ser robusta ante retornos `None` o `False`.
- El apilado (`_apilar_chunks`) ignora archivos marcados como "Sin datos" (columna única con ese nombre) y los bloques vacíos.

### Detección de errores SAP

| Error SAP | Función de detección | Acción |
|-----------|----------------------|--------|
| `MSITEM030` / `MSITEM033` — sin partidas | `_verificar_sin_partidas()` | Crea archivo vacío, retorna `False` |
| Popup subsidiaria/central (FBL5N) | `_cerrar_popup_subsidiaria()` | Presiona Continuar hasta que desaparezca |

### Flags de estado por sociedad

Después de cada descarga (normal o grandes) se guarda un archivo `_flags_{sociedad}.json` en la carpeta de proveedores con las claves `sin_proveedores` y `sin_clientes`. La consolidación lo lee para saber si debe generar matrices vacías en lugar de intentar leer archivos inexistentes.

### Persistencia de configuración

Las cuentas por sociedad se persisten como texto Python-eval-safe en archivos `.txt` dentro de `config/`. Al cargar, se usa `ast.literal_eval()` para parsear de forma segura sin `eval()`.

### Proceso de Consolidación — detalle

1. **Carga** de FBL1N, ZFIQ02, FBL3N (proveedores) y FBL5N, Catálogo Clientes, FBL3N Clientes
2. **Limpieza** de tabs, saltos de línea y espacios en todas las columnas (`clean_all_str`)
3. **Enriquecimiento**: merge FBL1N ↔ ZFIQ02 para agregar `Nombre 1` del proveedor; FBL3N ↔ FBL1N para agregar `Proveedor`, `Texto` y `Texto cab.documento`
4. **Filtrado por cuentas**: modo Manual aplica el diccionario de cuentas; modo Automático conserva todo
5. **Pivot / Matriz**: `groupby` + `pivot` de importe por proveedor/cliente y cuenta, con filas de Totales, Cuadre Balanza y variaciones
6. **Exportación** multi-hoja a `Intercompanias_Consolidado_{sociedad}.xlsx`

### Cuentas contables por sociedad — valores por defecto

Estos valores se usan en modo Manual cuando no hay archivo `.txt` guardado.

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

## ⚠️ Consideraciones Importantes

### SAP GUI Scripting
- Debe estar habilitado en SAP GUI (Opciones → Accesibilidad y Scripting)
- SAP debe estar abierto y con una sesión activa durante todo el proceso de descarga
- No interactuar con SAP mientras la descarga está en curso

### Archivos de entrada
- Los nombres de archivo son exactos y sensibles a mayúsculas/minúsculas
- `Clientes_Catalogo.xls` debe colocarse manualmente en `Input/Clientes/` antes de consolidar
- Los archivos (`Sociedad_Nombre.xls`, `Cuentas_Desc.xls`) deben ir directamente en `Input/` (padre de Proveedores/ y Clientes/)

### Performance
- La descarga normal puede tardar varios minutos por sociedad según el volumen
- La descarga grande puede tardar decenas de minutos: FBL3N se ejecuta N veces (una por bloque)
- La consolidación es rápida (< 1 minuto por sociedad)
- No cerrar la aplicación durante la ejecución

## 🐛 Solución de Problemas

### "Archivos no encontrados" al consolidar
- Verificar que existan todos los archivos de entrada en las rutas correctas
- Ejecutar primero el proceso de descarga
- Confirmar que `Clientes_Catalogo.xls` esté en `Input/Clientes/`

### Crash en el primer bloque de FBL3N Clientes (modo grandes)
- Síntoma: los bloques de proveedores terminan bien pero el primer bloque de clientes falla
- Causa: memoria SAP acumulada entre bloques
- Solución: el sistema ya inserta un reset de sesión automático + 15 s de pausa entre proveedores y clientes; si el problema persiste, aumentar el `time.sleep` en `_reset_sesion_sap()` dentro de `controller.py`

### SAP devuelve "Sin partidas" para una sociedad con movimientos
- Revisar el rango de cuentas FBL1N / FBL5N en Configuración → puede estar demasiado acotado
- Verificar que el periodo de fechas sea correcto

### Error en SAP (COM error)
- Verificar que SAP GUI Scripting esté habilitado
- Confirmar que la sesión SAP no esté bloqueada (login expirado, popup de error)
- Revisar que las variantes `PYTHON` (FBL5N) y `RDA_FBL3N` (FBL3N) existan en el sistema

### Error: "Formato de fecha inválido"
- Usar el formato `DD.MM.YYYY` exacto, ej: `01.01.2025`

### Sociedad sin datos en la matriz
- En modo Manual: verificar que la sociedad esté en la tabla de cuentas de la pestaña Configuración
- Si no está configurada se aplica el default (`6600022` para proveedores, `7201000` para clientes)
- En modo Automático: si la matriz sale vacía, revisar que el FBL3N descargado tenga datos

## 📝 Notas de Desarrollo

### Agregar una nueva sociedad con cuentas específicas

**Vía GUI (recomendado):**
1. Abrir la app → Pestaña Configuración
2. En la sección correspondiente (Proveedores o Clientes), ingresar la sociedad y sus cuentas separadas por coma
3. Clic en **Agregar / Actualizar**
4. Clic en **Guardar TXT** para persistir el cambio

**Vía código (hardcoded defaults):**
1. Editar `interfaz_GUI.py` → diccionarios `cuentas_proveedores_por_sociedad` y `cuentas_clientes_por_sociedad`
2. Estos valores solo aplican cuando no existe el archivo `.txt` guardado

### Ajustar tamaño de bloque para FBL3N grandes

Los defaults de clase están en `controller.py`:
```python
_BLOQUE_PROV_DEFAULT = 2500
_BLOQUE_CLI_DEFAULT  = 500
```
Pero se pueden cambiar en tiempo de ejecución desde los campos de la GUI sin tocar el código.

### Mantenimiento
- Revisar cuentas por sociedad si cambian las estructuras contables → actualizar vía GUI y guardar TXT
- Actualizar nombres de columnas en `Consolidacion_V2.py` si SAP cambia la estructura de los reportes
- Mantener sincronizadas las versiones de `pandas` / `openpyxl` en `requirements.txt`

## 📄 Licencia

Uso interno — Todos los derechos reservados
