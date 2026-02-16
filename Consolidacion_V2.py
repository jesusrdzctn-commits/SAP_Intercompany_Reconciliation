import pandas as pd
import os


# Helper simple
def clean_all_str(df):
    """Limpia tabs, saltos de línea y espacios en TODAS las columnas."""
    for col in df.columns:
        df[col] = df[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    return df


def ejecutar_consolidacion(ruta_input, ruta_output, callback_status=None):
    """
    Ejecuta el proceso completo de consolidación de intercompañías.
    
    Args:
        ruta_input (str): Ruta de la carpeta de entrada con los archivos fuente
        ruta_output (str): Ruta de la carpeta de salida para el archivo consolidado
        callback_status (function, optional): Función callback para actualizar estado (ej: lambda msg: print(msg))
    
    Returns:
        str: Ruta del archivo consolidado generado
    
    Raises:
        FileNotFoundError: Si algún archivo requerido no existe
        Exception: Si ocurre algún error durante el procesamiento
    """
    
    def update_status(message):
        """Helper para actualizar estado si hay callback"""
        if callback_status:
            callback_status(message)
    
    # =========================
    # === Rutas y archivos ===
    # =========================
    archivo_fbl1 = os.path.join(ruta_input, 'Proveedores', 'FBL1_Intercompañias 1.xlsx')
    archivo_zfiq02 = os.path.join(ruta_input, 'Proveedores', 'ZFIQ02_Intercompañias.xlsx')
    archivo_fbl3 = os.path.join(ruta_input, 'Proveedores', 'FBL301-10.xlsx')
    archivo_fbl5 = os.path.join(ruta_input, 'Clientes', 'fbl5n 4.xlsx')
    archivo_fbl3_clientes = os.path.join(ruta_input, 'Clientes', 'FBL3N Clientes.xlsx')
    archivo_cat_clientes = os.path.join(ruta_input, 'Clientes', 'CLIENTES.xlsx')
    
    # Verificar que existan los archivos requeridos
    archivos_requeridos = [
        archivo_fbl1, archivo_zfiq02, archivo_fbl3,
        archivo_fbl5, archivo_fbl3_clientes, archivo_cat_clientes
    ]
    
    for archivo in archivos_requeridos:
        if not os.path.exists(archivo):
            raise FileNotFoundError(f"Archivo no encontrado: {archivo}")
    
    # =========================
    # === PROVEEDORES
    # =========================
    
    update_status("📄 Procesando FBL1N...")
    
    # --- Procesar FBL1N ---
    fbl1n = pd.read_excel(archivo_fbl1, dtype=str)
    fbl1n = clean_all_str(fbl1n)
    
    # Limpieza específica (equivalente a Text to Columns) y Mover N° documento al inicio
    for col in ["N° documento", "Cuenta"]:
        if col in fbl1n.columns:
            fbl1n[col] = fbl1n[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    if "N° documento" in fbl1n.columns:
        cols = fbl1n.columns.tolist()
        cols = ["N° documento"] + [c for c in cols if c != "N° documento"]
        fbl1n = fbl1n[cols]
    
    update_status("📄 Procesando catálogo de proveedores...")
    
    # --- Procesar ZFIQ02 (Catálogo de proveedores) ---
    cat = pd.read_excel(archivo_zfiq02, dtype=str)
    cat = clean_all_str(cat)
    cat_reducido = cat[["Acreedor", "Nombre 1"]].drop_duplicates()
    
    # --- Merge catálogo → FBL1N (VLOOKUP Cuenta ↔ Acreedor) ---
    fbl1n = fbl1n.merge(
        cat_reducido,
        how="left",
        left_on="Cuenta",
        right_on="Acreedor",
        suffixes=("_x", "_y")
    )
    
    # Limpiar duplicadas y renombrar
    fbl1n = fbl1n.drop(columns=["Acreedor_y"], errors="ignore")
    fbl1n = fbl1n.rename(columns={"Acreedor_x": "Acreedor"})
    if "Nombre 1_x" in fbl1n.columns:
        fbl1n = fbl1n.drop(columns=["Nombre 1_x"])
    if "Nombre 1_y" in fbl1n.columns:
        fbl1n = fbl1n.rename(columns={"Nombre 1_y": "Nombre 1"})
    
    # Reordenar colocando "Nombre 1" cerca del inicio
    if "Nombre 1" in fbl1n.columns:
        columnas = fbl1n.columns.tolist()
        columnas.remove("Nombre 1")
        insert_pos = 4 if len(columnas) >= 4 else len(columnas)
        columnas.insert(insert_pos, "Nombre 1")
        fbl1n = fbl1n[columnas]
    
    update_status("📄 Procesando FBL3N Proveedores...")
    
    # --- Procesar FBL3N (Proveedores) ---
    fbl3n = pd.read_excel(archivo_fbl3, dtype=str)
    fbl3n = clean_all_str(fbl3n)
    for col in ["N° documento", "Cuenta"]:
        if col in fbl3n.columns:
            fbl3n[col] = fbl3n[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    
    # Crear columnas nuevas
    for col in ["Concepto Intercompañias", "UUID Auditor"]:
        if col not in fbl3n.columns:
            fbl3n[col] = ""
    # Filtrar Cuenta = 6600022
    if "Cuenta" in fbl3n.columns:
        fbl3n = fbl3n[fbl3n["Cuenta"] == "6600022"]
    # Determinar columnas de texto (por nombre; si no, por índice 19/20 como fallback)
    if "Texto" in fbl1n.columns:
        col_texto = "Texto"
    else:
        col_texto = fbl1n.columns[19] if len(fbl1n.columns) > 19 else fbl1n.columns[-1]
    if "Texto cab.documento" in fbl1n.columns:
        col_texto_cab = "Texto cab.documento"
    else:
        col_texto_cab = fbl1n.columns[20] if len(fbl1n.columns) > 20 else fbl1n.columns[-1]
    
    # Lookup #1 Proveedor (N° documento → Nombre 1)
    if "N° documento" in fbl1n.columns and "Nombre 1" in fbl1n.columns:
        fbl3n = fbl3n.merge(
            fbl1n[["N° documento", "Nombre 1"]].rename(columns={"Nombre 1": "Proveedor"}),
            how="left",
            left_on="N° documento",
            right_on="N° documento"
        )
    
    # Lookup #2 Texto
    if "N° documento" in fbl1n.columns and col_texto in fbl1n.columns:
        fbl3n = fbl3n.merge(
            fbl1n[["N° documento", col_texto]].rename(columns={col_texto: "Texto"}),
            how="left",
            left_on="N° documento",
            right_on="N° documento"
        )
    else:
        fbl3n["Texto"] = ""
    
    # Lookup #3 Texto cab.documento
    if "N° documento" in fbl1n.columns and col_texto_cab in fbl1n.columns:
        fbl3n = fbl3n.merge(
            fbl1n[["N° documento", col_texto_cab]].rename(columns={col_texto_cab: "Texto cab.documento"}),
            how="left",
            left_on="N° documento",
            right_on="N° documento"
        )
    else:
        fbl3n["Texto cab.documento"] = ""
    
    orden_deseado = ["Proveedor", "Texto", "Texto cab.documento", "Concepto Intercompañias", "UUID Auditor"]
    presentes = [c for c in orden_deseado if c in fbl3n.columns]
    resto = [c for c in fbl3n.columns if c not in presentes]
    fbl3n = fbl3n[resto + presentes]
    
    # Llenar Concepto Intercompañias (Proveedores)
    fbl3n["Concepto Intercompañias"] = "Servicios Administrativos"
    
    # =========================
    # === MATRIZ PROVEEDORES
    # =========================
    
    update_status("📊 Creando matriz de proveedores...")
    
    # Normalizar Importe en moneda local
    if "Importe en moneda local" in fbl3n.columns:
        fbl3n["Importe en moneda local"] = (
            fbl3n["Importe en moneda local"].astype(str).str.replace(",", "", regex=False)
        )
        fbl3n["Importe en moneda local"] = pd.to_numeric(
            fbl3n["Importe en moneda local"], errors="coerce"
        ).fillna(0)
    else:
        fbl3n["Importe en moneda local"] = 0.0
    cuentas_prov = (
        fbl3n["Cuenta"].astype(str).str.strip().dropna().unique().tolist()
        if "Cuenta" in fbl3n.columns else []
    )
    proveedores_list = (
        fbl3n["Proveedor"].astype(str).str.strip().replace({"": None}).dropna().unique().tolist()
        if "Proveedor" in fbl3n.columns else []
    )
    if "Proveedor" in fbl3n.columns and "Cuenta" in fbl3n.columns:
        pivot_prov = (
            fbl3n.groupby(["Proveedor", "Cuenta"], as_index=False)["Importe en moneda local"]
                 .sum()
                 .pivot(index="Proveedor", columns="Cuenta", values="Importe en moneda local")
        )
    else:
        pivot_prov = pd.DataFrame()
    
    pivot_prov = pivot_prov.reindex(index=proveedores_list, columns=cuentas_prov).fillna(0)
    
    totales_prov = pivot_prov.sum(axis=0).to_frame().T
    totales_prov.index = ["Totales"]
    cuadre_prov = pd.DataFrame([[0] * len(cuentas_prov)], index=["Cuadre Balanza"], columns=cuentas_prov)
    variaciones_prov = pd.DataFrame([[0] * len(cuentas_prov)], index=["variaciones"], columns=cuentas_prov)
    matriz_prov = pd.concat([pivot_prov, totales_prov, cuadre_prov, variaciones_prov], axis=0)
    matriz_prov.insert(0, "Proveedores", matriz_prov.index)
    matriz_prov = matriz_prov.reset_index(drop=True)
    
    # =========================
    # === CLIENTES
    # =========================
    
    update_status("📄 Procesando archivos de clientes...")
    
    # --- Procesar FBL5N ---
    fbl5n = pd.read_excel(archivo_fbl5, dtype=str)
    fbl5n = clean_all_str(fbl5n)
    for col in ["N° documento", "Cuenta"]:
        if col in fbl5n.columns:
            fbl5n[col] = fbl5n[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    
    # Mover N° documento al inicio
    if "N° documento" in fbl5n.columns:
        cols5 = fbl5n.columns.tolist()
        cols5 = ["N° documento"] + [c for c in cols5 if c != "N° documento"]
        fbl5n = fbl5n[cols5]
    
    # --- Catálogo CLIENTES.xlsx ---
    clientes = pd.read_excel(archivo_cat_clientes, dtype=str)
    clientes = clean_all_str(clientes)
    
    # Limpieza tipo Text to Columns sobre 'numero'
    if "numero" in clientes.columns:
        clientes["numero"] = clientes["numero"].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    
    # Reducir y VLOOKUP Cuenta (FBL5N) ↔ numero (CLIENTES) → nombre
    if "numero" in clientes.columns and "nombre" in clientes.columns:
        cat_clientes_reducido = clientes[["numero", "nombre"]].drop_duplicates()
        fbl5n = fbl5n.merge(
            cat_clientes_reducido,
            how="left",
            left_on="Cuenta",
            right_on="numero",
            suffixes=("", "_catcli")
        )
    else:
        fbl5n["Nombre"] = ""
    
    # Eliminar columna Nombre original (NaN) y la llave auxiliar
    fbl5n.drop(columns=["Nombre", "numero"], inplace=True, errors="ignore")
    # Renombrar columna correcta desde el catálogo
    fbl5n.rename(columns={"nombre": "Nombre"}, inplace=True)
    if "Nombre" in fbl5n.columns:
        col_nombre = fbl5n.pop("Nombre")
        fbl5n.insert(3, "Nombre", col_nombre)
    
    update_status("📄 Procesando FBL3N Clientes...")
    
    # --- Procesar FBL3N (Clientes) ---
    fbl3n_cli = pd.read_excel(archivo_fbl3_clientes, dtype=str)
    fbl3n_cli = clean_all_str(fbl3n_cli)
    for col in ["N° documento", "Cuenta"]:
        if col in fbl3n_cli.columns:
            fbl3n_cli[col] = fbl3n_cli[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    
    # === Crear columnas nuevas ===
    for col in ["Concepto Intercompañias", "UUID Auditor"]:
        fbl3n_cli[col] = ""
    
    # Filtrar por cuentas 7000005, 7001002, 7201000
    cuentas_clientes_filtrar = ["7000005", "7001002", "7201000"]
    fbl3n_cli = fbl3n_cli[fbl3n_cli["Cuenta"].isin(cuentas_clientes_filtrar)]
    
    # Eliminar filas con N° documento vacío en FBL5N
    fbl5n["N° documento"] = fbl5n["N° documento"].astype(str).str.strip()
    fbl5n = fbl5n[
        fbl5n["N° documento"].notna() &
        (fbl5n["N° documento"] != "") &
        (fbl5n["N° documento"].str.lower() != "nan")
    ].copy()
    
    col_texto2 = "Texto" 
    col_texto_cab2 = "Asignación"
    
    # === Lookup #1 Cliente ===
    fbl3n_cli = fbl3n_cli.merge(
        fbl5n[["N° documento", "Nombre"]].rename(columns={"Nombre": "Cliente"}),
        how="left",
        left_on="N° documento",
        right_on="N° documento"
    )
    
    # === Lookup #2 Texto ===
    fbl3n_cli = fbl3n_cli.merge(
        fbl5n[["N° documento", col_texto2]].rename(columns={col_texto2: "Texto"}),
        how="left",
        left_on="N° documento",
        right_on="N° documento"
    )
    
    # === Lookup #3 Texto cab.documento ===
    fbl3n_cli = fbl3n_cli.merge(
        fbl5n[["N° documento", col_texto_cab2]].rename(columns={col_texto_cab2: "Texto cab.documento"}),
        how="left",
        left_on="N° documento",
        right_on="N° documento"
    )
    fbl3n_cli = fbl3n_cli[fbl3n_cli["N° documento"].notna() & (fbl3n_cli["N° documento"] != "") & (fbl3n_cli["N° documento"].str.lower() != "nan")]
    
    # === Reordenar columnas FBL3N ===
    orden_deseado2 = ["Cliente", "Texto", "Texto cab.documento", "Concepto Intercompañias", "UUID Auditor"]
    presentes2 = [c for c in orden_deseado2 if c in fbl3n_cli.columns]
    resto2 = [c for c in fbl3n_cli.columns if c not in presentes2]
    fbl3n_cli = fbl3n_cli[resto2 + presentes2]
    
    # Sobrescrituras según Cuenta
    if "Cuenta" in fbl3n_cli.columns:
        fbl3n_cli.loc[fbl3n_cli["Cuenta"] == "7000005", "Concepto Intercompañias"] = "Servicios"
        fbl3n_cli.loc[fbl3n_cli["Cuenta"] == "7001002", "Concepto Intercompañias"] = "Arrendamiento Inmuebles"
        fbl3n_cli.loc[fbl3n_cli["Cuenta"] == "7201000", "Concepto Intercompañias"] = "Intereses"
    
    # =========================
    # === MATRIZ CLIENTES
    # =========================
    
    update_status("📊 Creando matriz de clientes...")
    
    if "Importe en moneda local" in fbl3n_cli.columns:
        fbl3n_cli["Importe en moneda local"] = (
            fbl3n_cli["Importe en moneda local"]
            .astype(str)
            .str.replace(",", "", regex=False)
        )
        fbl3n_cli["Importe en moneda local"] = pd.to_numeric(
            fbl3n_cli["Importe en moneda local"],
            errors="coerce"
        ).fillna(0)
    else:
        fbl3n_cli["Importe en moneda local"] = 0.0
    
    cuentas_cli = (
        fbl3n_cli["Cuenta"].astype(str).str.strip().dropna().unique().tolist()
        if "Cuenta" in fbl3n_cli.columns else []
    )
    clientes_list = (
        fbl3n_cli["Cliente"].astype(str).str.strip().replace({"": None}).dropna().unique().tolist()
        if "Cliente" in fbl3n_cli.columns else []
    )
    
    if "Cliente" in fbl3n_cli.columns and "Cuenta" in fbl3n_cli.columns:
        pivot_cli = (
            fbl3n_cli.groupby(["Cliente", "Cuenta"], as_index=False)["Importe en moneda local"]
                     .sum()
                     .pivot(index="Cliente", columns="Cuenta", values="Importe en moneda local")
        )
    else:
        pivot_cli = pd.DataFrame()
    
    pivot_cli = pivot_cli.reindex(index=clientes_list, columns=cuentas_cli).fillna(0)
    totales_cli = pivot_cli.sum(axis=0).to_frame().T
    totales_cli.index = ["Totales"]
    cuadre_cli = pd.DataFrame([[0] * len(cuentas_cli)], index=["Cuadre Balanza"], columns=cuentas_cli)
    variaciones_cli = pd.DataFrame([[0] * len(cuentas_cli)], index=["variaciones"], columns=cuentas_cli)
    matriz_cli = pd.concat([pivot_cli, totales_cli, cuadre_cli, variaciones_cli], axis=0)
    matriz_cli.insert(0, "Clientes", matriz_cli.index)
    matriz_cli = matriz_cli.reset_index(drop=True)
    cols_numericas = matriz_cli.columns.difference(["Clientes"])
    matriz_cli[cols_numericas] = matriz_cli[cols_numericas].abs()
    
    # =========================
    # === Exportar Excel consolidado con todas las hojas ===
    # =========================
    
    update_status("💾 Guardando archivo consolidado...")
    
    # Crear carpeta de salida si no existe
    os.makedirs(ruta_output, exist_ok=True)
    
    archivo_consolidado = os.path.join(ruta_output, "Intercompanias_Consolidado.xlsx")
    
    with pd.ExcelWriter(archivo_consolidado, engine="openpyxl") as writer:
        # Proveedores
        fbl1n.to_excel(writer, sheet_name="FBL1N", index=False)
        cat.to_excel(writer, sheet_name="Cat Proveedores", index=False)
        fbl3n.to_excel(writer, sheet_name="FBL3N Proveedores", index=False)
        matriz_prov.to_excel(writer, sheet_name="Matriz Proveedores", index=False)
        
        # Clientes
        fbl5n.to_excel(writer, sheet_name="FBL5N", index=False)
        clientes.to_excel(writer, sheet_name="Cat Clientes", index=False)
        fbl3n_cli.to_excel(writer, sheet_name="FBL3N Clientes", index=False)
        matriz_cli.to_excel(writer, sheet_name="Matriz Clientes", index=False)
    
    update_status("✅ Consolidación completada")
    
    return archivo_consolidado


# =========================
# === Ejecución directa ===
# =========================
if __name__ == "__main__":
    """
    Permite ejecutar el script directamente con rutas dinámicas
    adaptadas al usuario actual del sistema
    """
    # Obtener perfil de usuario dinámicamente
    user_profile = os.environ.get('USERPROFILE') or os.path.expanduser('~')
    base_path = os.path.join(user_profile, 'Documents', 'Intercompañias', 'RDA_Intercompanias', 'src')
    
    ruta_input = os.path.join(base_path, 'Input')
    ruta_output = os.path.join(base_path, 'Output')
    
    try:
        archivo_generado = ejecutar_consolidacion(
            ruta_input, 
            ruta_output,
            callback_status=lambda msg: print(msg)
        )
        
        print("\n📘 Archivo consolidado generado: Intercompanias_Consolidado.xlsx")
        print("   - Incluye Proveedores: FBL1N, Cat Proveedores, FBL3N Proveedores, Matriz Proveedores")
        print("   - Incluye Clientes: FBL5N, Cat Clientes, FBL3N Clientes, Matriz Clientes")
        print("🎉 PROCESO COMPLETO FINALIZADO EXITOSAMENTE 🎉")
        print(f"\nArchivo guardado en: {archivo_generado}")
        
    except Exception as e:
        print(f"❌ Error durante la consolidación: {str(e)}")
        raise
