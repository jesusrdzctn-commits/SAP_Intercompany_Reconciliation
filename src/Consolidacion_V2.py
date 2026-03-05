import pandas as pd
import os


# Helper simple
def clean_all_str(df):
    """Limpia tabs, saltos de línea y espacios en TODAS las columnas."""
    for col in df.columns:
        df[col] = df[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    return df


def ejecutar_consolidacion_por_sociedad(ruta_input, ruta_output, sociedad, callback_status=None):
    """
    Ejecuta el proceso de consolidación para UNA sociedad específica.
    
    Args:
        ruta_input (str): Ruta de la carpeta de entrada
        ruta_output (str): Ruta de la carpeta de salida
        sociedad (str): Código de la sociedad (ej: 'MX73')
        callback_status (function, optional): Función callback para actualizar estado
    
    Returns:
        str: Ruta del archivo consolidado generado
    """
    
    def update_status(message):
        if callback_status:
            callback_status(message)
    


    # =========================
    # === Rutas y archivos ===
    # =========================

    archivo_fbl1 = os.path.join(ruta_input, 'Proveedores', f'FBL1_Proveedores_{sociedad}.xlsx')
    archivo_zfiq02 = os.path.join(ruta_input, 'Proveedores', f'ZFIQ02_Proveedores_{sociedad}.xlsx')
    archivo_fbl3 = os.path.join(ruta_input, 'Proveedores', f'FBL3N_Proveedores_{sociedad}.xlsx')
    # archivo_fbl5 = os.path.join(ruta_input, 'Clientes', f'fbl5n_{sociedad}.xlsx')
    # archivo_fbl3_clientes = os.path.join(ruta_input, 'Clientes', f'FBL3N_Clientes_{sociedad}.xlsx')
    # archivo_cat_clientes = os.path.join(ruta_input, 'Clientes', 'CLIENTES.xlsx')  # Este es común
    
    # Verificar archivos
    archivos_requeridos = [
        archivo_fbl1, archivo_zfiq02, archivo_fbl3,
        # archivo_fbl5, archivo_fbl3_clientes, archivo_cat_clientes
    ]
    
    for archivo in archivos_requeridos:
        if not os.path.exists(archivo):
            raise FileNotFoundError(f"Archivo no encontrado: {archivo}")
    
    update_status(f"📄 Procesando FBL1N - {sociedad}...")
    

    
    # =========================
    # === PROVEEDORES
    # =========================
    
    # --- Procesar FBL1N ---
    fbl1n = pd.read_excel(archivo_fbl1, dtype=str)
    fbl1n = clean_all_str(fbl1n)
    
    # Limpieza específica (equivalente a Text to Columns) y Mover Nº documento al inicio
    for col in ["Nº documento", "Cuenta"]:
        if col in fbl1n.columns:
            fbl1n[col] = fbl1n[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    if "Nº documento" in fbl1n.columns:
        cols = fbl1n.columns.tolist()
        cols = ["Nº documento"] + [c for c in cols if c != "Nº documento"]
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
    fbl1n = fbl1n.drop(columns=["Nombre 1_x"])
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
    for col in ["Nº documento", "Cuenta"]:
        if col in fbl3n.columns:
            fbl3n[col] = fbl3n[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    
    # Crear columnas nuevas
    for col in ["Concepto Intercompañias", "UUID Auditor"]:
        if col not in fbl3n.columns:
            fbl3n[col] = ""

    # Filtrar cuentas según sociedad
    #NOTA: Cambiar en un futuro las cuentas por sociedad en caso de agregar/quitar algunas cuentas
    CUENTAS_POR_SOCIEDAD = {
    "MX01": ["6600022", "7201000", "7204000"],
    "MX05": ["7201000"],
    "MX22": ["6600021", "2050000", "6600022", "6700040", "6700043", "6700048", "6900010"],
    "MX30": ["6600022", "7204000", "6900010"],
    "MX73": ["6600022"],
    }
    cuentas_filtro = CUENTAS_POR_SOCIEDAD.get(sociedad.upper(), ["6600022"])
    if "Cuenta" in fbl3n.columns:
        fbl3n = fbl3n[fbl3n["Cuenta"].isin(cuentas_filtro)]

    # Determinar columnas de texto (por nombre; si no, por índice 19/20 como fallback)
    if "Texto" in fbl1n.columns:
        col_texto = "Texto"
    else:
        col_texto = fbl1n.columns[19] if len(fbl1n.columns) > 19 else fbl1n.columns[-1]
    if "Texto cab.documento" in fbl1n.columns:
        col_texto_cab = "Texto cab.documento"
    else:
        col_texto_cab = fbl1n.columns[20] if len(fbl1n.columns) > 20 else fbl1n.columns[-1]
    
    # Lookup #1 Proveedor (Nº documento → Nombre 1)
    if "Nº documento" in fbl1n.columns and "Nombre 1" in fbl1n.columns:
        fbl3n = fbl3n.merge(
            fbl1n[["Nº documento", "Nombre 1"]].rename(columns={"Nombre 1": "Proveedor"}),
            how="left",
            left_on="Nº documento",
            right_on="Nº documento"
        )
    
    # Lookup #2 Texto
    if "Nº documento" in fbl1n.columns and col_texto in fbl1n.columns:
        fbl3n = fbl3n.merge(
            fbl1n[["Nº documento", col_texto]].rename(columns={col_texto: "Texto"}),
            how="left",
            left_on="Nº documento",
            right_on="Nº documento"
        )
    else:
        fbl3n["Texto"] = ""
    
    fbl3n = fbl3n.drop(columns=["Texto_x"])
    fbl3n = fbl3n.rename(columns={"Texto_y": "Texto"})

    # Lookup #3 Texto cab.documento
    if "Nº documento" in fbl1n.columns and col_texto_cab in fbl1n.columns:
        fbl3n = fbl3n.merge(
            fbl1n[["Nº documento", col_texto_cab]].rename(columns={col_texto_cab: "Texto cab.documento"}),
            how="left",
            left_on="Nº documento",
            right_on="Nº documento"
        )
    else:
        fbl3n["Texto cab.documento"] = ""
    
    #fbl3n = fbl3n.drop(columns=["Texto cab.documento_x"])
    fbl3n = fbl3n.rename(columns={"Texto cab.documento_y": "Texto cab.documento"})

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
    
    '''
    '''

    
    # =========================
    # === Exportar Excel consolidado con todas las hojas ===
    # =========================
    
    update_status("💾 Guardando archivo consolidado...")
    
    # Crear carpeta de salida si no existe
    os.makedirs(ruta_output, exist_ok=True)
    
    archivo_consolidado = os.path.join(ruta_output, f"Intercompanias_Consolidado_{sociedad}.xlsx")
    
    with pd.ExcelWriter(archivo_consolidado, engine="openpyxl") as writer:
        # Proveedores
        fbl1n.to_excel(writer, sheet_name="FBL1N", index=False)
        cat.to_excel(writer, sheet_name="Cat Proveedores", index=False)
        fbl3n.to_excel(writer, sheet_name="FBL3N Proveedores", index=False)
        matriz_prov.to_excel(writer, sheet_name="Matriz Proveedores", index=False)
        
        '''
        # Clientes
        fbl5n.to_excel(writer, sheet_name="FBL5N", index=False)
        clientes.to_excel(writer, sheet_name="Cat Clientes", index=False)
        fbl3n_cli.to_excel(writer, sheet_name="FBL3N Clientes", index=False)
        matriz_cli.to_excel(writer, sheet_name="Matriz Clientes", index=False)
        '''
    
    update_status(f"✅ Consolidación completada - {sociedad}")
    
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
        archivo_generado = ejecutar_consolidacion_por_sociedad(
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