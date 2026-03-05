    # =========================
    # === CLIENTES
    # =========================
    
    update_status("📄 Procesando archivos de clientes...")
    
    # --- Procesar FBL5N ---
    fbl5n = pd.read_excel(archivo_fbl5, dtype=str)
    fbl5n = clean_all_str(fbl5n)
    for col in ["Nº documento", "Cuenta"]:
        if col in fbl5n.columns:
            fbl5n[col] = fbl5n[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    
    # Mover Nº documento al inicio
    if "Nº documento" in fbl5n.columns:
        cols5 = fbl5n.columns.tolist()
        cols5 = ["Nº documento"] + [c for c in cols5 if c != "Nº documento"]
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
    for col in ["Nº documento", "Cuenta"]:
        if col in fbl3n_cli.columns:
            fbl3n_cli[col] = fbl3n_cli[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    
    # === Crear columnas nuevas ===
    for col in ["Concepto Intercompañias", "UUID Auditor"]:
        fbl3n_cli[col] = ""
    
    # Filtrar por cuentas 7000005, 7001002, 7201000
    cuentas_clientes_filtrar = ["7000005", "7001002", "7201000"]
    fbl3n_cli = fbl3n_cli[fbl3n_cli["Cuenta"].isin(cuentas_clientes_filtrar)]
    
    # Eliminar filas con Nº documento vacío en FBL5N
    fbl5n["Nº documento"] = fbl5n["Nº documento"].astype(str).str.strip()
    fbl5n = fbl5n[
        fbl5n["Nº documento"].notna() &
        (fbl5n["Nº documento"] != "") &
        (fbl5n["Nº documento"].str.lower() != "nan")
    ].copy()
    
    col_texto2 = "Texto" 
    col_texto_cab2 = "Asignación"
    
    # === Lookup #1 Cliente ===
    fbl3n_cli = fbl3n_cli.merge(
        fbl5n[["Nº documento", "Nombre"]].rename(columns={"Nombre": "Cliente"}),
        how="left",
        left_on="Nº documento",
        right_on="Nº documento"
    )
    
    # === Lookup #2 Texto ===
    fbl3n_cli = fbl3n_cli.merge(
        fbl5n[["Nº documento", col_texto2]].rename(columns={col_texto2: "Texto"}),
        how="left",
        left_on="Nº documento",
        right_on="Nº documento"
    )
    
    # === Lookup #3 Texto cab.documento ===
    fbl3n_cli = fbl3n_cli.merge(
        fbl5n[["Nº documento", col_texto_cab2]].rename(columns={col_texto_cab2: "Texto cab.documento"}),
        how="left",
        left_on="Nº documento",
        right_on="Nº documento"
    )
    fbl3n_cli = fbl3n_cli[fbl3n_cli["Nº documento"].notna() & (fbl3n_cli["Nº documento"] != "") & (fbl3n_cli["Nº documento"].str.lower() != "nan")]
    
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