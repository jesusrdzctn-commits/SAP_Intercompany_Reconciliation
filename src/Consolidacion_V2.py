import pandas as pd
import os


# Helper simple
def clean_all_str(df):
    """Limpia tabs, saltos de línea y espacios en TODAS las columnas."""
    for col in df.columns:
        df[col] = df[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
    return df


def _matriz_sin_movimientos(col_nombre):
    """
    Devuelve un DataFrame con el formato estándar de matriz pero con una sola
    fila 'SIN MOVIMIENTOS', cuenta default 7201000 y todos los valores en cero.
    """
    cuenta_default = "7201000"
    filas = ["SIN MOVIMIENTOS", "Totales", "Cuadre Balanza", "variaciones"]
    df = pd.DataFrame({
        col_nombre: filas,
        cuenta_default: [0, 0, 0, 0],
    })
    return df


# Valores por defecto para cuentas por sociedad (modo manual sin TXT cargado)
_DEFAULT_CUENTAS_PROVEEDORES = {
    "MX01": ["6600022", "7201000", "7204000"],
    "MX05": ["7201000"],
    "MX22": ["6600021", "2050000", "6600022", "6700040", "6700043", "6700048", "6900010"],
    "MX30": ["6600022", "7204000", "6900010"],
    "MX73": ["6600022"],
}

_DEFAULT_CUENTAS_CLIENTES = {
    "MX01": ["7000005", "7000020", "7201000"],
    "MX05": ["7201000"],
    "MX22": ["4300010", "7000005", "7001002", "7010005", "7201000"],
    "MX30": ["7001000", "7001002", "7001005", "7011000", "7500000"],
    "MX31": ["7201000"],
    "MX32": ["7201000"],
    "MX73": ["7000005", "7001002", "7201000"],
    "MX80": ["7201000"],
}


def ejecutar_consolidacion_por_sociedad(
    ruta_input,
    ruta_output,
    sociedad,
    sin_proveedores=False,
    sin_clientes=False,
    callback_status=None,
    cuentas_proveedores=None,
    cuentas_clientes=None,
):
    """
    Ejecuta el proceso de consolidación para UNA sociedad específica.

    Args:
        ruta_input (str): Ruta de la carpeta de entrada (padre de Proveedores/ y Clientes/)
        ruta_output (str): Ruta de la carpeta de salida
        sociedad (str): Código de la sociedad (ej: 'MX73')
        sin_proveedores (bool): Si True, omite procesamiento de proveedores
        sin_clientes (bool): Si True, omite procesamiento de clientes
        callback_status (function, optional): Función callback para actualizar estado
        cuentas_proveedores (dict | None): Diccionario {sociedad: [cuentas]} para proveedores.
            - Si se pasa un dict  → modo Manual: filtra sólo esas cuentas.
            - Si se pasa None     → modo Automático: usa TODAS las cuentas del archivo FBL3N.
        cuentas_clientes (dict | None): Igual que cuentas_proveedores pero para clientes.

    Returns:
        str: Ruta del archivo consolidado generado
    """

    def update_status(message):
        if callback_status:
            callback_status(message)

    # Determinar si estamos en modo automático (None = no filtrar)
    modo_automatico_prov = cuentas_proveedores is None
    modo_automatico_cli  = cuentas_clientes    is None

    # Cuando es manual y no se pasó dict, usar defaults
    CUENTAS_PROVEEDORES_POR_SOCIEDAD = cuentas_proveedores if not modo_automatico_prov else {}
    CUENTAS_CLIENTES_POR_SOCIEDAD    = cuentas_clientes    if not modo_automatico_cli  else {}

    # =========================
    # === Rutas y archivos ===
    # =========================

    archivo_fbl1            = os.path.join(ruta_input, 'Proveedores', f'FBL1_Proveedores_{sociedad}.xlsx')
    archivo_zfiq02          = os.path.join(ruta_input, 'Proveedores', f'ZFIQ02_Proveedores_{sociedad}.xlsx')
    archivo_fbl3            = os.path.join(ruta_input, 'Proveedores', f'FBL3N_Proveedores_{sociedad}.xlsx')
    archivo_fbl5            = os.path.join(ruta_input, 'Clientes',    f'FBL5N_Clientes_{sociedad}.xlsx')
    archivo_fbl3_clientes   = os.path.join(ruta_input, 'Clientes',    f'FBL3N_Clientes_{sociedad}.xlsx')
    archivo_cat_clientes    = os.path.join(ruta_input, 'Clientes',    'Clientes_Catalogo.xls')
    archivo_sociedad_nombre = os.path.join(ruta_input, 'Sociedad_Nombre.xls')
    archivo_cuentas_desc    = os.path.join(ruta_input, 'Cuentas_Desc.xls')

    # Verificar archivos requeridos
    archivos_requeridos = [archivo_cat_clientes]
    if not sin_proveedores:
        archivos_requeridos += [archivo_fbl1, archivo_zfiq02, archivo_fbl3]
    if not sin_clientes:
        archivos_requeridos += [archivo_fbl5, archivo_fbl3_clientes]

    for archivo in archivos_requeridos:
        if not os.path.exists(archivo):
            raise FileNotFoundError(f"Archivo no encontrado: {archivo}")

    # Nombre de la sociedad (opcional)
    nombre_sociedad = ""
    if os.path.exists(archivo_sociedad_nombre):
        try:
            df_soc_nombres = pd.read_csv(archivo_sociedad_nombre, dtype=str, header=0, sep='\t', encoding='utf-16')
            match = df_soc_nombres[df_soc_nombres["Soc."] == sociedad.upper()]
            if not match.empty:
                nombre_sociedad = match.iloc[0]["Nombre de la empresa"]
        except Exception:
            nombre_sociedad = ""

    # Descripciones de cuentas (opcional)
    dict_cuentas_desc = {}
    if os.path.exists(archivo_cuentas_desc):
        try:
            df_cuentas = pd.read_csv(archivo_cuentas_desc, dtype=str, sep='\t', encoding='utf-16')
            df_cuentas["Cuenta"]      = df_cuentas["Cuenta"].astype(str).str.strip()
            df_cuentas["Descripción"] = df_cuentas["Descripción"].astype(str).str.strip()
            dict_cuentas_desc = dict(zip(df_cuentas["Cuenta"], df_cuentas["Descripción"]))
        except Exception:
            dict_cuentas_desc = {}

    # =========================
    # === PROVEEDORES
    # =========================

    if sin_proveedores:
        update_status(f"⚠️ Sin movimientos de proveedores - {sociedad}, generando matrices vacías...")
        fbl1n       = pd.DataFrame()
        cat         = pd.DataFrame()
        fbl3n       = pd.DataFrame()
        matriz_prov = _matriz_sin_movimientos("Proveedores")
    else:
        update_status(f"📄 Procesando FBL1N - {sociedad}...")

        fbl1n = pd.read_excel(archivo_fbl1, dtype=str)
        fbl1n = clean_all_str(fbl1n)

        for col in ["Nº documento", "Cuenta"]:
            if col in fbl1n.columns:
                fbl1n[col] = fbl1n[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
        if "Nº documento" in fbl1n.columns:
            cols = fbl1n.columns.tolist()
            cols = ["Nº documento"] + [c for c in cols if c != "Nº documento"]
            fbl1n = fbl1n[cols]

        update_status("📄 Procesando catálogo de proveedores...")

        cat = pd.read_excel(archivo_zfiq02, dtype=str)
        cat = clean_all_str(cat)
        cat_reducido = cat[["Acreedor", "Nombre 1"]].drop_duplicates()

        fbl1n = fbl1n.merge(
            cat_reducido, how="left",
            left_on="Cuenta", right_on="Acreedor", suffixes=("_x", "_y")
        )
        fbl1n = fbl1n.drop(columns=["Acreedor_y"], errors="ignore")
        fbl1n = fbl1n.rename(columns={"Acreedor_x": "Acreedor"})
        fbl1n = fbl1n.drop(columns=["Nombre 1_x"])
        fbl1n = fbl1n.rename(columns={"Nombre 1_y": "Nombre 1"})

        if "Nombre 1" in fbl1n.columns:
            columnas = fbl1n.columns.tolist()
            columnas.remove("Nombre 1")
            insert_pos = 4 if len(columnas) >= 4 else len(columnas)
            columnas.insert(insert_pos, "Nombre 1")
            fbl1n = fbl1n[columnas]

        update_status("📄 Procesando FBL3N Proveedores...")

        fbl3n = pd.read_excel(archivo_fbl3, dtype=str)
        fbl3n = clean_all_str(fbl3n)
        for col in ["Nº documento", "Cuenta"]:
            if col in fbl3n.columns:
                fbl3n[col] = fbl3n[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()

        for col in ["Concepto Intercompañias", "UUID Auditor"]:
            if col not in fbl3n.columns:
                fbl3n[col] = ""

        # ── Filtro de cuentas ──────────────────────────────────────────────
        # Modo Manual: filtra por las cuentas definidas para la sociedad.
        # Modo Automático: usa todas las cuentas presentes en el archivo FBL3N.
        if not modo_automatico_prov and "Cuenta" in fbl3n.columns:
            cuentas_filtro = CUENTAS_PROVEEDORES_POR_SOCIEDAD.get(sociedad.upper(), ["6600022"])
            fbl3n = fbl3n[fbl3n["Cuenta"].isin(cuentas_filtro)]
        # En modo automático no se aplica ningún filtro → se conservan todas las cuentas.

        col_texto = "Texto" if "Texto" in fbl1n.columns else (
            fbl1n.columns[19] if len(fbl1n.columns) > 19 else fbl1n.columns[-1]
        )
        col_texto_cab = "Texto cab.documento" if "Texto cab.documento" in fbl1n.columns else (
            fbl1n.columns[20] if len(fbl1n.columns) > 20 else fbl1n.columns[-1]
        )

        if "Nº documento" in fbl1n.columns and "Nombre 1" in fbl1n.columns:
            fbl3n = fbl3n.merge(
                fbl1n[["Nº documento", "Nombre 1"]].rename(columns={"Nombre 1": "Proveedor"}),
                how="left", on="Nº documento"
            )

        if "Nº documento" in fbl1n.columns and col_texto in fbl1n.columns:
            fbl3n = fbl3n.merge(
                fbl1n[["Nº documento", col_texto]].rename(columns={col_texto: "Texto"}),
                how="left", on="Nº documento"
            )
        else:
            fbl3n["Texto"] = ""

        fbl3n = fbl3n.drop(columns=["Texto_x"], errors="ignore")
        fbl3n = fbl3n.rename(columns={"Texto_y": "Texto"})

        if "Nº documento" in fbl1n.columns and col_texto_cab in fbl1n.columns:
            fbl3n = fbl3n.merge(
                fbl1n[["Nº documento", col_texto_cab]].rename(columns={col_texto_cab: "Texto cab.documento"}),
                how="left", on="Nº documento"
            )
        else:
            fbl3n["Texto cab.documento"] = ""

        fbl3n = fbl3n.rename(columns={"Texto cab.documento_y": "Texto cab.documento"})

        orden_deseado = ["Proveedor", "Texto", "Texto cab.documento", "Concepto Intercompañias", "UUID Auditor"]
        presentes = [c for c in orden_deseado if c in fbl3n.columns]
        resto = [c for c in fbl3n.columns if c not in presentes]
        fbl3n = fbl3n[resto + presentes]

        # === MATRIZ PROVEEDORES ===
        update_status("📊 Creando matriz de proveedores...")

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

        totales_prov     = pivot_prov.sum(axis=0).to_frame().T
        totales_prov.index = ["Totales"]
        cuadre_prov      = pd.DataFrame([[0] * len(cuentas_prov)], index=["Cuadre Balanza"], columns=cuentas_prov)
        variaciones_prov = pd.DataFrame([[0] * len(cuentas_prov)], index=["variaciones"],    columns=cuentas_prov)
        matriz_prov = pd.concat([pivot_prov, totales_prov, cuadre_prov, variaciones_prov], axis=0)
        matriz_prov.insert(0, "Proveedores", matriz_prov.index)
        matriz_prov = matriz_prov.reset_index(drop=True)

        fila_cero_prov = pd.DataFrame({col: [""] for col in matriz_prov.columns})
        matriz_prov = pd.concat([fila_cero_prov, matriz_prov], ignore_index=True)

        col_sociedad_prov = [""] * len(matriz_prov)
        col_sociedad_prov[0] = sociedad.upper()
        col_sociedad_prov[1] = nombre_sociedad
        matriz_prov.insert(0, "Sociedad", col_sociedad_prov)

        for cuenta in cuentas_prov:
            if cuenta in matriz_prov.columns:
                matriz_prov.at[0, cuenta] = dict_cuentas_desc.get(str(cuenta).strip(), "")

        cols_numericas_prov = matriz_prov.columns.difference(["Sociedad", "Proveedores"])
        matriz_prov.loc[1:, cols_numericas_prov] = matriz_prov.loc[1:, cols_numericas_prov].apply(
            pd.to_numeric, errors='coerce'
        )

    # =========================
    # === CLIENTES
    # =========================

    if sin_clientes:
        update_status(f"⚠️ Sin movimientos de clientes - {sociedad}, generando matrices vacías...")
        fbl5n      = pd.DataFrame()
        clientes   = pd.DataFrame()
        fbl3n_cli  = pd.DataFrame()
        matriz_cli = _matriz_sin_movimientos("Clientes")
    else:
        update_status("📄 Procesando archivos de clientes...")

        fbl5n = pd.read_excel(archivo_fbl5, dtype=str)
        fbl5n = clean_all_str(fbl5n)
        for col in ["Nº documento", "Cuenta"]:
            if col in fbl5n.columns:
                fbl5n[col] = fbl5n[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()

        if "Nº documento" in fbl5n.columns:
            cols5 = fbl5n.columns.tolist()
            cols5 = ["Nº documento"] + [c for c in cols5 if c != "Nº documento"]
            fbl5n = fbl5n[cols5]

        clientes = pd.read_csv(archivo_cat_clientes, dtype=str, sep='\t', encoding='utf-16')
        clientes = clean_all_str(clientes)
        clientes["_cli_num"]    = clientes.iloc[:, 1].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()
        clientes["_cli_nombre"] = clientes.iloc[:, 3].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()

        cat_clientes_reducido = clientes[["_cli_num", "_cli_nombre"]].drop_duplicates()
        fbl5n = fbl5n.merge(cat_clientes_reducido, how="left",
                            left_on="Cuenta", right_on="_cli_num", suffixes=("", "_catcli"))
        fbl5n.drop(columns=["Nombre", "_cli_num"], inplace=True, errors="ignore")
        fbl5n.rename(columns={"_cli_nombre": "Nombre"}, inplace=True)

        if "Nombre" in fbl5n.columns:
            col_nombre = fbl5n.pop("Nombre")
            fbl5n.insert(3, "Nombre", col_nombre)

        update_status("📄 Procesando FBL3N Clientes...")

        fbl3n_cli = pd.read_excel(archivo_fbl3_clientes, dtype=str)
        fbl3n_cli = clean_all_str(fbl3n_cli)
        for col in ["Nº documento", "Cuenta"]:
            if col in fbl3n_cli.columns:
                fbl3n_cli[col] = fbl3n_cli[col].astype(str).str.replace(r"[\t\r\n]", "", regex=True).str.strip()

        for col in ["Concepto Intercompañias", "UUID Auditor"]:
            fbl3n_cli[col] = ""

        # ── Filtro de cuentas clientes ─────────────────────────────────────
        # Modo Manual: filtra por las cuentas definidas para la sociedad.
        # Modo Automático: usa todas las cuentas presentes en el archivo FBL3N Clientes.
        if not modo_automatico_cli and "Cuenta" in fbl3n_cli.columns:
            cuentas_clientes_filtrar = CUENTAS_CLIENTES_POR_SOCIEDAD.get(sociedad.upper(), ["7201000"])
            fbl3n_cli = fbl3n_cli[fbl3n_cli["Cuenta"].isin(cuentas_clientes_filtrar)]
        # En modo automático no se aplica ningún filtro → se conservan todas las cuentas.

        fbl5n["Nº documento"] = fbl5n["Nº documento"].astype(str).str.strip()
        fbl5n = fbl5n[
            fbl5n["Nº documento"].notna() &
            (fbl5n["Nº documento"] != "") &
            (fbl5n["Nº documento"].str.lower() != "nan")
        ].copy()

        col_texto2     = "Texto"
        col_texto_cab2 = "Asignación"

        fbl3n_cli = fbl3n_cli.merge(
            fbl5n[["Nº documento", "Nombre"]].rename(columns={"Nombre": "Cliente"}),
            how="left", on="Nº documento"
        )

        if col_texto2 in fbl5n.columns:
            fbl3n_cli = fbl3n_cli.merge(
                fbl5n[["Nº documento", col_texto2]].rename(columns={col_texto2: "Texto"}),
                how="left", on="Nº documento"
            )
        else:
            fbl3n_cli["Texto"] = ""

        if col_texto_cab2 in fbl5n.columns:
            fbl3n_cli = fbl3n_cli.merge(
                fbl5n[["Nº documento", col_texto_cab2]].rename(columns={col_texto_cab2: "Texto cab.documento"}),
                how="left", on="Nº documento"
            )
        else:
            fbl3n_cli["Texto cab.documento"] = ""

        fbl3n_cli = fbl3n_cli[
            fbl3n_cli["Nº documento"].notna() &
            (fbl3n_cli["Nº documento"] != "") &
            (fbl3n_cli["Nº documento"].str.lower() != "nan")
        ]

        orden_deseado2 = ["Cliente", "Texto", "Texto cab.documento", "Concepto Intercompañias", "UUID Auditor"]
        presentes2 = [c for c in orden_deseado2 if c in fbl3n_cli.columns]
        resto2 = [c for c in fbl3n_cli.columns if c not in presentes2]
        fbl3n_cli = fbl3n_cli[resto2 + presentes2]

        # === MATRIZ CLIENTES ===
        update_status("📊 Creando matriz de clientes...")

        if "Importe en moneda local" in fbl3n_cli.columns:
            fbl3n_cli["Importe en moneda local"] = (
                fbl3n_cli["Importe en moneda local"].astype(str).str.replace(",", "", regex=False)
            )
            fbl3n_cli["Importe en moneda local"] = pd.to_numeric(
                fbl3n_cli["Importe en moneda local"], errors="coerce"
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
        totales_cli    = pivot_cli.sum(axis=0).to_frame().T
        totales_cli.index = ["Totales"]
        cuadre_cli      = pd.DataFrame([[0] * len(cuentas_cli)], index=["Cuadre Balanza"], columns=cuentas_cli)
        variaciones_cli = pd.DataFrame([[0] * len(cuentas_cli)], index=["variaciones"],    columns=cuentas_cli)
        matriz_cli = pd.concat([pivot_cli, totales_cli, cuadre_cli, variaciones_cli], axis=0)
        matriz_cli.insert(0, "Clientes", matriz_cli.index)
        matriz_cli = matriz_cli.reset_index(drop=True)

        fila_cero_cli = pd.DataFrame({col: [""] for col in matriz_cli.columns})
        matriz_cli = pd.concat([fila_cero_cli, matriz_cli], ignore_index=True)

        col_sociedad_cli = [""] * len(matriz_cli)
        col_sociedad_cli[0] = sociedad.upper()
        col_sociedad_cli[1] = nombre_sociedad
        matriz_cli.insert(0, "Sociedad", col_sociedad_cli)

        for cuenta in cuentas_cli:
            if cuenta in matriz_cli.columns:
                matriz_cli.at[0, cuenta] = dict_cuentas_desc.get(str(cuenta).strip(), "")

        cols_numericas = matriz_cli.columns.difference(["Sociedad", "Clientes"])
        matriz_cli.loc[1:, cols_numericas] = matriz_cli.loc[1:, cols_numericas].apply(
            pd.to_numeric, errors='coerce'
        )

    # =========================
    # === Exportar Excel consolidado ===
    # =========================

    update_status("💾 Guardando archivo consolidado...")

    os.makedirs(ruta_output, exist_ok=True)
    archivo_consolidado = os.path.join(ruta_output, f"Intercompanias_Consolidado_{sociedad}.xlsx")

    with pd.ExcelWriter(archivo_consolidado, engine="openpyxl") as writer:
        fbl1n.to_excel(writer, sheet_name="FBL1N", index=False)
        cat.to_excel(writer, sheet_name="Cat Proveedores", index=False)
        fbl3n.to_excel(writer, sheet_name="FBL3N Proveedores", index=False)
        matriz_prov.to_excel(writer, sheet_name="Matriz Proveedores", index=False)
        fbl5n.to_excel(writer, sheet_name="FBL5N", index=False)
        clientes.drop(columns=["_cli_num", "_cli_nombre"], errors="ignore").to_excel(
            writer, sheet_name="Cat Clientes", index=False
        )
        fbl3n_cli.to_excel(writer, sheet_name="FBL3N Clientes", index=False)
        matriz_cli.to_excel(writer, sheet_name="Matriz Clientes", index=False)

    update_status(f"✅ Consolidación completada - {sociedad}")

    return archivo_consolidado


# =========================
# === Ejecución directa ===
# =========================
if __name__ == "__main__":
    user_profile = os.environ.get('USERPROFILE') or os.path.expanduser('~')
    base_path = os.path.join(user_profile, 'Documents', 'Intercompañias', 'RDA_Intercompanias', 'src')

    ruta_input  = os.path.join(base_path, 'Input')
    ruta_output = os.path.join(base_path, 'Output')

    try:
        archivo_generado = ejecutar_consolidacion_por_sociedad(
            ruta_input,
            ruta_output,
            sociedad="MX73",
            callback_status=lambda msg: print(msg)
            # cuentas_proveedores y cuentas_clientes omitidos → modo Automático
        )
        print(f"\n🎉 PROCESO COMPLETO FINALIZADO EXITOSAMENTE 🎉")
        print(f"Archivo guardado en: {archivo_generado}")

    except Exception as e:
        print(f"❌ Error durante la consolidación: {str(e)}")
        raise
