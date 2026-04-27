import os
import time
import json
import shutil
from datetime import datetime, timedelta
import pandas as pd
import win32com.client as win32
from tkinter import messagebox
from DescargaSAP import FBL1N_Intercompañias, ZFIQ02_Intercompañias, FBL3N, FBL5_Intercompañias
from Consolidacion_V2 import ejecutar_consolidacion_por_sociedad


class IntercompaniasController:
    """Controller class that handles the business logic for the Intercompañías application"""

    def __init__(self, gui):
        self.gui = gui
        self.gui.on_download = self.execute_download
        self.gui.on_consolidation = self.execute_consolidation
        self.gui.on_download_large = self.execute_download_large   # nuevo callback

    # =========================================================
    # DESCARGA NORMAL
    # =========================================================
    def execute_download(self):
        """Execute the document download process from SAP"""
        config = self.gui.get_config()

        if not config['sociedades']:
            messagebox.showerror("Error", "Debe agregar al menos una sociedad")
            return

        if not self.gui.validate_dates():
            return

        confirm = messagebox.askyesno(
            "Confirmar",
            f"¿Desea iniciar la descarga para {len(config['sociedades'])} sociedad(es)?\n"
            f"Periodo: {config['date_from']} - {config['date_to']}\n\n"
            f"Los archivos de proveedores se guardarán en:\n{config['input_path']}\n\n"
            f"Los archivos de clientes se guardarán en:\n{config['clientes_path']}"
        )
        if not confirm:
            return

        os.makedirs(config['input_path'], exist_ok=True)
        os.makedirs(config['clientes_path'], exist_ok=True)

        self.gui.disable_buttons()
        self.gui.set_status("⏳ Procesando... Por favor espere")

        try:
            self._run_download_process(config)
            self.gui.set_status("✅ ¡Proceso completado exitosamente!")
            messagebox.showinfo(
                "Éxito",
                "La descarga de documentos se completó correctamente.\n\n"
                f"Archivos de proveedores guardados en:\n{config['input_path']}\n\n"
                f"Archivos de clientes guardados en:\n{config['clientes_path']}"
            )
        except Exception as e:
            self.gui.set_status("❌ Error en el proceso")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso:\n\n{str(e)}")
        finally:
            self.gui.enable_buttons()
            if "completado exitosamente" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    def _run_download_process(self, config):
        """Proceso de descarga normal (sin chunking)."""
        try:
            excel = win32.GetObject(Class="Excel.Application")
            excel.Interactive = False
        except Exception:
            pass

        try:
            DateFrom     = config['date_from']
            DateTo       = config['date_to']
            FolderPath   = config['input_path']
            ClientesPath = config['clientes_path']
            sociedades   = config['sociedades']
            fbl1n_from   = config['fbl1n_range_from']
            fbl1n_to     = config['fbl1n_range_to']
            fbl5n_from   = config['fbl5n_range_from']
            fbl5n_to     = config['fbl5n_range_to']

            for idx, sociedad in enumerate(sociedades, 1):
                self.gui.set_status(f"📥 Procesando sociedad {sociedad} ({idx}/{len(sociedades)})...")
                self._descargar_sociedad(
                    sociedad, DateFrom, DateTo,
                    FolderPath, ClientesPath,
                    fbl1n_from, fbl1n_to, fbl5n_from, fbl5n_to,
                    idx, len(sociedades),
                )
        finally:
            try:
                excel = win32.GetObject(Class="Excel.Application")
                excel.Interactive = True
            except Exception:
                pass

    def _descargar_sociedad(
        self, sociedad, DateFrom, DateTo,
        FolderPath, ClientesPath,
        fbl1n_from, fbl1n_to, fbl5n_from, fbl5n_to,
        idx, total,
    ):
        """
        Descarga todos los archivos SAP para una sociedad en el periodo dado.
        Retorna (fbl1n_con_datos, fbl5n_con_datos).
        """
        # ── FBL1N ────────────────────────────────────────────────────
        self.gui.set_status(f"📥 Descargando FBL1N - {sociedad}...")
        FBL1_FileName = f"FBL1_Proveedores_{sociedad}.xlsx"
        fbl1n_con_datos = FBL1N_Intercompañias(
            [sociedad], DateFrom, DateTo, FolderPath, FBL1_FileName, fbl1n_from, fbl1n_to
        )
        FBL1_File = os.path.join(FolderPath, FBL1_FileName)

        time.sleep(10)
        self._cerrar_workbook_excel(FBL1_File)

        # ── ZFIQ02 ───────────────────────────────────────────────────
        self.gui.set_status(f"📥 Descargando ZFIQ02 - {sociedad}...")
        ZFIQ02_FileName = f"ZFIQ02_Proveedores_{sociedad}.xlsx"
        ZFIQ02_File = os.path.join(FolderPath, ZFIQ02_FileName)
        ZFIQ02_Intercompañias([sociedad], ZFIQ02_File)

        # ── FBL3N Proveedores ─────────────────────────────────────────
        FBL3N_FileName = f"FBL3N_Proveedores_{sociedad}.xlsx"
        FBL3N_File = os.path.join(FolderPath, FBL3N_FileName)

        if fbl1n_con_datos:
            self.gui.set_status(f"📄 Procesando documentos - {sociedad}...")
            df_FBL1 = pd.read_excel(FBL1_File, engine='openpyxl')
            colNDocument = df_FBL1.columns[6]
            resultado = df_FBL1[colNDocument].dropna().astype(str).unique().tolist()

            self.gui.set_status(f"📥 Descargando FBL3N - {sociedad}...")
            if os.path.exists(FBL3N_File):
                os.remove(FBL3N_File)
            FBL3N(resultado, [sociedad], DateFrom, DateTo, FolderPath, FBL3N_FileName)
            time.sleep(8)
            self._cerrar_workbook_excel(FBL3N_File)
        else:
            self.gui.set_status(f"⚠️ FBL1N sin movimientos - {sociedad}, se omite FBL3N Proveedores")
            self._crear_excel_vacio(FBL3N_File)

        self.gui.set_status(f"✅ Proveedores {sociedad} completados ({idx}/{total})")
        time.sleep(5)

        # ── FBL5N ────────────────────────────────────────────────────
        self.gui.set_status(f"📥 Descargando FBL5N - {sociedad}...")
        FBL5_FileName = f"FBL5N_Clientes_{sociedad}.xlsx"
        FBL5_File = os.path.join(ClientesPath, FBL5_FileName)
        fbl5n_con_datos = FBL5_Intercompañias(
            [sociedad], DateFrom, DateTo, ClientesPath, FBL5_FileName, fbl5n_from, fbl5n_to
        )

        time.sleep(10)
        self._cerrar_workbook_excel(FBL5_File)

        # ── FBL3N Clientes ────────────────────────────────────────────
        FBL3N_Clientes_FileName = f"FBL3N_Clientes_{sociedad}.xlsx"
        FBL3N_Clientes_File = os.path.join(ClientesPath, FBL3N_Clientes_FileName)

        if fbl5n_con_datos:
            self.gui.set_status(f"📄 Procesando documentos clientes - {sociedad}...")
            df_FBL5 = pd.read_excel(FBL5_File, engine='openpyxl')
            colNDocument_cli = df_FBL5.columns[8]
            resultado_cli = df_FBL5[colNDocument_cli].dropna().astype(str).unique().tolist()

            self.gui.set_status(f"📥 Descargando FBL3N Clientes - {sociedad}...")
            if os.path.exists(FBL3N_Clientes_File):
                os.remove(FBL3N_Clientes_File)
            FBL3N(resultado_cli, [sociedad], DateFrom, DateTo, ClientesPath, FBL3N_Clientes_FileName)
            time.sleep(8)
            self._cerrar_workbook_excel(FBL3N_Clientes_File)
        else:
            self.gui.set_status(f"⚠️ FBL5N sin movimientos - {sociedad}, se omite FBL3N Clientes")
            self._crear_excel_vacio(FBL3N_Clientes_File)

        # ── Guardar flags ─────────────────────────────────────────────
        flags = {
            'sin_proveedores': not fbl1n_con_datos,
            'sin_clientes':    not fbl5n_con_datos,
        }
        flags_path = os.path.join(FolderPath, f"_flags_{sociedad}.json")
        with open(flags_path, 'w') as f:
            json.dump(flags, f)

        self.gui.set_status(f"✅ Sociedad {sociedad} completada ({idx}/{total})")
        time.sleep(5)

        return fbl1n_con_datos, fbl5n_con_datos

    # =========================================================
    # DESCARGA POR CHUNKS (SOCIEDADES GRANDES)
    # =========================================================
    def execute_download_large(self):
        """Execute the chunked download process for large companies."""
        config = self.gui.get_config()

        if not config['sociedades']:
            messagebox.showerror("Error", "Debe agregar al menos una sociedad")
            return

        if not self.gui.validate_dates():
            return

        chunk_days = self.gui.validate_chunk_days()
        if chunk_days is None:
            return

        confirm = messagebox.askyesno(
            "Confirmar Descarga por Chunks",
            f"¿Desea iniciar la descarga por chunks para {len(config['sociedades'])} sociedad(es)?\n\n"
            f"Periodo: {config['date_from']} - {config['date_to']}\n"
            f"Días por chunk: {chunk_days}\n\n"
            f"Los archivos temporales se guardarán en subcarpetas por sociedad dentro de Input.\n"
            f"Al finalizar, se consolidarán en un solo archivo en Input/Proveedores e Input/Clientes."
        )
        if not confirm:
            return

        self.gui.disable_buttons()
        self.gui.set_status("⏳ Procesando chunks... Por favor espere")

        try:
            self._run_chunked_download(config, chunk_days)
            self.gui.set_status("✅ ¡Descarga por chunks completada!")
            messagebox.showinfo(
                "Éxito",
                "La descarga por chunks se completó correctamente.\n\n"
                f"Archivos consolidados en:\n"
                f"  • {config['input_path']}\n"
                f"  • {config['clientes_path']}"
            )
        except Exception as e:
            self.gui.set_status("❌ Error en descarga por chunks")
            messagebox.showerror("Error", f"Ocurrió un error durante la descarga:\n\n{str(e)}")
        finally:
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    def _run_chunked_download(self, config, chunk_days):
        """Lógica principal del chunking: divide el periodo, descarga cada bloque y apila."""
        try:
            excel = win32.GetObject(Class="Excel.Application")
            excel.Interactive = False
        except Exception:
            pass

        try:
            date_from    = datetime.strptime(config['date_from'], "%d.%m.%Y")
            date_to      = datetime.strptime(config['date_to'],   "%d.%m.%Y")
            FolderPath   = config['input_path']
            ClientesPath = config['clientes_path']
            fbl1n_from   = config['fbl1n_range_from']
            fbl1n_to     = config['fbl1n_range_to']
            fbl5n_from   = config['fbl5n_range_from']
            fbl5n_to     = config['fbl5n_range_to']

            # Generar lista de chunks
            chunks = self._generar_chunks(date_from, date_to, chunk_days)

            for idx_soc, sociedad in enumerate(config['sociedades'], 1):
                self.gui.set_status(
                    f"🏢 Procesando sociedad grande {sociedad} "
                    f"({idx_soc}/{len(config['sociedades'])}) — {len(chunks)} chunks..."
                )

                # Carpetas temporales por sociedad
                tmp_prov = os.path.join(FolderPath,   f"_chunks_{sociedad}")
                tmp_cli  = os.path.join(ClientesPath, f"_chunks_{sociedad}")
                os.makedirs(tmp_prov, exist_ok=True)
                os.makedirs(tmp_cli,  exist_ok=True)

                chunks_fbl1n     = []   # rutas de archivos FBL1N por chunk con datos
                chunks_zfiq02    = []   # ZFIQ02 (solo necesitamos uno, pero descargamos por chunk)
                chunks_fbl3n_p   = []   # FBL3N Proveedores por chunk con datos
                chunks_fbl5n     = []   # FBL5N por chunk con datos
                chunks_fbl3n_c   = []   # FBL3N Clientes por chunk con datos

                for idx_chunk, (chunk_from, chunk_to) in enumerate(chunks, 1):
                    chunk_from_str = chunk_from.strftime("%d.%m.%Y")
                    chunk_to_str   = chunk_to.strftime("%d.%m.%Y")

                    self.gui.set_status(
                        f"📦 [{sociedad}] Chunk {idx_chunk}/{len(chunks)}: "
                        f"{chunk_from_str} → {chunk_to_str}"
                    )

                    # ── FBL1N chunk ───────────────────────────────────
                    fbl1n_fname = f"FBL1_Proveedores_{sociedad}_chunk{idx_chunk:03d}.xlsx"
                    fbl1n_path  = os.path.join(tmp_prov, fbl1n_fname)
                    fbl1n_ok = FBL1N_Intercompañias(
                        [sociedad], chunk_from_str, chunk_to_str,
                        tmp_prov, fbl1n_fname, fbl1n_from, fbl1n_to
                    )
                    time.sleep(8)
                    self._cerrar_workbook_excel(fbl1n_path)

                    if fbl1n_ok:
                        chunks_fbl1n.append(fbl1n_path)

                    # ── ZFIQ02 (una sola vez, primer chunk con éxito) ──
                    if not chunks_zfiq02:
                        zfiq02_fname = f"ZFIQ02_Proveedores_{sociedad}.xlsx"
                        zfiq02_path  = os.path.join(tmp_prov, zfiq02_fname)
                        ZFIQ02_Intercompañias([sociedad], zfiq02_path)
                        if os.path.exists(zfiq02_path):
                            chunks_zfiq02.append(zfiq02_path)

                    # ── FBL3N Proveedores chunk ───────────────────────
                    if fbl1n_ok:
                        df_fbl1 = pd.read_excel(fbl1n_path, engine='openpyxl')
                        col_ndoc = df_fbl1.columns[6]
                        resultado = df_fbl1[col_ndoc].dropna().astype(str).unique().tolist()

                        fbl3n_p_fname = f"FBL3N_Proveedores_{sociedad}_chunk{idx_chunk:03d}.xlsx"
                        fbl3n_p_path  = os.path.join(tmp_prov, fbl3n_p_fname)
                        if os.path.exists(fbl3n_p_path):
                            os.remove(fbl3n_p_path)

                        fbl3n_p_ok = FBL3N(
                            resultado, [sociedad],
                            chunk_from_str, chunk_to_str,
                            tmp_prov, fbl3n_p_fname
                        )
                        time.sleep(8)
                        self._cerrar_workbook_excel(fbl3n_p_path)

                        if fbl3n_p_ok:
                            chunks_fbl3n_p.append(fbl3n_p_path)

                    # ── FBL5N chunk ───────────────────────────────────
                    fbl5n_fname = f"FBL5N_Clientes_{sociedad}_chunk{idx_chunk:03d}.xlsx"
                    fbl5n_path  = os.path.join(tmp_cli, fbl5n_fname)
                    fbl5n_ok = FBL5_Intercompañias(
                        [sociedad], chunk_from_str, chunk_to_str,
                        tmp_cli, fbl5n_fname, fbl5n_from, fbl5n_to
                    )
                    time.sleep(8)
                    self._cerrar_workbook_excel(fbl5n_path)

                    if fbl5n_ok:
                        chunks_fbl5n.append(fbl5n_path)

                    # ── FBL3N Clientes chunk ──────────────────────────
                    if fbl5n_ok:
                        df_fbl5 = pd.read_excel(fbl5n_path, engine='openpyxl')
                        col_ndoc_cli = df_fbl5.columns[8]
                        resultado_cli = df_fbl5[col_ndoc_cli].dropna().astype(str).unique().tolist()

                        fbl3n_c_fname = f"FBL3N_Clientes_{sociedad}_chunk{idx_chunk:03d}.xlsx"
                        fbl3n_c_path  = os.path.join(tmp_cli, fbl3n_c_fname)
                        if os.path.exists(fbl3n_c_path):
                            os.remove(fbl3n_c_path)

                        fbl3n_c_ok = FBL3N(
                            resultado_cli, [sociedad],
                            chunk_from_str, chunk_to_str,
                            tmp_cli, fbl3n_c_fname
                        )
                        time.sleep(8)
                        self._cerrar_workbook_excel(fbl3n_c_path)

                        if fbl3n_c_ok:
                            chunks_fbl3n_c.append(fbl3n_c_path)

                    time.sleep(3)

                # ── Apilar chunks en archivos finales ─────────────────
                self.gui.set_status(f"📋 [{sociedad}] Apilando chunks en archivos finales...")

                fbl1n_final   = os.path.join(FolderPath,   f"FBL1_Proveedores_{sociedad}.xlsx")
                zfiq02_final  = os.path.join(FolderPath,   f"ZFIQ02_Proveedores_{sociedad}.xlsx")
                fbl3n_p_final = os.path.join(FolderPath,   f"FBL3N_Proveedores_{sociedad}.xlsx")
                fbl5n_final   = os.path.join(ClientesPath, f"FBL5N_Clientes_{sociedad}.xlsx")
                fbl3n_c_final = os.path.join(ClientesPath, f"FBL3N_Clientes_{sociedad}.xlsx")

                self._apilar_chunks(chunks_fbl1n,   fbl1n_final,   sociedad, "FBL1N")
                self._apilar_chunks(chunks_fbl3n_p, fbl3n_p_final, sociedad, "FBL3N Proveedores")
                self._apilar_chunks(chunks_fbl5n,   fbl5n_final,   sociedad, "FBL5N")
                self._apilar_chunks(chunks_fbl3n_c, fbl3n_c_final, sociedad, "FBL3N Clientes")

                # ZFIQ02: copiar el primero que se descargó (no necesita apilarse)
                if chunks_zfiq02:
                    shutil.copy2(chunks_zfiq02[0], zfiq02_final)
                else:
                    self._crear_excel_vacio(zfiq02_final)

                # ── Guardar flags ─────────────────────────────────────
                flags = {
                    'sin_proveedores': len(chunks_fbl1n) == 0,
                    'sin_clientes':    len(chunks_fbl5n) == 0,
                }
                flags_path = os.path.join(FolderPath, f"_flags_{sociedad}.json")
                with open(flags_path, 'w') as f:
                    json.dump(flags, f)

                # ── Limpiar carpetas temporales ───────────────────────
                self.gui.set_status(f"🧹 [{sociedad}] Limpiando archivos temporales...")
                shutil.rmtree(tmp_prov, ignore_errors=True)
                shutil.rmtree(tmp_cli,  ignore_errors=True)

                self.gui.set_status(f"✅ [{sociedad}] ¡Completado! ({idx_soc}/{len(config['sociedades'])})")
                time.sleep(3)

        finally:
            try:
                excel = win32.GetObject(Class="Excel.Application")
                excel.Interactive = True
            except Exception:
                pass

    @staticmethod
    def _generar_chunks(date_from, date_to, chunk_days):
        """
        Divide el rango [date_from, date_to] en bloques de `chunk_days` días.
        El último bloque puede ser menor si el rango no es divisible exactamente.

        Returns:
            list[tuple[datetime, datetime]]: Lista de (inicio, fin) de cada chunk.
        """
        chunks = []
        current = date_from
        while current <= date_to:
            end = min(current + timedelta(days=chunk_days - 1), date_to)
            chunks.append((current, end))
            current = end + timedelta(days=1)
        return chunks

    @staticmethod
    def _apilar_chunks(rutas_chunks, ruta_final, sociedad, nombre_reporte):
        """
        Lee todos los archivos de chunk, los concatena en un solo DataFrame
        y lo guarda en ruta_final. Si no hay chunks con datos, crea un archivo vacío.

        Args:
            rutas_chunks (list[str]): Rutas de archivos Excel con datos.
            ruta_final (str): Ruta del archivo Excel de salida.
            sociedad (str): Código de sociedad (para logs).
            nombre_reporte (str): Nombre del reporte (para logs).
        """
        if not rutas_chunks:
            # Sin datos en ningún chunk → archivo vacío placeholder
            import openpyxl as _oxl
            wb = _oxl.Workbook()
            wb.active.title = "Sin datos"
            wb.save(ruta_final)
            print(f"[CHUNKS] {nombre_reporte} {sociedad}: sin datos en ningún chunk → archivo vacío creado.")
            return

        dfs = []
        for ruta in rutas_chunks:
            try:
                df = pd.read_excel(ruta, dtype=str, engine='openpyxl')
                # Ignorar hojas/archivos marcados como "Sin datos"
                if df.empty or (len(df.columns) == 1 and "Sin datos" in df.columns[0]):
                    continue
                dfs.append(df)
            except Exception as e:
                print(f"[CHUNKS] Advertencia al leer chunk {ruta}: {e}")

        if not dfs:
            import openpyxl as _oxl
            wb = _oxl.Workbook()
            wb.active.title = "Sin datos"
            wb.save(ruta_final)
            return

        df_total = pd.concat(dfs, ignore_index=True)
        df_total.to_excel(ruta_final, index=False, engine='openpyxl')
        print(f"[CHUNKS] {nombre_reporte} {sociedad}: {len(dfs)} chunks apilados → {ruta_final}")

    # =========================================================
    # CONSOLIDACIÓN
    # =========================================================
    def execute_consolidation(self):
        """Execute the consolidation process"""
        config = self.gui.get_config()

        mode = config.get("consolidation_mode", "manual")
        mode_label = "Manual" if mode == "manual" else "Automático"

        confirm = messagebox.askyesno(
            "Confirmar Consolidación",
            f"¿Desea ejecutar el proceso de consolidación?\n\n"
            f"Modo de cuentas: {mode_label}\n\n"
            f"Se leerán archivos de proveedores de:\n{config['input_path']}\n"
            f"Se leerán archivos de clientes de:\n{config['clientes_path']}\n\n"
            f"El archivo consolidado se guardará en:\n{config['output_path']}"
        )
        if not confirm:
            return

        os.makedirs(config['output_path'], exist_ok=True)

        self.gui.disable_buttons()
        self.gui.set_status("⏳ Consolidando... Por favor espere")

        try:
            self._run_consolidation_process(config)
            self.gui.set_status("✅ ¡Consolidación completada!")
            messagebox.showinfo(
                "Éxito",
                "El proceso de consolidación se completó correctamente.\n\n"
                f"Archivo guardado en:\n{config['output_path']}"
            )
        except FileNotFoundError as e:
            self.gui.set_status("❌ Archivos no encontrados")
            messagebox.showerror(
                "Error - Archivos Faltantes",
                "No se encontraron todos los archivos necesarios.\n\n"
                "Asegúrese de ejecutar primero 'Descarga de Documentos'\n"
                "y que todos los archivos estén en la carpeta de entrada.\n\n"
                f"Detalles: {str(e)}"
            )
        except Exception as e:
            self.gui.set_status("❌ Error en consolidación")
            messagebox.showerror("Error", f"Ocurrió un error durante la consolidación:\n\n{str(e)}")
        finally:
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    def _run_consolidation_process(self, config):
        ruta_input_prov = config['input_path']
        ruta_input      = os.path.dirname(ruta_input_prov)
        ruta_output     = config['output_path']

        mode = config.get("consolidation_mode", "manual")
        cuentas_proveedores = config['cuentas_proveedores_por_sociedad'] if mode == "manual" else None
        cuentas_clientes    = config['cuentas_clientes_por_sociedad']    if mode == "manual" else None

        for sociedad in config['sociedades']:
            flags_path = os.path.join(ruta_input_prov, f"_flags_{sociedad}.json")
            flags = {'sin_proveedores': False, 'sin_clientes': False}
            if os.path.exists(flags_path):
                try:
                    with open(flags_path) as f:
                        flags = json.load(f)
                except Exception:
                    pass

            ejecutar_consolidacion_por_sociedad(
                ruta_input,
                ruta_output,
                sociedad=sociedad,
                sin_proveedores=flags['sin_proveedores'],
                sin_clientes=flags['sin_clientes'],
                callback_status=self.gui.set_status,
                cuentas_proveedores=cuentas_proveedores,
                cuentas_clientes=cuentas_clientes,
            )

        num_archivos = len(config['sociedades'])
        self.gui.set_status(f"✅ {num_archivos} archivo(s) generado(s)")

    # =========================================================
    # HELPERS
    # =========================================================
    @staticmethod
    def _cerrar_workbook_excel(ruta_archivo):
        """Cierra un workbook de Excel si está abierto."""
        try:
            excel = win32.GetObject(Class="Excel.Application")
            for wb in list(excel.Workbooks):
                try:
                    if os.path.abspath(wb.FullName) == os.path.abspath(ruta_archivo):
                        wb.Close(SaveChanges=False)
                except Exception:
                    pass
        except Exception:
            pass

    @staticmethod
    def _crear_excel_vacio(ruta_archivo):
        """Crea un archivo Excel vacío como placeholder."""
        import openpyxl as _oxl
        wb = _oxl.Workbook()
        wb.active.title = "Sin datos"
        wb.save(ruta_archivo)
