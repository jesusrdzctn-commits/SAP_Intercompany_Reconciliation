import os
import time
import json
import pandas as pd
import win32com.client as win32
from tkinter import messagebox
from DescargaSAP import FBL1N_Intercompañias, ZFIQ02_Intercompañias, FBL3N, FBL5_Intercompañias
from Consolidacion_V2 import ejecutar_consolidacion_por_sociedad


class IntercompaniasController:
    """Controller class that handles the business logic for the Intercompañías application"""

    def __init__(self, gui):
        """
        Initialize controller with GUI reference

        Args:
            gui: IntercompaniasGUI instance
        """
        self.gui = gui

        # Bind controller methods to GUI buttons
        self.gui.on_download = self.execute_download
        self.gui.on_consolidation = self.execute_consolidation

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
        """
        Internal method to execute the download process

        Args:
            config: Configuration dictionary from GUI
        """
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

            # Step 1: FBL1N
            self.gui.set_status(f"📥 Descargando FBL1 - {sociedad}...")
            FBL1_FileName = f"FBL1_Proveedores_{sociedad}.xlsx"
            fbl1n_con_datos = FBL1N_Intercompañias(
                [sociedad], DateFrom, DateTo, FolderPath, FBL1_FileName, fbl1n_from, fbl1n_to
            )
            FBL1_Intercompañias_File = os.path.join(FolderPath, FBL1_FileName)

            time.sleep(5)

            try:
                excel = win32.GetObject(Class="Excel.Application")
                for wb in list(excel.Workbooks):
                    try:
                        if os.path.abspath(wb.FullName) == FBL1_Intercompañias_File:
                            wb.Close(SaveChanges=False)
                    except Exception:
                        pass
            except Exception:
                pass

            # Step 2: ZFIQ02
            self.gui.set_status(f"📥 Descargando ZFIQ02 - {sociedad}...")
            ZFIQ02_FileName = f"ZFIQ02_Proveedores_{sociedad}.xlsx"
            ZFIQ02_Intercompañias_File = os.path.join(FolderPath, ZFIQ02_FileName)
            ZFIQ02_Intercompañias([sociedad], ZFIQ02_Intercompañias_File)

            # Step 3: FBL3N Proveedores
            FBL3N_FileName = f"FBL3N_Proveedores_{sociedad}.xlsx"
            FBL3N_File = os.path.join(FolderPath, FBL3N_FileName)

            if fbl1n_con_datos:
                self.gui.set_status(f"📄 Procesando documentos - {sociedad}...")
                df_FBL1 = pd.read_excel(FBL1_Intercompañias_File, engine='openpyxl')
                colNDocument = df_FBL1.columns[6]
                resultado = df_FBL1[colNDocument].dropna().astype(str).unique().tolist()

                self.gui.set_status(f"📥 Descargando FBL3N - {sociedad}...")
                if os.path.exists(FBL3N_File):
                    os.remove(FBL3N_File)
                FBL3N(resultado, [sociedad], DateFrom, DateTo, FolderPath, FBL3N_FileName)
                time.sleep(3)
            else:
                self.gui.set_status(f"⚠️ FBL1N sin movimientos - {sociedad}, se omite FBL3N Proveedores")
                import openpyxl as _oxl
                wb_vacio = _oxl.Workbook()
                wb_vacio.active.title = "Sin datos"
                wb_vacio.save(FBL3N_File)

            self.gui.set_status(f"✅ Proveedores {sociedad} completados ({idx}/{len(sociedades)})")
            time.sleep(2)

            # Step 4: FBL5N
            self.gui.set_status(f"📥 Descargando FBL5N - {sociedad}...")
            FBL5_FileName = f"FBL5N_Clientes_{sociedad}.xlsx"
            FBL5_File = os.path.join(ClientesPath, FBL5_FileName)
            fbl5n_con_datos = FBL5_Intercompañias(
                [sociedad], DateFrom, DateTo, ClientesPath, FBL5_FileName, fbl5n_from, fbl5n_to
            )

            time.sleep(5)

            try:
                excel = win32.GetObject(Class="Excel.Application")
                for wb in list(excel.Workbooks):
                    try:
                        if os.path.abspath(wb.FullName) == os.path.abspath(FBL5_File):
                            wb.Close(SaveChanges=False)
                    except Exception:
                        pass
            except Exception:
                pass

            # Step 5: FBL3N Clientes
            self.gui.set_status(f"📥 Descargando FBL3N Clientes - {sociedad}...")
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
                time.sleep(3)
            else:
                self.gui.set_status(f"⚠️ FBL5N sin movimientos - {sociedad}, se omite FBL3N Clientes")
                import openpyxl as _oxl
                wb_vacio = _oxl.Workbook()
                wb_vacio.active.title = "Sin datos"
                wb_vacio.save(FBL3N_Clientes_File)

            # Guardar flags por sociedad para que la consolidación los lea
            flags_sin_movimientos = {
                'sin_proveedores': not fbl1n_con_datos,
                'sin_clientes':    not fbl5n_con_datos,
            }
            flags_path = os.path.join(FolderPath, f"_flags_{sociedad}.json")
            with open(flags_path, 'w') as f:
                json.dump(flags_sin_movimientos, f)

            self.gui.set_status(f"✅ Sociedad {sociedad} completada ({idx}/{len(sociedades)})")
            time.sleep(2)

    def execute_consolidation(self):
        """Execute the consolidation process"""
        config = self.gui.get_config()

        confirm = messagebox.askyesno(
            "Confirmar Consolidación",
            "¿Desea ejecutar el proceso de consolidación?\n\n"
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
        """
        Internal method to execute the consolidation process

        Args:
            config: Configuration dictionary from GUI
        """
        # Las rutas vienen directamente del GUI (ya incluyen Proveedores / Clientes)
        ruta_input_prov = config['input_path']
        ruta_input      = os.path.dirname(ruta_input_prov)   # carpeta Input padre
        ruta_output     = config['output_path']

        for sociedad in config['sociedades']:
            # Leer flags dejados por la descarga (si existen)
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
                cuentas_proveedores=config['cuentas_proveedores_por_sociedad'],
                cuentas_clientes=config['cuentas_clientes_por_sociedad'],
            )

        num_archivos = len(config['sociedades'])
        self.gui.set_status(f"✅ {num_archivos} archivo(s) generado(s)")
