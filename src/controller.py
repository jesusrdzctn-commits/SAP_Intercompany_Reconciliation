from logging import config
import os
import time
import pandas as pd
import win32com.client as win32
from tkinter import messagebox
from FBL1_Intercompañias import FBL1N_Intercompañias, ZFIQ02_Intercompañias, FBL3N, FBL5_Intercompañias
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
        # Get configuration from GUI
        config = self.gui.get_config()
        
        # Validate inputs
        if not config['sociedades']:
            messagebox.showerror("Error", "Debe agregar al menos una sociedad")
            return
        
        if not self.gui.validate_dates():
            return
        
        # Confirm action
        confirm = messagebox.askyesno(
            "Confirmar", 
            f"¿Desea iniciar la descarga para {len(config['sociedades'])} sociedad(es)?\n"
            f"Periodo: {config['date_from']} - {config['date_to']}\n\n"
            f"Los archivos se guardarán en:\n{config['input_path']}"
        )
        
        if not confirm:
            return
        
        # Create output directory if it doesn't exist
        os.makedirs(config['input_path'], exist_ok=True)
        os.makedirs(config['clientes_path'], exist_ok=True)
        
        # Disable buttons during execution
        self.gui.disable_buttons()
        self.gui.set_status("⏳ Procesando... Por favor espere")
        
        try:
            # Execute the download process
            self._run_download_process(config)
            
            # Success message
            self.gui.set_status("✅ ¡Proceso completado exitosamente!")
            messagebox.showinfo(
                "Éxito", 
                "La descarga de documentos se completó correctamente.\n\n"
                f"Archivos guardados en:\n{config['input_path']}"
            )
            
        except Exception as e:
            self.gui.set_status("❌ Error en el proceso")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso:\n\n{str(e)}")
        
        finally:
            # Re-enable buttons
            self.gui.enable_buttons()
            if "completado exitosamente" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")
    
    def _run_download_process(self, config):
        """
        Internal method to execute the download process
        
        Args:
            config: Configuration dictionary from GUI
        """
        DateFrom = config['date_from']
        DateTo = config['date_to']
        FolderPath = config['input_path']
        ClientesPath = config['clientes_path']
        sociedades = config['sociedades']
        
        # Procesar cada sociedad por separado
        for idx, sociedad in enumerate(sociedades, 1):
            self.gui.set_status(f"📥 Procesando sociedad {sociedad} ({idx}/{len(sociedades)})...")
            
            # Step 1: FBL1 Download por sociedad
            self.gui.set_status(f"📥 Descargando FBL1 - {sociedad}...")
            FileName = f"FBL1_Proveedores_{sociedad}.xlsx"
            FBL1N_Intercompañias([sociedad], DateFrom, DateTo, FolderPath, FileName)
            FBL1_Intercompañias_File = os.path.join(FolderPath, FileName)
            
            time.sleep(5)  # Wait for file to be saved
            
            # Close Excel workbook if open
            try:
                excel = win32.GetObject(Class="Excel.Application")
                for wb in list(excel.Workbooks):
                    try:
                        if os.path.abspath(wb.FullName) == FBL1_Intercompañias_File:
                            wb.Close(SaveChanges=False)
                    except Exception:
                        pass
            except:
                pass  # Excel might not be running
            
            # Step 2: ZFIQ02 Download por sociedad
            self.gui.set_status(f"📥 Descargando ZFIQ02 - {sociedad}...")
            ZFIQ02_FileName = f"ZFIQ02_Proveedores_{sociedad}.xlsx"
            ZFIQ02_Intercompañias_File = os.path.join(FolderPath, ZFIQ02_FileName)
            ZFIQ02_Intercompañias([sociedad], ZFIQ02_Intercompañias_File)
            
            # Step 3: Read FBL1 and extract document numbers
            self.gui.set_status(f"📄 Procesando documentos - {sociedad}...")
            df_FBL1 = pd.read_excel(FBL1_Intercompañias_File, engine='openpyxl')
            colNDocument = df_FBL1.columns[6]
            resultado = df_FBL1[colNDocument].dropna().astype(str).unique().tolist()
            
            # Step 4: FBL3N Download por sociedad
            self.gui.set_status(f"📥 Descargando FBL3N - {sociedad}...")
            FBL3N(resultado, [sociedad], DateFrom, DateTo,FolderPath, FileName)
            
            # Step 5: Save FBL3N with proper name
            self.gui.set_status(f"💾 Guardando FBL3N - {sociedad}...")
            time.sleep(3)
            
            try:
                excel = win32.GetObject(Class="Excel.Application")
                FBL3N_FileName = f"FBL3N_Proveedores_{sociedad}.xlsx"
                FBL3N_File = os.path.join(FolderPath, FBL3N_FileName)
                
                if os.path.exists(FBL3N_File):
                    os.remove(FBL3N_File)
                
                for wb in list(excel.Workbooks):
                    try:
                        wb.SaveAs(FBL3N_File)
                        wb.Close(SaveChanges=False)
                        break
                    except Exception:
                        pass
            except Exception as e:
                self.gui.set_status(f"⚠️ Advertencia: Error guardando FBL3N para {sociedad}")
            
            self.gui.set_status(f"✅ Sociedad {sociedad} completada ({idx}/{len(sociedades)})")
            time.sleep(2)

            self.gui.set_status(f"📥 Descargando FBL5N - {sociedad}...")
            FBL5_FileName = f"FBL5N_Clientes_{sociedad}.xlsx"
            FBL5_File = os.path.join(ClientesPath, FBL5_FileName)
            FBL5_Intercompañias([sociedad], DateFrom, DateTo, ClientesPath, FBL5_FileName)

            time.sleep(5)

            try:
                excel = win32.GetObject(Class="Excel.Application")
                for wb in list(excel.Workbooks):
                    try:
                        if os.path.abspath(wb.FullName) == os.path.abspath(FBL5_File):
                            wb.Close(SaveChanges=False)
                    except Exception:
                        pass
            except:
                pass

            self.gui.set_status(f"✅ Sociedad {sociedad} completada ({idx}/{len(sociedades)})")
            time.sleep(2)
    
    def execute_consolidation(self):
        """Execute the consolidation process"""
        # Get configuration from GUI
        config = self.gui.get_config()
        
        # Confirm action
        confirm = messagebox.askyesno(
            "Confirmar Consolidación", 
            "¿Desea ejecutar el proceso de consolidación?\n\n"
            f"Se leerán archivos de:\n{config['input_path']}\n\n"
            f"El archivo consolidado se guardará en:\n{config['output_path']}"
        )
        
        if not confirm:
            return
        
        # Create output directory if it doesn't exist
        os.makedirs(config['output_path'], exist_ok=True)
        
        # Disable buttons during execution
        self.gui.disable_buttons()
        self.gui.set_status("⏳ Consolidando... Por favor espere")
        
        try:
            # Execute the consolidation process
            self._run_consolidation_process(config)
            
            # Success message
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
            # Re-enable buttons
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")
    
    def _run_consolidation_process(self, config):
        """
        Internal method to execute the consolidation process
        
        Args:
            config: Configuration dictionary from GUI
        """
        user_profile = os.environ.get('USERPROFILE') or os.path.expanduser('~')
        base_path = os.path.join(user_profile, 'Documents', 'Intercompañias', 'RDA_Intercompanias', 'src')
        
        ruta_input = os.path.join(base_path, 'Input')
        ruta_output = os.path.join(base_path, 'Output')
        
        for sociedad in config['sociedades']:
            archivo = ejecutar_consolidacion_por_sociedad(
                ruta_input,
                ruta_output,
                sociedad=sociedad,
                callback_status=self.gui.set_status
            )
        
        # Mostrar resumen
        num_archivos = len(config['sociedades'])
        self.gui.set_status(f"✅ {num_archivos} archivo(s) generado(s)")