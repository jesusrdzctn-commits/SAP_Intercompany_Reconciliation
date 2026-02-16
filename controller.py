import os
import time
import pandas as pd
import win32com.client as win32
from tkinter import messagebox
from FBL1_Intercompañias import FBL1_Intercompañias, ZFIQ02_Intercompañias, FBL3N
from Consolidacion_V2 import ejecutar_consolidacion


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
        self.gui.run_download = self.execute_download
        self.gui.run_consolidation = self.execute_consolidation
    
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
        FileName = "FBL1_Intercompañias.xlsx"
        sociedades = config['sociedades']
        
        # Step 1: FBL1 Download
        self.gui.set_status("📥 Descargando FBL1...")
        FBL1_Intercompañias(sociedades, DateFrom, DateTo, FolderPath, FileName)
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
        
        # Step 2: ZFIQ02 Download
        self.gui.set_status("📥 Descargando ZFIQ02...")
        ZFIQ02_FolderPath = FolderPath
        ZFIQ02_FileName = "ZFIQ02_Intercompañias.xlsx"
        ZFIQ02_Intercompañias_File = os.path.join(ZFIQ02_FolderPath, ZFIQ02_FileName)
        ZFIQ02_Intercompañias(sociedades, ZFIQ02_Intercompañias_File)
        
        # Step 3: Read FBL1 and extract document numbers
        self.gui.set_status("📄 Procesando documentos...")
        df_FBL1 = pd.read_excel(FBL1_Intercompañias_File, engine='openpyxl')
        colNDocument = df_FBL1.columns[6]
        resultado = df_FBL1[colNDocument].dropna().astype(str).unique().tolist()
        
        # Step 4: FBL3N Download
        self.gui.set_status("📥 Descargando FBL3N...")
        arr_Sociedades = "\n".join(sociedades)
        FBL3N(resultado, arr_Sociedades, DateFrom, DateTo)
    
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
        # Get user profile for dynamic paths
        user_profile = os.environ.get('USERPROFILE') or os.path.expanduser('~')
        base_path = os.path.join(user_profile, 'Documents', 'Intercompañias', 'RDA_Intercompanias', 'src')
        
        ruta_input = os.path.join(base_path, 'Input')
        ruta_output = os.path.join(base_path, 'Output')
        
        # Ejecutar consolidación usando la función importada
        # Pasamos el método set_status de la GUI como callback para actualizar el estado
        self.gui.set_status("📊 Iniciando consolidación...")
        archivo_consolidado = ejecutar_consolidacion(
            ruta_input, 
            ruta_output,
            callback_status=self.gui.set_status
        )
        
        self.gui.set_status(f"✅ Archivo generado: {os.path.basename(archivo_consolidado)}")
