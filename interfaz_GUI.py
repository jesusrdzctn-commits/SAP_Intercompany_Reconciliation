import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import time
import pandas as pd
import win32com.client as win32


class IntercompaniasGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Extracción de Documentos - Intercompañías")
        self.root.geometry("650x620")
        self.root.resizable(False, False)
        
        # Modern color scheme
        self.bg_color = "#FFFFFF"
        self.primary_color = "#000000"
        self.secondary_color = "#333333"
        self.accent_color = "#2C2C2C"
        self.light_gray = "#F5F5F5"
        self.border_color = "#E0E0E0"
        
        # Configure root background
        self.root.configure(bg=self.bg_color)
        
        # Variables
        self.sociedades_list = []
        
        # Get dynamic output path
        user_profile = os.environ.get('USERPROFILE') or os.path.expanduser('~')
        self.output_path = os.path.join(user_profile, 'Documents', 'Intercompañias', 'src', 'Output')
        
        # Create main frame with modern styling
        main_frame = tk.Frame(root, bg=self.bg_color, padx=30, pady=20)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title with modern font
        title_label = tk.Label(
            main_frame, 
            text="Sistema de Extracción de Documentos", 
            font=('Segoe UI', 18, 'bold'),
            bg=self.bg_color,
            fg=self.primary_color
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 25))
        
        # ===== Date Section =====
        date_frame = tk.LabelFrame(
            main_frame, 
            text="  Intervalo de Tiempo  ", 
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.primary_color,
            bd=1,
            relief=tk.SOLID,
            padx=15,
            pady=15
        )
        date_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 18))
        
        # FROM Date
        tk.Label(
            date_frame, 
            text="Fecha Desde:", 
            font=('Segoe UI', 9),
            bg=self.bg_color,
            fg=self.secondary_color
        ).grid(row=0, column=0, sticky=tk.W, pady=8)
        
        self.date_from_var = tk.StringVar(value="01.01.2025")
        self.date_from_entry = tk.Entry(
            date_frame, 
            textvariable=self.date_from_var, 
            width=15,
            font=('Segoe UI', 10),
            bg=self.light_gray,
            fg=self.primary_color,
            relief=tk.FLAT,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.primary_color
        )
        self.date_from_entry.grid(row=0, column=1, padx=10, pady=8, sticky=tk.W)
        
        tk.Label(
            date_frame, 
            text="(DD.MM.YYYY)", 
            font=('Segoe UI', 8),
            bg=self.bg_color,
            fg="#999999"
        ).grid(row=0, column=2, sticky=tk.W)
        
        # TO Date
        tk.Label(
            date_frame, 
            text="Fecha Hasta:", 
            font=('Segoe UI', 9),
            bg=self.bg_color,
            fg=self.secondary_color
        ).grid(row=1, column=0, sticky=tk.W, pady=8)
        
        self.date_to_var = tk.StringVar(value="31.12.2025")
        self.date_to_entry = tk.Entry(
            date_frame, 
            textvariable=self.date_to_var, 
            width=15,
            font=('Segoe UI', 10),
            bg=self.light_gray,
            fg=self.primary_color,
            relief=tk.FLAT,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.primary_color
        )
        self.date_to_entry.grid(row=1, column=1, padx=10, pady=8, sticky=tk.W)
        
        tk.Label(
            date_frame, 
            text="(DD.MM.YYYY)", 
            font=('Segoe UI', 8),
            bg=self.bg_color,
            fg="#999999"
        ).grid(row=1, column=2, sticky=tk.W)
        
        # ===== Sociedades Section =====
        sociedades_frame = tk.LabelFrame(
            main_frame, 
            text="  Sociedades  ", 
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.primary_color,
            bd=1,
            relief=tk.SOLID,
            padx=15,
            pady=15
        )
        sociedades_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 18))
        
        # Input for new sociedad
        input_frame = tk.Frame(sociedades_frame, bg=self.bg_color)
        input_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        tk.Label(
            input_frame, 
            text="Agregar Sociedad:", 
            font=('Segoe UI', 9),
            bg=self.bg_color,
            fg=self.secondary_color
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        self.sociedad_var = tk.StringVar()
        self.sociedad_entry = tk.Entry(
            input_frame, 
            textvariable=self.sociedad_var, 
            width=15,
            font=('Segoe UI', 10),
            bg=self.light_gray,
            fg=self.primary_color,
            relief=tk.FLAT,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.primary_color
        )
        self.sociedad_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.sociedad_entry.bind('<Return>', lambda e: self.add_sociedad())
        
        # Add button with modern style
        add_btn = tk.Button(
            input_frame, 
            text="Agregar",
            command=self.add_sociedad,
            font=('Segoe UI', 9),
            bg=self.primary_color,
            fg=self.bg_color,
            relief=tk.FLAT,
            bd=0,
            padx=20,
            pady=5,
            cursor="hand2",
            activebackground=self.secondary_color,
            activeforeground=self.bg_color
        )
        add_btn.pack(side=tk.LEFT)
        
        # Listbox to show sociedades
        tk.Label(
            sociedades_frame, 
            text="Sociedades Seleccionadas:", 
            font=('Segoe UI', 9),
            bg=self.bg_color,
            fg=self.secondary_color
        ).grid(row=1, column=0, sticky=tk.W, pady=(10, 5))
        
        listbox_frame = tk.Frame(sociedades_frame, bg=self.bg_color)
        listbox_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        scrollbar = tk.Scrollbar(listbox_frame, bg=self.light_gray)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.sociedades_listbox = tk.Listbox(
            listbox_frame, 
            height=5, 
            width=60,
            font=('Segoe UI', 9),
            bg=self.light_gray,
            fg=self.primary_color,
            relief=tk.FLAT,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.primary_color,
            selectbackground=self.accent_color,
            selectforeground=self.bg_color,
            yscrollcommand=scrollbar.set
        )
        self.sociedades_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.sociedades_listbox.yview)
        
        # Remove button
        remove_btn = tk.Button(
            sociedades_frame, 
            text="Eliminar Seleccionada",
            command=self.remove_sociedad,
            font=('Segoe UI', 9),
            bg=self.bg_color,
            fg=self.secondary_color,
            relief=tk.SOLID,
            bd=1,
            padx=20,
            pady=5,
            cursor="hand2",
            activebackground=self.light_gray,
            activeforeground=self.primary_color
        )
        remove_btn.grid(row=3, column=0, columnspan=3, pady=(10, 0))
        
        # ===== Action Buttons =====
        button_frame = tk.Frame(main_frame, bg=self.bg_color)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(20, 0))
        
        # Main action button with modern style - MORE VISIBLE
        self.run_btn = tk.Button(
            button_frame, 
            text="⚡ Correr Descarga de Documentos",
            command=self.run_download,
            font=('Segoe UI', 12, 'bold'),
            bg=self.primary_color,
            fg=self.bg_color,
            relief=tk.RAISED,
            bd=2,
            padx=35,
            pady=15,
            cursor="hand2",
            activebackground=self.secondary_color,
            activeforeground=self.bg_color
        )
        self.run_btn.pack(pady=15)
        
        # Status label
        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        status_label = tk.Label(
            main_frame, 
            textvariable=self.status_var,
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg="#666666"
        )
        status_label.grid(row=4, column=0, columnspan=3, pady=(5, 0))
        
        # Output path label
        output_label = tk.Label(
            main_frame,
            text=f"📁 Carpeta de salida: {self.output_path}",
            font=('Segoe UI', 8),
            bg=self.bg_color,
            fg="#999999",
            wraplength=550
        )
        output_label.grid(row=5, column=0, columnspan=3, pady=(8, 0))
    
    def add_sociedad(self):
        """Add a sociedad to the list"""
        sociedad = self.sociedad_var.get().strip().upper()
        
        if not sociedad:
            messagebox.showwarning("Advertencia", "Por favor ingrese una sociedad")
            return
        
        if sociedad in self.sociedades_list:
            messagebox.showwarning("Advertencia", f"La sociedad {sociedad} ya está en la lista")
            return
        
        self.sociedades_list.append(sociedad)
        self.sociedades_listbox.insert(tk.END, sociedad)
        self.sociedad_var.set("")
        self.sociedad_entry.focus()
    
    def remove_sociedad(self):
        """Remove selected sociedad from the list"""
        selection = self.sociedades_listbox.curselection()
        
        if not selection:
            messagebox.showwarning("Advertencia", "Por favor seleccione una sociedad para eliminar")
            return
        
        index = selection[0]
        sociedad = self.sociedades_listbox.get(index)
        
        self.sociedades_listbox.delete(index)
        self.sociedades_list.remove(sociedad)
    
    def validate_dates(self):
        """Validate date format"""
        try:
            datetime.strptime(self.date_from_var.get(), "%d.%m.%Y")
            datetime.strptime(self.date_to_var.get(), "%d.%m.%Y")
            return True
        except ValueError:
            messagebox.showerror("Error", "Formato de fecha inválido. Use DD.MM.YYYY")
            return False
    
    def run_download(self):
        """Execute the document download process"""
        # Validate inputs
        if not self.sociedades_list:
            messagebox.showerror("Error", "Debe agregar al menos una sociedad")
            return
        
        if not self.validate_dates():
            return
        
        # Confirm action
        confirm = messagebox.askyesno(
            "Confirmar", 
            f"¿Desea iniciar la descarga para {len(self.sociedades_list)} sociedad(es)?\n"
            f"Periodo: {self.date_from_var.get()} - {self.date_to_var.get()}\n\n"
            f"Los archivos se guardarán en:\n{self.output_path}"
        )
        
        if not confirm:
            return
        
        # Create output directory if it doesn't exist
        os.makedirs(self.output_path, exist_ok=True)
        
        # Disable button during execution
        self.run_btn.config(state='disabled', bg="#666666")
        self.status_var.set("⏳ Procesando... Por favor espere")
        self.root.update()
        
        try:
            # Import functions here to avoid import errors if module not available
            from FBL1_Intercompañias import FBL1_Intercompañias, ZFIQ02_Intercompañias, FBL3N
            
            # Execute the download process
            DateFrom = self.date_from_var.get()
            DateTo = self.date_to_var.get()
            FolderPath = self.output_path
            FileName = "FBL1_Intercompañias.xlsx"
            sociedades = self.sociedades_list
            
            # Step 1: FBL1 Download
            self.status_var.set("📥 Descargando FBL1...")
            self.root.update()
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
            self.status_var.set("📥 Descargando ZFIQ02...")
            self.root.update()
            ZFIQ02_FolderPath = FolderPath
            ZFIQ02_FileName = "ZFIQ02_Intercompañias.xlsx"
            ZFIQ02_Intercompañias_File = os.path.join(ZFIQ02_FolderPath, ZFIQ02_FileName)
            ZFIQ02_Intercompañias(sociedades, ZFIQ02_Intercompañias_File)
            
            # Step 3: Read FBL1 and extract document numbers
            self.status_var.set("🔄 Procesando documentos...")
            self.root.update()
            df_FBL1 = pd.read_excel(FBL1_Intercompañias_File, engine='openpyxl')
            colNDocument = df_FBL1.columns[6]
            resultado = df_FBL1[colNDocument].dropna().astype(str).unique().tolist()
            
            # Step 4: FBL3N Download
            self.status_var.set("📥 Descargando FBL3N...")
            self.root.update()
            arr_Sociedades = "\n".join(sociedades)
            FBL3N(resultado, arr_Sociedades, DateFrom, DateTo)
            
            # Success message
            self.status_var.set("✅ ¡Proceso completado exitosamente!")
            messagebox.showinfo(
                "Éxito", 
                "La descarga de documentos se completó correctamente.\n\n"
                f"Archivos guardados en:\n{FolderPath}"
            )
            
        except Exception as e:
            self.status_var.set("❌ Error en el proceso")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso:\n\n{str(e)}")
        
        finally:
            # Re-enable button
            self.run_btn.config(state='normal', bg=self.primary_color)
            if "completado exitosamente" not in self.status_var.get():
                self.status_var.set("✓ Listo para comenzar")


def main():
    root = tk.Tk()
    app = IntercompaniasGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
