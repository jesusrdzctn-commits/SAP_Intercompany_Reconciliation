import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os


class IntercompaniasGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Extracción de Documentos - Intercompañías")
        self.root.geometry("650x730")
        self.root.resizable(False, False)
        
        # Modern color scheme
        self.bg_color = "#FFFFFF"
        self.primary_color = "#000000"
        self.secondary_color = "#333333"
        self.accent_color = "#2C2C2C"
        self.light_gray = "#F5F5F5"
        self.border_color = "#E0E0E0"
        self.success_color = "#28A745"
        
        # Configure root background
        self.root.configure(bg=self.bg_color)
        
        # Variables
        self.sociedades_list = []
        self.on_download = None
        self.on_consolidation = None
        
        # Get dynamic output path
        user_profile = os.environ.get('USERPROFILE') or os.path.expanduser('~')
        base_path = os.path.join(user_profile, 'Documents', 'Intercompañias', 'RDA_Intercompanias', 'src')
        self.input_path = os.path.join(base_path, 'Input', 'Proveedores')
        self.clientes_path = os.path.join(base_path, 'Input', 'Clientes')
        self.output_path = os.path.join(base_path, 'Output')
        
        self._build_ui()
    
    def _build_ui(self):
        """Build the complete user interface"""
        # Create main frame with modern styling
        main_frame = tk.Frame(self.root, bg=self.bg_color, padx=30, pady=20)
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
        self._build_date_section(main_frame)
        
        # ===== Sociedades Section =====
        self._build_sociedades_section(main_frame)
        
        # ===== Action Buttons =====
        self._build_action_buttons(main_frame)
        
        # ===== Status and Output Info =====
        self._build_status_section(main_frame)
    
    def _build_date_section(self, parent):
        """Build the date input section"""
        date_frame = tk.LabelFrame(
            parent, 
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
    
    def _build_sociedades_section(self, parent):
        """Build the sociedades management section"""
        sociedades_frame = tk.LabelFrame(
            parent, 
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
    
    def _build_action_buttons(self, parent):
        """Build the main action buttons"""
        button_frame = tk.Frame(parent, bg=self.bg_color)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(20, 0))
        
        # Download button
        self.download_btn = tk.Button(
            button_frame, 
            text="⚡ Descargar Documentos",
            command=self._handle_download,
            font=('Segoe UI', 11, 'bold'),
            bg=self.primary_color,
            fg=self.bg_color,
            relief=tk.RAISED,
            bd=2,
            padx=30,
            pady=12,
            cursor="hand2",
            activebackground=self.secondary_color,
            activeforeground=self.bg_color
        )
        self.download_btn.pack(pady=(0, 10))
        
        # Consolidation button
        self.consolidate_btn = tk.Button(
            button_frame, 
            text="📊 Consolidación",
            command=self._handle_consolidation,
            font=('Segoe UI', 11, 'bold'),
            bg=self.success_color,
            fg=self.bg_color,
            relief=tk.RAISED,
            bd=2,
            padx=30,
            pady=12,
            cursor="hand2",
            activebackground="#218838",
            activeforeground=self.bg_color
        )
        self.consolidate_btn.pack()
    
    def _build_status_section(self, parent):
        """Build status and information section"""
        # Status label
        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        status_label = tk.Label(
            parent, 
            textvariable=self.status_var,
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg="#666666"
        )
        status_label.grid(row=4, column=0, columnspan=3, pady=(15, 0))
        
        # Output path labels
        input_label = tk.Label(
            parent,
            text=f"📥 Carpeta de entrada: {self.input_path}",
            font=('Segoe UI', 8),
            bg=self.bg_color,
            fg="#999999",
            wraplength=550
        )
        input_label.grid(row=5, column=0, columnspan=3, pady=(8, 0))
        
        output_label = tk.Label(
            parent,
            text=f"📤 Carpeta de salida: {self.output_path}",
            font=('Segoe UI', 8),
            bg=self.bg_color,
            fg="#999999",
            wraplength=550
        )
        output_label.grid(row=6, column=0, columnspan=3, pady=(3, 0))
    
    # ===== Event Handlers =====
    
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
    
    def get_config(self):
        """Get current configuration as dictionary"""
        return {
            'sociedades': self.sociedades_list,
            'date_from': self.date_from_var.get(),
            'date_to': self.date_to_var.get(),
            'input_path': self.input_path,
            'clientes_path': self.clientes_path,
            'output_path': self.output_path
        }
    
    def set_status(self, message):
        """Update status message"""
        self.status_var.set(message)
        self.root.update()
    
    def disable_buttons(self):
        """Disable all action buttons"""
        self.download_btn.config(state='disabled', bg="#666666")
        self.consolidate_btn.config(state='disabled', bg="#666666")
    
    def enable_buttons(self):
        """Enable all action buttons"""
        self.download_btn.config(state='normal', bg=self.primary_color)
        self.consolidate_btn.config(state='normal', bg=self.success_color)
    
    # ===== Action Methods (to be implemented in controller) =====
    
    def _handle_download(self):
        if self.on_download:
            self.on_download()
        else:
            messagebox.showinfo("Info", "Funcionalidad no conectada")

    def _handle_consolidation(self):
        if self.on_consolidation:
            self.on_consolidation()
        else:
            messagebox.showinfo("Info", "Funcionalidad no conectada")

def main():
    """Main entry point for standalone GUI testing"""
    root = tk.Tk()
    app = IntercompaniasGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
