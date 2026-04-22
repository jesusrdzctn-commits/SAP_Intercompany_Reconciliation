import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import ast


class IntercompaniasGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Extracción de Documentos - Intercompañías")
        self.root.geometry("900x780")
        self.root.resizable(False, False)

        # Color scheme
        self.bg_color = "#FFFFFF"
        self.primary_color = "#000000"
        self.secondary_color = "#333333"
        self.accent_color = "#2C2C2C"
        self.light_gray = "#F5F5F5"
        self.border_color = "#E0E0E0"
        self.success_color = "#28A745"
        self.large_color = "#1565C0"   # azul para el botón de sociedades grandes

        self.root.configure(bg=self.bg_color)

        # Variables
        self.sociedades_list = []
        self.on_download = None
        self.on_consolidation = None
        self.on_download_large = None   # nuevo callback para descarga por chunks
        self.consolidation_mode = tk.StringVar(value="manual")

        # Diccionarios configurables
        self.cuentas_proveedores_por_sociedad = {
            "MX01": ["6600022", "7201000", "7204000"],
            "MX05": ["7201000"],
            "MX22": ["6600021", "2050000", "6600022", "6700040", "6700043", "6700048", "6900010"],
            "MX30": ["6600022", "7204000", "6900010"],
            "MX73": ["6600022"],
        }

        self.cuentas_clientes_por_sociedad = {
            "MX01": ["7000005", "7000020", "7201000"],
            "MX05": ["7201000"],
            "MX22": ["4300010", "7000005", "7001002", "7010005", "7201000"],
            "MX30": ["7001000", "7001002", "7001005", "7011000", "7500000"],
            "MX31": ["7201000"],
            "MX32": ["7201000"],
            "MX73": ["7000005", "7001002", "7201000"],
            "MX80": ["7201000"],
        }

        # Dynamic paths
        user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        base_path = os.path.join(
            user_profile, "Documents", "Intercompañias", "RDA_Intercompanias", "src"
        )
        self.input_path = os.path.join(base_path, "Input", "Proveedores")
        self.clientes_path = os.path.join(base_path, "Input", "Clientes")
        self.output_path = os.path.join(base_path, "Output")
        self.config_path = os.path.join(base_path, "config")

        self.proveedores_txt_path = os.path.join(
            self.config_path, "cuentas_proveedores_por_sociedad.txt"
        )
        self.clientes_txt_path = os.path.join(
            self.config_path, "cuentas_clientes_por_sociedad.txt"
        )

        self._create_project_folders()
        self._build_ui()
        self._load_all_configs()
        self._refresh_proveedores_tree()
        self._refresh_clientes_tree()

    def _create_project_folders(self):
        folders = [self.input_path, self.clientes_path, self.output_path, self.config_path]
        for folder in folders:
            os.makedirs(folder, exist_ok=True)

    def _build_ui(self):
        main_frame = tk.Frame(self.root, bg=self.bg_color, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)

        title_label = tk.Label(
            main_frame,
            text="Sistema de Extracción de Documentos",
            font=("Segoe UI", 18, "bold"),
            bg=self.bg_color,
            fg=self.primary_color,
        )
        title_label.pack(pady=(0, 20))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook", background=self.bg_color, borderwidth=0)
        style.configure("TNotebook.Tab", padding=[15, 8], font=("Segoe UI", 10, "bold"))
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=24)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)

        self.process_tab = tk.Frame(self.notebook, bg=self.bg_color)
        self.config_tab = tk.Frame(self.notebook, bg=self.bg_color)

        self.notebook.add(self.process_tab, text="Proceso")
        self.notebook.add(self.config_tab, text="Configuración")

        self._build_process_tab()
        self._build_config_tab()
        self._build_status_section(main_frame)

    def _build_process_tab(self):
        process_container = tk.Frame(self.process_tab, bg=self.bg_color, padx=15, pady=15)
        process_container.pack(fill="both", expand=True)

        subtitle = tk.Label(
            process_container,
            text="Inicio de proceso de descarga y conciliación",
            font=("Segoe UI", 12, "bold"),
            bg=self.bg_color,
            fg=self.primary_color,
        )
        subtitle.pack(anchor="w", pady=(0, 15))

        self._build_date_section(process_container)
        self._build_action_buttons(process_container)

    def _build_config_tab(self):
        container = tk.Frame(self.config_tab, bg=self.bg_color)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=self.bg_color, highlightthickness=0, bd=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        config_container = tk.Frame(canvas, bg=self.bg_color, padx=15, pady=15)
        self.config_canvas_window = canvas.create_window((0, 0), window=config_container, anchor="nw")

        config_container.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfig(self.config_canvas_window, width=e.width)
        )
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        )

        subtitle = tk.Label(
            config_container,
            text="Configuración general",
            font=("Segoe UI", 12, "bold"),
            bg=self.bg_color,
            fg=self.primary_color,
        )
        subtitle.pack(anchor="w", pady=(0, 15))

        self._build_sociedades_section(config_container)
        self._build_paths_section(config_container)
        self._build_proveedores_section(config_container)
        self._build_clientes_section(config_container)

    def _build_date_section(self, parent):
        date_frame = tk.LabelFrame(
            parent,
            text="  Intervalo de Tiempo  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color,
            fg=self.primary_color,
            bd=1,
            relief=tk.SOLID,
            padx=15,
            pady=15,
        )
        date_frame.pack(fill="x", pady=(0, 20))

        tk.Label(
            date_frame, text="Fecha Desde:", font=("Segoe UI", 9),
            bg=self.bg_color, fg=self.secondary_color,
        ).grid(row=0, column=0, sticky=tk.W, pady=8)

        self.date_from_var = tk.StringVar(value="01.01.2025")
        self.date_from_entry = tk.Entry(
            date_frame, textvariable=self.date_from_var, width=15,
            font=("Segoe UI", 10), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        )
        self.date_from_entry.grid(row=0, column=1, padx=10, pady=8, sticky=tk.W)

        tk.Label(date_frame, text="(DD.MM.YYYY)", font=("Segoe UI", 8),
                 bg=self.bg_color, fg="#999999").grid(row=0, column=2, sticky=tk.W)

        tk.Label(
            date_frame, text="Fecha Hasta:", font=("Segoe UI", 9),
            bg=self.bg_color, fg=self.secondary_color,
        ).grid(row=1, column=0, sticky=tk.W, pady=8)

        self.date_to_var = tk.StringVar(value="31.12.2025")
        self.date_to_entry = tk.Entry(
            date_frame, textvariable=self.date_to_var, width=15,
            font=("Segoe UI", 10), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        )
        self.date_to_entry.grid(row=1, column=1, padx=10, pady=8, sticky=tk.W)

        tk.Label(date_frame, text="(DD.MM.YYYY)", font=("Segoe UI", 8),
                 bg=self.bg_color, fg="#999999").grid(row=1, column=2, sticky=tk.W)

    def _build_sociedades_section(self, parent):
        sociedades_frame = tk.LabelFrame(
            parent, text="  Sociedades  ", font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color, bd=1, relief=tk.SOLID,
            padx=15, pady=15,
        )
        sociedades_frame.pack(fill="x", pady=(0, 20))

        content_frame = tk.Frame(sociedades_frame, bg=self.bg_color)
        content_frame.pack(fill="x")

        # ── Columna izquierda ──
        left_frame = tk.Frame(content_frame, bg=self.bg_color)
        left_frame.pack(side=tk.LEFT, fill="both", expand=True)

        input_frame = tk.Frame(left_frame, bg=self.bg_color)
        input_frame.pack(fill="x", pady=(0, 10))

        tk.Label(input_frame, text="Agregar Sociedad:", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).pack(side=tk.LEFT, padx=(0, 10))

        self.sociedad_var = tk.StringVar()
        self.sociedad_entry = tk.Entry(
            input_frame, textvariable=self.sociedad_var, width=15,
            font=("Segoe UI", 10), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        )
        self.sociedad_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.sociedad_entry.bind("<Return>", lambda e: self.add_sociedad())

        tk.Button(
            input_frame, text="Agregar", command=self.add_sociedad,
            font=("Segoe UI", 9), bg=self.primary_color, fg=self.bg_color,
            relief=tk.FLAT, bd=0, padx=20, pady=5, cursor="hand2",
            activebackground=self.secondary_color, activeforeground=self.bg_color,
        ).pack(side=tk.LEFT)

        tk.Label(left_frame, text="Sociedades Seleccionadas:", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).pack(anchor="w", pady=(10, 5))

        listbox_frame = tk.Frame(left_frame, bg=self.bg_color)
        listbox_frame.pack(fill="x", pady=5)

        scrollbar = tk.Scrollbar(listbox_frame, bg=self.light_gray)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.sociedades_listbox = tk.Listbox(
            listbox_frame, height=5, width=40, font=("Segoe UI", 9),
            bg=self.light_gray, fg=self.primary_color, relief=tk.FLAT, bd=1,
            highlightthickness=1, highlightbackground=self.border_color,
            highlightcolor=self.primary_color, selectbackground=self.accent_color,
            selectforeground=self.bg_color, yscrollcommand=scrollbar.set,
        )
        self.sociedades_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.sociedades_listbox.yview)

        tk.Button(
            left_frame, text="Eliminar Seleccionada", command=self.remove_sociedad,
            font=("Segoe UI", 9), bg=self.bg_color, fg=self.secondary_color,
            relief=tk.SOLID, bd=1, padx=20, pady=5, cursor="hand2",
            activebackground=self.light_gray, activeforeground=self.primary_color,
        ).pack(pady=(10, 0), anchor="w")

        # ── Separador vertical ──
        tk.Frame(content_frame, bg=self.border_color, width=1).pack(side=tk.LEFT, fill="y", padx=15)

        # ── Columna derecha: rangos ──
        right_frame = tk.Frame(content_frame, bg=self.bg_color)
        right_frame.pack(side=tk.LEFT, fill="y", padx=(0, 5))

        tk.Label(right_frame, text="Rangos de cuentas", font=("Segoe UI", 9, "bold"),
                 bg=self.bg_color, fg=self.primary_color).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        # FBL1N
        tk.Label(right_frame, text="FBL1N (proveedores):", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).grid(row=1, column=0, sticky="w", pady=4)

        self.fbl1n_from_var = tk.StringVar(value="4000000000")
        tk.Entry(right_frame, textvariable=self.fbl1n_from_var, width=13,
                 font=("Segoe UI", 9), bg=self.light_gray, fg=self.primary_color,
                 relief=tk.FLAT, bd=1, highlightthickness=1,
                 highlightbackground=self.border_color, highlightcolor=self.primary_color,
                 ).grid(row=1, column=1, padx=(8, 4), pady=4)

        tk.Label(right_frame, text="—", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).grid(row=1, column=2)

        self.fbl1n_to_var = tk.StringVar(value="7399999999")
        tk.Entry(right_frame, textvariable=self.fbl1n_to_var, width=13,
                 font=("Segoe UI", 9), bg=self.light_gray, fg=self.primary_color,
                 relief=tk.FLAT, bd=1, highlightthickness=1,
                 highlightbackground=self.border_color, highlightcolor=self.primary_color,
                 ).grid(row=1, column=3, padx=(4, 0), pady=4)

        # FBL5N
        tk.Label(right_frame, text="FBL5N (clientes):", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).grid(row=2, column=0, sticky="w", pady=4)

        self.fbl5n_from_var = tk.StringVar(value="200000")
        tk.Entry(right_frame, textvariable=self.fbl5n_from_var, width=13,
                 font=("Segoe UI", 9), bg=self.light_gray, fg=self.primary_color,
                 relief=tk.FLAT, bd=1, highlightthickness=1,
                 highlightbackground=self.border_color, highlightcolor=self.primary_color,
                 ).grid(row=2, column=1, padx=(8, 4), pady=4)

        tk.Label(right_frame, text="—", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).grid(row=2, column=2)

        self.fbl5n_to_var = tk.StringVar(value="299999")
        tk.Entry(right_frame, textvariable=self.fbl5n_to_var, width=13,
                 font=("Segoe UI", 9), bg=self.light_gray, fg=self.primary_color,
                 relief=tk.FLAT, bd=1, highlightthickness=1,
                 highlightbackground=self.border_color, highlightcolor=self.primary_color,
                 ).grid(row=2, column=3, padx=(4, 0), pady=4)

    def _build_paths_section(self, parent):
        paths_frame = tk.LabelFrame(
            parent, text="  Rutas de trabajo  ", font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color, bd=1, relief=tk.SOLID,
            padx=15, pady=15,
        )
        paths_frame.pack(fill="x", pady=(0, 20))

        rows = [
            ("Entrada proveedores:", "input_path_var", self.input_path, "select_input_path"),
            ("Entrada clientes:",    "clientes_path_var", self.clientes_path, "select_clientes_path"),
            ("Carpeta de salida:",   "output_path_var",  self.output_path,   "select_output_path"),
        ]
        for i, (label, varname, default, cmd) in enumerate(rows):
            tk.Label(paths_frame, text=label, font=("Segoe UI", 9),
                     bg=self.bg_color, fg=self.secondary_color).grid(row=i, column=0, sticky="w", pady=8)
            var = tk.StringVar(value=default)
            setattr(self, varname, var)
            tk.Entry(paths_frame, textvariable=var, width=60,
                     font=("Segoe UI", 9), bg=self.light_gray, relief=tk.FLAT,
                     ).grid(row=i, column=1, padx=10, pady=8, sticky="we")
            tk.Button(paths_frame, text="Buscar", command=getattr(self, cmd),
                      font=("Segoe UI", 9), bg=self.primary_color, fg="white",
                      relief=tk.FLAT, padx=15, pady=4,
                      ).grid(row=i, column=2, pady=8)

        paths_frame.columnconfigure(1, weight=1)

    def _build_proveedores_section(self, parent):
        frame = tk.LabelFrame(
            parent, text="  Cuentas proveedores por sociedad  ",
            font=("Segoe UI", 10, "bold"), bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=15,
        )
        frame.pack(fill="both", expand=True, pady=(0, 20))

        form_frame = tk.Frame(frame, bg=self.bg_color)
        form_frame.pack(fill="x", pady=(0, 15))

        tk.Label(form_frame, text="Sociedad:", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).grid(row=0, column=0, sticky="w", pady=6)
        self.prov_sociedad_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.prov_sociedad_var, width=15,
                 font=("Segoe UI", 10), bg=self.light_gray, relief=tk.FLAT,
                 ).grid(row=0, column=1, padx=10, pady=6, sticky="w")

        tk.Label(form_frame, text="Cuentas (separadas por coma):", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).grid(row=1, column=0, sticky="w", pady=6)
        self.prov_cuentas_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.prov_cuentas_var, width=70,
                 font=("Segoe UI", 10), bg=self.light_gray, relief=tk.FLAT,
                 ).grid(row=1, column=1, padx=10, pady=6, sticky="we")
        form_frame.columnconfigure(1, weight=1)

        btns = tk.Frame(frame, bg=self.bg_color)
        btns.pack(fill="x", pady=(0, 10))

        tk.Button(btns, text="Agregar / Actualizar", command=self.add_or_update_proveedor,
                  font=("Segoe UI", 9), bg=self.primary_color, fg=self.bg_color,
                  relief=tk.FLAT, bd=0, padx=20, pady=5, cursor="hand2",
                  ).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btns, text="Eliminar Seleccionada", command=self.remove_proveedor,
                  font=("Segoe UI", 9), bg=self.bg_color, fg=self.secondary_color,
                  relief=tk.SOLID, bd=1, padx=20, pady=5, cursor="hand2",
                  ).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btns, text="Guardar TXT Proveedores", command=self.save_proveedores_to_txt,
                  font=("Segoe UI", 9, "bold"), bg=self.success_color, fg=self.bg_color,
                  relief=tk.FLAT, bd=0, padx=20, pady=5, cursor="hand2",
                  ).pack(side=tk.LEFT)

        tree_frame = tk.Frame(frame, bg=self.bg_color)
        tree_frame.pack(fill="both", expand=True)

        sy = tk.Scrollbar(tree_frame, orient="vertical")
        sy.pack(side="right", fill="y")
        sx = tk.Scrollbar(tree_frame, orient="horizontal")
        sx.pack(side="bottom", fill="x")

        self.proveedores_tree = ttk.Treeview(
            tree_frame, columns=("sociedad", "cuentas"), show="headings",
            yscrollcommand=sy.set, xscrollcommand=sx.set, height=7,
        )
        self.proveedores_tree.pack(fill="both", expand=True)
        sy.config(command=self.proveedores_tree.yview)
        sx.config(command=self.proveedores_tree.xview)

        self.proveedores_tree.heading("sociedad", text="Sociedad")
        self.proveedores_tree.heading("cuentas", text="Cuentas")
        self.proveedores_tree.column("sociedad", width=120, anchor="center")
        self.proveedores_tree.column("cuentas", width=600, anchor="w")
        self.proveedores_tree.bind("<<TreeviewSelect>>", self.on_select_proveedor)

    def _build_clientes_section(self, parent):
        frame = tk.LabelFrame(
            parent, text="  Cuentas clientes por sociedad  ",
            font=("Segoe UI", 10, "bold"), bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=15,
        )
        frame.pack(fill="both", expand=True, pady=(0, 20))

        form_frame = tk.Frame(frame, bg=self.bg_color)
        form_frame.pack(fill="x", pady=(0, 15))

        tk.Label(form_frame, text="Sociedad:", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).grid(row=0, column=0, sticky="w", pady=6)
        self.cli_sociedad_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.cli_sociedad_var, width=15,
                 font=("Segoe UI", 10), bg=self.light_gray, relief=tk.FLAT,
                 ).grid(row=0, column=1, padx=10, pady=6, sticky="w")

        tk.Label(form_frame, text="Cuentas (separadas por coma):", font=("Segoe UI", 9),
                 bg=self.bg_color, fg=self.secondary_color).grid(row=1, column=0, sticky="w", pady=6)
        self.cli_cuentas_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.cli_cuentas_var, width=70,
                 font=("Segoe UI", 10), bg=self.light_gray, relief=tk.FLAT,
                 ).grid(row=1, column=1, padx=10, pady=6, sticky="we")
        form_frame.columnconfigure(1, weight=1)

        btns = tk.Frame(frame, bg=self.bg_color)
        btns.pack(fill="x", pady=(0, 10))

        tk.Button(btns, text="Agregar / Actualizar", command=self.add_or_update_cliente,
                  font=("Segoe UI", 9), bg=self.primary_color, fg=self.bg_color,
                  relief=tk.FLAT, bd=0, padx=20, pady=5, cursor="hand2",
                  ).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btns, text="Eliminar Seleccionada", command=self.remove_cliente,
                  font=("Segoe UI", 9), bg=self.bg_color, fg=self.secondary_color,
                  relief=tk.SOLID, bd=1, padx=20, pady=5, cursor="hand2",
                  ).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btns, text="Guardar TXT Clientes", command=self.save_clientes_to_txt,
                  font=("Segoe UI", 9, "bold"), bg=self.success_color, fg=self.bg_color,
                  relief=tk.FLAT, bd=0, padx=20, pady=5, cursor="hand2",
                  ).pack(side=tk.LEFT)

        tree_frame = tk.Frame(frame, bg=self.bg_color)
        tree_frame.pack(fill="both", expand=True)

        sy = tk.Scrollbar(tree_frame, orient="vertical")
        sy.pack(side="right", fill="y")
        sx = tk.Scrollbar(tree_frame, orient="horizontal")
        sx.pack(side="bottom", fill="x")

        self.clientes_tree = ttk.Treeview(
            tree_frame, columns=("sociedad", "cuentas"), show="headings",
            yscrollcommand=sy.set, xscrollcommand=sx.set, height=7,
        )
        self.clientes_tree.pack(fill="both", expand=True)
        sy.config(command=self.clientes_tree.yview)
        sx.config(command=self.clientes_tree.xview)

        self.clientes_tree.heading("sociedad", text="Sociedad")
        self.clientes_tree.heading("cuentas", text="Cuentas")
        self.clientes_tree.column("sociedad", width=120, anchor="center")
        self.clientes_tree.column("cuentas", width=600, anchor="w")
        self.clientes_tree.bind("<<TreeviewSelect>>", self.on_select_cliente)

    # ------------------------------------------------------------------
    # Pestaña Proceso – botones de acción
    # ------------------------------------------------------------------
    def _build_action_buttons(self, parent):
        button_frame = tk.Frame(parent, bg=self.bg_color)
        button_frame.pack(pady=(10, 0))

        # ── Botón Descargar Normal ─────────────────────────────────────
        self.download_btn = tk.Button(
            button_frame,
            text="⚡ Descargar Documentos",
            command=self._handle_download,
            font=("Segoe UI", 11, "bold"),
            bg=self.primary_color,
            fg=self.bg_color,
            relief=tk.RAISED,
            bd=2,
            padx=30,
            pady=12,
            cursor="hand2",
            activebackground=self.secondary_color,
            activeforeground=self.bg_color,
            width=25,
        )
        self.download_btn.pack(pady=(0, 10))

        # ── Grupo Descarga por Chunks (Sociedades Grandes) ─────────────
        large_group = tk.LabelFrame(
            button_frame,
            text="  Descarga de Sociedades Grandes  ",
            font=("Segoe UI", 9, "bold"),
            bg=self.bg_color,
            fg=self.large_color,
            bd=1,
            relief=tk.SOLID,
            padx=20,
            pady=12,
        )
        large_group.pack(fill="x", pady=(0, 10))

        # Campo: días por chunk
        chunk_frame = tk.Frame(large_group, bg=self.bg_color)
        chunk_frame.pack(anchor="w", pady=(0, 6))

        tk.Label(
            chunk_frame,
            text="Días por chunk:",
            font=("Segoe UI", 9),
            bg=self.bg_color,
            fg=self.secondary_color,
        ).pack(side=tk.LEFT, padx=(0, 10))

        self.chunk_days_var = tk.StringVar(value="5")
        tk.Entry(
            chunk_frame,
            textvariable=self.chunk_days_var,
            width=5,
            font=("Segoe UI", 10),
            bg=self.light_gray,
            fg=self.primary_color,
            relief=tk.FLAT,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.large_color,
            justify="center",
        ).pack(side=tk.LEFT)

        tk.Label(
            chunk_frame,
            text="(ej: 5 → descarga de 5 días en 5 días)",
            font=("Segoe UI", 8),
            bg=self.bg_color,
            fg="#888888",
        ).pack(side=tk.LEFT, padx=(10, 0))

        # Descripción
        tk.Label(
            large_group,
            text=(
                "Divide el intervalo de fechas en bloques para evitar el error de memoria "
                "en SAP. Cada bloque se descarga por separado y luego se apilan en un solo "
                "archivo dentro de Input/Proveedores e Input/Clientes."
            ),
            font=("Segoe UI", 8),
            bg=self.bg_color,
            fg="#666666",
            wraplength=520,
            justify=tk.LEFT,
        ).pack(anchor="w", pady=(0, 10))

        # Botón ejecutar descarga por chunks
        self.download_large_btn = tk.Button(
            large_group,
            text="🏢 Descargar Sociedades Grandes",
            command=self._handle_download_large,
            font=("Segoe UI", 11, "bold"),
            bg=self.large_color,
            fg=self.bg_color,
            relief=tk.RAISED,
            bd=2,
            padx=30,
            pady=12,
            cursor="hand2",
            activebackground="#0D47A1",
            activeforeground=self.bg_color,
            width=25,
        )
        self.download_large_btn.pack()

        # ── Grupo Consolidación ────────────────────────────────────────
        consolidation_group = tk.LabelFrame(
            button_frame,
            text="  Consolidación  ",
            font=("Segoe UI", 9, "bold"),
            bg=self.bg_color,
            fg=self.primary_color,
            bd=1,
            relief=tk.SOLID,
            padx=20,
            pady=12,
        )
        consolidation_group.pack(fill="x", pady=(0, 5))

        mode_frame = tk.Frame(consolidation_group, bg=self.bg_color)
        mode_frame.pack(anchor="w", pady=(0, 6))

        tk.Label(
            mode_frame,
            text="Modo de cuentas:",
            font=("Segoe UI", 9),
            bg=self.bg_color,
            fg=self.secondary_color,
        ).pack(side=tk.LEFT, padx=(0, 12))

        tk.Radiobutton(
            mode_frame,
            text="Manual",
            variable=self.consolidation_mode,
            value="manual",
            font=("Segoe UI", 9),
            bg=self.bg_color,
            fg=self.primary_color,
            activebackground=self.bg_color,
            selectcolor=self.light_gray,
            cursor="hand2",
            command=self._on_mode_change,
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Radiobutton(
            mode_frame,
            text="Automático",
            variable=self.consolidation_mode,
            value="automatico",
            font=("Segoe UI", 9),
            bg=self.bg_color,
            fg=self.primary_color,
            activebackground=self.bg_color,
            selectcolor=self.light_gray,
            cursor="hand2",
            command=self._on_mode_change,
        ).pack(side=tk.LEFT)

        self.mode_desc_var = tk.StringVar(
            value="Usa las cuentas configuradas manualmente en la pestaña Configuración."
        )
        tk.Label(
            consolidation_group,
            textvariable=self.mode_desc_var,
            font=("Segoe UI", 8),
            bg=self.bg_color,
            fg="#666666",
            wraplength=520,
            justify=tk.LEFT,
        ).pack(anchor="w", pady=(0, 12))

        self.consolidate_btn = tk.Button(
            consolidation_group,
            text="📊 Conciliación / Consolidación",
            command=self._handle_consolidation,
            font=("Segoe UI", 11, "bold"),
            bg=self.success_color,
            fg=self.bg_color,
            relief=tk.RAISED,
            bd=2,
            padx=30,
            pady=12,
            cursor="hand2",
            activebackground="#218838",
            activeforeground=self.bg_color,
            width=25,
        )
        self.consolidate_btn.pack()

    def _on_mode_change(self):
        if self.consolidation_mode.get() == "manual":
            self.mode_desc_var.set(
                "Usa las cuentas configuradas manualmente en la pestaña Configuración."
            )
        else:
            self.mode_desc_var.set(
                "Toma TODAS las cuentas presentes en los archivos FBL3N descargados "
                "(proveedores y clientes), sin aplicar ningún filtro por sociedad."
            )

    def _build_status_section(self, parent):
        footer_frame = tk.Frame(parent, bg=self.bg_color)
        footer_frame.pack(fill="x", pady=(15, 0))

        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        tk.Label(
            footer_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            bg=self.bg_color,
            fg="#666666",
        ).pack()

    # =========================
    # PROVEEDORES
    # =========================
    def _refresh_proveedores_tree(self):
        for item in self.proveedores_tree.get_children():
            self.proveedores_tree.delete(item)
        for sociedad in sorted(self.cuentas_proveedores_por_sociedad.keys()):
            cuentas = self.cuentas_proveedores_por_sociedad[sociedad]
            self.proveedores_tree.insert("", "end", values=(sociedad, ", ".join(cuentas)))

    def add_or_update_proveedor(self):
        sociedad = self.prov_sociedad_var.get().strip().upper()
        cuentas_text = self.prov_cuentas_var.get().strip()
        if not sociedad:
            messagebox.showwarning("Advertencia", "Capture una sociedad para proveedores")
            return
        if not cuentas_text:
            messagebox.showwarning("Advertencia", "Capture las cuentas de proveedores separadas por coma")
            return
        cuentas = [c.strip() for c in cuentas_text.split(",") if c.strip()]
        self.cuentas_proveedores_por_sociedad[sociedad] = cuentas
        self._refresh_proveedores_tree()
        self.prov_sociedad_var.set("")
        self.prov_cuentas_var.set("")
        self.set_status(f"✓ Proveedores actualizados para {sociedad}")

    def remove_proveedor(self):
        selected = self.proveedores_tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Seleccione una sociedad de proveedores para eliminar")
            return
        sociedad = self.proveedores_tree.item(selected[0], "values")[0]
        self.cuentas_proveedores_por_sociedad.pop(sociedad, None)
        self._refresh_proveedores_tree()
        self.prov_sociedad_var.set("")
        self.prov_cuentas_var.set("")
        self.set_status(f"✓ Proveedores eliminados para {sociedad}")

    def on_select_proveedor(self, event=None):
        selected = self.proveedores_tree.selection()
        if not selected:
            return
        sociedad, cuentas = self.proveedores_tree.item(selected[0], "values")
        self.prov_sociedad_var.set(sociedad)
        self.prov_cuentas_var.set(cuentas)

    def save_proveedores_to_txt(self):
        try:
            with open(self.proveedores_txt_path, "w", encoding="utf-8") as f:
                f.write("CUENTAS_PROVEEDORES_POR_SOCIEDAD = {\n")
                for sociedad in sorted(self.cuentas_proveedores_por_sociedad.keys()):
                    cuentas = self.cuentas_proveedores_por_sociedad[sociedad]
                    cuentas_str = ", ".join([f'"{c}"' for c in cuentas])
                    f.write(f'    "{sociedad}": [{cuentas_str}],\n')
                f.write("}\n")
            self.set_status("✓ TXT de proveedores guardado correctamente")
            messagebox.showinfo("Éxito", f"Archivo guardado correctamente:\n{self.proveedores_txt_path}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el TXT de proveedores:\n{e}")

    def _load_proveedores_from_txt(self):
        if not os.path.exists(self.proveedores_txt_path):
            return
        try:
            with open(self.proveedores_txt_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
            if "=" in content:
                _, dict_text = content.split("=", 1)
                data = ast.literal_eval(dict_text.strip())
                if isinstance(data, dict):
                    self.cuentas_proveedores_por_sociedad = {
                        str(k).upper(): [str(x).strip() for x in v]
                        for k, v in data.items()
                    }
        except Exception as e:
            messagebox.showwarning(
                "Advertencia",
                f"No se pudo cargar el TXT de proveedores.\nSe usarán valores por defecto.\n\nDetalle:\n{e}"
            )

    # =========================
    # CLIENTES
    # =========================
    def _refresh_clientes_tree(self):
        for item in self.clientes_tree.get_children():
            self.clientes_tree.delete(item)
        for sociedad in sorted(self.cuentas_clientes_por_sociedad.keys()):
            cuentas = self.cuentas_clientes_por_sociedad[sociedad]
            self.clientes_tree.insert("", "end", values=(sociedad, ", ".join(cuentas)))

    def add_or_update_cliente(self):
        sociedad = self.cli_sociedad_var.get().strip().upper()
        cuentas_text = self.cli_cuentas_var.get().strip()
        if not sociedad:
            messagebox.showwarning("Advertencia", "Capture una sociedad para clientes")
            return
        if not cuentas_text:
            messagebox.showwarning("Advertencia", "Capture las cuentas de clientes separadas por coma")
            return
        cuentas = [c.strip() for c in cuentas_text.split(",") if c.strip()]
        self.cuentas_clientes_por_sociedad[sociedad] = cuentas
        self._refresh_clientes_tree()
        self.cli_sociedad_var.set("")
        self.cli_cuentas_var.set("")
        self.set_status(f"✓ Clientes actualizados para {sociedad}")

    def remove_cliente(self):
        selected = self.clientes_tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Seleccione una sociedad de clientes para eliminar")
            return
        sociedad = self.clientes_tree.item(selected[0], "values")[0]
        self.cuentas_clientes_por_sociedad.pop(sociedad, None)
        self._refresh_clientes_tree()
        self.cli_sociedad_var.set("")
        self.cli_cuentas_var.set("")
        self.set_status(f"✓ Clientes eliminados para {sociedad}")

    def on_select_cliente(self, event=None):
        selected = self.clientes_tree.selection()
        if not selected:
            return
        sociedad, cuentas = self.clientes_tree.item(selected[0], "values")
        self.cli_sociedad_var.set(sociedad)
        self.cli_cuentas_var.set(cuentas)

    def save_clientes_to_txt(self):
        try:
            with open(self.clientes_txt_path, "w", encoding="utf-8") as f:
                f.write("CUENTAS_CLIENTES_POR_SOCIEDAD = {\n")
                for sociedad in sorted(self.cuentas_clientes_por_sociedad.keys()):
                    cuentas = self.cuentas_clientes_por_sociedad[sociedad]
                    cuentas_str = ", ".join([f'"{c}"' for c in cuentas])
                    f.write(f'    "{sociedad}": [{cuentas_str}],\n')
                f.write("}\n")
            self.set_status("✓ TXT de clientes guardado correctamente")
            messagebox.showinfo("Éxito", f"Archivo guardado correctamente:\n{self.clientes_txt_path}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el TXT de clientes:\n{e}")

    def _load_clientes_from_txt(self):
        if not os.path.exists(self.clientes_txt_path):
            return
        try:
            with open(self.clientes_txt_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
            if "=" in content:
                _, dict_text = content.split("=", 1)
                data = ast.literal_eval(dict_text.strip())
                if isinstance(data, dict):
                    self.cuentas_clientes_por_sociedad = {
                        str(k).upper(): [str(x).strip() for x in v]
                        for k, v in data.items()
                    }
        except Exception as e:
            messagebox.showwarning(
                "Advertencia",
                f"No se pudo cargar el TXT de clientes.\nSe usarán valores por defecto.\n\nDetalle:\n{e}"
            )

    def _load_all_configs(self):
        self._load_proveedores_from_txt()
        self._load_clientes_from_txt()

    # ===== Helpers de rutas =====
    def select_input_path(self):
        path = filedialog.askdirectory(title="Selecciona carpeta de entrada (proveedores)")
        if path:
            path = os.path.normpath(path)
            self.input_path_var.set(path)
            self.input_path = path

    def select_clientes_path(self):
        path = filedialog.askdirectory(title="Selecciona carpeta de entrada (clientes)")
        if path:
            path = os.path.normpath(path)
            self.clientes_path_var.set(path)
            self.clientes_path = path

    def select_output_path(self):
        path = filedialog.askdirectory(title="Selecciona carpeta de salida")
        if path:
            path = os.path.normpath(path)
            self.output_path_var.set(path)
            self.output_path = path

    # ===== Event Handlers =====
    def add_sociedad(self):
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
        selection = self.sociedades_listbox.curselection()
        if not selection:
            messagebox.showwarning("Advertencia", "Por favor seleccione una sociedad para eliminar")
            return
        index = selection[0]
        sociedad = self.sociedades_listbox.get(index)
        self.sociedades_listbox.delete(index)
        self.sociedades_list.remove(sociedad)

    def validate_dates(self):
        try:
            fecha_desde = datetime.strptime(self.date_from_var.get(), "%d.%m.%Y")
            fecha_hasta = datetime.strptime(self.date_to_var.get(), "%d.%m.%Y")
            if fecha_desde > fecha_hasta:
                messagebox.showerror("Error", "La fecha desde no puede ser mayor a la fecha hasta")
                return False
            return True
        except ValueError:
            messagebox.showerror("Error", "Formato de fecha inválido. Use DD.MM.YYYY")
            return False

    def validate_chunk_days(self):
        """Valida que el campo de días por chunk sea un entero positivo."""
        try:
            days = int(self.chunk_days_var.get().strip())
            if days <= 0:
                raise ValueError
            return days
        except ValueError:
            messagebox.showerror("Error", "Los días por chunk deben ser un número entero positivo.")
            return None

    def get_config(self):
        self.input_path = os.path.normpath(self.input_path_var.get().strip())
        self.clientes_path = os.path.normpath(self.clientes_path_var.get().strip())
        self.output_path = os.path.normpath(self.output_path_var.get().strip())

        return {
            "sociedades": self.sociedades_list,
            "date_from": self.date_from_var.get(),
            "date_to": self.date_to_var.get(),
            "input_path": self.input_path,
            "clientes_path": self.clientes_path,
            "output_path": self.output_path,
            "fbl1n_range_from": self.fbl1n_from_var.get().strip(),
            "fbl1n_range_to":   self.fbl1n_to_var.get().strip(),
            "fbl5n_range_from": self.fbl5n_from_var.get().strip(),
            "fbl5n_range_to":   self.fbl5n_to_var.get().strip(),
            "cuentas_proveedores_por_sociedad": self.cuentas_proveedores_por_sociedad,
            "cuentas_clientes_por_sociedad": self.cuentas_clientes_por_sociedad,
            "consolidation_mode": self.consolidation_mode.get(),
            "chunk_days": self.chunk_days_var.get().strip(),
        }

    def set_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

    def disable_buttons(self):
        self.download_btn.config(state="disabled", bg="#666666")
        self.download_large_btn.config(state="disabled", bg="#666666")
        self.consolidate_btn.config(state="disabled", bg="#666666")

    def enable_buttons(self):
        self.download_btn.config(state="normal", bg=self.primary_color)
        self.download_large_btn.config(state="normal", bg=self.large_color)
        self.consolidate_btn.config(state="normal", bg=self.success_color)

    def _handle_download(self):
        if not self.validate_dates():
            return
        if self.on_download:
            self.on_download()
        else:
            messagebox.showinfo("Info", f"Funcionalidad no conectada\n\n{self.get_config()}")

    def _handle_download_large(self):
        if not self.validate_dates():
            return
        if self.validate_chunk_days() is None:
            return
        if self.on_download_large:
            self.on_download_large()
        else:
            messagebox.showinfo("Info", f"Funcionalidad no conectada\n\n{self.get_config()}")

    def _handle_consolidation(self):
        if not self.validate_dates():
            return
        if self.on_consolidation:
            self.on_consolidation()
        else:
            messagebox.showinfo("Info", f"Funcionalidad no conectada\n\n{self.get_config()}")


def main():
    root = tk.Tk()
    app = IntercompaniasGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
