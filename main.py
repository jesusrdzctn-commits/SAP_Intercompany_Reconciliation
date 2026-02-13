"""
Sistema de Extracción y Consolidación de Documentos - Intercompañías

Este es el punto de entrada principal de la aplicación.
Conecta la interfaz gráfica (GUI) con la lógica de negocio (Controller).

Autor: Sistema de Intercompañías
Fecha: 2026
"""

import tkinter as tk
from interfaz_GUI import IntercompaniasGUI
from controller import IntercompaniasController


def main():
    """
    Función principal que inicializa la aplicación
    """
    # Crear ventana principal de Tkinter
    root = tk.Tk()
    
    # Inicializar la GUI
    gui = IntercompaniasGUI(root)
    
    # Inicializar el controlador y conectarlo con la GUI
    controller = IntercompaniasController(gui)
    
    # Iniciar el loop de eventos de la aplicación
    root.mainloop()


if __name__ == "__main__":
    main()
