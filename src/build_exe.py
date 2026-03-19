"""
Script para crear el ejecutable del Sistema de Intercompañías
Ejecutar desde la carpeta donde están los .py: python build_exe.py
"""

import subprocess
import sys
import os


def install_pyinstaller():
    """Instalar PyInstaller si no está disponible"""
    try:
        import PyInstaller
        print("✅ PyInstaller ya está instalado")
    except ImportError:
        print("📦 Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✅ PyInstaller instalado correctamente")


def create_executable():
    """Crear el ejecutable"""
    print("🚀 Creando ejecutable...")

    cmd = [
        "pyinstaller",
        "--onefile",            # Un solo archivo .exe
        "--windowed",           # Sin ventana de consola (app con GUI)
        "--name=Intercompanias",# Nombre del ejecutable
        "--icon=NONE",          # Sin icono personalizado (cambiar si hay un .ico)

        # Módulos ocultos que PyInstaller no detecta automáticamente
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.cell._writer",
        "--hidden-import=pandas",
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=pyperclip",
        "--hidden-import=xlrd",

        # Punto de entrada
        "main.py"
    ]

    try:
        subprocess.run(cmd, check=True)
        print("✅ Ejecutable creado exitosamente!")
        print("📁 Ubicación: dist/Intercompanias.exe")

        exe_path = os.path.join("dist", "Intercompanias.exe")
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"📊 Tamaño del ejecutable: {size_mb:.1f} MB")

            print("\n📋 INSTRUCCIONES PARA EL USUARIO:")
            print("1. Distribuir el archivo: dist/Intercompanias.exe")
            print("2. El usuario solo necesita ejecutar el .exe")
            print("3. No requiere tener Python instalado")
            print("4. SAP GUI debe estar abierto y conectado antes de usar la descarga")
            print("5. La primera ejecución creará automáticamente las carpetas necesarias en Documentos")
        else:
            print("❌ Error: No se encontró el ejecutable después de la compilación")

    except subprocess.CalledProcessError as e:
        print(f"❌ Error creando ejecutable: {e}")
        print("\n💡 Sugerencias:")
        print("   - Verifica que estés en la carpeta donde están los archivos .py")
        print("   - Ejecuta: pip install pyinstaller")
        print("   - Si el error menciona 'win32com', ejecuta: pip install pywin32")
        print("   - Si el error menciona 'openpyxl', ejecuta: pip install openpyxl")


def check_source_files():
    """Verificar que todos los archivos fuente necesarios existen"""
    archivos_requeridos = [
        "main.py",
        "interfaz_GUI.py",
        "controller.py",
        "DescargaSAP.py",
        "Consolidacion_V2.py",
    ]

    print("🔍 Verificando archivos fuente...")
    todos_ok = True
    for archivo in archivos_requeridos:
        if os.path.exists(archivo):
            print(f"   ✅ {archivo}")
        else:
            print(f"   ❌ {archivo} — NO ENCONTRADO")
            todos_ok = False

    return todos_ok


def main():
    print("🔨 CONSTRUCCIÓN DE EJECUTABLE - SISTEMA INTERCOMPAÑÍAS")
    print("=" * 60)

    # Verificar archivos fuente
    if not check_source_files():
        print("\n❌ Faltan archivos fuente. Asegúrate de ejecutar este script")
        print("   desde la carpeta donde están todos los archivos .py del proyecto.")
        return

    print()

    # Instalar PyInstaller si es necesario
    install_pyinstaller()

    print()

    # Crear ejecutable
    create_executable()

    print("\n🎉 ¡Proceso completado!")


if __name__ == "__main__":
    main()
