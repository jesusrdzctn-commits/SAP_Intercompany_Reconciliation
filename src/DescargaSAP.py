import pandas as pd
import os
import time
from datetime import datetime, timedelta
import win32com.client
import openpyxl  # ensure PyInstaller bundles the engine
import locale
import win32com.client as win32
import pyperclip
import traceback

def FBL1N_Intercompañias(sociedades,DateFrom, Date_To, FolderPath, FileName, account_from="4000000000", account_to="7399999999"):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL1"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").text  = account_from
    session.findById("wnd[0]/usr/ctxtKD_LIFNR-HIGH").text = account_to
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").setFocus
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").caretPosition = 2
    #session.findById("wnd[0]").sendVKey(2)
    #session.findById("wnd[1]").close
    session.findById("wnd[0]/usr/radX_AISEL").setFocus
    session.findById("wnd[0]/usr/radX_AISEL").selected = True
    session.findById("wnd[0]/usr/chkX_SHBV").selected = True
    session.findById("wnd[0]/usr/chkX_MERK").selected = True
    session.findById("wnd[0]/usr/chkX_PARK").selected = True
    session.findById("wnd[0]/usr/btn%_KD_BUKRS_%_APP_%-VALU_PUSH").press()
    for i, valor in enumerate(sociedades, start=0):
    # Construimos el ID dinámicamente usando el contador en la parte [1,i]
        field_id = (
            f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
            f"ssubSCREEN_HEADER:SAPLALDB:3010/"
            f"tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
        )

        # Asignamos el texto
        session.findById(field_id).text = valor

        # Ponemos el foco
        session.findById(field_id).setFocus()

        # Ajustamos la posición del cursor (ejemplo: al final del texto)
        session.findById(field_id).caretPosition = len(valor)

        #print(f"Fila {i} completada con valor {valor}")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = DateFrom
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = Date_To
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/TAXVJG"
    session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 7
    #Descarga de archivo
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(4)
    # --- Detectar mensaje "Sin partidas" (MSITEM033) ---
    ruta_completa = os.path.join(FolderPath, FileName)
    if _verificar_sin_partidas(session, ruta_completa):
        return False  # sin datos — archivo vacío ya creado
    
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    return True  # con datos

def _verificar_sin_partidas(session, ruta_archivo):
    """
    Detecta el mensaje MSITEM033 'No se ha seleccionado ninguna partida'.
    Si aparece, cierra el popup, crea un Excel vacío como placeholder y
    devuelve True para que el llamador sepa que no hay datos y debe saltar
    la exportación normal.

    Args:
        session   : Sesión SAP activa.
        ruta_archivo (str): Ruta completa donde se crearía el archivo vacío.

    Returns:
        bool: True si SAP reportó sin partidas (archivo vacío creado),
              False si hay datos y se puede continuar normalmente.
    """
    try:
        # SAP muestra el error en la barra de status (wnd[0]/sbar) o como popup wnd[1]
        # Intentamos leer el texto de la barra de estado primero
        sbar_text = session.findById("wnd[0]/sbar").Text.strip()
    except Exception:
        sbar_text = ""

    MENSAJES_SIN_DATOS = ("MSITEM033", "MSITEM030", "ninguna partida", "ninguna cuenta")

    sin_partidas = any(m in sbar_text or m in sbar_text.lower() for m in MENSAJES_SIN_DATOS)

    # También puede venir como wnd[1] modal
    if not sin_partidas:
        try:
            msg = session.findById("wnd[1]/usr/txtMESSTXT1").Text.strip()
            sin_partidas = any(m in msg or m in msg.lower() for m in MENSAJES_SIN_DATOS)
        except Exception:
            pass

    if sin_partidas:
        print(f"[SAP] Sin partidas para este criterio → creando archivo vacío: {ruta_archivo}")
        # Cerrar el popup si existe
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            pass
        # Volver a pantalla de selección (F3)
        try:
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
        except Exception:
            pass
        # Crear Excel vacío para que el resto del flujo no falle por archivo faltante
        import openpyxl as _oxl
        wb = _oxl.Workbook()
        wb.active.title = "Sin datos"
        os.makedirs(os.path.dirname(ruta_archivo), exist_ok=True)
        wb.save(ruta_archivo)
        return True

    return False
    
def _cerrar_popup_subsidiaria(session, max_intentos=100):
    """
    Detecta y cierra el popup de subsidiaria/central que aparece en FBL5N.
    El diálogo avisa que una cuenta es sucursal y pregunta si incluir partidas
    de la central. Se presiona 'Continuar' (Enter / btn[0]) hasta que desaparezca.

    Args:
        session: Sesión SAP activa.
        max_intentos: Número máximo de cuentas/popups a confirmar antes de abortar.
    """
    for _ in range(max_intentos):
        try:
            # wnd[1] existe → puede ser el popup de subsidiaria u otro diálogo
            ventana = session.findById("wnd[1]")
            titulo = ventana.Text.strip().lower()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(0.5)
            print(f"[FBL5N] Popup cerrado (título: '{titulo}')")
        except Exception:
            # wnd[1] ya no existe → no hay más popups pendientes
            break

def FBL5_Intercompañias(sociedades,DateFrom, Date_To, FolderPath, FileName, account_from="200000", account_to="299999"):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL5"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = account_from
    session.findById("wnd[0]/usr/ctxtDD_KUNNR-HIGH").text = account_to
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").setFocus
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_DD_BUKRS_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    for i, valor in enumerate(sociedades, start=0):
    # Construimos el ID dinámicamente usando el contador en la parte [1,i]
        field_id = (
            f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
            f"ssubSCREEN_HEADER:SAPLALDB:3010/"
            f"tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
        )

        # Asignamos el texto
        session.findById(field_id).text = valor

        # Ponemos el foco
        session.findById(field_id).setFocus()

        # Ajustamos la posición del cursor (ejemplo: al final del texto)
        session.findById(field_id).caretPosition = len(valor)

        #print(f"Fila {i} completada con valor {valor}")

    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/radX_AISEL").select()
    session.findById("wnd[0]/usr/chkX_SHBV").selected = True
    session.findById("wnd[0]/usr/chkX_MERK").selected = True
    session.findById("wnd[0]/usr/chkX_PARK").selected = True
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = DateFrom
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = Date_To
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "PYTHON"
    #Descarga de archivo
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(2)
    _cerrar_popup_subsidiaria(session)

    # --- Detectar mensaje "Sin partidas" (MSITEM033) ---
    ruta_completa = os.path.join(FolderPath, FileName)
    if _verificar_sin_partidas(session, ruta_completa):
        return False  # sin datos — archivo vacío ya creado

    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    return True  # con datos

def ZFIQ02_Intercompañias(sociedades,ZFIQ02_Intercompañias_File):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZFIQ02"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").press()
    for i, valor in enumerate(sociedades, start=0):
    # Construimos el ID dinámicamente usando el contador en la parte [1,i]
        field_id = (
            f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
            f"ssubSCREEN_HEADER:SAPLALDB:3010/"
            f"tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
        )

        # Asignamos el texto
        session.findById(field_id).text = valor

        # Ponemos el foco
        session.findById(field_id).setFocus()

        # Ajustamos la posición del cursor (ejemplo: al final del texto)
        session.findById(field_id).caretPosition = len(valor)

        #print(f"Fila {i} completada con valor {valor}")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(7)
    time.sleep(2)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    excel = win32.GetObject(Class="Excel.Application")  # se conecta a Excel abierto
    excel.DisplayAlerts = True
    for wb in list(excel.Workbooks):
            try:
                if os.path.exists(ZFIQ02_Intercompañias_File):
                    os.remove(ZFIQ02_Intercompañias_File)

                wb.SaveAs(ZFIQ02_Intercompañias_File)
                wb.Close(SaveChanges=False)
                excel.DisplayAlerts = True
            except Exception:
                pass

    #session.findById("wnd[0]").sendVKey(8)
    time.sleep(2)
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    #session.findById("wnd[0]/tbar[0]/btn[3]").press()

def esperar_archivo(ruta_archivo, timeout=30):
    """
    Espera hasta que el archivo exista y tenga tamaño mayor a 0.
    Retorna True si existe, False si no aparece dentro del tiempo límite.
    """
    tiempo_inicio = time.time()

    while time.time() - tiempo_inicio < timeout:
        if os.path.exists(ruta_archivo) and os.path.getsize(ruta_archivo) > 0:
            return True

        time.sleep(1)

    return False


def FBL3N(resultado, sociedades, DateFrom, DateTo, FolderPath, FileName):
    ruta_archivo = os.path.join(FolderPath, FileName)

    try:
        print("Iniciando descarga FBL3N...")

        # Validaciones iniciales
        if not resultado:
            print("No hay documentos para procesar.")
            return False

        if not sociedades:
            print("No hay sociedades configuradas.")
            return False

        if not DateFrom or not DateTo:
            print("Fechas no configuradas correctamente.")
            return False

        # Crear carpeta si no existe
        os.makedirs(FolderPath, exist_ok=True)

        # Si el archivo ya existe, lo eliminamos para validar una descarga nueva
        if os.path.exists(ruta_archivo):
            os.remove(ruta_archivo)
            print(f"Archivo anterior eliminado: {ruta_archivo}")

        # Conexión SAP
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine

        if application.Children.Count == 0:
            print("SAP está abierto, pero no hay conexión activa.")
            return False

        connection = application.Children(0)

        if connection.Children.Count == 0:
            print("Hay conexión SAP, pero no hay sesión activa.")
            return False

        session = connection.Children(0)

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL3N"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/tbar[1]/btn[16]").press()

        session.findById(
            "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/"
            "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/"
            "shellcont/shellcont/shell/shellcont[1]/shell"
        ).collapseNode("          1")

        session.findById(
            "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/"
            "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/"
            "shellcont/shellcont/shell/shellcont[1]/shell"
        ).collapseNode("         13")

        session.findById(
            "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/"
            "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/"
            "shellcont/shellcont/shell/shellcont[1]/shell"
        ).selectNode("         35")

        session.findById(
            "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/"
            "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/"
            "shellcont/shellcont/shell/shellcont[1]/shell"
        ).topNode = "          1"

        session.findById(
            "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/"
            "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/"
            "shellcont/shellcont/shell/shellcont[1]/shell"
        ).doubleClickNode("         35")

        session.findById(
            "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/"
            "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:"
            "SAPLSSEL:1106/btn%_%%DYN006_%_APP_%-VALU_PUSH"
        ).press()

        # Pegar documentos desde portapapeles
        documents = os.linesep.join(map(str, resultado))
        pyperclip.copy(documents)

        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/radX_AISEL").select()
        session.findById("wnd[0]/usr/radX_AISEL").setFocus()

        session.findById("wnd[0]/tbar[1]/btn[16]").press()

        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = DateFrom
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = DateTo
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 10

        # Sociedades
        session.findById("wnd[0]/usr/btn%_SD_BUKRS_%_APP_%-VALU_PUSH").press()

        for i, valor in enumerate(sociedades, start=0):
            valor = str(valor).strip().upper()

            field_id = (
                f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                f"ssubSCREEN_HEADER:SAPLALDB:3010/"
                f"tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
            )

            campo = session.findById(field_id)
            campo.text = valor
            campo.setFocus()
            campo.caretPosition = len(valor)

            print(f"Sociedad agregada: {valor}")

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # Layout
        session.findById("wnd[0]/usr/ctxtPA_VARI").text = "RDA_FBL3N"
        session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
        session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 9

        # Ejecutar
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        time.sleep(2)

        # Exportar
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(FileName)

        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # Esperar archivo
        archivo_descargado = esperar_archivo(ruta_archivo, timeout=30)

        # Regresar en SAP
        try:
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
        except Exception:
            print("No se pudo regresar en SAP, pero la descarga ya fue validada.")

        if archivo_descargado:
            print(f"Archivo descargado correctamente: {ruta_archivo}")
            return True
        else:
            print(f"No se encontró el archivo descargado: {ruta_archivo}")
            return False

    except Exception as e:
        print("Error en descarga FBL3N")
        print(f"Detalle: {e}")
        print(traceback.format_exc())

        return False