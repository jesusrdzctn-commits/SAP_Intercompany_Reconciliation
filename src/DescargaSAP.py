import pandas as pd
import os
import time
from datetime import datetime, timedelta
import win32com.client
import openpyxl as _oxl  # ensure PyInstaller bundles the engine
import locale
import win32com.client as win32
import pyperclip
import traceback
import pywintypes


def FBL1N_Intercompañias(sociedades, DateFrom, Date_To, FolderPath, FileName, account_from="4000000000", account_to="7399999999"):
    session = None
    ruta_completa = os.path.join(FolderPath, FileName)

    try:
        # =========================
        # Validaciones iniciales
        # =========================
        if not sociedades:
            raise ValueError("La lista 'sociedades' está vacía.")
        if not isinstance(sociedades, (list, tuple)):
            raise TypeError("'sociedades' debe ser una lista o tupla.")

        os.makedirs(FolderPath, exist_ok=True)

        # =========================
        # Conexión SAP
        # =========================
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
        except Exception as e:
            raise ConnectionError(
                "No fue posible conectarse a SAP GUI. "
                "Verifica que SAP esté abierto y con una sesión activa."
            ) from e

        # =========================
        # Inicio transacción
        # =========================
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL1"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").text  = account_from
        session.findById("wnd[0]/usr/ctxtKD_LIFNR-HIGH").text = account_to
        session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").caretPosition = 2

        session.findById("wnd[0]/usr/radX_AISEL").setFocus()
        session.findById("wnd[0]/usr/radX_AISEL").selected = True
        session.findById("wnd[0]/usr/chkX_SHBV").selected = True
        session.findById("wnd[0]/usr/chkX_MERK").selected = True
        session.findById("wnd[0]/usr/chkX_PARK").selected = True

        # =========================
        # Cargar sociedades
        # =========================
        session.findById("wnd[0]/usr/btn%_KD_BUKRS_%_APP_%-VALU_PUSH").press()

        for i, valor in enumerate(sociedades, start=0):
            try:
                field_id = (
                    f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                    f"ssubSCREEN_HEADER:SAPLALDB:3010/"
                    f"tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
                )
                session.findById(field_id).text = valor
                session.findById(field_id).setFocus()
                session.findById(field_id).caretPosition = len(valor)
            except Exception as e:
                raise RuntimeError(
                    f"Error al capturar la sociedad '{valor}' en la fila {i}."
                ) from e

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # =========================
        # Fechas y variante
        # =========================
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text  = DateFrom
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = Date_To
        session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/TAXVJG"
        session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
        session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 7

        # =========================
        # Ejecutar reporte
        # =========================
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(4)

        if _verificar_sin_partidas(session, ruta_completa):
            return False  # sin datos — archivo vacío ya creado

        # =========================
        # Exportar archivo
        # =========================
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        return True  # con datos

    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP] FBL1N_Intercompañias: {e}")
        return False

    except (ValueError, TypeError, ConnectionError, RuntimeError) as e:
        print(f"[ERROR CONTROLADO] FBL1N_Intercompañias: {e}")
        return False

    except Exception as e:
        print(f"[ERROR NO CONTROLADO] FBL1N_Intercompañias: {e}")
        print(traceback.format_exc())
        return False

    finally:
        if session is not None:
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass


def _verificar_sin_partidas(session, ruta_archivo):
    """
    Detecta el mensaje MSITEM033 / MSITEM030 'No se ha seleccionado ninguna partida'.
    Si aparece, cierra el popup, crea un Excel vacío como placeholder y
    devuelve True para que el llamador sepa que no hay datos y debe saltar
    la exportación normal.
    """
    try:
        if not ruta_archivo or not isinstance(ruta_archivo, str):
            raise ValueError("La ruta del archivo es inválida.")

        sbar_text  = ""
        sin_partidas = False

        try:
            sbar_text = session.findById("wnd[0]/sbar").Text.strip()
        except pywintypes.com_error:
            sbar_text = ""
        except Exception:
            sbar_text = ""

        MENSAJES_SIN_DATOS = ("MSITEM033", "MSITEM030", "ninguna partida", "ninguna cuenta")

        if any(m in sbar_text or m in sbar_text.lower() for m in MENSAJES_SIN_DATOS):
            sin_partidas = True

        if not sin_partidas:
            try:
                popup_text = session.findById("wnd[1]/usr/txtMESSTXT1").Text.strip()
                if any(m in popup_text or m in popup_text.lower() for m in MENSAJES_SIN_DATOS):
                    sin_partidas = True
            except pywintypes.com_error:
                pass
            except Exception:
                pass

        if sin_partidas:
            print(f"[SAP] Sin partidas para este criterio → creando archivo vacío: {ruta_archivo}")
            try:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except Exception:
                pass
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass
            try:
                carpeta = os.path.dirname(ruta_archivo)
                if carpeta:
                    os.makedirs(carpeta, exist_ok=True)
                wb = _oxl.Workbook()
                ws = wb.active
                ws.title = "Sin datos"
                wb.save(ruta_archivo)
                wb.close()
            except Exception as e:
                raise RuntimeError(
                    f"No se pudo crear el archivo vacío en: {ruta_archivo}"
                ) from e
            return True

        return False

    except ValueError as e:
        print(f"[ERROR VALIDACIÓN] _verificar_sin_partidas: {e}")
        return False
    except RuntimeError as e:
        print(f"[ERROR ARCHIVO] _verificar_sin_partidas: {e}")
        return False
    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP] _verificar_sin_partidas: {e}")
        return False
    except Exception as e:
        print(f"[ERROR NO CONTROLADO] _verificar_sin_partidas: {e}")
        return False


def _cerrar_popup_subsidiaria(session, max_intentos=100):
    """
    Detecta y cierra el popup de subsidiaria/central que aparece en FBL5N.
    Se presiona 'Continuar' (Enter / btn[0]) hasta que desaparezca.
    """
    for _ in range(max_intentos):
        try:
            ventana = session.findById("wnd[1]")
            titulo  = ventana.Text.strip().lower()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(0.5)
            print(f"[FBL5N] Popup cerrado (título: '{titulo}')")
        except Exception:
            break


def FBL5_Intercompañias(sociedades, DateFrom, Date_To, FolderPath, FileName, account_from="200000", account_to="299999"):
    session = None
    ruta_completa = os.path.join(FolderPath, FileName)

    try:
        # =========================
        # Validaciones iniciales
        # =========================
        if not sociedades:
            raise ValueError("La lista 'sociedades' está vacía.")
        if not isinstance(sociedades, (list, tuple)):
            raise TypeError("'sociedades' debe ser una lista o tupla.")
        if not FolderPath or not isinstance(FolderPath, str):
            raise ValueError("'FolderPath' no es válido.")
        if not FileName or not isinstance(FileName, str):
            raise ValueError("'FileName' no es válido.")

        os.makedirs(FolderPath, exist_ok=True)

        # =========================
        # Conexión SAP
        # =========================
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
        except Exception as e:
            raise ConnectionError(
                "No fue posible conectarse a SAP GUI. "
                "Verifica que SAP esté abierto y con una sesión activa."
            ) from e

        # =========================
        # Inicio transacción
        # =========================
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL5"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text  = account_from
        session.findById("wnd[0]/usr/ctxtDD_KUNNR-HIGH").text = account_to
        session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").caretPosition = 0

        # =========================
        # Cargar sociedades
        # =========================
        session.findById("wnd[0]/usr/btn%_DD_BUKRS_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()

        for i, valor in enumerate(sociedades, start=0):
            try:
                field_id = (
                    f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                    f"ssubSCREEN_HEADER:SAPLALDB:3010/"
                    f"tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
                )
                session.findById(field_id).text = valor
                session.findById(field_id).setFocus()
                session.findById(field_id).caretPosition = len(valor)
            except Exception as e:
                raise RuntimeError(
                    f"Error al capturar la sociedad '{valor}' en la fila {i}."
                ) from e

        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # =========================
        # Parámetros de ejecución
        # =========================
        session.findById("wnd[0]/usr/radX_AISEL").select()
        session.findById("wnd[0]/usr/chkX_SHBV").selected = True
        session.findById("wnd[0]/usr/chkX_MERK").selected = True
        session.findById("wnd[0]/usr/chkX_PARK").selected = True
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text  = DateFrom
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = Date_To
        session.findById("wnd[0]/usr/ctxtPA_VARI").text = "PYTHON"

        # =========================
        # Ejecutar reporte
        # =========================
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(2)

        try:
            _cerrar_popup_subsidiaria(session)
        except Exception as e:
            raise RuntimeError("Error al intentar cerrar el popup de subsidiaria.") from e

        if _verificar_sin_partidas(session, ruta_completa):
            return False  # sin datos — archivo vacío ya creado

        # =========================
        # Exportar archivo
        # =========================
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        return True  # con datos

    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP] FBL5_Intercompañias: {e}")
        return False

    except (ValueError, TypeError, ConnectionError, RuntimeError) as e:
        print(f"[ERROR CONTROLADO] FBL5_Intercompañias: {e}")
        return False

    except Exception as e:
        print(f"[ERROR NO CONTROLADO] FBL5_Intercompañias: {e}")
        print(traceback.format_exc())
        return False

    finally:
        if session is not None:
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass


def ZFIQ02_Intercompañias(sociedades, ZFIQ02_Intercompañias_File):
    session = None
    excel   = None

    try:
        # =========================
        # Validaciones iniciales
        # =========================
        if not sociedades:
            raise ValueError("La lista 'sociedades' está vacía.")
        if not isinstance(sociedades, (list, tuple)):
            raise TypeError("'sociedades' debe ser una lista o tupla.")
        if not ZFIQ02_Intercompañias_File or not isinstance(ZFIQ02_Intercompañias_File, str):
            raise ValueError("'ZFIQ02_Intercompañias_File' no es válido.")

        carpeta_destino = os.path.dirname(ZFIQ02_Intercompañias_File)
        if carpeta_destino:
            os.makedirs(carpeta_destino, exist_ok=True)

        # =========================
        # Conexión SAP
        # =========================
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
        except Exception as e:
            raise ConnectionError(
                "No fue posible conectarse a SAP GUI. "
                "Verifica que SAP esté abierto y con una sesión activa."
            ) from e

        # =========================
        # Inicio transacción
        # =========================
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZFIQ02"
        session.findById("wnd[0]").sendVKey(0)

        # =========================
        # Cargar sociedades
        # =========================
        session.findById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").press()

        for i, valor in enumerate(sociedades, start=0):
            try:
                field_id = (
                    f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                    f"ssubSCREEN_HEADER:SAPLALDB:3010/"
                    f"tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
                )
                session.findById(field_id).text = valor
                session.findById(field_id).setFocus()
                session.findById(field_id).caretPosition = len(valor)
            except Exception as e:
                raise RuntimeError(
                    f"Error al capturar la sociedad '{valor}' en la fila {i}."
                ) from e

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # =========================
        # Ejecutar reporte
        # =========================
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]").sendVKey(7)
        time.sleep(2)

        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
            "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]"
        ).select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
            "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]"
        ).setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # =========================
        # Conectar Excel y guardar
        # =========================
        try:
            excel = win32.GetObject(Class="Excel.Application")
            excel.DisplayAlerts = False
        except Exception as e:
            raise RuntimeError(
                "No fue posible conectarse a Excel abierto. "
                "Verifica que el archivo exportado se haya abierto en Excel."
            ) from e

        try:
            if os.path.exists(ZFIQ02_Intercompañias_File):
                os.remove(ZFIQ02_Intercompañias_File)
        except Exception as e:
            raise RuntimeError(
                f"No se pudo eliminar el archivo existente: {ZFIQ02_Intercompañias_File}"
            ) from e

        workbook_guardado = False
        for wb in list(excel.Workbooks):
            try:
                wb.SaveAs(ZFIQ02_Intercompañias_File)
                wb.Close(SaveChanges=False)
                workbook_guardado = True
                print(f"Archivo guardado correctamente en: {ZFIQ02_Intercompañias_File}")
                break
            except Exception:
                continue

        if not workbook_guardado:
            raise RuntimeError(
                "No fue posible identificar o guardar el workbook exportado desde SAP."
            )

        time.sleep(2)

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        return True

    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP/Excel] ZFIQ02_Intercompañias: {e}")
        return False

    except (ValueError, TypeError, ConnectionError, RuntimeError) as e:
        print(f"[ERROR CONTROLADO] ZFIQ02_Intercompañias: {e}")
        return False

    except Exception as e:
        print(f"[ERROR NO CONTROLADO] ZFIQ02_Intercompañias: {e}")
        print(traceback.format_exc())
        return False

    finally:
        if excel is not None:
            try:
                excel.DisplayAlerts = True
            except Exception:
                pass
        if session is not None:
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass


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
    session    = None
    ruta_archivo = os.path.join(FolderPath, FileName)

    try:
        # =========================
        # Validaciones iniciales
        # =========================
        if not resultado:
            raise ValueError("La lista 'resultado' está vacía.")
        if not isinstance(resultado, (list, tuple)):
            raise TypeError("'resultado' debe ser una lista o tupla.")
        if not sociedades:
            raise ValueError("La lista 'sociedades' está vacía.")
        if not isinstance(sociedades, (list, tuple)):
            raise TypeError("'sociedades' debe ser una lista o tupla.")
        if not FolderPath or not isinstance(FolderPath, str):
            raise ValueError("'FolderPath' no es válido.")
        if not FileName or not isinstance(FileName, str):
            raise ValueError("'FileName' no es válido.")

        os.makedirs(FolderPath, exist_ok=True)

        if os.path.exists(ruta_archivo):
            os.remove(ruta_archivo)
            print(f"Archivo anterior eliminado: {ruta_archivo}")

        # =========================
        # Conexión SAP
        # =========================
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            if application.Children.Count == 0:
                raise ConnectionError("SAP está abierto, pero no hay conexión activa.")
            connection = application.Children(0)
            if connection.Children.Count == 0:
                raise ConnectionError("Hay conexión SAP, pero no hay sesión activa.")
            session = connection.Children(0)
        except ConnectionError:
            raise
        except Exception as e:
            raise ConnectionError(
                "No fue posible conectarse a SAP GUI. "
                "Verifica que SAP esté abierto y con una sesión activa."
            ) from e

        # =========================
        # Inicio transacción
        # =========================
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL3N"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/tbar[1]/btn[16]").press()

        arbol = (
            "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/"
            "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/"
            "shellcont/shellcont/shell/shellcont[1]/shell"
        )
        session.findById(arbol).collapseNode("          1")
        session.findById(arbol).collapseNode("         13")
        session.findById(arbol).selectNode("         35")
        session.findById(arbol).topNode = "          1"
        session.findById(arbol).doubleClickNode("         35")

        session.findById(
            "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/"
            "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:"
            "SAPLSSEL:1106/btn%_%%DYN006_%_APP_%-VALU_PUSH"
        ).press()

        # =========================
        # Cargar documentos
        # =========================
        try:
            documents = os.linesep.join(map(str, resultado))
            pyperclip.copy(documents)
        except Exception as e:
            raise RuntimeError("No fue posible copiar los documentos al portapapeles.") from e

        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # =========================
        # Parámetros de ejecución
        # =========================
        session.findById("wnd[0]/usr/radX_AISEL").select()
        session.findById("wnd[0]/usr/radX_AISEL").setFocus()

        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text  = DateFrom
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = DateTo
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 10

        # =========================
        # Cargar sociedades
        # =========================
        session.findById("wnd[0]/usr/btn%_SD_BUKRS_%_APP_%-VALU_PUSH").press()

        for i, valor in enumerate(sociedades, start=0):
            try:
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
            except Exception as e:
                raise RuntimeError(
                    f"Error al capturar la sociedad '{valor}' en la fila {i}."
                ) from e

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # =========================
        # Layout y ejecución
        # =========================
        session.findById("wnd[0]/usr/ctxtPA_VARI").text = "RDA_FBL3N"
        session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
        session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 9
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        time.sleep(2)

        # =========================
        # Exportar archivo
        # =========================
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(FileName)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # =========================
        # Esperar y regresar
        # =========================
        archivo_descargado = esperar_archivo(ruta_archivo, timeout=30)

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

    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP] FBL3N: {e}")
        return False

    except (ValueError, TypeError, ConnectionError, RuntimeError) as e:
        print(f"[ERROR CONTROLADO] FBL3N: {e}")
        return False

    except Exception as e:
        print(f"[ERROR NO CONTROLADO] FBL3N: {e}")
        print(traceback.format_exc())
        return False

    finally:
        if session is not None:
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass
