import pandas as pd
import os
import time
from datetime import datetime, timedelta
import win32com.client
import openpyxl  # ensure PyInstaller bundles the engine
import locale
import win32com.client as win32
import pyperclip

def FBL1_Intercompañias(sociedades,DateFrom, Date_To, FolderPath, FileName):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "FBL1"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").text = "6000000000"
    session.findById("wnd[0]/usr/ctxtKD_LIFNR-HIGH").text = "6999999999"
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

        print(f"Fila {i} completada con valor {valor}")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = DateFrom
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = Date_To
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/TAXVJG"
    session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 7
    #Descarga de archivo
    session.findById("wnd[0]").sendVKey(8)
    #session.findById("wnd[0]/tbar[1]/btn[8]").press
    #session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").press
    session.findById("wnd[0]").sendVKey(16)
    time.sleep(2)
    session.findById("wnd[0]").sendVKey(8)
    #session.findById("wnd[1]/tbar[0]/btn[0]").Setfocus
    #session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    #session.findById("wnd[0]/tbar[0]/btn[3]").press()

def ZFIQ02_Intercompañias(sociedades,ZFIQ02_Intercompañias_File):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZFIQ02"
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

        print(f"Fila {i} completada con valor {valor}")
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

def FBL3N(resultado,sociedades,DateFrom, DateTo):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "FBL3N"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[16]").press()
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").collapseNode("          1")
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").collapseNode( "         13")
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode("         35")
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "          1"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode("         35")
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN006_%_APP_%-VALU_PUSH").press()
    #session.findById("wnd[1]/tbar[0]/btn[24]").press()
    for i, valor in enumerate(resultado, start=0):
    # Construimos el ID dinámicamente usando el contador en la parte [1,i]
        field_id = (
            f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
            f"ssubSCREEN_HEADER:SAPLALDB:3010/"
            f"tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,{i}]"
        )

        # Asignamos el texto
        session.findById(field_id).text = valor

        # Ponemos el foco
        session.findById(field_id).setFocus()

        # Ajustamos la posición del cursor (ejemplo: al final del texto)
        session.findById(field_id).caretPosition = len(valor)

        print(f"Fila {i} completada con valor {valor}")
        
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/radX_AISEL").select()
    session.findById("wnd[0]/usr/radX_AISEL").setFocus()
    session.findById("wnd[0]/tbar[1]/btn[16]").press()
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = DateFrom
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = DateTo
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
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

        print(f"Fila {i} completada con valor {valor}")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").setFocus
    session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").caretPosition = 4
    session.findById("wnd[0]/tbar[1]/btn[8]").press()


DateFrom = "01.01.2025"
DateTo = "31.12.2025"
FolderPath = r"C:\Users\80337365\Documents\Intercompañias"
FileName = "FBL1_Intercompañias.xlsx"
#sociedades = ["MX70","MX00","MX30","MX01", "MX60", "MX21"]
sociedades = ["MX73","MX30","MX80","MX31","MX60","MX01"]
FBL1_Intercompañias(sociedades,DateFrom, DateTo, FolderPath, FileName)
FBL1_Intercompañias_File = os.path.join(FolderPath, FileName)
time.sleep(5)  # Wait for the file to be fully saved

excel = win32.GetObject(Class="Excel.Application")  # se conecta a Excel abierto

for wb in list(excel.Workbooks):
        try:
            if os.path.abspath(wb.FullName) == FBL1_Intercompañias_File:
                wb.Close(SaveChanges=False)
        except Exception:
            pass

ZFIQ02_FolderPath = r"C:\Users\80337365\Documents\Intercompañias"
ZFIQ02_FileName = "ZFIQ02_Intercompañias.xlsx"
ZFIQ02_Intercompañias_File = os.path.join(ZFIQ02_FolderPath, ZFIQ02_FileName)
ZFIQ02_Intercompañias(sociedades,ZFIQ02_Intercompañias_File)
df_FBL1 = pd.read_excel(FBL1_Intercompañias_File, engine='openpyxl')
colNDocument = df_FBL1.columns[6]

resultado = df_FBL1[colNDocument].dropna().astype(str).unique().tolist()

pd.Series([resultado]).to_clipboard(index=False, header=False)

arr_Sociedades = "\n".join(sociedades)
FBL3N(resultado,arr_Sociedades,DateFrom, DateTo)
#ZFIQ02 = pd.read_excel(ZFIQ02_Intercompañias_File, engine='openpyxl')
#print(ZFIQ02.head())
