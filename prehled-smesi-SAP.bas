Option Explicit
Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub SAP_mb52()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

Worksheets("AKTUALIZACE").Range("I9") = "Probíhá stahování MB52 ze SAP..."

'session.findById("wnd[0]").Maximize
'session.findById("wnd[0]/tbar[0]/okcd").Text = "mb52"
'session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = "4*"
'session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "1130"

session.findById("wnd[0]").Maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "MB52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "4*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "6*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "1130"

session.findById("wnd[0]/usr/ctxtLGORT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "3140"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "3121"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "3123"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "3124"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "3125"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "3192"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[1,4]").SetFocus
session.findById("wnd[1]/usr/lbl[1,4]").caretPosition = 2
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\Makra exporty"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Export_mb52_smesi.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Worksheets("AKTUALIZACE").Range("I9") = "Probíhá stahování ZMB5M ze SAP..."

Call SAP_zmb5m

End Sub

Sub SAP_zmb5m()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

session.findById("wnd[0]").Maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "zmb5m"
session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = "4*"
'session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "1130"

session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "4*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "6*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "1130"


'session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
session.findById("wnd[0]/usr/ctxtLGORT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "3140"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "3121"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "3123"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "3124"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "3125"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "3192"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/txtRESTZEIT").Text = "1000"
session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 7
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[1,3]").SetFocus
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 4
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").FirstVisibleRow = 11
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").FirstVisibleRow = 0
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\Makra exporty"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Export_zmb5m_smesi.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]").ResizeWorkingPane 238, 38, False
session.findById("wnd[0]/tbar[0]/btn[3]").press

Worksheets("AKTUALIZACE").Range("I9") = "Stahování ze SAP hotovo..."

Call Najit_poznamky

End Sub
