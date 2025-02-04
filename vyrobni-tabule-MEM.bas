Option Explicit

Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub aktualizace_zak_mem()

Aktualizace.Range("K15").Value = "OK"
Aktualizace.Range("K16").Value = "x"

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "COOIS"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[17]").Press
session.FindById("wnd[1]/usr/txtENAME-LOW").Text = "hoskoond"
session.FindById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.FindById("wnd[1]/usr/txtENAME-LOW").CaretPosition = 8
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "0"
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").Text = Aktualizace.Range("AC9")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").Text = Aktualizace.Range("AC10")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").SetFocus
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").CaretPosition = 9
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectContextMenuItem "&XXL"
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Plan tabule\MEM"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT_ZAK.XLSX"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 10
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

MsgBox "Data propsána, následuje aktualizace Směsí a KZ"

End Sub

Sub aktualizace_KZSM_mem()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "COOIS"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[17]").Press
session.FindById("wnd[1]/usr/txtENAME-LOW").Text = "sedlajos"
session.FindById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.FindById("wnd[1]/usr/txtENAME-LOW").CaretPosition = 8
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "0"
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").Text = Aktualizace.Range("AC9")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").Text = Aktualizace.Range("AC10")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").SetFocus
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").CaretPosition = 9
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectContextMenuItem "&XXL"
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Plan tabule\MEM"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT_KOMP.XLSX"
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

Aktualizace.Range("K16").Value = "OK"
Aktualizace.Range("AC6") = Date
Aktualizace.Range("AC7") = Now
Aktualizace.Range("AC8") = Environ("username")
MsgBox "Aktualizace výrobní tabule MEM dokončena"

End Sub
