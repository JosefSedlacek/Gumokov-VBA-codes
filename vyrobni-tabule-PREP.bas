Option Explicit
Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub AktualizaceFosfat()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

Aktualizace.Range("F15").Value = "running"

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "COOIS"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[17]").Press
session.FindById("wnd[1]/usr/txtENAME-LOW").Text = "balvidav"
session.FindById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.FindById("wnd[1]/usr/txtENAME-LOW").CaretPosition = 8
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow = 3
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "3"
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").Text = Aktualizace.Range("AB9")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").Text = Aktualizace.Range("AB10")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").SetFocus
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").CaretPosition = 8
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectContextMenuItem "&XXL"
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Plan tabule\PREP"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "fosfat.XLSX"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 6
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

Aktualizace.Range("F15").Value = "OK"
Aktualizace.Range("F16").Value = "x"
Aktualizace.Range("F17").Value = "x"
Aktualizace.Range("F18").Value = "x"

MsgBox "Zakázky PREP aktualizovány, následuje ODMAŠTĚNÍ"

End Sub

Sub AktualizaceOdmtry()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

Aktualizace.Range("F16").Value = "running"

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "COOIS"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[17]").Press
session.FindById("wnd[1]/usr/txtENAME-LOW").Text = "balvidav"
session.FindById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.FindById("wnd[1]/usr/txtENAME-LOW").CaretPosition = 8
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow = 5
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "5"
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").Text = Aktualizace.Range("AB9")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").Text = Aktualizace.Range("AB10")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").SetFocus
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").CaretPosition = 8
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectContextMenuItem "&XXL"
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Plan tabule\PREP"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Odmtry.XLSX"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 6
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

Aktualizace.Range("F16").Value = "OK"

MsgBox "Data ODMAŠTĚNÍ PREP aktualizovány, následuje NÁNOS"

End Sub

Sub AktualizaceNanos()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

Aktualizace.Range("F17").Value = "running"

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "COOIS"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[17]").Press
session.FindById("wnd[1]/usr/txtENAME-LOW").Text = "balvidav"
session.FindById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.FindById("wnd[1]/usr/txtENAME-LOW").CaretPosition = 8
session.FindById("wnd[1]").SendVKey 0
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow = 4
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "4"
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").Text = Aktualizace.Range("AB9")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").Text = Aktualizace.Range("AB10")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").SetFocus
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").CaretPosition = 8
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectContextMenuItem "&XXL"
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Plan tabule\PREP"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Nanos.XLSX"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 5
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

Aktualizace.Range("F17").Value = "OK"

MsgBox "Data NÁNOS PREP aktualizovány, následují KOMPONENTY"

End Sub

Sub AktualizaceKomponentyPREP()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

Aktualizace.Range("F18").Value = "running"

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "COOIS"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[17]").Press
session.FindById("wnd[1]/usr/txtENAME-LOW").Text = "BALVIDAV"
session.FindById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.FindById("wnd[1]/usr/txtENAME-LOW").CaretPosition = 8
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow = 6
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "6"
session.FindById("wnd[1]/tbar[0]/btn[2]").Press
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").Text = Aktualizace.Range("AB9")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").Text = Aktualizace.Range("AB10")
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").SetFocus
session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").CaretPosition = 9
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectContextMenuItem "&XXL"
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Plan tabule"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "PREP_komponenty.XLSX"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 15
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press


Aktualizace.Range("F18").Value = "OK"

'info o aktualizaci - datum a zaměstnanec
Aktualizace.Range("AB6") = Date
Aktualizace.Range("AB7") = Now
Aktualizace.Range("AB8") = Environ("username")
MsgBox "Aktualizace výrobní tabule PREP dokončena"

End Sub
