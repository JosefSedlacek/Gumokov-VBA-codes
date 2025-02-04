Option Explicit

Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub Aktualizovat_ZSD_ODV()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

Dim den As String
Dim datum As Date

den = Worksheets("AKTUALIZACE").Range("A1").Value
datum = Worksheets("AKTUALIZACE").Range("A2").Value

Worksheets("AKTUALIZACE").Range("H4") = "Probíhá"
Worksheets("AKTUALIZACE").Range("H7") = ""
Worksheets("AKTUALIZACE").Range("I7") = ""
Worksheets("AKTUALIZACE").Range("J7") = ""

Dim zacatek As Date
zacatek = Worksheets("AKTUALIZACE").Range("B4").Value

Dim konec As Date
konec = Worksheets("AKTUALIZACE").Range("B5").Value

'session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZSD_ODV"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 2
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/ctxtSP$00003-LOW").Text = zacatek
session.findById("wnd[0]/usr/ctxtSP$00003-HIGH").Text = konec
session.findById("wnd[0]/usr/ctxtSP$00003-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtSP$00003-HIGH").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\Makra exporty"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT_ZSD_ODV.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


Worksheets("AKTUALIZACE").Range("H4") = "Hotovo"
Worksheets("AKTUALIZACE").Range("I4") = den
Worksheets("AKTUALIZACE").Range("J4") = datum

End Sub

Sub propsat_data()

Dim sourceWorkbook As Workbook
Dim targetWorkbook As Workbook
Dim sourceSheet As Worksheet
Dim targetSheet As Worksheet
Dim sourceRange As Range
Dim targetRange As Range
Dim filePath As String

Worksheets("AKTUALIZACE").Range("H7") = "Probíhá"

Dim den As String
Dim datum As Date

den = Worksheets("AKTUALIZACE").Range("A1").Value
datum = Worksheets("AKTUALIZACE").Range("A2").Value

' Define the file path of the source workbook
filePath = "P:\All Access\Makra exporty\EXPORT_ZSD_ODV.xlsx"
Set targetWorkbook = ThisWorkbook
Set sourceWorkbook = Workbooks.Open(filePath)
Set sourceSheet = sourceWorkbook.Sheets(1)
Set sourceRange = sourceSheet.Range("A2:U1500")
Set targetSheet = targetWorkbook.Sheets("SAP sheet")
Set targetRange = targetSheet.Range("A2")
targetRange.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count).Value = sourceRange.Value

' Close the source workbook without saving
sourceWorkbook.Close False

Worksheets("AKTUALIZACE").Range("H7") = "Hotovo"
Worksheets("AKTUALIZACE").Range("I7") = den
Worksheets("AKTUALIZACE").Range("J7") = datum

Dim currentDateTime As String
currentDateTime = Format(Now, "DD.MM.YYYY - H:MM")
Worksheets("KP").Range("B2").Value = currentDateTime

' Informovat o ukončení makra
MsgBox "Hotovo, data jsou aktuální", vbInformation

End Sub
