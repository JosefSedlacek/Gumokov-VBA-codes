Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub KAP_Stahni_zpetne_hlaseni()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

    Dim ws As Worksheet
    Dim i As Integer
    Dim Zacatek As String
    Dim Konec As String
    Dim Nazev_zphl As String
    
    Set ws = ThisWorkbook.Sheets("Kapacity")
    For i = 6 To 53
        If ws.Cells(i, "I").Value <> "OK" Then
            Zacatek = ws.Cells(i, "J").Value
            Konec = ws.Cells(i, "K").Value
            Nazev_zphl = ws.Cells(i, "L").Value
            ws.Cells(i, "I").Value = "OK"
            Exit For
        End If
    Next i
    
    MsgBox "Zacatek: " & Zacatek & vbCrLf & _
           "Konec: " & Konec & vbCrLf & _
           "Nazev_zphl: " & Nazev_zphl
           
Dim cesta As Variant
Dim vcera As Variant
Dim dnes As Variant

cesta = ws.Range("I4").Value
vcera = ws.Range("A1").Value
dnes = ws.Range("A2").Value

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "ZPP_ZPHL_NORMA"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[17]").Press
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow = 2
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "2"
session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell

'Definovat rozmezí v SAP transakci
session.FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = Zacatek
session.FindById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = Konec
session.FindById("wnd[0]/usr/ctxtIDATUM").Text = dnes

session.FindById("wnd[0]/usr/ctxtIDATUM").SetFocus
session.FindById("wnd[0]/usr/ctxtIDATUM").CaretPosition = 10
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").PressToolbarContextButton "&MB_VARIANT"
session.FindById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").SelectContextMenuItem "&LOAD"
session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").CurrentCellRow = 1
session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").SelectedRows = "1"
session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").ClickCurrentCell
session.FindById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
session.FindById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").SelectContextMenuItem "&PC"
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = cesta
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = Nazev_zphl
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 16
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

ws.Range("F24") = "posledně proběhlo: " & dnes & vbCrLf & "aktualizoval: " & Environ("Username")


End Sub

Sub KAP_CM02()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Kapacity")
Dim dnes As Variant
dnes = ws.Range("A2").Value

ws.Range("B5") = ""
ws.Range("B9") = ""
ws.Range("B13") = ""
ws.Range("B17") = ""

ws.Range("F6") = ""
ws.Range("F10") = ""
ws.Range("F14") = ""
ws.Range("F18") = ""

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

'##############################
'########## Verze 1 ###########

'session.FindById("wnd[0]").Maximize
'session.FindById("wnd[0]/tbar[0]/okcd").Text = "CM02"
'session.FindById("wnd[0]").SendVKey 0
'session.FindById("wnd[0]/usr/txt[35,7]").Text = "1130"
'session.FindById("wnd[0]/usr/txt[35,7]").SetFocus
'session.FindById("wnd[0]/usr/txt[35,7]").CaretPosition = 4
'session.FindById("wnd[0]/tbar[1]/btn[6]").Press
'session.FindById("wnd[1]/tbar[0]/btn[0]").Press
'session.FindById("wnd[1]/tbar[0]/btn[0]").Press
'session.FindById("wnd[1]/tbar[0]/btn[0]").Press
'session.FindById("wnd[0]/tbar[1]/btn[20]").Press
'session.FindById("wnd[1]/tbar[0]/btn[0]").Press
'session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Kapacity"
'session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "CM02.txt"
'session.FindById("wnd[1]/tbar[0]/btn[11]").Press
'session.FindById("wnd[0]/tbar[0]/btn[3]").Press
'session.FindById("wnd[0]/tbar[0]/btn[3]").Press
'session.FindById("wnd[0]/tbar[0]/btn[3]").Press

'##############################
'########## Verze 2 ###########
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "CM02"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/txt[35,7]").Text = "1130"
session.FindById("wnd[0]/usr/txt[35,7]").SetFocus
session.FindById("wnd[0]/usr/txt[35,7]").CaretPosition = 4
session.FindById("wnd[0]/tbar[1]/btn[6]").Press
session.FindById("wnd[1]/tbar[0]/btn[6]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[0]/usr/lbl[49,0]").SetFocus
session.FindById("wnd[0]/usr/lbl[49,0]").CaretPosition = 24
session.FindById("wnd[0]/tbar[1]/btn[20]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Kapacity"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "CM02.txt"
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

ws.Range("F6") = "posledně proběhlo: " & dnes & vbCrLf & "aktualizoval: " & Environ("Username")
ws.Range("B5") = "OK"

End Sub

Sub KAP_CM38_empty()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Kapacity")
Dim dnes As Variant
dnes = ws.Range("A2").Value

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "CM38"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/txt[35,9]").Text = "1130"
session.FindById("wnd[0]/usr/txt[35,9]").SetFocus
session.FindById("wnd[0]/usr/txt[35,9]").CaretPosition = 4
session.FindById("wnd[0]/tbar[1]/btn[6]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[0]/tbar[1]/btn[20]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Kapacity"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "CM38_empty.txt"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 10
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

ws.Range("F10") = "posledně proběhlo: " & dnes & vbCrLf & "aktualizoval: " & Environ("Username")
ws.Range("B9") = "OK"

End Sub

Sub KAP_CM38_35()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Kapacity")
Dim dnes As Variant
dnes = ws.Range("A2").Value

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "CM38"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/txt[35,3]").Text = "35"
session.FindById("wnd[0]/usr/txt[35,9]").Text = "1130"
session.FindById("wnd[0]/usr/txt[35,3]").CaretPosition = 2
session.FindById("wnd[0]/tbar[1]/btn[6]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[0]/tbar[1]/btn[20]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Kapacity"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "CM38_35.txt"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 7
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

ws.Range("F14") = "posledně proběhlo: " & dnes & vbCrLf & "aktualizoval: " & Environ("Username")
ws.Range("B13") = "OK"

End Sub

Sub KAP_CM38_32()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Kapacity")
Dim dnes As Variant
dnes = ws.Range("A2").Value

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "CM38"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/txt[35,3]").Text = "32"
session.FindById("wnd[0]/usr/txt[35,9]").Text = "1130"
session.FindById("wnd[0]/usr/txt[35,3]").CaretPosition = 2
session.FindById("wnd[0]/tbar[1]/btn[6]").Press
session.FindById("wnd[0]/tbar[1]/btn[20]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Kapacity"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "CM38_32.txt"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 7
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

ws.Range("F18") = "posledně proběhlo: " & dnes & vbCrLf & "aktualizoval: " & Environ("Username")
ws.Range("B17") = "OK"

MsgBox "PowerBI sestava Kapacity byla aktualizována"

End Sub
