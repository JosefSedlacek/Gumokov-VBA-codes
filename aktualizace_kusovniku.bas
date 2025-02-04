Option Explicit

Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub SAP_MB51()

Worksheets("AKTUALIZACE").Range("A3") = ""
Worksheets("AKTUALIZACE").Range("A6") = ""
Worksheets("AKTUALIZACE").Range("A9") = ""
Worksheets("AKTUALIZACE").Range("A12") = ""
Worksheets("AKTUALIZACE").Range("A15") = ""
Worksheets("AKTUALIZACE").Range("A18") = ""
Worksheets("AKTUALIZACE").Range("A21") = ""
Worksheets("AKTUALIZACE").Range("A25") = ""
Worksheets("AKTUALIZACE").Range("A28") = ""
Worksheets("AKTUALIZACE").Range("A31") = ""
Worksheets("AKTUALIZACE").Range("A34") = ""

'###########################################
'Napojení na SAP GUI API
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

'##########################################
'Definovat rozmezí období pro SAP transakci
Dim Zacatek As Date
    Zacatek = Worksheets("AKTUALIZACE").Range("G3").Value
Dim Konec As Date
    Konec = Worksheets("AKTUALIZACE").Range("G4").Value

'##############################################
'Scripting SAP - proběhne transakce MB51
'Nastavení varianty SEDLAJOS a layoutu /OBECNE

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "MB51"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "SEDLAJOS"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[8]").press

' Zadat správné datumy do transakce:
session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = Zacatek
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = Konec


session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 33
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 22
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 19
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "22"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]").sendVKey 9
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "MB51_101_pohyby.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


Worksheets("AKTUALIZACE").Range("A3") = "OK"

End Sub

Sub VyplnitData()

Dim wsRozklad As Worksheet
Dim wsAktualizace As Worksheet
Dim posledniRadek As Long
Dim aktualizaceDatum As String
Dim i As Long

Set wsRozklad = ThisWorkbook.Worksheets("rozklad.txt")
Set wsAktualizace = ThisWorkbook.Worksheets("AKTUALIZACE")

aktualizaceDatum = wsAktualizace.Range("G4").Value
posledniRadek = wsRozklad.Cells(wsRozklad.Rows.Count, "A").End(xlUp).Row

For i = 1 To posledniRadek
    If wsRozklad.Cells(i, "A").Value <> "" Then
        ' Pokud buňka není prázdná, vyplň datum a hodnotu 1000
        wsRozklad.Cells(i, "B").Value = aktualizaceDatum
        wsRozklad.Cells(i, "C").Value = 1000
    End If
Next i

MsgBox "Úspěšně dokončeno!", vbInformation

Worksheets("AKTUALIZACE").Range("A9") = "OK"

End Sub

Sub UlozitJakoTXT()

Dim wsRozklad As Worksheet
Dim posledniRadek As Long
Dim cesta As String
Dim textFile As String
Dim fs As Object
Dim textStream As Object
Dim radek As String
Dim i As Long

' Nastavení listu
Set wsRozklad = ThisWorkbook.Worksheets("rozklad.txt")

' Zjištění posledního obsazeného řádku ve sloupci A
posledniRadek = wsRozklad.Cells(wsRozklad.Rows.Count, "A").End(xlUp).Row

' Cesta k uloženému souboru
cesta = "C:\Kusovnik\rozklad.txt"

' Kontrola, zda složka existuje, pokud ne, vytvořit ji
If Dir("C:\Kusovnik", vbDirectory) = "" Then
    MkDir "C:\Kusovnik"
End If

' Vytvoření a zápis do textového souboru
Set fs = CreateObject("Scripting.FileSystemObject")
Set textStream = fs.CreateTextFile(cesta, True, False)

' Procházení řádků a zápis dat s tabulátorovými oddělovači
For i = 1 To posledniRadek
    radek = wsRozklad.Cells(i, "A").Value & vbTab & _
            wsRozklad.Cells(i, "B").Value & vbTab & _
            wsRozklad.Cells(i, "C").Value
    textStream.WriteLine radek
Next i

' Zavření textového souboru
textStream.Close

' Zpráva po dokončení
MsgBox "Soubor rozklad.txt byl úspěšně uložen do C:\Kusovnik!", vbInformation

Worksheets("AKTUALIZACE").Range("A12") = "OK"

End Sub

Sub ZPP_ROZKL()

'###########################################
'Napojení na SAP GUI API
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

'##########################################
'Definovat rozmezí období pro SAP transakci
Dim rozpad_zdroj As String
    rozpad_zdroj = Worksheets("AKTUALIZACE").Range("G15").Value
Dim email As String
    email = Worksheets("AKTUALIZACE").Range("G16").Value

'###########################################
' průběh transakce ZPP_ROZKL:
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZPP_ROZKL"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtGS_SEL-PBDNR").Text = "JS1130"
session.findById("wnd[0]/usr/txtGS_SEL-WERKS").Text = "1130"

' vypsání zdrojového souboru a emailu:
session.findById("wnd[0]/usr/ctxtGS_SEL-FILENAME").Text = rozpad_zdroj
session.findById("wnd[0]/usr/txtGS_SEL-EMAIL").Text = email

session.findById("wnd[0]/usr/txtGS_SEL-EMAIL").SetFocus
session.findById("wnd[0]/usr/txtGS_SEL-EMAIL").caretPosition = 25
session.findById("wnd[0]/usr/btn%#AUTOTEXT004").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/btn%#AUTOTEXT003").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/btn%#AUTOTEXT005").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Worksheets("AKTUALIZACE").Range("A15") = "OK"

End Sub

Sub ZPP_ROZKL_ulozit_excel()

'###########################################
'Napojení na SAP GUI API
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

'##########################################
'Definovat rozmezí období pro SAP transakci
Dim kam_ulozit As String
    kam_ulozit = Worksheets("AKTUALIZACE").Range("G18").Value
Dim nazev As String
    nazev = Worksheets("AKTUALIZACE").Range("G19").Value

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZPP_ROZKL"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtGS_SEL-PBDNR").Text = "JS1130"
session.findById("wnd[0]/usr/txtGS_SEL-WERKS").Text = "1130"
session.findById("wnd[0]/usr/txtGS_SEL-WERKS").SetFocus
session.findById("wnd[0]/usr/txtGS_SEL-WERKS").caretPosition = 4
session.findById("wnd[0]/usr/btn%#AUTOTEXT009").press
session.findById("wnd[0]/usr/cntlSCR300_CC1/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlSCR300_CC1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press

'vyplnění cesty a názvu uložení xlsx souboru:
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = kam_ulozit
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nazev
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Worksheets("AKTUALIZACE").Range("A18") = "OK"

End Sub

Sub ZPP_KALVZ()

'###########################################
'Napojení na SAP GUI API
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

'##########################################
'Definovat rozmezí období pro SAP transakci
Dim datum_kalkulace As Date
    datum_kalkulace = Worksheets("AKTUALIZACE").Range("G21").Value
Dim kam_ulozit2 As String
    kam_ulozit2 = Worksheets("AKTUALIZACE").Range("G22").Value
Dim nazev2 As String
    nazev2 = Worksheets("AKTUALIZACE").Range("G23").Value

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZPP_KALVZ"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell

' nastavení správného datumu kalkulace
session.findById("wnd[0]/usr/ctxtKADKY-LOW").Text = datum_kalkulace

session.findById("wnd[0]/usr/ctxtKADKY-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtKADKY-LOW").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press

' správná cesta k uložení souboru
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = kam_ulozit2
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nazev2

session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Worksheets("AKTUALIZACE").Range("A21") = "OK"

End Sub

Sub CopyExcelFiles()

Dim sourceFolder As String
Dim targetFolder As String
Dim fileNames As Variant
Dim fileName As Variant
Dim fso As Object

' Nastavení cest ke složkám
sourceFolder = "C:\Kusovnik\"
targetFolder = "P:\All Access\TB HRA KPIs\podklady\Kusovniky\"

' Názvy souborů, které chceme zkopírovat
fileNames = Array("KALKULACE_VPC2.XLSX", "ROZPAD_KUSOVNIKU.XLSX")

' Vytvoření instance FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Kontrola, zda cílová složka existuje, pokud ne, vytvoří se
If Not fso.FolderExists(targetFolder) Then
    fso.CreateFolder targetFolder
End If

' Kopírování jednotlivých souborů
For Each fileName In fileNames
    ' Zkopíruj soubor ze zdrojové složky do cílové
    If fso.FileExists(sourceFolder & fileName) Then
        fso.CopyFile Source:=sourceFolder & fileName, Destination:=targetFolder & fileName, OverwriteFiles:=True
    Else
        MsgBox "Soubor " & fileName & " nebyl nalezen ve složce " & sourceFolder, vbExclamation
    End If
Next fileName

MsgBox "Soubory byly úspěšně zkopírovány.", vbInformation

Worksheets("AKTUALIZACE").Range("A25") = "OK"

End Sub

Sub ZPPPOSTUP()

'###########################################
'Napojení na SAP GUI API
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

'##########################################
'Definovat rozmezí období pro SAP transakci
Dim datum_platnosti As Date
    datum_platnosti = Worksheets("AKTUALIZACE").Range("I28").Value
Dim kam_ulozit3 As String
    kam_ulozit3 = Worksheets("AKTUALIZACE").Range("G28").Value
Dim nazev3 As String
    nazev3 = Worksheets("AKTUALIZACE").Range("G29").Value



session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "ZPPPOSTUP"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell

' datum:
session.findById("wnd[0]/usr/ctxtPN_DATUV").Text = datum_platnosti

session.findById("wnd[0]/usr/ctxtPN_DATUV").SetFocus
session.findById("wnd[0]/usr/ctxtPN_DATUV").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

'cesta:
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = kam_ulozit3
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nazev3


session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

MsgBox "Postupy jsou staženy.", vbInformation

Worksheets("AKTUALIZACE").Range("A28") = "OK"

End Sub

Sub Podklady_POSTUPY()

Dim ws As Worksheet
Dim newWorkbook As Workbook
Dim savePath As String

Set ws = ThisWorkbook.Sheets("POSTUPY")
savePath = "P:\All Access\TB HRA KPIs\podklady\Kusovniky\POSTUPY.XLSX"
ws.Copy

Set newWorkbook = ActiveWorkbook

newWorkbook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

newWorkbook.Close SaveChanges:=False

MsgBox "POSTUPY uloženy v podkladech pro PowerBI.", vbInformation

Worksheets("AKTUALIZACE").Range("A34") = "OK"

End Sub
