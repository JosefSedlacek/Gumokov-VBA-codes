Option Explicit

Sub Ulozit_zmeny()

'######################################################################################
'################### Definování objektů a proměnných ##################################

    Dim zdrojList As Worksheet, cilList As Worksheet
    Dim zdrojData As Range, cilData As Range
    Dim zdrojCell As Range, cilCell As Range
    Dim nalezeno As Boolean

    Dim zdrojList2 As Worksheet, cilList2 As Worksheet
    Dim zdrojData2 As Range, cilData2 As Range
    Dim zdrojCell2 As Range, cilCell2 As Range
    Dim nalezeno2 As Boolean

    ThisWorkbook.Worksheets("Hodnocení lisaře").Range("O8").Value = Date

'######################################################################################
'###################### POL data ######################################################

'################### Definování objektů pro POL data ##################################

    ' Nastavení listů
    Set zdrojList = ThisWorkbook.Worksheets("Hodnocení lisaře")
    Set cilList = ThisWorkbook.Worksheets("POL data")

    ' Rozsah dat ke zkopírování
    Set zdrojData = zdrojList.Range("A12:D46")

    ' Rozsah dat v listu "POL data" (sloupec A)
    Set cilData = cilList.Range("A2:A" & cilList.Cells(cilList.Rows.Count, "A").End(xlUp).Row)

'zdrojList.Range("O8").Value = Date


'################### Cyklus - hledání a zápis do POL data #############################

    ' Procházení každé buňky ve zdrojovém rozsahu
    For Each zdrojCell In zdrojData.Columns(1).Cells
        nalezeno = False

        ' Procházení každé buňky ve sloupci A v listu "POL data"
        For Each cilCell In cilData
            ' Pokud se hodnoty shodují, aktualizovat hodnoty ve sloupcích B a C
            If zdrojCell.Value = cilCell.Value Then
                cilCell.Offset(0, 1).Value = zdrojCell.Offset(0, 1).Value
                cilCell.Offset(0, 2).Value = zdrojCell.Offset(0, 2).Value
                cilCell.Offset(0, 3).Value = zdrojCell.Offset(0, 3).Value
                nalezeno = True
                Exit For
            End If
        Next cilCell

        ' Pokud nebyla hodnota nalezena, přidat nová data na konec sloupce A v POL data
        If Not nalezeno Then
            Dim posledniRadek As Long
            posledniRadek = cilList.Cells(cilList.Rows.Count, "A").End(xlUp).Row + 1
            cilList.Cells(posledniRadek, 1).Value = zdrojCell.Value
            cilList.Cells(posledniRadek, 2).Value = zdrojCell.Offset(0, 1).Value
            cilList.Cells(posledniRadek, 3).Value = zdrojCell.Offset(0, 2).Value
            cilList.Cells(posledniRadek, 4).Value = zdrojCell.Offset(0, 3).Value
        End If

    Next zdrojCell


'################### Načtení dat z POL data ###########################################

    'Vložení a protažení vyhledávacího vzorečku

    zdrojList.Range("N12").Select
    ActiveCell.Formula2R1C1 = _
    "=IF(XLOOKUP(RC[-13],'POL data'!C1,'POL data'!C3,"""")=0,"""",XLOOKUP(RC[-13],'POL data'!C1,'POL data'!C3,""""))"
    Range("N12").Select
    Selection.AutoFill Destination:=Range("N12:N46"), Type:=xlFillValues

    zdrojList.Range("O12").Select
    ActiveCell.Formula2R1C1 = _
    "=IF(XLOOKUP(RC[-14],'POL data'!C1,'POL data'!C4,"""")=0,"""",XLOOKUP(RC[-14],'POL data'!C1,'POL data'!C4,""""))"
    Range("O12").Select
    Selection.AutoFill Destination:=Range("O12:O46"), Type:=xlFillValues


'######################################################################################
'###################### LAST SAVE data ################################################

'################### Definování objektů pro LAST SAVE data ############################

    ' Nastavení listů
    Set zdrojList2 = ThisWorkbook.Worksheets("Hodnocení lisaře")
    Set cilList2 = ThisWorkbook.Worksheets("LAST SAVE data")

    ' Rozsah dat ke zkopírování
    Set zdrojData2 = zdrojList.Range("A7:D7")

    ' Rozsah dat v listu "LAST SAVE data" (sloupec A)
    Set cilData2 = cilList2.Range("A2:A" & cilList2.Cells(cilList2.Rows.Count, "A").End(xlUp).Row)


'################### Cyklus - hledání a zápis do LAST SAVE data #######################

    ' Procházení každé buňky ve zdrojovém rozsahu
    For Each zdrojCell2 In zdrojData2.Columns(1).Cells
        nalezeno2 = False

        ' Procházení každé buňky ve sloupci A v listu "LAST SAVE data"
        For Each cilCell2 In cilData2
            ' Pokud se hodnoty shodují, aktualizovat hodnoty ve sloupcích B a C
            If zdrojCell2.Value = cilCell2.Value Then
                cilCell2.Offset(0, 1).Value = zdrojCell2.Offset(0, 1).Value
                cilCell2.Offset(0, 2).Value = zdrojCell2.Offset(0, 2).Value
                cilCell2.Offset(0, 3).Value = zdrojCell2.Offset(0, 3).Value
                nalezeno2 = True
                Exit For
            End If
        Next cilCell2

        ' Pokud nebyla hodnota nalezena, přidat nová data na konec sloupce A v LAST SAVE data
        If Not nalezeno2 Then
            Dim posledniRadek2 As Long
            posledniRadek2 = cilList2.Cells(cilList2.Rows.Count, "A").End(xlUp).Row + 1
            cilList2.Cells(posledniRadek2, 1).Value = zdrojCell2.Value
            cilList2.Cells(posledniRadek2, 2).Value = zdrojCell2.Offset(0, 1).Value
            cilList2.Cells(posledniRadek2, 3).Value = zdrojCell2.Offset(0, 2).Value
            cilList2.Cells(posledniRadek2, 4).Value = zdrojCell2.Offset(0, 3).Value
        End If

    Next zdrojCell2


'################### Načtení dat z LAST SAVE data ######################################

    zdrojList.Range("O8").Select
    ActiveCell.Formula2R1C1 = _
    "=XLOOKUP(R[-1]C[-14],'LAST SAVE data'!C1,'LAST SAVE data'!C2,""žádný zápis"")"
    Range("O8").Select


'######################################################################################
'################### Vymazat výběr - prázdná hodnota ##################################

    Worksheets("Hodnocení lisaře").Range("G5").Value = ""
    Worksheets("Hodnocení lisaře").Range("G5").Select
    MsgBox "Změny uloženy - můžete vybrat dalšího lisaře", vbInformation

End Sub

Sub Ulozit_zmeny_verze2()

    Dim wsHodnoceni As Worksheet
    Dim wsPOLData As Worksheet
    Dim wsLastSave As Worksheet

    Set wsHodnoceni = ThisWorkbook.Worksheets("Hodnocení lisaře")
    Set wsPOLData = ThisWorkbook.Worksheets("POL data")
    Set wsLastSave = ThisWorkbook.Worksheets("LAST SAVE data")

    ' Zapsat dnešní datum (uložení dat k dnešnímu dni)
    ThisWorkbook.Worksheets("Hodnocení lisaře").Range("O8").Value = Date
    
'######################################################################################
'################### Vložení dat do POL data ##########################################
    
    Dim hledanaHodnota As Variant
    Dim posledniRadek As Long
    Dim nalezeno As Range
    
    hledanaHodnota = wsHodnoceni.Range("A12").Value
    Set nalezeno = wsPOLData.Columns("A").Find(What:=hledanaHodnota, LookAt:=xlWhole)
    
    If Not nalezeno Is Nothing Then
        wsHodnoceni.Range("A12:D46").Copy
        nalezeno.PasteSpecial xlPasteValues
    Else
        ' Pokud hodnota není nalezena, vlož na konec tabulky
        posledniRadek = wsPOLData.Cells(wsPOLData.Rows.Count, "A").End(xlUp).Row + 1
        wsHodnoceni.Range("A12:D46").Copy
        wsPOLData.Cells(posledniRadek, 1).PasteSpecial xlPasteValues
    End If

    Application.CutCopyMode = False
    
'######################################################################################
'################### Vložení dat do Last Save #########################################

    Dim hledanaHodnotaLS As Variant
    Dim posledniRadekLS As Long
    Dim nalezenoLS As Range
    
    hledanaHodnotaLS = wsHodnoceni.Range("A7").Value
    Set nalezenoLS = wsLastSave.Columns("A").Find(What:=hledanaHodnotaLS, LookAt:=xlWhole)
    
    If Not nalezenoLS Is Nothing Then
        wsHodnoceni.Range("A7:D7").Copy
        nalezenoLS.PasteSpecial xlPasteValues
    Else
        ' Pokud hodnota není nalezena, vlož na konec tabulky
        posledniRadekLS = wsLastSave.Cells(wsLastSave.Rows.Count, "A").End(xlUp).Row + 1
        wsHodnoceni.Range("A7:D7").Copy
        wsLastSave.Cells(posledniRadekLS, 1).PasteSpecial xlPasteValues
    End If

    Application.CutCopyMode = False


'################### Načtení dat z POL data ###########################################
    
    'Vložení a protažení vyhledávacího vzorečku
        
    wsHodnoceni.Range("N12").Select
    ActiveCell.Formula2R1C1 = _
    "=IF(XLOOKUP(RC[-13],'POL data'!C1,'POL data'!C3,"""")=0,"""",XLOOKUP(RC[-13],'POL data'!C1,'POL data'!C3,""""))"
    Range("N12").Select
    Selection.AutoFill Destination:=Range("N12:N46"), Type:=xlFillValues
    
    wsHodnoceni.Range("O12").Select
    ActiveCell.Formula2R1C1 = _
    "=IF(XLOOKUP(RC[-14],'POL data'!C1,'POL data'!C4,"""")=0,"""",XLOOKUP(RC[-14],'POL data'!C1,'POL data'!C4,""""))"
    Range("O12").Select
    Selection.AutoFill Destination:=Range("O12:O46"), Type:=xlFillValues
    
'################### Načtení dat z LAST SAVE data ######################################
    
    wsHodnoceni.Range("O8").Select
    ActiveCell.Formula2R1C1 = _
    "=XLOOKUP(R[-1]C[-14],'LAST SAVE data'!C1,'LAST SAVE data'!C2,""žádný zápis"")"
    Range("O8").Select
    

'######################################################################################
'################### Vymazat výběr - prázdná hodnota ##################################
    
    Worksheets("Hodnocení lisaře").Range("G5").Value = ""
    Worksheets("Hodnocení lisaře").Range("G5").Select
    MsgBox "Změny uloženy - můžete vybrat dalšího lisaře", vbInformation

End Sub


Sub Aktualizovat()

    ActiveWorkbook.RefreshAll
    
End Sub
