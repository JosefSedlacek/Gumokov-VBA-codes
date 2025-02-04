Option Explicit

Sub ZalohaPoznamek()

Worksheets("AKTUALIZACE").Select
Worksheets("AKTUALIZACE").Range("I9") = "Probíhá záloha poznámek"
Worksheets("AKTUALIZACE").Range("I11") = ""

'-----------------------------------
'########### Zkopírování ###########

Application.Wait Now + TimeValue("00:00:02")

Dim zdrojList As Worksheet
Dim cilList As Worksheet
Dim zdrojData1 As Range
Dim cilBunka1 As Range
Dim zdrojData2 As Range
Dim cilBunka2 As Range

Set zdrojList = ThisWorkbook.Sheets("SEZNAM ŠARŽÍ")
Set cilList = ThisWorkbook.Sheets("Zaloha")
Set zdrojData1 = zdrojList.Range("B7:B1500")
Set cilBunka1 = cilList.Range("A2")
Set zdrojData2 = zdrojList.Range("L7:L1500")
Set cilBunka2 = cilList.Range("B2")

cilBunka1.Resize(zdrojData1.Rows.Count, 1).Value = zdrojData1.Value
cilBunka2.Resize(zdrojData2.Rows.Count, 1).Value = zdrojData2.Value

MsgBox "Poznámky uloženy", vbInformation

Application.Wait Now + TimeValue("00:00:03")

Call SAP_mb52

End Sub

Sub Najit_poznamky()

Application.Wait Now + TimeValue("00:00:02")
Worksheets("AKTUALIZACE").Range("I9") = "Hledání poznámek ke směsím se zálohy..."
Application.Wait Now + TimeValue("00:00:03")

'------------------------------------------------------
'######## Zrušení filtrů na listu SEZNAM ŠARŽÍ ########

Application.ScreenUpdating = False

    Worksheets("SEZNAM ŠARŽÍ").Select
    ActiveWorkbook.SlicerCaches("Průřez_Sklad").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Průřez_Stav").ClearManualFilter
    ActiveSheet.Range("$A$6:$A$1500").AutoFilter Field:=1
    ActiveSheet.ListObjects("SARZE").ShowAutoFilterDropDown = False

'------------------------------------------------------
'########### Protažení vyhledávacího vzorce ###########

Worksheets("SEZNAM ŠARŽÍ").Select
Worksheets("SEZNAM ŠARŽÍ").Range("L7").Select
ActiveCell.FormulaR1C1 = _
"=IF(IFERROR(XLOOKUP(RC[-10],Zaloha!C[-11],Zaloha!C[-10]),"""")=0,"""",IFERROR(XLOOKUP(RC[-10],Zaloha!C[-11],Zaloha!C[-10]),""""))"
Range("L7").Select
Selection.AutoFill Destination:=Range("L7:L1500"), Type:=xlFillValues

Application.Wait Now + TimeValue("00:00:02")

Worksheets("SEZNAM ŠARŽÍ").Range("B7").Select
Selection.AutoFill Destination:=Range("SARZE[Finder]")

Worksheets("TESTOVÁNÍ").Select
Worksheets("TESTOVÁNÍ").Range("C7").Select
Selection.AutoFill Destination:=Range("TEST[Finder_all_rows]")

Worksheets("AKTUALIZACE").Select
Application.ScreenUpdating = True

Call Aktualizovat

End Sub

Sub Aktualizovat_vsechno()
    ActiveWorkbook.RefreshAll
    Application.Wait Now + TimeValue("00:00:03")
    MsgBox "Data byla načtena, nyní jsou aktuální"
End Sub

Sub Aktualizovat()

Application.Wait Now + TimeValue("00:00:02")
ActiveWorkbook.RefreshAll
Application.Wait Now + TimeValue("00:00:02")
MsgBox "Data byla načtena, nyní jsou aktuální"


Worksheets("AKTUALIZACE").Range("I9") = "Hotovo."
MsgBox "Aktualizace přehledu směsí dokončena"

'------------------------------------------
'########### Info o aktualizaci ###########

Dim aktualniDatumCas As String
Dim uzivatel As String

aktualniDatumCas = Now
uzivatel = Environ("Username")

Worksheets("AKTUALIZACE").Range("I11") = aktualniDatumCas & vbCrLf & uzivatel

'--------------------------------------
'########### Zamknutí listů ###########

Worksheets("Zaloha").Visible = xlSheetHidden

Worksheets("AKTUALIZACE").Protect Password:="123456", _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=True, _
    AllowFormattingCells:=False, _
    AllowFormattingColumns:=False, _
    AllowFormattingRows:=False, _
    AllowSorting:=False, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
            
Worksheets("SEZNAM ŠARŽÍ").Protect Password:="123456", _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=True, _
    AllowFormattingCells:=False, _
    AllowFormattingColumns:=False, _
    AllowFormattingRows:=False, _
    AllowSorting:=False, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
            
Worksheets("TESTOVÁNÍ").Protect Password:="123456", _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=True, _
    AllowFormattingCells:=False, _
    AllowFormattingColumns:=False, _
    AllowFormattingRows:=False, _
    AllowSorting:=False, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
    
Worksheets("PŘEHLED LIKVIDACE").Protect Password:="123456", _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=True, _
    AllowFormattingCells:=False, _
    AllowFormattingColumns:=False, _
    AllowFormattingRows:=False, _
    AllowSorting:=False, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False


End Sub
