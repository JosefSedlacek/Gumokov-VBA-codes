Option Explicit

Sub PropsatDataDoReportingu()

If ZapisVyroba.Range("BJ4") = ZapisVyroba.Range("AK4") Then
MsgBox "Zkontrolujte, zde máte zadané správné datum a směnu !!!", Title:="POZOR:"
Else
Application.ScreenUpdating = False
DataVyrobaAVS.Unprotect "123456"

ZapisVyroba.Range("AY4") = Now
ZapisVyroba.Range("BE4") = ZapisVyroba.Range("Y4")
ZapisVyroba.Range("BJ4") = ZapisVyroba.Range("AK4")

VyrobaDB.Range("A2:J59").Copy
DataVyrobaAVS.Activate
DataVyrobaAVS.Range("C50000").End(xlUp).Select
Selection.Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
ZapisVyroba.Activate
DataVyrobaAVS.Protect
DataVyrobaAVS.Protect "123456", False, Contents:=True, Scenarios:=True, AllowFiltering:=True
Application.ScreenUpdating = True
End If
End Sub

Sub poznamka_planovaci()

Dim puvodni_poznamka As String
Dim nova_poznamka As String
Dim hledana_zakazka As String
Dim pocet_najitych_zakazek As Long
Dim radek_zakazky As Long
Dim poznamka As String

puvodni_poznamka = Reporting.Range("G2")
hledana_zakazka = Reporting.Range("G4")
nova_poznamka = InputBox("Zadejte / upravte poznámku:", "POZNÁMKA", puvodni_poznamka)
pocet_najitych_zakazek = Application.WorksheetFunction.CountIf(PozDB.Range("A2:A20000"), hledana_zakazka)
radek_zakazky = PozDB.Range("F2")
poznamka = Reporting.Range("G2").Value

      If pocet_najitych_zakazek > 0 Then
            PozDB.Range("B" & radek_zakazky) = nova_poznamka
      Else
            PozDB.Activate
            PozDB.Range("B30000").End(xlUp).Select
            Selection.Offset(1, 0) = nova_poznamka
            PozDB.Range("A20000").End(xlUp).Select
            Selection.Offset(1, 0) = hledana_zakazka
            Reporting.Activate
'            Reporting.Range("G2").FormulaLocal = "=IFERROR(XLOOKUP($G$4;ZakazkyDB!$B$3:$B$2001;ZakazkyDB!$V$3:$V$2001);"")"
      End If

End Sub
